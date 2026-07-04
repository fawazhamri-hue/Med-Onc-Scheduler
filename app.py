"""
OncoScheduler Web — Flask backend.

Routes:
  GET  /                  → main page (week setup)
  GET  /api/people        → list all active people
  GET  /api/template      → clinic template
  POST /api/solve         → run solver, return schedule JSON
  POST /api/export/xlsx   → return Excel file
  POST /api/export/docx   → return Word file
  GET  /api/settings      → solver settings
  POST /api/settings      → save solver settings
  POST /api/import/rota   → parse ROTA file, return proposals
"""

import io
import json
import os
import tempfile
import traceback
from datetime import datetime, timedelta, date

import pandas as pd
from flask import Flask, jsonify, request, send_file

from db import Database
from exporter_docx import export_docx
from exporter_xlsx import export_xlsx
from solver_core import SolverConfig, solve

app = Flask(__name__, static_folder="static", static_url_path="")

DB_PATH = os.environ.get("DB_PATH", "oncoscheduler.db")

DAYS = ["sun", "mon", "tue", "wed", "thu"]


def get_db():
    return Database(DB_PATH)


# ── helpers ──────────────────────────────────────────────────────────────────

def _parse_slots(val):
    slots = []
    for part in (val or "").split(","):
        part = part.strip()
        if not part or ":" not in part:
            continue
        dd, ss = part.split(":", 1)
        dd, ss = dd.strip().lower()[:3], ss.strip().upper()
        if dd in DAYS and ss in ("AM", "PM"):
            slots.append((dd, ss))
    return slots


def _build_cfg(db, chemo_slots_override=None):
    cfg = SolverConfig()
    for key, val in db.all_settings().items():
        if key == "solver_timeout_s":
            try:
                cfg.solver_timeout_sec = float(val)
            except (TypeError, ValueError):
                pass
        elif key == "chemo_active":
            cfg.chemo_active = val.strip() not in ("0", "false", "no", "")
        elif key == "chemo_slots":
            slots = _parse_slots(val)
            if slots:
                cfg.chemo_slots = slots
        elif hasattr(cfg, key):
            try:
                setattr(cfg, key, int(val))
            except (TypeError, ValueError):
                pass
    # Per-week override: the Schedule page lets the chief tick which day/session
    # slots need chemo THIS week, since it isn't the same every week (e.g. some
    # weeks have no Sunday AM chemo). This replaces the Rules-page default only
    # for this solve — Rules stays the fallback for weeks with no override saved.
    if chemo_slots_override is not None:
        cfg.chemo_slots = [(d, s) for d, s in chemo_slots_override
                          if d in DAYS and s in ("AM", "PM")]
        cfg.chemo_active = len(cfg.chemo_slots) > 0
    return cfg


def _assemble_from_payload(payload):
    """
    Build DataFrames directly from the JSON payload sent by the frontend.
    Payload shape:
      {
        fellows:   [{name, rotation, role, min_clinics, max_clinics, status}],
        residents: [{name, type, min_clinics, max_clinics, status}],
        clinics:   [{day, session, consultant, specialty, clinic_type, patients}],
        pins:      [{person, day, session, consultant}],
        availability: [{person, day, session, reason}],
        week_start: "2026-06-14"
      }
    """
    fellows_df = pd.DataFrame(payload.get("fellows", []))
    if not fellows_df.empty:
        fellows_df = fellows_df.rename(columns={
            "name": "Fellow", "role": "Role", "rotation": "Rotation",
            "min_clinics": "Min Clinics", "max_clinics": "Max Clinics",
            "status": "Status"
        })

    residents_df = pd.DataFrame(payload.get("residents", []))
    if not residents_df.empty:
        residents_df = residents_df.rename(columns={
            "name": "Resident", "type": "Type",
            "min_clinics": "Min Clinics", "max_clinics": "Max Clinics",
            "status": "Status"
        })

    clinics_df = pd.DataFrame(payload.get("clinics", []))
    if not clinics_df.empty:
        clinics_df = clinics_df.rename(columns={
            "day": "Day", "session": "Session",
            "consultant": "Consultant", "specialty": "Specialty",
            "clinic_type": "ClinicType", "patients": "Patients"
        })
        # Capitalise Day so solver matches
        clinics_df["Day"] = clinics_df["Day"].str.capitalize()

    pins_df = pd.DataFrame(payload.get("pins", []))
    if not pins_df.empty:
        pins_df = pins_df.rename(columns={
            "person": "Fellow", "day": "Day",
            "session": "Session", "consultant": "Consultant"
        })
    else:
        pins_df = pd.DataFrame(columns=["Fellow", "Day", "Session", "Consultant"])

    avail_list = payload.get("availability", [])
    availability_df = pd.DataFrame(avail_list) if avail_list else None
    if availability_df is not None and not availability_df.empty:
        availability_df = availability_df.rename(columns={
            "person": "Person", "day": "Day",
            "session": "Session", "reason": "Reason"
        })

    week_start = payload.get("week_start", "")
    try:
        ws = datetime.strptime(week_start, "%Y-%m-%d").date()
        date_map = {
            DAYS[i].capitalize(): (ws + timedelta(days=i)).strftime("%d %b")
            for i in range(5)
        }
    except (ValueError, TypeError):
        date_map = {d.capitalize(): d.capitalize() for d in DAYS}

    return fellows_df, residents_df, clinics_df, pins_df, availability_df, date_map


# ── routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return app.send_static_file("index.html")


@app.route("/api/people")
def api_people():
    db = get_db()
    people = db.list_people(include_inactive=True)
    return jsonify(people)


@app.route("/api/template")
def api_template():
    db = get_db()
    return jsonify(db.get_template())


@app.route("/api/settings", methods=["GET"])
def api_get_settings():
    db = get_db()
    return jsonify(db.all_settings())


@app.route("/api/settings", methods=["POST"])
def api_save_settings():
    db = get_db()
    data = request.get_json()
    for key, val in (data or {}).items():
        db.set_setting(key, str(val))
    return jsonify({"ok": True})


def _merge_db_roster(payload, db):
    """
    Build the full solver payload from the database roster + the per-week data
    the frontend sent (clinics, week dates, one-off pins/exclusions).

    This is the web equivalent of the desktop app's Database.assemble_week:
    it resolves date-range leaves, rotation windows, recurring LTC pins, and
    recurring exclusions from the DB so the Schedule page never has to carry
    permanent roster state. Active/inactive comes straight from the DB row.
    """
    from datetime import datetime, timedelta
    DAYS_ = ["sun", "mon", "tue", "wed", "thu"]

    week_start = payload.get("week_start", "")
    try:
        wstart = datetime.strptime(week_start, "%Y-%m-%d").date()
    except (ValueError, TypeError):
        # No valid week → fall back to whatever the frontend sent verbatim
        return payload
    day_dates = {DAYS_[i]: wstart + timedelta(days=i) for i in range(5)}
    wk_end = day_dates["thu"]

    def parse_d(s):
        if not s:
            return None
        return datetime.strptime(str(s)[:10], "%Y-%m-%d").date()

    def in_window(p):
        sd, ed = parse_d(p["start_date"]), parse_d(p["end_date"])
        if sd and sd > wk_end:
            return False
        if ed and ed < wstart:
            return False
        return True

    people = [p for p in db.list_people(include_inactive=True)]
    by_id = {p["id"]: p for p in people}

    # Date-range leave resolution: full week → Leave status; partial → exclusions.
    full_week_leave = set()
    partial_excl = []
    for lv in db.list_leaves():
        p = by_id.get(lv["person_id"])
        if not p:
            continue
        lsd, led = parse_d(lv["start_date"]), parse_d(lv["end_date"])
        if not lsd or not led:
            continue
        covered = [d for d in DAYS_ if lsd <= day_dates[d] <= led]
        if not covered:
            continue
        if len(covered) == 5:
            full_week_leave.add(p["name"])
        else:
            for d in covered:
                partial_excl.append({"person": p["name"], "day": d,
                                     "session": "All", "reason": lv["note"] or "leave"})

    fellows, residents = [], []
    for p in people:
        # Inactive in DB OR outside rotation window OR full-week leave → status Leave
        is_off = (not p["active"]) or (not in_window(p)) or (p["name"] in full_week_leave)
        status = "Leave" if is_off else "Active"
        if p["kind"] in ("fellow", "np"):
            fellows.append({
                "name": p["name"], "rotation": p["rotation"] or "",
                "role": "NP" if p["kind"] == "np" else "Fellow",
                "min_clinics": p["min_clinics"], "max_clinics": p["max_clinics"],
                "status": status,
            })
        elif p["kind"] in ("resident", "assistant"):
            residents.append({
                "name": p["name"],
                "type": "Assistant" if p["kind"] == "assistant" else "Resident",
                "min_clinics": p["min_clinics"], "max_clinics": p["max_clinics"],
                "status": status,
            })

    # Recurring LTC pins from DB (skip people who are off this week — the solver
    # would skip them anyway, but filtering here keeps demand accounting clean).
    db_pins = []
    off_names = {f["name"] for f in fellows if f["status"] == "Leave"} | \
                {r["name"] for r in residents if r["status"] == "Leave"}
    for rp in db.list_recurring_pins():
        p = by_id.get(rp["person_id"])
        if not p or p["name"] in off_names:
            continue
        db_pins.append({"person": p["name"], "day": rp["day"],
                        "session": rp["session"], "consultant": rp["consultant"]})

    # Recurring exclusions from DB
    db_excl = []
    for re_ in db.list_recurring_exclusions():
        p = by_id.get(re_["person_id"])
        if not p:
            continue
        db_excl.append({"person": p["name"], "day": re_["day"],
                        "session": re_["session"], "reason": "recurring"})

    # Combine: DB pins + frontend one-off pins; DB exclusions + partial leaves
    # + frontend one-off exclusions.
    all_pins = db_pins + payload.get("pins", [])
    all_excl = db_excl + partial_excl + payload.get("availability", [])

    return {
        **payload,
        "fellows": fellows,
        "residents": residents,
        "pins": all_pins,
        "availability": all_excl,
    }


# =====================================================
# STEP 2: post-solve validator (independent audit, warn-only, never blocks)
# =====================================================
# =====================================================
# Recompute Workload Summary totals from a (possibly hand-edited) schedule.
# The solver's own fellows_summary/residents_summary/assistants_summary are
# only correct for the schedule it produced; once the chief edits assignments
# in the browser, those counts go stale. Export must reflect what's actually
# on the page, not what the solver originally said.
# =====================================================
def _recompute_summaries(schedule, fellows_df, residents_df):
    def names_of(row):
        raw = str(row.get("Assigned", "") or "")
        return [n.strip().rstrip("*") for n in raw.replace("/", ",").split(",")
                if n.strip() and n.strip() not in ("—", "-")]

    totals = {}
    inside = {}
    for r in (schedule or []):
        specialty = str(r.get("Specialty", "") or "")
        for n in names_of(r):
            key = n.lower()
            totals[key] = totals.get(key, 0) + 1
            if specialty:
                inside.setdefault(key, 0)

    # Fellows/NPs: pull Rotation + Role from the merged roster DataFrame.
    fellows_summary = []
    if fellows_df is not None and not fellows_df.empty:
        for _, row in fellows_df.iterrows():
            name = str(row.get("Fellow", "")).strip()
            key = name.lower()
            if key not in totals:
                continue  # not assigned anywhere in the (edited) schedule
            rotation = str(row.get("Rotation", "") or "")
            role = str(row.get("Role", "Fellow") or "Fellow")
            total = totals.get(key, 0)
            ins = sum(1 for r in (schedule or [])
                     if key in [n.lower() for n in names_of(r)]
                     and str(r.get("Specialty", "")) == rotation)
            fellows_summary.append({"Fellow": name, "Rotation": rotation,
                                    "Role": role, "Total": total, "Inside": ins})

    residents_summary = []
    if residents_df is not None and not residents_df.empty:
        for _, row in residents_df.iterrows():
            name = str(row.get("Resident", "")).strip()
            key = name.lower()
            if key not in totals:
                continue
            residents_summary.append({"Resident": name, "Total": totals.get(key, 0)})

    return fellows_summary, residents_summary


def validate_schedule(schedule, fellows, residents):
    """
    Re-checks the solver's OUTPUT (not the model) against the hard rules:
      V1  same person in 2+ clinics, same day+session   -> error
      V2  AddOn with helper count != 1                  -> error
      V3  person assigned more clinics than their Max   -> error
      V4  chemo slot with 0 helpers                      -> error
      V5  person on Leave/inactive but still assigned    -> error
      V6  person under their Min                          -> info (not urgent)
    Returns a list of {"level": "error"|"info", "msg": str}.
    """
    out = []
    limits = {}
    for p in (fellows or []):
        limits[str(p.get("name", "")).strip().lower()] = (
            int(p.get("min_clinics", 0) or 0),
            int(p.get("max_clinics", 99) or 99),
            str(p.get("status", "Active")))
    for p in (residents or []):
        limits[str(p.get("name", "")).strip().lower()] = (
            int(p.get("min_clinics", 0) or 0),
            int(p.get("max_clinics", 99) or 99),
            str(p.get("status", "Active")))

    def names_of(row):
        raw = str(row.get("Assigned", "") or "")
        return [n.strip().rstrip("*") for n in raw.replace("/", ",").split(",")
                if n.strip() and n.strip() not in ("—", "-")]

    # V1: same-session double-booking
    per_session = {}
    for r in (schedule or []):
        key = (r.get("Day"), r.get("Session"))
        for n in names_of(r):
            per_session.setdefault(key, {}).setdefault(n.lower(), []).append(
                r.get("Consultant", "?"))
    for (day, sess), people in per_session.items():
        for n, cons_list in people.items():
            if len(cons_list) > 1:
                out.append({"level": "error",
                            "msg": f"DOUBLE-BOOKED: {n.title()} on {day} {sess} "
                                   f"in {len(cons_list)} clinics ({', '.join(cons_list)})"})

    # V2: AddOn helper count must be exactly 1
    for r in (schedule or []):
        if r.get("IsAddOn"):
            k = len(names_of(r))
            if k != 1:
                out.append({"level": "error",
                            "msg": f"ADDON STAFFING: {r.get('Day')} {r.get('Session')} "
                                   f"{r.get('Consultant')} has {k} helper(s) (must be exactly 1)"})

    # V3 / V5 / V6: totals vs Min/Max/Status
    totals = {}
    for r in (schedule or []):
        for n in names_of(r):
            totals[n.lower()] = totals.get(n.lower(), 0) + 1
    for n, t in totals.items():
        if n in limits:
            mn, mx, status = limits[n]
            if str(status).lower() == "leave":
                out.append({"level": "error",
                            "msg": f"ON-LEAVE ASSIGNED: {n.title()} is marked Leave/"
                                   f"inactive but has {t} clinic(s) this week"})
            if t > mx:
                out.append({"level": "error",
                            "msg": f"OVER MAX: {n.title()} has {t} clinics (max {mx})"})
            elif t < mn:
                out.append({"level": "info",
                            "msg": f"Under min: {n.title()} has {t} clinics (min {mn})"})

    # V4: chemo slot uncovered
    for r in (schedule or []):
        cons = str(r.get("Consultant", "")).lower()
        if "chemo" in cons and len(names_of(r)) == 0:
            out.append({"level": "error",
                        "msg": f"CHEMO UNCOVERED: {r.get('Day')} {r.get('Session')}"})
    return out


# =====================================================
# STEP 3: pre-generate summary (what WILL be sent to the solver, no solving)
# =====================================================
@app.route("/api/solve-preview", methods=["POST"])
def api_solve_preview():
    try:
        payload = request.get_json() or {}
        db = get_db()
        merged = _merge_db_roster(dict(payload), db)
        fellows = merged.get("fellows", [])
        residents = merged.get("residents", [])
        clinics = payload.get("clinics", [])

        active_fellows = [f for f in fellows if f.get("status") != "Leave"
                          and str(f.get("role", "")).lower() != "np"]
        active_nps = [f for f in fellows if f.get("status") != "Leave"
                     and str(f.get("role", "")).lower() == "np"]
        active_residents = [r for r in residents if r.get("status") != "Leave"]
        excluded = ([f["name"] for f in fellows if f.get("status") == "Leave"] +
                   [r["name"] for r in residents if r.get("status") == "Leave"])

        addons = [c for c in clinics if (c.get("clinic_type") or "") == "AddOn"]
        total_patients = sum(int(c.get("patients", 0) or 0) for c in clinics)

        return jsonify({
            "active_fellows": len(active_fellows),
            "active_nps": len(active_nps),
            "active_residents": len(active_residents),
            "excluded": excluded,
            "pins": merged.get("pins", []),
            "clinic_count": len(clinics),
            "addon_count": len(addons),
            "total_patients": total_patients,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/solve", methods=["POST"])
def api_solve():
    try:
        payload = request.get_json()
        db = get_db()
        # Merge permanent roster data from the DB so the Schedule page doesn't
        # have to carry it. The frontend sends clinics + week + any one-off
        # pins/exclusions; the server adds people, recurring LTC pins,
        # recurring exclusions, and date-range leaves.
        payload = _merge_db_roster(payload, db)

        # --- DIAGNOSTIC TAP (temporary, for double-booking investigation) ---
        # Dumps the exact merged payload the solver receives, every solve,
        # overwriting the previous one. Does not affect solving in any way.
        try:
            dump_path = os.path.join(os.path.dirname(DB_PATH) or ".",
                                     "last_solve_payload.json")
            with open(dump_path, "w") as f:
                json.dump({"timestamp": datetime.now().isoformat(),
                          "payload": payload}, f, indent=2, default=str)
        except Exception:
            pass  # diagnostic failure must never break solving
        # --- END DIAGNOSTIC TAP ---

        fellows_df, residents_df, clinics_df, pins_df, availability_df, date_map = \
            _assemble_from_payload(payload)
        cfg = _build_cfg(db, chemo_slots_override=payload.get("chemo_slots"))
        result = solve(clinics_df, fellows_df, residents_df, pins_df,
                       config=cfg, availability_path=availability_df)

        # STEP 2: independent post-solve audit of the OUTPUT (warn-only).
        validation_warnings = validate_schedule(
            result.get("schedule", []), payload.get("fellows", []),
            payload.get("residents", []))

        return jsonify({
            "status": result["status"],
            "schedule": result.get("schedule", []),
            "fellows_summary": result.get("fellows_summary", []),
            "residents_summary": result.get("residents_summary", []),
            "assistants_summary": result.get("assistants_summary", []),
            "on_leave": result.get("on_leave", []),
            "purple_orphans_leave": result.get("purple_orphans_leave", []),
            "availability_orphans": result.get("availability_orphans", []),
            "date_map": date_map,
            "validation_warnings": validation_warnings,
        })
    except Exception as e:
        return jsonify({"error": traceback.format_exc(), "message": str(e)}), 500


@app.route("/api/validate", methods=["POST"])
def api_validate():
    """STEP 2: re-run the validator against an (edited) schedule the person
    is looking at, without re-solving. Used by the in-app result editor."""
    try:
        data = request.get_json() or {}
        db = get_db()
        merged = _merge_db_roster({"fellows": [], "residents": [], "pins": [],
                                   "availability": [], "clinics": [],
                                   "week_start": data.get("week_start", "")}, db)
        warnings = validate_schedule(data.get("schedule", []),
                                     merged.get("fellows", []),
                                     merged.get("residents", []))
        return jsonify(warnings)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/week/<week_start>", methods=["GET"])
def api_week_load(week_start):
    """STEP 4: load a previously saved week's clinics/pins/exclusions/chemo slots."""
    try:
        db = get_db()
        wk = next((w for w in db.list_weeks() if w["start_date"] == week_start), None)
        chemo_raw = db.get_setting(f"chemo_slots:{week_start}", None)
        chemo_slots = _parse_slots(chemo_raw) if chemo_raw else None
        if not wk:
            return jsonify({"exists": False, "clinics": [], "pins": [],
                            "exclusions": [], "chemo_slots": chemo_slots})
        clinics = [{"day": c["day"], "session": c["session"],
                    "consultant": c["consultant"], "specialty": c["specialty"],
                    "clinic_type": c["clinic_type"] or "", "patients": c["patients"]}
                   for c in db.list_week_clinics(wk["id"])]
        pins = [{"person": p["person_name"], "day": p["day"],
                 "session": p["session"], "consultant": p["consultant"]}
                for p in db.list_week_pins(wk["id"])]
        excl = [{"person": x["person_name"], "day": x["day"],
                 "session": x["session"], "reason": x["reason"]}
                for x in db.list_week_exclusions(wk["id"])]
        return jsonify({"exists": True, "clinics": clinics, "pins": pins,
                        "exclusions": excl, "chemo_slots": chemo_slots})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/week/<week_start>", methods=["POST"])
def api_week_save(week_start):
    """STEP 4: save/overwrite this week's clinics/pins/exclusions/chemo slots."""
    try:
        data = request.get_json() or {}
        db = get_db()
        wid = db.create_week(week_start)  # idempotent: returns existing id if present
        for c in db.list_week_clinics(wid):
            db.delete_week_clinic(c["id"])
        for p in db.list_week_pins(wid):
            db.delete_week_pin(p["id"])
        for x in db.list_week_exclusions(wid):
            db.delete_week_exclusion(x["id"])
        for c in data.get("clinics", []):
            db.add_week_clinic(wid, c["day"], c["session"], c["consultant"],
                               c.get("specialty", ""), c.get("clinic_type", ""),
                               int(c.get("patients", 0) or 0))
        for p in data.get("pins", []):
            db.add_week_pin(wid, p["person"], p["day"], p["session"], p["consultant"])
        for x in data.get("exclusions", []):
            db.add_week_exclusion(wid, x["person"], x["day"], x["session"],
                                  x.get("reason", ""))
        # Per-week chemo slot override, stored as "day:sess, day:sess, ..."
        if "chemo_slots" in data:
            slots_str = ", ".join(f"{d}:{s}" for d, s in data["chemo_slots"])
            db.set_setting(f"chemo_slots:{week_start}", slots_str)
        return jsonify({"ok": True, "week_id": wid})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/weeks", methods=["GET"])
def api_weeks_list():
    db = get_db()
    return jsonify([w["start_date"] for w in db.list_weeks()])


@app.route("/api/debug/last-payload", methods=["GET"])
def api_debug_last_payload():
    """Temporary diagnostic route — returns the exact payload the solver
    received on the most recent /api/solve call. Used to investigate the
    double-booking bug with real data instead of reconstructions."""
    dump_path = os.path.join(os.path.dirname(DB_PATH) or ".",
                             "last_solve_payload.json")
    if not os.path.exists(dump_path):
        return jsonify({"error": "No solve has run yet on this server."}), 404
    with open(dump_path) as f:
        return jsonify(json.load(f))


@app.route("/api/export/xlsx", methods=["POST"])
def api_export_xlsx():
    try:
        payload = request.get_json()
        result_data = payload.get("result")
        db = get_db()
        payload = _merge_db_roster(payload, db)
        fellows_df, residents_df, clinics_df, pins_df, availability_df, date_map = \
            _assemble_from_payload(payload)
        # Drop Leave-status people from the workload summary (inactive / on-leave /
        # out-of-window). They contribute nothing and shouldn't clutter the export.
        if not fellows_df.empty and "Status" in fellows_df.columns:
            fellows_df = fellows_df[fellows_df["Status"] != "Leave"].reset_index(drop=True)
        if not residents_df.empty and "Status" in residents_df.columns:
            residents_df = residents_df[residents_df["Status"] != "Leave"].reset_index(drop=True)
        week_start = payload.get("week_start", "")
        week_end = payload.get("week_end", "")

        # Rebuild result dict with DataFrames for the exporter
        result = {**result_data, "availability": [], "availability_orphans": [],
                  "purple_orphans_leave": [], "on_leave": []}

        # Recompute Workload Summary from the schedule actually being exported
        # (which may have been hand-edited in the browser after solving).
        fs, rs = _recompute_summaries(result.get("schedule", []), fellows_df, residents_df)
        if fs:
            result["fellows_summary"] = fs
        if rs:
            result["residents_summary"] = rs

        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        export_xlsx(result, fellows_df, residents_df, tmp.name,
                    week_start, week_end, date_map)
        fname = f"Schedule_{week_start}.xlsx"
        return send_file(tmp.name, as_attachment=True,
                         download_name=fname,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return jsonify({"error": traceback.format_exc(), "message": str(e)}), 500


@app.route("/api/export/docx", methods=["POST"])
def api_export_docx():
    try:
        payload = request.get_json()
        result_data = payload.get("result")
        db = get_db()
        payload = _merge_db_roster(payload, db)
        fellows_df, residents_df, clinics_df, pins_df, availability_df, date_map = \
            _assemble_from_payload(payload)
        week_start = payload.get("week_start", "")
        week_end = payload.get("week_end", "")

        result = {**result_data, "availability": [], "availability_orphans": [],
                  "purple_orphans_leave": [], "on_leave": []}

        # Recompute Workload Summary from the schedule actually being exported
        # (which may have been hand-edited in the browser after solving).
        fs, rs = _recompute_summaries(result.get("schedule", []), fellows_df, residents_df)
        if fs:
            result["fellows_summary"] = fs
        if rs:
            result["residents_summary"] = rs

        day_info = [(d.capitalize(), d.upper(), date_map.get(d.capitalize(), ""))
                    for d in DAYS]
        tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        tmp.close()
        export_docx(result, tmp.name, week_start, week_end, day_info)
        fname = f"Schedule_{week_start}.docx"
        return send_file(tmp.name, as_attachment=True,
                         download_name=fname,
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
        return jsonify({"error": traceback.format_exc(), "message": str(e)}), 500


@app.route("/api/import/patients", methods=["POST"])
def api_import_patients():
    try:
        from importer_rota import parse_patient_numbers, list_patient_sheets
        f = request.files.get("file")
        if not f:
            return jsonify({"error": "No file uploaded"}), 400
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        f.save(tmp.name)
        tmp.close()
        sheet = request.form.get("sheet") or None
        sheets = list_patient_sheets(tmp.name)
        # If the workbook has multiple sheets and none chosen yet, return the
        # list so the frontend can ask which week.
        if len(sheets) > 1 and not sheet:
            os.unlink(tmp.name)
            return jsonify({"need_sheet": True, "sheets": sheets})
        clinics = parse_patient_numbers(tmp.name, sheet=sheet)
        os.unlink(tmp.name)
        return jsonify({"clinics": clinics})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/import/rota", methods=["POST"])
def api_import_rota():
    try:
        from importer_rota import parse_rota, diff_against_db, apply_proposal
        f = request.files.get("file")
        if not f:
            return jsonify({"error": "No file uploaded"}), 400
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        f.save(tmp.name)
        tmp.close()
        year = int(request.form.get("year", datetime.now().year))
        db = get_db()
        props = parse_rota(tmp.name, default_year=year)
        props = diff_against_db(props, db)
        os.unlink(tmp.name)
        return jsonify(props)
    except Exception as e:
        return jsonify({"error": traceback.format_exc(), "message": str(e)}), 500


@app.route("/api/import/rota/apply", methods=["POST"])
def api_import_rota_apply():
    try:
        from importer_rota import apply_proposal
        proposals = request.get_json()
        db = get_db()
        results = []
        for p in (proposals or []):
            try:
                msg = apply_proposal(p, db)
                results.append({"ok": True, "msg": msg})
            except Exception as e:
                results.append({"ok": False, "msg": str(e)})
        return jsonify(results)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/people/export", methods=["GET"])
def api_people_export():
    """
    Full roster backup — every person plus their leaves, recurring pins, and
    recurring exclusions, as one JSON file. Re-importable via /api/people/import.
    Doubles as a general backup: download occasionally to keep a local copy
    independent of the server.
    """
    try:
        db = get_db()
        people = db.list_people(include_inactive=True)
        out = []
        for p in people:
            out.append({
                "name": p["name"], "kind": p["kind"], "rotation": p["rotation"],
                "min_clinics": p["min_clinics"], "max_clinics": p["max_clinics"],
                "start_date": p["start_date"], "end_date": p["end_date"],
                "notes": p["notes"], "active": p["active"],
                "leaves": [{"start_date": l["start_date"], "end_date": l["end_date"],
                           "note": l["note"]} for l in db.list_leaves(p["id"])],
                "pins": [{"day": x["day"], "session": x["session"],
                         "consultant": x["consultant"]}
                        for x in db.list_recurring_pins(p["id"])],
                "exclusions": [{"day": x["day"], "session": x["session"]}
                              for x in db.list_recurring_exclusions(p["id"])],
            })
        payload = {"exported_at": datetime.now().isoformat(timespec="seconds"),
                  "people": out}
        buf = io.BytesIO(json.dumps(payload, indent=2).encode())
        return send_file(buf, as_attachment=True,
                         download_name=f"oncoscheduler_roster_{date.today().isoformat()}.json",
                         mimetype="application/json")
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/people/import", methods=["POST"])
def api_people_import():
    """
    Restore a roster previously downloaded via /api/people/export.
    Safe by default: a person whose name already exists in the DB is SKIPPED
    entirely (including their leaves/pins/exclusions) so a re-import can never
    silently overwrite live data. Returns a per-person report.
    """
    try:
        f = request.files.get("file")
        if not f:
            return jsonify({"error": "No file uploaded"}), 400
        data = json.load(f.stream)
        people_in = data.get("people", [])
        db = get_db()
        existing_names = {p["name"].strip().lower()
                          for p in db.list_people(include_inactive=True)}
        results = []
        for p in people_in:
            name = str(p.get("name", "")).strip()
            if not name:
                continue
            if name.lower() in existing_names:
                results.append({"name": name, "action": "skipped",
                                "reason": "already exists"})
                continue
            pid = db.add_person(
                name, p.get("kind", "fellow"), p.get("rotation", ""),
                int(p.get("min_clinics", 0) or 0), int(p.get("max_clinics", 5) or 5),
                p.get("start_date") or None, p.get("end_date") or None,
                p.get("notes", ""))
            if not p.get("active", 1):
                db.update_person(pid, active=0)
            for lv in p.get("leaves", []):
                db.add_leave(pid, lv["start_date"], lv["end_date"], lv.get("note", ""))
            for pin in p.get("pins", []):
                db.add_recurring_pin(pid, pin["day"], pin["session"], pin["consultant"])
            for ex in p.get("exclusions", []):
                db.add_recurring_exclusion(pid, ex["day"], ex["session"])
            results.append({"name": name, "action": "added"})
        added = sum(1 for r in results if r["action"] == "added")
        skipped = sum(1 for r in results if r["action"] == "skipped")
        return jsonify({"ok": True, "added": added, "skipped": skipped,
                        "details": results})
    except Exception as e:
        return jsonify({"error": traceback.format_exc(), "message": str(e)}), 500


@app.route("/api/people", methods=["POST"])
def api_add_person():
    db = get_db()
    data = request.get_json()
    try:
        pid = db.add_person(
            data["name"], data["kind"], data.get("rotation", ""),
            int(data.get("min_clinics", 0)), int(data.get("max_clinics", 5)),
            data.get("start_date") or None, data.get("end_date") or None,
            data.get("notes", ""))
        return jsonify({"id": pid, "ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/people/<int:pid>", methods=["PUT"])
def api_update_person(pid):
    db = get_db()
    data = request.get_json()
    db.update_person(pid, **data)
    return jsonify({"ok": True})


@app.route("/api/people/<int:pid>", methods=["DELETE"])
def api_delete_person(pid):
    db = get_db()
    db.delete_person(pid)
    return jsonify({"ok": True})


@app.route("/api/people/<int:pid>/leaves", methods=["GET"])
def api_get_leaves(pid):
    db = get_db()
    return jsonify(db.list_leaves(pid))


@app.route("/api/people/<int:pid>/leaves", methods=["POST"])
def api_add_leave(pid):
    db = get_db()
    data = request.get_json()
    lid = db.add_leave(pid, data["start_date"], data["end_date"],
                       data.get("note", ""))
    return jsonify({"id": lid, "ok": True})


@app.route("/api/leaves/<int:lid>", methods=["DELETE"])
def api_delete_leave(lid):
    db = get_db()
    db.delete_leave(lid)
    return jsonify({"ok": True})


@app.route("/api/people/<int:pid>/pins", methods=["GET"])
def api_get_pins(pid):
    db = get_db()
    return jsonify(db.list_recurring_pins(pid))


@app.route("/api/people/<int:pid>/pins", methods=["POST"])
def api_add_pin(pid):
    db = get_db()
    data = request.get_json()
    rid = db.add_recurring_pin(pid, data["day"], data["session"], data["consultant"])
    return jsonify({"id": rid, "ok": True})


@app.route("/api/pins/<int:rid>", methods=["DELETE"])
def api_delete_pin(rid):
    db = get_db()
    db.delete_recurring_pin(rid)
    return jsonify({"ok": True})


@app.route("/api/people/<int:pid>/exclusions", methods=["GET"])
def api_get_exclusions(pid):
    db = get_db()
    return jsonify(db.list_recurring_exclusions(pid))


@app.route("/api/people/<int:pid>/exclusions", methods=["POST"])
def api_add_exclusion(pid):
    db = get_db()
    data = request.get_json()
    eid = db.add_recurring_exclusion(pid, data["day"], data["session"])
    return jsonify({"id": eid, "ok": True})


@app.route("/api/exclusions/<int:eid>", methods=["DELETE"])
def api_delete_exclusion(eid):
    db = get_db()
    db.delete_recurring_exclusion(eid)
    return jsonify({"ok": True})


@app.route("/api/template", methods=["POST"])
def api_save_template():
    db = get_db()
    db.set_template(request.get_json())
    return jsonify({"ok": True})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
