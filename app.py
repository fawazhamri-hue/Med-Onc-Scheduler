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
from datetime import datetime, timedelta

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


def _build_cfg(db):
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
        fellows_df, residents_df, clinics_df, pins_df, availability_df, date_map = \
            _assemble_from_payload(payload)
        cfg = _build_cfg(db)
        result = solve(clinics_df, fellows_df, residents_df, pins_df,
                       config=cfg, availability_path=availability_df)
        return jsonify({
            "status": result["status"],
            "schedule": result.get("schedule", []),
            "fellows_summary": result.get("fellows_summary", []),
            "residents_summary": result.get("residents_summary", []),
            "on_leave": result.get("on_leave", []),
            "purple_orphans_leave": result.get("purple_orphans_leave", []),
            "availability_orphans": result.get("availability_orphans", []),
            "date_map": date_map,
        })
    except Exception as e:
        return jsonify({"error": traceback.format_exc(), "message": str(e)}), 500


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
        _, _, _, _, _, date_map = _assemble_from_payload(payload)
        week_start = payload.get("week_start", "")
        week_end = payload.get("week_end", "")

        result = {**result_data, "availability": [], "availability_orphans": [],
                  "purple_orphans_leave": [], "on_leave": []}

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
