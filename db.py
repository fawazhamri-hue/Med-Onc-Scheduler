"""
OncoScheduler v2 — persistence layer.

One SQLite file (oncoscheduler.db) next to the executable.
Backup = copy the file.

Tables
------
people                permanent roster (fellows, NPs, residents, assistants, externals)
leaves                date-range leaves per person (entered once, applied automatically)
recurring_pins        LTC-style standing purple pins (apply every week)
recurring_exclusions  standing availability blocks ("every Tue PM excused")
clinic_template       the standing weekly clinic grid (consultants rarely change)
weeks                 one row per scheduled week
week_clinics          this week's clinics + patient counts + AddOn flags
week_pins             one-off purple pins for a single week
week_exclusions       one-off leave/availability rows for a single week
settings              solver weights & config (editable without code change)
schedule_history      JSON snapshot of every generated schedule
"""

import json
import os
import sqlite3
from datetime import date, datetime, timedelta

import pandas as pd

DAYS = ["sun", "mon", "tue", "wed", "thu"]

SCHEMA = """
CREATE TABLE IF NOT EXISTS people (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE,
    kind TEXT NOT NULL CHECK (kind IN ('fellow','np','resident','assistant','external')),
    rotation TEXT DEFAULT '',          -- for fellows/NPs
    min_clinics INTEGER DEFAULT 0,
    max_clinics INTEGER DEFAULT 5,
    start_date TEXT,                   -- ISO date or NULL = always
    end_date TEXT,                     -- ISO date or NULL = always
    notes TEXT DEFAULT '',
    active INTEGER DEFAULT 1           -- soft delete
);
CREATE TABLE IF NOT EXISTS leaves (
    id INTEGER PRIMARY KEY,
    person_id INTEGER NOT NULL REFERENCES people(id) ON DELETE CASCADE,
    start_date TEXT NOT NULL,
    end_date TEXT NOT NULL,
    note TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS recurring_pins (
    id INTEGER PRIMARY KEY,
    person_id INTEGER NOT NULL REFERENCES people(id) ON DELETE CASCADE,
    day TEXT NOT NULL,
    session TEXT NOT NULL,
    consultant TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS recurring_exclusions (
    id INTEGER PRIMARY KEY,
    person_id INTEGER NOT NULL REFERENCES people(id) ON DELETE CASCADE,
    day TEXT NOT NULL,
    session TEXT NOT NULL              -- AM / PM / ALL
);
CREATE TABLE IF NOT EXISTS clinic_template (
    id INTEGER PRIMARY KEY,
    day TEXT NOT NULL,
    session TEXT NOT NULL,
    consultant TEXT NOT NULL,
    specialty TEXT NOT NULL,
    clinic_type TEXT DEFAULT ''        -- '' | 'AddOn'
);
CREATE TABLE IF NOT EXISTS weeks (
    id INTEGER PRIMARY KEY,
    start_date TEXT NOT NULL UNIQUE,   -- the Sunday
    end_date TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS week_clinics (
    id INTEGER PRIMARY KEY,
    week_id INTEGER NOT NULL REFERENCES weeks(id) ON DELETE CASCADE,
    day TEXT NOT NULL,
    session TEXT NOT NULL,
    consultant TEXT NOT NULL,
    specialty TEXT NOT NULL,
    clinic_type TEXT DEFAULT '',
    patients INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS week_pins (
    id INTEGER PRIMARY KEY,
    week_id INTEGER NOT NULL REFERENCES weeks(id) ON DELETE CASCADE,
    person_name TEXT NOT NULL,         -- name, not id: allows externals not in people
    day TEXT NOT NULL,
    session TEXT NOT NULL,
    consultant TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS week_exclusions (
    id INTEGER PRIMARY KEY,
    week_id INTEGER NOT NULL REFERENCES weeks(id) ON DELETE CASCADE,
    person_name TEXT NOT NULL,
    day TEXT NOT NULL,
    session TEXT NOT NULL,
    reason TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS schedule_history (
    id INTEGER PRIMARY KEY,
    week_id INTEGER NOT NULL REFERENCES weeks(id) ON DELETE CASCADE,
    generated_at TEXT NOT NULL,
    result_json TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS favor_log (
    id INTEGER PRIMARY KEY,
    entry_date TEXT NOT NULL,          -- date the favor happened (editable)
    person_name TEXT NOT NULL,         -- name, not id: survives even if person later deleted
    description TEXT NOT NULL,         -- free text: what happened
    tag TEXT DEFAULT '',               -- optional: Clinic coverage / Inpatient help / Other
    compensated INTEGER DEFAULT 0,     -- 0 = outstanding, 1 = compensated
    compensation_note TEXT DEFAULT '', -- how/when it was paid back
    compensated_at TEXT DEFAULT '',    -- auto-set the moment it's toggled compensated
    created_at TEXT NOT NULL           -- when the log entry itself was added
);
"""

# Default solver settings exposed in the GUI (Phase B fills the editor;
# the table exists from day one so weights live in data, not code).
DEFAULT_SETTINGS = {
    "w_chemo_shortfall": "10000",
    "w_addon_uncovered": "10000",
    "w_fellow_short": "1500",
    "w_capacity_short": "250",
    "w_assist_for_chemo": "250",
    "w_under_min_res": "200",
    "w_rotation_fairness": "200",
    "w_under_min_fellow": "150",
    "w_over_assign": "110",
    "w_rotation_match": "100",
    "w_fairness": "35",
    "w_band_pref": "25",
    "chemo_active": "1",
    "chemo_slots": "sun:PM, mon:AM, tue:AM, wed:AM, thu:AM",
    "solver_timeout_s": "30",
}


def _iso(d):
    if d is None or d == "":
        return None
    if isinstance(d, (date, datetime)):
        return d.strftime("%Y-%m-%d")
    return str(d)[:10]


def _parse(d):
    if d is None or d == "":
        return None
    return datetime.strptime(str(d)[:10], "%Y-%m-%d").date()


class Database:
    def __init__(self, path="oncoscheduler.db"):
        self.path = path
        self.con = sqlite3.connect(path, check_same_thread=False)
        self.con.row_factory = sqlite3.Row
        self.con.execute("PRAGMA foreign_keys = ON")
        self.con.executescript(SCHEMA)
        for k, v in DEFAULT_SETTINGS.items():
            self.con.execute(
                "INSERT OR IGNORE INTO settings (key, value) VALUES (?,?)", (k, v))
        self.con.commit()

    # ---------------- people ----------------
    def add_person(self, name, kind, rotation="", min_clinics=0, max_clinics=5,
                   start_date=None, end_date=None, notes=""):
        cur = self.con.execute(
            "INSERT INTO people (name, kind, rotation, min_clinics, max_clinics,"
            " start_date, end_date, notes) VALUES (?,?,?,?,?,?,?,?)",
            (name.strip(), kind, rotation.strip(), int(min_clinics),
             int(max_clinics), _iso(start_date), _iso(end_date), notes))
        self.con.commit()
        return cur.lastrowid

    def update_person(self, pid, **fields):
        allowed = {"name", "kind", "rotation", "min_clinics", "max_clinics",
                   "start_date", "end_date", "notes", "active"}
        sets, vals = [], []
        for k, v in fields.items():
            if k not in allowed:
                continue
            if k in ("start_date", "end_date"):
                v = _iso(v)
            sets.append(f"{k}=?")
            vals.append(v)
        if not sets:
            return
        vals.append(pid)
        self.con.execute(f"UPDATE people SET {', '.join(sets)} WHERE id=?", vals)
        self.con.commit()

    def delete_person(self, pid):
        self.con.execute("DELETE FROM people WHERE id=?", (pid,))
        self.con.commit()

    def get_person_by_name(self, name):
        row = self.con.execute(
            "SELECT * FROM people WHERE lower(name)=lower(?)", (name.strip(),)).fetchone()
        return dict(row) if row else None

    def list_people(self, kind=None, include_inactive=False):
        q = "SELECT * FROM people"
        conds, vals = [], []
        if kind:
            conds.append("kind=?")
            vals.append(kind)
        if not include_inactive:
            conds.append("active=1")
        if conds:
            q += " WHERE " + " AND ".join(conds)
        q += " ORDER BY kind, name"
        return [dict(r) for r in self.con.execute(q, vals)]

    # ---------------- leaves ----------------
    def add_leave(self, person_id, start_date, end_date, note=""):
        cur = self.con.execute(
            "INSERT INTO leaves (person_id, start_date, end_date, note) VALUES (?,?,?,?)",
            (person_id, _iso(start_date), _iso(end_date), note))
        self.con.commit()
        return cur.lastrowid

    def delete_leave(self, leave_id):
        self.con.execute("DELETE FROM leaves WHERE id=?", (leave_id,))
        self.con.commit()

    def list_leaves(self, person_id=None):
        if person_id:
            rows = self.con.execute(
                "SELECT l.*, p.name FROM leaves l JOIN people p ON p.id=l.person_id"
                " WHERE person_id=? ORDER BY start_date", (person_id,))
        else:
            rows = self.con.execute(
                "SELECT l.*, p.name FROM leaves l JOIN people p ON p.id=l.person_id"
                " ORDER BY start_date")
        return [dict(r) for r in rows]

    # ---------------- recurring pins / exclusions ----------------
    def add_recurring_pin(self, person_id, day, session, consultant):
        cur = self.con.execute(
            "INSERT INTO recurring_pins (person_id, day, session, consultant) VALUES (?,?,?,?)",
            (person_id, day.lower()[:3], session.upper(), consultant.strip()))
        self.con.commit()
        return cur.lastrowid

    def delete_recurring_pin(self, pin_id):
        self.con.execute("DELETE FROM recurring_pins WHERE id=?", (pin_id,))
        self.con.commit()

    def list_recurring_pins(self, person_id=None):
        q = ("SELECT rp.*, p.name FROM recurring_pins rp"
             " JOIN people p ON p.id = rp.person_id")
        if person_id:
            rows = self.con.execute(q + " WHERE rp.person_id=?", (person_id,))
        else:
            rows = self.con.execute(q)
        return [dict(r) for r in rows]

    def add_recurring_exclusion(self, person_id, day, session):
        cur = self.con.execute(
            "INSERT INTO recurring_exclusions (person_id, day, session) VALUES (?,?,?)",
            (person_id, day.lower()[:3], session.upper()))
        self.con.commit()
        return cur.lastrowid

    def delete_recurring_exclusion(self, ex_id):
        self.con.execute("DELETE FROM recurring_exclusions WHERE id=?", (ex_id,))
        self.con.commit()

    def list_recurring_exclusions(self, person_id=None):
        q = ("SELECT re.*, p.name FROM recurring_exclusions re"
             " JOIN people p ON p.id = re.person_id")
        if person_id:
            rows = self.con.execute(q + " WHERE re.person_id=?", (person_id,))
        else:
            rows = self.con.execute(q)
        return [dict(r) for r in rows]

    # ---------------- clinic template ----------------
    def set_template(self, rows):
        """rows: list of dicts (day, session, consultant, specialty, clinic_type)."""
        self.con.execute("DELETE FROM clinic_template")
        for r in rows:
            self.con.execute(
                "INSERT INTO clinic_template (day, session, consultant, specialty, clinic_type)"
                " VALUES (?,?,?,?,?)",
                (r["day"].lower()[:3], r["session"].upper(), r["consultant"].strip(),
                 r["specialty"].strip(), r.get("clinic_type", "")))
        self.con.commit()

    def get_template(self):
        return [dict(r) for r in self.con.execute(
            "SELECT * FROM clinic_template ORDER BY id")]

    # ---------------- weeks ----------------
    def create_week(self, start_date):
        sd = _parse(_iso(start_date))
        ed = sd + timedelta(days=4)  # Sun..Thu
        row = self.con.execute(
            "SELECT id FROM weeks WHERE start_date=?", (_iso(sd),)).fetchone()
        if row:
            return row["id"]
        cur = self.con.execute(
            "INSERT INTO weeks (start_date, end_date) VALUES (?,?)",
            (_iso(sd), _iso(ed)))
        self.con.commit()
        wid = cur.lastrowid
        # Seed week_clinics from template
        for t in self.get_template():
            self.con.execute(
                "INSERT INTO week_clinics (week_id, day, session, consultant, specialty,"
                " clinic_type, patients) VALUES (?,?,?,?,?,?,0)",
                (wid, t["day"], t["session"], t["consultant"], t["specialty"],
                 t["clinic_type"]))
        self.con.commit()
        return wid

    def get_week(self, week_id):
        row = self.con.execute("SELECT * FROM weeks WHERE id=?", (week_id,)).fetchone()
        return dict(row) if row else None

    def list_weeks(self):
        return [dict(r) for r in self.con.execute(
            "SELECT * FROM weeks ORDER BY start_date DESC")]

    # ---------------- week clinics ----------------
    def list_week_clinics(self, week_id):
        return [dict(r) for r in self.con.execute(
            "SELECT * FROM week_clinics WHERE week_id=? ORDER BY id", (week_id,))]

    def add_week_clinic(self, week_id, day, session, consultant, specialty,
                        clinic_type="", patients=0):
        cur = self.con.execute(
            "INSERT INTO week_clinics (week_id, day, session, consultant, specialty,"
            " clinic_type, patients) VALUES (?,?,?,?,?,?,?)",
            (week_id, day.lower()[:3], session.upper(), consultant.strip(),
             specialty.strip(), clinic_type, int(patients)))
        self.con.commit()
        return cur.lastrowid

    def update_week_clinic(self, clinic_id, **fields):
        allowed = {"day", "session", "consultant", "specialty", "clinic_type", "patients"}
        sets, vals = [], []
        for k, v in fields.items():
            if k in allowed:
                sets.append(f"{k}=?")
                vals.append(v)
        if not sets:
            return
        vals.append(clinic_id)
        self.con.execute(f"UPDATE week_clinics SET {', '.join(sets)} WHERE id=?", vals)
        self.con.commit()

    def delete_week_clinic(self, clinic_id):
        self.con.execute("DELETE FROM week_clinics WHERE id=?", (clinic_id,))
        self.con.commit()

    # ---------------- week pins / exclusions ----------------
    def add_week_pin(self, week_id, person_name, day, session, consultant):
        cur = self.con.execute(
            "INSERT INTO week_pins (week_id, person_name, day, session, consultant)"
            " VALUES (?,?,?,?,?)",
            (week_id, person_name.strip(), day.lower()[:3], session.upper(),
             consultant.strip()))
        self.con.commit()
        return cur.lastrowid

    def delete_week_pin(self, pin_id):
        self.con.execute("DELETE FROM week_pins WHERE id=?", (pin_id,))
        self.con.commit()

    def list_week_pins(self, week_id):
        return [dict(r) for r in self.con.execute(
            "SELECT * FROM week_pins WHERE week_id=?", (week_id,))]

    def add_week_exclusion(self, week_id, person_name, day, session, reason=""):
        cur = self.con.execute(
            "INSERT INTO week_exclusions (week_id, person_name, day, session, reason)"
            " VALUES (?,?,?,?,?)",
            (week_id, person_name.strip(), day.lower()[:3], session.upper(), reason))
        self.con.commit()
        return cur.lastrowid

    def delete_week_exclusion(self, ex_id):
        self.con.execute("DELETE FROM week_exclusions WHERE id=?", (ex_id,))
        self.con.commit()

    def list_week_exclusions(self, week_id):
        return [dict(r) for r in self.con.execute(
            "SELECT * FROM week_exclusions WHERE week_id=?", (week_id,))]

    # ---------------- settings ----------------
    def get_setting(self, key, default=None):
        row = self.con.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default

    def set_setting(self, key, value):
        self.con.execute(
            "INSERT INTO settings (key, value) VALUES (?,?)"
            " ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, str(value)))
        self.con.commit()

    def all_settings(self):
        return {r["key"]: r["value"] for r in self.con.execute("SELECT * FROM settings")}

    # ---------------- history ----------------
    def save_history(self, week_id, result):
        slim = {k: result.get(k) for k in
                ("status", "schedule", "fellows_summary", "residents_summary",
                 "assistants_summary", "on_leave", "availability_orphans",
                 "purple_orphans_leave")}
        self.con.execute(
            "INSERT INTO schedule_history (week_id, generated_at, result_json)"
            " VALUES (?,?,?)",
            (week_id, datetime.now().isoformat(timespec="seconds"),
             json.dumps(slim, default=str)))
        self.con.commit()

    def list_history(self, week_id=None):
        if week_id:
            rows = self.con.execute(
                "SELECT id, week_id, generated_at FROM schedule_history"
                " WHERE week_id=? ORDER BY generated_at DESC", (week_id,))
        else:
            rows = self.con.execute(
                "SELECT id, week_id, generated_at FROM schedule_history"
                " ORDER BY generated_at DESC")
        return [dict(r) for r in rows]

    # =====================================================================
    # Favor log: unofficial coverage, inpatient help, ad-hoc favors the chief
    # asks fellows for and needs to remember whether he's paid back.
    # =====================================================================
    def add_favor(self, entry_date, person_name, description, tag=""):
        cur = self.con.execute(
            "INSERT INTO favor_log (entry_date, person_name, description, tag,"
            " compensated, compensation_note, compensated_at, created_at)"
            " VALUES (?,?,?,?,0,'','',?)",
            (entry_date, person_name.strip(), description.strip(), tag,
             datetime.now().isoformat(timespec="seconds")))
        self.con.commit()
        return cur.lastrowid

    def list_favors(self, person_name=None, outstanding_only=False):
        q = "SELECT * FROM favor_log"
        conds, params = [], []
        if person_name:
            conds.append("person_name=?")
            params.append(person_name)
        if outstanding_only:
            conds.append("compensated=0")
        if conds:
            q += " WHERE " + " AND ".join(conds)
        q += " ORDER BY entry_date DESC, id DESC"
        return [dict(r) for r in self.con.execute(q, params)]

    def get_favor(self, favor_id):
        row = self.con.execute(
            "SELECT * FROM favor_log WHERE id=?", (favor_id,)).fetchone()
        return dict(row) if row else None

    def update_favor(self, favor_id, **fields):
        allowed = {"entry_date", "person_name", "description", "tag",
                  "compensated", "compensation_note"}
        sets, vals = [], []
        for k, v in fields.items():
            if k in allowed:
                sets.append(f"{k}=?")
                vals.append(v)
        if not sets:
            return
        # Auto-timestamp the moment something is marked compensated.
        if fields.get("compensated") == 1:
            sets.append("compensated_at=?")
            vals.append(datetime.now().isoformat(timespec="seconds"))
        elif fields.get("compensated") == 0:
            sets.append("compensated_at=?")
            vals.append("")
        vals.append(favor_id)
        self.con.execute(f"UPDATE favor_log SET {', '.join(sets)} WHERE id=?", vals)
        self.con.commit()

    def delete_favor(self, favor_id):
        self.con.execute("DELETE FROM favor_log WHERE id=?", (favor_id,))
        self.con.commit()

    # =====================================================================
    # The bridge: assemble solver inputs for a given week from the database.
    # Produces DataFrames matching the v1 xlsx schemas exactly, so the v12.2
    # solver core runs unchanged.
    # =====================================================================
    def assemble_week(self, week_id):
        wk = self.get_week(week_id)
        if not wk:
            raise ValueError(f"No week with id {week_id}")
        wstart = _parse(wk["start_date"])
        day_dates = {DAYS[i]: wstart + timedelta(days=i) for i in range(5)}

        def is_active(p):
            sd = _parse(p["start_date"])
            ed = _parse(p["end_date"])
            if sd and sd > day_dates["thu"]:
                return False
            if ed and ed < wstart:
                return False
            return bool(p["active"])

        people = [p for p in self.list_people() if is_active(p)]
        by_id = {p["id"]: p for p in people}

        # Leave resolution: full-week leave -> Status=Leave;
        # partial-week -> per-day exclusions.
        leave_status = {}        # person name -> True if whole week
        partial_excl = []        # rows for availability df
        for lv in self.list_leaves():
            p = by_id.get(lv["person_id"])
            if not p:
                continue
            lsd, led = _parse(lv["start_date"]), _parse(lv["end_date"])
            covered = [d for d in DAYS if lsd <= day_dates[d] <= led]
            if not covered:
                continue
            if len(covered) == 5:
                leave_status[p["name"]] = True
            else:
                for d in covered:
                    partial_excl.append({"Person": p["name"], "Day": d,
                                         "Session": "All",
                                         "Reason": lv["note"] or "leave"})

        # fellows_df: fellows + NPs
        fellows_rows = []
        for p in people:
            if p["kind"] not in ("fellow", "np"):
                continue
            fellows_rows.append({
                "Fellow": p["name"], "Rotation": p["rotation"],
                "Min Clinics": p["min_clinics"], "Max Clinics": p["max_clinics"],
                "Role": "NP" if p["kind"] == "np" else "Fellow",
                "Status": "Leave" if leave_status.get(p["name"]) else "Active",
            })
        fellows_df = pd.DataFrame(fellows_rows)

        # residents_df: residents + assistants
        res_rows = []
        for p in people:
            if p["kind"] not in ("resident", "assistant"):
                continue
            res_rows.append({
                "Resident": p["name"],
                "Min Clinics": p["min_clinics"], "Max Clinics": p["max_clinics"],
                "Type": "Assistant" if p["kind"] == "assistant" else "Resident",
                "Status": "Leave" if leave_status.get(p["name"]) else "Active",
            })
        residents_df = pd.DataFrame(res_rows)

        # clinics_df
        cl_rows = []
        for c in self.list_week_clinics(week_id):
            cl_rows.append({
                "Day": c["day"].capitalize(), "Session": c["session"],
                "Consultant": c["consultant"], "Specialty": c["specialty"],
                "ClinicType": c["clinic_type"] or None,
                "Patients": c["patients"],
            })
        clinics_df = pd.DataFrame(cl_rows)

        # purple_df: recurring pins (active people only) + one-off pins
        purple_rows = []
        for rp in self.list_recurring_pins():
            p = by_id.get(rp["person_id"])
            if not p:
                continue  # person inactive/out of window this week
            purple_rows.append({"Fellow": p["name"], "Day": rp["day"],
                                "Session": rp["session"], "Consultant": rp["consultant"]})
        for wp in self.list_week_pins(week_id):
            purple_rows.append({"Fellow": wp["person_name"], "Day": wp["day"],
                                "Session": wp["session"], "Consultant": wp["consultant"]})
        purple_df = pd.DataFrame(purple_rows) if purple_rows else pd.DataFrame(
            columns=["Fellow", "Day", "Session", "Consultant"])

        # availability_df: recurring exclusions + partial leaves + one-off exclusions
        av_rows = []
        for re_ in self.list_recurring_exclusions():
            p = by_id.get(re_["person_id"])
            if not p:
                continue
            av_rows.append({"Person": p["name"], "Day": re_["day"],
                            "Session": re_["session"], "Reason": "recurring"})
        av_rows.extend(partial_excl)
        for we in self.list_week_exclusions(week_id):
            av_rows.append({"Person": we["person_name"], "Day": we["day"],
                            "Session": we["session"], "Reason": we["reason"]})
        availability_df = pd.DataFrame(av_rows) if av_rows else None

        date_map = {d.capitalize(): day_dates[d].strftime("%d %b") for d in DAYS}
        return {
            "fellows_df": fellows_df,
            "residents_df": residents_df,
            "clinics_df": clinics_df,
            "purple_df": purple_df,
            "availability_df": availability_df,
            "date_map": date_map,
            "week": wk,
        }
