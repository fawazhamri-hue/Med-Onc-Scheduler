"""
Solver core for the Medical Oncology weekly scheduler.
Importable; takes paths in, returns a result dict.
"""

from ortools.sat.python import cp_model
import pandas as pd
import re
import os
from dataclasses import dataclass, field
from typing import Optional


# =====================================================
# CONFIG OBJECT (defaults — override via SolverConfig)
# =====================================================
@dataclass
class SolverConfig:
    # Weights
    # Penalty weights — ORDER MATTERS.
    # Hard rule: chemo coverage must NEVER lose to any other soft trade-off.
    # w_chemo_shortfall is set high enough to dominate the AGGREGATE of plausible
    # competing penalties, not just each individually. With cap_short at 250/patient
    # on clinics up to ~26 patients, a single capacity gap can hit ~1,500; multiple
    # gaps compound. 10,000 keeps chemo protected against any realistic aggregate.
    # (Soft-not-hard so the solver still returns SOME schedule under severe
    # undersupply, with a CHEMO UNCOVERED warning rather than a silent INFEASIBLE.)
    w_rotation_match:       int = 100
    w_shortfall:            int = 5     # multiplied by patient count for severity weighting
    w_chemo_shortfall:      int = 10000 # MUST exceed AGGREGATE of all other penalties below
    # w_fellow_short — penalty per missing fellow on a 17+ pt clinic.
    # Bumped from 400 in v12 after the Wed AM Sabah case (22pt, 0 fellows) came back
    # OPTIMAL at the lower weight — the math didn't motivate the solver to pull a fellow
    # in. At 1500 it dominates the typical "cost of pulling a fellow from elsewhere"
    # (over_assign 110 + small band_pref deviations) by enough to drive redistribution.
    # Still below w_chemo_shortfall and w_addon_uncovered (both 10000).
    w_fellow_short:         int = 1500
    w_capacity_short:       int = 250   # missing patient capacity (per patient unit short)
    w_over_assign:          int = 110   # > rotation_match: stops solver from piling fellows for free rotation points
    w_over_excess:          int = 200   # 2+ helpers beyond required — kept BELOW w_chemo_shortfall
    w_under_min_fellow:     int = 150
    w_under_min_res:        int = 200   # residents/NPs hit their min; still < chemo
    w_imbalance:            int = 10    # legacy: kept for backwards compat in diagnostics
    w_fairness:             int = 35    # per clinic above min, per person — pushes balanced distribution
    # Per-rotation in-rotation fairness — penalizes spread (max - min) of in-rotation
    # clinic counts among fellows of the SAME rotation. Without this term, the objective
    # rewards total rotation matches but not their distribution, leading to "Maryam 5,
    # Othman 3, Fawaz 0" patterns where one fellow soaks up rotation matches while a
    # rotation-mate gets pulled entirely out of rotation due to availability quirks.
    # Bumped from 50 to 200 in v12: at 50 the term was dominated by over_assign (110),
    # so the solver wouldn't add a fellow to a small (0-required) in-rotation clinic
    # even when doing so saved 50 in spread. At 200, a spread reduction from 5 to 4
    # saves 200 — comfortably > over_assign (110), so the swap is preferred. Still well
    # below w_fellow_short (1500) so it can't override a genuine fellow-coverage need.
    w_rotation_fairness:    int = 200
    # True workload balancing among fellows. This directly penalizes the gap
    # between the busiest and least busy fellow, unlike total-over-min which can
    # be identical for 7/5/5 and 6/6/5 distributions.
    w_workload_spread:      int = 350
    # A sixth clinic is acceptable. A seventh should be used only when it
    # prevents a more important coverage problem.
    w_seventh_clinic:       int = 1200
    # Soft educational target for fellows. Kept soft because some weeks do not
    # contain enough clinics in every rotation, but missing the target should be
    # expensive enough to affect the schedule materially.
    w_in_rotation_short:    int = 900
    preferred_in_rotation_min: int = 3
    w_double_session:       int = 2
    w_assist_for_chemo:     int = 250   # makes assistants preferred over fellows on chemo;
                                        # must exceed w_rotation_match (100) so the solver
                                        # picks an assistant on chemo even when a fellow
                                        # could earn rotation-match elsewhere
    # Repeated chemoassessment is allowed when necessary, but more than two
    # sessions per fellow should be increasingly unattractive.
    w_chemo_repeat:         int = 900
    preferred_chemo_max:    int = 2
    # Add-ons prefer a fellow from the same rotation. A cross-rotation fellow or
    # assistant is a fallback rather than forbidden. This penalty is lower than
    # the seventh-clinic penalty, so the solver normally uses the matching fellow
    # up to six clinics, then considers fallback coverage instead of forcing seven.
    w_cross_rotation_addon: int = 650
    w_band_pref:            int = 25   # soft preference penalty per band-rule deviation
    # ADD-ON clinics (consultant on leave; ~8-10 patients covered solo by a fellow or assistant).
    # Per chief: leaving an add-on unfilled is as unacceptable as leaving chemo unfilled,
    # so this weight matches w_chemo_shortfall (both must dominate aggregate competing penalties).
    w_addon_uncovered:      int = 10000
    max_double_days:        int = 3

    # --- Capacity model ---
    consultant_capacity: int = 11   # patients a consultant alone can handle
    fellow_capacity:     int = 6    # additional patients per fellow / assistant
    resident_capacity:   int = 5    # additional patients per resident
    np_capacity:         int = 5    # additional patients per nurse practitioner
    educational_floor:   int = 9    # min patients triggering at least 1 helper
    two_fellow_threshold: int = 24  # patients at/above which 2 fellows required

    # Per-person min overrides (rarely needed; usually set in fellows.xlsx directly)
    min_overrides: dict = field(default_factory=lambda: {})

    # Chemoassessment
    chemo_active: bool = True
    chemo_slots:  list = field(default_factory=lambda: [
        ("sun", "PM"), ("mon", "AM"), ("tue", "AM"),
        ("wed", "AM"), ("thu", "AM"),
    ])

    solver_timeout_sec: int = 90    # bumped from 30s in v12 after add-on + validation
                                    # constraints grew the search space; 30s was producing
                                    # tie-break suboptimality on dense weeks.


# =====================================================
# HELPERS
# =====================================================
def parse_patients(value):
    if pd.isna(value):
        return 0
    s = str(value).strip()
    m = re.search(r"(\d+)\s*[-–to]+\s*(\d+)", s)
    if m:
        return (int(m.group(1)) + int(m.group(2))) // 2
    m = re.search(r"\d+", s)
    return int(m.group()) if m else 0


def compute_requirements(patients, cfg):
    """Patient-count band requirements (hard minima).

    Returns dict:
      min_total       : minimum total helpers required (HARD)
      min_fellows     : minimum fellow-equivalents required (HARD; chemo-write safety)
      patient_cap_req : extra patient capacity helpers must provide (HARD; safety)

    Soft preferences for HOW the helpers are composed within these floors
    (e.g., "12–14 prefers resident/NP not fellow") are encoded as objective
    penalties in the solver, not here.

    Patient-count bands (12–14, 15–16, 17–19, 20–23, 24–26, 27+) are derived
    from the chief fellow's clinical judgment, not pure capacity math.
    """
    if patients < 12:
        # 0–11: consultant alone
        return {"min_total": 0, "min_fellows": 0, "patient_cap_req": 0}

    if patients <= 14:
        # 12–14: 1 helper (any), prefer NP/resident (soft)
        return {"min_total": 1, "min_fellows": 0,
                "patient_cap_req": max(0, patients - cfg.consultant_capacity)}

    if patients <= 16:
        # 15–16: 1 fellow OR 2 of anything; min_total=1 satisfies the lower bound
        # (the OR-of-2-anything is handled by the soft preference layer)
        return {"min_total": 1, "min_fellows": 0,
                "patient_cap_req": max(0, patients - cfg.consultant_capacity)}

    if patients <= 19:
        # 17–19: 1 fellow + 1 non-fellow (2 helpers, ≥1 fellow)
        # The "non-fellow" preference is soft; HARD floor is 2 helpers, ≥1 fellow.
        return {"min_total": 2, "min_fellows": 1,
                "patient_cap_req": patients - cfg.consultant_capacity}

    if patients <= 23:
        # 20–23: 2 fellows alone OR 1 fellow + 2 non-fellows
        # HARD floor: 2 helpers, ≥1 fellow. Soft preference for the patterns.
        return {"min_total": 2, "min_fellows": 1,
                "patient_cap_req": patients - cfg.consultant_capacity}

    if patients <= 26:
        # 24–26: 2 fellows + 1 non-fellow (3 helpers, ≥2 fellows)
        return {"min_total": 3, "min_fellows": 2,
                "patient_cap_req": patients - cfg.consultant_capacity}

    # 27+: 3 fellows OR 2 fellows + 2 non-fellows (≥3 helpers, ≥2 fellows)
    return {"min_total": 3, "min_fellows": 2,
            "patient_cap_req": patients - cfg.consultant_capacity}


# =====================================================
# MAIN SOLVE
# =====================================================
def solve(clinics_path, fellows_path, residents_path, purple_path,
          config: Optional[SolverConfig] = None,
          availability_path: Optional[str] = None):
    """
    Returns a dict with:
      status            : "OPTIMAL", "FEASIBLE", "INFEASIBLE", "UNKNOWN"
      schedule          : list of dicts (Day/Session/Consultant/Specialty/Patients/Required/Assigned)
      fellows_summary   : list of dicts
      residents_summary : list of dicts
      assistants_summary: list of dicts (may be empty)
      diagnostics       : dict (rotation_matches, shortfalls, under_min, spread, demand, capacity)
      message           : str (human-readable status note)
      availability      : list of (person, day, session, reason) entries applied

    availability_path is optional. If provided AND the file exists, persons listed
    are excluded from the indicated day/session combinations.
    """
    cfg = config or SolverConfig()

    # v2: each *_path may be a filesystem path OR an already-built DataFrame.
    def _as_df(x):
        return x if isinstance(x, pd.DataFrame) else pd.read_excel(x)
    clinics_df   = _as_df(clinics_path)
    fellows_df   = _as_df(fellows_path)
    residents_df = _as_df(residents_path)
    purple_df    = _as_df(purple_path)

    # --- Build people list ---
    # v12.2: Status column (optional). Values: Active (default) / Leave.
    # If Leave, the person is excluded from the solver entirely. Their rows in
    # availability.xlsx are silently skipped (no orphan warning). Their rows in
    # purple_clinics.xlsx still raise an error — pinning + leave is a real intent
    # conflict, not stale data.
    on_leave = []  # list of names marked Leave this week (for Notes sheet)
    fellows = []
    for idx, row in fellows_df.iterrows():
        name = str(row.get("Fellow", "")).strip()
        if not name or name.lower() == "nan":
            continue  # skip blank/ghost rows
        # Check Status column (v12.2)
        if "Status" in fellows_df.columns and pd.notna(row.get("Status")):
            status_val = str(row["Status"]).strip().lower()
            if status_val in ("leave", "off", "absent", "on leave"):
                on_leave.append(name)
                continue  # skip — person excluded from solver
            elif status_val not in ("active", ""):
                raise ValueError(
                    f"fellows.xlsx row {idx + 2}: '{name}' has unknown Status "
                    f"'{row['Status']}'. Valid values: Active, Leave."
                )
        role = "Fellow"
        if "Role" in fellows_df.columns and pd.notna(row.get("Role")):
            role = str(row["Role"]).strip().capitalize()
        try:
            min_c = int(row["Min Clinics"]) if pd.notna(row["Min Clinics"]) else 0
            max_c = int(row["Max Clinics"]) if pd.notna(row["Max Clinics"]) else 0
        except (ValueError, TypeError):
            raise ValueError(
                f"fellows.xlsx row {idx + 2}: '{name}' has invalid "
                f"Min/Max Clinics values. Both must be whole numbers."
            )
        if max_c == 0:
            raise ValueError(
                f"fellows.xlsx row {idx + 2}: '{name}' has Max Clinics blank or 0. "
                f"This person cannot be assigned to anything. Set Max >= Min, "
                f"set Status=Leave if they are off this week, "
                f"or remove the row if they are not on the team."
            )
        if max_c < min_c:
            raise ValueError(
                f"fellows.xlsx row {idx + 2}: '{name}' has Max ({max_c}) < Min ({min_c}). "
                f"Max must be >= Min."
            )
        rotation = str(row.get("Rotation", "")).strip()
        if not rotation or rotation.lower() == "nan":
            raise ValueError(f"fellows.xlsx row {idx + 2}: '{name}' is missing Rotation.")
        fellows.append({
            "name": name,
            "rotation": rotation,
            "min": min_c,
            "max": max_c,
            "is_resident": False,
            "is_assistant": False,
            "is_np": role.upper() == "NP",
            "role": role.lower(),
        })
    for idx, row in residents_df.iterrows():
        name = str(row.get("Resident", "")).strip()
        if not name or name.lower() == "nan":
            continue  # skip blank/ghost rows
        # Check Status column (v12.2)
        if "Status" in residents_df.columns and pd.notna(row.get("Status")):
            status_val = str(row["Status"]).strip().lower()
            if status_val in ("leave", "off", "absent", "on leave"):
                on_leave.append(name)
                continue  # skip — person excluded from solver
            elif status_val not in ("active", ""):
                raise ValueError(
                    f"residents.xlsx row {idx + 2}: '{name}' has unknown Status "
                    f"'{row['Status']}'. Valid values: Active, Leave."
                )
        ptype = "Resident"
        if "Type" in residents_df.columns and pd.notna(row.get("Type")):
            ptype = str(row["Type"]).strip().capitalize()
        try:
            min_c = int(row["Min Clinics"]) if pd.notna(row["Min Clinics"]) else 0
            max_c = int(row["Max Clinics"]) if pd.notna(row["Max Clinics"]) else 0
        except (ValueError, TypeError):
            raise ValueError(
                f"residents.xlsx row {idx + 2}: '{name}' has invalid "
                f"Min/Max Clinics values. Both must be whole numbers."
            )
        if max_c == 0:
            raise ValueError(
                f"residents.xlsx row {idx + 2}: '{name}' has Max Clinics blank or 0. "
                f"This person cannot be assigned to anything. Set Max >= Min, "
                f"set Status=Leave if they are off this week, "
                f"or remove the row if they are not on the team."
            )
        if max_c < min_c:
            raise ValueError(
                f"residents.xlsx row {idx + 2}: '{name}' has Max ({max_c}) < Min ({min_c}). "
                f"Max must be >= Min."
            )
        fellows.append({
            "name": name,
            "rotation": "Any",
            "min": min_c,
            "max": max_c,
            "is_resident": ptype == "Resident",
            "is_assistant": ptype == "Assistant",
            "is_np": False,
            "role": ptype.lower(),
        })
    fellow_names = [f["name"] for f in fellows]
    for f in fellows:
        if f["name"] in cfg.min_overrides:
            f["min"] = cfg.min_overrides[f["name"]]

    # --- Availability (optional file) ---
    # Reads availability_path xlsx if provided + exists.
    # Each row: Person | Day | Session(AM|PM|All) | Reason(optional)
    # Person is excluded from any clinic at the listed day/session.
    #
    # v12.2: orphan handling. If 'Person' is not in this week's roster:
    #   - If they're on the 'on_leave' list (Status=Leave in fellows/residents):
    #     silently skip — redundant info, already handled.
    #   - Otherwise: skip the row, append to availability_orphans for the Notes
    #     sheet. Used to be a hard ValueError, but that blocked weekly schedules
    #     because availability.xlsx accumulates entries while the roster changes.
    availability = []  # list of dicts: {person, day, session, reason}
    availability_orphans = []  # list of dicts: {person, day, session, reason} - unrecognized names
    avail_lookup = set()  # set of (person_name, day_lowercase, session_uppercase) for fast O(1) check
    name_lookup = {n.lower(): n for n in fellow_names}  # case-insensitive name match
    on_leave_lookup = {n.lower() for n in on_leave}     # case-insensitive on-leave match
    if availability_path is not None and (
            isinstance(availability_path, pd.DataFrame) or os.path.isfile(availability_path)):
        avail_df = _as_df(availability_path)
        for idx, row in avail_df.iterrows():
            person_raw = str(row.get("Person", "")).strip()
            if not person_raw or person_raw.lower() == "nan":
                continue
            day_raw = str(row.get("Day", "")).strip().lower()
            sess_raw = str(row.get("Session", "")).strip().upper()
            reason = ""
            if "Reason" in avail_df.columns and pd.notna(row.get("Reason")):
                reason = str(row["Reason"]).strip()

            # Validate person name (v12.2: orphans no longer raise)
            canonical_name = name_lookup.get(person_raw.lower())
            if canonical_name is None:
                # Two cases:
                #   (a) Person is on Leave this week — silently skip (info redundant).
                #   (b) Genuinely unrecognized — record as orphan for Notes sheet.
                if person_raw.lower() in on_leave_lookup:
                    continue
                availability_orphans.append({
                    "person": person_raw,
                    "day": day_raw,
                    "session": sess_raw,
                    "reason": reason,
                })
                continue

            # Validate day
            valid_days = {"sun","mon","tue","wed","thu","fri","sat",
                          "sunday","monday","tuesday","wednesday","thursday","friday","saturday"}
            if day_raw not in valid_days:
                raise ValueError(
                    f"availability.xlsx row {idx + 2}: invalid Day '{day_raw}'. "
                    f"Use sun/mon/tue/wed/thu (or full names)."
                )
            # Normalize to short
            day_short = day_raw[:3] if len(day_raw) > 3 else day_raw

            # Validate session
            if sess_raw not in ("AM", "PM", "ALL"):
                raise ValueError(
                    f"availability.xlsx row {idx + 2}: invalid Session '{sess_raw}'. "
                    f"Use AM, PM, or All."
                )

            # Expand "All" into AM + PM entries
            sessions_to_block = ["AM", "PM"] if sess_raw == "ALL" else [sess_raw]
            for s in sessions_to_block:
                avail_lookup.add((canonical_name, day_short, s))
                availability.append({
                    "person": canonical_name,
                    "day": day_short,
                    "session": s,
                    "reason": reason,
                })

    # --- Purple rules ---
    purple_rules = []
    for idx, row in purple_df.iterrows():
        fellow = str(row.get("Fellow", "")).strip()
        if not fellow or fellow.lower() == "nan":
            continue
        purple_rules.append({
            "fellow":     fellow,
            "day":        str(row.get("Day", "")).strip().lower(),
            "session":    str(row.get("Session", "")).strip().upper(),
            "consultant": str(row.get("Consultant", "")).strip().lower(),
        })

    # v12.2: filter out purple rules referencing on-Leave persons BEFORE downstream
    # logic uses purple_rules. Otherwise externals logic miscounts the on-Leave
    # person as an external covering the slot, silently reducing demand. These
    # filtered rules are surfaced in the Notes sheet (purple_orphans_leave).
    purple_orphans_leave = []
    active_purple_rules = []
    for r in purple_rules:
        if r["fellow"].lower() in on_leave_lookup:
            purple_orphans_leave.append({
                "person": r["fellow"],
                "day": r["day"],
                "session": r["session"],
                "consultant": r["consultant"],
            })
        else:
            active_purple_rules.append(r)
    purple_rules = active_purple_rules

    # --- Clinics ---
    # ClinicType column (optional, defaults to "Regular") supports add-on clinics —
    # smaller (~8-10 pt) sessions running in an absent consultant's slot, covered
    # solo by a fellow (matching rotation) or an assistant. Band/capacity/min-fellow
    # rules do NOT apply to add-ons (the chief sets patient count low enough that
    # one person handles it; the normal capacity model would erroneously flag a gap).
    VALID_CLINIC_TYPES = {"regular", "addon", "add-on", "add on"}
    clinics = []
    for idx, row in clinics_df.iterrows():
        cons = str(row.get("Consultant", "")).strip().lower()
        if not cons or cons in ("no clinics", "nan"):
            continue
        day  = str(row.get("Day", "")).strip().lower()
        sess = str(row.get("Session", "")).strip().upper()
        if not day or not sess:
            raise ValueError(
                f"clinics.xlsx row {idx + 2}: '{cons}' is missing Day or Session."
            )

        # Parse + validate ClinicType (optional column; backward compatible)
        is_addon = False
        if "ClinicType" in clinics_df.columns and pd.notna(row.get("ClinicType")):
            ctype_raw = str(row["ClinicType"]).strip().lower()
            if ctype_raw and ctype_raw != "nan":
                if ctype_raw not in VALID_CLINIC_TYPES:
                    raise ValueError(
                        f"clinics.xlsx row {idx + 2}: unknown ClinicType '{row['ClinicType']}'. "
                        f"Valid values: Regular, AddOn."
                    )
                is_addon = ctype_raw in ("addon", "add-on", "add on")

        patients = parse_patients(row.get("Patients"))
        # Skip clinics with 0 patients — they're not actually running this week
        # and shouldn't consume helper capacity. Chemo clinics are a separate code
        # path, so this only filters regular consultant clinics.
        if patients == 0:
            continue

        if is_addon:
            # Add-ons: 1 helper required, no consultant capacity, no band/min-fellow rules.
            reqs = {"min_total": 1, "min_fellows": 0, "patient_cap_req": 0}
        else:
            reqs = compute_requirements(patients, cfg)

        is_purple = any(r["day"] == day and r["session"] == sess
                        and r["consultant"] == cons for r in purple_rules)
        if is_purple:
            # Purple slot must have at least 1 helper (the forced person)
            reqs["min_total"] = max(1, reqs["min_total"])
        clinics.append({
            "id": idx, "day": day, "session": sess, "consultant": cons,
            "specialty": str(row["Specialty"]).strip(),
            "patients": patients,
            "required_helpers": reqs["min_total"],
            "min_fellows":      reqs["min_fellows"],
            "patient_cap_req":  reqs["patient_cap_req"],
            "is_chemo": False,
            "is_addon": is_addon,
        })

    # External purple staff (non-fellow) cover one slot, treated as fellow-equivalent
    for rule in purple_rules:
        if rule["fellow"] not in fellow_names:
            for c in clinics:
                if (c["day"] == rule["day"] and c["session"] == rule["session"]
                    and c["consultant"] == rule["consultant"]):
                    c["required_helpers"] = max(0, c["required_helpers"] - 1)
                    c["min_fellows"]      = max(0, c["min_fellows"] - 1)
                    c["patient_cap_req"]  = max(0, c["patient_cap_req"] - cfg.fellow_capacity)

    # --- Chemo clinics ---
    chemo_ids = set()
    if cfg.chemo_active:
        next_id = max((c["id"] for c in clinics), default=0) + 1
        for day, sess in cfg.chemo_slots:
            cid = next_id
            next_id += 1
            clinics.append({
                "id": cid, "day": day, "session": sess,
                "consultant": "chemoassessment", "specialty": "Chemo",
                "patients": 0, "required_helpers": 1,
                "min_fellows": 0, "patient_cap_req": 0,
                "is_chemo": True,
                "is_addon": False,
            })
            chemo_ids.add(cid)

    # --- Pre-flight purple validation ---
    # Walks every purple rule and checks it against the constraints that would
    # otherwise silently render the model INFEASIBLE (or quietly suboptimal):
    #   - Pinning someone marked Leave this week (v12.2: hard error — intent conflict)
    #   - Pinning someone to a slot they're on leave for (per availability.xlsx)
    #   - Pinning to a non-existent clinic (warn, don't error — could be cosmetic)
    #
    # v12.2 change: AddOn rotation-match is NO LONGER pre-flight-validated for
    # internal fellows. Purple is the chief's explicit override — same principle
    # already applied to externals on AddOns. The eligibility constraint below
    # (in the AddOn HARD constraint loop) is updated in parallel.
    name_to_fellow = {f["name"].lower(): f for f in fellows}
    purple_errors = []
    for i, rule in enumerate(purple_rules):
        person = rule["fellow"]
        day = rule["day"]
        sess = rule["session"]
        cons = rule["consultant"]
        if not person or not day or not sess or not cons:
            continue  # skip malformed rows (already filtered earlier but defensive)
        # Note: on-Leave persons already filtered out of purple_rules upstream
        # (v12.2). Their purple entries are tracked in purple_orphans_leave for
        # the Notes sheet.

        # Find the target clinic (regular, not chemo)
        matches = [c for c in clinics
                   if c["day"] == day and c["session"] == sess
                   and c["consultant"] == cons and not c["is_chemo"]]
        if not matches:
            # Purple points at a clinic that doesn't exist this week — likely a
            # leftover row from a prior week. Warn-style (won't block solver).
            continue
        target = matches[0]

        # Check 1: leave conflict (canonicalize person name for case-insensitive match)
        canonical_person = name_lookup.get(person.lower(), person)
        if (canonical_person, day, sess) in avail_lookup:
            purple_errors.append(
                f"purple_clinics.xlsx row {i + 2}: '{person}' is pinned to "
                f"{day.title()} {sess} {cons.title()} but is on leave that session "
                f"(per availability.xlsx). Remove one of the conflicting entries."
            )
            continue

        # Check 2: AddOn eligibility (v12.2: rotation check dropped for fellows)
        if target.get("is_addon"):
            f_obj = name_to_fellow.get(person.lower())
            if f_obj is None:
                # External purple on an AddOn — chief explicitly pinned them.
                continue
            role = f_obj["role"]
            # v12.2: rotation match for fellows on AddOns is no longer pre-flight
            # validated. Purple is the chief's deliberate override (same as for
            # externals). Solver's eligibility constraint is also relaxed below.
            if role == "np":
                purple_errors.append(
                    f"purple_clinics.xlsx row {i + 2}: '{person}' is an NP, but is "
                    f"pinned to AddOn clinic {day.title()} {sess} {cons.title()}. "
                    f"NPs cannot cover AddOns."
                )
            elif role == "resident":
                purple_errors.append(
                    f"purple_clinics.xlsx row {i + 2}: '{person}' is a resident "
                    f"(non-assistant), but is pinned to AddOn clinic {day.title()} "
                    f"{sess} {cons.title()}. Only fellows and assistants can cover AddOns."
                )
            # role == "fellow" or "assistant": always eligible with purple override

    if purple_errors:
        raise ValueError(
            "Purple rule conflicts found (these would make the schedule infeasible "
            "or silently broken):\n" + "\n".join(f"  - {e}" for e in purple_errors)
        )

    # --- Pre-flight diagnostics ---
    demand = sum(c["required_helpers"] for c in clinics)
    capacity_min = sum(f["min"] for f in fellows)
    capacity_max = sum(f["max"] for f in fellows)

    DAYS = sorted({c["day"] for c in clinics})
    SESSIONS = ["AM", "PM"]

    # Fast hard-min feasibility screen. Because one person cannot attend two
    # clinics in the same session, the relevant upper bound is the number of
    # distinct eligible day/session slots, not the raw number of clinics.
    minimum_slot_issues = []
    for f in fellows:
        eligible_sessions = set()
        for c in clinics:
            if (f["name"], c["day"], c["session"]) in avail_lookup:
                continue
            if c["is_chemo"]:
                eligible = f["role"] in ("fellow", "assistant")
            elif c["is_addon"]:
                external_present = any(
                    r["fellow"] not in fellow_names
                    and r["day"] == c["day"]
                    and r["session"] == c["session"]
                    and r["consultant"] == c["consultant"]
                    for r in purple_rules
                )
                eligible = (not external_present
                            and f["role"] in ("fellow", "assistant"))
            elif f["is_np"]:
                eligible = c["specialty"].lower() == f["rotation"].lower()
            else:
                eligible = True
            if eligible:
                eligible_sessions.add((c["day"], c["session"]))

        if len(eligible_sessions) < f["min"]:
            minimum_slot_issues.append({
                "person": f["name"],
                "minimum": f["min"],
                "eligible_sessions": len(eligible_sessions),
                "role": f["role"],
            })

    if minimum_slot_issues:
        detail = "; ".join(
            f"{x['person']} needs {x['minimum']} but has only "
            f"{x['eligible_sessions']} eligible session(s)"
            for x in minimum_slot_issues
        )
        return {
            "status": "INFEASIBLE",
            "message": "Hard minimum clinic requirements cannot be met. " + detail,
            "diagnostics": {
                "demand": demand,
                "capacity_min": capacity_min,
                "capacity_max": capacity_max,
                "minimum_slot_issues": minimum_slot_issues,
            },
            "schedule": [], "fellows_summary": [],
            "residents_summary": [], "assistants_summary": [],
            "availability": availability,
            "availability_orphans": availability_orphans,
            "purple_orphans_leave": purple_orphans_leave,
            "on_leave": on_leave,
        }

    # --- Build CP-SAT model ---
    model = cp_model.CpModel()
    assign = {(c["id"], f["name"]): model.NewBoolVar(f"c{c['id']}_{f['name']}")
              for c in clinics for f in fellows}

    # Chemo eligibility: only fellows + assistants can be assigned.
    # Residents and NPs cannot write chemo.
    # v12.2: HARD cap of at most 1 helper per chemo slot. Matches the AddOn pattern.
    # Without this cap, the solver dumps under-min fellows onto chemo as 2nd helpers
    # because under_min_fellow (150) > over_assign (110) — eating chemo as a slack
    # absorber instead of distributing fellows across regular clinics. 0 is still
    # allowed (CHEMO UNCOVERED warning) to avoid INFEASIBLE under severe shortage.
    for c in clinics:
        if c["is_chemo"]:
            for f in fellows:
                eligible = (f["role"] in ("fellow", "assistant"))
                if not eligible:
                    model.Add(assign[(c["id"], f["name"])] == 0)
            # Hard cap: max 1 helper per chemo slot (v12.2)
            model.Add(sum(assign[(c["id"], f["name"])] for f in fellows) <= 1)

    # Add-on eligibility (HARD):
    #   - Fellows may cover any add-on, but same-rotation coverage is strongly
    #     preferred in the objective. This creates a safe fallback when the only
    #     matching fellow would otherwise be pushed to an excessive workload.
    #   - Assistant: eligible regardless of specialty (assistants are cross-trained,
    #     and residents.xlsx has no Rotation column to filter on).
    #   - Resident (non-assistant): not eligible.
    #   - NP: not eligible.
    #   Also enforces a HARD upper bound of 1 helper per add-on (chief: "always 1 person";
    #   without this, two rotation-matching fellows would overstaff the slot since each
    #   earns +100 rotation reward against only -110 over_assign cost).
    # Pre-compute purple-pinned set. Explicit chief pins are exempt from the
    # cross-rotation preference penalty later in the objective.
    purple_addon_pins = set()  # set of (person_lower, clinic_id)
    for c in clinics:
        if not c.get("is_addon"):
            continue
        for r in purple_rules:
            if (r["day"] == c["day"] and r["session"] == c["session"]
                and r["consultant"] == c["consultant"]):
                purple_addon_pins.add((r["fellow"].lower(), c["id"]))

    for c in clinics:
        if not c["is_addon"]:
            continue
        for f in fellows:
            if f["role"] == "assistant":
                eligible = True
            elif f["role"] == "fellow":
                eligible = True
            else:
                eligible = False  # residents (non-assistant) and NPs
            if not eligible:
                model.Add(assign[(c["id"], f["name"])] == 0)
        # Hard cap: max 1 helper on an add-on (counting externals via reserved capacity).
        # If an external is already pinned via purple, no internal should also be
        # assigned — the cap becomes 0 internal. If no external, cap is 1 internal.
        ext_on_addon = sum(
            1 for r in purple_rules
            if r["fellow"] not in fellow_names
            and r["day"] == c["day"]
            and r["session"] == c["session"]
            and r["consultant"] == c["consultant"]
        )
        model.Add(sum(assign[(c["id"], f["name"])] for f in fellows) <= max(0, 1 - ext_on_addon))

    # NP rotation lock: NPs can only be assigned to clinics matching their rotation.
    # (Their rotation is treated as a HARD constraint, unlike fellows where it's a
    #  preference encoded in the objective.)
    for f in fellows:
        if f["is_np"]:
            for c in clinics:
                if c["is_chemo"]:
                    continue  # already excluded above
                if c["specialty"].lower() != f["rotation"].lower():
                    model.Add(assign[(c["id"], f["name"])] == 0)

    # HARD: Availability — exclude people from clinics on their leave day/session.
    # avail_lookup is a set of (person, day, session) triples loaded above.
    if avail_lookup:
        for c in clinics:
            for f in fellows:
                if (f["name"], c["day"], c["session"]) in avail_lookup:
                    model.Add(assign[(c["id"], f["name"])] == 0)

    # HARD: 20-23 patient clinics — combined cap rule.
    # External fellow-equivalent purples (Rakan, Faisal, Shaikhah, Alsallom)
    # COUNT toward the 3-fellow cap and the 2-fellow threshold.
    # Per chief: "2 fellows alone OK, OR 1 fellow + 2 non-fellows; no 3 fellows".
    for c in clinics:
        if c["is_chemo"] or c["is_addon"] or not (20 <= c["patients"] <= 23):
            continue
        # Count external fellow-equiv purples already pinned to this clinic
        ext_fe_count = sum(1 for r in purple_rules
                           if r["fellow"] not in fellow_names
                           and r["day"] == c["day"]
                           and r["session"] == c["session"]
                           and r["consultant"] == c["consultant"])
        all_internal = [assign[(c["id"], f["name"])] for f in fellows]
        int_fellow_eq = [assign[(c["id"], f["name"])] for f in fellows
                         if f["role"] in ("fellow", "assistant")]
        # Total fellow_eq (internal + external) must be ≤ 2
        max_internal_fe = max(0, 2 - ext_fe_count)
        if int_fellow_eq:
            model.Add(sum(int_fellow_eq) <= max_internal_fe)
        # If total fellow_eq ≤ 1, total helpers must be ≥ 3
        # Equivalently: if int_fellow_eq + ext_fe_count < 2, int_total + ext_fe_count ≥ 3
        if int_fellow_eq:
            has_2f_total = model.NewBoolVar(f"has_2f_{c['id']}")
            model.Add(sum(int_fellow_eq) + ext_fe_count >= 2).OnlyEnforceIf(has_2f_total)
            model.Add(sum(int_fellow_eq) + ext_fe_count <= 1).OnlyEnforceIf(has_2f_total.Not())
            model.Add(sum(all_internal) + ext_fe_count >= 3).OnlyEnforceIf(has_2f_total.Not())

    # HARD: 24-26 — exactly 2F + 1 non-fellow is preferred but not enforced
    #              (min_total=3, min_fellows=2 already cover the floor)
    # HARD: 27+ — same floor (3, 2). The "3F or 2F+2 non-fellow" preference is soft.

    # Soft helper count: at LEAST required_helpers (floor, not exact).
    # Over-assigning is allowed and lightly penalized to keep numbers sensible
    # but doesn't waste fellows when others need their min hit.
    # Pre-compute external (purple) count per clinic. Externals don't have assign
    # variables — they're pinned via the purple file and displayed at output time.
    # But they DO count as people covering the clinic. Without this lookup, the
    # solver's shortfall accounting treats external-covered slots as uncovered
    # (false ADD-ON UNFILLED / HEAD SHORT warnings) and the AddOn hard cap of 1
    # helper can't see them either.
    ext_count_by_clinic = {}
    for c in clinics:
        ext_count_by_clinic[c["id"]] = sum(
            1 for r in purple_rules
            if r["fellow"] not in fellow_names
            and r["day"] == c["day"]
            and r["session"] == c["session"]
            and r["consultant"] == c["consultant"]
        )

    shortfall = {}
    over_assign = {}
    over_excess = {}
    for c in clinics:
        sf = model.NewIntVar(0, c["required_helpers"], f"sf_{c['id']}")
        shortfall[c["id"]] = sf
        # Total assigned >= required - shortfall (so shortfall absorbs deficit)
        # Total assigned can also exceed required (no upper cap from this rule).
        # Include externals — they count as people covering the slot.
        total_assigned = (sum(assign[(c["id"], f["name"])] for f in fellows)
                          + ext_count_by_clinic[c["id"]])
        model.Add(total_assigned + sf >= c["required_helpers"])

        # Track over-assignment for objective.
        # 'over' = total - required (linear, mild penalty per over)
        # 'excess' = max(0, over - 1) — the part beyond +1, penalized heavily
        # so over-assignments spread across clinics rather than concentrate.
        over = model.NewIntVar(0, len(fellows) + ext_count_by_clinic[c["id"]],
                               f"over_{c['id']}")
        model.Add(over >= total_assigned - c["required_helpers"])
        over_assign[c["id"]] = over

        excess = model.NewIntVar(0, len(fellows) + ext_count_by_clinic[c["id"]],
                                 f"excess_{c['id']}")
        model.Add(excess >= over - 1)
        over_excess[c["id"]] = excess

    # SOFT: minimum fellow-equivalent helpers per clinic (chemo-write rule).
    # Fellows + assistants count. Violations heavily penalized but allowed in
    # severe under-staffing weeks (so we always return SOME schedule).
    fellow_short = {}   # clinic_id -> IntVar tracking fellow shortfall
    for c in clinics:
        if c["is_chemo"] or c["min_fellows"] == 0:
            continue
        fs = model.NewIntVar(0, c["min_fellows"], f"fshort_{c['id']}")
        fellow_short[c["id"]] = fs
        fellow_eq = [assign[(c["id"], f["name"])] for f in fellows
                     if f["role"] in ("fellow", "assistant")]
        if fellow_eq:
            model.Add(sum(fellow_eq) + fs >= c["min_fellows"])

    # SOFT: patient capacity. Same rationale.
    def _cap_unit(person):
        if person["role"] in ("fellow", "assistant"):
            return cfg.fellow_capacity
        if person["is_np"]:
            return cfg.np_capacity
        return cfg.resident_capacity

    cap_short = {}      # clinic_id -> IntVar tracking patient capacity shortfall
    for c in clinics:
        if c["is_chemo"] or c["patient_cap_req"] <= 0:
            continue
        cs = model.NewIntVar(0, c["patient_cap_req"], f"cshort_{c['id']}")
        cap_short[c["id"]] = cs
        cap_terms = [assign[(c["id"], f["name"])] * _cap_unit(f) for f in fellows]
        model.Add(sum(cap_terms) + cs >= c["patient_cap_req"])

    # Workload bounds: hard min and hard max.
    # The previous soft-min model could deliberately leave people below their
    # configured minimum when other objective penalties were cheaper. Minimums
    # are now true roster requirements. The zero-valued under_min variable is
    # retained for backward-compatible summaries and diagnostics.
    totals = {}
    under_min = {}
    for f in fellows:
        t = model.NewIntVar(0, f["max"], f"total_{f['name']}")
        model.Add(t == sum(assign[(c["id"], f["name"])] for c in clinics))
        model.Add(t >= f["min"])
        model.Add(t <= f["max"])
        um = model.NewIntVar(0, 0, f"under_{f['name']}")
        totals[f["name"]] = t
        under_min[f["name"]] = um

    # Forced internal purple
    for rule in purple_rules:
        if rule["fellow"] in fellow_names:
            for c in clinics:
                if (c["day"] == rule["day"] and c["session"] == rule["session"]
                    and c["consultant"] == rule["consultant"]):
                    model.Add(assign[(c["id"], rule["fellow"])] == 1)

    # No double-booking
    for f in fellows:
        for d in DAYS:
            for s in SESSIONS:
                same = [assign[(c["id"], f["name"])] for c in clinics
                        if c["day"] == d and c["session"] == s]
                if same:
                    model.Add(sum(same) <= 1)

    # AM+PM same-day controls
    double_day_vars = []
    for f in fellows:
        if f["is_resident"]:
            continue
        per_day = []
        for d in DAYS:
            am = [assign[(c["id"], f["name"])] for c in clinics
                  if c["day"] == d and c["session"] == "AM"]
            pm = [assign[(c["id"], f["name"])] for c in clinics
                  if c["day"] == d and c["session"] == "PM"]
            if not am or not pm:
                continue
            am_used = model.NewBoolVar(f"am_{f['name']}_{d}")
            pm_used = model.NewBoolVar(f"pm_{f['name']}_{d}")
            model.Add(sum(am) >= 1).OnlyEnforceIf(am_used)
            model.Add(sum(am) == 0).OnlyEnforceIf(am_used.Not())
            model.Add(sum(pm) >= 1).OnlyEnforceIf(pm_used)
            model.Add(sum(pm) == 0).OnlyEnforceIf(pm_used.Not())
            both = model.NewBoolVar(f"both_{f['name']}_{d}")
            model.AddBoolAnd([am_used, pm_used]).OnlyEnforceIf(both)
            model.AddBoolOr([am_used.Not(), pm_used.Not()]).OnlyEnforceIf(both.Not())
            per_day.append(both)
            double_day_vars.append(both)
        if per_day:
            model.Add(sum(per_day) <= cfg.max_double_days)

    # Fairness: per-person overage above min
    # The previous "spread" metric (max - min) was brittle — only changed if the
    # absolute extremes moved. Replaced with per-person overage which pushes the
    # solver to spread above-min clinics across MANY people rather than stacking
    # on a few. Applies to fellows + NPs (residents have min=max so always 0).
    over_min = {}
    for f in fellows:
        if f["is_resident"]:
            continue  # residents already have tight min/max
        om = model.NewIntVar(0, f["max"], f"over_min_{f['name']}")
        model.Add(om >= totals[f["name"]] - f["min"])
        over_min[f["name"]] = om
    total_over_min = sum(over_min.values()) if over_min else model.NewConstant(0)

    # Fellow workload ceiling preference. Six clinics is acceptable; only the
    # seventh-and-beyond portion carries an additional strong penalty.
    actual_fellows = [f for f in fellows if f["role"] == "fellow"]
    seventh_clinic_units = {}
    for f in actual_fellows:
        seventh = model.NewIntVar(0, max(0, f["max"] - 6),
                                  f"seventh_units_{f['name']}")
        model.Add(seventh >= totals[f["name"]] - 6)
        seventh_clinic_units[f["name"]] = seventh
    total_seventh_units = (sum(seventh_clinic_units.values())
                           if seventh_clinic_units else model.NewConstant(0))

    # Soft educational floor for in-rotation exposure. The target is capped by
    # the number of eligible in-rotation sessions in this specific week, so the
    # solver is not punished for clinics that simply do not exist.
    in_rotation_counts = {}
    in_rotation_short = {}
    for f in actual_fellows:
        matching = [c for c in clinics
                    if not c["is_chemo"]
                    and c["specialty"].lower() == f["rotation"].lower()]
        potential_sessions = set()
        for c in matching:
            if (f["name"], c["day"], c["session"]) in avail_lookup:
                continue
            if c["is_addon"]:
                external_present = any(
                    r["fellow"] not in fellow_names
                    and r["day"] == c["day"]
                    and r["session"] == c["session"]
                    and r["consultant"] == c["consultant"]
                    for r in purple_rules
                )
                if external_present:
                    continue
            potential_sessions.add((c["day"], c["session"]))

        max_count = len(matching)
        cnt = model.NewIntVar(0, max_count, f"in_rotation_{f['name']}")
        model.Add(cnt == sum(assign[(c["id"], f["name"])] for c in matching))
        in_rotation_counts[f["name"]] = cnt

        target = min(cfg.preferred_in_rotation_min, f["min"],
                     len(potential_sessions))
        short = model.NewIntVar(0, target, f"in_rotation_short_{f['name']}")
        model.Add(cnt + short >= target)
        in_rotation_short[f["name"]] = short
    total_in_rotation_short = (sum(in_rotation_short.values())
                               if in_rotation_short else model.NewConstant(0))

    # Repeated chemoassessment preference. This is intentionally soft because
    # staffing shortages may require one fellow to cover more than two sessions.
    chemo_repeat_excess = {}
    for f in actual_fellows:
        chemo_count_expr = sum(assign[(c["id"], f["name"])]
                               for c in clinics if c["is_chemo"])
        max_excess = max(0, sum(1 for c in clinics if c["is_chemo"])
                         - cfg.preferred_chemo_max)
        excess = model.NewIntVar(0, max_excess,
                                 f"chemo_repeat_{f['name']}")
        model.Add(excess >= chemo_count_expr - cfg.preferred_chemo_max)
        chemo_repeat_excess[f["name"]] = excess
    total_chemo_repeat = (sum(chemo_repeat_excess.values())
                          if chemo_repeat_excess else model.NewConstant(0))

    # Per-rotation in-rotation fairness (v12).
    # Group fellows by rotation. For each rotation with ≥2 fellows, compute the
    # in-rotation clinic count per fellow and penalize the spread (max - min).
    # Applies to FELLOWS only (NPs, residents, assistants don't have a meaningful
    # rotation distribution concern — their roster role spans specialties).
    rotation_groups = {}
    for f in fellows:
        if f["is_resident"] or f["is_assistant"] or f["is_np"]:
            continue
        rotation_groups.setdefault(f["rotation"].lower(), []).append(f)

    rotation_spread_vars = {}  # rotation -> spread IntVar (for diagnostics + objective)
    for rot_lower, group in rotation_groups.items():
        if len(group) < 2:
            continue  # solo-fellow rotations have no distribution to balance
        in_rot_counts = []
        max_possible = sum(1 for c in clinics
                           if c["specialty"].lower() == rot_lower)
        if max_possible == 0:
            continue
        for f in group:
            in_rot_counts.append(in_rotation_counts[f["name"]])
        rot_max = model.NewIntVar(0, max_possible, f"max_in_{rot_lower}")
        rot_min = model.NewIntVar(0, max_possible, f"min_in_{rot_lower}")
        rot_spread = model.NewIntVar(0, max_possible, f"spread_in_{rot_lower}")
        model.AddMaxEquality(rot_max, in_rot_counts)
        model.AddMinEquality(rot_min, in_rot_counts)
        model.Add(rot_spread == rot_max - rot_min)
        rotation_spread_vars[rot_lower] = rot_spread

    total_rotation_spread = (sum(rotation_spread_vars.values())
                             if rotation_spread_vars else model.NewConstant(0))

    # True fellow workload spread, now included in the objective.
    fellow_only = [totals[f["name"]] for f in actual_fellows]
    if fellow_only:
        max_t = model.NewIntVar(0, 100, "max_t")
        min_t = model.NewIntVar(0, 100, "min_t")
        model.AddMaxEquality(max_t, fellow_only)
        model.AddMinEquality(min_t, fellow_only)
        spread = model.NewIntVar(0, 100, "spread")
        model.Add(spread == max_t - min_t)
    else:
        spread = model.NewConstant(0)

    # Objective
    rotation_matches = sum(
        assign[(c["id"], f["name"])]
        for c in clinics for f in fellows
        if not f["is_resident"] and not f["is_assistant"]
        and c["specialty"].lower() == f["rotation"].lower()
    )
    total_shortfall = sum(shortfall.values())
    # Shortfall penalty:
    #   - chemo: high flat cost (w_chemo_shortfall)
    #   - add-on: high flat cost (w_addon_uncovered) — per chief, as critical as chemo
    #   - regular: scales with patient volume so busy clinics are protected most
    weighted_shortfall_terms = []
    for c in clinics:
        if c["is_chemo"]:
            penalty = cfg.w_chemo_shortfall
        elif c["is_addon"]:
            penalty = cfg.w_addon_uncovered
        else:
            penalty = 50 + 5 * c["patients"]         # base 50 + patient severity
        weighted_shortfall_terms.append(penalty * shortfall[c["id"]])
    weighted_shortfall = sum(weighted_shortfall_terms)
    total_over = sum(over_assign.values())
    total_excess = sum(over_excess.values())
    under_min_f = sum(under_min[f["name"]] for f in fellows
                      if not f["is_resident"] and not f["is_assistant"])
    under_min_r = sum(under_min[f["name"]] for f in fellows
                      if f["is_resident"] or f["is_assistant"])
    assist_chemo = sum(assign[(c["id"], f["name"])] for c in clinics for f in fellows
                       if c["is_chemo"] and f["is_assistant"])

    # Add-on fallback penalty. Same-rotation fellows are preferred. A
    # cross-rotation fellow or an assistant remains available when concentrating
    # all add-ons on one matching fellow would otherwise create an excessive load.
    # Explicit purple pins are deliberate chief overrides and carry no penalty.
    cross_rotation_addon = sum(
        assign[(c["id"], f["name"])]
        for c in clinics if c["is_addon"]
        for f in fellows
        if f["role"] in ("fellow", "assistant")
        and (f["role"] == "assistant"
             or c["specialty"].lower() != f["rotation"].lower())
        and (f["name"].lower(), c["id"]) not in purple_addon_pins
    )

    total_fellow_short = sum(fellow_short.values()) if fellow_short else 0
    total_cap_short    = sum(cap_short.values())    if cap_short    else 0

    # -----------------------------------------------------------------
    # Band-based soft preferences (chief fellow's stated patterns)
    # -----------------------------------------------------------------
    # 12–14: prefer NP or resident as the helper (save fellows for busier)
    # 17–19: prefer "1 fellow + 1 non-fellow" — discourage 2 fellows here
    # 20–23: prefer "2 fellows alone" over "1 fellow + 2 non-fellows"
    #        (i.e., discourage adding a 3rd helper when 2 fellows already suffice)
    # 24–26: 3rd helper should be non-fellow (discourage 3 fellows)
    # 27+ : prefer "2 fellows + 2 non-fellows" over "3 fellows"
    band_penalty_terms = []
    for c in clinics:
        if c["is_chemo"] or c["is_addon"]:
            continue
        p = c["patients"]
        # External fellow-equivalent purples already pinned here
        ext_fe = sum(1 for r in purple_rules
                     if r["fellow"] not in fellow_names
                     and r["day"] == c["day"]
                     and r["session"] == c["session"]
                     and r["consultant"] == c["consultant"])
        # Per-clinic helper assignments split by role
        fellow_eq_assigned = [assign[(c["id"], f["name"])] for f in fellows
                              if f["role"] in ("fellow", "assistant")]
        non_fellow_assigned = [assign[(c["id"], f["name"])] for f in fellows
                               if f["role"] not in ("fellow", "assistant")]

        if 12 <= p <= 14:
            # Penalty per fellow_eq present (internal + external).
            # ext_fe is constant; only internal contributes to the variable cost.
            band_penalty_terms.append(cfg.w_band_pref * (sum(fellow_eq_assigned) + ext_fe))

        elif 17 <= p <= 19:
            # Penalize 2nd+ fellow_eq (counting externals).
            extra_f = model.NewIntVar(0, 10, f"f_extra19_{c['id']}")
            model.Add(extra_f >= sum(fellow_eq_assigned) + ext_fe - 1)
            band_penalty_terms.append(cfg.w_band_pref * extra_f)

        elif 20 <= p <= 23:
            # Prefer 2-fellow alone over 1-fellow + 2-others.
            # Penalize each non-fellow assigned.
            band_penalty_terms.append(cfg.w_band_pref * sum(non_fellow_assigned))

        elif 24 <= p <= 26:
            # 3rd helper should be non-fellow. Penalize 3rd+ fellow_eq (counting externals).
            extra_f3 = model.NewIntVar(0, 10, f"f_extra26_{c['id']}")
            model.Add(extra_f3 >= sum(fellow_eq_assigned) + ext_fe - 2)
            band_penalty_terms.append(cfg.w_band_pref * extra_f3)

        elif p >= 27:
            # Prefer 2F+2 non-fellow over 3F. Penalize 3rd+ fellow_eq (counting externals).
            extra_f27 = model.NewIntVar(0, 10, f"f_extra27_{c['id']}")
            model.Add(extra_f27 >= sum(fellow_eq_assigned) + ext_fe - 2)
            band_penalty_terms.append(cfg.w_band_pref * extra_f27)

    band_pref_penalty = sum(band_penalty_terms) if band_penalty_terms else 0

    terms = [
        cfg.w_rotation_match    * rotation_matches,
        -1                      * weighted_shortfall,    # weights baked in above
        -cfg.w_fellow_short     * total_fellow_short,
        -cfg.w_capacity_short   * total_cap_short,
        -cfg.w_over_assign      * total_over,
        -cfg.w_over_excess      * total_excess,
        -cfg.w_under_min_fellow * under_min_f,
        -cfg.w_under_min_res    * under_min_r,
        -cfg.w_fairness         * total_over_min,
        -cfg.w_rotation_fairness * total_rotation_spread,
        -cfg.w_workload_spread  * spread,
        -cfg.w_seventh_clinic   * total_seventh_units,
        -cfg.w_in_rotation_short * total_in_rotation_short,
        -cfg.w_chemo_repeat     * total_chemo_repeat,
        -cfg.w_cross_rotation_addon * cross_rotation_addon,
        cfg.w_assist_for_chemo  * assist_chemo,
        -1                      * band_pref_penalty,     # weight baked in above
    ]
    if cfg.w_double_session and double_day_vars:
        terms.append(-cfg.w_double_session * sum(double_day_vars))
    model.Maximize(sum(terms))

    # Solve
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = cfg.solver_timeout_sec
    status = solver.Solve(model)

    status_name = {
        cp_model.OPTIMAL:    "OPTIMAL",
        cp_model.FEASIBLE:   "FEASIBLE",
        cp_model.INFEASIBLE: "INFEASIBLE",
        cp_model.MODEL_INVALID: "MODEL_INVALID",
    }.get(status, "UNKNOWN")

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return {
            "status": status_name,
            "message": (
                f"No feasible schedule under the hard minimum/maximum, role, "
                f"availability, pin, and no-double-booking rules. "
                f"Total demand={demand}, sum of minimums={capacity_min}, "
                f"sum of maximums={capacity_max}. Review the roster minimums "
                f"and restricted sessions first."
            ),
            "diagnostics": {"demand": demand,
                            "capacity_min": capacity_min,
                            "capacity_max": capacity_max,
                            "minimum_slot_issues": minimum_slot_issues},
            "schedule": [], "fellows_summary": [],
            "residents_summary": [], "assistants_summary": [],
            "availability": availability,
            "availability_orphans": availability_orphans,
            "purple_orphans_leave": purple_orphans_leave,
            "on_leave": on_leave,
        }

    # --- Extract schedule ---
    day_order = {"sunday":0,"sun":0,"monday":1,"mon":1,"tuesday":2,"tue":2,
                 "wednesday":3,"wed":3,"thursday":4,"thu":4}
    sorted_clinics = sorted(clinics,
                            key=lambda x:(day_order.get(x["day"], 9), x["session"]))

    DAY_TITLE = {"sunday":"Sun","sun":"Sun","monday":"Mon","mon":"Mon",
                 "tuesday":"Tue","tue":"Tue","wednesday":"Wed","wed":"Wed",
                 "thursday":"Thu","thu":"Thu"}

    schedule = []
    for c in sorted_clinics:
        internal = [f["name"] for f in fellows
                    if solver.Value(assign[(c["id"], f["name"])]) == 1]
        external = [r["fellow"] + "*" for r in purple_rules
                    if r["fellow"] not in fellow_names
                    and r["day"] == c["day"]
                    and r["session"] == c["session"]
                    and r["consultant"] == c["consultant"]]
        sf_val = solver.Value(shortfall[c["id"]])
        fs_val = solver.Value(fellow_short[c["id"]]) if c["id"] in fellow_short else 0
        cs_val = solver.Value(cap_short[c["id"]])    if c["id"] in cap_short    else 0
        warnings = []
        if fs_val > 0:
            warnings.append(f"NEEDS {fs_val}× FELLOW")
        if cs_val > 0:
            warnings.append(f"CAP SHORT {cs_val}")
        if sf_val > 0 and not c["is_chemo"] and not c["is_addon"]:
            warnings.append(f"HEAD SHORT {sf_val}")
        if c["is_chemo"] and sf_val > 0:
            warnings.append("CHEMO UNCOVERED")
        if c["is_addon"] and sf_val > 0:
            warnings.append("ADD-ON UNFILLED")
        schedule.append({
            "Day":        DAY_TITLE.get(c["day"], c["day"].title()),
            "Session":    c["session"],
            "Consultant": c["consultant"].title(),
            "Specialty":  c["specialty"],
            "Patients":   c["patients"],
            "Required":   c["required_helpers"],
            "MinFellows": c.get("min_fellows", 0),
            "Assigned":   ", ".join(internal + external) if (internal or external) else "—",
            "Shortfall":  sf_val,
            "FellowShort":fs_val,
            "CapShort":   cs_val,
            "Warnings":   "; ".join(warnings),
            "IsChemo":    c["is_chemo"],
            "IsAddOn":    c["is_addon"],
        })

    # Summaries
    fellows_summary = []
    for f in fellows:
        if f["is_resident"] or f["is_assistant"]:
            continue
        total = solver.Value(totals[f["name"]])
        inside = sum(solver.Value(assign[(c["id"], f["name"])]) for c in clinics
                     if c["specialty"].lower() == f["rotation"].lower())
        fellows_summary.append({
            "Fellow": f["name"], "Rotation": f["rotation"],
            "Role": f.get("role", "fellow").title(),
            "Min": f["min"], "Max": f["max"], "Total": total,
            "Inside": inside, "Outside": total - inside,
            "BelowMin": solver.Value(under_min[f["name"]]),
        })

    residents_summary = []
    for f in fellows:
        if not f["is_resident"]:
            continue
        residents_summary.append({
            "Resident": f["name"], "Min": f["min"], "Max": f["max"],
            "Total": solver.Value(totals[f["name"]]),
            "BelowMin": solver.Value(under_min[f["name"]]),
        })

    assistants_summary = []
    for f in fellows:
        if not f["is_assistant"]:
            continue
        chemo_count = sum(solver.Value(assign[(c["id"], f["name"])])
                          for c in clinics if c["is_chemo"])
        assistants_summary.append({
            "Assistant": f["name"], "Min": f["min"], "Max": f["max"],
            "Total": solver.Value(totals[f["name"]]),
            "Chemo": chemo_count,
            "BelowMin": solver.Value(under_min[f["name"]]),
        })

    return {
        "status": status_name,
        "message": "Schedule generated successfully.",
        "schedule": schedule,
        "fellows_summary": fellows_summary,
        "residents_summary": residents_summary,
        "assistants_summary": assistants_summary,
        "availability": availability,
        "availability_orphans": availability_orphans,
        "purple_orphans_leave": purple_orphans_leave,
        "on_leave": on_leave,
        "diagnostics": {
            "rotation_matches": solver.Value(rotation_matches),
            "shortfalls": solver.Value(total_shortfall),
            "fellow_short": solver.Value(total_fellow_short) if fellow_short else 0,
            "cap_short": solver.Value(total_cap_short) if cap_short else 0,
            "under_min_fellows": solver.Value(under_min_f),
            "under_min_residents": solver.Value(under_min_r),
            "spread": solver.Value(spread),
            "seventh_clinic_units": solver.Value(total_seventh_units),
            "in_rotation_short": solver.Value(total_in_rotation_short),
            "chemo_repeat_excess": solver.Value(total_chemo_repeat),
            "cross_rotation_addons": solver.Value(cross_rotation_addon),
            "demand": demand,
            "capacity_min": capacity_min,
            "capacity_max": capacity_max,
            "minimum_slot_issues": minimum_slot_issues,
        },
    }
