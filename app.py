
import io
import re
import base64
from datetime import date, datetime

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

TODAY = date.today()

# -----------------------------
# Business rules learned from supplied examples
# -----------------------------
VALID_RELATIONS = {
    "Self": (18, 80, 1),
    "Spouse": (18, 80, 1),
    "Child": (0, 25, 2),
    "Parent": (40, 80, 2),
    "Parent-in-law": (40, 80, 2),
}

# Bambrew's supplied final workbook allows 4 children.
BAMBrew_VALID_RELATIONS = {
    "Self": (18, 80, 1),
    "Spouse": (18, 80, 1),
    "Child": (0, 25, 4),
    "Parent": (40, 80, 2),
    "Parent-in-law": (40, 80, 2),
}

GENDERS = {"Male", "Female", "Transgender", "Other"}

REL_ALIASES = {
    "self": "Self", "myself": "Self", "my self": "Self", "employee": "Self",
    "spouse": "Spouse", "wife": "Spouse", "husband": "Spouse",
    "spouse(wife)": "Spouse", "spouse (wife)": "Spouse",
    "child": "Child", "son": "Child", "daughter": "Child",
    "child 1": "Child", "child 2": "Child", "child 3": "Child",
    "dependent child1": "Child", "dependent child2": "Child",
    "dependent child 1": "Child", "dependent child 2": "Child",
    "parent": "Parent", "father": "Parent", "fathers": "Parent",
    "mother": "Parent", "mothers": "Parent",
    "parent-in-law": "Parent-in-law", "parent in law": "Parent-in-law",
    "parent law": "Parent-in-law", "parent-law": "Parent-in-law",
    "father-in-law": "Parent-in-law", "mother-in-law": "Parent-in-law",
    "father law": "Parent-in-law", "mother law": "Parent-in-law",
    "father-law": "Parent-in-law", "mother-law": "Parent-in-law",
    "m law": "Parent-in-law", "f law": "Parent-in-law",
    "m-law": "Parent-in-law", "f-law": "Parent-in-law",

    # Standard relationship/member codes supplied by source files.
    "mmbr": "Self", "member": "Self",
    "spse": "Spouse",
    "udtr": "Child", "daughter": "Child",
    "sonm": "Child", "son": "Child",
    "emot": "Parent", "employee mother": "Parent",
    "efat": "Parent", "employee father": "Parent",
    "eflw": "Parent-in-law", "employee father in law": "Parent-in-law",
    "emlw": "Parent-in-law", "employee mother in law": "Parent-in-law",
}
GENDER_ALIASES = {
    "m": "Male", "male": "Male", "mail": "Male",
    "f": "Female", "female": "Female", "femail": "Female",
    "transgender": "Transgender", "other": "Other",
}

FINAL_COLUMNS = [
    "Employee ID", "Member Name", "Date of Birth", "Age", "Age Band",
    "Gender", "Relationship to Self", "GHI Sum Insured",
    "Mobile Number", "Email",
    "Corrections", "Correction Rules",
    "Remark"
]


def local_svg_data_uri(filename):
    """Embed a bundled SVG safely in the Streamlit title banner."""
    try:
        with open(filename, "rb") as logo_file:
            encoded = base64.b64encode(logo_file.read()).decode("ascii")
        return f"data:image/svg+xml;base64,{encoded}"
    except OSError:
        return ""


def load_source_sheet(workbook_bytes, sheet_name):
    """Read a source sheet once per uploaded workbook/session."""
    cache = st.session_state.setdefault("_source_sheet_cache", {})
    cache_key = (hash(workbook_bytes), sheet_name)
    if cache_key not in cache:
        cache[cache_key] = pd.read_excel(
            io.BytesIO(workbook_bytes), sheet_name=sheet_name, header=None
        )
    return cache[cache_key]

def clean_text(v):
    if pd.isna(v):
        return ""
    return " ".join(str(v).replace("\xa0", " ").split()).strip()

def norm(v):
    return clean_text(v).lower()
def normalize_si(value):
    """Normalize source Sum Insured to a numeric rupee amount."""
    if value is None:
        return np.nan
    try:
        if pd.isna(value):
            return np.nan
    except Exception:
        pass
    s = clean_text(value).lower().replace(",", "").replace("₹", "").strip()
    if not s:
        return np.nan
    # Accept both spaced and compact Indian amount shorthand, e.g. 2lakhs,
    # 2 Thousand, and 2 CR.  Removing the suffix leaves the numeric portion.
    multiplier = 1
    suffixes = [
        (r"(?:lakh|lakhs|lac|lacs)", 100000),
        (r"(?:thousand|thousands|k)", 1000),
        (r"(?:crore|crores|cr)", 10000000),
    ]
    for suffix, amount in suffixes:
        if re.search(rf"{suffix}\b", s):
            multiplier = amount
            s = re.sub(rf"{suffix}\b", "", s).strip()
            break
    try:
        return float(s) * multiplier
    except Exception:
        return np.nan


def parse_date(v):
    if pd.isna(v) or clean_text(v) == "":
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.date()
    if isinstance(v, date):
        return v
    s = clean_text(v)
    # The reference final workbooks complete year-only dates as 01-Jan-year.
    if re.fullmatch(r"\d{4}", s):
        try:
            return date(int(s), 1, 1)
        except Exception:
            return None
    try:
        d = pd.to_datetime(s, errors="coerce", dayfirst=True)
        return None if pd.isna(d) else d.date()
    except Exception:
        return None

def employee_id_sort_key(value):
    """Natural alphanumeric Employee ID ordering."""
    s = clean_text(value).casefold()
    parts = re.split(r"(\d+)", s)
    return tuple((1, int(p)) if p.isdigit() else (0, p) for p in parts)

def age_from_dob(dob, as_of=None):
    """Safely calculate age from Excel/Pandas/string DOB values."""
    if dob is None:
        return np.nan
    try:
        if pd.isna(dob):
            return np.nan
    except Exception:
        pass
    try:
        if isinstance(dob, str):
            value = dob.strip()
            if not value:
                return np.nan
            parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
        else:
            parsed = pd.to_datetime(dob, errors="coerce")
        if pd.isna(parsed):
            return np.nan

        dob_date = parsed.date() if hasattr(parsed, "date") else parsed
        today = as_of if as_of is not None else date.today()
        if isinstance(today, pd.Timestamp):
            today = today.date()
        elif isinstance(today, datetime):
            today = today.date()

        years = int(today.year) - int(dob_date.year)
        if (int(today.month), int(today.day)) < (int(dob_date.month), int(dob_date.day)):
            years -= 1
        return years if years >= 0 else np.nan
    except Exception:
        return np.nan


def age_band(age, style="hyphen"):
    if age is None:
        return ""
    if isinstance(age, str) and clean_text(age) == "No age mentioned":
        return ""
    age_num = pd.to_numeric(age, errors="coerce")
    if pd.isna(age_num):
        return ""
    age = float(age_num)
    sep = "/" if style == "slash" else "-"
    if age > 85: return "From 86"
    if age > 80: return f"81{sep}85"
    if age > 75: return f"76{sep}80"
    if age > 70: return f"71{sep}75"
    if age > 65: return f"66{sep}70"
    if age > 60: return f"61{sep}65"
    if age > 55: return f"56{sep}60"
    if age > 50: return f"51{sep}55"
    if age > 45: return f"46{sep}50"
    if age > 40: return f"41{sep}45"
    if age > 35: return f"36{sep}40"
    if age > 30: return f"31{sep}35"
    if age > 25: return f"26{sep}30"
    if age > 18: return f"19{sep}25"
    return f"0{sep}18"


def log_correction(logs, source_row, field, old, new, rule, confidence="High"):
    old_s = "" if pd.isna(old) else clean_text(old)
    new_s = "" if pd.isna(new) else clean_text(new)
    if old_s != new_s:
        logs.append({
            "Source Row": source_row,
            "Field": field,
            "Original Value": old_s,
            "Corrected Value": new_s,
            "Rule Applied": rule,
            "Confidence": confidence,
        })

def normalize_relationship(value):
    s = norm(value)

    # IN-LAW PRECEDENCE RULE:
    # Any relationship label/code that clearly refers to a law/in-law
    # relationship must map to Parent-in-law, never Parent.
    #
    # Covers examples such as:
    # parent-in-law, parent in law, father-in-law, mother-in-law,
    # father law, mother law, father-law, mother-law,
    # M law, F law, M-law, F-law, M/Law, F/Law,
    # employee father in law, employee mother in law,
    # EFLW, EMLW, and similar variants.
    compact = re.sub(r"[^a-z0-9]+", "", s)
    spaced = re.sub(r"[^a-z0-9]+", " ", s).strip()

    law_terms = (
        "inlaw" in compact
        or "law" in spaced.split()
        or "law" in compact
    )

    if law_terms:
        return "Parent-in-law"

    # Sources often encode children as Dependent Child1/2/3/... . Treat every
    # numbered child variant consistently instead of relying on a fixed list.
    if re.fullmatch(r"(?:dependent\s*)?child\s*\d+", s):
        return "Child"

    return REL_ALIASES.get(s, clean_text(value))

def normalize_gender(value):
    s = norm(value)
    return GENDER_ALIASES.get(s, clean_text(value))


def parent_name_indicates_unavailable(value):
    """Return True when a parent name is a deceased/unavailable placeholder."""
    name = norm(value)
    return bool(
        re.search(r"\blate\b", name)
        or re.fullmatch(r"n\s*/?\s*a\.?", name)
        or re.search(r"\bnot\s+avai(?:lable|ble)\b", name)
    )

def find_header_row(raw, expected_keywords):
    best_row, best_score = 0, -1
    for i in range(min(20, len(raw))):
        vals = [norm(x) for x in raw.iloc[i].tolist()]
        score = sum(
            any(k in v for k in expected_keywords)
            for v in vals
        )
        if score > best_score:
            best_row, best_score = i, score
    return best_row, best_score


def classify_sheet(raw):
    """
    Generalized semantic classifier.
    No client names, fixed column positions, or mandatory Employee ID.
    """
    if raw is None or raw.empty:
        return "unknown", 0

    vocab = [
        "employee","emp","member","insured","name","date of birth","dob",
        "age","spouse","wife","husband","child","father","mother","parent",
        "relationship","relation","mobile","phone"
    ]

    best_row, best_score = 0, -1
    for r in range(min(20, len(raw))):
        vals = [str(v).strip().lower() for v in raw.iloc[r].tolist()]
        vals = [v for v in vals if v and v != "nan"]
        score = sum(any(token in v for token in vocab) for v in vals)
        if score > best_score:
            best_row, best_score = r, score

    headers = [str(v).strip().lower() for v in raw.iloc[best_row].tolist()]
    headers = [v for v in headers if v and v != "nan"]

    employee_name = any(
        ("employee" in h or h.startswith("emp") or
         "member" in h or "insured" in h) and "name" in h
        for h in headers
    )
    employee_dob = any(
        "employee" in h and
        ("date of birth" in h or "dob" in h or "birth" in h)
        for h in headers
    )
    spouse = any(any(k in h for k in ("spouse","wife","husband")) for h in headers)
    child = any(re.search(r"child\s*\d+", h) for h in headers)
    parent = any(any(k in h for k in ("father","mother","parent","dad","mom")) for h in headers)

    if employee_name and (employee_dob or spouse or child or parent):
        return "wide_tetris", best_row

    if "self name" in headers and (spouse or child):
        return "wide_bambrew", best_row

    has_id = any(
        any(k in h for k in (
            "member id","memberid","empid","employee id", "employeeid",
            "employee code","emp id","member code"
        )) for h in headers
    )
    has_dob = any(
        "date of birth" in h or "dateofbirth" in h or h == "dob" or "birth date" in h
        for h in headers
    )
    has_rel = any("relationship" in h or "relation" in h or "dependent type" in h for h in headers)
    has_name = any("name" in h for h in headers)

    if has_id and has_dob and (has_rel or has_name):
        return "long", best_row

    if has_name and has_dob and (spouse or child or parent):
        return "wide_tetris", best_row

    return "unknown", best_row


def normalize_phone(value):
    """Normalize a phone/contact value while preserving valid digits."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, float):
        if pd.isna(value):
            return ""
        if value.is_integer():
            return str(int(value))
    text = clean_text(value)
    if not text:
        return ""
    plus = text.startswith("+")
    digits = re.sub(r"\D", "", text)
    if len(digits) < 8 or len(digits) > 15:
        return ""
    return ("+" if plus else "") + digits


def looks_like_phone(value):
    if value is None or isinstance(value, (pd.Timestamp, datetime, date)):
        return False
    text = clean_text(value)
    if not text:
        return False
    digits = re.sub(r"\D", "", text)
    return 8 <= len(digits) <= 15


def find_phone_column(headers, data):
    """
    Find a phone/mobile/contact column from its header. If the header is
    missing or generic, inspect the actual values across all columns.
    No fixed column number is used.
    """
    h = [norm(x) for x in headers]

    strong = []
    for i, x in enumerate(h):
        if any(k in x for k in (
            "mobile number", "mobile no", "mobile", "phone number",
            "phone no", "telephone", "contact number", "contact no",
            "contact", "phone"
        )):
            if not any(k in x for k in (
                "employee id", "member id", "emp id",
                "number of children", "sum insured", "age"
            )):
                strong.append(i)
    if strong:
        return strong[0]

    # Header-independent fallback.
    candidates = []
    for i, x in enumerate(h):
        if any(k in x for k in (
            "id", "dob", "date", "age", "children", "child",
            "sum insured", "sr no", "serial", "s.no"
        )):
            continue
        values = [v for v in data.iloc[:, i].tolist() if clean_text(v) != ""]
        if not values:
            continue
        matches = sum(looks_like_phone(v) for v in values)
        ratio = matches / len(values)
        if matches >= 1 and ratio >= 0.60:
            candidates.append((ratio, matches, i))
    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][2]
    return None


def make_long_column_map(headers):
    h = [norm(x) for x in headers]

    def find(*terms):
        for i, x in enumerate(h):
            if any(t in x for t in terms):
                return i
        return None

    return {
        "id": find("memberid", "empid", "employeeid", "employee code", "employee id"),
        "name": find("insured name", "insuredname", "member name", "membername", "name"),
        "dob": find("date of birth", "dateofbirth", "dob"),
        "age": find("completed age", "age"),
        "gender": find("gender", "sex"),
        "rel": find("relationship to self", "relationship", "relation 2", "relation"),
        "si": find("ghi sum insured", "sum insured", "suminsured"),
        "phone": find(
            "mobile number", "mobile no", "mobile",
            "phone number", "phone no", "telephone",
            "contact number", "contact no", "contact"
        ),
    }

def process_long(raw, header_row, client_hint):
    headers = [clean_text(x) for x in raw.iloc[header_row].tolist()]
    data = raw.iloc[header_row + 1:].copy()
    cmap = make_long_column_map(headers)

    if cmap["id"] is None or cmap["dob"] is None or cmap["rel"] is None:
        raise ValueError(
            "Could not identify Employee ID, Date of Birth and Relationship columns."
        )

    phone_idx = cmap["phone"]
    if phone_idx is None:
        phone_idx = find_phone_column(headers, data)

    logs = []
    out = []

    for excel_row, (_, row) in enumerate(data.iterrows(), start=header_row + 2):
        raw_id = row.iloc[cmap["id"]]
        raw_name = row.iloc[cmap["name"]] if cmap["name"] is not None else ""
        raw_dob = row.iloc[cmap["dob"]]
        raw_gender = row.iloc[cmap["gender"]] if cmap["gender"] is not None else ""
        raw_rel = row.iloc[cmap["rel"]]
        raw_phone = row.iloc[phone_idx] if phone_idx is not None else ""
        raw_si = row.iloc[cmap["si"]] if cmap["si"] is not None else np.nan

        if clean_text(raw_id) == "" and clean_text(raw_name) == "":
            continue

        empid = clean_text(raw_id)
        name = clean_text(raw_name)
        dob = parse_date(raw_dob)
        gender = normalize_gender(raw_gender)
        rel = normalize_relationship(raw_rel)

        if isinstance(raw_name, str) and name != raw_name:
            log_correction(
                logs, excel_row, "Member Name", raw_name, name,
                "Removed extra/hidden whitespace and standardized text spacing.", "High"
            )
        if gender != clean_text(raw_gender):
            log_correction(
                logs, excel_row, "Gender", raw_gender, gender,
                "Normalized gender alias/abbreviation/typo."
            )
        if rel != clean_text(raw_rel):
            log_correction(
                logs, excel_row, "Relationship to Self", raw_rel, rel,
                "Normalized relationship alias."
            )

        out.append({
            "Employee ID": empid,
            "Member Name": name,
            "Date of Birth": dob,
            "Age": clean_text(row.iloc[cmap["age"]]) if cmap["age"] is not None else "",
            "Gender": gender,
            "Relationship to Self": rel,
            "GHI Sum Insured": normalize_si(raw_si),
            "Mobile Number": normalize_phone(raw_phone),
            "_source_row": excel_row,
            "_gender_source": "source",
            "_employee_id_generated": False,
        })

    df = pd.DataFrame(out)
    if df.empty:
        raise ValueError("No member rows were found in this sheet.")

    # If a long-format family has Sum Insured on only one member row,
    # propagate that family-level value to the other members.
    if "GHI Sum Insured" in df.columns and "Employee ID" in df.columns:
        df["GHI Sum Insured"] = (
            df.groupby("Employee ID", dropna=False)["GHI Sum Insured"]
              .transform(lambda s: s.ffill().bfill())
        )

    # AGE RULE:
    # 1) If DOB exists, calculate current age from DOB.
    # 2) If DOB is missing, retain a valid source Age.
    # 3) If both are missing, show "No age mentioned".
    calculated_age = df["Date of Birth"].apply(age_from_dob)
    source_age = pd.to_numeric(df["Age"], errors="coerce")
    resolved_age = calculated_age.where(calculated_age.notna(), source_age)
    df["Age"] = resolved_age.apply(
        lambda x: int(x) if pd.notna(x) else "No age mentioned"
    )
    style = "hyphen" if client_hint != "bambrew" else "slash"
    df["Age Band"] = df["Age"].apply(
        lambda x: age_band(x, style)
        if isinstance(x, (int, float, np.integer, np.floating)) and pd.notna(x)
        else ""
    )
    if "Mobile Number" not in df.columns:
        df["Mobile Number"] = ""
    df["Email"] = ""

    return df, logs, {
        "orientation": "Long / already member-level",
        "notes": "Rows preserved; relationship values normalized; Age and Age Band recalculated from DOB; phone/mobile/contact values preserved when found."
    }

def process_tetris(raw, header_row, runtime_config=None):
    """Smart header-driven wide-to-long parser; no fixed column positions."""
    runtime_config = runtime_config or {}
    headers = [clean_text(x) for x in raw.iloc[header_row].tolist()]
    data = raw.iloc[header_row + 1:].copy()
    h = [norm(x) for x in headers]

    def name_header(x):
        return "name" in x or x in {"employee", "insured", "member"}

    def dob_header(x):
        return "date of birth" in x or x == "dob" or "birth date" in x

    def age_header(x):
        return x == "age" or x.endswith(" age") or " age " in f" {x} "

    def employee_header(x):
        return any(k in x for k in (
            "employee full name", "employee name", "employee",
            "insured name", "member name"
        ))

    def child_no(x):
        m = re.search(r"child\s*(\d+)", x)
        return int(m.group(1)) if m else None

    def nearest_field(name_idx, predicate):
        candidates = [i for i, x in enumerate(h) if predicate(x)]
        right = [i for i in candidates if i > name_idx]
        if right:
            next_name = [i for i in range(name_idx + 1, len(h)) if name_header(h[i])]
            boundary = min(next_name) if next_name else len(h)
            bounded = [i for i in right if i < boundary]
            if bounded:
                return min(bounded, key=lambda i: i - name_idx)
            return min(right, key=lambda i: i - name_idx)
        left = [i for i in candidates if i < name_idx]
        return min(left, key=lambda i: name_idx - i) if left else None

    employee_candidates = [i for i, x in enumerate(h) if employee_header(x)]
    i_employee = employee_candidates[0] if employee_candidates else None
    if i_employee is None:
        raise ValueError("Could not identify the employee/member name column from the headers.")

    spouse_desc = [
        (i, nearest_field(i, dob_header), nearest_field(i, age_header))
        for i, x in enumerate(h) if "spouse" in x and name_header(x)
    ]

    child_desc = []
    for i, x in enumerate(h):
        if "child" in x and name_header(x):
            n = child_no(x)
            child_desc.append((
                n if n is not None else 999,
                i,
                nearest_field(i, dob_header),
                nearest_field(i, age_header)
            ))
    child_desc.sort(key=lambda x: (x[0], x[1]))

    parent_desc = []
    seen = set()
    for i, x in enumerate(h):
        if name_header(x) and any(k in x for k in (
            "father", "mother", "parent", "dad", "mom"
        )):
            if i not in seen:
                seen.add(i)
                parent_desc.append((
                    i,
                    nearest_field(i, dob_header),
                    nearest_field(i, age_header),
                    headers[i]
                ))

    employee_dob_idx = nearest_field(i_employee, dob_header)
    employee_age_idx = nearest_field(i_employee, age_header)

    # Family-level Sum Insured: locate it by header meaning, never by fixed
    # column position. The value on the family row is applied to every member
    # generated from that family row.
    si_candidates = [
        i for i, x in enumerate(h)
        if any(term in x for term in (
            "sum insured", "ghi sum insured", "ghi", "si amount",
            "insured amount", "coverage amount", "sum assured"
        ))
    ]
    si_idx = si_candidates[0] if si_candidates else None

    phone_idx = find_phone_column(headers, data)

    self_rows, spouse_rows, child_rows, parent_rows = [], [], {}, []
    logs = []
    employee_id = 1

    for excel_row, (_, row) in enumerate(data.iterrows(), start=header_row + 2):
        raw_name = row.iloc[i_employee]
        raw_dob = row.iloc[employee_dob_idx] if employee_dob_idx is not None else None
        raw_phone = row.iloc[phone_idx] if phone_idx is not None else ""

        name = clean_text(raw_name)
        dob = parse_date(raw_dob)
        if not name and dob is None:
            continue

        family_si = row.iloc[si_idx] if si_idx is not None else np.nan

        def make_member(member_name, member_dob=None, rel="Self",
                        source_age=None, source_label="", phone="", sum_insured=None):
            member_name = clean_text(member_name)
            if not member_name:
                return None

            parsed_dob = parse_date(member_dob)
            age_value = None
            if source_age is not None and clean_text(source_age):
                try:
                    age_value = float(source_age)
                    if age_value.is_integer():
                        age_value = int(age_value)
                except Exception:
                    age_value = None

            return {
                "Employee ID": employee_id,
                "Member Name": member_name,
                "Date of Birth": parsed_dob,
                "Age": age_value,
                "Gender": "",
                "Relationship to Self": rel,
                "GHI Sum Insured": normalize_si(sum_insured),
                "Mobile Number": normalize_phone(phone) if rel == "Self" else "",
                "_source_row": excel_row,
                "_gender_source": "not supplied",
                "_source_label": source_label,
                "_employee_id_generated": True,
            }

        self_member = make_member(
            name,
            dob,
            "Self",
            row.iloc[employee_age_idx] if employee_age_idx is not None else None,
            headers[i_employee],
            raw_phone,
            sum_insured=family_si
        )
        if self_member:
            self_rows.append(self_member)

        for ni, di, ai in spouse_desc:
            member = make_member(
                row.iloc[ni],
                row.iloc[di] if di is not None else None,
                "Spouse",
                row.iloc[ai] if ai is not None else None,
                headers[ni],
                sum_insured=family_si
            )
            if member:
                spouse_rows.append(member)

        for n, ni, di, ai in child_desc:
            member = make_member(
                row.iloc[ni],
                row.iloc[di] if di is not None else None,
                "Child",
                row.iloc[ai] if ai is not None else None,
                headers[ni],
                sum_insured=family_si
            )
            if member:
                child_rows.setdefault(n, []).append(member)

        for ni, di, ai, label in parent_desc:
            member = make_member(
                row.iloc[ni],
                row.iloc[di] if di is not None else None,
                "Parent",
                row.iloc[ai] if ai is not None else None,
                label,
                sum_insured=family_si
            )
            if member:
                parent_rows.append(member)

        employee_id += 1

    out = self_rows + spouse_rows
    for n in sorted(child_rows):
        out.extend(child_rows[n])
    out.extend(parent_rows)

    df = pd.DataFrame(out)
    if df.empty:
        raise ValueError("No member rows could be identified from the wide-format headers.")

    # AGE RULE:
    # 1) If DOB exists, calculate current age from DOB.
    # 2) If DOB is missing, retain a valid source Age.
    # 3) If both are missing, show "No age mentioned".
    calculated_age = df["Date of Birth"].apply(age_from_dob)
    source_age = pd.to_numeric(df["Age"], errors="coerce")
    resolved_age = calculated_age.where(calculated_age.notna(), source_age)
    df["Age"] = resolved_age.apply(
        lambda x: int(x) if pd.notna(x) else "No age mentioned"
    )
    df["Age Band"] = df["Age"].apply(
        lambda x: age_band(x, "hyphen")
        if isinstance(x, (int, float, np.integer, np.floating)) and pd.notna(x)
        else ""
    )
    if "Mobile Number" not in df.columns:
        df["Mobile Number"] = ""
    df["Email"] = ""

    if "_source_label" in df.columns:
        df.drop(columns=["_source_label"], inplace=True)

    return df, logs, {
        "orientation": "Wide → Long",
        "profile": "Smart header-driven wide-format parser",
        "header_detection": "Semantic header matching; no fixed column positions.",
        "parent_rule": "Father/Mother/Parent name fields are detected from header text and mapped to Parent.",
        "child_rule": "All Child N name headers are discovered dynamically.",
        "associated_fields": "DOB/Age are associated by header context within each member group.",
        "phone_rule": "Phone/mobile/contact is detected from headers or, if needed, from phone-like values across all columns.",
        "employee_id_rule": "Sequential ID based on source employee-row order."
    }

def process_bambrew(raw, header_row):
    headers = [clean_text(x) for x in raw.iloc[header_row].tolist()]
    data = raw.iloc[header_row + 1:].copy()
    h = [norm(x) for x in headers]

    def find(*terms):
        for i, x in enumerate(h):
            if any(t in x for t in terms):
                return i
        return None

    phone_idx = find_phone_column(headers, data)

    idx = {
        "id": find("emp id"),
        "self_name": find("self name"),
        "self_dob": find("self dob"),
        "spouse_name": find("spouse name"),
        "spouse_dob": find("spouse dob"),
        "child1_name": find("children 1", "child 1"),
        "child1_dob": find("children 1 dob", "child 1 dob"),
        "child2_name": find("children 2", "child 2"),
        "child2_dob": find("children 2 dob", "child 2 dob"),
        "mother_name": find("mother name"),
        "mother_dob": find("mother dob"),
        "father_name": find("father name"),
        "father_dob": find("father dob"),
        "si": find("ghi sum insured", "sum insured", "si amount", "insured amount", "coverage amount", "sum assured"),
    }
    if idx["id"] is None or idx["self_name"] is None:
        raise ValueError("Could not identify Bambrew Employee ID and Self Name columns.")

    logs = []
    out = []

    for excel_row, (_, row) in enumerate(data.iterrows(), start=header_row + 2):
        raw_empid = row.iloc[idx["id"]]
        empid = raw_empid
        if clean_text(raw_empid) == "":
            continue

        family_si = row.iloc[idx["si"]] if idx["si"] is not None else np.nan

        def add_member(name_i, dob_i, rel, gender=""):
            if name_i is None or dob_i is None:
                return
            raw_name = row.iloc[name_i]
            raw_dob = row.iloc[dob_i]
            if clean_text(raw_name) == "":
                return
            dob = parse_date(raw_dob)
            # DOB is normalized internally for the Corrected Data and Age calculation,
            # but it is not treated as an automatic correction in the audit.
            out.append({
                "Employee ID": empid,
                "Member Name": clean_text(raw_name),
                "Date of Birth": dob,
                "Gender": gender,
                "Relationship to Self": rel,
                "GHI Sum Insured": normalize_si(family_si),
                "Mobile Number": normalize_phone(row.iloc[phone_idx]) if (rel == "Self" and phone_idx is not None) else "",
                "_source_row": excel_row,
                "_gender_source": "role-derived" if gender else "not supplied",
                "_employee_id_generated": False,
            })

        add_member(idx["self_name"], idx["self_dob"], "Self", "")
        add_member(idx["spouse_name"], idx["spouse_dob"], "Spouse", "")
        add_member(idx["child1_name"], idx["child1_dob"], "Child", "")
        add_member(idx["child2_name"], idx["child2_dob"], "Child", "")
        # Mother/Father are deterministically gendered by their source column.
        add_member(idx["mother_name"], idx["mother_dob"], "Parent", "Female")
        add_member(idx["father_name"], idx["father_dob"], "Parent", "Male")

    df = pd.DataFrame(out)
    if df.empty:
        raise ValueError("No Bambrew member rows were found.")

    # AGE RULE:
    # Calculate current age from DOB; if DOB is absent, show "No age mentioned".
    calculated_age = df["Date of Birth"].apply(age_from_dob)
    df["Age"] = calculated_age.apply(
        lambda x: int(x) if pd.notna(x) else "No age mentioned"
    )
    df["Age Band"] = df["Age"].apply(
        lambda x: age_band(x, "slash")
        if isinstance(x, (int, float, np.integer, np.floating)) and pd.notna(x)
        else ""
    )
    if "Mobile Number" not in df.columns:
        df["Mobile Number"] = ""
    df["Email"] = ""

    # Match the supplied final workbook's mixed-type Excel sort:
    # numeric Employee IDs first in numeric ascending order, then text IDs lexicographically.
    def bambrew_sort_key(v):
        if isinstance(v, (int, float, np.integer, np.floating)) and not pd.isna(v):
            return (0, float(v), "")
        return (1, 0, str(v))

    df["_sort_key"] = df["Employee ID"].map(bambrew_sort_key)
    df = df.sort_values("_sort_key", kind="stable").drop(columns="_sort_key").reset_index(drop=True)

    return df, logs, {
        "orientation": "Wide → Long",
        "profile": "Bambrew",
                "included_relationships": "Self, Spouse, Child, Mother/Father as Parent",
        "excluded_relationships": "Mother-in-law/Father-in-law columns are not carried into the supplied final workbook.",
        "employee_id_rule": "Use source Emp Id as Employee ID; preserve leading zeros when Excel stores it as text.",
        "gender_rule": "Mother/Father gender is deterministic from the source column. Self/Spouse/Child gender was manually populated in the supplied final workbook and cannot be safely reproduced from the raw columns alone; the tool leaves those blank for review."
    }

def add_validation(df, relations, config=None):
    df = df.copy()
    config = config or {}
    relations = {k: tuple(v) for k, v in relations.items()}
    skip_configured_rules = config.get("skip_validation_rules", False)

    # These limits are the user-configurable rules in the UI. They are
    # intentionally excluded when the user elects not to apply UI rules.
    if not skip_configured_rules:
        for rel, key in [("Child", "child_limit"), ("Spouse", "spouse_limit"),
                         ("Parent", "parent_limit"), ("Parent-in-law", "parent_inlaw_limit")]:
            if config.get(key) not in (None, ""):
                lo, hi, _ = relations.get(rel, (0, 100, 0))
                relations[rel] = (lo, hi, int(config[key]))

    remarks = [clean_text(x) if "_pre_remark" in df.columns else "" for x in df.get(
        "_pre_remark", pd.Series([""] * len(df), index=df.index))]
    validation_errors = []
    # Error-heavy workbooks can generate many remarks. Resolve index positions
    # once instead of scanning the entire DataFrame index for every error.
    index_positions = {}
    for position, index_value in enumerate(df.index):
        index_positions.setdefault(index_value, position)

    def add_remark(idx, text):
        pos = index_positions[idx]
        remarks[pos] = (remarks[pos] + " " + text).strip() if remarks[pos] else text

    def add_error(idx, rule, field, message, severity="Error"):
        row = df.loc[idx]
        validation_errors.append({
            "Employee ID": row.get("Employee ID", ""),
            "Source Row": row.get("_source_row", ""),
            "Field": field,
            "Rule": rule,
            "Severity": severity,
            "Value": row.get(field, "") if field in df.columns else "",
            "Message": message,
        })
        add_remark(idx, message)

    # Field-level validation.
    for idx, r in df.iterrows():
        if not clean_text(r.get("Member Name", "")):
            add_error(idx, "Member name required", "Member Name", "Member Name is missing.")
        if r.get("Date of Birth") is None:
            add_error(idx, "DOB required", "Date of Birth", "Date of Birth is missing/incomplete.")
        gender = clean_text(r.get("Gender", ""))
        if gender and gender not in GENDERS:
            add_error(idx, "Gender normalization/validation", "Gender", "Invalid gender value.")
        rel = r.get("Relationship to Self", "")
        if rel not in relations:
            add_error(
                idx,
                "Relationship hierarchy/code validation",
                "Relationship to Self",
                "Relationship value/code could not be mapped to a supported hierarchy."
            )
        if rel in {"Parent", "Parent-in-law"} and parent_name_indicates_unavailable(
            r.get("Member Name", "")
        ):
            add_error(
                idx,
                "Deceased/unavailable parent name",
                "Member Name",
                "Parent or Parent-in-law name indicates a deceased or unavailable member; please remove or replace this dependent with valid details.",
                "Review",
            )
        elif (
            clean_text(r.get("Age", "")) != "No age mentioned"
            and pd.to_numeric(r.get("Age"), errors="coerce") == pd.to_numeric(r.get("Age"), errors="coerce")
        ):
            lo, hi, _ = relations[rel]
            age_value = float(r["Age"])
            if not (lo <= age_value <= hi):
                add_error(idx, "Relationship age limit", "Age", f"{rel} age should be between {lo} and {hi}.")

    # Employee-level validation and deterministic spouse-gender inference.
    #
    # Rule:
    #   Self = Male + blank Spouse gender -> Spouse = Female
    #   Self = Female + blank Spouse gender -> Spouse = Male
    #
    # Existing spouse gender is never overwritten.
    for empid, group in df.groupby(df["Employee ID"].astype(str)):
        self_rows = group[group["Relationship to Self"] == "Self"]
        spouse_rows = group[group["Relationship to Self"] == "Spouse"]

        if len(self_rows) and len(spouse_rows):
            self_gender = normalize_gender(self_rows.iloc[0].get("Gender", ""))
            inferred_gender = "Female" if self_gender == "Male" else (
                "Male" if self_gender == "Female" else ""
            )

            if inferred_gender:
                for idx in spouse_rows.index:
                    if not clean_text(df.loc[idx, "Gender"]):
                        df.loc[idx, "Gender"] = inferred_gender
                        add_remark(
                            idx,
                            f"Spouse gender inferred as {inferred_gender} from Self gender."
                        )

    # Employee-level validation.
    for empid, group in df.groupby(df["Employee ID"].astype(str)):
        if not skip_configured_rules:
            rel_counts = group["Relationship to Self"].value_counts().to_dict()
            for rel, count in rel_counts.items():
                if rel in relations and count > relations[rel][2]:
                    for idx in group[group["Relationship to Self"] == rel].index:
                        add_error(idx, f"{rel} count limit", "Relationship to Self", f"{rel} count ({count}) exceeds the allowed limit of {relations[rel][2]}.")

            total_dependents = sum(v for k, v in rel_counts.items() if k != "Self")
            if config.get("max_dependents") not in (None, "") and total_dependents > int(config["max_dependents"]):
                for idx in group.index:
                    if group.loc[idx, "Relationship to Self"] != "Self":
                        add_error(idx, "Total dependent count", "Relationship to Self", f"Total dependents ({total_dependents}) exceeds the configured maximum of {int(config['max_dependents'])}.")

            if rel_counts.get("Parent", 0) and rel_counts.get("Parent-in-law", 0):
                if config.get("allow_parent_and_inlaw_together") is False:
                    for idx in group.index:
                        if group.loc[idx, "Relationship to Self"] in {"Parent", "Parent-in-law"}:
                            add_error(idx, "Parent vs Parent-in-law conflict", "Relationship to Self", "Both Parent and Parent-in-law details are present. Parent has been considered as the primary relationship; please confirm with the employee which dependent details should be considered.", "Review")

            if config.get("spouse_same_gender") is False:
                sr = group[group["Relationship to Self"] == "Self"]
                spr = group[group["Relationship to Self"] == "Spouse"]
                if len(sr) and len(spr):
                    sg = clean_text(sr.iloc[0]["Gender"]).lower()
                    for idx in spr.index:
                        if sg and clean_text(group.loc[idx, "Gender"]).lower() == sg:
                            add_error(idx, "Same-gender spouse rule", "Gender", "Spouse has the same gender as Self and same-gender spouse is not allowed.")

            if config.get("parent_same_gender") is False:
                for rel in ("Parent", "Parent-in-law"):
                    rows = group[group["Relationship to Self"] == rel]
                    if len(rows) == 2:
                        genders = [clean_text(x).lower() for x in rows["Gender"]]
                        if genders[0] and genders[0] == genders[1]:
                            for idx in rows.index:
                                add_error(idx, "Same-gender parent rule", "Gender", f"Same gender is not allowed for two {rel} members under the configured rule.")

            sr = group[group["Relationship to Self"] == "Self"]
            spouse_rows = group[group["Relationship to Self"] == "Spouse"]
            if len(sr) or len(spouse_rows):
                self_age = sr.iloc[0]["Age"] if len(sr) else None
                for idx, row in group.iterrows():
                    age = row["Age"]
                    if age is None:
                        continue
                    rel = row["Relationship to Self"]
                    if rel == "Child" and config.get("child_gap") not in (None, "") and pd.notna(pd.to_numeric(self_age, errors="coerce")) and pd.notna(pd.to_numeric(age, errors="coerce")) and float(pd.to_numeric(self_age, errors="coerce")) - float(pd.to_numeric(age, errors="coerce")) < float(config["child_gap"]):
                        add_error(idx, "Self-to-Child age gap", "Age", f"Minimum Self-to-Child age gap of {config['child_gap']} years is not met.")
                    if rel == "Parent" and config.get("parent_gap") not in (None, "") and pd.notna(pd.to_numeric(self_age, errors="coerce")) and pd.notna(pd.to_numeric(age, errors="coerce")) and float(pd.to_numeric(age, errors="coerce")) - float(pd.to_numeric(self_age, errors="coerce")) < float(config["parent_gap"]):
                        add_error(idx, "Self-to-Parent age gap", "Age", f"Minimum Self-to-Parent age gap of {config['parent_gap']} years is not met.")
                    if rel == "Parent-in-law" and config.get("parent_inlaw_gap") not in (None, "") and len(spouse_rows):
                        spouse_age = spouse_rows.iloc[0]["Age"]
                        if pd.notna(pd.to_numeric(spouse_age, errors="coerce")) and pd.notna(pd.to_numeric(age, errors="coerce")) and float(pd.to_numeric(age, errors="coerce")) - float(pd.to_numeric(spouse_age, errors="coerce")) < float(config["parent_inlaw_gap"]):
                            add_error(idx, "Spouse-to-Parent-in-law age gap", "Age", f"Minimum Spouse-to-Parent-in-law age gap of {config['parent_inlaw_gap']} years is not met.")

    df["Remark"] = remarks
    if "_pre_remark" in df.columns:
        df.drop(columns=["_pre_remark"], inplace=True)
    # Keep the individual validation columns used by the earlier tool.
    # Only the single combined Error column is intentionally removed.
    def err_name(r):
        return "Please enter Member Name." if not clean_text(r.get("Member Name", "")) else ""

    def err_dob(r):
        return "DoB is incomplete." if r.get("Date of Birth") is None else ""

    def err_gender(r):
        g = clean_text(r.get("Gender", ""))
        return "Invalid gender." if g and g not in GENDERS else ""

    def err_rel(r):
        return "Invalid relationship to employee." if r.get("Relationship to Self") not in relations else ""

    def err_age(r):
        rel = r.get("Relationship to Self")
        age = r.get("Age")
        if rel not in relations or age is None:
            return ""
        if clean_text(age) == "No age mentioned":
            return ""
        age_num = pd.to_numeric(age, errors="coerce")
        if pd.isna(age_num):
            return ""
        lo, hi, _ = relations[rel]
        lo_num = float(pd.to_numeric(lo, errors="coerce"))
        hi_num = float(pd.to_numeric(hi, errors="coerce"))
        age_num = float(age_num)
        return "" if lo_num <= age_num <= hi_num else f"{rel} age should be between {int(lo_num) if lo_num.is_integer() else lo_num} and {int(hi_num) if hi_num.is_integer() else hi_num}."

    # Contact validation was removed from the rules. The legacy individual
    # columns remain available for compatibility and are intentionally blank.
    df["Error_Name"] = df.apply(err_name, axis=1)
    df["Error_DoB"] = df.apply(err_dob, axis=1)
    df["Error_Age"] = df.apply(err_age, axis=1)
    df["Error_Gender"] = df.apply(err_gender, axis=1)
    df["Error_Relationship"] = df.apply(err_rel, axis=1)
    df["Error_Ph"] = ""
    df["Error_Email"] = ""
    df.drop(columns=["Error"], errors="ignore", inplace=True)
    return df, validation_errors, relations

def build_count(df, relations, validation_enabled=True):
    ids = list(dict.fromkeys(df["Employee ID"].tolist()))
    employee_keys = df["Employee ID"].astype(str)
    # Cache each string-equivalent employee group once. This intentionally
    # retains the prior behavior where values that stringify identically are
    # counted together, while preserving the original output ID order.
    employee_groups = {
        key: group for key, group in df.groupby(employee_keys, sort=False)
    }
    rows = []
    for eid in ids:
        sub = employee_groups[str(eid)]
        counts = {rel: int((sub["Relationship to Self"] == rel).sum()) for rel in relations}
        self_ok = counts.get("Self", 0) == relations["Self"][2]
        spouse_ok = counts.get("Spouse", 0) <= relations["Spouse"][2]
        child_ok = counts.get("Child", 0) <= relations["Child"][2]
        parent_ok = counts.get("Parent", 0) <= relations["Parent"][2]
        inlaw_ok = counts.get("Parent-in-law", 0) <= relations["Parent-in-law"][2]
        cross_ok = not (counts.get("Parent", 0) > 0 and counts.get("Parent-in-law", 0) > 0)
        valid = self_ok and spouse_ok and child_ok and parent_ok and inlaw_ok and cross_ok

        rows.append({
            "MEMBERID": eid,
            "Self": counts.get("Self", 0),
            "Spouse": counts.get("Spouse", 0),
            "Child": counts.get("Child", 0),
            "Parent": counts.get("Parent", 0),
            "Parent-in-law": counts.get("Parent-in-law", 0),
            "Validate": ("Y" if valid else "INVALID") if validation_enabled else "Not checked",
        })
    return pd.DataFrame(rows)

def write_sheet(ws, df):
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws.append(row)

    header_fill = PatternFill("solid", fgColor="D9EAF7")

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30

    for col in range(1, ws.max_column + 1):
        values = [
            str(ws.cell(r, col).value or "")
            for r in range(1, min(ws.max_row, 150) + 1)
        ]
        width = min(max(max((len(v) for v in values), default=10) + 2, 10), 45)
        ws.column_dimensions[get_column_letter(col)].width = width

    # Date column.
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(1, c).value).lower() == "date of birth":
            for r in range(2, ws.max_row + 1):
                ws.cell(r, c).number_format = "dd-mm-yyyy"
        if str(ws.cell(1, c).value).lower() == "ghi sum insured":
            for r in range(2, ws.max_row + 1):
                # Indian digit grouping: 2,00,000 / 2,00,00,000.
                ws.cell(r, c).number_format = "#,##,##0"

def build_output_workbook(results, raw_sheets, runtime_config):
    wb = Workbook(); wb.remove(wb.active)
    def unique_name(desired):
        name=re.sub(r'[:\\/?*\[\]]','-',str(desired)).strip() or 'Sheet'; name=name[:31]
        base=name; n=2
        while name in wb.sheetnames:
            suf=f" ({n})"; name=base[:31-len(suf)]+suf; n+=1
        return name
    for source_sheet,result in results.items():
        raw_ws = wb.create_sheet(unique_name(f"{source_sheet} (Raw Data)")); raw_df=raw_sheets[source_sheet]
        # Stream rows directly to OpenPyXL; avoid allocating one giant nested
        # Python list for large source tabs.
        for row in raw_df.itertuples(index=False, name=None): raw_ws.append(row)
        if raw_ws.max_row:
            for cell in raw_ws[1]: cell.font=Font(bold=True); cell.fill=PatternFill("solid",fgColor="EDEDED")
        raw_ws.freeze_panes="A2"
        if result["status"]!="Processed":
            continue
        prefix="" if len(results)==1 else f" - {source_sheet}"
        ws=wb.create_sheet(unique_name("Corrected Data"+prefix)); write_sheet(ws,result["final"][FINAL_COLUMNS])
        cw=wb.create_sheet(unique_name("Corrections Made"+prefix)); corr_cols=["Source Row","Field","Original Value","Corrected Value","Rule Applied","Confidence"]
        write_sheet(cw,pd.DataFrame(result["corrections"])[corr_cols] if result["corrections"] else pd.DataFrame(columns=corr_cols))
        ew=wb.create_sheet(unique_name("To be Corrected"+prefix)); err_cols=["Employee ID","Source Row","Field","Rule","Severity","Value","Message"]
        write_sheet(ew,pd.DataFrame(result["validation_errors"])[err_cols] if result["validation_errors"] else pd.DataFrame(columns=err_cols))
        qw=wb.create_sheet(unique_name("Count"+prefix)); write_sheet(qw,result["count"])

        # Age Band Pivot Table
        pivot_name = unique_name("Age Band Pivot Table"+prefix)
        pws = wb.create_sheet(pivot_name)

        age_band_series = result["final"]["Age Band"].fillna("").astype(str).str.strip()
        age_band_series = age_band_series[age_band_series != ""]

        if len(age_band_series):
            pivot = (
                age_band_series.value_counts()
                .rename_axis("Age Band")
                .reset_index(name="Total Members")
            )

            def age_band_sort_key(value):
                m = re.search(r"(\d+)", str(value))
                return int(m.group(1)) if m else 999999

            pivot["_sort"] = pivot["Age Band"].map(age_band_sort_key)
            pivot = pivot.sort_values(["_sort", "Age Band"]).drop(columns="_sort")
            pivot.loc[len(pivot)] = ["Grand Total", int(pivot["Total Members"].sum())]
        else:
            pivot = pd.DataFrame({"Age Band": ["Grand Total"], "Total Members": [0]})

        write_sheet(pws, pivot)
        for cell in pws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
        if pws.max_row >= 2:
            for cell in pws[pws.max_row]:
                cell.font = Font(bold=True)
        pws.column_dimensions["A"].width = 22
        pws.column_dimensions["B"].width = 18
        pws.freeze_panes = "A2"

    vw=wb.create_sheet("Validation rules", 0); vw.append(["Rule","Configured Value"])
    # When validation is bypassed, none of the configured limits or checks
    # apply to this run. Show zero for each rule so the audit tab accurately
    # reflects that no validation thresholds were active.
    validation_skipped = runtime_config.get("skip_validation_rules", False)
    for k,v in runtime_config.items():
        # Keep the bypass flag visible as TRUE; all validation rules themselves
        # show 0 because they were not applied for this run.
        configured_value = v if k == "skip_validation_rules" else (0 if validation_skipped else v)
        vw.append([k, configured_value])
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

def process_one_sheet(raw, sheet_name, filename, runtime_config=None, wide_si=None):
    runtime_config = runtime_config or {}
    kind, header_row = classify_sheet(raw)
    lower_file = filename.lower()
    client_hint = "tetris" if "tetris" in lower_file else ("bambrew" if "bambrew" in lower_file else "")

    if kind == "wide_tetris":
        final, logs, meta = process_tetris(raw, header_row, runtime_config=runtime_config)
        relations = VALID_RELATIONS
    elif kind == "wide_bambrew":
        final, logs, meta = process_bambrew(raw, header_row)
        relations = BAMBrew_VALID_RELATIONS
    elif kind == "long":
        final, logs, meta = process_long(raw, header_row, client_hint)
        relations = VALID_RELATIONS
    else:
        raise ValueError("This sheet does not match a recognized member-data structure. Select a member-data tab or add a mapping rule for this source format.")

    # Runtime GHI/limits override learned client defaults.
    # Preserve gender values before validation so deterministic spouse-gender
    # inference can be recorded in the Corrections audit.
    _gender_before_validation = final.get(
        "Gender", pd.Series([""] * len(final), index=final.index)
    ).copy()

    # Built-in data-quality validation always runs. The checkbox only skips
    # the configurable business rules displayed in the UI.
    validation_enabled = not runtime_config.get("skip_validation_rules", False)
    final, validation_errors, relations = add_validation(final, relations, runtime_config)

    if "_source_row" not in final.columns:
        final["_source_row"] = np.arange(len(final)) + header_row + 2

    # Audit deterministic spouse-gender inference as a correction.
    if "Gender" in final.columns:
        for _pos, (_before, _after) in enumerate(
            zip(_gender_before_validation.tolist(), final["Gender"].tolist())
        ):
            if not clean_text(_before) and clean_text(_after) in {"Male", "Female"}:
                _rel = clean_text(final.iloc[_pos].get("Relationship to Self", ""))
                if _rel == "Spouse":
                    logs.append({
                        "Source Row": final.iloc[_pos].get("_source_row", ""),
                        "Field": "Gender",
                        "Original Value": "",
                        "Corrected Value": clean_text(_after),
                        "Rule Applied": "Self gender used to infer blank Spouse gender.",
                        "Confidence": "High",
                    })

    correction_map = {}
    for c in logs:
        correction_map.setdefault(c["Source Row"], []).append(c)

    # Build correction summaries in one pass rather than three DataFrame-wide
    # `.apply()` calls. `Correction Confidence` was never exported, so avoid
    # calculating that discarded column for large workbooks.
    source_log_items = [
        correction_map.get(source_row, []) for source_row in final["_source_row"]
    ]
    final["Corrections"] = [
        "; ".join(
            f"{item['Field']}: {item['Original Value']} → {item['Corrected Value']}"
            for item in items
        )
        for items in source_log_items
    ]
    final["Correction Rules"] = [
        "; ".join(dict.fromkeys(item["Rule Applied"] for item in items))
        for items in source_log_items
    ]

    # Preserve whether Employee ID was supplied by the source or generated by the tool.
    generated_id = final.get(
        "_employee_id_generated",
        pd.Series(False, index=final.index)
    ).fillna(False).astype(bool).reset_index(drop=True)

    # Construct the complete final output first.
    final = final[FINAL_COLUMNS + ["_source_row"]].reset_index(drop=True)
    final["_employee_id_generated"] = generated_id

    # FINAL DATA STEP:
    # After transformation, normalization, corrections, validation, remarks,
    # and final column construction are complete, sort Employee ID A-to-Z
    # using natural alphanumeric order.
    if "Employee ID" in final.columns:
        final["_employee_id_sort"] = final["Employee ID"].map(employee_id_sort_key)
        final = final.sort_values(
            "_employee_id_sort", kind="stable"
        ).drop(columns="_employee_id_sort").reset_index(drop=True)

    # Count before display labels are added, so all members retain the same ID.
    count = build_count(final, relations, validation_enabled=validation_enabled)

    # Only tool-generated IDs receive the explanatory label.
    # Source Employee ID / Employee Code values are left untouched.
    generated_mask = final["_employee_id_generated"].astype(bool)
    final["Employee ID"] = final["Employee ID"].astype(object)
    final.loc[generated_mask, "Employee ID"] = (
        final.loc[generated_mask, "Employee ID"]
        .map(lambda v: f"{clean_text(v)} (Given by Tool)")
    )

    meta["detected_kind"] = kind
    meta["source_sheet"] = sheet_name
    return final.drop(columns=["_source_row", "_employee_id_generated"]), logs, validation_errors, count, meta


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(
    page_title="Data Validation Tool",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Bright indigo/coral visual system.
st.markdown("""
<style>
:root {
    --primary-color: #08a6df;
    --dt-ink: #456b7d;
    --dt-muted: #7893a0;
    --dt-violet: #a7dfef;
    --dt-indigo: #c9effa;
    --dt-coral: #e0f4fa;
    --dt-surface: #ffffff;
    --dt-canvas: #f7fbfd;
    --dt-line: #e3edf1;
    --dt-shadow: rgba(121, 167, 184, .16);
}

.stApp,
[data-testid="stAppViewContainer"],
[data-testid="stAppViewBlockContainer"],
main {
    background: radial-gradient(circle at 8% 1%, #e4f7fc 0, transparent 28%), radial-gradient(circle at 92% 16%, #f0f8fa 0, transparent 25%), var(--dt-canvas) !important;
    color: var(--dt-ink) !important;
}

[data-testid="stHeader"] {
    background: rgba(250,253,254,.92) !important;
    border-bottom: 1px solid var(--dt-line) !important;
    backdrop-filter: blur(16px);
}

.block-container {
    max-width: 1320px !important;
    padding-top: 2.8rem !important;
    padding-bottom: 4rem !important;
}

h1, h2, h3, h4, h5, h6,
p, label, [data-testid="stMarkdownContainer"] {
    color: var(--dt-ink) !important;
}

h1 {
    font-size: 2.55rem !important;
    font-weight: 800 !important;
    letter-spacing: -0.045em !important;
}

h2 {
    font-size: 1.35rem !important;
    font-weight: 750 !important;
}

.hs-title {
    background: #13293d;
    border: 1px solid #13293d;
    border-radius: 20px;
    padding: 38px 42px 36px;
    margin: 0;
    box-shadow: 0 16px 34px rgba(19, 41, 61, .20);
    position: relative;
    min-height: 146px;
}

.hs-title h1,
.hs-title p {
    color: #ffffff !important;
    margin: 0 !important;
}

.hs-brandmark {
    position: absolute;
    right: 30px;
    bottom: 24px;
    display: flex;
    align-items: end;
    gap: 10px;
}

.hs-brandmark span {
    padding-bottom: 2px;
    color: #e6e8ea;
    font-size: .78rem;
    font-weight: 600;
    letter-spacing: .02em;
}

.hs-brandmark img {
    display: block;
    width: 142px;
    height: auto;
}

.upload-header-label {
    margin: 3px 0 10px;
    color: var(--dt-ink);
    font-size: 1.02rem;
    font-weight: 700;
}

.upload-next-step {
    margin-top: 10px;
    padding: 10px 12px;
    color: var(--dt-muted);
    background: #ffffff;
    border: 1px solid var(--dt-line);
    border-radius: 10px;
    font-size: .87rem;
    line-height: 1.35;
}

.hs-title h1 {
    font-size: 2.45rem !important;
    letter-spacing: -0.045em;
}

.hs-title p {
    margin-top: 8px !important;
    color: #e6e8ea !important;
    opacity: .94;
}

[data-testid="stFileUploader"] {
    background: rgba(255,255,255,.78) !important;
    border: 2px dashed #b9e4f0 !important;
    border-radius: 18px !important;
    padding: 10px !important;
    box-shadow: 0 10px 26px var(--dt-shadow) !important;
}

[data-testid="stFileUploader"] button {
    background: #edf8fb !important;
    color: var(--dt-ink) !important;
    border: 1px solid #cceaf2 !important;
    border-radius: 10px !important;
}

.stButton > button,
.stDownloadButton > button {
    width: 100%;
    min-height: 50px;
    background: linear-gradient(105deg, #d9f4fb, #c9effa 55%, #eaf7fb) !important;
    color: var(--dt-ink) !important;
    border: 1px solid #bfe4ef !important;
    border-radius: 14px !important;
    font-weight: 700 !important;
    letter-spacing: 0.01em !important;
    box-shadow: 0 8px 18px rgba(137, 183, 199, .18) !important;
    transition: transform .15s ease, box-shadow .15s ease !important;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 12px 24px rgba(137, 183, 199, .24) !important;
}

[data-baseweb="input"] > div,
[data-baseweb="select"] > div,
[data-baseweb="textarea"] > div,
input,
textarea {
    background: rgba(255,255,255,.93) !important;
    color: var(--dt-ink) !important;
    border-color: var(--dt-line) !important;
    border-radius: 12px !important;
}

[data-baseweb="select"] > div:focus-within,
[data-baseweb="input"] > div:focus-within,
[data-baseweb="textarea"] > div:focus-within {
    border-color: #b9e4f0 !important;
    box-shadow: 0 0 0 3px rgba(185,228,240,.26) !important;
}

[data-testid="stCheckbox"] label,
[data-testid="stRadio"] label {
    color: var(--dt-ink) !important;
}

[data-testid="stCheckbox"] {
    min-height: 62px;
    display: flex;
    align-items: center;
    padding: 10px 12px;
    background: rgba(255,255,255,.84);
    border: 1px solid var(--dt-line);
    border-radius: 12px;
    box-shadow: 0 5px 13px rgba(137,183,199,.09);
}

/* Streamlit/BaseWeb selected-sheet chips and checked boxes. */
[data-baseweb="tag"],
[data-testid="stMultiSelect"] span[data-baseweb="tag"],
[data-testid="stMultiSelect"] [data-baseweb="tag"] {
    background-color: #08a6df !important;
    border-color: #08a6df !important;
}

[data-baseweb="tag"] *,
[data-baseweb="tag"] svg,
[data-testid="stMultiSelect"] [data-baseweb="tag"] * {
    color: #ffffff !important;
    fill: #ffffff !important;
}

[data-baseweb="checkbox"] input[type="checkbox"] {
    accent-color: #08a6df !important;
}

[data-baseweb="checkbox"] div[role="checkbox"][aria-checked="true"],
[data-testid="stCheckbox"] div[role="checkbox"][aria-checked="true"],
[data-testid="stCheckbox"] label:has(input:checked) div[role="checkbox"],
[data-testid="stCheckbox"] input:checked + div {
    background-color: #08a6df !important;
    border-color: #08a6df !important;
}

[data-baseweb="checkbox"] div[role="checkbox"][aria-checked="true"] svg,
[data-testid="stCheckbox"] label:has(input:checked) svg {
    color: #ffffff !important;
    fill: #ffffff !important;
}

[data-testid="stProgress"] {
    padding: 12px 0 8px 0;
}

[data-testid="stProgressBar"] {
    background: #e7f4f8 !important;
    border: 0 !important;
    border-radius: 999px !important;
    overflow: hidden !important;
}

[data-testid="stProgressBar"] > div {
    background: linear-gradient(90deg, var(--dt-indigo), var(--dt-violet), var(--dt-coral)) !important;
    border-radius: 999px !important;
}

[data-testid="stDataFrame"] {
    border: 1px solid var(--dt-line) !important;
    border-radius: 12px !important;
    overflow: hidden !important;
    box-shadow: 0 6px 16px var(--dt-shadow) !important;
}

[data-testid="stAlert"] {
    background: rgba(255,255,255,.85) !important;
    color: var(--dt-ink) !important;
    border: 1px solid var(--dt-line) !important;
    border-radius: 12px !important;
}

.stCaption,
[data-testid="stCaptionContainer"] {
    color: var(--dt-muted) !important;
}

.processing-card {
    display: flex;
    align-items: center;
    gap: 14px;
    margin: 22px 0 10px;
    padding: 18px 20px;
    color: var(--dt-ink);
    background: rgba(255,255,255,.86);
    border: 1px solid var(--dt-line);
    border-radius: 16px;
    box-shadow: 0 12px 28px var(--dt-shadow);
}

.processing-orb {
    width: 26px;
    height: 26px;
    flex: 0 0 auto;
    border: 4px solid #e7f4f8;
    border-top-color: var(--dt-violet);
    border-right-color: var(--dt-coral);
    border-radius: 50%;
    animation: data-tool-spin .85s linear infinite;
}

.processing-card strong { display: block; font-size: 1rem; }
.processing-card span { color: var(--dt-muted); font-size: .9rem; }

@keyframes data-tool-spin { to { transform: rotate(360deg); } }
</style>
""", unsafe_allow_html=True)

brand_logo_uri = local_svg_data_uri("logo-white.svg")
brandmark_html = (
    f"<div class='hs-brandmark'><span>Built by</span><img src='{brand_logo_uri}' alt='BharatSure'></div>"
    if brand_logo_uri else ""
)
heading_col, upload_col = st.columns([7, 3], gap="large")
with heading_col:
    st.markdown(f"""
<div class="hs-title">
    <h1>Data Validation Tool</h1>
    <p>Upload → Validate → Correct → Audit → Download</p>
    {brandmark_html}
</div>
""", unsafe_allow_html=True)
with upload_col:
    st.markdown("<div class='upload-header-label'>Upload Raw Data Excel</div>", unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Upload Raw Data Excel",
        type=["xlsx", "xlsm"],
        label_visibility="collapsed",
    )
    if uploaded:
        st.markdown(
            "<div class='upload-next-step'><strong>Workbook uploaded.</strong> "
            "Scroll down to select the sheet and configure business rules.</div>",
            unsafe_allow_html=True,
        )

st.subheader("Recommended Headers — Copy & Paste")
st.caption(
    "Optional helper: use these semantic headers when your source headers are unclear. "
    "Column order can be changed; the tool reads header meaning, not fixed A/B/C positions."
)
recommended_headers = (
    "Employee ID\tEmployee Name\tEmployee Date of Birth\tEmployee Age\t"
    "Gender\tRelationship to Self\tSpouse Name\tSpouse Date of Birth\t"
    "Child 1 - Name\tChild 1 - Date of Birth\tChild 1 - Age\t"
    "Child 2 - Name\tChild 2 - Date of Birth\tChild 2 - Age\t"
    "Child 3 - Name\tChild 3 - Date of Birth\tChild 3 - Age\t"
    "Father's Name\tFather's Age\tFather's Date of Birth\t"
    "Mother's Name\tMother's Age\tMother's Date of Birth\t"
    "Mobile Number\tEmail"
)
st.code(recommended_headers, language="text")

if uploaded:
    try:
        upload_banner = st.empty()
        upload_banner.markdown(
            "<div class='processing-card'><div class='processing-orb'></div>"
            "<div><strong>Reading your workbook</strong>"
            "<span>Inspecting the uploaded file and preparing its sheets.</span></div></div>",
            unsafe_allow_html=True,
        )
        upload_status = st.status("Reading the uploaded workbook…", expanded=True)
        upload_status.write("Opening the Excel file.")
        upload_progress = st.progress(5, text="Opening your workbook...")
        uploaded_bytes = uploaded.getvalue()
        xls = pd.ExcelFile(io.BytesIO(uploaded_bytes))
        upload_progress.progress(20, text="Workbook opened. Inspecting available sheets...")
        st.subheader("1. Select Sheet(s) to Process")
        selected = []

        if len(xls.sheet_names) == 1:
            # Do not parse the data tab here. The selected sheet is parsed once
            # when Convert is clicked, which avoids a single-sheet-only early
            # read failure and keeps its flow identical to multi-sheet runs.
            upload_status.update(label="Workbook ready", state="running", expanded=True)
            upload_progress.progress(100, text="Workbook ready.")
            upload_status.update(label="Workbook ready", state="complete", expanded=False)
            upload_banner.empty()
            selected = [xls.sheet_names[0]]
            st.success(
                f"Only one sheet detected: **{selected[0]}**. It will be processed automatically."
            )
        else:
            st.info(
                "Multiple sheets detected. Choose exactly which sheet(s) to process. "
                "The tool will not force a sheet based on a hard-coded rule."
            )

            info_rows = []
            total_uploaded_sheets = len(xls.sheet_names)
            for sheet_number, sh in enumerate(xls.sheet_names, start=1):
                upload_status.update(
                    label=f"Inspecting sheet {sheet_number} of {total_uploaded_sheets}",
                    state="running", expanded=True
                )
                upload_progress.progress(
                    int(20 + (sheet_number - 1) / total_uploaded_sheets * 75),
                    text=f"Inspecting sheet {sheet_number} of {total_uploaded_sheets}: {sh}"
                )
                raw = load_source_sheet(uploaded_bytes, sh)
                kind, hr = classify_sheet(raw)
                hint = {
                    "wide_tetris": "Wide family-data format",
                    "wide_bambrew": "Wide family-data format",
                    "long": "Member-level format",
                    "unknown": "Unknown / needs mapping"
                }[kind]
                info_rows.append({
                    "Sheet": sh,
                    "Rows": len(raw),
                    "Columns": len(raw.columns),
                    "Detected Format": hint
                })

            upload_progress.progress(100, text="Workbook ready.")
            upload_status.update(label="Workbook ready", state="complete", expanded=False)
            upload_banner.empty()

            st.dataframe(
                pd.DataFrame(info_rows),
                use_container_width=True,
                hide_index=True
            )

            all_selected = st.checkbox("Select all sheets")
            if all_selected:
                selected = list(xls.sheet_names)
            else:
                selected = st.multiselect(
                    "Choose sheet(s)",
                    options=list(xls.sheet_names)
                )

        if selected:
            st.write("**Selected:** " + ", ".join(selected))
            st.subheader("2. Configure Business Rules")
            skip_validation_rules = st.checkbox(
                "Check for No validation rules",
                help="Skip only the configurable business rules below. Built-in data-quality checks and corrections still run."
            )
            # Row 1: dependent / spouse limits.
            c1,c2,c3=st.columns(3)
            with c1: child_limit=st.text_input("Maximum Children / Employee *",placeholder="Enter limit, e.g. 2")
            with c2: max_dependents=st.number_input("Maximum Dependents / Employee (excl. Self)",min_value=0,value=5,step=1)
            with c3: spouse_limit=st.number_input("Maximum Spouse(s)",min_value=0,value=1,step=1)

            # Row 2: parent limits.
            c4,c5=st.columns(2)
            with c4: parent_limit=st.number_input("Maximum Parents",min_value=0,value=2,step=1)
            with c5: parent_inlaw_limit=st.number_input("Maximum Parents-in-law",min_value=0,value=2,step=1)

            # Row 3: all age-gap rules, followed by the three option cards.
            a1,a2,a3=st.columns(3)
            with a1: child_gap=st.number_input("Self → Child Min Age Gap",min_value=0,value=18,step=1)
            with a2: parent_gap=st.number_input("Self → Parent Min Age Gap",min_value=0,value=18,step=1)
            with a3: parent_inlaw_gap=st.number_input("Spouse → Parent-in-law Min Age Gap",min_value=0,value=18,step=1)

            b1,b2,b3=st.columns(3)
            with b1: spouse_same_gender=st.checkbox("Allow same-gender spouse")
            with b2: parent_same_gender=st.checkbox("Allow same-gender parent pair")
            with b3: allow_parent_and_inlaw_together=st.checkbox("Allow Parent + Parent-in-law together")
            runtime_config={"child_limit":child_limit,"max_dependents":max_dependents,"spouse_limit":spouse_limit,"parent_limit":parent_limit,"parent_inlaw_limit":parent_inlaw_limit,"spouse_same_gender":spouse_same_gender,"parent_same_gender":parent_same_gender,"child_gap":child_gap,"parent_gap":parent_gap,"parent_inlaw_gap":parent_inlaw_gap,"allow_parent_and_inlaw_together":allow_parent_and_inlaw_together,"skip_validation_rules":skip_validation_rules}
            if st.button("Proceed & Generate Corrected Data", type="primary"):
                if not skip_validation_rules and not str(child_limit).strip():
                    st.error("Please enter Maximum Children / Employee before proceeding.")
                    st.stop()
                if not skip_validation_rules:
                    try:
                        runtime_config["child_limit"] = int(float(child_limit))
                        if runtime_config["child_limit"] < 0:
                            raise ValueError
                    except ValueError:
                        st.error("Maximum Children / Employee must be a valid non-negative number.")
                        st.stop()

                results = {}
                processing_banner = st.empty()
                processing_banner.markdown(
                    "<div class='processing-card'><div class='processing-orb'></div>"
                    "<div><strong>Working on your workbook</strong>"
                    "<span>Please keep this tab open while we prepare your download.</span></div></div>",
                    unsafe_allow_html=True,
                )
                process_status = st.status("Preparing your workbook…", expanded=True)
                process_status.write("Reading the selected worksheet(s).")
                progress = st.progress(5, text="Preparing selected sheets...")
                total_sheets = len(selected)

                for sheet_no, sh in enumerate(selected, start=1):
                    progress.progress(
                        int(10 + (sheet_no - 1) / total_sheets * 65),
                        text=f"Processing sheet {sheet_no} of {total_sheets}: {sh}"
                    )
                    process_status.update(
                        label=f"Processing sheet {sheet_no} of {total_sheets}",
                        state="running", expanded=True
                    )
                    try:
                        raw = load_source_sheet(uploaded_bytes, sh)
                        final, logs, validation_errors, count, meta = process_one_sheet(
                            raw, sh, uploaded.name, runtime_config=runtime_config
                        )
                        results[sh] = {
                            "status": "Processed",
                            "final": final,
                            "corrections": logs,
                            "validation_errors": validation_errors,
                            "count": count,
                            "meta": meta,
                            "error": ""
                        }
                    except Exception as e:
                        results[sh] = {
                            "status": "Not Processed",
                            "final": pd.DataFrame(),
                            "corrections": [],
                            "validation_errors": [],
                            "count": pd.DataFrame(),
                            "meta": {},
                            "error": f"{type(e).__name__}: {e}"
                        }

                progress.progress(75, text="Finalizing processed data...")
                processed = {k:v for k,v in results.items() if v["status"] == "Processed"}
                failed = {k:v for k,v in results.items() if v["status"] != "Processed"}
                if processed:
                    try:
                        process_status.update(
                            label="Creating your download…", state="running", expanded=True
                        )
                        process_status.write("Building the formatted Excel workbook.")
                        progress.progress(85, text="Building your downloadable workbook...")
                        with st.spinner("Building the final Excel workbook..."):
                            raw_sheets = {
                                sh: load_source_sheet(uploaded_bytes, sh)
                                for sh in selected
                            }
                            output = build_output_workbook(
                                results, raw_sheets, runtime_config
                            )

                        progress.progress(100, text="Your workbook is ready.")
                        process_status.update(
                            label="Workbook ready", state="complete", expanded=False
                        )
                        processing_banner.empty()
                        st.success("Your updated workbook is ready to download.")
                        st.download_button(
                            "Download the updated data",
                            data=output,
                            file_name="Final Data - Automated.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
                    except Exception as e:
                        process_status.update(
                            label="Workbook generation failed", state="error", expanded=True
                        )
                        st.error(
                            f"Workbook generation failed: {type(e).__name__}: {e}"
                        )
                if failed:
                    if not processed:
                        progress.progress(100, text="Processing finished with errors.")
                        process_status.update(
                            label="Processing finished with errors", state="error", expanded=True
                        )
                    st.warning("Some selected sheets could not be processed. Review the error shown for each sheet above.")
        else:
            if len(xls.sheet_names) > 1:
                st.warning("Select at least one sheet to continue.")

    except Exception as e:
        st.error(f"Could not read the workbook: {e}")
else:
    st.info("Upload a raw Excel workbook to start.")
