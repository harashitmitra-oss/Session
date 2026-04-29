import re
from pathlib import Path
from typing import List, Optional, Tuple
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Batch Attendance Mapper", layout="wide")

st.markdown(
    """
    <style>
    .main {
        background: linear-gradient(180deg, #f8fbff 0%, #eef4ff 100%);
    }
    .block-container {
        padding-top: 1.4rem;
        padding-bottom: 2rem;
        max-width: 1200px;
    }
    .hero-card {
        background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 100%);
        padding: 1.4rem 1.5rem;
        border-radius: 22px;
        color: white;
        box-shadow: 0 10px 30px rgba(29, 78, 216, 0.20);
        margin-bottom: 1rem;
    }
    .hero-title {
        font-size: 2rem;
        font-weight: 800;
        margin-bottom: 0.2rem;
    }
    .hero-subtitle {
        font-size: 1rem;
        opacity: 0.92;
    }
    .section-card {
        background: white;
        border-radius: 20px;
        padding: 1rem 1rem 0.75rem 1rem;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
        border: 1px solid rgba(148, 163, 184, 0.18);
        margin-bottom: 1rem;
    }
    .batch-card {
        background: white;
        border-radius: 18px;
        padding: 1rem;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
        border: 1px solid rgba(148, 163, 184, 0.18);
        margin: 0.8rem 0 0.35rem 0;
    }
    .batch-title {
        font-size: 1.1rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.15rem;
    }
    .batch-subtitle {
        color: #475569;
        font-size: 0.95rem;
    }
    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid rgba(148, 163, 184, 0.18);
        padding: 0.85rem;
        border-radius: 18px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# Helpers
# -----------------------------
def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    value = str(value).strip().lower()
    value = re.sub(r"\s+", " ", value)
    return value


def normalize_email(value) -> str:
    return normalize_text(value)


def normalize_name(value) -> str:
    value = normalize_text(value)
    value = re.sub(r"\b(dr|mr|mrs|ms|prof)\.?,?\b", "", value)
    value = re.sub(r"[^a-z0-9 ]", "", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def tokens_from_name(value: str) -> List[str]:
    value = normalize_name(value)
    return [token for token in value.split() if token]


def canonical_name_key(value: str) -> str:
    tokens = tokens_from_name(value)
    if not tokens:
        return ""
    # Remove one-letter tail tokens like B, M, C when present
    if len(tokens) >= 2 and len(tokens[-1]) == 1:
        tokens = tokens[:-1]
    return " ".join(tokens)


def compact_name_key(value: str) -> str:
    return canonical_name_key(value).replace(" ", "")


def clean_batch(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def parse_event_details(filename: str) -> Tuple[str, str]:
    stem = Path(filename).stem
    stem = re.sub(r"\s*\(\d+\)$", "", stem).strip()
    parts = stem.split("--")
    if len(parts) >= 3:
        event_name = parts[1].strip()
        event_date = parts[2].strip()
    elif len(parts) == 2:
        event_name = parts[1].strip()
        event_date = "Unknown"
    else:
        event_name = stem
        event_date = "Unknown"
    return event_name, event_date


def is_placeholder_attendee(name: str, email: str) -> bool:
    combined = f"{normalize_name(name)} {normalize_email(email)}"
    markers = [
        "notetaker",
        "otter.ai",
        "fireflies.ai",
        "tetr events",
        "tetr college of business",
        "zoom",
        "host",
    ]
    return any(marker in combined for marker in markers)


def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def safe_best_name_match(att_row: pd.Series, students_df: pd.DataFrame) -> Optional[pd.Series]:
    att_name = att_row.get("Attendance Name", "")
    att_email = normalize_email(att_row.get("Attendance Email", ""))
    if canonical_name_key(att_name) == "":
        return None
    if is_placeholder_attendee(att_name, att_email):
        return None

    att_canonical = canonical_name_key(att_name)
    att_tokens = set(tokens_from_name(att_name))
    att_compact = compact_name_key(att_name)

    # Stage A: exact canonical match among unique names
    exact = students_df[students_df["canonical_name_key"] == att_canonical].copy()
    if len(exact) == 1:
        return exact.iloc[0]

    # Stage B: compact match catches punctuation/spacing variants
    compact = students_df[students_df["compact_name_key"] == att_compact].copy()
    if len(compact) == 1:
        return compact.iloc[0]

    # Stage C: token-subset / near-match heuristic, only if one safe candidate exists
    candidates = []
    for _, stu in students_df.iterrows():
        stu_name = stu.get("Name", "")
        stu_tokens = set(tokens_from_name(stu_name))
        stu_canonical = stu.get("canonical_name_key", "")
        if not stu_tokens or not att_tokens:
            continue

        inter = att_tokens & stu_tokens
        overlap = len(inter)
        sim = similarity(att_canonical, stu_canonical)

        # Require strong agreement: either almost exact similarity or substantial token overlap
        if sim >= 0.92 or (overlap >= max(2, min(len(att_tokens), len(stu_tokens))) and sim >= 0.80):
            candidates.append((sim, overlap, stu))

    if not candidates:
        return None

    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    best = candidates[0]

    # Avoid ambiguous matches if top two are too close
    if len(candidates) > 1:
        second = candidates[1]
        if abs(best[0] - second[0]) < 0.03 and best[1] == second[1]:
            return None

    return best[2]


# -----------------------------
# Data loaders
# -----------------------------
def load_students(student_file_path: str) -> pd.DataFrame:
    df = pd.read_excel(student_file_path)
    df.columns = [str(c).strip() for c in df.columns]

    required = ["Name", "Email", "UG/PG", "Batch"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Students file is missing required columns: {', '.join(missing)}")

    df = df[required].copy()
    df["Name"] = df["Name"].fillna("").astype(str).str.strip()
    df["Email"] = df["Email"].fillna("").astype(str).str.strip()
    df["UG/PG"] = df["UG/PG"].fillna("").astype(str).str.strip().str.upper()
    df["Batch"] = df["Batch"].map(clean_batch)
    df["Batch Label"] = df["UG/PG"] + " B" + df["Batch"]

    df["email_key"] = df["Email"].map(normalize_email)
    df["name_key"] = df["Name"].map(normalize_name)
    df["canonical_name_key"] = df["Name"].map(canonical_name_key)
    df["compact_name_key"] = df["Name"].map(compact_name_key)

    df = df.drop_duplicates(subset=["email_key", "canonical_name_key", "Batch Label"]).copy()
    return df


def load_personas(persona_file_path: str) -> pd.DataFrame:
    df = pd.read_excel(persona_file_path)
    df.columns = [str(c).strip() for c in df.columns]

    rename_map = {
        "Name": "Student Name",
        "Name ": "Student Name",
        "Persona Name": "Persona Name",
        "Phone": "Phone",
        "Email": "Persona Email 1",
        "Email 2 (if exists)": "Persona Email 2",
        "Email 2 (if exists) ": "Persona Email 2",
        "UG/PG": "UG/PG",
    }
    df = df.rename(columns=rename_map)

    required = ["Student Name", "Persona Name"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Persona file is missing required columns: {', '.join(missing)}")

    for col in ["Student Name", "Persona Name", "Phone", "Persona Email 1", "Persona Email 2", "UG/PG"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["UG/PG"] = df["UG/PG"].str.upper()
    df["persona_name_key"] = df["Persona Name"].map(normalize_name)
    df["persona_canonical_name_key"] = df["Persona Name"].map(canonical_name_key)
    df["persona_compact_name_key"] = df["Persona Name"].map(compact_name_key)
    df["persona_email_1_key"] = df["Persona Email 1"].map(normalize_email)
    df["persona_email_2_key"] = df["Persona Email 2"].map(normalize_email)

    df = df.drop_duplicates(subset=["Student Name", "persona_canonical_name_key", "persona_email_1_key", "persona_email_2_key"]).copy()
    return df


@st.cache_data(show_spinner=False)
def load_attendance(file_bytes: bytes, filename: str) -> pd.DataFrame:
    from io import BytesIO

    df = pd.read_excel(BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]

    if "Name" not in df.columns and "Email" not in df.columns:
        raise ValueError("Attendance file must contain at least Name or Email column.")

    if "Name" not in df.columns:
        df["Name"] = ""
    if "Email" not in df.columns:
        df["Email"] = ""

    df["Name"] = df["Name"].fillna("").astype(str).str.strip()
    df["Email"] = df["Email"].fillna("").astype(str).str.strip()
    df["email_key"] = df["Email"].map(normalize_email)
    df["name_key"] = df["Name"].map(normalize_name)
    df["canonical_name_key"] = df["Name"].map(canonical_name_key)
    df["compact_name_key"] = df["Name"].map(compact_name_key)
    return df


# -----------------------------
# Matching
# -----------------------------
def match_attendees(attendance_df: pd.DataFrame, students_df: pd.DataFrame):
    attendance = attendance_df.copy().rename(columns={"Name": "Attendance Name", "Email": "Attendance Email"})
    students = students_df.copy()

    students_email = students[students["email_key"] != ""].drop_duplicates(subset=["email_key"])

    matched_rows = []
    unmatched_rows = []

    for _, att in attendance.iterrows():
        attendance_name = att.get("Attendance Name", "")
        attendance_email = att.get("Attendance Email", "")
        email_key = att.get("email_key", "")

        if is_placeholder_attendee(attendance_name, attendance_email):
            unmatched_rows.append({
                "Attendance Name": attendance_name,
                "Attendance Email": attendance_email,
                "Reason": "Placeholder / internal attendee",
            })
            continue

        email_hit = students_email[students_email["email_key"] == email_key]
        if len(email_hit) >= 1:
            # Exact email match is definitive for student matching, even if the attendance name differs.
            # If duplicate student rows share the same email, prefer the first de-duplicated record.
            stu = email_hit.iloc[0]
            matched_rows.append({
                "Batch Label": stu["Batch Label"],
                "UG/PG": stu["UG/PG"],
                "Batch": stu["Batch"],
                "Student Name": stu["Name"],
                "Student Email": stu["Email"],
                "Attendance Name": attendance_name,
                "Attendance Email": attendance_email,
                "match_type": "Email",
            })
            continue

        stu = safe_best_name_match(att, students)
        if stu is not None:
            match_type = "Name"
            if normalize_email(attendance_email) != "" and normalize_email(attendance_email) != normalize_email(stu["Email"]):
                match_type = "Name (email differs)"
            matched_rows.append({
                "Batch Label": stu["Batch Label"],
                "UG/PG": stu["UG/PG"],
                "Batch": stu["Batch"],
                "Student Name": stu["Name"],
                "Student Email": stu["Email"],
                "Attendance Name": attendance_name,
                "Attendance Email": attendance_email,
                "match_type": match_type,
            })
            continue

        unmatched_rows.append({
            "Attendance Name": attendance_name,
            "Attendance Email": attendance_email,
            "Reason": "No safe student match found",
        })

    matched = pd.DataFrame(matched_rows)
    unmatched = pd.DataFrame(unmatched_rows)

    if matched.empty:
        matched = pd.DataFrame(columns=[
            "Batch Label", "UG/PG", "Batch", "Student Name", "Student Email",
            "Attendance Name", "Attendance Email", "match_type"
        ])
    else:
        matched["dedupe_key"] = matched["Student Email"].map(normalize_email)
        blank_mask = matched["dedupe_key"] == ""
        matched.loc[blank_mask, "dedupe_key"] = matched.loc[blank_mask, "Student Name"].map(canonical_name_key)
        matched = matched.drop_duplicates(subset=["Batch Label", "dedupe_key"]).copy()
        matched = matched.sort_values(by=["Batch Label", "Student Name"], ascending=[True, True])

    if unmatched.empty:
        unmatched = pd.DataFrame(columns=["Attendance Name", "Attendance Email", "Reason"])
    else:
        unmatched = unmatched.drop_duplicates().sort_values(by=["Attendance Name", "Attendance Email"])

    return matched, unmatched


def safe_best_persona_match(att_row: pd.Series, personas_df: pd.DataFrame) -> Optional[pd.Series]:
    att_name = att_row.get("Attendance Name", "")
    att_email = normalize_email(att_row.get("Attendance Email", ""))
    if canonical_name_key(att_name) == "" and att_email == "":
        return None
    if is_placeholder_attendee(att_name, att_email):
        return None

    att_canonical = canonical_name_key(att_name)
    att_compact = compact_name_key(att_name)
    att_tokens = set(tokens_from_name(att_name))

    def choose_by_persona_name(candidates_df: pd.DataFrame) -> Optional[pd.Series]:
        if candidates_df.empty:
            return None

        exact = candidates_df[candidates_df["persona_canonical_name_key"] == att_canonical].copy()
        if len(exact) == 1:
            return exact.iloc[0]

        compact = candidates_df[candidates_df["persona_compact_name_key"] == att_compact].copy()
        if len(compact) == 1:
            return compact.iloc[0]

        prefix_token_hits = []
        for _, persona in candidates_df.iterrows():
            persona_canonical = persona.get("persona_canonical_name_key", "")
            persona_tokens = set(tokens_from_name(persona.get("Persona Name", "")))
            if not persona_tokens or not att_tokens:
                continue
            if persona_canonical and (
                att_canonical.startswith(persona_canonical + " ")
                or persona_canonical.startswith(att_canonical + " ")
                or persona_tokens.issubset(att_tokens)
                or att_tokens.issubset(persona_tokens)
            ):
                prefix_token_hits.append(persona)

        if len(prefix_token_hits) == 1:
            return prefix_token_hits[0]

        candidates = []
        for _, persona in candidates_df.iterrows():
            persona_name = persona.get("Persona Name", "")
            persona_tokens = set(tokens_from_name(persona_name))
            persona_canonical = persona.get("persona_canonical_name_key", "")
            if not persona_tokens or not att_tokens:
                continue
            overlap = len(att_tokens & persona_tokens)
            sim = similarity(att_canonical, persona_canonical)
            if sim >= 0.85 or (overlap >= 1 and sim >= 0.60):
                candidates.append((sim, overlap, persona))

        if not candidates:
            return None

        candidates.sort(key=lambda x: (x[1], x[0]), reverse=True)
        best = candidates[0]
        if len(candidates) > 1:
            second = candidates[1]
            if abs(best[0] - second[0]) < 0.03 and best[1] == second[1]:
                return None
        return best[2]

    if att_email != "":
        email_matches = personas_df[
            (personas_df["persona_email_1_key"] == att_email) |
            (personas_df["persona_email_2_key"] == att_email)
        ].copy()
        if len(email_matches) == 1:
            return email_matches.iloc[0]
        if len(email_matches) > 1:
            chosen = choose_by_persona_name(email_matches)
            if chosen is not None:
                return chosen

    exact = personas_df[personas_df["persona_canonical_name_key"] == att_canonical].copy()
    if len(exact) == 1:
        return exact.iloc[0]

    compact = personas_df[personas_df["persona_compact_name_key"] == att_compact].copy()
    if len(compact) == 1:
        return compact.iloc[0]

    chosen = choose_by_persona_name(personas_df)
    if chosen is not None:
        return chosen

    return None
    if is_placeholder_attendee(att_name, att_email):
        return None

    att_canonical = canonical_name_key(att_name)
    att_compact = compact_name_key(att_name)
    att_tokens = set(tokens_from_name(att_name))

    def choose_by_persona_name(candidates_df: pd.DataFrame) -> Optional[pd.Series]:
        if candidates_df.empty:
            return None
        exact = candidates_df[candidates_df["persona_canonical_name_key"] == att_canonical].copy()
        if len(exact) == 1:
            return exact.iloc[0]

        compact = candidates_df[candidates_df["persona_compact_name_key"] == att_compact].copy()
        if len(compact) == 1:
            return compact.iloc[0]

        candidates = []
        for _, persona in candidates_df.iterrows():
            persona_name = persona.get("Persona Name", "")
            persona_tokens = set(tokens_from_name(persona_name))
            persona_canonical = persona.get("persona_canonical_name_key", "")
            if not persona_tokens or not att_tokens:
                continue
            overlap = len(att_tokens & persona_tokens)
            sim = similarity(att_canonical, persona_canonical)
            if sim >= 0.90 or (overlap >= max(1, min(len(att_tokens), len(persona_tokens))) and sim >= 0.75):
                candidates.append((sim, overlap, persona))

        if not candidates:
            return None
        candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
        best = candidates[0]
        if len(candidates) > 1:
            second = candidates[1]
            if abs(best[0] - second[0]) < 0.03 and best[1] == second[1]:
                return None
        return best[2]

    if att_email != "":
        email_matches = personas_df[
            (personas_df["persona_email_1_key"] == att_email) |
            (personas_df["persona_email_2_key"] == att_email)
        ].copy()
        if len(email_matches) == 1:
            return email_matches.iloc[0]
        if len(email_matches) > 1:
            chosen = choose_by_persona_name(email_matches)
            if chosen is not None:
                return chosen

    exact = personas_df[personas_df["persona_canonical_name_key"] == att_canonical].copy()
    if len(exact) == 1:
        return exact.iloc[0]

    compact = personas_df[personas_df["persona_compact_name_key"] == att_compact].copy()
    if len(compact) == 1:
        return compact.iloc[0]

    chosen = choose_by_persona_name(personas_df)
    if chosen is not None:
        return chosen

    return None


def match_personas(attendance_df: pd.DataFrame, personas_df: pd.DataFrame):
    attendance = attendance_df.copy().rename(columns={"Name": "Attendance Name", "Email": "Attendance Email"})

    matched_rows = []
    unmatched_rows = []

    for _, att in attendance.iterrows():
        attendance_name = att.get("Attendance Name", "")
        attendance_email = att.get("Attendance Email", "")

        if is_placeholder_attendee(attendance_name, attendance_email):
            continue

        persona = safe_best_persona_match(att, personas_df)
        if persona is not None:
            matched_by = "Persona Name"
            att_email_key = normalize_email(attendance_email)
            if att_email_key != "" and (
                att_email_key == normalize_email(persona.get("Persona Email 1", "")) or
                att_email_key == normalize_email(persona.get("Persona Email 2", ""))
            ):
                matched_by = "Persona Email"
            matched_rows.append({
                "Persona Name": persona.get("Persona Name", ""),
                "Student Name": persona.get("Student Name", ""),
                "Persona Email 1": persona.get("Persona Email 1", ""),
                "Persona Email 2": persona.get("Persona Email 2", ""),
                "Phone": persona.get("Phone", ""),
                "UG/PG": persona.get("UG/PG", ""),
                "Attendance Name": attendance_name,
                "Attendance Email": attendance_email,
                "Matched By": matched_by,
            })
        else:
            unmatched_rows.append({
                "Attendance Name": attendance_name,
                "Attendance Email": attendance_email,
                "Reason": "No safe persona match found",
            })

    matched = pd.DataFrame(matched_rows)
    unmatched = pd.DataFrame(unmatched_rows)

    if matched.empty:
        matched = pd.DataFrame(columns=[
            "Persona Name", "Student Name", "Persona Email 1", "Persona Email 2", "Phone",
            "UG/PG", "Attendance Name", "Attendance Email", "Matched By"
        ])
    else:
        matched["dedupe_key"] = matched["Persona Email 1"].map(normalize_email)
        no_primary = matched["dedupe_key"] == ""
        matched.loc[no_primary, "dedupe_key"] = matched.loc[no_primary, "Persona Email 2"].map(normalize_email)
        still_blank = matched["dedupe_key"] == ""
        matched.loc[still_blank, "dedupe_key"] = matched.loc[still_blank, "Persona Name"].map(canonical_name_key)
        matched = matched.drop_duplicates(subset=["dedupe_key"]).copy()
        matched = matched.sort_values(by=["Persona Name", "Student Name"], ascending=[True, True])

    if unmatched.empty:
        unmatched = pd.DataFrame(columns=["Attendance Name", "Attendance Email", "Reason"])
    else:
        unmatched = unmatched.drop_duplicates().sort_values(by=["Attendance Name", "Attendance Email"])

    return matched, unmatched


def build_final_unmatched(attendance_df: pd.DataFrame, matched_students: pd.DataFrame, matched_personas: pd.DataFrame) -> pd.DataFrame:
    attendance = attendance_df.copy().rename(columns={"Name": "Attendance Name", "Email": "Attendance Email"})
    attendance["row_key"] = attendance.apply(
        lambda r: f"{canonical_name_key(r.get('Attendance Name', ''))}||{normalize_email(r.get('Attendance Email', ''))}",
        axis=1,
    )

    matched_keys = set()

    if not matched_students.empty:
        ms = matched_students.copy()
        ms["row_key"] = ms.apply(
            lambda r: f"{canonical_name_key(r.get('Attendance Name', ''))}||{normalize_email(r.get('Attendance Email', ''))}",
            axis=1,
        )
        matched_keys.update(ms["row_key"].tolist())

    if not matched_personas.empty:
        mp = matched_personas.copy()
        mp["row_key"] = mp.apply(
            lambda r: f"{canonical_name_key(r.get('Attendance Name', ''))}||{normalize_email(r.get('Attendance Email', ''))}",
            axis=1,
        )
        matched_keys.update(mp["row_key"].tolist())

    final_unmatched = attendance[~attendance["row_key"].isin(matched_keys)].copy()
    final_unmatched = final_unmatched[["Attendance Name", "Attendance Email"]].drop_duplicates()
    final_unmatched = final_unmatched.sort_values(by=["Attendance Name", "Attendance Email"])
    return final_unmatched


# -----------------------------
# App UI
# -----------------------------
st.markdown(
    """
    <div class="hero-card">
        <div class="hero-title">Batch Attendance Mapper</div>
        <div class="hero-subtitle">Upload an attendance sheet and instantly see attendees distributed into their batches with improved matching accuracy.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

STUDENTS_FILE = "Students.xlsx"
PERSONA_FILE = "Persona Records.xlsx"

if not Path(STUDENTS_FILE).exists():
    st.error("Students.xlsx was not found. Put Students.xlsx in the same folder as app.py.")
    st.stop()

# Important: no cache here, so new students added to Students.xlsx load immediately on refresh/redeploy
try:
    students_df = load_students(STUDENTS_FILE)
    student_file_mtime = Path(STUDENTS_FILE).stat().st_mtime
except Exception as e:
    st.error(f"Could not load Students.xlsx: {e}")
    st.stop()

persona_df = pd.DataFrame()
persona_file_available = Path(PERSONA_FILE).exists()
persona_file_mtime = None
if persona_file_available:
    try:
        persona_df = load_personas(PERSONA_FILE)
        persona_file_mtime = Path(PERSONA_FILE).stat().st_mtime
    except Exception as e:
        st.warning(f"Could not load Persona Records.xlsx: {e}")
        persona_df = pd.DataFrame()
        persona_file_available = False

with st.expander("Preview Students.xlsx", expanded=False):
    st.caption(f"Loaded rows: {len(students_df)}")
    st.caption(f"Last updated timestamp: {student_file_mtime}")
    st.dataframe(
        students_df[["Name", "Email", "UG/PG", "Batch", "Batch Label"]],
        use_container_width=True,
        height=280,
    )

if persona_file_available:
    with st.expander("Preview Persona Records.xlsx", expanded=False):
        st.caption(f"Loaded persona rows: {len(persona_df)}")
        st.caption(f"Last updated timestamp: {persona_file_mtime}")
        st.dataframe(
            persona_df[["Student Name", "Persona Name", "Persona Email 1", "Persona Email 2", "Phone", "UG/PG"]],
            use_container_width=True,
            height=280,
        )

uploaded_file = st.file_uploader("Upload attendance sheet", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("Upload the attendance file to see batch-wise attendee distribution.")
    st.stop()

try:
    attendance_bytes = uploaded_file.getvalue()
    attendance_df = load_attendance(attendance_bytes, uploaded_file.name)
    event_name, event_date = parse_event_details(uploaded_file.name)
    matched_students, unmatched_students = match_attendees(attendance_df, students_df)
    matched_personas, unmatched_personas = match_personas(attendance_df, persona_df) if persona_file_available and not persona_df.empty else (
        pd.DataFrame(columns=["Persona Name", "Student Name", "Persona Email 1", "Persona Email 2", "Phone", "UG/PG", "Attendance Name", "Attendance Email", "Matched By"]),
        pd.DataFrame(columns=["Attendance Name", "Attendance Email", "Reason"])
    )
    final_unmatched = build_final_unmatched(attendance_df, matched_students, matched_personas)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Event Details")
    st.text_input("Event Name", value=event_name, disabled=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Event Date", event_date)
    c2.metric("Matched Students", len(matched_students))
    c3.metric("Attendance Rows", len(attendance_df))
    c4.metric("Batches Present", matched_students["Batch Label"].nunique() if not matched_students.empty else 0)

    c5, c6 = st.columns(2)
    c5.metric("Final Unmatched", len(final_unmatched))
    if persona_file_available and not persona_df.empty:
        c6.metric("Matched Personas", len(matched_personas))
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Batch Attendees Count")
    if matched_students.empty:
        st.warning("No attendees matched with Students.xlsx.")
        batch_summary = pd.DataFrame(columns=["Batch Label", "Attendee Count"])
    else:
        batch_summary = (
            matched_students.groupby("Batch Label", dropna=False)
            .size()
            .reset_index(name="Attendee Count")
            .sort_values("Batch Label")
        )
        st.dataframe(batch_summary, use_container_width=True, height=260)
    st.markdown('</div>', unsafe_allow_html=True)

    if not matched_students.empty:
        st.subheader("Batch-wise Name and Email List")
        for _, row in batch_summary.iterrows():
            batch_label = row["Batch Label"]
            count = row["Attendee Count"]
            batch_df = matched_students[matched_students["Batch Label"] == batch_label].copy()
            batch_df = batch_df[["Student Name", "Student Email", "Attendance Name", "Attendance Email", "match_type"]].rename(
                columns={"match_type": "Matched By"}
            )

            st.markdown(
                f"""
                <div class="batch-card">
                    <div class="batch-title">{batch_label}</div>
                    <div class="batch-subtitle">{count} attendee(s)</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.dataframe(batch_df, use_container_width=True, height=min(360, 70 + len(batch_df) * 35))

    if persona_file_available and not persona_df.empty:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Persona Attendance")
        if matched_personas.empty:
            st.info("No persona attendees matched from Persona Records.xlsx.")
        else:
            persona_summary = (
                matched_personas.groupby(["UG/PG"], dropna=False)
                .size()
                .reset_index(name="Persona Attendee Count")
                .sort_values("UG/PG")
            )
            st.dataframe(persona_summary, use_container_width=True, height=180)
            st.dataframe(matched_personas, use_container_width=True, height=min(420, 70 + len(matched_personas) * 35))
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Download Output")
    output_path = Path("batch_attendance_output.xlsx")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        batch_summary.to_excel(writer, sheet_name="Batch Summary", index=False)
        matched_students.to_excel(writer, sheet_name="Matched Students", index=False)
        unmatched_students.to_excel(writer, sheet_name="Student Unmatched Debug", index=False)
        final_unmatched.to_excel(writer, sheet_name="Unmatched Students", index=False)
        if persona_file_available and not persona_df.empty:
            matched_personas.to_excel(writer, sheet_name="Matched Personas", index=False)
            unmatched_personas.to_excel(writer, sheet_name="Unmatched Personas", index=False)

    with open(output_path, "rb") as f:
        st.download_button(
            "Download Excel Output",
            data=f,
            file_name=f"batch_attendance_output_{Path(uploaded_file.name).stem}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    st.markdown('</div>', unsafe_allow_html=True)

    if not final_unmatched.empty:
        st.subheader("Unmatched Attendees")
        st.dataframe(final_unmatched, use_container_width=True, height=300)

    with st.expander("Preview uploaded attendance file", expanded=False):
        st.dataframe(attendance_df, use_container_width=True, height=280)

except Exception as e:
    st.error(f"Could not process the attendance file: {e}")
