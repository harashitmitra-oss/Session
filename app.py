import re
from pathlib import Path
from typing import Tuple

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


def looks_like_internal_or_proxy_email(email: str) -> bool:
    email = normalize_email(email)
    proxy_markers = [
        "@tetr.",
        "otter.ai",
        "fireflies.ai",
        "events",
        "marketing",
        "notetaker",
    ]
    return any(marker in email for marker in proxy_markers)


@st.cache_data
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

    # Only allow unique name matches to avoid pushing same-name students into wrong batches
    name_counts = df["name_key"].value_counts(dropna=False)
    df["name_is_unique"] = df["name_key"].map(name_counts) == 1

    return df


@st.cache_data
def load_attendance(uploaded_file) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file)
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

    return df


def match_attendees(attendance_df: pd.DataFrame, students_df: pd.DataFrame):
    attendance = attendance_df.copy()
    students = students_df.copy()

    attendance = attendance.rename(columns={"Name": "Attendance Name", "Email": "Attendance Email"})

    students_email = students[students["email_key"] != ""].drop_duplicates(subset=["email_key"])
    students_name = students[(students["name_key"] != "") & (students["name_is_unique"])].drop_duplicates(subset=["name_key"])

    # 1) Exact email match only
    email_match = attendance.merge(
        students_email[["email_key", "Name", "Email", "UG/PG", "Batch", "Batch Label"]],
        on="email_key",
        how="left",
    )
    email_match["match_type"] = email_match["Batch Label"].notna().map(lambda x: "Email" if x else "")

    email_matched = email_match[email_match["Batch Label"].notna()].copy()
    email_unmatched = email_match[email_match["Batch Label"].isna()].copy()

    # 2) Safer name fallback:
    # only when the student name is unique in Students.xlsx, and the attendance email is blank
    # or clearly an internal/proxy/notetaker email rather than the student's personal email.
    name_candidates = email_unmatched.copy()
    name_candidates["allow_name_match"] = name_candidates["Attendance Email"].map(
        lambda x: (normalize_email(x) == "") or looks_like_internal_or_proxy_email(x)
    )
    name_candidates = name_candidates[name_candidates["allow_name_match"]].copy()

    name_match = name_candidates[["Attendance Name", "Attendance Email", "name_key", "email_key"]].merge(
        students_name[["name_key", "Name", "Email", "UG/PG", "Batch", "Batch Label"]],
        on="name_key",
        how="left",
    )
    name_match["match_type"] = name_match["Batch Label"].notna().map(lambda x: "Name" if x else "")

    # Unmatched rows include rows blocked from unsafe name matching plus failed name matches
    blocked_unmatched = email_unmatched[~email_unmatched.index.isin(name_candidates.index)].copy()
    blocked_unmatched = blocked_unmatched[["Attendance Name", "Attendance Email"]].copy()

    combined = pd.concat([email_matched, name_match], ignore_index=True, sort=False)
    combined = combined.rename(columns={"Name": "Student Name", "Email": "Student Email"})

    matched = combined[combined["Batch Label"].notna()].copy()
    failed_name = combined[combined["Batch Label"].isna()].copy()

    if not matched.empty:
        matched["dedupe_key"] = matched["Student Email"].map(normalize_email)
        blank_mask = matched["dedupe_key"] == ""
        matched.loc[blank_mask, "dedupe_key"] = matched.loc[blank_mask, "Student Name"].map(normalize_name)
        matched = matched.drop_duplicates(subset=["Batch Label", "dedupe_key"]).copy()

    matched_cols = [
        "Batch Label",
        "UG/PG",
        "Batch",
        "Student Name",
        "Student Email",
        "Attendance Name",
        "Attendance Email",
        "match_type",
    ]
    for col in matched_cols:
        if col not in matched.columns:
            matched[col] = ""

    matched = matched[matched_cols].sort_values(by=["Batch Label", "Student Name"], ascending=[True, True])

    unmatched = pd.concat(
        [
            failed_name[["Attendance Name", "Attendance Email"]] if not failed_name.empty else pd.DataFrame(columns=["Attendance Name", "Attendance Email"]),
            blocked_unmatched,
        ],
        ignore_index=True,
    )
    unmatched = unmatched.drop_duplicates().sort_values(by=["Attendance Name", "Attendance Email"])

    return matched, unmatched


# -----------------------------
# App UI
# -----------------------------
st.markdown(
    """
    <div class="hero-card">
        <div class="hero-title">Batch Attendance Mapper</div>
        <div class="hero-subtitle">Upload an attendance sheet and instantly see attendees distributed into their batches.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

STUDENTS_FILE = "Students.xlsx"

if not Path(STUDENTS_FILE).exists():
    st.error("Students.xlsx was not found. Put Students.xlsx in the same folder as app.py.")
    st.stop()

try:
    students_df = load_students(STUDENTS_FILE)
except Exception as e:
    st.error(f"Could not load Students.xlsx: {e}")
    st.stop()

with st.expander("Preview Students.xlsx", expanded=False):
    st.dataframe(
        students_df[["Name", "Email", "UG/PG", "Batch", "Batch Label"]],
        use_container_width=True,
        height=280,
    )

uploaded_file = st.file_uploader("Upload attendance sheet", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("Upload the attendance file to see batch-wise attendee distribution.")
    st.stop()

try:
    attendance_df = load_attendance(uploaded_file)
    event_name, event_date = parse_event_details(uploaded_file.name)
    matched_students, unmatched_students = match_attendees(attendance_df, students_df)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Event Details")
    c1, c2, c3 = st.columns(3)
    c1.metric("Event Name", event_name)
    c2.metric("Event Date", event_date)
    c3.metric("Matched Students", len(matched_students))

    c4, c5, c6 = st.columns(3)
    c4.metric("Attendance Rows", len(attendance_df))
    c5.metric("Batches Present", matched_students["Batch Label"].nunique() if not matched_students.empty else 0)
    c6.metric("Unmatched", len(unmatched_students))
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
            batch_df = batch_df[["Student Name", "Student Email", "match_type"]].rename(
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
            st.dataframe(batch_df, use_container_width=True, height=min(320, 70 + len(batch_df) * 35))

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Download Output")
    output_path = Path("batch_attendance_output.xlsx")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        batch_summary.to_excel(writer, sheet_name="Batch Summary", index=False)
        matched_students.to_excel(writer, sheet_name="Matched Students", index=False)
        unmatched_students.to_excel(writer, sheet_name="Unmatched Students", index=False)

    with open(output_path, "rb") as f:
        st.download_button(
            "Download Excel Output",
            data=f,
            file_name=f"batch_attendance_output_{Path(uploaded_file.name).stem}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    st.markdown('</div>', unsafe_allow_html=True)

    if not unmatched_students.empty:
        st.subheader("Unmatched Attendees")
        st.dataframe(unmatched_students, use_container_width=True, height=280)

    with st.expander("Preview uploaded attendance file", expanded=False):
        st.dataframe(attendance_df, use_container_width=True, height=280)

except Exception as e:
    st.error(f"Could not process the attendance file: {e}")
