import re
from pathlib import Path
from typing import Tuple

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Batch Attendance Mapper", layout="wide")


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
    value = re.sub(r"[^a-z0-9 ]", "", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def parse_event_details(filename: str) -> Tuple[str, str]:
    """
    Example:
    Attendees--An Exclusive AMA with Tetr Co-Founder for Students and Parents _ Tarun Gangwar--18 Apr, 2026 (1).xlsx
    """
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
    df["Batch"] = df["Batch"].fillna("").astype(str).str.strip()

    df["Batch Label"] = df["UG/PG"] + " B" + df["Batch"]
    df["email_key"] = df["Email"].map(normalize_email)
    df["name_key"] = df["Name"].map(normalize_name)

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

    attendance = attendance.rename(columns={"Name": "Name_attendance", "Email": "Email_attendance"})

    students_email = students[students["email_key"] != ""].drop_duplicates(subset=["email_key"])
    students_name = students[students["name_key"] != ""].drop_duplicates(subset=["name_key"])

    email_match = attendance.merge(
        students_email[["email_key", "Name", "Email", "UG/PG", "Batch", "Batch Label"]],
        on="email_key",
        how="left",
    )
    email_match["match_type"] = email_match["Batch Label"].notna().map(lambda x: "Email" if x else "")

    email_matched = email_match[email_match["Batch Label"].notna()].copy()
    email_unmatched = email_match[email_match["Batch Label"].isna()].copy()

    name_match = email_unmatched[["Name_attendance", "Email_attendance", "name_key", "email_key"]].merge(
        students_name[["name_key", "Name", "Email", "UG/PG", "Batch", "Batch Label"]],
        on="name_key",
        how="left",
    )
    name_match["match_type"] = name_match["Batch Label"].notna().map(lambda x: "Name" if x else "")

    combined = pd.concat([email_matched, name_match], ignore_index=True, sort=False)

    combined = combined.rename(
        columns={
            "Name_attendance": "Attendance Name",
            "Email_attendance": "Attendance Email",
            "Name": "Student Name",
            "Email": "Student Email",
        }
    )

    matched = combined[combined["Batch Label"].notna()].copy()
    unmatched = combined[combined["Batch Label"].isna()].copy()

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

    matched = matched[matched_cols].sort_values(
        by=["Batch Label", "Student Name"],
        ascending=[True, True]
    )

    if not unmatched.empty:
        if "Attendance Name" not in unmatched.columns:
            unmatched["Attendance Name"] = ""
        if "Attendance Email" not in unmatched.columns:
            unmatched["Attendance Email"] = ""
        unmatched = unmatched[["Attendance Name", "Attendance Email"]].drop_duplicates().sort_values(
            by=["Attendance Name", "Attendance Email"]
        )
    else:
        unmatched = pd.DataFrame(columns=["Attendance Name", "Attendance Email"])

    return matched, unmatched


st.title("Batch Attendance Mapper")
st.caption("Upload an attendance sheet and automatically distribute attendees into their batches.")

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
        use_container_width=True
    )

uploaded_file = st.file_uploader("Upload attendance sheet", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("Upload the attendance file to see batch-wise attendee distribution.")
    st.stop()

try:
    attendance_df = load_attendance(uploaded_file)
    event_name, event_date = parse_event_details(uploaded_file.name)
    matched_students, unmatched_students = match_attendees(attendance_df, students_df)

    st.subheader("Event Details")
    c1, c2, c3 = st.columns(3)
    c1.metric("Event Name", event_name)
    c2.metric("Event Date", event_date)
    c3.metric("Matched Students", len(matched_students))

    c4, c5, c6 = st.columns(3)
    c4.metric("Attendance Rows", len(attendance_df))
    c5.metric("Batches Present", matched_students["Batch Label"].nunique() if not matched_students.empty else 0)
    c6.metric("Unmatched", len(unmatched_students))

    st.subheader("Batch Attendees Count")

    if matched_students.empty:
        st.warning("No attendees matched with Students.xlsx.")
    else:
        batch_summary = (
            matched_students.groupby("Batch Label", dropna=False)
            .size()
            .reset_index(name="Attendee Count")
            .sort_values("Batch Label")
        )

        st.dataframe(batch_summary, use_container_width=True)

        st.subheader("Batch-wise Name and Email List")

        for _, row in batch_summary.iterrows():
            batch_label = row["Batch Label"]
            count = row["Attendee Count"]

            batch_df = matched_students[matched_students["Batch Label"] == batch_label].copy()
            batch_df = batch_df[["Student Name", "Student Email", "match_type"]].rename(
                columns={"match_type": "Matched By"}
            )

            with st.expander(f"{batch_label} — {count} attendee(s)"):
                st.dataframe(batch_df, use_container_width=True)

    st.subheader("Download Output")

    output_path = Path("batch_attendance_output.xlsx")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        if matched_students.empty:
            pd.DataFrame(columns=["Batch Label", "Attendee Count"]).to_excel(
                writer, sheet_name="Batch Summary", index=False
            )
        else:
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

    if not unmatched_students.empty:
        st.subheader("Unmatched Attendees")
        st.dataframe(unmatched_students, use_container_width=True)

    with st.expander("Preview uploaded attendance file", expanded=False):
        st.dataframe(attendance_df, use_container_width=True)

except Exception as e:
    st.error(f"Could not process the attendance file: {e}")
