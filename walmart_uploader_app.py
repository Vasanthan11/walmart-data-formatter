import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils.dataframe import dataframe_to_rows

st.title("WALMART - EXCEL DATA SHEET CORRECTION PROCESS")

raw_text = st.text_area("📋 Paste raw proof content below:", height=400)

# --- Helper Functions ---

def detect_proof(name, path):
    name = name.upper()
    path = path.upper().replace("_", " ").replace("-", " ").replace(".", " ")

    if "PROOF1" in path:
        return "PROOF1"
    if "PRE PRESS" in path:
        return "PRE PRESS"
    if "AFTER PRESS" in path:
        return "AFTER PRESS"
    if "CPR" in path:
        return "CPR"
    if "PRINT READY" in path or "PRESS PRINT READY CHANGES" in path:
        return "PRINT READY"
    if "PRESS" in path:
        if "-AP" in name or "AP-" in name:
            return "AFTER PRESS"
        if "-PP" in name or "PP-" in name:
            return "PRE PRESS"
        return "PRESS"
    return ""

def clean_page_name(name):
    match = re.match(r'^(AP|PP)-(.+)', name)
    if match:
        suffix = match.group(1)
        rest = match.group(2).strip()
        return f"{rest} -{suffix}"
    return name

def parse_date_from_line(line, current_year=2025):
    match_day_time = re.search(r'(\w{3})\s+(\d{1,2}:\d{2})\s*[  ]*([APMapm]+)', line)
    match_full_date = re.search(r'([A-Za-z]{3,})\s+(\d{1,2}),\s+(\d{1,2}:\d{2})\s*([APMapm]+)', line)

    if match_full_date:
        month, day, time_str, am_pm = match_full_date.groups()
        dt = datetime.strptime(f"{month} {day} {current_year} {time_str} {am_pm}", "%b %d %Y %I:%M %p")
        return dt.strftime("%d/%m/%Y")

    elif match_day_time:
        day_abbr, time_str, am_pm = match_day_time.groups()
        weekday_map = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        target_index = weekday_map.index(day_abbr[:3])
        today = datetime.now()
        current_index = today.weekday()
        delta_days = (target_index - current_index) % 7
        base_date = today + timedelta(days=delta_days)
        dt = datetime.strptime(f"{base_date.date()} {time_str} {am_pm}", "%Y-%m-%d %I:%M %p")
        
        # Apply 4 PM rule
        if dt.time() < datetime.strptime("04:00 PM", "%I:%M %p").time():
            dt -= timedelta(days=1)

        return dt.strftime("%d/%m/%Y")

    return ""

def extract_data(raw_text):
    lines = [line.strip() for line in raw_text.split('\n') if line.strip()]

    unwanted_keywords = ["unread", "confirm", "annotation", "reduce"]
    cleaned_lines = [line for line in lines if not any(k in line.lower() for k in unwanted_keywords)]

    result = []
    i = 0
    while i < len(cleaned_lines) - 2:
        timestamp_line = cleaned_lines[i]
        name_line = cleaned_lines[i]
        page_line = cleaned_lines[i + 1]
        path_line = cleaned_lines[i + 2]

        if ',' in name_line and path_line.startswith("/Volumes"):
            assembler = name_line.split(',')[0].strip()
            date_str = parse_date_from_line(timestamp_line)
            page_name_cleaned = clean_page_name(page_line)

            week_match = re.search(r'WK\s*(\d+)', page_name_cleaned, re.IGNORECASE)
            week = f"week-{week_match.group(1)}" if week_match else ""
            proof = detect_proof(page_name_cleaned, path_line)
            qc = "Direct Upload" if page_line.strip().upper().startswith("D-") else "Hariharan"

            result.append({
                "Date": date_str,
                "Banner Name": "walmart",
                "Week": week,
                "Page Name": page_name_cleaned,
                "Proof": proof,
                "Language": "All zones",
                "Page Assembler": assembler,
                "QC": qc
            })

            i += 3
        else:
            i += 1

    return pd.DataFrame(result)

def apply_dropdowns(ws, start_row, end_row):
    dropdowns = {
        'E': ["PRESS", "CPR", "PRE PRESS", "AFTER PRESS", "PRINT READY", "PROOF1"],
        'F': ["All zones", "BIL", "ENG"],
        'G': ["", "Munish Balakrishnan", "Mohammed Siddik", "Sakthivel S", "Prasanth As", "Naveen Kumar"],
        'H': ["", "Direct Upload", "Hariharan"]
    }

    for col, options in dropdowns.items():
        formula = '"' + ",".join(options) + '"'
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        dv.error = "Invalid option"
        dv.errorTitle = "Dropdown Error"
        dv.prompt = "Please select from dropdown"
        dv.promptTitle = "Valid Options"
        ws.add_data_validation(dv)
        dv.add(f"{col}{start_row}:{col}{end_row}")

# --- Streamlit Interface ---

if st.button("✅ Generate Excel with Dropdowns"):
    if not raw_text.strip():
        st.warning("Please paste some raw data first.")
    else:
        df = extract_data(raw_text)
        st.success("🎉 Data processed successfully!")

        wb = Workbook()
        ws = wb.active
        ws.title = "Proof Data"

        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)

        apply_dropdowns(ws, start_row=2, end_row=ws.max_row)

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label="📥 Download Excel",
            data=output,
            file_name="walmart_proof_data_with_date.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
