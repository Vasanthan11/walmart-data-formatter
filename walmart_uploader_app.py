import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils.dataframe import dataframe_to_rows

st.title("WALMART - EXCEL DATA SHEET CORRECTION PROCESS")

raw_text = st.text_area("📋 Paste raw proof content below:", height=300)

# Page Assembler mapping
assembler_map = {
    "MU": "Munish",
    "SD": "Siddik",
    "SK": "Sakthivel",
    "PR": "Prasanth"
}

# QC mapping
qc_map = {
    "D": "Direct Upload",
    "ND": "Hariharan"
}

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
    # Move -AP or -PP to end if present in prefix
    match = re.match(r'^(AP|PP)-(.+)', name)
    if match:
        suffix = match.group(1)
        rest = match.group(2).strip()
        return f"{rest} -{suffix}"
    return name

def extract_data(raw_text):
    lines = [line.strip() for line in raw_text.split('\n') if line.strip()]
    pairs = [(lines[i], lines[i+1]) for i in range(0, len(lines), 2)]

    result = []

    for name, path in pairs:
        name = clean_page_name(name)
        week_match = re.search(r'WK\s*(\d+)', name, re.IGNORECASE)
        week = f"week-{week_match.group(1)}" if week_match else ""
        proof = detect_proof(name, path)

        # Extract prefix like MU-D or SK-ND
        prefix_match = re.match(r'([A-Z]{2})-([A-Z]{1,2})', name)
        assembler_code = prefix_match.group(1) if prefix_match else ""
        qc_code = prefix_match.group(2) if prefix_match else ""

        assembler = assembler_map.get(assembler_code, "")
        qc = qc_map.get(qc_code, "")

        result.append({
            "Banner Name": "walmart",
            "Week": week,
            "Page Name": name,
            "Proof": proof,
            "Language": "All zones",
            "Page Assembler": assembler,
            "QC": qc
        })

    return pd.DataFrame(result)

def apply_dropdowns(ws, start_row, end_row):
    dropdowns = {
        'D': ["PRESS", "CPR", "PRE PRESS", "AFTER PRESS", "PRINT READY", "PROOF1"],
        'E': ["All zones", "BIL", "ENG"],
        'F': ["", "Munish", "Siddik", "Sakthivel", "Prasanth"],
        'G': ["", "Direct Upload", "Hariharan"]
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

if st.button("✅ Generate Excel with Dropdowns"):
    if not raw_text.strip():
        st.warning("Please paste some raw data first.")
    else:
        df = extract_data(raw_text)
        st.success("🎉 Data processed successfully!")

        # Create Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Proof Data"

        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)

        apply_dropdowns(ws, start_row=2, end_row=ws.max_row)

        # Output Excel
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label="📥 Download Excel",
            data=output,
            file_name="walmart_proof_data_with_qc.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
