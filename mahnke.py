import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

def filter_excel_data(file):
    # Load the Excel file
    workbook = load_workbook(file, data_only=True)
    if "Druck Fahrer" not in workbook.sheetnames:
        st.error("Das Blatt 'Druck Fahrer' wurde nicht gefunden.")
        return None

    sheet = workbook["Druck Fahrer"]

    # Find the row where "Aushilfsfahrer" is present
    end_row = None
    for row in sheet.iter_rows(min_row=1, max_col=1, max_row=sheet.max_row):
        for cell in row:
            if cell.value == "Aushilfsfahrer":
                end_row = cell.row - 1
                break
        if end_row:
            break

    if not end_row:
        end_row = sheet.max_row

    # Create a new workbook for the filtered data
    filtered_workbook = openpyxl.Workbook()
    filtered_sheet = filtered_workbook.active
    filtered_sheet.title = "Gefilterte Daten"

    for row in sheet.iter_rows(min_row=1, max_row=end_row):
        for cell in row:
            new_cell = filtered_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
            # Copy cell styles if they exist
            if cell.has_style:
                new_cell.font = cell.font
                new_cell.border = cell.border
                new_cell.fill = cell.fill
                new_cell.number_format = cell.number_format
                new_cell.protection = cell.protection
                new_cell.alignment = cell.alignment

    return filtered_workbook

def to_bytes(workbook):
    output = BytesIO()
    workbook.save(output)
    processed_data = output.getvalue()
    return processed_data

# Streamlit app
st.title("Excel Datenfilter und Download")

uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    with st.spinner("Daten werden verarbeitet..."):
        filtered_workbook = filter_excel_data(uploaded_file)

    if filtered_workbook:
        st.success("Daten wurden erfolgreich gefiltert.")

        excel_data = to_bytes(filtered_workbook)
        st.download_button(
            label="Gefilterte Excel-Datei herunterladen",
            data=excel_data,
            file_name="gefilterte_daten.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
