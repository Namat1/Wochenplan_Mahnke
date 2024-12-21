import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from io import BytesIO

# Hinweis an den Benutzer
st.info("Die rot und grün gefärbten Zeilen müssen manuell eingetragen werden. Dispo und Aushilfen!")

# Funktion zum Extrahieren der relevanten Daten für einen Bereich
def extract_work_data_for_range(df, start_value, end_value):
    relevant_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A."]
    excluded_words = ["Hoffahrer", "Waschteam", "Aushilfsfahrer"]
    result = []

    df.iloc[:, 1] = df.iloc[:, 1].astype(str).str.strip().str.lower()

    if start_value not in df.iloc[:, 1].values or end_value not in df.iloc[:, 1].values:
        st.error(f"Die Werte '{start_value}' oder '{end_value}' wurden in Spalte B nicht gefunden.")
        st.stop()

    start_index = df[df.iloc[:, 1] == start_value].index[0]
    end_index = df[df.iloc[:, 1] == end_value].index[0]

    for row_index in range(start_index, end_index + 1):
        lastname = str(df.iloc[row_index, 1]).strip().title()
        firstname = str(df.iloc[row_index, 2]).strip().title()

        if not lastname or not firstname or lastname == "None" or firstname == "None":
            continue

        activities_row = row_index + 1
        row = {day: "" for day in ["Nachname", "Vorname", "Sonntag", "Montag", "Dienstag", 
                                   "Mittwoch", "Donnerstag", "Freitag", "Samstag"]}
        row["Nachname"], row["Vorname"] = lastname, firstname

        for day, (col1, col2) in enumerate(
            [(4, 5), (6, 7), (8, 9), (10, 11), (12, 13), (14, 15), (16, 17)]
        ):
            activity1 = str(df.iloc[activities_row, col1]).strip()
            activity2 = str(df.iloc[activities_row, col2]).strip()

            activity = " ".join(filter(lambda x: x and x != "0", [activity1, activity2])).strip()

            if (any(word in activity for word in relevant_words) and
                not any(excluded in activity for excluded in excluded_words)):
                weekday = ["Sonntag", "Montag", "Dienstag", "Mittwoch", 
                           "Donnerstag", "Freitag", "Samstag"][day]
                row[weekday] = activity

        result.append(row)

    return pd.DataFrame(result)

# Funktion zum Erstellen der Header-Daten mit Datumsangaben
def create_header_with_dates(df):
    dates = [
        pd.to_datetime(df.iloc[1, 4]).strftime('%d.%m.%Y'),
        pd.to_datetime(df.iloc[1, 6]).strftime('%d.%m.%Y'),
        pd.to_datetime(df.iloc[1, 8]).strftime('%d.%m.%Y'),
        pd.to_datetime(df.iloc[1, 10]).strftime('%d.%m.%Y'),
        pd.to_datetime(df.iloc[1, 12]).strftime('%d.%m.%Y'),
        pd.to_datetime(df.iloc[1, 14]).strftime('%d.%m.%Y'),
        pd.to_datetime(df.iloc[1, 16]).strftime('%d.%m.%Y'),
    ]
    return dates

# Funktion zum Stylen der Excel-Datei
def style_excel(ws, calendar_week, num_new_rows, total_rows):
    header_fill = PatternFill(start_color="FFADD8E6", end_color="FFADD8E6", fill_type="solid")
    alt_row_fill = PatternFill(start_color="FFFFF0AA", end_color="FFFFF0AA", fill_type="solid")
    title_fill = PatternFill(start_color="FF4682B4", end_color="FF4682B4", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    ws["A1"].value = f"Kalenderwoche: {calendar_week + 1}"
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].fill = title_fill
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)

    for col in ws.iter_cols(min_row=3, max_row=3):
        for cell in col:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.border = thin_border

    for row in range(4, ws.max_row + 1):
        for cell in ws[row]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            if row % 2 == 0:
                cell.fill = alt_row_fill

    adjust_column_width(ws)

# Funktion zum Anpassen der Spaltenbreite
def adjust_column_width(ws):
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

# Streamlit App
st.title("Wochenarbeitsbericht Fuhrpark")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=True)
    try:
        sheet = wb["Druck Fahrer"]
    except KeyError:
        st.error("Das Arbeitsblatt 'Druck Fahrer' wurde in der hochgeladenen Datei nicht gefunden.")
        st.stop()

    date_g2 = sheet["G2"].value
    if isinstance(date_g2, datetime):
        calendar_week = date_g2.isocalendar()[1]
        excel_filename = f"Wochenbericht_Fuhrpark_KW{calendar_week:02d}.xlsx"

        employee_data = [
            {"Nachname": "Castensen", "Vorname": "Martin"},
            {"Nachname": "Richter", "Vorname": "Clemens"},
            {"Nachname": "Gebauer", "Vorname": "Ronny"},
            {"Nachname": "Pham Manh", "Vorname": "Chris"},
            {"Nachname": "Ohlenroth", "Vorname": "Nadja"}
        ]
        columns = ["Nachname", "Vorname", "Sonntag", "Montag", "Dienstag", 
                   "Mittwoch", "Donnerstag", "Freitag", "Samstag"]
        new_data = pd.DataFrame([{**emp, **{col: "" for col in columns[2:]}} for emp in employee_data])

        dates = create_header_with_dates(sheet)
        extracted_data = pd.concat([new_data], ignore_index=True)
        extracted_data.columns = ["Nachname", "Vorname"] + [
            f"{day} ({date})" for day, date in zip(
                ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], dates
            )
        ]

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            extracted_data.to_excel(writer, index=False, sheet_name="Wochenübersicht", startrow=2)
            ws = writer.sheets["Wochenübersicht"]
            style_excel(ws, calendar_week, len(new_data), len(extracted_data))
        excel_data = output.getvalue()

        st.download_button(
            label="Download als Excel",
            data=excel_data,
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Fehler: In Zelle G2 steht kein gültiges Datum.")
