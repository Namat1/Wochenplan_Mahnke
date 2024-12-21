import streamlit as st
import pandas as pd
import re
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

    # Bereinige Spalte B (zweite Spalte) von Leerzeichen und setze alles in Kleinbuchstaben
    df.iloc[:, 1] = df.iloc[:, 1].astype(str).str.strip().str.lower()

    # Prüfe, ob der Bereich existiert
    if start_value not in df.iloc[:, 1].values or end_value not in df.iloc[:, 1].values:
        st.error(f"Die Werte '{start_value}' oder '{end_value}' wurden in Spalte B nicht gefunden.")
        st.stop()

    start_index = df[df.iloc[:, 1] == start_value].index[0]
    end_index = df[df.iloc[:, 1] == end_value].index[0]

    for row_index in range(start_index, end_index + 1):
        lastname = str(df.iloc[row_index, 1]).strip().title()
        firstname = str(df.iloc[row_index, 2]).strip().title()

        # Überspringe Zeilen, bei denen Vorname oder Nachname "Leer" ist
        if lastname == "Leer" or firstname == "Leer":
            continue

        # Überspringe Zeilen, bei denen Nachname oder Vorname fehlt oder 'None' ist
        if not lastname or not firstname or lastname == "None" or firstname == "None":
            continue

        activities_row = row_index + 1

        row = {
            "Nachname": lastname,
            "Vorname": firstname,
            "Sonntag": "",
            "Montag": "",
            "Dienstag": "",
            "Mittwoch": "",
            "Donnerstag": "",
            "Freitag": "",
            "Samstag": "",
        }

        for day, (col1, col2) in enumerate(
            [(4, 5), (6, 7), (8, 9), (10, 11), (12, 13), (14, 15), (16, 17)]
        ):
            activity1 = str(df.iloc[activities_row, col1]).strip()
            activity2 = str(df.iloc[activities_row, col2]).strip()
            activity = " ".join(filter(lambda x: x and x != "0", [activity1, activity2])).strip()

            if (any(word in activity for word in relevant_words) and
                not any(excluded in activity for excluded in excluded_words)):
                weekday = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day]
                row[weekday] = activity

        result.append(row)

    return pd.DataFrame(result)

# Funktion zum Extrahieren der KW aus dem Dateinamen
def extract_kw_from_filename(filename):
    match = re.search(r"KW(\d{2})", filename, re.IGNORECASE)
    if match:
        return int(match.group(1))
    return None

# Funktion, um die Tabelle optisch aufzubereiten
def style_excel(ws, calendar_week):
    header_fill = PatternFill(start_color="FFADD8E6", end_color="FFADD8E6", fill_type="solid")
    title_fill = PatternFill(start_color="FF4682B4", end_color="FF4682B4", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    ws["A1"].value = f"Kalenderwoche: {calendar_week}"
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].fill = title_fill
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)

    for col in ws.iter_cols(min_row=3, max_row=3, min_col=1, max_col=ws.max_column):
        for cell in col:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.font = Font(bold=True, size=12)
            cell.fill = header_fill
            cell.border = thin_border

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
    file_name = uploaded_file.name  # Name der hochgeladenen Datei
    calendar_week = extract_kw_from_filename(file_name)
    if calendar_week is None:
        st.error("Die Kalenderwoche konnte nicht aus dem Dateinamen extrahiert werden. Bitte Dateinamen prüfen.")
        st.stop()

    wb = load_workbook(uploaded_file, data_only=True)
    sheet = wb["Druck Fahrer"]
    data = pd.DataFrame(sheet.values)

    new_data = pd.DataFrame([
        {"Nachname": "Castensen", "Vorname": "Martin"},
        {"Nachname": "Richter", "Vorname": "Clemens"}
    ])

    extracted_data = pd.concat([new_data], ignore_index=True)

    columns = ["Nachname", "Vorname"]
    extracted_data.columns = columns

    excel_filename = f"Wochenbericht_Fuhrpark_KW{calendar_week:02d}.xlsx"

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        extracted_data.to_excel(writer, index=False, sheet_name="Wochenübersicht", startrow=2)
        ws = writer.sheets["Wochenübersicht"]
        style_excel(ws, calendar_week)
    excel_data = output.getvalue()

    st.download_button(
        label="Download als Excel",
        data=excel_data,
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
