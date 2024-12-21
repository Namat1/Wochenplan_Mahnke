import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from io import BytesIO

# Hinweis an den Benutzer
st.info("Die rot und grün gefärbten Zeilen müssen manuell eingetragen werden. Dispo und Aushilfen!")

# Funktion zum Extrahieren der Kalenderwoche aus dem Dateinamen
def extract_calendar_week(filename):
    import re
    match = re.search(r'KW(\d{2})', filename)
    if match:
        return int(match.group(1))
    else:
        st.error("Keine Kalenderwoche im Dateinamen gefunden!")
        st.stop()

# Funktion zum Überspringen von Zeilen mit "Leer"
def filter_rows(df):
    return df[~df[['Nachname', 'Vorname']].apply(lambda x: x.str.contains("Leer", na=False, case=False)).any(axis=1)]

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

# Streamlit App
st.title("Wochenarbeitsbericht Fuhrpark")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    # Extrahiere Kalenderwoche aus dem Dateinamen
    calendar_week = extract_calendar_week(uploaded_file.name)

    # Lade die Excel-Datei
    wb = load_workbook(uploaded_file, data_only=True)
    sheet = wb["Druck Fahrer"]
    data = pd.DataFrame(sheet.values)

    # Filtere Zeilen mit "Leer"
    data.columns = ["Spalte1", "Nachname", "Vorname"] + [f"Spalte{i}" for i in range(4, len(data.columns) + 1)]
    data = filter_rows(data)

    # Extrahiere die Datenbereiche
    extracted_data = pd.concat([
        extract_work_data_for_range(data, "adler", "zosel"),
        extract_work_data_for_range(data, "böhnke", "kleiber"),
        extract_work_data_for_range(data, "linke", "steckel")
    ], ignore_index=True)

    # Erstelle den Dateinamen mit Kalenderwoche
    excel_filename = f"Wochenbericht_Fuhrpark_KW{calendar_week:02d}.xlsx"

    # Daten als Excel-Datei exportieren
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        extracted_data.to_excel(writer, index=False, sheet_name="Wochenübersicht", startrow=2)
        ws = writer.sheets["Wochenübersicht"]
        ws["A1"] = f"Kalenderwoche: {calendar_week}"
    excel_data = output.getvalue()

    # Download-Option
    st.download_button(
        label="Download als Excel",
        data=excel_data,
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
