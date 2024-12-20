import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from io import BytesIO

# Funktion zum Extrahieren der relevanten Daten
def extract_work_data(df):
    relevant_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A."]
    result = []

    row_index = 10  # Start bei Zeile 11 (Index 10)
    while row_index <= 144:  # Bis Zeile 145 (Index 144)
        lastname = df.iloc[row_index, 1]  # Spalte B
        firstname = df.iloc[row_index, 2]  # Spalte C
        activities_row = row_index + 1

        if activities_row >= len(df):  # Ende der Daten erreicht
            break

        # Initialisiere Zeilen für die Ausgabe
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

        # Iteriere durch die Wochentage
        for day, (activity_start_col, date_col) in enumerate(
            [(4, 4), (6, 6), (8, 8), (10, 10), (12, 12), (14, 14), (16, 16)]
        ):
            activity = df.iloc[activities_row, activity_start_col]
            if any(word in str(activity) for word in relevant_words):
                weekday = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day]
                row[weekday] = activity

        result.append(row)
        row_index += 2  # Zwei Zeilen weiter

    return pd.DataFrame(result)

# Funktion, um die Datumszeile zu erstellen
def create_header_with_dates(df):
    dates = [
        df.iloc[1, 4],  # E2
        df.iloc[1, 6],  # G2
        df.iloc[1, 8],  # I2
        df.iloc[1, 10], # K2
        df.iloc[1, 12], # M2
        df.iloc[1, 14], # O2
        df.iloc[1, 16], # Q2
    ]
    return dates

# Funktion, um die Spaltenbreite anzupassen
def adjust_column_width(ws):
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2  # Padding für besseren Abstand

# Streamlit App
st.title("Übersicht der Wochenarbeit")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    # Lade die Excel-Datei
    wb = load_workbook(uploaded_file, data_only=True)
    sheet = wb["Druck Fahrer"]
    data = pd.DataFrame(sheet.values)

    # Extrahiere die Daten und das Datum
    extracted_data = extract_work_data(data)
    dates = create_header_with_dates(data)

    # Füge die Datumszeile unter die Wochentage hinzu
    extracted_data.columns = pd.MultiIndex.from_tuples(
        [("Nachname", ""), ("Vorname", "")] +
        [(weekday, date) for weekday, date in zip(["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], dates)]
    )

    # Debugging: Zeige die Daten
    st.write("Inhalt von extracted_data:")
    st.dataframe(extracted_data)

    # Daten als Excel-Datei exportieren
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        extracted_data.to_excel(writer, index=False, sheet_name="Wochenübersicht")
        ws = writer.sheets["Wochenübersicht"]
        adjust_column_width(ws)  # Passe die Spaltenbreite an
    excel_data = output.getvalue()

    # Download-Option
    st.download_button(
        label="Download als Excel",
        data=excel_data,
        file_name="Wochenübersicht.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
