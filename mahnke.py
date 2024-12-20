import streamlit as st
import pandas as pd
from openpyxl import load_workbook

# Funktion zum Extrahieren der relevanten Daten
def extract_work_data(df):
    relevant_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A."]
    result = []

    # Iteriere durch die Namen, beginnend bei B11 und C11
    row_index = 10  # Start bei Zeile 11 (Index 10)
    while row_index < len(df):
        lastname = df.iloc[row_index, 1]  # Spalte B
        firstname = df.iloc[row_index, 2]  # Spalte C

        # Aktivitäten in der nächsten Zeile
        activities_row = row_index + 1
        if activities_row >= len(df):  # Sicherstellen, dass die Zeile existiert
            break

        # Iteriere durch die Wochentage
        for day, (activity_start_col, date_col) in enumerate(
            [(4, 4), (6, 6), (8, 8), (10, 10), (12, 12), (14, 14), (16, 16)]
        ):
            activity = df.iloc[activities_row, activity_start_col]
            if any(word in str(activity) for word in relevant_words):
                result.append({
                    "Nachname": lastname,
                    "Vorname": firstname,
                    "Wochentag": ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day],
                    "Datum": df.iloc[1, date_col],  # Datum aus Zeile 2
                    "Tätigkeit": activity
                })

        # Springe zum nächsten Namen (zwei Zeilen weiter)
        row_index += 2

    return pd.DataFrame(result)

# Funktion, um mehrzeiligen Header zu erstellen
def create_header(dates):
    # Erster Header: Wochentage
    weekdays = ["Nachname", "Vorname"] + ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]
    # Zweiter Header: Datum
    sub_header = ["", ""] + list(dates)

    return pd.MultiIndex.from_arrays([weekdays, sub_header])

# Streamlit App
st.title("Übersicht der Wochenarbeit")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    # Lade die Excel-Datei
    wb = load_workbook(uploaded_file, data_only=True)
    sheet = wb["Druck Fahrer"]
    data = pd.DataFrame(sheet.values)

    # Holen der Datumswerte aus den definierten Spalten
    dates = [
        data.iloc[1, 4],  # E2
        data.iloc[1, 6],  # G2
        data.iloc[1, 8],  # I2
        data.iloc[1, 10], # K2
        data.iloc[1, 12], # M2
        data.iloc[1, 14], # O2
        data.iloc[1, 16], # Q2
    ]

    # Erstellen eines DataFrames mit mehrzeiligem Header
    header = create_header(dates)
    extracted_data = extract_work_data(data)

    # Zeige die Tabelle
    st.write("Tabellenübersicht:")
    st.dataframe(extracted_data)

    # Download-Option
    st.download_button(
        label="Download als Excel",
        data=extracted_data.to_excel(index=False, engine="openpyxl"),
        file_name="Wochenübersicht.xlsx"
    )
