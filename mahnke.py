import streamlit as st
import pandas as pd
from openpyxl import load_workbook

# Funktion zum Extrahieren der relevanten Daten
def extract_work_data(df):
    relevant_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A."]
    result = []

    # Iteriere durch die Zeilen, beginnend bei Zeile 11 (Index 10) und überspringe jede zweite Zeile
    for index in range(10, len(df), 2):  # Start bei B11 = Index 10
        lastname = df.iloc[index]["Nachname"]
        firstname = df.iloc[index]["Vorname"]

        # Abbruchbedingung
        if lastname == "Steckel":
            break

        # Iteriere durch die Wochentage
        for day, (col1, col2, date_col) in enumerate(
            [("E", "F", 4), ("G", "H", 6), ("I", "J", 8),
             ("K", "L", 10), ("M", "N", 12), ("O", "P", 14), ("Q", "R", 16)]
        ):
            activity_col1 = df.iloc[index + 1, ord(col1) - 65]  # Dynamische Konvertierung
            activity_col2 = df.iloc[index + 1, ord(col2) - 65]
            activity = f"{activity_col1} {activity_col2}".strip()

            if any(word in str(activity) for word in relevant_words):
                result.append({
                    "Nachname": lastname,
                    "Vorname": firstname,
                    "Wochentag": ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day],
                    "Datum": df.iloc[1, date_col],  # Datum aus der jeweiligen Spalte (4 = E, 6 = G, ...)
                    "Tätigkeit": activity
                })

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
    formatted_data = data.iloc[10:, 1:]  # Daten ab Zeile 11, Spalten ab B
    formatted_data.columns = header

    # Zeige die Tabelle
    st.write("Tabellenübersicht mit mehrzeiligem Header:")
    st.dataframe(formatted_data)

    # Download-Option
    st.download_button(
        label="Download als Excel",
        data=formatted_data.to_excel(index=False, engine="openpyxl"),
        file_name="Wochenübersicht.xlsx"
    )
