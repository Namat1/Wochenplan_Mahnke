import streamlit as st
import pandas as pd
from openpyxl import load_workbook

# Funktion zum Extrahieren der relevanten Daten
def extract_work_data(df):
    relevant_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A."]
    result = []

    # Iteriere durch die Zeilen, beginnend bei Zeile 11 (Index 10) und überspringe jede zweite Zeile
    for index in range(10, len(df), 2):  # Start bei B11 = Index 10
        lastname = df.iloc[index, 1]  # Spalte B = Index 1
        firstname = df.iloc[index, 2]  # Spalte C = Index 2

        # Abbruchbedingung
        if lastname == "Steckel":
            break

        # Iteriere durch die Wochentage
        for day, (col1, col2, date_col) in enumerate(
            [("E", "F", 4), ("G", "H", 6), ("I", "J", 8),
             ("K", "L", 10), ("M", "N", 12), ("O", "P", 14), ("Q", "R", 16)]
        ):
            activity_col1 = df.iloc[index + 1, col1]
            activity_col2 = df.iloc[index + 1, col2]
            activity = f"{activity_col1} {activity_col2}".strip()

            if any(word in str(activity) for word in relevant_words):
                result.append({
                    "Nachname": lastname,
                    "Vorname": firstname,
                    "Wochentag": ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day],
                    "Datum": df.iloc[1, date_col],
                    "Tätigkeit": activity
                })

    return pd.DataFrame(result)

# Funktion, um mehrzeiligen Header zu erstellen
def create_header(num_columns, dates):
    # Erster Header: Wochentage
    weekdays = ["Nachname", "Vorname"] + ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]
    weekdays = weekdays[:num_columns]  # Kürze auf die tatsächliche Spaltenanzahl
    # Zweiter Header: Datum
    sub_header = ["", ""] + list(dates)[:num_columns - 2]  # Kürze auf die tatsächliche Spaltenanzahl

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

    # Überprüfen der tatsächlichen Spaltenanzahl
    num_columns = len(data.columns) - 1  # Ab Spalte B
    header = create_header(num_columns, dates)

    # Daten formatieren
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
