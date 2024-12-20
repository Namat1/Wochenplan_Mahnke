import streamlit as st
import pandas as pd
from openpyxl import load_workbook

# Funktion zum Extrahieren der relevanten Daten
def extract_work_data(df):
    relevant_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A."]
    result = []

    # Iteriere durch die Zeilen, beginnend bei B11, und überspringe jede zweite Zeile
    for index in range(10, len(df), 2):  # 10 entspricht B11, da 0-basierter Index
        lastname = df.iloc[index, 1]  # Spalte B = Index 1
        firstname = df.iloc[index, 2]  # Spalte C = Index 2

        # Abbruchbedingung
        if lastname == "Steckel":
            break

        # Iteriere durch die Wochentage
        for day, (col1, col2, date_col) in enumerate(
            [("E", "F", "E2"), ("G", "H", "G2"), ("I", "J", "I2"),
             ("K", "L", "K2"), ("M", "N", "M2"), ("O", "P", "O2"), ("Q", "R", "Q2")]
        ):
            activity_col1 = df.iloc[index + 1, ord(col1) - 65]  # Zeile darunter, Spalte berechnet
            activity_col2 = df.iloc[index + 1, ord(col2) - 65]  # Zeile darunter, Spalte berechnet
            activity = f"{activity_col1} {activity_col2}".strip()

            if any(word in str(activity) for word in relevant_words):
                result.append({
                    "Nachname": lastname,
                    "Vorname": firstname,
                    "Wochentag": ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day],
                    "Datum": df.iloc[1, ord(date_col[0]) - 65],  # Nur das erste Zeichen von date_col verwenden
                    "Tätigkeit": activity
                })

    return pd.DataFrame(result)

# Funktion zum Einlesen von Excel-Daten ohne Formeln
def load_excel_without_formulas(file, sheet_name):
    wb = load_workbook(file, data_only=True)
    sheet = wb[sheet_name]

    # Lade die Daten aus dem Blatt in ein DataFrame
    data = pd.DataFrame(sheet.values)

    return data

# Streamlit App
st.title("Übersicht der Wochenarbeit")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    # Lade die Excel-Datei ohne Formeln
    df = load_excel_without_formulas(uploaded_file, sheet_name="Druck Fahrer")

    # Extrahiere die Daten
    data = extract_work_data(df)

    # Zeige die Tabelle
    st.write("Tabellenübersicht der Wochenarbeit:")
    st.dataframe(data)

    # Download-Option
    st.download_button(
        label="Download als Excel",
        data=data.to_excel(index=False, engine='openpyxl'),
        file_name="Wochenübersicht.xlsx"
    )
