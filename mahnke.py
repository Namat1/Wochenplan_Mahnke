import streamlit as st
import pandas as pd
from openpyxl import load_workbook

# Funktion zum Extrahieren der relevanten Daten
def extract_work_data(df):
    relevant_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A."]
    result = []

    # Iteriere durch die Zeilen, bis der Nachname "Steckel" erreicht wird
    for index, row in df.iterrows():
        lastname = row["Nachname"]
        firstname = row["Vorname"]

        # Abbruchbedingung
        if lastname == "Steckel":
            break

        # Iteriere durch die Wochentage
        for day, (col1, col2, date_col) in enumerate(
            [("E", "F", "E2"), ("G", "H", "G2"), ("I", "J", "I2"),
             ("K", "L", "K2"), ("M", "N", "M2"), ("O", "P", "O2"), ("Q", "R", "Q2")]
        ):
            activity_col1 = row.get(col1, "")
            activity_col2 = row.get(col2, "")
            activity = f"{activity_col1} {activity_col2}".strip()

            if any(word in activity for word in relevant_words):
                result.append({
                    "Nachname": lastname,
                    "Vorname": firstname,
                    "Wochentag": ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day],
                    "Datum": df.loc[1, date_col],
                    "Tätigkeit": activity
                })

    return pd.DataFrame(result)

# Funktion zum Einlesen von Excel-Daten ohne Formeln
def load_excel_without_formulas(file, sheet_name):
    wb = load_workbook(file, data_only=True)
    sheet = wb[sheet_name]

    # Lade die Daten aus dem Blatt in ein DataFrame
    data = pd.DataFrame(sheet.values)

    # Die erste Zeile als Spaltennamen setzen
    data.columns = data.iloc[0]
    data = data[1:]

    return data

# Streamlit App
st.title("Übersicht der Wochenarbeit")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    # Lade die Excel-Datei ohne Formeln
    df = load_excel_without_formulas(uploaded_file, sheet_name="Druck Fahrer")

    # Relevante Spalten umbenennen
    df["Nachname"] = df["B"]
    df["Vorname"] = df["C"]

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
