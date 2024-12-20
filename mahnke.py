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
            [("E", "F", "E2"), ("G", "H", "G2"), ("I", "J", "I2"),
             ("K", "L", "K2"), ("M", "N", "M2"), ("O", "P", "O2"), ("Q", "R", "Q2")]
        ):
            # Prüfen, ob Spalten vorhanden sind
            if col1 not in df.columns or col2 not in df.columns:
                st.error(f"Spalte {col1} oder {col2} existiert nicht im DataFrame.")
                break

            activity_col1 = df.iloc[index + 1, ord(col1) - 65]  # Dynamische Konvertierung
            activity_col2 = df.iloc[index + 1, ord(col2) - 65]
            activity = f"{activity_col1} {activity_col2}".strip()

            if any(word in str(activity) for word in relevant_words):
                result.append({
                    "Nachname": lastname,
                    "Vorname": firstname,
                    "Wochentag": ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day],
                    "Datum": df.iloc[1, ord(date_col[0]) - 65],  # Datum aus Zeile 2
                    "Tätigkeit": activity
                })

    return pd.DataFrame(result)

# Funktion zum Einlesen von Excel-Daten ohne Formeln mit dynamischem Header
def load_excel_with_header(file, sheet_name):
    wb = load_workbook(file, data_only=True)
    sheet = wb[sheet_name]

    # Lade die Daten aus dem Blatt
    data = pd.DataFrame(sheet.values)

    # Dynamisch Spaltennamen zuweisen
    headers = ["A", "Nachname", "Vorname"] + [f"Column_{i}" for i in range(3, len(data.columns))]
    data.columns = headers[:len(data.columns)]  # Passe die Header dynamisch an die Anzahl der Spalten an

    return data

# Streamlit App
st.title("Übersicht der Wochenarbeit")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    # Lade die Excel-Datei mit Header
    df = load_excel_with_header(uploaded_file, sheet_name="Druck Fahrer")

    # Zeige die Spaltennamen an
    st.write("Spaltennamen des DataFrames:", df.columns)

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
