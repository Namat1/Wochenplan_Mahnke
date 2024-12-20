import streamlit as st
import pandas as pd
from openpyxl import load_workbook
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

        for day, (activity_start_col, date_col) in enumerate(
            [(4, 4), (6, 6), (8, 8), (10, 10), (12, 12), (14, 14), (16, 16)]
        ):
            activity = df.iloc[activities_row, activity_start_col]
            if any(word in str(activity) for word in relevant_words):
                result.append({
                    "Nachname": lastname,
                    "Vorname": firstname,
                    "Wochentag": ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day],
                    "Datum": df.iloc[1, date_col],
                    "Tätigkeit": activity
                })

        row_index += 2  # Zwei Zeilen weiter

    return pd.DataFrame(result)

# Streamlit App
st.title("Übersicht der Wochenarbeit")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    # Lade die Excel-Datei
    wb = load_workbook(uploaded_file, data_only=True)
    sheet = wb["Druck Fahrer"]
    data = pd.DataFrame(sheet.values)

    # Extrahiere die Daten bis B145
    extracted_data = extract_work_data(data)

    # Debugging: Zeige die Daten
    st.write("Typ von extracted_data:", type(extracted_data))
    st.write("Inhalt von extracted_data:")
    st.dataframe(extracted_data)

    # Daten als Excel-Datei exportieren
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        extracted_data.to_excel(writer, index=False, sheet_name="Wochenübersicht")
    excel_data = output.getvalue()

    # Download-Option
    st.download_button(
        label="Download als Excel",
        data=excel_data,
        file_name="Wochenübersicht.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
