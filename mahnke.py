import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from io import BytesIO

# Hinweis an den Benutzer
st.info("Die rot und grün gefärbten Zeilen müssen manuell eingetragen werden. Dispo und Aushilfen!")

# Funktion zum Extrahieren der relevanten Daten für einen Bereich
def extract_work_data_for_range(df, start_value, end_value):
    relevant_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule"]
    excluded_words = ["Hoffahrer", "Waschteam", "Aushilfsfahrer"]
    result = []

    # Bereinige Spalte B (zweite Spalte) von Leerzeichen und setze alles in Kleinbuchstaben
    df.iloc[:, 1] = df.iloc[:, 1].astype(str).str.strip().str.lower()

    # Prüfe, ob der Bereich existiert
    if start_value not in df.iloc[:, 1].values or end_value not in df.iloc[:, 1].values:
        st.error(f"Die Werte '{start_value}' oder '{end_value}' wurden in Spalte B nicht gefunden.")
        st.stop()  # Beendet die Ausführung

    start_index = df[df.iloc[:, 1] == start_value].index[0]
    end_index = df[df.iloc[:, 1] == end_value].index[0]

    for row_index in range(start_index, end_index + 1):
        lastname = str(df.iloc[row_index, 1]).strip().title()  # Spalte B
        firstname = str(df.iloc[row_index, 2]).strip().title()  # Spalte C

        # Überspringe Zeilen, bei denen Nachname oder Vorname fehlt oder 'None'
        if not lastname or not firstname or lastname == "None" or firstname == "None":
            continue

        activities_row = row_index + 1

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

        # Iteriere durch die Wochentage und prüfe beide Zellen (z. B. E und F für Sonntag)
        for day, (col1, col2) in enumerate(
            [(4, 5), (6, 7), (8, 9), (10, 11), (12, 13), (14, 15), (16, 17)]
        ):
            # Aktivität aus beiden Zellen auslesen
            activity1 = str(df.iloc[activities_row, col1]).strip()
            activity2 = str(df.iloc[activities_row, col2]).strip()

            # Kombiniere beide Aktivitäten, falls sie nicht leer oder "0" sind
            activity = " ".join(filter(lambda x: x and x != "0", [activity1, activity2])).strip()

            # Prüfen, ob eine der relevanten Aktivitäten vorkommt und keine der ausgeschlossenen Wörter enthalten ist
            if (any(word in activity for word in relevant_words) and
                not any(excluded in activity for excluded in excluded_words)):
                weekday = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day]
                row[weekday] = activity

        result.append(row)

    return pd.DataFrame(result)

# Funktion, um die Tabelle optisch aufzubereiten
def style_excel(ws, calendar_week, num_new_rows, total_rows):
    # ... (wie zuvor, bleibt unverändert)

# Streamlit App
st.title("Wochenarbeitsbericht Fuhrpark")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    progress_bar = st.progress(0)
    progress_status = st.empty()

    # 1. Lade die Excel-Datei
    progress_status.text("Lade die Excel-Datei...")
    wb = load_workbook(uploaded_file, data_only=True)
    progress_bar.progress(20)

    # 2. Lese Daten aus der Excel-Datei
    progress_status.text("Lese Daten aus der Excel-Datei...")
    sheet = wb["Druck Fahrer"]
    data = pd.DataFrame(sheet.values)
    progress_bar.progress(40)

    # 3. Daten extrahieren
    progress_status.text("Extrahiere relevante Daten...")
    extracted_data_1 = extract_work_data_for_range(data, "adler", "steckel")
    progress_bar.progress(60)

    # 4. Daten vorbereiten
    progress_status.text("Bereite die Daten vor...")
    new_data = pd.DataFrame([{
        "Nachname": "Carstensen", "Vorname": "Martin", "Sonntag": "", "Montag": "", "Dienstag": "",
        "Mittwoch": "", "Donnerstag": "", "Freitag": "", "Samstag": ""
    }, {
        "Nachname": "Richter", "Vorname": "Clemens", "Sonntag": "", "Montag": "", "Dienstag": "",
        "Mittwoch": "", "Donnerstag": "", "Freitag": "", "Samstag": ""
    }])
    extracted_data = pd.concat([new_data, extracted_data_1], ignore_index=True)
    progress_bar.progress(80)

    # 5. Erstelle Excel-Datei
    progress_status.text("Erstelle die Excel-Datei...")
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        extracted_data.to_excel(writer, index=False, sheet_name="Wochenübersicht", startrow=2)
        ws = writer.sheets["Wochenübersicht"]
        style_excel(ws, 1, len(new_data), len(extracted_data))
    progress_bar.progress(100)

    # Verarbeitung abgeschlossen
    progress_status.text("Verarbeitung abgeschlossen!")
    st.success("Die Excel-Datei wurde erfolgreich verarbeitet.")

    # Download-Button
    st.download_button(
        label="Download als Excel",
        data=output.getvalue(),
        file_name="Fuhrpark_Wochenbericht.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
