import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Hinweis an den Benutzer
st.info("Die rot und grün gefärbten Zeilen müssen manuell eingetragen werden. Diese Zeilen dienen nur zur Markierung!")

# Funktion zum Extrahieren der relevanten Daten für einen Bereich
def extract_work_data_for_range(df, start_value, end_value):
    relevant_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A."]
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

# Funktion, um die Datumszeile zu erstellen
def create_header_with_dates(df):
    dates = [
        pd.to_datetime(df.iloc[1, 4]).strftime('%d.%m.%Y'),  # E2
        pd.to_datetime(df.iloc[1, 6]).strftime('%d.%m.%Y'),  # G2
        pd.to_datetime(df.iloc[1, 8]).strftime('%d.%m.%Y'),  # I2
        pd.to_datetime(df.iloc[1, 10]).strftime('%d.%m.%Y'), # K2
        pd.to_datetime(df.iloc[1, 12]).strftime('%d.%m.%Y'), # M2
        pd.to_datetime(df.iloc[1, 14]).strftime('%d.%m.%Y'), # O2
        pd.to_datetime(df.iloc[1, 16]).strftime('%d.%m.%Y'), # Q2
    ]
    return dates

# Funktion zum Erstellen einer PDF-Datei mit ReportLab
def generate_pdf(df, calendar_week, output_filename):
    c = canvas.Canvas(output_filename, pagesize=letter)
    width, height = letter

    # Header hinzufügen
    c.setFont("Helvetica", 14)
    c.drawString(100, height - 40, f"Wochenbericht Fuhrpark Kalenderwoche {calendar_week}")

    # Tabelle erstellen
    c.setFont("Helvetica", 10)
    x = 100
    y = height - 60
    for column in df.columns:
        c.drawString(x, y, column)
        x += 100

    y -= 20
    for index, row in df.iterrows():
        x = 100
        for value in row:
            c.drawString(x, y, str(value))
            x += 100
        y -= 20

    c.save()

# Streamlit App
st.title("Übersicht der Wochenarbeit")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    # Lade die Excel-Datei
    wb = load_workbook(uploaded_file, data_only=True)
    sheet = wb["Druck Fahrer"]
    data = pd.DataFrame(sheet.values)

    # Erstelle 6 Zeilen für die Mitarbeiter oberhalb von "Adler"
    new_data = pd.DataFrame([{
        "Nachname": "Castensen", "Vorname": "Martin", "Sonntag": "", "Montag": "", "Dienstag": "", 
        "Mittwoch": "", "Donnerstag": "", "Freitag": "", "Samstag": ""
    }, {
        "Nachname": "Richter", "Vorname": "Clemens", "Sonntag": "", "Montag": "", "Dienstag": "", 
        "Mittwoch": "", "Donnerstag": "", "Freitag": "", "Samstag": ""
    }, {
        "Nachname": "Gebauer", "Vorname": "Ronny", "Sonntag": "", "Montag": "", "Dienstag": "", 
        "Mittwoch": "", "Donnerstag": "", "Freitag": "", "Samstag": ""
    }, {
        "Nachname": "Pham Manh", "Vorname": "Chris", "Sonntag": "", "Montag": "", "Dienstag": "", 
        "Mittwoch": "", "Donnerstag": "", "Freitag": "", "Samstag": ""
    }, {
        "Nachname": "Ohlenroth", "Vorname": "Nadja", "Sonntag": "", "Montag": "", "Dienstag": "", 
        "Mittwoch": "", "Donnerstag": "", "Freitag": "", "Samstag": ""
    }])

    # Extrahiere die Daten für den Bereich (Adler bis Zosel)
    extracted_data_1 = extract_work_data_for_range(data, "adler", "zosel")

    # Extrahiere die Daten für den Bereich (Böhnke bis Kleiber)
    extracted_data_2 = extract_work_data_for_range(data, "böhnke", "kleiber")

    # Extrahiere die Daten für den Bereich (Linke bis Steckel)
    extracted_data_3 = extract_work_data_for_range(data, "linke", "steckel")

    # Füge alle Daten zusammen
    extracted_data = pd.concat([new_data, extracted_data_1, extracted_data_2, extracted_data_3], ignore_index=True)

    # Kalenderwoche berechnen
    dates = create_header_with_dates(data)
    first_date = pd.to_datetime(dates[0], format='%d.%m.%Y')
    calendar_week = first_date.isocalendar()[1]

    # Flache Spaltenüberschriften erstellen
    columns = ["Nachname", "Vorname"] + [f"{weekday} ({date})" for weekday, date in zip(
        ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], dates
    )]
    extracted_data.columns = columns

    # PDF-Dateiname mit Kalenderwoche erstellen
    pdf_filename = f"Wochenbericht_Fuhrpark_KW{calendar_week:02d}.pdf"

    # PDF generieren
    generate_pdf(extracted_data, calendar_week, pdf_filename)

    # Download-Option
    with open(pdf_filename, "rb") as f:
        st.download_button(
            label="Download als PDF",
            data=f,
            file_name=pdf_filename,
            mime="application/pdf"
        )
