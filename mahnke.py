import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import range_boundaries
from io import BytesIO

# Funktion zum Überspringen verbundener Zellen
def is_merged_cell(ws, row, col):
    """Prüft, ob die gegebene Zelle in einem verbundenen Bereich liegt."""
    for merged_range in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
        if min_row <= row <= max_row and min_col <= col <= max_col:
            return True
    return False

# Funktion, um den Bereich von "Adler" bis "Kleiber" zu finden
def find_range(ws, start_name, end_name, column=2):  # Spalte B = 2
    start_row = None
    end_row = None
    for row in range(1, ws.max_row + 1):
        value = ws.cell(row=row, column=column).value
        if value == start_name:
            start_row = row
        if value == end_name:
            end_row = row
        if start_row and end_row:
            break
    return start_row, end_row

# Extrahiere Daten zwischen "Adler" und "Kleiber"
def extract_range_data(ws, start_name="Adler", end_name="Kleiber"):
    start_row, end_row = find_range(ws, start_name, end_name)
    if not start_row or not end_row:
        raise ValueError(f"Bereich zwischen {start_name} und {end_name} wurde nicht gefunden.")

    relevant_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A."]
    result = []

    # Iteriere durch den Bereich und überspringe verbundene Zellen
    for row in range(start_row, end_row + 1, 2):  # Nimm nur ungerade Zeilen für Namen
        if is_merged_cell(ws, row, 2):  # Überspringe verbundene Zellen in Spalte B
            continue

        lastname = ws.cell(row=row, column=2).value  # Nachname
        firstname = ws.cell(row=row, column=3).value  # Vorname
        activities_row = row + 1  # Aktivitäten sind eine Zeile darunter

        row_data = {
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

        # Iteriere durch die Wochentage und lese Aktivitäten aus Spalten E bis R
        for day, (col1, col2) in enumerate(
            [(5, 6), (7, 8), (9, 10), (11, 12), (13, 14), (15, 16), (17, 18)]
        ):
            activity1 = ws.cell(row=activities_row, column=col1).value
            activity2 = ws.cell(row=activities_row, column=col2).value

            # Kombiniere beide Aktivitäten, falls sie nicht leer oder "0" sind
            activity = " ".join(filter(lambda x: x and x != "0", [str(activity1 or "").strip(), str(activity2 or "").strip()]))
            if any(word in activity for word in relevant_words):
                weekday = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day]
                row_data[weekday] = activity

        result.append(row_data)

    return pd.DataFrame(result)

# Anwendung in Streamlit
st.title("Übersicht der Wochenarbeit")

uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb["Druck Fahrer"]

    # Extrahiere Daten im Bereich von "Adler" bis "Kleiber"
    try:
        extracted_data = extract_range_data(ws, start_name="Adler", end_name="Kleiber")
        st.write("Inhalt der Tabelle:")
        st.dataframe(extracted_data)

        # Exportiere die Daten als Excel-Datei
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            extracted_data.to_excel(writer, index=False, sheet_name="Bereich Adler bis Kleiber")
        excel_data = output.getvalue()

        st.download_button(
            label="Download als Excel",
            data=excel_data,
            file_name="Bereich_Adler_bis_Kleiber.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except ValueError as e:
        st.error(str(e))
