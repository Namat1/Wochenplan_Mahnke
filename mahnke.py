import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from io import BytesIO

# Funktion zur Bearbeitung der Excel-Datei
def process_excel(file):
    # Laden der Excel-Datei
    wb = openpyxl.load_workbook(file)
    
    # Arbeitsblätter initialisieren
    if "Druck Fahrer" not in wb.sheetnames:
        return "Fehler: Das Blatt 'Druck Fahrer' fehlt in der Datei."
    src_sheet = wb["Druck Fahrer"]
    dest_sheet = wb.create_sheet(title="Bearbeitet")

    # Daten kopieren
    for row in src_sheet.iter_rows():
        dest_sheet.append([cell.value for cell in row])

    # Konstanten aus dem VBA-Skript
    keep_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A.", "n. A"]
    namen = ["Richter", "Carstensen", "Gebauer", "Pham Manh", "Ohlenroth"]
    namen1 = ["Clemens", "Martin", "Ronny", "Chris", "Nadja"]

    # Bereich für Spalten E bis R löschen (Zeilen 11 bis 270)
    for row in dest_sheet.iter_rows(min_row=11, max_row=270, min_col=5, max_col=18):
        for cell in row:
            if cell.value and not any(word in str(cell.value) for word in keep_words):
                cell.value = None

    # Zeilen 11 und 12 kopieren und Namen einfügen
    for i in range(5):
        dest_sheet.insert_rows(11)
        if i < len(namen):
            dest_sheet.cell(row=11, column=2, value=namen[i])
            dest_sheet.cell(row=11, column=3, value=namen1[i])

    # Zeilen 5 bis 10 löschen
    for _ in range(6):
        dest_sheet.delete_rows(5)

    # "Tour" aus Spalte D entfernen
    for cell in dest_sheet["D"]:
        if cell.value and "Tour" in str(cell.value):
            cell.value = str(cell.value).replace("Tour", "")

    # Zeilen unterhalb von "715" in Spalte A löschen
    for row in dest_sheet.iter_rows(min_col=1, max_col=1):
        for cell in row:
            if cell.value == 715:
                dest_sheet.delete_rows(cell.row + 1, dest_sheet.max_row - cell.row)
                break

    # Markierung der Zeilen 5 bis 14 in den Spalten B, C und D
    fill = PatternFill(start_color="F0E68C", end_color="F0E68C", fill_type="solid")
    for row in dest_sheet.iter_rows(min_row=5, max_row=14, min_col=2, max_col=4):
        for cell in row:
            cell.fill = fill

    # Datei in BytesIO speichern
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# Streamlit-App
st.title("Excel-Bearbeitung basierend auf VBA-Skript")

uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type="xlsx")

if uploaded_file:
    processed_file = process_excel(uploaded_file)

    if isinstance(processed_file, str):
        st.error(processed_file)
    else:
        st.success("Die Datei wurde erfolgreich bearbeitet!")
        st.download_button(
            label="Bearbeitete Datei herunterladen",
            data=processed_file,
            file_name="bearbeitet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
