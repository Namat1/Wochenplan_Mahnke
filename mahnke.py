import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# Hauptfunktion für die Bearbeitung
def bearbeiten_und_speichern(uploaded_file):
    # Lade die Excel-Datei
    wb = openpyxl.load_workbook(uploaded_file)
    src_sheet = wb["Druck Fahrer"]

    # Neues Workbook und Arbeitsblatt erstellen
    new_wb = openpyxl.Workbook()
    dest_sheet = new_wb.active
    dest_sheet.title = "Bearbeitet"

    # Inhalte vom Quellblatt kopieren
    for row in src_sheet.iter_rows(values_only=True):
        dest_sheet.append(row)

    # Bereich für die Bearbeitung definieren
    keep_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A.", "n. A"]
    namen = ["Richter", "Carstensen", "Gebauer", "Pham Manh", "Ohlenroth"]
    namen1 = ["Clemens", "Martin", "Ronny", "Chris", "Nadja"]

    # Inhalte löschen, die nicht in der Keep-List sind
    for row in dest_sheet.iter_rows(min_row=11, max_row=270, min_col=5, max_col=18):
        for cell in row:
            if cell.value and all(word not in str(cell.value) for word in keep_words):
                cell.value = None

    # Zusätzliche Bearbeitungen (z. B. Einfügen der Namen)
    for i in range(5):
        dest_sheet.insert_rows(11)
        dest_sheet["B11"] = namen[i] if i < len(namen) else ""
        dest_sheet["C11"] = namen1[i] if i < len(namen1) else ""

    # Tour-Wörter aus Spalte D entfernen
    for cell in dest_sheet["D"]:
        if cell.value and "Tour" in str(cell.value):
            cell.value = str(cell.value).replace("Tour", "")

    # Lösche Zeilen unterhalb der Zeile mit 715 in Spalte A
    for row in dest_sheet.iter_rows(min_col=1, max_col=1):
        for cell in row:
            if cell.value == 715:
                dest_sheet.delete_rows(cell.row + 1, dest_sheet.max_row - cell.row)
                break

    # Ergebnis speichern
    output = BytesIO()
    new_wb.save(output)
    output.seek(0)
    return output

# Streamlit App
st.title("Excel-Bearbeitungs-App")

uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch (mit einem Arbeitsblatt 'Druck Fahrer')", type="xlsx")

if uploaded_file:
    st.success("Datei erfolgreich hochgeladen!")
    if st.button("Bearbeiten und Speichern"):
        output_file = bearbeiten_und_speichern(uploaded_file)
        st.download_button(
            label="Bearbeitetes Excel herunterladen",
            data=output_file,
            file_name="Bearbeitet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
