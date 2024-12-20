import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import range_boundaries
from datetime import datetime
from io import BytesIO

# Funktion zum Berechnen der Kalenderwoche
def get_calendar_week(date_value):
    """Berechnet die Kalenderwoche aus einem Datum mit oder ohne Zeit."""
    try:
        # Versuche, Datum direkt zu parsen
        date = datetime.strptime(str(date_value).split()[0], "%Y-%m-%d")  # Entfernt den Zeitanteil
        return date.isocalendar()[1]
    except ValueError:
        raise ValueError(f"Ungültiges Datum: {date_value}")

# Funktion zur Formatierung der Excel-Datei
def style_excel(ws, calendar_week):
    # Farben und Stil für Header und Gitterlinien
    header_fill = PatternFill(start_color="FFADD8E6", end_color="FFADD8E6", fill_type="solid")  # Hellblau für Header
    alt_row_fill = PatternFill(start_color="FFFFF0AA", end_color="FFFFF0AA", fill_type="solid")  # Hellgelb für Zeilen
    title_fill = PatternFill(start_color="FF4682B4", end_color="FF4682B4", fill_type="solid")  # Dunkelblau für KW/Abteilung
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # KW-Eintrag oberhalb der Tabelle
    ws["A1"].value = f"Kalenderwoche: {calendar_week + 1}"  # KW + 1
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].fill = title_fill
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)

    # Abteilung unterhalb der KW
    ws["A2"].value = "Abteilung: Fuhrpark NFC"
    ws["A2"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A2"].fill = title_fill
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ws.max_column)

    # Header-Zeile fett, zentriert und farbig (nur die erste Zeile des Headers)
    for col in ws.iter_cols(min_row=3, max_row=3, min_col=1, max_col=ws.max_column):
        for cell in col:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.font = Font(bold=True, size=12)
            cell.fill = header_fill
            cell.border = thin_border

    # Datenzeilen formatieren (abwechselnd einfärben)
    for row in range(4, ws.max_row + 1):
        for cell in ws[row]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            if row % 2 == 0:  # Jede zweite Zeile einfärben
                cell.fill = alt_row_fill

    # Spaltenbreite anpassen
    adjust_column_width(ws)

    # Erste drei Zeilen fixieren
    ws.freeze_panes = "A4"

# Funktion, um die Spaltenbreite anzupassen
def adjust_column_width(ws):
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2  # Padding für besseren Abstand

# Anwendung in Streamlit
st.title("Übersicht der Wochenarbeit")

uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb["Druck Fahrer"]

    # Extrahiere Daten im Bereich von B11 bis Nachname "Kleiber"
    try:
        extracted_data = extract_range_data(ws, end_name="Kleiber")

        # Berechne die Kalenderwoche aus dem ersten Datum (z. B. Sonntag in Spalte E2)
        first_date_cell = ws.cell(row=2, column=5).value
        calendar_week = get_calendar_week(first_date_cell)

        # Exportiere die Daten als Excel-Datei
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            extracted_data.to_excel(writer, index=False, sheet_name="Bereich B11 bis Kleiber", startrow=2)
            styled_ws = writer.sheets["Bereich B11 bis Kleiber"]
            style_excel(styled_ws, calendar_week)

        st.write("Gefundene Daten:")
        st.dataframe(extracted_data)

        excel_data = output.getvalue()
        st.download_button(
            label="Download als Excel",
            data=excel_data,
            file_name="Bereich_B11_bis_Kleiber.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except ValueError as e:
        st.error(str(e))
