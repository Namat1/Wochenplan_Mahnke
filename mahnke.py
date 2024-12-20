import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import range_boundaries, get_column_letter
from datetime import datetime
from io import BytesIO

# Funktion zum Überspringen verbundener Zellen
def is_merged_cell_and_wide(ws, row, col, min_width=4):
    """Prüft, ob die gegebene Zelle Teil eines verbundenen Bereichs ist und breiter als `min_width`."""
    for merged_range in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
        if min_row <= row <= max_row and min_col <= col <= max_col:
            if max_col - min_col + 1 > min_width:  # Breite des Bereichs prüfen
                return True
    return False

# Funktion, um den Bereich von "B11" bis Nachname "Steckel" zu finden
def find_range(ws, end_name, column=2):  # Spalte B = 2
    start_row = 11
    end_row = None
    debug_values = []
    for row in range(start_row, ws.max_row + 1):
        value = ws.cell(row=row, column=column).value
        debug_values.append(value)  # Zum Debuggen alle Werte speichern
        if value == end_name:
            end_row = row
            break
    return start_row, end_row, debug_values

# Funktion zur Berechnung der Kalenderwoche
def get_calendar_week(date_value):
    """Berechnet die Kalenderwoche aus einem Datum."""
    try:
        date = datetime.strptime(str(date_value).split()[0], "%Y-%m-%d")  # Entfernt Zeitanteil
        return date.isocalendar()[1]
    except ValueError:
        raise ValueError(f"Ungültiges Datum: {date_value}")

# Extrahiere Daten
def extract_range_data(ws, end_name="Steckel"):
    """Extrahiert Daten von B11 bis einschließlich der Zeile mit Nachname 'Steckel'."""
    start_row, end_row, debug_values = find_range(ws, end_name)
    if not start_row or not end_row:
        raise ValueError(
            f"Bereich bis {end_name} wurde nicht gefunden. "
            f"Gefundene Werte in Spalte B: {debug_values}"
        )

    st.write("Gefundene Werte in Spalte B:", debug_values)  # Debugging-Ausgabe

    relevant_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A."]
    result = []

    # Iteriere durch den Bereich
    for row in range(start_row, end_row + 1, 2):
        if is_merged_cell_and_wide(ws, row, 2):  # Verbundene Zellen
            st.write(f"Überspringe verbundene Zelle in Zeile {row}")
            continue

        lastname = ws.cell(row=row, column=2).value  # Nachname
        firstname = ws.cell(row=row, column=3).value  # Vorname

        # Überspringe Nachnamen "Leer"
        if str(lastname).strip().lower() == "leer":
            st.write(f"Überspringe Nachname 'Leer' in Zeile {row}")
            continue

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

        # Iteriere durch die Wochentage
        for day, (col1, col2) in enumerate(
            [(5, 6), (7, 8), (9, 10), (11, 12), (13, 14), (15, 16), (17, 18)]
        ):
            activity1 = ws.cell(row=activities_row, column=col1).value
            activity2 = ws.cell(row=activities_row, column=col2).value

            # Aktivitäten kombinieren
            activity = " ".join(filter(lambda x: x and x != "0", [str(activity1 or "").strip(), str(activity2 or "").strip()]))
            if any(word in activity for word in relevant_words):
                weekday = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day]
                row_data[weekday] = activity

        result.append(row_data)

    return pd.DataFrame(result)

# Formatierung
def style_excel(ws, calendar_week):
    """Formatierung der Excel-Datei."""
    header_fill = PatternFill(start_color="FFADD8E6", end_color="FFADD8E6", fill_type="solid")
    alt_row_fill = PatternFill(start_color="FFFFF0AA", end_color="FFFFF0AA", fill_type="solid")
    title_fill = PatternFill(start_color="FF4682B4", end_color="FF4682B4", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # KW und Abteilung
    ws["A1"].value = f"Kalenderwoche: {calendar_week + 1}"
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].fill = title_fill
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)

    ws["A2"].value = "Abteilung: Fuhrpark NFC"
    ws["A2"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A2"].fill = title_fill
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ws.max_column)

    # Header
    for col in ws.iter_cols(min_row=3, max_row=3, min_col=1, max_col=ws.max_column):
        for cell in col:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.font = Font(bold=True, size=12)
            cell.fill = header_fill
            cell.border = thin_border

    for row in range(4, ws.max_row + 1):
        for cell in ws[row]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            if row % 2 == 0:
                cell.fill = alt_row_fill

    ws.freeze_panes = "A4"

# Spaltenbreite
def adjust_column_width(ws):
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

# Streamlit-App
st.title("Übersicht der Wochenarbeit")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb["Druck Fahrer"]

    try:
        extracted_data = extract_range_data(ws, end_name="Steckel")
        first_date_cell = ws.cell(row=2, column=5).value
        calendar_week = get_calendar_week(first_date_cell)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            extracted_data.to_excel(writer, index=False, sheet_name="Bereich B11 bis Steckel", startrow=2)
            styled_ws = writer.sheets["Bereich B11 bis Steckel"]
            style_excel(styled_ws, calendar_week)

        st.write("Gefundene Daten:")
        st.dataframe(extracted_data)

        excel_data = output.getvalue()
        st.download_button(
            label="Download als Excel",
            data=excel_data,
            file_name="Bereich_B11_bis_Steckel.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except ValueError as e:
        st.error(str(e))
