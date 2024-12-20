import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
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

        # Iteriere durch die Wochentage
        for day, (activity_start_col, date_col) in enumerate(
            [(4, 4), (6, 6), (8, 8), (10, 10), (12, 12), (14, 14), (16, 16)]
        ):
            activity = df.iloc[activities_row, activity_start_col]
            if any(word in str(activity) for word in relevant_words):
                weekday = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day]
                row[weekday] = activity

        result.append(row)
        row_index += 2  # Zwei Zeilen weiter

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

# Funktion, um die Tabelle optisch aufzubereiten
def style_excel(ws, calendar_week):
    # Farben und Stil für Header und Gitterlinien
    header_fill = PatternFill(start_color="FFCCFFCC", end_color="FFCCFFCC", fill_type="solid")  # Grün für Header
    alt_row_fill = PatternFill(start_color="FFF0F0F0", end_color="FFF0F0F0", fill_type="solid")  # Grau für Zeilen
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # KW-Eintrag oberhalb der Tabelle
    ws["A1"].value = f"Kalenderwoche: {calendar_week}"
    ws["A1"].font = Font(bold=True)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)

    # Abteilung unterhalb der KW
    ws["A2"].value = "Abteilung: Fuhrpark NFC"
    ws["A2"].font = Font(bold=True)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ws.max_column)

    # Header-Zeile fett, zentriert und farbig (nur die erste Zeile des Headers)
    for col in ws.iter_cols(min_row=3, max_row=3, min_col=1, max_col=ws.max_column):
        for cell in col:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.font = Font(bold=True)
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

# Funktion, um die Spaltenbreite anzupassen
def adjust_column_width(ws):
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2  # Padding für besseren Abstand

# Streamlit App
st.title("Übersicht der Wochenarbeit")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    # Lade die Excel-Datei
    wb = load_workbook(uploaded_file, data_only=True)
    sheet = wb["Druck Fahrer"]
    data = pd.DataFrame(sheet.values)

    # Extrahiere die Daten und das Datum
    extracted_data = extract_work_data(data)
    dates = create_header_with_dates(data)

    # Kalenderwoche berechnen
    first_date = pd.to_datetime(dates[0], format='%d.%m.%Y')
    calendar_week = first_date.isocalendar()[1]

    # Flache Spaltenüberschriften erstellen
    columns = ["Nachname", "Vorname"] + [f"{weekday} ({date})" for weekday, date in zip(
        ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], dates
    )]
    extracted_data.columns = columns

    # Debugging: Zeige die Daten
    st.write("Inhalt von extracted_data:")
    st.dataframe(extracted_data)

    # Daten als Excel-Datei exportieren
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        extracted_data.to_excel(writer, index=False, sheet_name="Wochenübersicht", startrow=2)
        ws = writer.sheets["Wochenübersicht"]
        style_excel(ws, calendar_week)  # Optische Anpassungen und KW-/Abteilungs-Eintrag
    excel_data = output.getvalue()

    # Download-Option
    st.download_button(
        label="Download als Excel",
        data=excel_data,
        file_name="Wochenübersicht.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
