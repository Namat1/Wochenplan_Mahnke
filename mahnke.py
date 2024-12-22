import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from io import BytesIO

st.title("Wochenarbeitsbericht Fuhrpark")
st.info("Die rot und grün gefärbten Zeilen müssen manuell eingetragen werden. Dispo und Aushilfen!")

# Funktionen bleiben unverändert
def extract_work_data_for_range(df, start_value, end_value):
    relevant_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule"]
    excluded_words = ["Hoffahrer", "Waschteam", "Aushilfsfahrer"]
    result = []

    df.iloc[:, 1] = df.iloc[:, 1].astype(str).str.strip().str.lower()
    if start_value not in df.iloc[:, 1].values or end_value not in df.iloc[:, 1].values:
        st.error(f"Die Werte '{start_value}' oder '{end_value}' wurden in Spalte B nicht gefunden.")
        st.stop()

    start_index = df[df.iloc[:, 1] == start_value].index[0]
    end_index = df[df.iloc[:, 1] == end_value].index[0]

    for row_index in range(start_index, end_index + 1):
        lastname = str(df.iloc[row_index, 1]).strip().title()
        firstname = str(df.iloc[row_index, 2]).strip().title()
        if not lastname or not firstname or lastname == "None" or firstname == "None":
            continue

        activities_row = row_index + 1
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

        for day, (col1, col2) in enumerate([(4, 5), (6, 7), (8, 9), (10, 11), (12, 13), (14, 15), (16, 17)]):
            activity1 = str(df.iloc[activities_row, col1]).strip()
            activity2 = str(df.iloc[activities_row, col2]).strip()
            activity = " ".join(filter(lambda x: x and x != "0", [activity1, activity2])).strip()
            if (any(word in activity for word in relevant_words) and
                    not any(excluded in activity for excluded in excluded_words)):
                weekday = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][day]
                row[weekday] = activity

        result.append(row)

    return pd.DataFrame(result)


def create_header_with_dates(df):
    dates = [
        pd.to_datetime(df.iloc[1, 4]).strftime('%d.%m.%Y'),
        pd.to_datetime(df.iloc[1, 6]).strftime('%d.%m.%Y'),
        pd.to_datetime(df.iloc[1, 8]).strftime('%d.%m.%Y'),
        pd.to_datetime(df.iloc[1, 10]).strftime('%d.%m.%Y'),
        pd.to_datetime(df.iloc[1, 12]).strftime('%d.%m.%Y'),
        pd.to_datetime(df.iloc[1, 14]).strftime('%d.%m.%Y'),
        pd.to_datetime(df.iloc[1, 16]).strftime('%d.%m.%Y'),
    ]
    return dates

# Hauptprogramm mit Fortschrittsanzeige
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
    progress_bar.progress(100)

    # Fertig
    progress_status.text("Verarbeitung abgeschlossen!")
    st.success("Die Excel-Datei wurde erfolgreich verarbeitet.")

    # Download-Button
    st.download_button(
        label="Download als Excel",
        data=output.getvalue(),
        file_name="Fuhrpark_Wochenbericht.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


    # KW-Eintrag oberhalb der Tabelle
    ws["A1"].value = f"Kalenderwoche: {calendar_week + 1}"  # KW + 1
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].fill = title_fill
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)

    # Abteilung unterhalb der KW
    ws["A2"].value = "Abteilung: Fuhrpark - NFC"
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

    # Formatierung für die letzten 6 Zeilen (abwechselnd grün und hellgrün)
    for row in range(ws.max_row - 5, ws.max_row + 1):
        for cell in ws[row]:
            if (row - (ws.max_row - 5)) % 2 == 0:  # Ungerade Zeilen
                cell.fill = last_row_fill_odd
            else:  # Gerade Zeilen
                cell.fill = last_row_fill_even

    # Formatierung für die ersten 6 Zeilen (abwechselnd rot und hellrot)
    for row in range(4, 4 + num_new_rows):
        for cell in ws[row]:
            if (row - 4) % 2 == 0:  # Ungerade Zeilen
                cell.fill = new_row_fill_odd
            else:  # Gerade Zeilen
                cell.fill = new_row_fill_even

    # Spaltenbreite anpassen
    adjust_column_width(ws)

    # Erste drei Zeilen fixieren
    ws.freeze_panes = "A4"

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
st.title("Wochenarbeitsbericht Fuhrpark")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    # Lade die Excel-Datei
    wb = load_workbook(uploaded_file, data_only=True)
    sheet = wb["Druck Fahrer"]
    data = pd.DataFrame(sheet.values)

    if uploaded_file:
    # Fortschrittsanzeige initialisieren
    progress_bar = st.progress(0)
    progress_status = st.empty()

    # 1. Fortschritt: Excel-Datei laden
    progress_status.text("Lade die Excel-Datei...")
    wb = load_workbook(uploaded_file, data_only=True)
    progress_bar.progress(20)

    # 2. Fortschritt: Daten auslesen
    progress_status.text("Lese Daten aus der Excel-Datei...")
    sheet = wb["Druck Fahrer"]
    data = pd.DataFrame(sheet.values)
    progress_bar.progress(40)

    # 3. Fortschritt: Daten extrahieren
    progress_status.text("Extrahiere relevante Daten...")
    extracted_data_1 = extract_work_data_for_range(data, "adler", "steckel")
    progress_bar.progress(60)

    # 4. Fortschritt: Daten vorbereiten
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

    # 5. Fortschritt: Excel-Datei erstellen
    progress_status.text("Erstelle Excel-Datei mit den verarbeiteten Daten...")
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        extracted_data.to_excel(writer, index=False, sheet_name="Wochenübersicht", startrow=2)
        ws = writer.sheets["Wochenübersicht"]
    progress_bar.progress(100)

    # Fortschrittsanzeige abschließen
    progress_status.text("Verarbeitung abgeschlossen!")
    st.success("Die Excel-Datei wurde erfolgreich verarbeitet.")

    # Erstelle 6 Zeilen für die Mitarbeiter oberhalb von "Adler"
    new_data = pd.DataFrame([{
        "Nachname": "Carstensen", "Vorname": "Martin", "Sonntag": "", "Montag": "", "Dienstag": "", 
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
    extracted_data_1 = extract_work_data_for_range(data, "adler", "steckel")

    

   

    # Füge alle Daten zusammen
    extracted_data = pd.concat([new_data, extracted_data_1,], ignore_index=True)

    # Kalenderwoche berechnen
    dates = create_header_with_dates(data)
    first_date = pd.to_datetime(dates[0], format='%d.%m.%Y')
    calendar_week = first_date.isocalendar()[1]

    # Flache Spaltenüberschriften erstellen
    columns = ["Nachname", "Vorname"] + [f"{weekday} ({date})" for weekday, date in zip(
        ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"], dates
    )]
    extracted_data.columns = columns

    # Excel-Dateiname mit Kalenderwoche erstellen
    excel_filename = f"Fuhrpark_Meldung_KW: {calendar_week + 1}.xlsx"


    # Daten als Excel-Datei exportieren
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        extracted_data.to_excel(writer, index=False, sheet_name="Wochenübersicht", startrow=2)
        ws = writer.sheets["Wochenübersicht"]
        style_excel(ws, calendar_week, len(new_data), len(extracted_data))  # Optische Anpassungen und KW-/Abteilungs-Eintrag
    excel_data = output.getvalue()

    # Download-Option
    st.download_button(
        label="Download als Excel",
        data=excel_data,
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
