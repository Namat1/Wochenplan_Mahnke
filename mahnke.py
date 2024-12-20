import streamlit as st
import pandas as pd
from io import BytesIO

# Funktion zur Verarbeitung der Excel-Daten
def process_excel(file):
    # Laden der Excel-Datei
    xls = pd.ExcelFile(file)
    
    # Blatt "Druck Fahrer" laden
    df = pd.read_excel(xls, sheet_name="Druck Fahrer", header=None)

    # Spaltennamen definieren (manuell, da keine Header vorhanden sind)
    df.columns = [f"Col{i}" for i in range(1, len(df.columns) + 1)]

    # Nachnamen und Vornamen extrahieren
    data_start_row = 10  # Start bei Zeile 11 (0-basierter Index = 10)
    df_names = df.iloc[data_start_row:, [1, 2]].dropna()
    df_names.columns = ["Nachname", "Vorname"]

    # Nur bis zum Nachnamen "Steckel"
    df_names = df_names[df_names["Nachname"] != "Steckel"]

    # Datenbereich (Spalte E2 bis Q2)
    date_row = df.iloc[1, 4:17]  # Zeile 2, Spalten E bis Q
    weekdays = date_row.values.tolist()

    # Wichtige Begriffe zur Suche
    keywords = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A."]

    # Ergebnisse initialisieren
    result = []

    for index, row in df_names.iterrows():
        for i, word in enumerate(keywords):
            for date_idx, date in enumerate(weekdays):
                cell_value = df.iloc[index + data_start_row, 4 + date_idx]  # Entsprechende Zellen prüfen
                if pd.notna(cell_value) and word in str(cell_value):
                    result.append({
                        "Nachname": row["Nachname"],
                        "Vorname": row["Vorname"],
                        "Datum": date,
                        "Wochentag": date_row.index[date_idx],
                        "Status": word
                    })

    # Ergebnis als DataFrame
    result_df = pd.DataFrame(result)
    return result_df

# Streamlit App
st.title("Excel-Daten extrahieren und analysieren")

uploaded_file = st.file_uploader("Bitte laden Sie eine Excel-Datei hoch", type="xlsx")

if uploaded_file is not None:
    # Verarbeiten der Datei
    processed_data = process_excel(uploaded_file)

    # Anzeigen der verarbeiteten Tabelle
    st.write("Verarbeitete Daten:")
    st.dataframe(processed_data)

    # Download-Link für die Ergebnisse
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        processed_data.to_excel(writer, index=False, sheet_name="Ergebnisse")
    output.seek(0)

    st.download_button(
        label="Download der Ergebnisse als Excel",
        data=output,
        file_name="Ergebnisse.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
