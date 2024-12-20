import streamlit as st
import pandas as pd
from io import BytesIO

def filter_excel_data(file):
    # Load the Excel file
    xls = pd.ExcelFile(file)
    if "Druck Fahrer" not in xls.sheet_names:
        st.error("Das Blatt 'Druck Fahrer' wurde nicht gefunden.")
        return None

    df = pd.read_excel(xls, sheet_name="Druck Fahrer")

    # Find the row where "Aushilfsfahrer" is present
    end_row = df[df.isin(["Aushilfsfahrer"]).any(axis=1)].index

    if not end_row.empty:
        # Select data until the row before "Aushilfsfahrer"
        filtered_data = df.iloc[:end_row[0]]
    else:
        filtered_data = df  # Copy the entire sheet if "Aushilfsfahrer" not found

    return filtered_data

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Gefilterte Daten')
    processed_data = output.getvalue()
    return processed_data

# Streamlit app
st.title("Excel Datenfilter und Download")

uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    with st.spinner("Daten werden verarbeitet..."):
        filtered_data = filter_excel_data(uploaded_file)

    if filtered_data is not None:
        st.success("Daten wurden erfolgreich gefiltert.")
        st.write("Gefilterte Daten:")
        st.dataframe(filtered_data)

        excel_data = to_excel(filtered_data)
        st.download_button(
            label="Gefilterte Excel-Datei herunterladen",
            data=excel_data,
            file_name="gefilterte_daten.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
