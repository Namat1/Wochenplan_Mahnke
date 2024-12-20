import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Streamlit app definition
st.title("Mahnke Wochenbericht")

# File upload section
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "csv"])

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".xlsx"):
            data = pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith(".csv"):
            data = pd.read_csv(uploaded_file)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()
else:
    st.info("Please upload an Excel or CSV file to proceed.")
    st.stop()

st.write("Original Data:", data)

# Retainable words and names
keep_words = ["Ausgleich", "Krank", "Sonderurlaub", "Urlaub", "Berufsschule", "Fahrschule", "n.A.", "n. A"]
namen = ["Richter", "Carstensen", "Gebauer", "Pham Manh", "Ohlenroth"]
namen1 = ["Clemens", "Martin", "Ronny", "Chris", "Nadja"]

# Filter data based on keep words
def filter_data(data):
    for col in data.columns[4:18]:  # Columns E to R
        data[col] = data[col].apply(
            lambda x: x if any(word in str(x) for word in keep_words) else ""
        )
    return data

filtered_data = filter_data(data.copy())
st.write("Filtered Data:", filtered_data)

# Insert names and process additional logic
def additional_processing(data):
    # Dynamische Anzahl der Spalten bestimmen
    num_columns = len(data.columns)
    empty_row = [""] * num_columns

    # Duplicate rows und Namen einfügen
    for i in range(5):
        data.loc[len(data)] = empty_row  # Leere Zeile hinzufügen
        if i < len(namen):
            data.iloc[-1, 1] = namen[i]  # Namen in Spalte B einfügen
        if i < len(namen1):
            data.iloc[-1, 2] = namen1[i]  # Namen in Spalte C einfügen

    # Zeilen 5 bis 10 löschen
    data = data.drop(index=range(5, 11), errors="ignore")

    # Wort "Tour" aus Spalte D entfernen
    if "D" in data.columns:
        data["D"] = data["D"].astype(str).str.replace("Tour", "", regex=False)

    return data

processed_data = additional_processing(filtered_data.copy())
st.write("Processed Data:", processed_data)

# Save to a new Excel file and offer for download
def convert_df_to_excel(data):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        data.to_excel(writer, index=False, sheet_name="Processed Data")
    processed_file = output.getvalue()
    return processed_file

excel_data = convert_df_to_excel(processed_data)
st.download_button(
    label="Download Processed Report as Excel",
    data=excel_data,
    file_name="processed_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
