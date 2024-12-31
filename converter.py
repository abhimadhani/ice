import pandas as pd
import json
import streamlit as st
from datetime import datetime

# Define function to process input Excel and produce JSON output
def process_excel_to_json(input_file):
    xls = pd.ExcelFile(input_file)

    # Assuming sheets match structure from the JSON file
    general_details_df = xls.parse("generalDetails")
    invoice_details_df = xls.parse("invoiceDetails")
    item_details_df = xls.parse("itemDetails")

    # General details (taking the first row as representative)
    general_details = general_details_df.iloc[0].fillna('').to_dict()

    # Invoice details
    invoice_details = invoice_details_df.fillna('').to_dict(orient="records")

    # Item details
    item_details = item_details_df.fillna('').to_dict(orient="records")

    # Combine into the desired structure
    output_data = {
        "generalDetailsRequest": general_details,
        "invoiceDetailsRequest": invoice_details,
        "itemDetailsRequest": item_details
    }

    return output_data

# Streamlit Web UI
st.title("Excel to JSON Converter")
st.write("Upload an Excel file to convert it into JSON format.")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")

if uploaded_file is not None:
    st.write("File uploaded successfully.")
    
    try:
       
        
        # Process the file
        xls = pd.ExcelFile(uploaded_file)  # Use the in-memory file directly
        output_data = process_excel_to_json(xls)

        # Display JSON preview
        st.write("Preview of the JSON output:")
        st.json(output_data)

        # Convert JSON to string and encode for download
        json_string = json.dumps(output_data, indent=4)
        st.download_button(
            label="Download JSON file",
            data=json_string,
            file_name="output.json",
            mime="application/json"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
