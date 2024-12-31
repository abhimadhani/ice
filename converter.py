import pandas as pd
import json
from datetime import datetime
import pytz

# Define function to process input Excel and produce JSON output
def process_excel_to_json(input_file, output_file):
    # Load data from Excel sheets
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

    # Determine user's system time zone
    user_timezone = datetime.now().astimezone().tzinfo

    # Write output to JSON
    with open(output_file, "w") as json_file:
        json.dump(output_data, json_file, indent=4, default=lambda x: (x.isoformat() + 'Z').replace('+00:00Z', '.000Z') if isinstance(x, pd.Timestamp) else str(x))

# Input and output file paths
input_file = "input.xlsx"
output_file = "output.json"

# Run the function
process_excel_to_json(input_file, output_file)

print(f"Data from {input_file} has been processed and saved to {output_file}.")
