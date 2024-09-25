import os
import logging
from flask import Flask, render_template, request
from google.oauth2 import service_account
from googleapiclient.discovery import build

app = Flask(__name__)

# Configure logging
logging.basicConfig(level=logging.DEBUG)

# Global variables for Google Sheets service and credentials
service = None
creds = None

# Path to your service account JSON file
SERVICE_ACCOUNT_FILE = '/Users/tijanamatias/Desktop/google-docs-update-app/credentials.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1dZ-iYwWfQjAsTFknw9nhI8nocV7fAdzE8NbSiSzRyy0'

# Initialize the Google Sheets service
try:
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    logging.info("Google Sheets service created successfully.")
except Exception as e:
    logging.error(f"Failed to create Google Sheets service: {e}")
    service = None  # Set service to None if initialization fails

def find_first_empty_row():
    if service is None:
        logging.error("Google Sheets service is not initialized.")
        raise Exception("Google Sheets service is not initialized.")

    try:
        # Get the current values in column A
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range='WH to CS OOS Comms!A:A'
        ).execute()
        values = result.get('values', [])

        # Find the first empty row
        if not values:
            return 1  # If no data, start at row 1
        return len(values) + 1
    except Exception as e:
        logging.error(f"Error finding first empty row: {e}")
        raise

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        order_number = request.form['order_number']
        sku = request.form['sku']

        # Ensure the Google Sheets service is available
        if service is None:
            logging.error("Google Sheets service is not available.")
            return "An error occurred: Google Sheets service is not available."

        # Prepare the data to be updated
        values = [[order_number, sku]]
        try:
            first_empty_row = find_first_empty_row()
            range_to_update = f'WH to CS OOS Comms!A{first_empty_row}:B{first_empty_row}'

            # Call the Sheets API to update the spreadsheet
            body = {'values': values}
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_to_update,
                valueInputOption='RAW',
                body=body
            ).execute()
            return 'Data updated successfully!'
        except Exception as e:
            logging.error(f"Error updating spreadsheet: {e}")
            return f'An error occurred: {e}'

    return render_template('index.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))  # Get port from environment variable
    app.run(host='0.0.0.0', port=port, debug=True)  # Bind to all available IP addresses
