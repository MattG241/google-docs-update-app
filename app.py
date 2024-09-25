import os
import logging
import json
from flask import Flask, render_template, request
from google.oauth2 import service_account
from googleapiclient.discovery import build

app = Flask(__name__)

# Configure logging
logging.basicConfig(level=logging.DEBUG)

# Global variables for Google Sheets service and credentials
service = None
creds = None

# Define the required scopes for your app (Google Sheets read and write access)
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1dZ-iYwWfQjAsTFknw9nhI8nocV7fAdzE8NbSiSzRyy0'  # Replace with your actual Spreadsheet ID

# Load Google service account credentials from the environment variable
SERVICE_ACCOUNT_INFO = os.getenv('GOOGLE_APPLICATION_CREDENTIALS_JSON')

if SERVICE_ACCOUNT_INFO:
    try:
        creds = service_account.Credentials.from_service_account_info(
            json.loads(SERVICE_ACCOUNT_INFO), scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        logging.info("Google Sheets service created successfully.")
    except Exception as e:
        logging.error(f"Failed to create Google Sheets service: {e}")
        service = None  # Set service to None if initialization fails
else:
    logging.error("Google Sheets credentials not found in environment variables.")
    service = None

# Function to find the first empty row in the Google Sheet
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

# Route for the form to input order data
@app.route('/', methods=['GET', 'POST'])
def index():
    success_message = None  # Initialize success message

    if request.method == 'POST':
        order_number = request.form['order_number']
        sku = request.form['sku']

        # Ensure the Google Sheets service is available
        if service is None:
            logging.error("Google Sheets service is not available.")
            return "An error occurred: Google Sheets service is not available."

        # Prepare the data to be updated
        values = [[order_number, sku]]
        logging.debug(f"Updating Google Sheet with values: {values}")

        try:
            first_empty_row = find_first_empty_row()
            range_to_update = f'WH to CS OOS Comms!A{first_empty_row}:B{first_empty_row}'

            # Log the range and body before execution
            logging.debug(f"Range to update: {range_to_update}")
            body = {'values': values}
            logging.debug(f"Body to update: {body}")

            # Call the Sheets API to update the spreadsheet
            response = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_to_update,
                valueInputOption='RAW',
                body=body
            ).execute()

            logging.info(f"Response from Google Sheets API: {response}")
            success_message = "OOS successfully added to Google Sheet."  # Set success message
        except Exception as e:
            logging.error(f"Error updating spreadsheet: {e}")
            return f'An error occurred: {e}'

    return render_template('index.html', success_message=success_message)

# Run the Flask app
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))  # Get port from environment variable
    app.run(host='0.0.0.0', port=port, debug=True)  # Bind to all available IP addresses
