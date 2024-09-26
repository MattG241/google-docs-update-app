from flask import Flask, render_template, request
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os
import logging

app = Flask(__name__)

# Set up logging
logging.basicConfig(level=logging.INFO)

# Google Sheets API configuration
SPREADSHEET_ID = '1dZ-iYwWfQjAsTFknw9nhI8nocV7fAdzE8NbSiSzRyy0'
CREDENTIALS_FILE = os.environ.get('GOOGLE_CREDENTIALS_FILE')  # Path to your credentials file from environment variable

def find_first_empty_row(service):
    range_name = 'WH to CS OOS Comms!A:A'  # Ensure this matches your sheet name
    result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
    values = result.get('values', [])
    return len(values) + 1

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        order_number = request.form['order_number']
        sku = request.form['sku']

        # Authenticate and build the Google Sheets service
        creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE)
        service = build('sheets', 'v4', credentials=creds)

        # Find the first empty row and update it
        first_empty_row = find_first_empty_row(service)
        range_to_update = f'WH to CS OOS Comms!A{first_empty_row}:B{first_empty_row}'

        values = [[order_number, sku]]
        body = {'values': values}
        service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=range_to_update,
                                                valueInputOption='RAW', body=body).execute()

        logging.info(f'Data updated: Order Number: {order_number}, SKU: {sku}')
        return 'Data updated successfully!'

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
