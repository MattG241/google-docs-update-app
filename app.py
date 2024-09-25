import os
from flask import Flask, render_template, request
from google.oauth2 import service_account
from googleapiclient.discovery import build

app = Flask(__name__)

# Path to your service account JSON file
SERVICE_ACCOUNT_FILE = '/Users/tijanamatias/Desktop/google-docs-update-app/credentials.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1dZ-iYwWfQjAsTFknw9nhI8nocV7fAdzE8NbSiSzRyy0'

# Create credentials from the service account file
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Build the service for Google Sheets
service = build('sheets', 'v4', credentials=creds)

def find_first_empty_row():
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

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        order_number = request.form['order_number']
        sku = request.form['sku']

        # Prepare the data to be updated
        values = [[order_number, sku]]
        first_empty_row = find_first_empty_row()
        range_to_update = f'WH to CS OOS Comms!A{first_empty_row}:B{first_empty_row}'

        # Call the Sheets API to update the spreadsheet
        body = {'values': values}
        try:
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_to_update,
                valueInputOption='RAW',
                body=body
            ).execute()
            return 'Data updated successfully!'
        except Exception as e:
            return f'An error occurred: {e}'

    return render_template('index.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))  # Get port from environment variable
    app.run(host='0.0.0.0', port=port, debug=True)  # Bind to all available IP addresses
