from flask import Flask, render_template, request
import google.auth
from googleapiclient.discovery import build
import os
app = Flask(__name__)

# Add your Google Sheets API credentials here
SPREADSHEET_ID = '1dZ-iYwWfQjAsTFknw9nhI8nocV7fAdzE8NbSiSzRyy0'

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
        creds, _ = google.auth.default()
        service = build('sheets', 'v4', credentials=creds)

        # Find the first empty row and update it
        first_empty_row = find_first_empty_row(service)
        range_to_update = f'WH to CS OOS Comms!A{first_empty_row}:B{first_empty_row}'  # Update the range to include columns A and B

        values = [[order_number, sku]]
        body = {'values': values}
        service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=range_to_update,
                                                valueInputOption='RAW', body=body).execute()

        return 'Data updated successfully!'

    return render_template('index.html')

if __name__ == '__main__':
    # Run the app on the specified host and port
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 5000)))
