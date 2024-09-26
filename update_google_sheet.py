from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime

# Path to your service account JSON file
SERVICE_ACCOUNT_FILE = '/Users/tijanamatias/Desktop/google-docs-update-app/credentials.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Create credentials from the service account file
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Build the service for Google Sheets
service = build('sheets', 'v4', credentials=creds)

# Spreadsheet ID
spreadsheet_id = '1dZ-iYwWfQjAsTFknw9nhI8nocV7fAdzE8NbSiSzRyy0'

def find_first_empty_row():
    # Get the current values in column A
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range='WH to CS OOS Comms!A:A'  # Adjust the sheet name if necessary
    ).execute()
    values = result.get('values', [])

    # Find the first empty row
    if not values:
        return 1  # If no data, start at row 1
    return len(values) + 1  # Return the next empty row

def update_google_sheet(order_number, sku):
    # Prepare the data to be updated
    date_today = datetime.now().strftime('%Y-%m-%d')  # Format the date
    values = [
        [date_today, order_number, sku]  # Fill the first row with date, order number, SKU
    ]
    body = {
        'values': values
    }

    # Find the first empty row in column A
    first_empty_row = find_first_empty_row()
    range_name = f'WH to CS OOS Comms!A{first_empty_row}'  # Specify the range based on the first empty row

    # Call the Sheets API to update the spreadsheet
    try:
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='RAW',
            body=body
        ).execute()
        print("Data successfully added to the spreadsheet.")
    except Exception as e:
        print("An error occurred:", e)

if __name__ == "__main__":
    order_number = input("Enter the order number: ")
    sku = input("Enter the SKU: ")
    update_google_sheet(order_number, sku)
