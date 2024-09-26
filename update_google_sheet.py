from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime

# Service account credentials
SERVICE_ACCOUNT_INFO = {
    "type": "service_account",
    "project_id": "autofilloos",
    "private_key_id": "c895475a1d6c10db19809969f4a441bc07de0650",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCmHP92i67QJgwv\nifWxmvqxAIQs1ksYNxhr7PtVNo7BHeOv3/PGyokAFIiA1RrWPwDWThDkXHYWMbmU\n23O5KRkwcTF/h96P9KDRuEczDIU2HwJzRF+ZE66znE5nSiM29wQ68wvnN5db5iBG\nxvZkWCfomSHg93BUMgaCmeIcKFnLBTJKkS8TYN/UqWx34PU4OePojqOWibKlPEri\nb6HhPOVQchYadn+7e0LFhYqowp6sI7+eRhbnJB+lm/jb8aAD9LucM2GTDpZTmQ0/\n15Mq/05JBi9+NRd5EN+DKSauknzbkTg7+1HaHnt+uRDIYgUTBykT0LSR4/iZ5TZX\nntyLQsAPAgMBAAECggEAOoZo8ylPlAJztKQUrlh+DrKx3uI5XvS6Y/wAqQspzJxt\nRd/PkbB2CFMzrMBoTiewcdDbXrm82SD305xl70ytlUWsPNRv86QqrPkSDMhSfrj6\nMgZa8CHhIWLmtLmIIqtxEBvli7coWrZ/lLAwyzXMCcU6DHrhVqixZn41DdqhmEdQ\nahyBKX8X3P6J44P4je9PY7z4oo/ES7Uqgp99xZ0+lrIKBwcNB7/scoZi4AuGsTQe\nGKr7Su7fgTixt7V04H8KMLOVXLpVqMreTGRaL/5Azngc6HZhbpqoZJbPAqkww10D\nHnqj/c7nEJI/dIC1eGwG8PEGl8vRgQ4B1la3+OmoEQKBgQDjWZhok3IwE0Fg+dzg\n979gaYUSFp9Kp466M9Vr3GudcjnopA/gZ4OJLVQ4qngRwn+rBA3IwJv+GydH/AuF\nIApmr58y9G8xUd1T9QqF9m34BjMZSFj3LCIfqH89mKB60AueSolHfiTd8FgjIdW8\n/bCisLegwUpyZiXE+NDmZLGR4wKBgQC7C+GF1w00iHV4/MPAUu3N5qF6nwV+U3m9\n62dcheG1oWk9qPIkJFlsAJnsE4098U3gFd957FKqFoFXh9rKLVWw7Wbbh6Gm6voQ\ntfBkA1vHEymHP2FkyQ/u2XP3ILhANw1YiQIUBa+LVbs+Pr/K1Nx8WooD6CqM9J+i\noyk/tMvA5QKBgE3PdkgkXqpxjKjCG4SrhkZbFv4v2+jTHBhCcULvN621UHh83iox\ng2VJrE+QmHOLm+JOCuGwejMn2/PZIaA4bRbj+JqZ6gx5NkTr0uQyiUSf6pE2n6xI\n4IzxQEs2l4Yw+ij83asoUznabm/nvp1mPjQQQ2izfuVUbIzTk7umrtd9AoGBAJfv\nj7K7PBPHIL01fQDVnCubwuGrGLhDsGlwNZa3fd+fDLC0cnSfPi/30RAt1ZZSU7LJ\nsa0FJSTagRgL19JQvwGn5dw/MTU4PAak828aN0vfKeWdu0w18oZPBt7gKiqnTWT0\nbca705t3+VAXgo2NGMi+dsuzpBS5hI6EwLXp83RtAoGBAJ1Zguh8tEX8aFEtDORu\nU2QYkFGT0lNNbQf8ApHqYYBQJ1igzEP2DH4bvLthnYAN+aAGaNPSB3htr4UPpFXU\nD0dIEifoEACLOqbO6QKXb2YNbWROGbBj+BXrUyKT2+iproIdP0nXRhTaO1mVX5Gr\nKtglNDOtVbmxPVBIXXvwPIXI\n-----END PRIVATE KEY-----\n",
    "client_email": "oosautofill@autofilloos.iam.gserviceaccount.com",
    "client_id": "118151718935133199043",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/oosautofill%40autofilloos.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

# Define the required scopes for the Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Create credentials from the service account info
creds = service_account.Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=SCOPES)

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
