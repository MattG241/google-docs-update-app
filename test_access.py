from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/documents.readonly']
SERVICE_ACCOUNT_FILE = '/Users/tijanamatias/Desktop/google-docs-update-app/autofilloos-e493bc91588b.json'

# Authenticate and create the Google Docs service
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

service = build('docs', 'v1', credentials=credentials)

document_id = '1dZ-iYwWfQjAsTFknw9nhI8nocV7fAdzE8NbSiSzRyy0'

try:
    document = service.documents().get(documentId=document_id).execute()
    print("Document Title:", document.get('title'))
except Exception as e:
    print("Error:", e)
