from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import os

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = ""

def main():
    credentials = None
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credentials.to_json())
    
    try:
        service = build("sheets", "v4", credentials=credentials)
        sheets = service.spreadsheets()

        # Assuming you want to check the first 100 rows (adjust as needed)
        range_to_check = "Blad1!A2:M101" 

        result = sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=range_to_check).execute()

        values = result.get("values", [])

        for i, row in enumerate(values, start=2):

            # Check if A to F filled
            if len(row) >= 6 and all(row[:6]):
                print(f"Row {i}: Hotelgegevens ingevuld")
            else:
                print(f"Row {i}: Hotelgegevens missing")

            # Check G to K
            if len(row) >= 11 and any(not cell for cell in row[6:11]): 
                print(f"Row {i}: VCC missing")

            # Check L
            if len(row) >= 12 and not row[11]:
                print(f"Row {i}: VCC not sent to hotel")

            # Check M
            if len(row) >= 13 and not row[12]:
                print(f"Row {i}: Old VCC not canceled")


    except HttpError as error:
        print(error)

if __name__ == "__main__":
    main()