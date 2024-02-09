import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
# The ID and range of a sample spreadsheet
SPREADSHEET_ID = "1GYcYqKxNklIk1h-uScDPCLtaC-PJyLT-L2idTNRM4bc"
def main():
    credentials = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else: 
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(credentials.to_json())
    try:
        service = build("sheets", "v4", credentials=credentials)
        sheets = service.spreadsheets()

        for row in range(2,7):
            num1 = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"Sheet1!A{row}").execute().get("values")[0][0])
            num2 = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"Sheet1!B{row}").execute().get("values")[0][0])
            
            calculation_result = num1 + num2
            print(f"Processing {num1} + {num2}")

            sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"Sheet1!C{row}",
                                   valueInputOption="USER_ENTERED", body={"values": [[f"{calculation_result}"]]}).execute()
            sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"Sheet1!D{row}",
                                   valueInputOption="USER_ENTERED", body={"values": [["Done"]]}).execute()
    except HttpError as error:
        print(error)

if __name__ == "__main__":
    main()