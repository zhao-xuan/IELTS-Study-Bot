import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1oV5wYMNoIezdXZveD-XQcgKssVTOsj3ZgblrbSYXCvI"
# SAMPLE_RANGE_NAME = "Class Data!A2:E"

def ensure_credential():
  """Shows basic usage of the Sheets API.
  Prints values from a sample spreadsheet.

  Return: the service object
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=53919)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials=creds)
    return service
  except HttpError as err:
    print(err)
  
# Function to update the sheet with new vocabulary
def add_new_vocabulary(sheet, vocabulary_list):
    if not service:
      print("Service is not available.")
      return

    # Prepare the header
    header = ["单词", "可以造句", "理解但造句难", "见过但意义模糊", "完全没见过"]

    # Prepare the data
    data = [header] + [[word, False, False, False, False] for word in vocabulary_list]

    # Prepare the body of the request
    body = {
        'values': data
    }

    # Update the sheet
    try:
        result = sheet.values().update(
            spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range='Sheet2!A1:E100',
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        print(f"{result.get('updatedCells')} cells updated.")
    except HttpError as err:
        print(err)

if __name__ == "__main__":
  # service = ensure_credential()
  # Call the Sheets API
  # sheet = service.spreadsheets()
  # add_new_vocabulary(sheet, ["apple", "banana", "orange"])
  pass