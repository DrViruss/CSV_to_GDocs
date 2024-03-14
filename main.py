import pandas as pd
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

TITLE = "customers-1000"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SHEET_COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
            'V', 'W', 'X', 'Y', 'Z']


def auth():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds


def create(_service, _title):
    try:
        spreadsheet = {"properties": {"title": _title}}
        spreadsheet = (
            _service.spreadsheets()
            .create(body=spreadsheet, fields="spreadsheetId")
            .execute()
        )
        print(f"Spreadsheet ID: {(spreadsheet.get('spreadsheetId'))}")
        return spreadsheet.get("spreadsheetId")
    except HttpError as error:
        print(f"An error occurred: {error}")
        exit()


def update(_service, _sheet_id, _value_input_option, _data):
    edit_range = "A1:" + str(SHEET_COLS[_data.columns.size - 1] + str(_data.shape[0]))

    try:
        body = {
            "values": _data.values[:, 1:].tolist()
        }
        result = (
            _service.spreadsheets()
            .values()
            .update(
                spreadsheetId=_sheet_id,
                range=edit_range,
                valueInputOption=_value_input_option,
                body=body,
            )
            .execute()
        )
        print(f"{result.get('updatedCells')} cells updated.")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


if __name__ == __name__:
    auth = auth()
    service = build("sheets", "v4", credentials=auth)
    sheet_id = create(service, TITLE)
    update(service, sheet_id, 'USER_ENTERED', pd.read_csv(TITLE+".csv"))
