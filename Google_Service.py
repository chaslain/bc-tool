import copy
import datetime
import pickle
import os
import re
import traceback

from googleapiclient.http import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build




class GoogleService:

    def __init__(self):
        self.flow = None
        self.drive_service = None
        self.sheet_service = None
        self.spreadsheet_object = None
        self.files = []

    def get_spreadsheets(self):
        pass

    def prepare_google_client(self):

        if not os.path.exists("client_secret.json"):
            exit(-1)

        self.flow = InstalledAppFlow.from_client_secrets_file('client_secret.json',
                                                              scopes=["https://www.googleapis.com/auth/drive"])

        if os.path.exists("tokens.pickle"):
            with open("tokens.pickle", "rb") as token:
                credentials = pickle.load(token)
        else:
            credentials = self.flow.run_local_server()
            # save the credentials
            with open("tokens.pickle", "wb") as file:
                pickle.dump(credentials, file)

        self.drive_service = build("drive", "v3", credentials=credentials)
        self.sheet_service = build("sheets", "v4", credentials=credentials)

    def download(self, index):
        file_to_download = self.files["files"][index]
        self.spreadsheet_object = self.sheet_service.spreadsheets().get(spreadsheetId=file_to_download["id"],
                                                                        includeGridData=True).execute()

    def get_sheets(self):
        result = []
        for sheet in self.spreadsheet_object["sheets"]:
            result.append(sheet["properties"]["title"])
        return result

    def get_drive_spreadsheets(self):
        if len(self.files) == 0:
            self.files = self.drive_service.files().list(q="mimeType = 'application/vnd.google-apps.spreadsheet' and trashed=false") \
                .execute()

        return self.files

    def generate_new_sheet(self, sheet_index, rows_to_highlight, sheet_column):
        # first, create new sheet
        new_spreadsheet = self.sheet_service.spreadsheets().create().execute()
        # create request
        request = self.sheet_service.spreadsheets().batchUpdate(spreadsheetId=new_spreadsheet['spreadsheetId'],
                                                                body={"requests":self.form_request(sheet_index, rows_to_highlight, sheet_column)})
        try:
            # send!
            request.execute()

        except HttpError as e:
            self.rollback(new_spreadsheet)
            traceback.print_exc()

    def form_request(self, sheet_index, rows_to_highlight, sheet_column):
        newSheet = copy.deepcopy(self.spreadsheet_object["sheets"][sheet_index])

        rows = []
        maxSize = 0
        for key, row in enumerate(newSheet['data'][0]['rowData']):
            if key == 0:
                continue

            values = []

            makeGreen = key in rows_to_highlight

            shouldAppend = True

            try:
                shouldAppend = row["values"][sheet_column]["effectiveValue"]["numberValue"] != 0
            except KeyError:
                shouldAppend = False
            except IndexError:
                pass

            if shouldAppend:
                if len(row['values']) > maxSize:
                    maxSize = len(row['values'])

                for j in row['values']:
                    cell = j
                    if "dataValidation" in j:
                        del j["dataValidation"]
                        # cleanse data validation as google will not accept it unless we bring the other page too

                    if makeGreen:
                        if 'userEnteredFormat' not in cell:
                            cell['userEnteredFormat'] = {}
                        cell['userEnteredFormat']['backgroundColor'] = {"red": 0, "green": 0.65, "blue": 0, "alpha": 1}

                    values.append(cell)

                rows.append({"values": values})

        requests = []

        # make spreadsheet wide enough for pathalogical rows... default limit is 25
        requests.append({
            "appendDimension":
                {
                    "sheetId": 0,
                    "dimension": "COLUMNS",
                    "length": maxSize
                }
        })

        # add the rows
        requests.append({
            "appendCells":
                {
                    "sheetId": 0,
                    "rows": rows,
                    "fields": "*",
                }
        })

        date = datetime.datetime.today().strftime('%Y-%m-%d')
        title = "Budget Comparator Report - " + date

        # rename
        requests.append({
            "updateSpreadsheetProperties":
                {
                    "properties":
                        {
                            "title": title
                        },
                    "fields": "title"

                },
        })

        return requests

    def rollback(self, new_spreadsheet):
        self.drive_service.files().delete(fileId = new_spreadsheet['spreadsheetId']).execute()
        print("cleaning up...")


def letter_to_number(letter):
    if len(letter) > 0:
        letter.lower()
        return ord(letter) - 97
    return None


def get_match_row_numbers(bank_transaction, sheet_transaction):

    result = []



def get_disparity_list(bank_transaction, sheet_transaction):
    result = []
    result2 = []

    # massage for our purposes- mainly to add a used flag
    sheet_transaction_use = []
    for i in sheet_transaction:
        sheet_transaction_use.append({"amount": abs(i), "used": False})

    for i in bank_transaction:
        found = False
        for key in range(len(sheet_transaction_use)):
            if float(i["amount"]) == sheet_transaction_use[key]["amount"] and not sheet_transaction_use[key]["used"] and not found:
                found = True
                sheet_transaction_use[key]["used"] = True
                result2.append(key+1) # append 1 as the header row throws it off

        if not found:
            result.append(i)

    return result, result2


def formatMoney(money):
    money = re.sub("^0+(?!\.)", "", money)
    add0(money)
    money = '$' + money
    return money


def add0(string, i=0):
    if not re.search("\.\d\d$", string):
        if i > 10:
            return
        if not "." in string:
            string += "."
        string += "0"

        return add0(string,i+1)
    return string