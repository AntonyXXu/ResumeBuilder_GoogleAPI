from googleapiclient.discovery import build
from google.oauth2 import service_account
import pandas as pd

class ApplicantBuilder:
    def __init__(self, sheets_scope, service_file, drive_scopes, docs_scope):
        self.__sheets_scope = sheets_scope
        self.__service_file = service_file
        # Sheets
        self.__sheet_cred = service_account.Credentials.from_service_account_file(
                            self.__service_file, 
                            scopes = self.__sheets_scope)
        self.__sheet_service = build("sheets", "v4", credentials = self.__sheet_cred)
        self.__sheet = self.__sheet_service.spreadsheets()
       
        # Drive
        self.__drive_scopes = drive_scopes
        self.__drive_cred = service_account.Credentials.from_service_account_file(
                            self.__service_file,
                            scopes = self.__drive_scopes)
        self.__drive_service = build("drive", "v3", credentials=self.__drive_cred)
        
        # Docs 
        self.__doc_scope = docs_scope
        self.__doc_cred = service_account.Credentials.from_service_account_file(
                            self.__service_file,
                            scopes = self.__doc_scope)
        self.__doc_service = build("docs", "v1", credentials=self.__doc_cred)
        
        # Saved Data
        self.__folder_ID = None
        self.__data = None
        self.__dictionary = None
        self.__doc_ID = []

    def read_sheet(self, sheet_ID , sheet_range):
        # Read data from sheet and returns it. Saves data within self.__data
        result = self.__sheet.values().get(
                        spreadsheetId = sheet_ID,
                        range = sheet_range,
                        majorDimension = "COLUMNS"
                        ).execute() 
                        # MajorDimension can be "ROWS" if desired
       
        # Pandas dataframe with index at column 1
        self.__data = pd.DataFrame(
                        result.get("values", [])[1:], 
                        columns = result.get("values", [])[0]
                        )
        return result
    
    def get_data(self):
        # Return data
        return self.__data
    
    def show_data(self):
        # Print self in console during testing    
        print(self.__data)
        return
    
    def write_sheet(self, sheet_ID, sheet_range, write_data):
        # Write to sheet
        update = self.__sheet.values().update(
                    spreadsheetId = sheet_ID, 
                    range = sheet_range, 
                    valueInputOption = "USER_ENTERED", 
                    body = {"values": write_data}
                    ).execute()
        return

    def create_dictionary(self, col_number=0): 
        # Dictionary based on Index (set as applicant for column = 0)
        self.__dictionary = self.__data.set_index(
                            self.__data.columns[col_number]).to_dict()
        return
    
    def get_dictionary(self):
        # Return dictionary
        return self.__dictionary

    def create_folder(self, folder_name , parent_ID=None ):
        # Create Google Drive Folder 
        folder = {
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder" }
        if parent_ID:
            folder["parents"] = parent_ID
        root_folder = self.__drive_service.files().create(
                    body = folder
                    ).execute()
        
        # update folder ID
        self.__folder_ID = root_folder.get("id")
        print ("Folder ID: %s" % self.__folder_ID)
        return
    
    def create_permissions(self, email=None):
        # Grant permissions to users. Otherwise, the files/folder are locked to service accts
        # Anyone can access 
        if not email:
            permissions = {
                "type": "anyone",
                "role": "writer"}
        # For individual users
        else: 
            permissions = {
                "type": "user",
                "role": "writer",
                "emailAddress": email}
       
        self.__drive_service.permissions().create(
                    fileId = self.__folder_ID,
                    body = permissions
                    ).execute()
        return

    def create_doc(self, file_name):
        # Create Google Docs within folder
        doc = {
            "name": file_name,
            "mimeType": "application/vnd.google-apps.document",
            "parents": [self.__folder_ID]}
        newDoc = self.__drive_service.files().create(
                            body = doc
                            ).execute()
        self.__doc_ID.append(newDoc.get("id"))
        return
    
    def get_doc_ID(self):
        return self.__doc_ID

    def doc_name(self, doc_ID):
        # Retrieve the documents contents from the Docs service.
        document = self.__doc_service.documents().get(
                        documentId=doc_ID
                        ).execute()
        return document.get("title")
    
    def write_doc(self, header, write_data, doc_ID):
        index = 1
        headings = [ 
                { # Insert heading
                    "insertText": { 
                        "location": {
                            "index": index },
                        "text": header + "\n"*2 } } ,  
                { # Format keys (headings) in the doc
                    "updateTextStyle": { 
                        "fields": "*",
                        "range": {
                            "startIndex": index,
                            "endIndex": index + len((header)) + 1 } ,                         
                        "textStyle": { 
                            "bold": True,
                            "fontSize": {
                                "magnitude": 14,
                                "unit": "PT" } ,
                            "underline": True } } } , 
                    ]
        result = self.__doc_service.documents().batchUpdate(
                    documentId=doc_ID, 
                    body={"requests": headings}
                    ).execute()
        index += len(header) + 2

        # Insert Dictionary Values
        for key,value in write_data.items():
            # Text to be inputted is dictionary items
            text = str(key) + ": \n\t" + str(value) + "\n"*2
            requests = [ 
                { # Insert text within the doc
                    "insertText": { 
                        "location": {
                            "index": index },
                        "text": text } } ,  
                { # Format keys in the doc
                    "updateTextStyle": { 
                        "fields": "*",
                        "range": {
                            "startIndex": index,
                            "endIndex": index + len((key)) + 1 } ,                         
                        "textStyle": { 
                            "bold": True,
                            "fontSize": {
                                "magnitude": 12,
                                "unit": "PT" } ,
                            "underline": False } } }  
                        ]

            result = self.__doc_service.documents().batchUpdate(
                    documentId=doc_ID, 
                    body={"requests": requests}
                    ).execute()
            # Increase index for input. If index > characters in the doc, write fails
            index += len(text)
