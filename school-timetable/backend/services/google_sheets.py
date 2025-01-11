from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os
import pickle
from fastapi import HTTPException

class GoogleSheetsService:
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    CREDENTIALS_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config', 'credentials.json')
    TOKEN_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config', 'token.pickle')
    OAUTH_HOST = "localhost"
    OAUTH_PORT = 8000
    OAUTH_REDIRECT_URI = f"http://{OAUTH_HOST}:{OAUTH_PORT}/oauth2callback"

    def __init__(self):
        self.service = self.get_service()

    def get_service(self):
        creds = None
        try:
            if os.path.exists(self.TOKEN_FILE):
                print(f"Found existing token file at {self.TOKEN_FILE}")
                with open(self.TOKEN_FILE, 'rb') as token:
                    creds = pickle.load(token)
                    print("Successfully loaded credentials from token file")
            
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    print("Refreshing expired credentials")
                    creds.refresh(Request())
                else:
                    print("Starting new OAuth2 flow...")
                    if not os.path.exists(self.CREDENTIALS_FILE):
                        raise FileNotFoundError(f"Credentials file not found at {os.path.abspath(self.CREDENTIALS_FILE)}")
                    
                    os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
                    
                    flow = InstalledAppFlow.from_client_secrets_file(
                        self.CREDENTIALS_FILE, 
                        self.SCOPES,
                        redirect_uri=self.OAUTH_REDIRECT_URI
                    )
                    
                    creds = flow.run_local_server(
                        host=self.OAUTH_HOST,
                        port=self.OAUTH_PORT,
                        success_message="Authorization successful! You can close this window.",
                        authorization_prompt_message="Please visit this URL to authorize access to your Google Sheets:"
                    )
                    print("OAuth2 flow completed successfully")
                
                os.makedirs(os.path.dirname(self.TOKEN_FILE), exist_ok=True)
                with open(self.TOKEN_FILE, 'wb') as token:
                    pickle.dump(creds, token)
                    print("Saved credentials to token file")
            
            print("Building Google Sheets service...")
            service = build('sheets', 'v4', credentials=creds)
            print("Successfully built Google Sheets service")
            return service
            
        except Exception as e:
            print(f"Error in get_service: {str(e)}")
            if isinstance(e, FileNotFoundError):
                raise HTTPException(
                    status_code=500,
                    detail=f"Configuration error: Credentials file not found. Please check {self.CREDENTIALS_FILE}"
                )
            if "Address already in use" in str(e):
                raise HTTPException(
                    status_code=500,
                    detail="OAuth server port is already in use. Please wait a moment and try again."
                )
            raise HTTPException(
                status_code=500,
                detail=f"Authentication error: {str(e)}"
            )

    async def get_sheet_data(self, spreadsheet_id: str, range: str):
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range
            ).execute()
            return {"values": result.get('values', [])}
        except Exception as e:
            print(f"Google Sheets API Error: {str(e)}")
            raise HTTPException(status_code=500, detail=f"Google Sheets API Error: {str(e)}")

    async def update_sheet_data(self, spreadsheet_id: str, range: str, values: list):
        try:
            body = {'values': values}
            result = self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            return {"updated_cells": result.get('updatedCells')}
        except Exception as e:
            print(f"Google Sheets API Error: {str(e)}")
            raise HTTPException(status_code=500, detail=f"Google Sheets API Error: {str(e)}")

    async def setup_structure(self, spreadsheet_id: str):
        try:
            print(f"Setting up structure for spreadsheet: {spreadsheet_id}")
            
            sheets = [
                {'properties': {'title': 'Teachers'}},
                {'properties': {'title': 'Classes'}},
                {'properties': {'title': 'Subjects'}},
                {'properties': {'title': 'Config'}},
                {'properties': {'title': 'Summary'}}
            ]
            
            batch_update_body = {
                'requests': [{'addSheet': sheet} for sheet in sheets]
            }
            
            try:
                print("Sending batch update request to create sheets")
                result = self.service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body=batch_update_body
                ).execute()
                print(f"Batch update successful: {result}")
                return {"message": "Structure setup completed"}
            except Exception as e:
                print(f"Google Sheets API Error in batch update: {str(e)}")
                error_msg = str(e)
                if "Invalid value" in error_msg:
                    raise HTTPException(status_code=400, detail=f"Invalid spreadsheet ID: {spreadsheet_id}")
                elif "insufficient permission" in error_msg.lower():
                    raise HTTPException(status_code=403, detail="Insufficient permissions. Make sure you have edit access to the spreadsheet.")
                else:
                    raise HTTPException(status_code=500, detail=f"Google Sheets API Error: {error_msg}")
        except Exception as e:
            print(f"Setup Structure Error: {str(e)}")
            if isinstance(e, HTTPException):
                raise e
            error_msg = str(e)
            if "credentials" in error_msg.lower():
                raise HTTPException(status_code=401, detail="Authentication required. Please visit http://localhost:8000/ to authenticate.")
            raise HTTPException(status_code=500, detail=f"Setup Structure Error: {error_msg}") 