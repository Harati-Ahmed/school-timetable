from fastapi import FastAPI, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os
import pickle
from typing import List, Optional, Any
from pydantic import BaseModel, BaseSettings

# Constants and Configuration
API_HOST = "localhost"  # FastAPI host
API_PORT = 8000  # FastAPI runs on 8000
OAUTH_HOST = "localhost"
OAUTH_PORT = 8001  # OAuth flow runs on 8001

# OAuth Configuration
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), 'config', 'credentials.json')
TOKEN_FILE = os.path.join(os.path.dirname(__file__), 'config', 'token.pickle')

# Column constants for periods
TEACHERS_FIRST_PERIOD = 4    # Column D (1st period)
TEACHERS_LAST_PERIOD = 14    # Column N (9th period)
TEACHERS_BREAK_COL = 7       # Column G (Break)
TEACHERS_LUNCH_COL = 11      # Column K (Lunch)

CLASSES_FIRST_PERIOD = 3     # Column C (1st period)
CLASSES_LAST_PERIOD = 13     # Column M (9th period)
CLASSES_BREAK_COL = 6        # Column F (Break)
CLASSES_LUNCH_COL = 10       # Column J (Lunch)

# Sheet name constants
TEACHERS_SHEET_NAME = 'Teachers'
CLASSES_SHEET_NAME = 'Classes'
CONFIG_TEACHERS_NAME = 'Config_Teachers'
CONFIG_CLASSES_NAME = 'Config_Classes'
CONFIG_SUBJECTS_NAME = 'Config_Subjects'
SUMMARY_SHEET_NAME = 'Summary'

# Column mapping between Teachers and Classes sheets
TEACHERS_TO_CLASSES_COLS = {
    4: 3,   # Period 1: Teachers D -> Classes C
    5: 4,   # Period 2: Teachers E -> Classes D
    6: 5,   # Period 3: Teachers F -> Classes E
    8: 7,   # Period 4: Teachers H -> Classes G
    9: 8,   # Period 5: Teachers I -> Classes H
    10: 9,  # Period 6: Teachers J -> Classes I
    12: 11, # Period 7: Teachers L -> Classes K
    13: 12, # Period 8: Teachers M -> Classes L
    14: 13  # Period 9: Teachers N -> Classes M
}

# Column mapping between Classes and Teachers sheets
CLASSES_TO_TEACHERS_COLS = {v: k for k, v in TEACHERS_TO_CLASSES_COLS.items()}

class SheetDataRequest(BaseModel):
    range: str
    values: List[List[str]]

class SheetDataResponse(BaseModel):
    values: List[List[str]]

class SetupStructureResponse(BaseModel):
    message: str

class Settings(BaseSettings):
    API_HOST: str = API_HOST
    API_PORT: int = API_PORT
    OAUTH_HOST: str = OAUTH_HOST
    OAUTH_PORT: int = OAUTH_PORT
    SCOPES: List[str] = SCOPES
    CREDENTIALS_FILE: str = CREDENTIALS_FILE
    TOKEN_FILE: str = TOKEN_FILE

    class Config:
        env_file = ".env"

settings = Settings()

app = FastAPI(
    title="School Timetable Management System",
    description="API for managing school timetables using Google Sheets",
    version="1.0.0"
)

# CORS middleware configuration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static files
app.mount("/static", StaticFiles(directory="../frontend/static"), name="static")
templates = Jinja2Templates(directory="../frontend/templates")

# Add cache management
class CacheManager:
    def __init__(self, timeout=21600):  # 6 hours in seconds
        self.timeout = timeout
        self._cache = {}

    def get(self, key: str) -> Any:
        if key in self._cache:
            return self._cache[key]
        return None

    def set(self, key: str, data: Any) -> bool:
        try:
            self._cache[key] = data
            return True
        except Exception as e:
            print(f"Cache set error for key {key}: {str(e)}")
            return False

    def remove(self, key: str) -> bool:
        try:
            if key in self._cache:
                del self._cache[key]
            return True
        except Exception as e:
            print(f"Cache remove error for key {key}: {str(e)}")
            return False

    def clear(self) -> bool:
        try:
            self._cache.clear()
            return True
        except Exception as e:
            print(f"Cache clear error: {str(e)}")
            return False

cache = CacheManager()

def get_config_data(spreadsheet_id: str, force_refresh: bool = False) -> dict:
    if not force_refresh:
        cached_data = cache.get('config_data')
        if cached_data:
            return cached_data

    try:
        service = get_google_sheets_service()
        
        config_data = {
            'teachers': get_sheet_data(service, spreadsheet_id, f"{CONFIG_TEACHERS_NAME}!A2:C"),
            'classes': get_sheet_data(service, spreadsheet_id, f"{CONFIG_CLASSES_NAME}!A2:B"),
            'subjects': get_sheet_data(service, spreadsheet_id, f"{CONFIG_SUBJECTS_NAME}!A2:B")
        }

        cache.set('config_data', config_data)
        return config_data
    except Exception as e:
        print(f"Error getting config data: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Failed to get config data: {str(e)}")

def get_sheet_data(service, spreadsheet_id: str, range_name: str) -> List[List[str]]:
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()
        return result.get('values', [])
    except Exception as e:
        print(f"Error getting sheet data: {str(e)}")
        return []

def get_google_sheets_service():
    creds = None
    try:
        if os.path.exists(TOKEN_FILE):
            print(f"Found existing token file at {TOKEN_FILE}")
            with open(TOKEN_FILE, 'rb') as token:
                creds = pickle.load(token)
                print("Successfully loaded credentials from token file")
        
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                print("Refreshing expired credentials")
                creds.refresh(Request())
            else:
                print("Starting new OAuth2 flow...")
                if not os.path.exists(CREDENTIALS_FILE):
                    print(f"Credentials file not found at: {os.path.abspath(CREDENTIALS_FILE)}")
                    raise FileNotFoundError(f"Credentials file not found at {os.path.abspath(CREDENTIALS_FILE)}")
                
                print(f"Loading credentials from: {CREDENTIALS_FILE}")
                with open(CREDENTIALS_FILE, 'r') as f:
                    print(f"Credentials content: {f.read()}")
                
                os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
                
                flow = InstalledAppFlow.from_client_secrets_file(
                    CREDENTIALS_FILE, 
                    SCOPES,
                    redirect_uri="http://localhost:8000/oauth2callback"
                )
                print("Successfully created OAuth flow")
                
                auth_url, _ = flow.authorization_url(
                    access_type='offline',
                    include_granted_scopes='true'
                )
                print(f"Generated auth URL: {auth_url}")
                
                # Return the auth URL in a 401 response
                return {
                    "auth_required": True,
                    "auth_url": auth_url
                }
        
        print("Building Google Sheets service...")
        service = build('sheets', 'v4', credentials=creds)
        print("Successfully built Google Sheets service")
        return service
        
    except Exception as e:
        print(f"Error in get_google_sheets_service: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Authentication error: {str(e)}"
        )

@app.get("/")
async def read_root():
    try:
        result = get_google_sheets_service()
        
        # Check if we need authentication
        if isinstance(result, dict) and result.get("auth_required"):
            return {
                "status": "unauthorized",
                "message": "Authentication required",
                "auth_url": result["auth_url"]
            }
        
        return {
            "message": "Authentication successful! You can now use the API.",
            "status": "authenticated",
            "api_version": "1.0.0"
        }
    except Exception as e:
        print(f"Error in root endpoint: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=str(e)
        )

@app.get("/oauth2callback")
async def oauth2callback(code: str, state: Optional[str] = None):
    try:
        print(f"Received OAuth callback with code: {code[:10]}...")
        
        flow = InstalledAppFlow.from_client_secrets_file(
            CREDENTIALS_FILE,
            SCOPES,
            redirect_uri="http://localhost:8000/oauth2callback"
        )
        
        flow.fetch_token(code=code)
        creds = flow.credentials
        
        # Save the credentials
        os.makedirs(os.path.dirname(TOKEN_FILE), exist_ok=True)
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)
        
        return {"message": "Authorization successful! You can close this window and return to the application."}
    except Exception as e:
        print(f"Error in OAuth callback: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"OAuth callback error: {str(e)}"
        )

@app.get("/api/sheets/{spreadsheet_id}")
async def get_sheet_data(spreadsheet_id: str, range: str) -> SheetDataResponse:
    try:
        service = get_google_sheets_service()
            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range
            ).execute()
            return {"values": result.get('values', [])}
        except Exception as e:
            print(f"Google Sheets API Error: {str(e)}")
            raise HTTPException(status_code=500, detail=f"Google Sheets API Error: {str(e)}")

@app.post("/api/sheets/{spreadsheet_id}")
async def update_sheet_data(spreadsheet_id: str, data: SheetDataRequest) -> dict[str, Any]:
    try:
        service = get_google_sheets_service()
        
        # Parse the range to get sheet name and cell reference
        parts = data.range.split('!')
        if len(parts) != 2:
            return {"message": "Invalid range format"}
            
        sheet_name = parts[0]
        cell_ref = parts[1]
        
        # Extract column letter and row number
        col_letter = ''.join(filter(str.isalpha, cell_ref))
        row = int(''.join(filter(str.isdigit, cell_ref)))
        col = ord(col_letter.upper()) - ord('A') + 1  # Convert column letter to number (1-based)
        
        # Get the value being set
        value = data.values[0][0] if data.values and data.values[0] else ''
        
        print(f"Processing update: sheet={sheet_name}, row={row}, col={col}, value={value}")
        
        # Initialize batch update request
        batch_updates = []
        
        # Add the primary update
        batch_updates.append({
            'range': data.range,
            'values': data.values
        })
        
        # Handle syncing between Teachers and Classes sheets
        if sheet_name == TEACHERS_SHEET_NAME and col in TEACHERS_TO_CLASSES_COLS:
            # Get teacher name
            teacher_range = f"{TEACHERS_SHEET_NAME}!B{row}"
            teacher_response = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=teacher_range
            ).execute()
            
            if 'values' in teacher_response and teacher_response['values']:
                teacher_name = teacher_response['values'][0][0]
                class_name = value
                
                if class_name and class_name.strip():
                    # Find the corresponding row in Classes sheet
                    classes_range = f"{CLASSES_SHEET_NAME}!B:B"
                    classes_response = service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id,
                        range=classes_range
                    ).execute()
                    
                    if 'values' in classes_response:
                        for i, row_data in enumerate(classes_response['values']):
                            if row_data and row_data[0] == class_name:
                                classes_col = TEACHERS_TO_CLASSES_COLS[col]
                                target_range = f"{CLASSES_SHEET_NAME}!{chr(64 + classes_col)}{i + 1}"
                                
                                print(f"Syncing to Classes sheet: range={target_range}, value={teacher_name}")
                                batch_updates.append({
                                    'range': target_range,
                                    'values': [[teacher_name]]
                                })
                                break
        
        elif sheet_name == CLASSES_SHEET_NAME and col in CLASSES_TO_TEACHERS_COLS:
            # Get class name
            class_range = f"{CLASSES_SHEET_NAME}!B{row}"
            class_response = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=class_range
            ).execute()
            
            if 'values' in class_response and class_response['values']:
                class_name = class_response['values'][0][0]
                teacher_name = value
                
                if teacher_name and teacher_name.strip():
                    # Find the corresponding row in Teachers sheet
                    teachers_range = f"{TEACHERS_SHEET_NAME}!B:B"
                    teachers_response = service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id,
                        range=teachers_range
                    ).execute()
                    
                    if 'values' in teachers_response:
                        for i, row_data in enumerate(teachers_response['values']):
                            if row_data and row_data[0] == teacher_name:
                                teachers_col = CLASSES_TO_TEACHERS_COLS[col]
                                target_range = f"{TEACHERS_SHEET_NAME}!{chr(64 + teachers_col)}{i + 1}"
                                
                                print(f"Syncing to Teachers sheet: range={target_range}, value={class_name}")
                                batch_updates.append({
                                    'range': target_range,
                                    'values': [[class_name]]
                                })
                                break
        
        # Execute batch update
        if batch_updates:
            print(f"Executing batch update with {len(batch_updates)} operations")
            result = service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    'valueInputOption': 'USER_ENTERED',
                    'data': batch_updates
                }
            ).execute()
            
            # Update summary after changes
            await update_summary(spreadsheet_id)
            
            return {"updated_cells": len(batch_updates)}
        else:
            return {"message": "No updates required"}
        
        except Exception as e:
        print(f"Error in update_sheet_data: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Update Error: {str(e)}")

@app.post("/api/timetable/setup-structure")
async def setup_structure(spreadsheet_id: str) -> SetupStructureResponse:
    try:
        service = get_google_sheets_service()
        
        # Delete existing sheets but keep one
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        delete_requests = []
        
        # Keep the first sheet, delete others
        for sheet in sheets[1:]:
            sheet_id = sheet['properties']['sheetId']
            delete_requests.append({
                'deleteSheet': {
                    'sheetId': sheet_id
                }
            })
        
        if delete_requests:
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': delete_requests}
            ).execute()
        
        # Enhanced sheet configurations with time periods
        sheet_configs = [
            {
                'title': CONFIG_TEACHERS_NAME,
                'headers': ['ID', 'Teacher Name', 'Subject'],
                'column_widths': [100, 200, 150]
            },
            {
                'title': CONFIG_CLASSES_NAME,
                'headers': ['ID', 'Class Name'],
                'column_widths': [100, 200]
            },
            {
                'title': CONFIG_SUBJECTS_NAME,
                'headers': ['ID', 'Subject Name'],
                'column_widths': [100, 200]
            },
            {
                'title': TEACHERS_SHEET_NAME,
                'headers': [
                    'SI', 'Teacher Name', 'Subject',
                    '1\n08:00-08:50', '2\n08:50-09:30', '3\n09:30-10:10',
                    'Break\n10:10-10:30',
                    '4\n10:30-11:10', '5\n11:10-11:50', '6\n11:50-12:30',
                    'Lunch\n12:30-01:00',
                    '7\n01:00-01:40', '8\n01:40-02:20', '9\n02:20-03:00'
                ],
                'column_widths': [50, 200, 150, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100]
            },
            {
                'title': CLASSES_SHEET_NAME,
                'headers': [
                    'SI', 'Class Name',
                    '1\n08:00-08:50', '2\n08:50-09:30', '3\n09:30-10:10',
                    'Break\n10:10-10:30',
                    '4\n10:30-11:10', '5\n11:10-11:50', '6\n11:50-12:30',
                    'Lunch\n12:30-01:00',
                    '7\n01:00-01:40', '8\n01:40-02:20', '9\n02:20-03:00'
                ],
                'column_widths': [50, 200, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100]
            },
            {
                'title': SUMMARY_SHEET_NAME,
                'headers': ['Teacher Name', 'Subject', 'Total Hours'],
                'column_widths': [200, 150, 100]
            }
        ]

        # Create sheets and apply formatting
        requests = []
        for config in sheet_configs:
            sheet_id = create_sheet(service, spreadsheet_id, config['title'])
            
            # Add headers
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{config['title']}!A1",
                valueInputOption='RAW',
                body={'values': [config['headers']]}
            ).execute()

            # Format headers with wrapping
            requests.extend([
                {
                    'repeatCell': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': 0,
                            'endRowIndex': 1,
                        },
                        'cell': {
                            'userEnteredFormat': {
                                'backgroundColor': {'red': 0.95, 'green': 0.95, 'blue': 0.95},
                                'textFormat': {'bold': True},
                                'horizontalAlignment': 'CENTER',
                                'verticalAlignment': 'MIDDLE',
                                'wrapStrategy': 'WRAP'
                            }
                        },
                        'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
                    }
                }
            ])

            # Set row height for header to accommodate wrapped text
            requests.append({
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': sheet_id,
                        'dimension': 'ROWS',
                        'startIndex': 0,
                        'endIndex': 1
                    },
                    'properties': {
                        'pixelSize': 60  # Increased height for wrapped text
                    },
                    'fields': 'pixelSize'
                }
            })

            # Set column widths
            for i, width in enumerate(config['column_widths']):
                requests.append({
                    'updateDimensionProperties': {
                        'range': {
                            'sheetId': sheet_id,
                            'dimension': 'COLUMNS',
                            'startIndex': i,
                            'endIndex': i + 1
                        },
                        'properties': {
                            'pixelSize': width
                        },
                        'fields': 'pixelSize'
                    }
                })

            # Add pink background for break and lunch columns
            if config['title'] in [TEACHERS_SHEET_NAME, CLASSES_SHEET_NAME]:
                break_col = TEACHERS_BREAK_COL if config['title'] == TEACHERS_SHEET_NAME else CLASSES_BREAK_COL
                lunch_col = TEACHERS_LUNCH_COL if config['title'] == TEACHERS_SHEET_NAME else CLASSES_LUNCH_COL
                
                requests.extend([
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": 0,
                                "endRowIndex": 100,
                                "startColumnIndex": break_col - 1,
                                "endColumnIndex": break_col
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "backgroundColor": {"red": 1.0, "green": 0.9, "blue": 0.9}
                                }
                            },
                            "fields": "userEnteredFormat(backgroundColor)"
                        }
                    },
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": 0,
                                "endRowIndex": 100,
                                "startColumnIndex": lunch_col - 1,
                                "endColumnIndex": lunch_col
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "backgroundColor": {"red": 1.0, "green": 0.9, "blue": 0.9}
                                }
                            },
                            "fields": "userEnteredFormat(backgroundColor)"
                        }
                    }
                ])

            # Add borders to all cells
            requests.append({
                "updateBorders": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 100,
                        "startColumnIndex": 0,
                        "endColumnIndex": len(config['headers'])
                    },
                    "top": {"style": "SOLID"},
                    "bottom": {"style": "SOLID"},
                    "left": {"style": "SOLID"},
                    "right": {"style": "SOLID"},
                    "innerHorizontal": {"style": "SOLID"},
                    "innerVertical": {"style": "SOLID"}
                }
            })

        # Apply all formatting
        if requests:
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': requests}
            ).execute()

        # Clear cache after setup
        cache.clear()

        return {"message": "Structure setup completed successfully"}
    except Exception as e:
        print(f"Error in setup_structure: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Setup Structure Error: {str(e)}"
        )

def create_sheet(service, spreadsheet_id: str, title: str) -> int:
    try:
            result = service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
            body={
                'requests': [{
                    'addSheet': {
                        'properties': {
                            'title': title
                        }
                    }
                }]
            }
            ).execute()
        return result['replies'][0]['addSheet']['properties']['sheetId']
        except Exception as e:
        # If sheet already exists, get its ID
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        for sheet in spreadsheet['sheets']:
            if sheet['properties']['title'] == title:
                return sheet['properties']['sheetId']
        raise e

@app.post("/api/timetable/deploy-config")
async def deploy_config(spreadsheet_id: str) -> dict[str, str]:
    try:
        service = get_google_sheets_service()
        
        # Get config data
        teachers_data = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{CONFIG_TEACHERS_NAME}!A2:C"
        ).execute().get('values', [])
        
        classes_data = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{CONFIG_CLASSES_NAME}!A2:B"
        ).execute().get('values', [])
        
        # Prepare data for Teachers sheet
        teachers_rows = []
        for i, teacher in enumerate(teachers_data, start=1):
            teachers_rows.append([str(i)] + teacher[1:])  # Add SI number and use name and subject
        
        # Prepare data for Classes sheet
        classes_rows = []
        for i, class_info in enumerate(classes_data, start=1):
            classes_rows.append([str(i)] + [class_info[1]])  # Add SI number and use only class name
        
        # Deploy to Teachers sheet
        if teachers_rows:
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{TEACHERS_SHEET_NAME}!A2",
                valueInputOption='USER_ENTERED',
                body={'values': teachers_rows}
            ).execute()
        
        # Deploy to Classes sheet
        if classes_rows:
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{CLASSES_SHEET_NAME}!A2",
                valueInputOption='USER_ENTERED',
                body={'values': classes_rows}
            ).execute()
        
        return {"message": "Configuration deployed successfully"}
    except Exception as e:
        print(f"Error in deploy_config: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/timetable/update-summary")
async def update_summary(spreadsheet_id: str) -> dict[str, str]:
    try:
        service = get_google_sheets_service()
        
        # Get teachers data including all periods
        teachers_data = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{TEACHERS_SHEET_NAME}!A2:N"
        ).execute().get('values', [])
        
        # Calculate summary with more details
        summary_data = []
        for row in teachers_data:
            if len(row) >= 3:  # Has at least SI, teacher name, and subject
                teacher_name = row[1]
                subject = row[2]
                
                # Count classes per period (excluding breaks and lunch)
                period_counts = []
                for cell in row[3:]:
                    if cell and cell.strip() and cell.lower() not in ['break', 'lunch']:
                        period_counts.append(cell)
                
                total_hours = len(period_counts)
                class_list = ', '.join(period_counts) if period_counts else 'No classes'
                
                summary_data.append([
                    teacher_name,
                    subject,
                    str(total_hours),
                    class_list
                ])
        
        # Update summary sheet with headers if needed
        headers = ['Teacher Name', 'Subject', 'Total Hours', 'Assigned Classes']
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{SUMMARY_SHEET_NAME}!A1",
            valueInputOption='RAW',
            body={'values': [headers]}
        ).execute()
        
        # Update summary data
        if summary_data:
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{SUMMARY_SHEET_NAME}!A2",
                valueInputOption='RAW',
                body={'values': summary_data}
            ).execute()
            
            # Format summary sheet
            sheet_id = None
            spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            for sheet in spreadsheet['sheets']:
                if sheet['properties']['title'] == SUMMARY_SHEET_NAME:
                    sheet_id = sheet['properties']['sheetId']
                    break
            
            if sheet_id:
                requests = [
                    {
                        'updateDimensionProperties': {
                            'range': {
                                'sheetId': sheet_id,
                                'dimension': 'COLUMNS',
                                'startIndex': 0,
                                'endIndex': 4
                            },
                            'properties': {
                                'pixelSize': 200  # Width for all columns
                            },
                            'fields': 'pixelSize'
                        }
                    },
                    {
                        'repeatCell': {
                            'range': {
                                'sheetId': sheet_id,
                                'startRowIndex': 0,
                                'endRowIndex': 1
                            },
                            'cell': {
                                'userEnteredFormat': {
                                    'backgroundColor': {'red': 0.95, 'green': 0.95, 'blue': 0.95},
                                    'textFormat': {'bold': True},
                                    'horizontalAlignment': 'CENTER'
                                }
                            },
                            'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
                        }
                    }
                ]
                
                service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={'requests': requests}
                ).execute()
        
        return {"message": "Summary updated successfully"}
    except Exception as e:
        print(f"Error updating summary: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/timetable/clear-all-data")
async def clear_all_data(spreadsheet_id: str) -> dict[str, str]:
    try:
        service = get_google_sheets_service()
        
        sheets_to_clear = [
            TEACHERS_SHEET_NAME,
            CLASSES_SHEET_NAME,
            CONFIG_TEACHERS_NAME,
            CONFIG_CLASSES_NAME,
            CONFIG_SUBJECTS_NAME,
            SUMMARY_SHEET_NAME
        ]
        
        for sheet_name in sheets_to_clear:
            try:
                service.spreadsheets().values().clear(
                    spreadsheetId=spreadsheet_id,
                    range=f"{sheet_name}!A2:Z1000"
                ).execute()
            except Exception as sheet_error:
                print(f"Error clearing sheet {sheet_name}: {str(sheet_error)}")
                continue
        
        # Clear cache
        cache.clear()
        
        return {"message": "All data cleared successfully"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/config/add-teacher")
async def add_teacher(spreadsheet_id: str, teacher_data: dict) -> dict[str, str]:
    try:
        service = get_google_sheets_service()
        
        # Get existing teachers
        existing_data = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{CONFIG_TEACHERS_NAME}!A:C"
        ).execute().get('values', [])
        
        # Generate new ID
        new_id = f"T{len(existing_data):03d}"
        
        # Add new teacher
        new_row = [new_id, teacher_data['name'], teacher_data['subject']]
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{CONFIG_TEACHERS_NAME}!A:C",
            valueInputOption='RAW',
            body={'values': [new_row]}
        ).execute()
        
        # Clear cache
        cache.remove('config_data')
        
        return {"message": f"Teacher {teacher_data['name']} added successfully", "id": new_id}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/config/add-class")
async def add_class(spreadsheet_id: str, class_data: dict) -> dict[str, str]:
    try:
        service = get_google_sheets_service()
        
        # Get existing classes
        existing_data = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{CONFIG_CLASSES_NAME}!A:B"
        ).execute().get('values', [])
        
        # Generate new ID
        new_id = f"C{len(existing_data):03d}"
        
        # Add new class (without section)
        new_row = [new_id, class_data['name']]
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{CONFIG_CLASSES_NAME}!A:B",
            valueInputOption='RAW',
            body={'values': [new_row]}
        ).execute()
        
        # Clear cache
        cache.remove('config_data')
        
        return {"message": f"Class {class_data['name']} added successfully", "id": new_id}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/config/add-subject")
async def add_subject(spreadsheet_id: str, subject_data: dict) -> dict[str, str]:
    try:
        service = get_google_sheets_service()
        
        # Get existing subjects
        existing_data = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{CONFIG_SUBJECTS_NAME}!A:B"
        ).execute().get('values', [])
        
        # Generate new ID
        new_id = f"S{len(existing_data):03d}"
        
        # Add new subject
        new_row = [new_id, subject_data['name']]
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{CONFIG_SUBJECTS_NAME}!A:B",
            valueInputOption='RAW',
            body={'values': [new_row]}
        ).execute()
        
        # Clear cache
        cache.remove('config_data')
        
        return {"message": f"Subject {subject_data['name']} added successfully", "id": new_id}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/timetable/format-structure")
async def format_structure(spreadsheet_id: str) -> dict[str, str]:
    try:
        service = get_google_sheets_service()
        
        # Get sheet IDs
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        sheet_ids = {}
        
        for sheet in sheets:
            title = sheet['properties']['title']
            sheet_id = sheet['properties']['sheetId']
            sheet_ids[title] = sheet_id
        
        requests = []
        
        # Format Teachers sheet
        if TEACHERS_SHEET_NAME in sheet_ids:
            requests.extend([
                # Format headers
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_ids[TEACHERS_SHEET_NAME],
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                                "textFormat": {"bold": True},
                                "horizontalAlignment": "CENTER",
                                "verticalAlignment": "MIDDLE",
                                "wrapStrategy": "WRAP"
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)"
                    }
                },
                # Format subject text in red
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_ids[TEACHERS_SHEET_NAME],
                            "startRowIndex": 1,
                            "endRowIndex": 100,
                            "startColumnIndex": 2,  # Subject column
                            "endColumnIndex": 3
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": {
                                    "foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}
                                }
                            }
                        },
                        "fields": "userEnteredFormat.textFormat.foregroundColor"
                    }
                },
                # Format break column with pink background
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_ids[TEACHERS_SHEET_NAME],
                            "startRowIndex": 0,
                            "endRowIndex": 100,
                            "startColumnIndex": TEACHERS_BREAK_COL - 1,
                            "endColumnIndex": TEACHERS_BREAK_COL
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 1.0, "green": 0.9, "blue": 0.9}
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor)"
                    }
                },
                # Format lunch column with pink background
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_ids[TEACHERS_SHEET_NAME],
                            "startRowIndex": 0,
                            "endRowIndex": 100,
                            "startColumnIndex": TEACHERS_LUNCH_COL - 1,
                            "endColumnIndex": TEACHERS_LUNCH_COL
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 1.0, "green": 0.9, "blue": 0.9}
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor)"
                    }
                },
                # Add borders to all cells
                {
                    "updateBorders": {
                        "range": {
                            "sheetId": sheet_ids[TEACHERS_SHEET_NAME],
                            "startRowIndex": 0,
                            "endRowIndex": 100,
                            "startColumnIndex": 0,
                            "endColumnIndex": TEACHERS_LAST_PERIOD
                        },
                        "top": {"style": "SOLID"},
                        "bottom": {"style": "SOLID"},
                        "left": {"style": "SOLID"},
                        "right": {"style": "SOLID"},
                        "innerHorizontal": {"style": "SOLID"},
                        "innerVertical": {"style": "SOLID"}
                    }
                },
                # Center align all cells
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_ids[TEACHERS_SHEET_NAME],
                            "startRowIndex": 1,
                            "endRowIndex": 100,
                            "startColumnIndex": 0,
                            "endColumnIndex": TEACHERS_LAST_PERIOD
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "horizontalAlignment": "LEFT",
                                "verticalAlignment": "MIDDLE"
                            }
                        },
                        "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
                    }
                }
            ])
        
        # Apply all formatting
        if requests:
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': requests}
            ).execute()
        
        return {"message": "Formatting applied successfully"}
        
    except Exception as e:
        print(f"Error in format_structure: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Format Structure Error: {str(e)}"
        )

@app.post("/api/timetable/setup-dropdowns")
async def setup_dropdowns(spreadsheet_id: str) -> dict[str, str]:
    try:
        service = get_google_sheets_service()
        
        # Get sheet IDs
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        sheet_ids = {}
        
        for sheet in sheets:
            title = sheet['properties']['title']
            sheet_id = sheet['properties']['sheetId']
            sheet_ids[title] = sheet_id
        
        requests = []
        
        # Set up dropdowns for Teachers sheet periods
        if TEACHERS_SHEET_NAME in sheet_ids:
            # Add dropdowns for period cells (excluding break and lunch)
            period_columns = [(3,4), (4,5), (5,6), (7,8), (8,9), (9,10), (11,12), (12,13), (13,14)]  # Columns for periods 1-9
            for start_col, end_col in period_columns:
                requests.append({
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet_ids[TEACHERS_SHEET_NAME],
                            "startRowIndex": 1,
                            "endRowIndex": 100,
                            "startColumnIndex": start_col,
                            "endColumnIndex": end_col
                        },
                        "rule": {
                            "condition": {
                                "type": "ONE_OF_RANGE",
                                "values": [{
                                    "userEnteredValue": f"={CONFIG_CLASSES_NAME}!B2:B"
                                }]
                            },
                            "showCustomUi": True,
                            "strict": True
                        }
                    }
                })
        
        # Set up dropdowns for Classes sheet periods
        if CLASSES_SHEET_NAME in sheet_ids:
            # Add dropdowns for period cells (excluding break and lunch)
            period_columns = [(2,3), (3,4), (4,5), (6,7), (7,8), (8,9), (10,11), (11,12), (12,13)]  # Columns for periods 1-9
            for start_col, end_col in period_columns:
                requests.append({
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet_ids[CLASSES_SHEET_NAME],
                            "startRowIndex": 1,
                            "endRowIndex": 100,
                            "startColumnIndex": start_col,
                            "endColumnIndex": end_col
                        },
                        "rule": {
                            "condition": {
                                "type": "ONE_OF_RANGE",
                                "values": [{
                                    "userEnteredValue": f"={CONFIG_TEACHERS_NAME}!B2:B"
                                }]
                            },
                            "showCustomUi": True,
                            "strict": True
                        }
                    }
                })
        
        # Set up subject dropdown in Config_Teachers sheet
        if CONFIG_TEACHERS_NAME in sheet_ids:
            requests.append({
                "setDataValidation": {
                    "range": {
                        "sheetId": sheet_ids[CONFIG_TEACHERS_NAME],
                        "startRowIndex": 1,
                        "endRowIndex": 100,
                        "startColumnIndex": 2,  # Column C (Subject)
                        "endColumnIndex": 3
                    },
                    "rule": {
                        "condition": {
                            "type": "ONE_OF_RANGE",
                            "values": [{
                                "userEnteredValue": f"={CONFIG_SUBJECTS_NAME}!B2:B"
                            }]
                        },
                        "showCustomUi": True,
                        "strict": True
                    }
                }
            })
        
        # Apply all validation rules
        if requests:
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': requests}
            ).execute()
        
        return {"message": "Dropdowns set up successfully"}
        
    except Exception as e:
        print(f"Error in setup_dropdowns: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Setup Dropdowns Error: {str(e)}"
        )

@app.post("/api/timetable/on-edit")
async def on_edit(spreadsheet_id: str, edit_data: dict) -> dict[str, str]:
    """Handle changes in the sheets and update dropdowns accordingly"""
    try:
        service = get_google_sheets_service()
        
        sheet_name = edit_data.get('sheet_name')
        row = edit_data.get('row', 0)
        col = edit_data.get('col', 0)
        value = edit_data.get('value', '')
        
        print(f"Processing edit: sheet={sheet_name}, row={row}, col={col}, value={value}")
        
        # Skip if not editing a period cell
        if sheet_name == TEACHERS_SHEET_NAME and (col < 4 or col > 14 or col in [7, 11]):  # Skip non-period columns
            return {"message": "Not a period cell"}
        elif sheet_name == CLASSES_SHEET_NAME and (col < 3 or col > 13 or col in [6, 10]):  # Skip non-period columns
            return {"message": "Not a period cell"}
        
        if sheet_name == TEACHERS_SHEET_NAME:
            # Get teacher name
            teacher_data = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"{TEACHERS_SHEET_NAME}!B{row}"
            ).execute().get('values', [[]])[0]
            
            if not teacher_data:
                print("No teacher data found")
                return {"message": "No teacher found"}
            
            teacher_name = teacher_data[0]
            class_name = value
            
            print(f"Teacher {teacher_name} assigned to class {class_name}")
            
            if class_name:
                # Calculate corresponding column in Classes sheet
                classes_col = col - 1  # Teachers sheet has one extra column (Subject)
                
                # Find the class row in Classes sheet
                class_data = service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=f"{CLASSES_SHEET_NAME}!B2:B100"
                ).execute().get('values', [])
                
                for i, row_data in enumerate(class_data, start=2):
                    if row_data and row_data[0] == class_name:
                        # Update the corresponding period in Classes sheet
                        target_range = f"{CLASSES_SHEET_NAME}!{chr(64+classes_col)}{i}"
                        print(f"Updating Classes sheet at {target_range} with {teacher_name}")
                        
                        try:
                            service.spreadsheets().values().update(
                                spreadsheetId=spreadsheet_id,
                                range=target_range,
                                valueInputOption='USER_ENTERED',
                                body={'values': [[teacher_name]]}
                            ).execute()
                            print(f"Successfully updated Classes sheet")
                        except Exception as e:
                            print(f"Error updating Classes sheet: {str(e)}")
                        break
        
        elif sheet_name == CLASSES_SHEET_NAME:
            # Get class name
            class_data = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"{CLASSES_SHEET_NAME}!B{row}"
            ).execute().get('values', [[]])[0]
            
            if not class_data:
                print("No class data found")
                return {"message": "No class found"}
            
            class_name = class_data[0]
            teacher_name = value
            
            print(f"Class {class_name} assigned to teacher {teacher_name}")
            
            if teacher_name:
                # Calculate corresponding column in Teachers sheet
                teachers_col = col + 1  # Teachers sheet has one extra column (Subject)
                
                # Find the teacher row in Teachers sheet
                teacher_data = service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=f"{TEACHERS_SHEET_NAME}!B2:B100"
                ).execute().get('values', [])
                
                for i, row_data in enumerate(teacher_data, start=2):
                    if row_data and row_data[0] == teacher_name:
                        # Update the corresponding period in Teachers sheet
                        target_range = f"{TEACHERS_SHEET_NAME}!{chr(64+teachers_col)}{i}"
                        print(f"Updating Teachers sheet at {target_range} with {class_name}")
                        
                        try:
                            service.spreadsheets().values().update(
                                spreadsheetId=spreadsheet_id,
                                range=target_range,
                                valueInputOption='USER_ENTERED',
                                body={'values': [[class_name]]}
                            ).execute()
                            print(f"Successfully updated Teachers sheet")
                        except Exception as e:
                            print(f"Error updating Teachers sheet: {str(e)}")
                        break
        
        # Update summary after changes
        await update_summary(spreadsheet_id)
        
        return {"message": "Changes synchronized successfully"}
        
    except Exception as e:
        print(f"Error in on_edit: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Edit Sync Error: {str(e)}"
        )

@app.post("/api/timetable/batch-deploy")
async def batch_deploy(spreadsheet_id: str) -> dict[str, str]:
    try:
        service = get_google_sheets_service()
        
        # Define all data
        subjects = [
            ["S001", "English"],
            ["S002", "Maths"],
            ["S003", "Science"],
            ["S004", "Hindi"],
            ["S005", "SST"],
            ["S006", "Sanskrit"]
        ]
        
        classes = [
            ["C001", "Nursery"],
            ["C002", "LKG - A"],
            ["C003", "LKG - B"],
            ["C004", "UKG - A"],
            ["C005", "UKG - B"],
            ["C006", "Grade - 1A"],
            ["C007", "Grade - 1B"],
            ["C008", "Grade - 2A"],
            ["C009", "Grade - 2B"],
            ["C010", "Grade - 3A"],
            ["C011", "Grade - 3B"],
            ["C012", "Grade - 4A"],
            ["C013", "Grade - 4B"],
            ["C014", "Grade - 5"],
            ["C015", "Grade - 6"],
            ["C016", "Grade - 7"],
            ["C017", "Grade - 8"],
            ["C018", "Grade - 9"],
            ["C019", "Grade - 10"],
            ["C020", "Grade - 11"],
            ["C021", "Grade - 12"]
        ]
        
        teachers = [
            ["T001", "Raju Bumb", "English"],
            ["T002", "Prabhat Karan", "Maths"],
            ["T003", "Shobha Hans", "Science"],
            ["T004", "Krishna Naidu", "Hindi"],
            ["T005", "Faraz Mangal", "SST"],
            ["T006", "Rimi Loke", "Sanskrit"],
            ["T007", "Amir Kar", "English"],
            ["T008", "Suraj Narayanan", "Maths"],
            ["T009", "Alaknanda Chaudry", "Science"],
            ["T010", "Preet Mittal", "English"],
            ["T011", "John Lalla", "English"],
            ["T012", "Ujwal Mohan", "Maths"],
            ["T013", "Aadish Mathur", "Science"],
            ["T014", "Iqbal Beharry", "Hindi"],
            ["T015", "Manjari Shenoy", "SST"],
            ["T016", "Aayushi Suri", "Sanskrit"],
            ["T017", "Parvez Mathur", "SST"],
            ["T018", "Qabool Malhotra", "Hindi"],
            ["T019", "Nagma Andra", "Sanskrit"],
            ["T020", "Krishna Arora", "Hindi"],
            ["T021", "John Lalla", "SST"],
            ["T022", "Nitin Banu", "Sanskrit"],
            ["T023", "Ananda Debnath", "Hindi"],
            ["T024", "Balaram Bhandari", "Hindi"],
            ["T025", "Ajay Chaudhri", "SST"],
            ["T026", "Niranjan Varma", "English"],
            ["T027", "Nur Patel", "Maths"],
            ["T028", "Aadish Mathur", "English"],
            ["T029", "Nur Patel", "Hindi"],
            ["T030", "John Lalla", "English"],
            ["T031", "Aadish Mathur", "SST"]
        ]
        
        # Clear all existing data
        await clear_all_data(spreadsheet_id)
        
        # Add all data in batch
        batch_requests = [
            {
                'range': f"{CONFIG_SUBJECTS_NAME}!A2:B{len(subjects)+1}",
                'values': subjects
            },
            {
                'range': f"{CONFIG_CLASSES_NAME}!A2:B{len(classes)+1}",
                'values': classes
            },
            {
                'range': f"{CONFIG_TEACHERS_NAME}!A2:C{len(teachers)+1}",
                'values': teachers
            }
        ]
        
        for request in batch_requests:
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=request['range'],
                valueInputOption='RAW',
                body={'values': request['values']}
            ).execute()
        
        # Deploy to main sheets
        await deploy_config(spreadsheet_id)
        
        # Set up dropdowns
        await setup_dropdowns(spreadsheet_id)
        
        # Apply formatting
        await format_structure(spreadsheet_id)
        
        # Sync sheets
        await sync_sheets(spreadsheet_id)
        
        return {"message": "All data deployed successfully"}
        
    except Exception as e:
        print(f"Error in batch_deploy: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Batch Deploy Error: {str(e)}"
        )

@app.post("/api/timetable/sync-sheets")
async def sync_sheets(spreadsheet_id: str) -> dict[str, str]:
    try:
        service = get_google_sheets_service()
        
        # Get data from config sheets
        teachers_config = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{CONFIG_TEACHERS_NAME}!A2:C"
        ).execute().get('values', [])
        
        classes_config = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{CONFIG_CLASSES_NAME}!A2:B"
        ).execute().get('values', [])
        
        # Get data from main sheets
        teachers_data = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{TEACHERS_SHEET_NAME}!A2:N"
        ).execute().get('values', [])
        
        classes_data = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{CLASSES_SHEET_NAME}!A2:M"
        ).execute().get('values', [])
        
        # Update Teachers sheet with latest config data
        teachers_rows = []
        for i, teacher in enumerate(teachers_config, start=1):
            # If teacher exists in main sheet, preserve their schedule
            existing_schedule = []
            for row in teachers_data:
                if len(row) >= 2 and row[1] == teacher[1]:  # Match by teacher name
                    existing_schedule = row[3:] if len(row) > 3 else []
                    break
            
            # Create new row with updated teacher info and existing schedule
            new_row = [str(i), teacher[1], teacher[2]] + existing_schedule
            teachers_rows.append(new_row)
        
        # Update Classes sheet with latest config data
        classes_rows = []
        for i, class_info in enumerate(classes_config, start=1):
            # If class exists in main sheet, preserve their schedule
            existing_schedule = []
            for row in classes_data:
                if len(row) >= 2 and row[1] == class_info[1]:  # Match by class name
                    existing_schedule = row[2:] if len(row) > 2 else []
                    break
            
            # Create new row with updated class info and existing schedule
            new_row = [str(i), class_info[1]] + existing_schedule
            classes_rows.append(new_row)
        
        # Deploy updated data to main sheets
        if teachers_rows:
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{TEACHERS_SHEET_NAME}!A2",
                valueInputOption='USER_ENTERED',
                body={'values': teachers_rows}
            ).execute()
        
        if classes_rows:
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{CLASSES_SHEET_NAME}!A2",
                valueInputOption='USER_ENTERED',
                body={'values': classes_rows}
            ).execute()
        
        # Update summary
        await update_summary(spreadsheet_id)
        
        return {"message": "Sheets synchronized successfully"}
        
    except Exception as e:
        print(f"Error in sync_sheets: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Sync Error: {str(e)}"
        )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host=API_HOST, port=API_PORT)
