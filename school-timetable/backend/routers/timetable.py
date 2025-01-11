from fastapi import APIRouter, HTTPException, Depends
from ..services.google_sheets import GoogleSheetsService
from ..models import SheetDataRequest, SheetDataResponse, SetupStructureResponse

router = APIRouter(prefix="/api/timetable")

@router.post("/setup-structure")
async def setup_structure(
    spreadsheet_id: str,
    sheets_service: GoogleSheetsService = Depends()
) -> SetupStructureResponse:
    # Move the setup structure logic here
    pass 