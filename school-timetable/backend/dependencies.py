from fastapi import Depends
from .services.google_sheets import GoogleSheetsService

async def get_sheets_service():
    service = GoogleSheetsService()
    yield service 