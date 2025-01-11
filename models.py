from pydantic import BaseModel
from typing import List, Optional

class SheetDataRequest(BaseModel):
    range: str
    values: List[List[str]]

class SheetDataResponse(BaseModel):
    values: List[List[str]]

class SetupStructureResponse(BaseModel):
    message: str 