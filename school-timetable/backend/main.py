from fastapi import FastAPI, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
from pydantic_settings import BaseSettings
from typing import List
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

class Settings(BaseSettings):
    api_host: str = os.getenv("API_HOST", "0.0.0.0")
    port: int = int(os.getenv("PORT", "8000"))
    environment: str = os.getenv("ENVIRONMENT", "development")
    google_client_id: str = os.getenv("GOOGLE_CLIENT_ID")
    google_project_id: str = os.getenv("GOOGLE_PROJECT_ID")
    google_client_secret: str = os.getenv("GOOGLE_CLIENT_SECRET")
    spreadsheet_id: str = os.getenv("SPREADSHEET_ID")
    oauth_redirect_uri: str = os.getenv("OAUTH_REDIRECT_URI", "http://localhost:8000/oauth2callback")
    allowed_origins: List[str] = os.getenv("ALLOWED_ORIGINS", "http://localhost:3000").split(",")

    class Config:
        env_file = ".env"

settings = Settings()

app = FastAPI(
    title="School Timetable API",
    description="API for managing school timetables using Google Sheets",
    version="1.0.0"
)

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.allowed_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ... existing code ...

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host=settings.api_host,
        port=settings.port,
        reload=settings.environment == "development"
    )
