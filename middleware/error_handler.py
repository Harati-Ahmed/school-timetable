from fastapi import Request
from fastapi.responses import JSONResponse
from google.auth.exceptions import GoogleAuthError

async def google_auth_exception_handler(request: Request, exc: GoogleAuthError):
    return JSONResponse(
        status_code=401,
        content={"detail": "Google authentication failed", "message": str(exc)}
    ) 