# School Timetable Management System

A web-based system for managing school timetables using Google Sheets integration.

## Features

- Google Sheets integration for timetable management
- FastAPI backend with Python
- Environment variable configuration
- CORS support
- OAuth2 authentication with Google

## Project Structure

```
school-timetable/
├── backend/
│   ├── main.py
│   ├── requirements.txt
│   ├── Procfile
│   ├── runtime.txt
│   └── .env.example
└── frontend/
    ├── static/
    └── templates/
```

## Setup Instructions

1. Clone the repository:
```bash
git clone https://github.com/your-username/school-timetable.git
cd school-timetable
```

2. Set up the backend:
```bash
cd backend
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

3. Configure environment variables:
```bash
cp .env.example .env
# Edit .env with your configuration
```

4. Run the development server:
```bash
uvicorn main:app --reload
```

## Deployment

### Backend (Render)
- Deploy using render.yaml configuration
- Set up environment variables in Render dashboard

### Frontend (GitHub Pages)
- Configure GitHub Pages in repository settings
- Update API endpoint in frontend code

## Environment Variables

See `.env.example` for required environment variables. 