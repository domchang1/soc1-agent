# SOC 1 Generator

Simple full-stack starter for uploading a Type II report and a blank SOC 1 Management Review sheet to generate a draft SOC 1 output.

## Backend (FastAPI)

```bash
cd backend
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

## Frontend (React + TypeScript)

```bash
cd frontend
npm install
npm run dev
```

Open `http://localhost:5173` to use the UI.
