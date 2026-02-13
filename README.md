# SOC 1 Type II Management Review Generator

AI-powered tool for processing SOC1 Type II audit reports and automatically filling management review Excel templates. Find [here](https://soc1-agent.vercel.app/)

## Features

- **PDF Extraction**: Uses `pdfplumber` to extract text and tables from SOC1 Type II PDF reports
- **AI-Powered Mapping**: Uses Google Gemini AI (free tier) to intelligently map extracted content to Excel template fields
- **Excel Generation**: Automatically fills management review templates with extracted data
- **Gap Analysis**: Analyzes the extraction for completeness and provides recommendations
- **User Feedback System**: Collects user feedback to help you continuously improve extraction quality

## Prerequisites

- Python 3.10+
- Node.js 18+
- Google API key (https://aistudio.google.com/apikey)

## Setup

### Backend (FastAPI)

```bash
cd backend
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -r requirements.txt

# Set your Google API key (free) in the .env file
GOOGLE_API_KEY="your-google-api-key"

# Run the server
uvicorn main:app --reload --port 8000
```

### Frontend (React + TypeScript)

```bash
cd frontend
npm install
npm run dev
```

Open `http://localhost:5173` to use the UI.

## Usage

1. **Upload Files**: 
   - Upload a SOC1 Type II report (PDF format)
   - Upload a blank management review template (Excel .xlsx format)

2. **Processing**:
   - The system extracts text and tables from the PDF
   - AI analyzes the content to identify controls, test results, and findings
   - The Excel template is populated with the extracted information

3. **Download**:
   - Once processing completes, download the filled management review
   - Review the analysis summary for key findings and recommendations

## API Endpoints

- `GET /api/health` - Health check endpoint
- `POST /api/upload` - Upload files and start processing
  - Accepts: `type_ii_report` (PDF), `management_review` (Excel)
  - Returns: `job_id` for status polling
- `GET /api/status/{job_id}` - Check processing status and get analysis summary
- `GET /api/download/{job_id}` - Download the filled Excel file
- `POST /api/feedback/{job_id}` - Submit user feedback on extraction quality
- `GET /api/feedback/stats` - Get aggregated feedback statistics (admin)
- `POST /api/cleanup-uploads` - Clear temporary upload files (maintenance)
- `POST /api/cleanup-old-files` - Remove output files older than 24 hours (maintenance)

## Architecture

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   Frontend      │────▶│   FastAPI       │────▶│   Gemini AI     │
│   (React)       │     │   Backend       │     │   (Free Tier)   │
└─────────────────┘     └─────────────────┘     └─────────────────┘
                              │                        │
                              ▼                        ▼
                       ┌─────────────────┐     ┌─────────────────┐
                       │   pdfplumber    │     │   openpyxl      │
                       │   (PDF Extract) │     │   (read-only)   │
                       └─────────────────┘     └─────────────────┘
                                                       │
                                               ┌───────┴───────┐
                                               │   xlsxwriter  │
                                               │  (Excel Write) │
                                               └───────────────┘
```

### Memory-Optimized Excel Pipeline

The Excel read/write pipeline is split into two libraries to stay well under a 2 GB container limit:

| Phase | Library | Mode | Memory |
|-------|---------|------|--------|
| **Read template** | openpyxl | `read_only=True` (streaming) | ~20 MB |
| **Parse layout** | stdlib `zipfile` + `ElementTree.iterparse` | Streaming XML | ~5 MB |
| **Write output** | xlsxwriter | Forward-only rows | ~10 MB |

Only the target tabs (`1.0` and `2.0.b`) are read from the template. Column widths and merged cells are parsed directly from the XLSX ZIP with streaming XML, so the full workbook object model is never loaded into memory.

## Dependencies

### Backend
- `fastapi` - Web framework
- `uvicorn` - ASGI server
- `pdfplumber` - PDF text and table extraction
- `openpyxl` - Excel template reading (read-only streaming)
- `xlsxwriter` - Excel output writing (memory-efficient)
- `google-genai` - Google Gemini AI client
- `python-dotenv` - Environment variable management

### Frontend
- React 18
- TypeScript
- Vite

## Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `GOOGLE_API_KEY` | Google AI API key (free tier available) | Yes |
| `VITE_API_URL` | Backend API URL (frontend only) | No* |

*If not set, defaults to `https://soc1-management-review-generator.onrender.com/api`

Get your free Google API key at: https://aistudio.google.com/apikey

## Deployment

### Backend Deployment (Render)

1. Push your code to GitHub
2. Go to [render.com](https://render.com) and sign in with GitHub
3. Create a new Web Service:
   - Select your repository
   - Set **Root Directory**: `backend`
   - **Runtime**: Python 3.11
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `uvicorn main:app --host 0.0.0.0 --port 8000`
4. Add environment variable in Settings:
   - `GOOGLE_API_KEY`: Your free API key from https://aistudio.google.com/apikey
5. Deploy! Render will provide a URL like `https://your-app.onrender.com`

**Note**: Render free tier spins down after 15 minutes of inactivity. Upgrade to paid tier for 24/7 availability.

#### Memory Limits

Render's free tier provides 2 GB RAM. The backend is optimized to stay well within this:

```bash
# Test locally with the same 2 GB constraint
docker build -t soc1-agent:latest .
docker run --rm -it \
  --memory=2g --memory-swap=2g \
  -p 8000:8000 \
  -e GOOGLE_API_KEY="your-key" \
  -e ENABLE_MEM_LOG=1 \
  soc1-agent:latest
```

Set `ENABLE_MEM_LOG=1` to log RSS memory every 0.5 s for debugging.

### Frontend Deployment (Vercel)

1. Go to [vercel.com](https://vercel.com) and sign in with GitHub
2. Import your repository
3. Configure:
   - **Framework**: Vite
   - **Root Directory**: `frontend`
   - **Build Command**: `npm run build`
   - **Output Directory**: `dist`
4. Add environment variable:
   - `VITE_API_URL`: `https://your-backend.onrender.com/api` (from Render deployment)
5. Deploy!

### CORS Configuration

The backend CORS is configured to accept requests from:
- `http://localhost:5173` (local development)
- `https://soc1-agent.vercel.app/` (production)

To update for your frontend domain, edit `backend/main.py` line 20 and add your frontend URL to the `allow_origins` list:

```python
allow_origins=[
    "http://localhost:5173", 
    "https://your-frontend-domain.vercel.app"
]
```

## User Feedback System

After downloading results, users can optionally provide feedback (star rating, issue categories, comments). This helps you identify and fix common extraction issues.

**Analyze feedback:**
```bash
cd backend
python analyze_feedback.py              # View summary statistics
python analyze_feedback.py --detailed   # See all feedback entries
python analyze_feedback.py --export     # Export to CSV
```

**Feedback data stored in:**
- `backend/feedback/feedback_log.json` - All feedback entries
- `backend/feedback/{job_id}_corrected.xlsx` - User-corrected files (when provided)

**Use feedback to improve:**
1. Run analysis tool to identify common issues
2. Review corrected files to see what the AI missed
3. Update AI prompts in `agent.py` with better instructions/examples
4. Test with previous problem cases
5. Deploy improvements

## Programmatic Usage

You can also use the agent directly in Python:

```python
from agent import process_soc1_sync

result = process_soc1_sync(
    type_ii_path="path/to/soc1-report.pdf",
    management_review_path="path/to/template.xlsx",
    output_dir="path/to/output",
)

print(f"Filled template saved to: {result['output_path']}")
print(f"Analysis: {result['analysis']}")
```
