# SOC 1 Type II Management Review Generator

AI-powered agent for processing SOC1 Type II audit reports and automatically filling management review Excel templates.

## Features

- **PDF Extraction**: Uses `pdfplumber` to extract text and tables from SOC1 Type II PDF reports
- **AI-Powered Mapping**: Supports multiple AI providers (Google Gemini free tier, Anthropic Claude)
- **Excel Generation**: Automatically fills management review templates with extracted data
- **Gap Analysis**: Analyzes the extraction for completeness and provides recommendations

## Prerequisites

- Python 3.10+
- Node.js 18+
- AI API key: **Google Gemini** (free tier)

## Setup

### Backend (FastAPI)

```bash
cd backend
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -r requirements.txt

# Set your AI API key (choose one):
export GOOGLE_API_KEY="your-google-api-key"

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

- `GET /api/health` - Health check
- `POST /api/upload` - Upload files and start processing
- `GET /api/status/{job_id}` - Check processing status
- `GET /api/download/{job_id}` - Download the filled Excel file

## Architecture

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   Frontend      │────▶│   FastAPI       │────▶│   Agent         │
│   (React)       │     │   Backend       │     │   (Claude AI)   │
└─────────────────┘     └─────────────────┘     └─────────────────┘
                               │                        │
                               ▼                        ▼
                        ┌─────────────────┐     ┌─────────────────┐
                        │   pdfplumber    │     │   openpyxl      │
                        │   (PDF Extract) │     │   (Excel Write) │
                        └─────────────────┘     └─────────────────┘
```

## Dependencies

### Backend
- `fastapi` - Web framework
- `uvicorn` - ASGI server
- `pdfplumber` - PDF text and table extraction
- `openpyxl` - Excel file manipulation
- `anthropic` - Claude AI API client

### Frontend
- React 18
- TypeScript
- Vite

## Programmatic Usage

You can also use the agent directly in Python:

```python
from agent import process_soc1_sync

# Auto-detect provider based on available API keys
result = process_soc1_sync(
    type_ii_path="path/to/soc1-report.pdf",
    management_review_path="path/to/template.xlsx",
    output_dir="path/to/output",
)

# Or specify a provider explicitly
result = process_soc1_sync(
    type_ii_path="path/to/soc1-report.pdf",
    management_review_path="path/to/template.xlsx",
    output_dir="path/to/output",
    provider="gemini",  # or "anthropic"
)

print(f"Filled template saved to: {result['output_path']}")
print(f"Analysis: {result['analysis']}")
```
