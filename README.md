# SOC 1 Type II Management Review Generator

AI-powered tool for processing SOC1 Type II audit reports and automatically filling management review Excel templates.

## Features

- **PDF Extraction**: Uses `pdfplumber` to extract text and tables from SOC1 Type II PDF reports
- **AI-Powered Mapping**: Uses Google Gemini AI (free tier) to intelligently map extracted content to Excel template fields
- **Excel Generation**: Automatically fills management review templates with extracted data
- **Gap Analysis**: Analyzes the extraction for completeness and provides recommendations

## Prerequisites

- Python 3.10+
- Node.js 18+
- Google API key (free) - Get at https://aistudio.google.com/apikey

## Setup

### Backend (FastAPI)

```bash
cd backend
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -r requirements.txt

# Set your Google API key (free)
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
│   Frontend      │────▶│   FastAPI       │────▶│   Gemini AI     │
│   (React)       │     │   Backend       │     │   (Free Tier)   │
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
- `google-generativeai` - Google Gemini AI client

### Frontend
- React 18
- TypeScript
- Vite

## Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `GOOGLE_API_KEY` | Google AI API key (free tier available) | Yes |

Get your free API key at: https://aistudio.google.com/apikey

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
