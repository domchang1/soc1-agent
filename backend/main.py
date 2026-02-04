from __future__ import annotations

import os
import traceback
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, UploadFile, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from agent import process_soc1_documents

app = FastAPI(title="SOC 1 Generator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory job status storage (use Redis/DB in production)
job_status: dict[str, dict[str, Any]] = {}


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


async def process_job(job_id: str, type_ii_path: Path, management_path: Path, output_dir: Path):
    """Background task to process the SOC1 documents."""
    try:
        job_status[job_id]["status"] = "processing"
        job_status[job_id]["message"] = "Extracting PDF content and mapping to Excel template..."

        result = await process_soc1_documents(
            type_ii_path=type_ii_path,
            management_review_path=management_path,
            output_dir=output_dir,
        )

        job_status[job_id].update({
            "status": "completed",
            "message": "SOC 1 management review generated successfully.",
            "result": result,
            "output_path": result["output_path"],
        })

    except Exception as e:
        job_status[job_id].update({
            "status": "failed",
            "message": f"Processing failed: {str(e)}",
            "error": traceback.format_exc(),
        })


@app.post("/api/upload")
async def upload(
    background_tasks: BackgroundTasks,
    type_ii_report: UploadFile = File(...),
    management_review: UploadFile = File(...),
) -> dict[str, Any]:
    upload_root = Path("uploads")
    output_root = Path("outputs")
    upload_root.mkdir(parents=True, exist_ok=True)
    output_root.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S")
    job_id = f"job-{timestamp}"

    type_ii_path = upload_root / f"type-ii-{timestamp}-{type_ii_report.filename}"
    management_path = upload_root / f"management-review-{timestamp}-{management_review.filename}"

    type_ii_bytes = await type_ii_report.read()
    management_bytes = await management_review.read()

    type_ii_path.write_bytes(type_ii_bytes)
    management_path.write_bytes(management_bytes)

    # Initialize job status
    job_status[job_id] = {
        "status": "queued",
        "message": "Upload complete. SOC 1 generation starting...",
        "created_at": timestamp,
        "type_ii_report": type_ii_report.filename,
        "management_review": management_review.filename,
    }

    # Check if Google API key is configured
    if not os.environ.get("GOOGLE_API_KEY"):
        job_status[job_id].update({
            "status": "failed",
            "message": "GOOGLE_API_KEY not configured. Get a free key at https://aistudio.google.com/apikey",
        })
        return {
            "job_id": job_id,
            "message": "Upload complete but processing cannot start - GOOGLE_API_KEY not configured.",
            "type_ii_report": {
                "filename": type_ii_report.filename,
                "bytes": len(type_ii_bytes),
            },
            "management_review": {
                "filename": management_review.filename,
                "bytes": len(management_bytes),
            },
            "soc1_output": {
                "status": "failed",
                "preview": "Please set GOOGLE_API_KEY environment variable. Get a free key at https://aistudio.google.com/apikey",
            },
        }

    # Start background processing
    background_tasks.add_task(
        process_job,
        job_id,
        type_ii_path,
        management_path,
        output_root,
    )

    return {
        "job_id": job_id,
        "message": "Upload complete. SOC 1 generation started.",
        "type_ii_report": {
            "filename": type_ii_report.filename,
            "bytes": len(type_ii_bytes),
        },
        "management_review": {
            "filename": management_review.filename,
            "bytes": len(management_bytes),
        },
        "soc1_output": {
            "status": "processing",
            "preview": "Processing has started. Poll /api/status/{job_id} for updates.",
        },
    }


@app.get("/api/status/{job_id}")
def get_status(job_id: str) -> dict[str, Any]:
    """Get the status of a processing job."""
    if job_id not in job_status:
        return {"error": "Job not found", "job_id": job_id}

    status = job_status[job_id].copy()

    # Add analysis preview if completed
    if status.get("status") == "completed" and status.get("result"):
        analysis = status["result"].get("analysis", {})
        status["analysis_summary"] = {
            "total_controls": analysis.get("total_controls_identified", "N/A"),
            "exceptions": analysis.get("controls_with_exceptions", "N/A"),
            "summary": analysis.get("summary", ""),
            "key_findings": analysis.get("key_findings", [])[:5],
        }

    return status


@app.get("/api/download/{job_id}")
def download_result(job_id: str):
    """Download the generated Excel file."""
    if job_id not in job_status:
        return {"error": "Job not found"}

    status = job_status[job_id]
    if status.get("status") != "completed":
        return {"error": "Job not completed yet", "status": status.get("status")}

    output_path = status.get("output_path")
    if not output_path or not Path(output_path).exists():
        return {"error": "Output file not found"}

    return FileResponse(
        path=output_path,
        filename=Path(output_path).name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
