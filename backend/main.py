from __future__ import annotations

import io
import os
import traceback
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, UploadFile, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from agent import process_soc1_documents

app = FastAPI(title="SOC 1 Generator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "https://soc1-agent.vercel.app/"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory job status storage (use Redis/DB in production)
# Store both metadata and the generated Excel file bytes in memory
job_status: dict[str, dict[str, Any]] = {}
job_files: dict[str, bytes] = {}  # Store generated Excel files in memory


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


async def process_job(job_id: str, type_ii_path: Path, management_path: Path):
    """Background task to process the SOC1 documents."""
    try:
        job_status[job_id]["status"] = "processing"
        job_status[job_id]["message"] = "Extracting PDF content and mapping to Excel template..."

        # Use a temporary directory for the processing pipeline
        import tempfile
        import shutil
        
        temp_dir = tempfile.mkdtemp()

        result = await process_soc1_documents(
            type_ii_path=type_ii_path,
            management_review_path=management_path,
            output_dir=Path(temp_dir),
        )

        # Read the generated Excel file into memory
        output_path = Path(result["output_path"])
        if output_path.exists():
            file_bytes = output_path.read_bytes()
            job_files[job_id] = file_bytes
            output_filename = output_path.name
        else:
            raise FileNotFoundError(f"Generated file not found: {output_path}")

        job_status[job_id].update({
            "status": "completed",
            "message": "SOC 1 management review generated successfully.",
            "result": result,
            "output_filename": output_filename,
        })

        # Clean up temporary files after reading
        shutil.rmtree(temp_dir, ignore_errors=True)
        
        # Clean up uploaded files
        type_ii_path.unlink(missing_ok=True)
        management_path.unlink(missing_ok=True)

    except Exception as e:
        job_status[job_id].update({
            "status": "failed",
            "message": f"Processing failed: {str(e)}",
            "error": traceback.format_exc(),
        })
        
        # Clean up uploaded files even on error
        type_ii_path.unlink(missing_ok=True)
        management_path.unlink(missing_ok=True)


@app.post("/api/upload")
async def upload(
    background_tasks: BackgroundTasks,
    type_ii_report: UploadFile = File(...),
    management_review: UploadFile = File(...),
) -> dict[str, Any]:
    upload_root = Path("uploads")
    upload_root.mkdir(parents=True, exist_ok=True)

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
            "total_cuecs": analysis.get("total_cuecs_identified", "N/A"),
            "cells_needing_review": analysis.get("cells_needing_review", {
                "low_confidence": 0,
                "medium_confidence": 0,
            }),
            "summary": analysis.get("summary", ""),
            "key_findings": analysis.get("key_findings", [])[:5],
            "cuec_findings": analysis.get("cuec_findings", [])[:5],
        }

    return status


@app.get("/api/download/{job_id}")
def download_result(job_id: str):
    """Download the generated Excel file from memory."""
    if job_id not in job_status:
        return {"error": "Job not found"}

    status = job_status[job_id]
    if status.get("status") != "completed":
        return {"error": "Job not completed yet", "status": status.get("status")}

    if job_id not in job_files:
        return {"error": "Generated file not found in memory"}

    file_bytes = job_files[job_id]
    filename = status.get("output_filename", "soc1_management_review.xlsx")

    # Return the file as a streaming response
    return StreamingResponse(
        io.BytesIO(file_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.post("/api/cleanup-uploads")
def cleanup_uploads() -> dict[str, Any]:
    """Manually clear all files in the uploads folder."""
    import shutil
    
    upload_root = Path("uploads")
    
    if not upload_root.exists():
        return {
            "status": "success",
            "message": "Uploads folder does not exist",
            "files_deleted": 0,
        }
    
    files_deleted = 0
    errors = []
    
    try:
        for file_path in upload_root.iterdir():
            if file_path.is_file():
                try:
                    file_path.unlink()
                    files_deleted += 1
                except Exception as e:
                    errors.append(f"Failed to delete {file_path.name}: {str(e)}")
        
        return {
            "status": "success",
            "message": f"Deleted {files_deleted} file(s) from uploads folder",
            "files_deleted": files_deleted,
            "errors": errors if errors else None,
        }
    except Exception as e:
        return {
            "status": "failed",
            "message": f"Cleanup failed: {str(e)}",
            "files_deleted": 0,
            "error": str(e),
        }
