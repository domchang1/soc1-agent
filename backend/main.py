from __future__ import annotations

import os
import json
import traceback
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any

import psutil
from fastapi import FastAPI, File, UploadFile, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

from agent import process_soc1_documents, log_memory

app = FastAPI(title="SOC 1 Generator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "https://soc1-agent.vercel.app"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------------------------------------------------------------------
# Storage
# ---------------------------------------------------------------------------
# In-memory job status storage (use Redis/DB in production)
job_status: dict[str, dict[str, Any]] = {}
# Map job_id -> Path on disk for the generated Excel file
job_files: dict[str, Path] = {}

# Persistent directory for generated output files
OUTPUT_DIR = Path("output_files")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# How long to keep completed jobs before cleanup (1 hour)
JOB_TTL = timedelta(hours=1)


class FeedbackRequest(BaseModel):
    """Request model for feedback submission."""
    rating: int
    feedback_text: str = ""
    issues: list[str] = []


# ---------------------------------------------------------------------------
# TTL cleanup
# ---------------------------------------------------------------------------
def cleanup_expired_jobs() -> int:
    """Remove job_status, job_files, and output files older than JOB_TTL.

    Returns the number of jobs cleaned up.
    """
    now = datetime.now(timezone.utc)
    expired_ids: list[str] = []

    for job_id, status in job_status.items():
        created_str = status.get("created_at", "")
        if not created_str:
            continue
        try:
            created = datetime.strptime(created_str, "%Y%m%d%H%M%S").replace(
                tzinfo=timezone.utc
            )
            if now - created > JOB_TTL:
                expired_ids.append(job_id)
        except ValueError:
            continue

    for job_id in expired_ids:
        # Remove output file from disk
        file_path = job_files.pop(job_id, None)
        if file_path and file_path.exists():
            file_path.unlink(missing_ok=True)
        # Remove status metadata
        job_status.pop(job_id, None)

    if expired_ids:
        print(f"[CLEANUP] Removed {len(expired_ids)} expired job(s): {expired_ids}")

    return len(expired_ids)


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------
@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/api/debug/memory")
def debug_memory() -> dict[str, Any]:
    """Return current process memory usage for debugging."""
    process = psutil.Process(os.getpid())
    mem = process.memory_info()
    return {
        "rss_mb": round(mem.rss / (1024 * 1024), 1),
        "vms_mb": round(mem.vms / (1024 * 1024), 1),
        "active_jobs": sum(
            1 for s in job_status.values() if s.get("status") == "processing"
        ),
        "stored_jobs": len(job_status),
        "stored_files": len(job_files),
        "stored_files_total_mb": round(
            sum(
                p.stat().st_size for p in job_files.values() if p.exists()
            ) / (1024 * 1024),
            1,
        ),
    }


async def process_job(job_id: str, type_ii_path: Path, management_path: Path):
    """Background task to process the SOC1 documents."""
    try:
        job_status[job_id]["status"] = "processing"
        job_status[job_id]["message"] = "Extracting PDF content and mapping to Excel template..."

        log_memory(f"process_job start ({job_id})")

        result = await process_soc1_documents(
            type_ii_path=type_ii_path,
            management_review_path=management_path,
            output_dir=OUTPUT_DIR,
        )

        # Move the generated file to the persistent output directory
        # (process_soc1_documents already writes to OUTPUT_DIR)
        output_path = Path(result["output_path"])
        if not output_path.exists():
            raise FileNotFoundError(f"Generated file not found: {output_path}")

        # Store path reference (NOT the file bytes)
        job_files[job_id] = output_path
        output_filename = output_path.name

        # Store only the lightweight analysis summary, not the full result dict
        analysis = result.get("analysis", {})
        analysis_summary = {
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

        job_status[job_id].update({
            "status": "completed",
            "message": "SOC 1 management review generated successfully.",
            "output_filename": output_filename,
            "analysis_summary": analysis_summary,
            "pdf_metadata": result.get("pdf_metadata"),
            "template_sheets": result.get("template_sheets"),
        })

        log_memory(f"process_job completed ({job_id})")

        # Clean up uploaded source files
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
    # Run TTL cleanup on each new upload to keep memory bounded
    cleanup_expired_jobs()

    upload_root = Path("uploads")
    upload_root.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S")
    job_id = f"job-{timestamp}"

    type_ii_path = upload_root / f"type-ii-{timestamp}-{type_ii_report.filename}"
    management_path = upload_root / f"management-review-{timestamp}-{management_review.filename}"

    # Stream files to disk in chunks instead of reading entirely into memory
    type_ii_size = 0
    with open(type_ii_path, "wb") as f:
        while chunk := await type_ii_report.read(1024 * 64):  # 64 KB chunks
            f.write(chunk)
            type_ii_size += len(chunk)

    management_size = 0
    with open(management_path, "wb") as f:
        while chunk := await management_review.read(1024 * 64):  # 64 KB chunks
            f.write(chunk)
            management_size += len(chunk)

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
                "bytes": type_ii_size,
            },
            "management_review": {
                "filename": management_review.filename,
                "bytes": management_size,
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
            "bytes": type_ii_size,
        },
        "management_review": {
            "filename": management_review.filename,
            "bytes": management_size,
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
    # analysis_summary is already pre-computed and stored directly in job_status
    return status


@app.get("/api/download/{job_id}")
def download_result(job_id: str):
    """Download the generated Excel file from disk."""
    if job_id not in job_status:
        return {"error": "Job not found"}

    status = job_status[job_id]
    if status.get("status") != "completed":
        return {"error": "Job not completed yet", "status": status.get("status")}

    file_path = job_files.get(job_id)
    if file_path is None or not file_path.exists():
        return {"error": "Generated file not found"}

    filename = status.get("output_filename", "soc1_management_review.xlsx")

    # Serve directly from disk — no need to load into memory
    return FileResponse(
        path=str(file_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )


@app.post("/api/feedback/{job_id}")
async def submit_feedback(
    job_id: str,
    feedback: FeedbackRequest,
) -> dict[str, Any]:
    """
    Submit feedback for a completed job.

    Args:
        job_id: The job ID to provide feedback for
        feedback: Feedback data including rating, text, and issues
    """
    if job_id not in job_status:
        return {"error": "Job not found"}

    feedback_dir = Path("feedback")
    feedback_dir.mkdir(exist_ok=True)

    # Store feedback metadata
    feedback_data = {
        "job_id": job_id,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "rating": feedback.rating,
        "feedback_text": feedback.feedback_text,
        "issues": feedback.issues,
        "job_metadata": {
            "type_ii_report": job_status[job_id].get("type_ii_report"),
            "management_review": job_status[job_id].get("management_review"),
            "analysis_summary": job_status[job_id].get("analysis_summary"),
        },
    }

    # Append feedback to JSON log
    feedback_log = feedback_dir / "feedback_log.json"

    if feedback_log.exists():
        with open(feedback_log, "r") as f:
            all_feedback = json.load(f)
    else:
        all_feedback = []

    all_feedback.append(feedback_data)

    with open(feedback_log, "w") as f:
        json.dump(all_feedback, f, indent=2)

    return {
        "status": "success",
        "message": "Thank you for your feedback! This helps us improve the extraction quality.",
        "feedback_id": f"{job_id}_{len(all_feedback)}",
    }


@app.get("/api/feedback/stats")
def get_feedback_stats() -> dict[str, Any]:
    """Get aggregated feedback statistics (for admin/monitoring)."""
    feedback_dir = Path("feedback")
    feedback_log = feedback_dir / "feedback_log.json"

    if not feedback_log.exists():
        return {
            "total_feedback": 0,
            "average_rating": 0,
            "common_issues": {},
        }

    with open(feedback_log, "r") as f:
        all_feedback = json.load(f)

    if not all_feedback:
        return {
            "total_feedback": 0,
            "average_rating": 0,
            "common_issues": {},
        }

    # Calculate statistics
    total = len(all_feedback)
    avg_rating = sum(f["rating"] for f in all_feedback) / total

    # Count issue frequencies
    issue_counts: dict[str, int] = {}
    for feedback in all_feedback:
        for issue in feedback.get("issues", []):
            issue_counts[issue] = issue_counts.get(issue, 0) + 1

    return {
        "total_feedback": total,
        "average_rating": round(avg_rating, 2),
        "common_issues": issue_counts,
        "recent_feedback": all_feedback[-5:],  # Last 5 feedback entries
    }


@app.post("/api/cleanup-uploads")
def cleanup_uploads() -> dict[str, Any]:
    """Manually clear all files in the uploads folder."""
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


@app.post("/api/cleanup-old-files")
def cleanup_old_files(max_age_hours: int = 24) -> dict[str, Any]:
    """
    Clean up old output files to prevent disk space issues.
    MEMORY FIX: This prevents accumulation of generated files.
    """
    from datetime import datetime, timedelta, timezone
    import time
    
    cutoff_time = time.time() - (max_age_hours * 3600)
    files_deleted = 0
    errors = []
    
    try:
        # Clean old output files
        if output_dir.exists():
            for file_path in output_dir.iterdir():
                if file_path.is_file():
                    try:
                        if file_path.stat().st_mtime < cutoff_time:
                            file_path.unlink()
                            files_deleted += 1
                    except Exception as e:
                        errors.append(f"Failed to delete {file_path.name}: {str(e)}")
        
        # Clean old job status entries
        old_jobs = [
            job_id for job_id, status in job_status.items()
            if status.get("created_at") and 
            datetime.fromisoformat(status["created_at"]).timestamp() < cutoff_time
        ]
        for job_id in old_jobs:
            del job_status[job_id]
        
        return {
            "status": "success",
            "message": f"Cleaned up {files_deleted} old file(s) and {len(old_jobs)} job entries",
            "files_deleted": files_deleted,
            "jobs_cleaned": len(old_jobs),
            "errors": errors if errors else None,
        }
    except Exception as e:
        return {
            "status": "failed",
            "message": f"Cleanup failed: {str(e)}",
            "error": str(e),
        }
