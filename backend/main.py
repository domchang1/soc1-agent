from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="SOC 1 Generator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/upload")
async def upload(
    type_ii_report: UploadFile = File(...),
    management_review: UploadFile = File(...),
) -> dict[str, Any]:
    upload_root = Path("uploads")
    upload_root.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S")
    type_ii_path = upload_root / f"type-ii-{timestamp}-{type_ii_report.filename}"
    management_path = upload_root / f"management-review-{timestamp}-{management_review.filename}"

    type_ii_bytes = await type_ii_report.read()
    management_bytes = await management_review.read()

    type_ii_path.write_bytes(type_ii_bytes)
    management_path.write_bytes(management_bytes)

    return {
        "message": "Upload complete. SOC 1 generation queued.",
        "type_ii_report": {
            "filename": type_ii_report.filename,
            "bytes": len(type_ii_bytes),
        },
        "management_review": {
            "filename": management_review.filename,
            "bytes": len(management_bytes),
        },
        "soc1_output": {
            "status": "pending",
            "preview": "This is a placeholder. The generated SOC 1 Type II will appear here.",
        },
    }
