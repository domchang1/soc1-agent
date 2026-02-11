# OOM Debug: Mirror Render's container constraints (Python 3.11, single worker)
# Run locally with: docker run --rm -it --memory=2g --memory-swap=2g -p 8000:8000 -e GOOGLE_API_KEY=xxx soc1-agent:latest

FROM python:3.11-slim

WORKDIR /app

# Install backend dependencies (matches Render)
COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy backend code
COPY backend/ .

# Render uses PORT env (often 10000); default 8000 for local
ENV PORT=8000
EXPOSE 8000

# Single worker to match Render and avoid worker multiplication OOM
# Set ENABLE_MEM_LOG=1 to enable high-frequency memory logging for debugging
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000} --workers 1"]
