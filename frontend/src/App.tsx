import { useCallback, useEffect, useMemo, useRef, useState } from "react";

type UploadResponse = {
  job_id: string;
  message: string;
  type_ii_report: { filename: string; bytes: number };
  management_review: { filename: string; bytes: number };
  soc1_output: { status: string; preview: string };
};

type JobStatus = {
  status: "queued" | "processing" | "completed" | "failed";
  message: string;
  error?: string;
  output_path?: string;
  analysis_summary?: {
    total_controls: string | number;
    exceptions: string | number;
    summary: string;
    key_findings: string[];
  };
};

const API_BASE = "http://localhost:8000/api";

export default function App() {
  const [typeIiFile, setTypeIiFile] = useState<File | null>(null);
  const [reviewFile, setReviewFile] = useState<File | null>(null);
  const [progress, setProgress] = useState(0);
  const [status, setStatus] = useState("Awaiting files.");
  const [error, setError] = useState<string | null>(null);
  const [output, setOutput] = useState<string>("");
  const [isUploading, setIsUploading] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [jobId, setJobId] = useState<string | null>(null);
  const [canDownload, setCanDownload] = useState(false);
  const pollingRef = useRef<ReturnType<typeof setInterval> | null>(null);

  const canSubmit = useMemo(
    () => Boolean(typeIiFile && reviewFile && !isUploading && !isProcessing),
    [typeIiFile, reviewFile, isUploading, isProcessing]
  );

  const pollJobStatus = useCallback(async (id: string) => {
    try {
      const response = await fetch(`${API_BASE}/status/${id}`);
      const data: JobStatus = await response.json();

      setStatus(data.message);

      if (data.status === "processing") {
        setProgress(50);
        setOutput(
          [
            "Status: Processing...",
            "",
            "The AI agent is:",
            "• Extracting text and tables from the PDF",
            "• Analyzing the SOC1 Type II report content",
            "• Mapping controls to the Excel template",
            "",
            "This may take 1-2 minutes depending on document size.",
          ].join("\n")
        );
      } else if (data.status === "completed") {
        setProgress(100);
        setIsProcessing(false);
        setCanDownload(true);

        const summary = data.analysis_summary;
        const outputLines = [
          "✓ Processing Complete!",
          "",
          "Analysis Summary:",
          `• Total Controls Identified: ${summary?.total_controls ?? "N/A"}`,
          `• Controls with Exceptions: ${summary?.exceptions ?? "N/A"}`,
          "",
        ];

        if (summary?.summary) {
          outputLines.push("Summary:", summary.summary, "");
        }

        if (summary?.key_findings && summary.key_findings.length > 0) {
          outputLines.push("Key Findings:");
          summary.key_findings.forEach((finding, i) => {
            outputLines.push(`  ${i + 1}. ${finding}`);
          });
          outputLines.push("");
        }

        outputLines.push("Click 'Download Result' to get the filled Excel file.");

        setOutput(outputLines.join("\n"));

        // Stop polling
        if (pollingRef.current) {
          clearInterval(pollingRef.current);
          pollingRef.current = null;
        }
      } else if (data.status === "failed") {
        setProgress(0);
        setIsProcessing(false);
        setError(data.message);
        setOutput(
          [
            "✗ Processing Failed",
            "",
            data.message,
            "",
            data.error ? `Error details:\n${data.error}` : "",
          ].join("\n")
        );

        // Stop polling
        if (pollingRef.current) {
          clearInterval(pollingRef.current);
          pollingRef.current = null;
        }
      }
    } catch (err) {
      console.error("Polling error:", err);
    }
  }, []);

  useEffect(() => {
    return () => {
      if (pollingRef.current) {
        clearInterval(pollingRef.current);
      }
    };
  }, []);

  const handleUpload = () => {
    if (!typeIiFile || !reviewFile) {
      setError("Please select both files before uploading.");
      return;
    }

    setError(null);
    setStatus("Uploading files...");
    setProgress(0);
    setIsUploading(true);
    setCanDownload(false);
    setJobId(null);

    const formData = new FormData();
    formData.append("type_ii_report", typeIiFile);
    formData.append("management_review", reviewFile);

    const xhr = new XMLHttpRequest();
    xhr.open("POST", `${API_BASE}/upload`);

    xhr.upload.onprogress = (event) => {
      if (event.lengthComputable) {
        const percent = Math.round((event.loaded / event.total) * 25);
        setProgress(percent);
      }
    };

    xhr.onload = () => {
      setIsUploading(false);
      if (xhr.status >= 200 && xhr.status < 300) {
        const response: UploadResponse = JSON.parse(xhr.responseText);
        setStatus(response.message);
        setJobId(response.job_id);

        if (response.soc1_output.status === "failed") {
          setError(response.soc1_output.preview);
          setOutput(
            [
              "Upload successful, but processing cannot start.",
              "",
              `Type II report: ${response.type_ii_report.filename} (${response.type_ii_report.bytes} bytes)`,
              `Management review: ${response.management_review.filename} (${response.management_review.bytes} bytes)`,
              "",
              "Error: " + response.soc1_output.preview,
            ].join("\n")
          );
        } else {
          setProgress(30);
          setIsProcessing(true);
          setOutput(
            [
              "Files uploaded successfully. Starting AI processing...",
              "",
              `Type II report: ${response.type_ii_report.filename}`,
              `Management review: ${response.management_review.filename}`,
              "",
              `Job ID: ${response.job_id}`,
            ].join("\n")
          );

          // Start polling for job status
          pollingRef.current = setInterval(() => {
            pollJobStatus(response.job_id);
          }, 2000);
        }
      } else {
        setError("Upload failed. Please try again.");
        setStatus("Upload failed.");
      }
    };

    xhr.onerror = () => {
      setIsUploading(false);
      setError("Network error. Is the backend running?");
      setStatus("Network error.");
    };

    xhr.send(formData);
  };

  const handleDownload = () => {
    if (!jobId) return;
    window.open(`${API_BASE}/download/${jobId}`, "_blank");
  };

  const handleReset = () => {
    setTypeIiFile(null);
    setReviewFile(null);
    setProgress(0);
    setStatus("Awaiting files.");
    setError(null);
    setOutput("");
    setIsUploading(false);
    setIsProcessing(false);
    setJobId(null);
    setCanDownload(false);

    if (pollingRef.current) {
      clearInterval(pollingRef.current);
      pollingRef.current = null;
    }

    // Reset file inputs
    const inputs = document.querySelectorAll<HTMLInputElement>('input[type="file"]');
    inputs.forEach((input) => {
      input.value = "";
    });
  };

  return (
    <div className="page">
      <section className="hero">
        <h1>SOC 1 Type II Generator</h1>
        <p>
          Upload a Type II report (PDF) and the blank SOC 1 Management Review worksheet (Excel)
          to generate a filled management review. AI-powered extraction maps controls,
          test results, and findings automatically.
        </p>
      </section>

      <section className="card upload-card">
        <div className="form-grid">
          <div>
            <label htmlFor="type-ii">Type II Report (PDF)</label>
            <input
              id="type-ii"
              type="file"
              accept=".pdf"
              onChange={(event) => setTypeIiFile(event.target.files?.[0] ?? null)}
              disabled={isUploading || isProcessing}
            />
          </div>
          <div>
            <label htmlFor="review">Management Review Template (Excel)</label>
            <input
              id="review"
              type="file"
              accept=".xlsx,.xls"
              onChange={(event) => setReviewFile(event.target.files?.[0] ?? null)}
              disabled={isUploading || isProcessing}
            />
          </div>
          <div className="button-group">
            <button type="button" onClick={handleUpload} disabled={!canSubmit}>
              {isUploading ? "Uploading..." : isProcessing ? "Processing..." : "Generate SOC 1"}
            </button>
            {(canDownload || jobId) && (
              <button
                type="button"
                onClick={handleReset}
                className="secondary"
              >
                Reset
              </button>
            )}
          </div>
        </div>
      </section>

      <section className="card">
        <div className="progress-wrap">
          <div className="status">{status}</div>
          <div className="progress-bar">
            <div
              className={`progress-value ${isProcessing ? "progress-animated" : ""}`}
              style={{ width: `${progress}%` }}
            />
          </div>
          {error ? <div className="error">{error}</div> : null}
        </div>
      </section>

      <section className="card">
        <div className="output-header">
          <h3>Output</h3>
          {canDownload && (
            <button type="button" onClick={handleDownload} className="download-btn">
              Download Result
            </button>
          )}
        </div>
        <div className={`output ${output ? "" : "output-empty"}`}>
          {output || "Upload files to begin."}
        </div>
      </section>
    </div>
  );
}
