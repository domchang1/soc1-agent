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
    total_cuecs: string | number;
    cells_needing_review?: {
      low_confidence: number;
      medium_confidence: number;
    };
    summary: string;
    key_findings: string[];
    cuec_findings?: string[];
  };
};

const API_BASE = import.meta.env.VITE_API_URL || "http://localhost:8000/api";

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
  const [showFeedback, setShowFeedback] = useState(false);
  const [feedbackRating, setFeedbackRating] = useState(0);
  const [feedbackText, setFeedbackText] = useState("");
  const [feedbackIssues, setFeedbackIssues] = useState<string[]>([]);
  const [feedbackSubmitted, setFeedbackSubmitted] = useState(false);
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
            "• Mapping controls to '1.0 Management Review' sheet",
            "• Extracting Complementary User Entity Controls (CUECs)",
            "• Assessing confidence levels for each cell",
            "",
            "This may take 1-2 minutes depending on document size.",
          ].join("\n")
        );
      } else if (data.status === "completed") {
        setProgress(100);
        setIsProcessing(false);
        setCanDownload(true);

        const summary = data.analysis_summary;
        
        // Build HTML output with sections
        const outputHTML = `
<div class="output-section success">
  <h3 class="output-section-title">✓ Processing Complete</h3>
  <div class="output-stat">
    <div class="stat-item">
      <div class="stat-label">Total Controls</div>
      <div class="stat-value">${summary?.total_controls ?? "N/A"}</div>
    </div>
    <div class="stat-item">
      <div class="stat-label">With Exceptions</div>
      <div class="stat-value">${summary?.exceptions ?? "0"}</div>
    </div>
    <div class="stat-item">
      <div class="stat-label">CUECs Found</div>
      <div class="stat-value">${summary?.total_cuecs ?? "0"}</div>
    </div>
  </div>
</div>

${summary?.cells_needing_review ? `
<div class="output-section">
  <h3 class="output-section-title">Confidence Review</h3>
  <div class="output-stat">
    <div class="stat-item">
      <div class="stat-label">Low Confidence</div>
      <div class="stat-value" style="color: #fca5a5;">${summary.cells_needing_review.low_confidence ?? 0}</div>
    </div>
    <div class="stat-item">
      <div class="stat-label">Medium Confidence</div>
      <div class="stat-value" style="color: #fef08a;">${summary.cells_needing_review.medium_confidence ?? 0}</div>
    </div>
  </div>
</div>
` : ""}

${summary?.summary ? `
<div class="output-section">
  <h3 class="output-section-title">Summary</h3>
  <div class="output-content">${summary.summary}</div>
</div>
` : ""}

${summary?.key_findings && summary.key_findings.length > 0 ? `
<div class="output-section">
  <h3 class="output-section-title">Key Findings</h3>
  <ul class="output-findings">
    ${summary.key_findings.map((f: string) => `<li>${f}</li>`).join("")}
  </ul>
</div>
` : ""}

${summary?.cuec_findings && summary.cuec_findings.length > 0 ? `
<div class="output-section">
  <h3 class="output-section-title">CUEC Findings</h3>
  <ul class="output-findings cuec">
    ${summary.cuec_findings.map((f: string) => `<li>${f}</li>`).join("")}
  </ul>
</div>
` : ""}

<div class="output-section">
  <h3 class="output-section-title">Color Legend</h3>
  <div class="output-legend">
    <div class="legend-item">
      <div class="legend-color high"></div>
      <span>High confidence (found in PDF)</span>
    </div>
    <div class="legend-item">
      <div class="legend-color medium"></div>
      <span>Medium confidence (partial/inferred)</span>
    </div>
    <div class="legend-item">
      <div class="legend-color low"></div>
      <span>Low confidence (not found)</span>
    </div>
  </div>
</div>
`;

        // Store as HTML instead of plain text
        setOutput(outputHTML);

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
    // Show feedback form after download
    setTimeout(() => setShowFeedback(true), 1000);
  };

  const handleFeedbackSubmit = async () => {
    if (!jobId || feedbackRating === 0) {
      alert("Please provide a rating before submitting.");
      return;
    }

    try {
      const payload = {
        rating: feedbackRating,
        feedback_text: feedbackText,
        issues: feedbackIssues,
      };

      const response = await fetch(`${API_BASE}/feedback/${jobId}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      if (response.ok) {
        setFeedbackSubmitted(true);
        setTimeout(() => {
          setShowFeedback(false);
          setFeedbackSubmitted(false);
        }, 2000);
      } else {
        console.error("Response status:", response.status);
        const errorData = await response.json();
        console.error("Error details:", errorData);
        alert("Failed to submit feedback. Please try again.");
      }
    } catch (err) {
      console.error("Failed to submit feedback:", err);
      alert("Failed to submit feedback. Please try again.");
    }
  };

  const toggleIssue = (issue: string) => {
    setFeedbackIssues((prev) =>
      prev.includes(issue) ? prev.filter((i) => i !== issue) : [...prev, issue]
    );
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
    setShowFeedback(false);
    setFeedbackRating(0);
    setFeedbackText("");
    setFeedbackIssues([]);
    setFeedbackSubmitted(false);

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
        <h1>SOC-1 Management Review Generator</h1>
        <p>
          Upload a Type II report (PDF) and the blank SOC-1 Management Review worksheet (Excel)
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
              {isUploading ? "Uploading..." : isProcessing ? "Processing..." : "Generate SOC-1"}
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
          {output ? (
            <div dangerouslySetInnerHTML={{ __html: output }} />
          ) : (
            "Upload files to begin."
          )}
        </div>
      </section>

      {showFeedback && !feedbackSubmitted && (
        <section className="card feedback-card">
          <h3>Help Us Improve</h3>
          <p>How accurate was the extraction? Your feedback helps us improve the AI.</p>
          
          <div className="feedback-rating">
            <label>Rating:</label>
            <div className="stars">
              {[1, 2, 3, 4, 5].map((star) => (
                <button
                  key={star}
                  type="button"
                  className={`star ${feedbackRating >= star ? "filled" : ""}`}
                  onClick={() => setFeedbackRating(star)}
                >
                  ★
                </button>
              ))}
            </div>
          </div>

          <div className="feedback-issues">
            <label>What issues did you encounter? (optional)</label>
            <div className="issue-checkboxes">
              {[
                { id: "missing_controls", label: "Missing controls" },
                { id: "incorrect_mapping", label: "Incorrect field mapping" },
                { id: "low_confidence", label: "Too many low confidence cells" },
                { id: "missing_cuecs", label: "Missing CUECs" },
                { id: "formatting", label: "Formatting issues" },
                { id: "other", label: "Other" },
              ].map((issue) => (
                <label key={issue.id} className="checkbox-label">
                  <input
                    type="checkbox"
                    checked={feedbackIssues.includes(issue.id)}
                    onChange={() => toggleIssue(issue.id)}
                  />
                  {issue.label}
                </label>
              ))}
            </div>
          </div>

          <div className="feedback-text">
            <label>Additional comments (optional):</label>
            <textarea
              value={feedbackText}
              onChange={(e) => setFeedbackText(e.target.value)}
              placeholder="Tell us more about your experience..."
              rows={3}
            />
          </div>

          <div className="feedback-actions">
            <button type="button" onClick={handleFeedbackSubmit} className="submit-feedback">
              Submit Feedback
            </button>
            <button type="button" onClick={() => setShowFeedback(false)} className="secondary">
              Skip
            </button>
          </div>
        </section>
      )}

      {feedbackSubmitted && (
        <section className="card feedback-success">
          <h3>✓ Thank you!</h3>
          <p>Your feedback has been submitted and will help improve future extractions.</p>
        </section>
      )}
    </div>
  );
}
