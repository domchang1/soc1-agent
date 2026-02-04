import { useMemo, useState } from "react";

type UploadResponse = {
  message: string;
  type_ii_report: { filename: string; bytes: number };
  management_review: { filename: string; bytes: number };
  soc1_output: { status: string; preview: string };
};

const API_URL = "http://localhost:8000/api/upload";

export default function App() {
  const [typeIiFile, setTypeIiFile] = useState<File | null>(null);
  const [reviewFile, setReviewFile] = useState<File | null>(null);
  const [progress, setProgress] = useState(0);
  const [status, setStatus] = useState("Awaiting files.");
  const [error, setError] = useState<string | null>(null);
  const [output, setOutput] = useState<string>("");
  const [isUploading, setIsUploading] = useState(false);

  const canSubmit = useMemo(
    () => Boolean(typeIiFile && reviewFile && !isUploading),
    [typeIiFile, reviewFile, isUploading]
  );

  const handleUpload = () => {
    if (!typeIiFile || !reviewFile) {
      setError("Please select both files before uploading.");
      return;
    }

    setError(null);
    setStatus("Uploading files...");
    setProgress(0);
    setIsUploading(true);

    const formData = new FormData();
    formData.append("type_ii_report", typeIiFile);
    formData.append("management_review", reviewFile);

    const xhr = new XMLHttpRequest();
    xhr.open("POST", API_URL);

    xhr.upload.onprogress = (event) => {
      if (event.lengthComputable) {
        const percent = Math.round((event.loaded / event.total) * 100);
        setProgress(percent);
      }
    };

    xhr.onload = () => {
      setIsUploading(false);
      if (xhr.status >= 200 && xhr.status < 300) {
        const response: UploadResponse = JSON.parse(xhr.responseText);
        setStatus(response.message);
        setOutput(
          [
            `SOC 1 output status: ${response.soc1_output.status}`,
            "",
            response.soc1_output.preview,
            "",
            "Files received:",
            `- Type II report: ${response.type_ii_report.filename} (${response.type_ii_report.bytes} bytes)`,
            `- Management review: ${response.management_review.filename} (${response.management_review.bytes} bytes)`,
          ].join("\n")
        );
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

  return (
    <div className="page">
      <section className="hero">
        <h1>SOC 1 Type II Generator</h1>
        <p>
          Upload a Type II report and the blank SOC 1 Management Review worksheet
          to generate a draft SOC 1 report. Your documents stay local to your
          environment.
        </p>
      </section>

      <section className="card upload-card">
        <div className="form-grid">
          <div>
            <label htmlFor="type-ii">Type II report (PDF or DOCX)</label>
            <input
              id="type-ii"
              type="file"
              onChange={(event) => setTypeIiFile(event.target.files?.[0] ?? null)}
            />
          </div>
          <div>
            <label htmlFor="review">SOC 1 Management Review (Excel)</label>
            <input
              id="review"
              type="file"
              onChange={(event) => setReviewFile(event.target.files?.[0] ?? null)}
            />
          </div>
          <button type="button" onClick={handleUpload} disabled={!canSubmit}>
            {isUploading ? "Uploading..." : "Generate SOC 1"}
          </button>
        </div>
      </section>

      <section className="card">
        <div className="progress-wrap">
          <div className="status">{status}</div>
          <div className="progress-bar">
            <div className="progress-value" style={{ width: `${progress}%` }} />
          </div>
          {error ? <div className="error">{error}</div> : null}
        </div>
      </section>

      <section className="card">
        <h3>Output</h3>
        <div className={`output ${output ? "" : "output-empty"}`}>
          {output || "Upload files to begin."}
        </div>
      </section>
    </div>
  );
}
