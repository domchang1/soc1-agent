"""
SOC1 Type II Report Processing Agent

This module provides functionality to:
1. Extract text and tables from PDF Type II reports
2. Read Excel management review templates
3. Use Google Gemini AI to intelligently map PDF content to Excel fields
4. Return a filled-out Excel management review
"""

from __future__ import annotations

import gc
import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import psutil
from google import genai
from google.genai import types
import openpyxl
import pdfplumber
from dotenv import load_dotenv
from openpyxl.styles import PatternFill


def log_memory(label: str) -> float:
    """Log current process RSS memory usage. Returns RSS in MB."""
    process = psutil.Process(os.getpid())
    rss_mb = process.memory_info().rss / (1024 * 1024)
    print(f"[MEMORY] {label}: {rss_mb:.1f} MB RSS")
    return rss_mb

# Confidence level color fills
FILL_HIGH_CONFIDENCE = None  # No fill - confident
FILL_MEDIUM_CONFIDENCE = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")  # Light yellow
FILL_LOW_CONFIDENCE = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")  # Light red

load_dotenv()

GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
if not GOOGLE_API_KEY:
    raise ValueError("GOOGLE_API_KEY not found in environment variables. Please set it in .env file.")


@dataclass
class ExtractedPDFContent:
    """Container for extracted PDF content."""

    full_text: str
    pages: list[str]
    tables: list[list[list[str]]]
    metadata: dict[str, Any]


@dataclass
class ExcelTemplate:
    """Container for Excel template structure."""

    filepath: Path
    sheet_names: list[str]
    headers: dict[str, list[str]]  # sheet_name -> list of column headers
    header_to_col: dict[str, dict[str, int]]  # sheet_name -> {header: col_index}
    header_row: dict[str, int]  # sheet_name -> header row number
    structure: dict[str, list[dict[str, Any]]]  # sheet_name -> list of row dicts
    sheet_type: dict[str, str] = None  # sheet_name -> "form" or "table"
    form_fields: dict[str, list[dict[str, Any]]] = None  # sheet_name -> list of form fields
    
    def __post_init__(self):
        if self.sheet_type is None:
            self.sheet_type = {}
        if self.form_fields is None:
            self.form_fields = {}


class PDFExtractor:
    """Extracts text and tables from PDF documents using pdfplumber."""

    @staticmethod
    def extract(pdf_path: Path) -> ExtractedPDFContent:
        """
        Extract all text and tables from a PDF file.

        Args:
            pdf_path: Path to the PDF file

        Returns:
            ExtractedPDFContent with full text, per-page text, tables, and metadata
        """
        # MEMORY FIX: Build full_text directly instead of storing all pages
        text_parts: list[str] = []
        tables: list[list[list[str]]] = []
        metadata: dict[str, Any] = {}

        with pdfplumber.open(pdf_path) as pdf:
            metadata = {
                "num_pages": len(pdf.pages),
                "metadata": pdf.metadata or {},
            }

            # MEMORY FIX: Limit to first 50 pages to prevent memory issues
            max_pages = min(50, len(pdf.pages))
            
            for i, page in enumerate(pdf.pages[:max_pages]):
                # Extract text from page
                page_text = page.extract_text() or ""
                text_parts.append(page_text)

                # Extract tables from page (MEMORY FIX: limit to 20 tables total)
                if len(tables) < 20:
                    page_tables = page.extract_tables() or []
                    for table in page_tables:
                        if len(tables) >= 20:
                            break
                        # Clean up table data
                        cleaned_table = [
                            [str(cell) if cell is not None else "" for cell in row]
                            for row in table
                        ]
                        tables.append(cleaned_table)

        full_text = "\n\n--- Page Break ---\n\n".join(text_parts)
        
        # MEMORY FIX: Don't store individual pages, just return empty list
        # This saves significant memory for large PDFs
        return ExtractedPDFContent(
            full_text=full_text,
            pages=[],  # Empty to save memory
            tables=tables,
            metadata=metadata,
        )


class ExcelHandler:
    """Handles reading and writing Excel files using openpyxl."""

    @staticmethod
    def read_template(excel_path: Path) -> tuple[openpyxl.Workbook, ExcelTemplate]:
        """
        Read an Excel template and extract its structure.
        Detects whether each sheet is a "form" (questionnaire) or "table" (tabular data).

        Args:
            excel_path: Path to the Excel file

        Returns:
            Tuple of (workbook, ExcelTemplate)
        """
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet_names = wb.sheetnames
        headers: dict[str, list[str]] = {}
        header_to_col: dict[str, dict[str, int]] = {}
        header_rows: dict[str, int] = {}
        structure: dict[str, list[dict[str, Any]]] = {}
        sheet_types: dict[str, str] = {}
        form_fields: dict[str, list[dict[str, Any]]] = {}

        for sheet_name in sheet_names:
            ws = wb[sheet_name]
            sheet_headers: list[str] = []
            sheet_header_to_col: dict[str, int] = {}
            sheet_structure: list[dict[str, Any]] = []
            sheet_form_fields: list[dict[str, Any]] = []

            # Analyze sheet structure to determine if it's a form or table
            # Count how many rows have content only in column A (form-like)
            form_like_rows = 0
            table_like_rows = 0
            best_header_row = None
            best_header_count = 0
            
            for row_idx in range(1, min(30, ws.max_row + 1)):
                row_values = []
                for col in range(1, min(15, ws.max_column + 1)):
                    val = ws.cell(row=row_idx, column=col).value
                    if val is not None:
                        row_values.append((col, val))
                
                if len(row_values) == 1 and row_values[0][0] == 1:
                    # Only column A has content - form-like
                    form_like_rows += 1
                elif len(row_values) >= 3:
                    # Multiple columns have content - could be header row
                    table_like_rows += 1
                    if len(row_values) > best_header_count:
                        best_header_count = len(row_values)
                        best_header_row = row_idx

            # Determine sheet type
            is_form = form_like_rows > table_like_rows * 2
            sheet_types[sheet_name] = "form" if is_form else "table"
            
            if is_form:
                # Extract form fields from column A
                for row_idx in range(1, min(100, ws.max_row + 1)):
                    label = ws.cell(row=row_idx, column=1).value
                    if label and isinstance(label, str) and len(label.strip()) > 0:
                        # Check if this row has any existing data in other columns
                        existing_values = {}
                        for col in range(2, min(10, ws.max_column + 1)):
                            val = ws.cell(row=row_idx, column=col).value
                            if val is not None:
                                existing_values[col] = val
                        
                        sheet_form_fields.append({
                            "row": row_idx,
                            "label": label.strip(),
                            "answer_col": 2,  # Default answer column is B
                            "existing": existing_values,
                        })
                
                form_fields[sheet_name] = sheet_form_fields
                # For forms, create pseudo-headers based on common answer columns
                sheet_headers = ["Label", "Answer", "Notes", "Reference"]
                for i, h in enumerate(sheet_headers, 1):
                    sheet_header_to_col[h] = i
                    sheet_header_to_col[h.lower()] = i
                found_header_row = 1
                
            else:
                # Table sheet - find the best header row
                found_header_row = best_header_row or 1
                
                # Extract headers from the identified row
                for col_idx in range(1, ws.max_column + 1):
                    v = ws.cell(row=found_header_row, column=col_idx).value
                    header_name = str(v).strip() if v else f"Column_{col_idx}"
                    # Clean up header names (remove newlines, extra spaces)
                    header_name = ' '.join(header_name.split())
                    sheet_headers.append(header_name)
                    sheet_header_to_col[header_name] = col_idx
                    sheet_header_to_col[header_name.lower()] = col_idx
                    # Also map without special characters for fuzzy matching
                    clean_name = ''.join(c for c in header_name.lower() if c.isalnum() or c.isspace())
                    sheet_header_to_col[clean_name] = col_idx

            headers[sheet_name] = sheet_headers
            header_to_col[sheet_name] = sheet_header_to_col
            header_rows[sheet_name] = found_header_row

            # Extract existing data structure (for context to AI)
            if not is_form:
                for row_idx in range(found_header_row + 1, min(found_header_row + 100, ws.max_row + 1)):
                    row_data: dict[str, Any] = {"_row": row_idx}
                    has_data = False
                    for col_idx, header in enumerate(sheet_headers, 1):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        row_data[header] = {
                            "value": cell.value,
                            "column": col_idx,
                            "row": row_idx,
                        }
                        if cell.value is not None:
                            has_data = True
                    if has_data:
                        sheet_structure.append(row_data)

            structure[sheet_name] = sheet_structure

        return wb, ExcelTemplate(
            filepath=excel_path,
            sheet_names=sheet_names,
            headers=headers,
            header_to_col=header_to_col,
            header_row=header_rows,
            structure=structure,
            sheet_type=sheet_types,
            form_fields=form_fields,
        )

    @staticmethod
    def fill_template(
        wb: openpyxl.Workbook,
        template: ExcelTemplate,
        mappings: dict[str, list[dict[str, Any]]],
        output_path: Path,
    ) -> Path:
        """
        Fill the Excel template with mapped data and apply confidence-based coloring.

        Args:
            wb: The workbook to fill
            template: The ExcelTemplate with header mappings
            mappings: Dict of sheet_name -> list of row updates
            output_path: Path to save the filled template

        Returns:
            Path to the saved file

        Color coding based on confidence:
            - High confidence: No background color (default)
            - Medium confidence: Light yellow background
            - Low confidence: Light red background (missing info)
        """
        print(f"\n{'='*60}")
        print(f"FILL_TEMPLATE DEBUG")
        print(f"{'='*60}")
        print(f"Sheets in workbook: {wb.sheetnames}")
        print(f"Mappings received for sheets: {list(mappings.keys())}")
        print(f"Sheet types: {template.sheet_type}")
        
        # Debug: Show what data we received
        for sheet_name, rows in mappings.items():
            print(f"\n--- Data for '{sheet_name}' ({len(rows)} rows) ---")
            for i, row in enumerate(rows[:5]):  # Show first 5 rows
                print(f"  Row {i}: {row}")
        
        for sheet_name, rows in mappings.items():
            # Find matching sheet (exact or fuzzy match)
            actual_sheet_name = None
            if sheet_name in wb.sheetnames:
                actual_sheet_name = sheet_name
            else:
                # Try fuzzy matching for common variations
                sheet_name_lower = sheet_name.lower()
                for ws_name in wb.sheetnames:
                    if sheet_name_lower in ws_name.lower() or ws_name.lower() in sheet_name_lower:
                        actual_sheet_name = ws_name
                        break
                    # Also try matching key parts like "management review" or "user entity"
                    if "management review" in sheet_name_lower and "management review" in ws_name.lower():
                        actual_sheet_name = ws_name
                        break
                    if "user entity" in sheet_name_lower and "user entity" in ws_name.lower():
                        actual_sheet_name = ws_name
                        break
                    if "cuec" in sheet_name_lower and "cuec" in ws_name.lower():
                        actual_sheet_name = ws_name
                        break

            if actual_sheet_name is None:
                print(f"Warning: Sheet '{sheet_name}' not found in workbook. Available: {wb.sheetnames}")
                continue

            print(f"\nProcessing sheet '{sheet_name}' -> actual: '{actual_sheet_name}' with {len(rows)} rows")
            ws = wb[actual_sheet_name]
            # Use the actual sheet name for header lookup
            lookup_sheet_name = actual_sheet_name
            header_map = template.header_to_col.get(lookup_sheet_name, {})
            header_row = template.header_row.get(lookup_sheet_name, 1)
            sheet_type = template.sheet_type.get(lookup_sheet_name, "table")
            
            print(f"  Sheet type: {sheet_type}")
            print(f"  Header row: {header_row}")
            print(f"  Header map keys: {list(header_map.keys())[:10]}...")
            
            # Normalize confidence values
            def normalize_confidence(c):
                if c in ("h", "high"):
                    return "high"
                elif c in ("m", "medium", "med"):
                    return "medium"
                elif c in ("l", "low"):
                    return "low"
                return "high"  # default

            for row_idx_in, row_update in enumerate(rows):
                row_idx = row_update.get("_row")
                if row_idx is None:
                    print(f"    Row {row_idx_in}: No _row field, skipping. Data: {row_update}")
                    continue

                cells_written = 0
                confidence_map = row_update.get("_confidence", row_update.get("_c", {}))
                row_confidence = row_update.get("_row_confidence", "high")
                
                # Handle form-style sheets differently
                if sheet_type == "form":
                    # For form sheets, write the "Answer" to column B at the specified row
                    answer = row_update.get("Answer") or row_update.get("answer")
                    if answer:
                        cell = ws.cell(row=row_idx, column=2, value=answer)
                        cells_written += 1
                    # Also check for any other column data
                    for col_name, value in row_update.items():
                        if col_name.startswith("_") or col_name.lower() == "answer":
                            continue
                        if value and isinstance(value, str):
                            # Try to map by column name
                            col_idx = header_map.get(col_name) or header_map.get(col_name.lower())
                            if col_idx:
                                ws.cell(row=row_idx, column=col_idx, value=value)
                                cells_written += 1
                else:
                    # Table-style sheet - match by column headers
                    # For table sheets, row should be after header row
                    if row_idx <= header_row:
                        row_idx = header_row + 1 + row_idx_in  # Auto-increment if row is invalid
                    
                    for col_name, value in row_update.items():
                        if col_name.startswith("_"):
                            continue

                        # Handle value with embedded confidence
                        cell_value = value
                        raw_confidence = confidence_map.get(col_name, "high")
                        cell_confidence = normalize_confidence(raw_confidence)

                        if isinstance(value, dict):
                            cell_value = value.get("value", value.get("v"))
                            raw_conf = value.get("confidence", value.get("c", raw_confidence))
                            cell_confidence = normalize_confidence(raw_conf)

                        # Skip if no value or empty string
                        if cell_value is None or cell_value == "":
                            continue

                        # Find the column index for this header - try multiple matching strategies
                        col_idx = None
                        
                        # Strategy 1: Exact match
                        col_idx = header_map.get(col_name)
                        
                        # Strategy 2: Case-insensitive match
                        if col_idx is None:
                            col_idx = header_map.get(col_name.lower().strip())
                        
                        # Strategy 3: Partial match (for long header names)
                        if col_idx is None:
                            col_name_clean = col_name.lower().strip()
                            for header_key, idx in header_map.items():
                                if isinstance(header_key, str):
                                    header_clean = header_key.lower().strip()
                                    # Check if col_name is contained in header or vice versa
                                    if col_name_clean in header_clean or header_clean in col_name_clean:
                                        col_idx = idx
                                        print(f"      Fuzzy match: '{col_name}' -> '{header_key}' (col {idx})")
                                        break
                        
                        # Strategy 4: Match by key words
                        if col_idx is None:
                            col_words = set(col_name.lower().split())
                            best_match_score = 0
                            for header_key, idx in header_map.items():
                                if isinstance(header_key, str):
                                    header_words = set(header_key.lower().split())
                                    common = col_words & header_words
                                    if len(common) > best_match_score and len(common) >= 2:
                                        best_match_score = len(common)
                                        col_idx = idx

                        if col_idx:
                            cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
                            cells_written += 1

                            # Apply confidence-based coloring
                            if cell_confidence == "low":
                                cell.fill = FILL_LOW_CONFIDENCE
                            elif cell_confidence == "medium":
                                cell.fill = FILL_MEDIUM_CONFIDENCE
                        else:
                            print(f"      Could not find column for '{col_name}' (value: {str(cell_value)[:50]}...)")
                
                if cells_written > 0:
                    print(f"    Row {row_idx}: Wrote {cells_written} cells")
                else:
                    print(f"    Row {row_idx}: No cells written! Data: {list(row_update.keys())}")

                # Apply row-level confidence coloring for empty/missing cells
                if row_confidence == "low":
                    for col_idx in range(1, len(template.headers.get(lookup_sheet_name, [])) + 1):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        if cell.value is None:
                            cell.fill = FILL_LOW_CONFIDENCE

        wb.save(output_path)
        return output_path


class SOC1Agent:
    """
    AI-powered agent for processing SOC1 Type II reports.

    Uses Google Gemini (free tier available) for intelligent content extraction.
    """

    def __init__(self, api_key: str | None = None):
        """
        Initialize the SOC1 Agent.

        Args:
            api_key: Google API key. If not provided, uses GOOGLE_API_KEY env var.
                     Get a free key at: https://aistudio.google.com/apikey
        """
        self.api_key = api_key or os.environ.get("GOOGLE_API_KEY")
        if not self.api_key:
            raise ValueError(
                "Google API key required. Set GOOGLE_API_KEY environment variable "
                "or pass api_key parameter.\n"
                "Get a free key at: https://aistudio.google.com/apikey"
            )
        # Initialize the genai client with API key
        self.client = genai.Client(api_key=self.api_key)
        # Use Gemini 2.5 Flash for free tier (latest, fast and capable)
        self.model = "gemini-2.5-flash"

    def _generate(self,
        prompt: str,
        max_tokens: int = 8192,
        retries: int = 3,
        expect_json: bool = False) -> str:
        """Generate a response from Gemini with retry logic."""
        import time
        
        last_error = None
        for attempt in range(retries):
            try:

                config_kwargs = dict(
                    max_output_tokens=max_tokens,
                    temperature=0.1,
                )
                if expect_json:
                    config_kwargs["response_mime_type"] = "application/json"

                response = self.client.models.generate_content(
                    model=self.model,
                    contents=prompt,
                    config=types.GenerateContentConfig(**config_kwargs),
                )
                
                # Check if response has text
                if getattr(response, "text", None):
                    return response.text

                # Fallback: pull text from candidates/parts when response.text is empty
                if getattr(response, "candidates", None):
                    cand0 = response.candidates[0]
                    content = getattr(cand0, "content", None)
                    parts = getattr(content, "parts", None) if content else None
                    if parts:
                        for p in parts:
                            t = getattr(p, "text", None)
                            if t:
                                return t

                raise ValueError("Empty response from Gemini")
                
            except Exception as e:
                last_error = e
                print(f"Attempt {attempt + 1}/{retries} failed: {str(e)}")
                if attempt < retries - 1:
                    wait_time = (attempt + 1) * 2  # Exponential backoff
                    print(f"Waiting {wait_time}s before retry...")
                    time.sleep(wait_time)
        
        raise RuntimeError(f"Failed after {retries} attempts. Last error: {last_error}")

    def _parse_json_response(self, response_text: str) -> dict[str, Any]:
        """Parse JSON from AI response, handling markdown code blocks and incomplete JSON."""
        original_text = response_text
        
        # Step 1: Remove markdown code block markers (handle both complete and incomplete blocks)
        # First try to match complete code blocks
        json_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", response_text)
        if json_match:
            response_text = json_match.group(1).strip()
        else:
            # Handle incomplete code blocks (no closing ```)
            if response_text.strip().startswith("```"):
                # Remove opening ``` and optional json tag
                response_text = re.sub(r"^```(?:json)?\s*", "", response_text.strip())
            # Also remove any trailing ``` if present
            response_text = re.sub(r"\s*```\s*$", "", response_text)

        # Step 2: Try direct parsing
        try:
            return json.loads(response_text)
        except json.JSONDecodeError:
            pass

        # Step 3: Find the JSON object boundaries
        json_start = response_text.find("{")
        if json_start < 0:
            print(f"No JSON object found in response: {original_text[:500]}")
            raise ValueError(f"No JSON object found in AI response. Response: {original_text[:200]}...")

        # Step 4: Try to extract valid JSON
        json_candidate = response_text[json_start:]
        
        # Try parsing as-is first
        try:
            return json.loads(json_candidate)
        except json.JSONDecodeError:
            pass

        # Step 5: Try to repair truncated JSON
        # Find the last valid closing brace
        json_candidate = self._repair_truncated_json(json_candidate)
        
        try:
            result = json.loads(json_candidate)
            print(f"Successfully parsed JSON after repair (found {len(result)} keys)")
            return result
        except json.JSONDecodeError as e:
            # Log detailed error for debugging
            print(f"Failed to parse JSON response after repair attempts.")
            print(f"JSON error: {e.msg} at position {e.pos}")
            print(f"Context around error: ...{json_candidate[max(0,e.pos-50):e.pos+50]}...")
            print(f"Original text length: {len(original_text)}")
            print(f"Repaired text length: {len(json_candidate)}")
            
            # Last resort: try to extract just the first complete sheet
            try:
                # Find the first complete array in the JSON
                first_array_end = json_candidate.find(']')
                if first_array_end > 0:
                    partial = json_candidate[:first_array_end + 1] + '}'
                    result = json.loads(partial)
                    print(f"Recovered partial JSON with {len(result)} keys")
                    return result
            except:
                pass
            
            raise ValueError(
                f"Could not parse AI response as JSON. The response may be truncated. "
                f"Response preview: {original_text[:300]}..."
            )

    def _repair_truncated_json(self, json_str: str) -> str:
        """Attempt to repair truncated JSON by closing open brackets and braces."""
        # Strategy: Find the last complete row entry and truncate there
        
        # First, try to find the last complete object in an array
        # Look for pattern: }, { which indicates complete objects in array
        last_complete = json_str.rfind('},')
        if last_complete > 0:
            # Check if cutting here gives us valid-ish JSON
            test_str = json_str[:last_complete + 1]  # Include the }
            open_braces = test_str.count('{') - test_str.count('}')
            open_brackets = test_str.count('[') - test_str.count(']')
            
            # If we have more opens than closes, this might be a good cut point
            if open_braces >= 0 and open_brackets >= 0:
                json_str = test_str
        
        # Remove any trailing incomplete content
        # Pattern 1: Incomplete key-value with nested object: "key": {incomplete
        json_str = re.sub(r',\s*"[^"]*"\s*:\s*\{[^{}]*$', '', json_str)
        # Pattern 2: Incomplete key-value with string: "key": "incomplete
        json_str = re.sub(r',\s*"[^"]*"\s*:\s*"[^"]*$', '', json_str)
        # Pattern 3: Incomplete key-value with object having confidence: "key": {"value": "x", "confidence
        json_str = re.sub(r',\s*"[^"]*"\s*:\s*\{[^{}]*$', '', json_str)
        # Pattern 4: Just an incomplete key: , "key
        json_str = re.sub(r',\s*"[^"]*$', '', json_str)
        # Pattern 5: Incomplete array element
        json_str = re.sub(r',\s*\[[^\[\]]*$', '', json_str)
        json_str = re.sub(r',\s*\{[^{}]*$', '', json_str)
        
        # Remove trailing comma if present
        json_str = re.sub(r',\s*$', '', json_str)
        
        # Count and balance brackets/braces
        open_braces = json_str.count('{') - json_str.count('}')
        open_brackets = json_str.count('[') - json_str.count(']')
        
        print(f"Repair: {open_brackets} unclosed brackets, {open_braces} unclosed braces")
        
        # Close any remaining open structures (close brackets before braces)
        for _ in range(max(0, open_brackets)):
            json_str += ']'
        for _ in range(max(0, open_braces)):
            json_str += '}'
        
        return json_str

    def _create_extraction_prompt(
        self,
        pdf_content: ExtractedPDFContent,
        template: ExcelTemplate,
    ) -> str:
        """Create a prompt for the AI to extract and map data."""
        
        # Find sheets that match management review and CUEC patterns
        management_sheet = None
        cuec_sheet = None
        deviations_sheet = None
        
        for sheet_name in template.sheet_names:
            sheet_lower = sheet_name.lower()
            if "management review" in sheet_lower and "1.0" in sheet_name:
                management_sheet = sheet_name
            if "user entity" in sheet_lower or "cuec" in sheet_lower or "comp user" in sheet_lower:
                cuec_sheet = sheet_name
            if "deviation" in sheet_lower:
                deviations_sheet = sheet_name

        prompt_parts = []
        prompt_parts.append("You are extracting data from a SOC1 Type II audit report to fill an Excel management review template.")
        prompt_parts.append("\n\n## PDF REPORT CONTENT:\n")
        # MEMORY FIX: Reduced from 80000 to 40000 chars to save memory
        prompt_parts.append(pdf_content.full_text[:40000])
        
        # Add extracted tables specifically (MEMORY FIX: reduced from 15 to 10 tables)
        if pdf_content.tables:
            prompt_parts.append("\n\n## EXTRACTED TABLES FROM PDF:\n")
            for i, table in enumerate(pdf_content.tables[:10], 1):
                prompt_parts.append(f"\nTable {i}:")
                # MEMORY FIX: Only show first 5 rows instead of 10
                for row in table[:5]:
                    prompt_parts.append(f"  {row}")
        
        prompt_parts.append("\n\n## EXCEL SHEETS TO FILL:\n")
        
        # Handle Management Review sheet (form-style)
        if management_sheet and template.sheet_type.get(management_sheet) == "form":
            form_fields = template.form_fields.get(management_sheet, [])
            prompt_parts.append(f"\n### Sheet: {management_sheet} (FORM-STYLE)")
            prompt_parts.append("This is a questionnaire. For each question, provide the answer from the PDF.")
            prompt_parts.append("Fields to fill:")
            for field in form_fields[:30]:  # Show up to 30 fields
                label = field['label'][:80]
                prompt_parts.append(f"  - Row {field['row']}: \"{label}\"")
        
        # Handle CUEC sheet (table-style)
        if cuec_sheet and template.sheet_type.get(cuec_sheet) == "table":
            cuec_headers = template.headers.get(cuec_sheet, [])
            prompt_parts.append(f"\n### Sheet: {cuec_sheet} (TABLE)")
            prompt_parts.append(f"Header row: {template.header_row.get(cuec_sheet)}")
            prompt_parts.append("Column headers (USE THESE EXACT NAMES):")
            for i, h in enumerate(cuec_headers[:10], 1):
                prompt_parts.append(f"  {i}. \"{h}\"")
            prompt_parts.append("\nLook for 'Complementary User Entity Controls' section in the PDF.")
            prompt_parts.append("Extract each CUEC with its control objective.")
        
        # Handle Deviations sheet
        if deviations_sheet and template.sheet_type.get(deviations_sheet) == "table":
            dev_headers = template.headers.get(deviations_sheet, [])
            prompt_parts.append(f"\n### Sheet: {deviations_sheet} (TABLE)")
            prompt_parts.append(f"Header row: {template.header_row.get(deviations_sheet)}")
            prompt_parts.append("Column headers:")
            for i, h in enumerate(dev_headers[:10], 1):
                prompt_parts.append(f"  {i}. \"{h}\"")
        
        # JSON format instructions
        prompt_parts.append("\n\n## RETURN FORMAT:")
        prompt_parts.append("Return ONLY valid JSON. No markdown code blocks, no commentary.")
        prompt_parts.append("{")
        
        if management_sheet:
            prompt_parts.append(f'  "{management_sheet}": [')
            prompt_parts.append('    {"_row": 4, "Answer": "Okta, Inc."},  // Row 4 answer')
            prompt_parts.append('    {"_row": 5, "Answer": "SOC1 Type II Report"},')
            prompt_parts.append("    ...")
            prompt_parts.append("  ],")
        
        if cuec_sheet:
            cuec_h = template.headers.get(cuec_sheet, ["No.", "Description"])
            h1 = cuec_h[0] if len(cuec_h) > 0 else "No."
            h2 = cuec_h[2] if len(cuec_h) > 2 else "Description"
            h3 = cuec_h[3] if len(cuec_h) > 3 else "Control Objective"
            start_row = template.header_row.get(cuec_sheet, 1) + 1
            prompt_parts.append(f'  "{cuec_sheet}": [')
            prompt_parts.append(f'    {{"_row": {start_row}, "{h1}": "1", "{h2}": "User entities are responsible for...", "{h3}": "CO 2 - Logical access"}},')
            prompt_parts.append(f'    {{"_row": {start_row + 1}, "{h1}": "2", "{h2}": "Another CUEC...", "{h3}": "CO 2 - Logical access"}}')
            prompt_parts.append("  ]")
        
        prompt_parts.append("}")
        
        prompt_parts.append("\n\n## IMPORTANT RULES:")
        prompt_parts.append("1. Use EXACT column names from the headers listed above")
        prompt_parts.append("2. For table sheets, _row is the row number (starts after header row)")
        prompt_parts.append("3. For form sheets, _row matches the row of the question label")
        prompt_parts.append("4. Extract ALL CUECs from the 'Complementary User Entity Controls' section")
        prompt_parts.append("5. Keep values concise but complete")
        prompt_parts.append("6. Return valid JSON only - no markdown, no extra text")
        
        return "\n".join(prompt_parts)

    def extract_and_map(
        self,
        pdf_content: ExtractedPDFContent,
        template: ExcelTemplate,
    ) -> dict[str, list[dict[str, Any]]]:
        """
        Use AI to extract data from PDF and map to Excel template.

        Args:
            pdf_content: Extracted PDF content
            template: Excel template structure

        Returns:
            Dict mapping sheet names to lists of row updates
        """
        try:
            prompt = self._create_extraction_prompt(pdf_content, template)
            print(f"\n{'='*60}")
            print("AI EXTRACTION")
            print(f"{'='*60}")
            print(f"Prompt length: {len(prompt)} chars")
            print(f"Template sheets: {template.sheet_names}")
            print(f"Sheet types: {template.sheet_type}")
            
            # MEMORY FIX: Reduced from 65536 to 32768 tokens to save memory
            response_text = self._generate(prompt, max_tokens=32768, expect_json=True)
            
            print(f"\nAI Response length: {len(response_text)} chars")
            print(f"Response preview: {response_text[:500]}...")
            
            result = self._parse_json_response(response_text)
            
            print(f"\nParsed result keys: {list(result.keys())}")
            for sheet, rows in result.items():
                print(f"  {sheet}: {len(rows)} rows")
                if rows:
                    print(f"    First row keys: {list(rows[0].keys())}")
            
            return result
        except Exception as e:
            print(f"\nMain extraction failed: {e}")
            import traceback
            traceback.print_exc()
            print("Trying simplified extraction...")
            return self._extract_simplified(pdf_content, template)

    def _extract_simplified(
        self,
        pdf_content: ExtractedPDFContent,
        template: ExcelTemplate,
    ) -> dict[str, list[dict[str, Any]]]:
        """Simplified extraction as fallback - extracts one sheet at a time."""
        result = {}
        
        for sheet_name in template.sheet_names:
            sheet_lower = sheet_name.lower()
            
            # Only process relevant sheets
            if not any(x in sheet_lower for x in ["management review", "user entity", "cuec", "comp user"]):
                continue
            
            headers = template.headers.get(sheet_name, [])
            header_row = template.header_row.get(sheet_name, 1)
            sheet_type = template.sheet_type.get(sheet_name, "table")
            
            print(f"\n--- Simplified extraction for '{sheet_name}' ---")
            print(f"  Sheet type: {sheet_type}, Headers: {headers[:5]}")
            
            if "user entity" in sheet_lower or "cuec" in sheet_lower or "comp user" in sheet_lower:
                # Special handling for CUEC sheet - look specifically for the CUEC section
                # MEMORY FIX: Only send relevant portion of PDF for CUECs
                # Look for CUEC section in the text
                cuec_section = ""
                full_text_lower = pdf_content.full_text.lower()
                cuec_start = full_text_lower.find("complementary user entity control")
                if cuec_start == -1:
                    cuec_start = full_text_lower.find("cuec")
                
                if cuec_start != -1:
                    # Extract 20000 chars around the CUEC section
                    start = max(0, cuec_start - 5000)
                    end = min(len(pdf_content.full_text), cuec_start + 15000)
                    cuec_section = pdf_content.full_text[start:end]
                else:
                    # Fallback: use last 20000 chars (CUECs often at end)
                    cuec_section = pdf_content.full_text[-20000:]
                
                prompt = f"""Extract Complementary User Entity Controls (CUECs) from this SOC1 Type II report.

Look for the section titled "Complementary User Entity Controls" or "CUECs".

PDF Content (relevant section):
{cuec_section}

Excel columns to fill (USE THESE EXACT NAMES):
{chr(10).join(f'  - "{h}"' for h in headers[:7])}

The header row is {header_row}, so data starts at row {header_row + 1}.

Return JSON array with one object per CUEC found:
[
  {{"_row": {header_row + 1}, "No.": "1", "{headers[2] if len(headers) > 2 else 'Description'}": "User entities are responsible for...", "{headers[3] if len(headers) > 3 else 'Control Objective'}": "CO 2 - Logical access"}},
  {{"_row": {header_row + 2}, "No.": "2", "{headers[2] if len(headers) > 2 else 'Description'}": "Another control...", "{headers[3] if len(headers) > 3 else 'Control Objective'}": "CO 2 - Logical access"}}
]

Return ONLY valid JSON array. No markdown, no commentary."""

            elif "management review" in sheet_lower:
                # Management review form
                form_fields = template.form_fields.get(sheet_name, [])
                fields_text = "\n".join(f"  Row {f['row']}: {f['label'][:60]}" for f in form_fields[:25])
                
                # MEMORY FIX: Reduced from 50000 to 30000 chars
                prompt = f"""Extract answers for this SOC1 Management Review questionnaire.

PDF Content:
{pdf_content.full_text[:30000]}

Questions to answer (row number: question):
{fields_text}

Return JSON array with answers:
[
  {{"_row": 4, "Answer": "Okta, Inc."}},
  {{"_row": 5, "Answer": "SOC1 Type II Report"}},
  ...
]

Return ONLY valid JSON array. No markdown."""
            else:
                # Generic extraction
                prompt = f"""Extract data from this SOC1 report for sheet "{sheet_name}".

Columns: {', '.join(headers[:10])}

PDF Content:
{pdf_content.full_text[:40000]}

Return JSON array:
[{{"_row": {header_row + 1}, "{headers[0]}": "value"}}]

Return ONLY JSON."""

            try:
                response = self._generate(prompt, max_tokens=32768, expect_json=True)
                print(f"  Response length: {len(response)} chars")
                print(f"  Response preview: {response[:300]}...")
                
                rows = self._parse_json_response(response)
                
                # Handle both array and object responses
                if isinstance(rows, list):
                    result[sheet_name] = rows
                    print(f"  Extracted {len(rows)} rows")
                elif isinstance(rows, dict):
                    if sheet_name in rows:
                        result[sheet_name] = rows[sheet_name]
                    else:
                        result[sheet_name] = [rows] if "_row" in rows else []
            except Exception as e:
                print(f"  Failed: {e}")
                import traceback
                traceback.print_exc()
                result[sheet_name] = []
        
        return result

    def analyze_for_gaps(
        self,
        pdf_content: ExtractedPDFContent,
        filled_data: dict[str, list[dict[str, Any]]],
    ) -> dict[str, Any]:
        """
        Analyze the filled data for gaps or issues that need attention.

        Args:
            pdf_content: The original PDF content
            filled_data: The data that was filled into the template

        Returns:
            Dict containing analysis results and recommendations
        """
        prompt = f"""Analyze the following SOC1 Type II report data extraction for completeness and accuracy.

## Extracted Data (filled into management review template):
{json.dumps(filled_data, indent=2, default=str)[:20000]}

## Original Report Summary (first 10000 chars):
{pdf_content.full_text[:10000]}

## Analysis Required:
1. Identify any controls mentioned in the report that may not have been captured
2. Flag any exceptions or findings that need management attention
3. Note any missing information that should be followed up on
4. Provide a summary of the overall control environment
5. Count cells marked with low/medium confidence that need review
6. Analyze the Complementary User Entity Controls (CUECs) extraction

Return a JSON object:
{{
    "total_controls_identified": <number>,
    "controls_with_exceptions": <number>,
    "total_cuecs_identified": <number>,
    "cells_needing_review": {{
        "low_confidence": <number of cells with low confidence>,
        "medium_confidence": <number of cells with medium confidence>
    }},
    "missing_information": ["list of missing items"],
    "key_findings": ["list of key findings"],
    "cuec_findings": ["list of findings related to CUECs"],
    "recommendations": ["list of recommendations"],
    "summary": "brief summary of the SOC1 report"
}}

Return ONLY the JSON object."""

        response_text = self._generate(prompt, max_tokens=4096, expect_json=True)

        try:
            return self._parse_json_response(response_text)
        except (json.JSONDecodeError, ValueError):
            return {"error": "Could not parse analysis", "raw": response_text[:1000]}


async def process_soc1_documents(
    type_ii_path: Path,
    management_review_path: Path,
    output_dir: Path | None = None,
    api_key: str | None = None,
) -> dict[str, Any]:
    """
    Main processing function for SOC1 Type II documents.

    Args:
        type_ii_path: Path to the Type II report PDF
        management_review_path: Path to the management review Excel template
        output_dir: Directory to save output files (defaults to same as input)
        api_key: Optional Google API key

    Returns:
        Dict containing:
            - output_path: Path to the filled Excel file
            - analysis: Analysis results and recommendations
            - status: Processing status
    """
    if output_dir is None:
        output_dir = management_review_path.parent

    output_dir.mkdir(parents=True, exist_ok=True)

    log_memory("start of process_soc1_documents")

    # Step 1: Extract PDF content
    print(f"Extracting content from PDF: {type_ii_path}")
    pdf_content = PDFExtractor.extract(type_ii_path)
    print(f"  - Extracted {len(pdf_content.pages)} pages")
    print(f"  - Found {len(pdf_content.tables)} tables")
    log_memory("after PDF extraction")

    # Step 2: Read Excel template
    print(f"Reading Excel template: {management_review_path}")
    workbook, template = ExcelHandler.read_template(management_review_path)
    print(f"  - Found sheets: {template.sheet_names}")
    for sheet, headers in template.headers.items():
        print(f"  - {sheet}: {len(headers)} columns")
    log_memory("after Excel template load")

    # Step 3: Initialize AI agent and process
    print("Initializing Google Gemini AI agent...")
    agent = SOC1Agent(api_key=api_key)

    print("Extracting and mapping content using AI...")
    mappings = agent.extract_and_map(pdf_content, template)
    log_memory("after AI extraction")

    # Step 4: Fill the template
    output_filename = f"filled_{management_review_path.name}"
    output_path = output_dir / output_filename

    print("Filling Excel template...")
    ExcelHandler.fill_template(workbook, template, mappings, output_path)
    print(f"  - Saved to: {output_path}")
    log_memory("after workbook save")

    # Free the workbook now that it's saved to disk
    workbook.close()
    del workbook
    gc.collect()
    log_memory("after workbook freed")

    # Step 5: Analyze for gaps (only needs a subset of PDF text)
    # Trim pdf_content to reduce memory before gap analysis
    pdf_metadata = pdf_content.metadata
    template_sheets = template.sheet_names
    pdf_text_for_analysis = pdf_content.full_text[:10000]

    # Free the bulk of PDF content
    pdf_content.full_text = ""
    pdf_content.pages.clear()
    pdf_content.tables.clear()
    del template
    gc.collect()
    log_memory("after freeing PDF content and template")

    print("Analyzing extraction for completeness...")
    # Build a lightweight PDF content object for analysis
    analysis_pdf = ExtractedPDFContent(
        full_text=pdf_text_for_analysis,
        pages=[],
        tables=[],
        metadata=pdf_metadata,
    )
    analysis = agent.analyze_for_gaps(analysis_pdf, mappings)
    del analysis_pdf, pdf_text_for_analysis
    gc.collect()
    log_memory("after gap analysis")

    return {
        "output_path": str(output_path),
        "analysis": analysis,
        "status": "completed",
        "pdf_metadata": pdf_metadata,
        "template_sheets": template_sheets,
    }


# Convenience function for synchronous usage
def process_soc1_sync(
    type_ii_path: str | Path,
    management_review_path: str | Path,
    output_dir: str | Path | None = None,
    api_key: str | None = None,
) -> dict[str, Any]:
    """
    Synchronous wrapper for process_soc1_documents.

    Args:
        type_ii_path: Path to the Type II report PDF
        management_review_path: Path to the management review Excel template
        output_dir: Directory to save output files
        api_key: Optional Google API key

    Returns:
        Dict containing processing results
    """
    import asyncio

    type_ii_path = Path(type_ii_path)
    management_review_path = Path(management_review_path)
    output_dir = Path(output_dir) if output_dir else None

    return asyncio.run(
        process_soc1_documents(
            type_ii_path,
            management_review_path,
            output_dir,
            api_key,
        )
    )
