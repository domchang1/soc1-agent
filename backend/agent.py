"""
SOC1 Type II Report Processing Agent

This module provides functionality to:
1. Extract text and tables from PDF Type II reports
2. Read Excel management review templates
3. Use Google Gemini AI to intelligently map PDF content to Excel fields
4. Return a filled-out Excel management review
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from google import genai
from google.genai import types
import openpyxl
import pdfplumber
from dotenv import load_dotenv
from openpyxl.styles import PatternFill

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
        pages: list[str] = []
        tables: list[list[list[str]]] = []
        metadata: dict[str, Any] = {}

        with pdfplumber.open(pdf_path) as pdf:
            metadata = {
                "num_pages": len(pdf.pages),
                "metadata": pdf.metadata or {},
            }

            for page in pdf.pages:
                # Extract text from page
                page_text = page.extract_text() or ""
                pages.append(page_text)

                # Extract tables from page
                page_tables = page.extract_tables() or []
                for table in page_tables:
                    # Clean up table data
                    cleaned_table = [
                        [str(cell) if cell is not None else "" for cell in row]
                        for row in table
                    ]
                    tables.append(cleaned_table)

        full_text = "\n\n--- Page Break ---\n\n".join(pages)

        return ExtractedPDFContent(
            full_text=full_text,
            pages=pages,
            tables=tables,
            metadata=metadata,
        )


class ExcelHandler:
    """Handles reading and writing Excel files using openpyxl."""

    @staticmethod
    def read_template(excel_path: Path) -> tuple[openpyxl.Workbook, ExcelTemplate]:
        """
        Read an Excel template and extract its structure.

        Args:
            excel_path: Path to the Excel file

        Returns:
            Tuple of (workbook, ExcelTemplate)
        """
        wb = openpyxl.load_workbook(excel_path)
        sheet_names = wb.sheetnames
        headers: dict[str, list[str]] = {}
        header_to_col: dict[str, dict[str, int]] = {}
        header_rows: dict[str, int] = {}
        structure: dict[str, list[dict[str, Any]]] = {}

        for sheet_name in sheet_names:
            ws = wb[sheet_name]
            sheet_headers: list[str] = []
            sheet_header_to_col: dict[str, int] = {}
            sheet_structure: list[dict[str, Any]] = []

            # Find headers (assume first row with content)
            found_header_row = 1
            for row_idx in range(1, min(10, ws.max_row + 1)):  # Check first 10 rows
                row_values = [
                    ws.cell(row=row_idx, column=col).value
                    for col in range(1, ws.max_column + 1)
                ]
                non_empty = [v for v in row_values if v is not None]
                if len(non_empty) >= 2:  # Found a row with multiple values
                    found_header_row = row_idx
                    for col_idx, v in enumerate(row_values, 1):
                        header_name = str(v) if v else f"Column_{col_idx}"
                        sheet_headers.append(header_name)
                        # Map both original and normalized (lowercase, stripped) header names
                        sheet_header_to_col[header_name] = col_idx
                        sheet_header_to_col[header_name.lower().strip()] = col_idx
                    break

            headers[sheet_name] = sheet_headers
            header_to_col[sheet_name] = sheet_header_to_col
            header_rows[sheet_name] = found_header_row

            # Extract existing data structure (for context to AI)
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
        print(f"fill_template called with {len(mappings)} sheets")
        print(f"Available sheets in workbook: {wb.sheetnames}")
        print(f"Mappings keys: {list(mappings.keys())}")
        
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

            print(f"Processing sheet '{sheet_name}' -> actual: '{actual_sheet_name}' with {len(rows)} rows")
            ws = wb[actual_sheet_name]
            # Use the actual sheet name for header lookup
            sheet_name = actual_sheet_name
            header_map = template.header_to_col.get(sheet_name, {})
            header_row = template.header_row.get(sheet_name, 1)
            print(f"  Header row: {header_row}, columns: {len(header_map)}")

            for row_idx_in, row_update in enumerate(rows):
                row_idx = row_update.get("_row")
                if row_idx is None:
                    print(f"    Row {row_idx_in}: No _row field, skipping")
                    continue

                # Ensure row_idx is after header row
                if row_idx <= header_row:
                    row_idx = header_row + 1
                
                cells_written = 0

                # Get confidence levels for this row (supports multiple formats)
                # Format 1: "_confidence": {"col_name": "high", ...}
                # Format 2: "_c": {"col_name": "h", ...} (abbreviated)
                # Format 3: "col_name": {"value": "x", "confidence": "high"}
                confidence_map = row_update.get("_confidence", row_update.get("_c", {}))
                row_confidence = row_update.get("_row_confidence", "high")
                
                # Normalize abbreviated confidence values
                def normalize_confidence(c):
                    if c in ("h", "high"):
                        return "high"
                    elif c in ("m", "medium", "med"):
                        return "medium"
                    elif c in ("l", "low"):
                        return "low"
                    return "high"  # default

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
                        # Mark empty cells with low confidence
                        col_idx = header_map.get(col_name) or header_map.get(col_name.lower().strip())
                        if col_idx and confidence_map.get(col_name) == "low":
                            cell = ws.cell(row=row_idx, column=col_idx)
                            cell.fill = FILL_LOW_CONFIDENCE
                        continue

                    # Find the column index for this header
                    col_idx = header_map.get(col_name)
                    if col_idx is None:
                        col_idx = header_map.get(col_name.lower().strip())

                    if col_idx:
                        cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
                        cells_written += 1

                        # Apply confidence-based coloring
                        if cell_confidence == "low":
                            cell.fill = FILL_LOW_CONFIDENCE
                        elif cell_confidence == "medium":
                            cell.fill = FILL_MEDIUM_CONFIDENCE
                        # High confidence = no fill (keep default)
                
                if cells_written > 0:
                    print(f"    Row {row_idx}: Wrote {cells_written} cells")

                # Apply row-level confidence coloring for empty/missing cells
                if row_confidence == "low":
                    for col_idx in range(1, len(template.headers.get(actual_sheet_name, [])) + 1):
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
        # Prepare template structure description
        template_desc = []
        for sheet_name, headers in template.headers.items():
            template_desc.append(f"\nSheet: {sheet_name}")
            template_desc.append(f"Columns: {', '.join(headers)}")

            # Show sample of existing data
            if template.structure.get(sheet_name):
                sample_rows = template.structure[sheet_name][:3]
                template_desc.append("Sample existing rows:")
                for row in sample_rows:
                    row_vals = {
                        k: v.get("value") if isinstance(v, dict) else v
                        for k, v in row.items()
                        if not k.startswith("_")
                    }
                    template_desc.append(f"  {row_vals}")

        template_info = "\n".join(template_desc)

        # Prepare tables info if any
        tables_info = ""
        if pdf_content.tables:
            tables_info = "\n\nExtracted Tables from PDF:\n"
            for i, table in enumerate(pdf_content.tables[:10], 1):  # Limit to 10 tables
                tables_info += f"\nTable {i}:\n"
                for row in table[:20]:  # Limit rows per table
                    tables_info += f"  {row}\n"

        # Find sheets that match management review and CUEC patterns
        management_sheet = None
        cuec_sheet = None
        for sheet_name in template.sheet_names:
            sheet_lower = sheet_name.lower()
            if "management review" in sheet_lower:
                management_sheet = sheet_name
            if "user entity" in sheet_lower or "cuec" in sheet_lower or "comp user" in sheet_lower:
                cuec_sheet = sheet_name

        # Get column names for the key sheets
        mgmt_cols = template.headers.get(management_sheet, [])[:10] if management_sheet else []
        cuec_cols = template.headers.get(cuec_sheet, [])[:7] if cuec_sheet else []

        return f"""Extract SOC1 data to fill Excel template. Be CONCISE.

SHEETS TO FILL:
1. "{management_sheet}" - columns: {', '.join(mgmt_cols[:8])}...
2. "{cuec_sheet}" - columns: {', '.join(cuec_cols)}

PDF CONTENT:
{pdf_content.full_text[:60000]}

Return ONLY valid JSON. No markdown, no commentary.
{{
  "{management_sheet}": [
    {{"_row": 2, "_c": {{"col1": "h"}}, "col1": "value", "col2": "value"}},
    {{"_row": 3, "_c": {{"col1": "m"}}, "col1": "value"}}
  ],
  "{cuec_sheet}": [
    {{"_row": 2, "col1": "CUEC description"}}
  ]
}}

CONFIDENCE in "_c" field: "h"=high (found), "m"=medium (inferred), "l"=low (not found)

RULES:
- Use EXACT column names from template
- Keep values SHORT (max 100 chars)
- One row per control/CUEC
- Skip columns with no data
- For CUECs: find "Complementary User Entity Controls" section"""

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
            # Use larger token limit to avoid truncation
            response_text = self._generate(prompt, max_tokens=65536, expect_json=True)
            return self._parse_json_response(response_text)
        except Exception as e:
            print(f"Main extraction failed: {e}")
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
            if not headers:
                continue
            
            prompt = f"""Extract data from this SOC1 report to fill the Excel sheet "{sheet_name}".

Sheet columns: {', '.join(headers[:15])}  

PDF Content (excerpt):
{pdf_content.full_text[:40000]}

Return a JSON array of rows. Each row has "_row" (starting at 2) and column values.
Keep values SHORT and concise. Example:
[
  {{"_row": 2, "{headers[0] if headers else 'Col1'}": "value1", "{headers[1] if len(headers) > 1 else 'Col2'}": "value2"}},
  {{"_row": 3, "{headers[0] if headers else 'Col1'}": "value3"}}
]

Return ONLY the JSON array."""

            try:
                response = self._generate(prompt, max_tokens=16384, expect_json=True)
                rows = self._parse_json_response(response)
                
                # Handle both array and object responses
                if isinstance(rows, list):
                    result[sheet_name] = rows
                elif isinstance(rows, dict):
                    # If it returned an object with the sheet name as key
                    if sheet_name in rows:
                        result[sheet_name] = rows[sheet_name]
                    else:
                        # Wrap single row in array
                        result[sheet_name] = [rows] if "_row" in rows else []
            except Exception as e:
                print(f"Failed to extract sheet {sheet_name}: {e}")
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

    # Step 1: Extract PDF content
    print(f"Extracting content from PDF: {type_ii_path}")
    pdf_content = PDFExtractor.extract(type_ii_path)
    print(f"  - Extracted {len(pdf_content.pages)} pages")
    print(f"  - Found {len(pdf_content.tables)} tables")

    # Step 2: Read Excel template
    print(f"Reading Excel template: {management_review_path}")
    workbook, template = ExcelHandler.read_template(management_review_path)
    print(f"  - Found sheets: {template.sheet_names}")
    for sheet, headers in template.headers.items():
        print(f"  - {sheet}: {len(headers)} columns")

    # Step 3: Initialize AI agent and process
    print("Initializing Google Gemini AI agent...")
    agent = SOC1Agent(api_key=api_key)

    print("Extracting and mapping content using AI...")
    mappings = agent.extract_and_map(pdf_content, template)

    # Step 4: Fill the template
    output_filename = f"filled_{management_review_path.name}"
    output_path = output_dir / output_filename

    print("Filling Excel template...")
    ExcelHandler.fill_template(workbook, template, mappings, output_path)
    print(f"  - Saved to: {output_path}")

    # Step 5: Analyze for gaps
    print("Analyzing extraction for completeness...")
    analysis = agent.analyze_for_gaps(pdf_content, mappings)

    return {
        "output_path": str(output_path),
        "analysis": analysis,
        "status": "completed",
        "pdf_metadata": pdf_content.metadata,
        "template_sheets": template.sheet_names,
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
