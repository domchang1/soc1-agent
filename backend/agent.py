"""
SOC1 Type II Report Processing Agent

This module provides functionality to:
1. Extract text and tables from PDF Type II reports
2. Read Excel management review templates
3. Use AI to intelligently map PDF content to Excel fields
4. Return a filled-out Excel management review

Supports multiple AI providers:
- Google Gemini (free tier available)
- Anthropic Claude
"""

from __future__ import annotations

import json
import os
import re
from abc import ABC, abstractmethod
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any

import openpyxl
import pdfplumber


class AIProvider(Enum):
    """Supported AI providers."""
    GEMINI = "gemini"
    ANTHROPIC = "anthropic"


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
        Fill the Excel template with mapped data.

        Args:
            wb: The workbook to fill
            template: The ExcelTemplate with header mappings
            mappings: Dict of sheet_name -> list of row updates
            output_path: Path to save the filled template

        Returns:
            Path to the saved file
        """
        for sheet_name, rows in mappings.items():
            if sheet_name not in wb.sheetnames:
                continue

            ws = wb[sheet_name]
            header_map = template.header_to_col.get(sheet_name, {})
            header_row = template.header_row.get(sheet_name, 1)

            for row_update in rows:
                row_idx = row_update.get("_row")
                if row_idx is None:
                    continue

                # Ensure row_idx is after header row
                if row_idx <= header_row:
                    row_idx = header_row + 1

                for col_name, value in row_update.items():
                    if col_name.startswith("_"):
                        continue

                    # Skip if value is a dict (metadata) or None
                    if isinstance(value, dict) or value is None:
                        continue

                    # Find the column index for this header
                    # Try exact match first, then normalized match
                    col_idx = header_map.get(col_name)
                    if col_idx is None:
                        col_idx = header_map.get(col_name.lower().strip())

                    if col_idx:
                        ws.cell(row=row_idx, column=col_idx, value=value)

        wb.save(output_path)
        return output_path


class BaseAIClient(ABC):
    """Abstract base class for AI clients."""

    @abstractmethod
    def generate(self, prompt: str, max_tokens: int = 8192) -> str:
        """Generate a response from the AI model."""
        pass


class GeminiClient(BaseAIClient):
    """Google Gemini AI client (free tier available)."""

    def __init__(self, api_key: str | None = None):
        import google.generativeai as genai

        self.api_key = api_key or os.environ.get("GOOGLE_API_KEY")
        if not self.api_key:
            raise ValueError(
                "Google API key required. Set GOOGLE_API_KEY environment variable "
                "or pass api_key parameter.\n"
                "Get a free key at: https://aistudio.google.com/apikey"
            )
        genai.configure(api_key=self.api_key)
        # Use Gemini 1.5 Flash for free tier (fast and capable)
        self.model = genai.GenerativeModel("gemini-1.5-flash")

    def generate(self, prompt: str, max_tokens: int = 8192) -> str:
        response = self.model.generate_content(
            prompt,
            generation_config={
                "max_output_tokens": max_tokens,
                "temperature": 0.1,  # Low temperature for accuracy
            },
        )
        return response.text


class AnthropicClient(BaseAIClient):
    """Anthropic Claude AI client."""

    def __init__(self, api_key: str | None = None):
        import anthropic

        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
        if not self.api_key:
            raise ValueError(
                "Anthropic API key required. Set ANTHROPIC_API_KEY environment variable "
                "or pass api_key parameter."
            )
        self.client = anthropic.Anthropic(api_key=self.api_key)
        self.model_name = "claude-sonnet-4-20250514"

    def generate(self, prompt: str, max_tokens: int = 8192) -> str:
        response = self.client.messages.create(
            model=self.model_name,
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}],
        )
        return response.content[0].text


def get_ai_client(provider: AIProvider | str | None = None, api_key: str | None = None) -> BaseAIClient:
    """
    Factory function to get the appropriate AI client.

    Auto-detects provider based on available API keys if not specified.

    Args:
        provider: The AI provider to use (gemini, anthropic)
        api_key: Optional API key (uses env vars if not provided)

    Returns:
        An AI client instance
    """
    if isinstance(provider, str):
        provider = AIProvider(provider.lower())

    # Auto-detect provider based on available API keys
    if provider is None:
        if os.environ.get("GOOGLE_API_KEY"):
            provider = AIProvider.GEMINI
        elif os.environ.get("ANTHROPIC_API_KEY"):
            provider = AIProvider.ANTHROPIC
        else:
            raise ValueError(
                "No AI API key found. Please set one of:\n"
                "  - GOOGLE_API_KEY (free tier at https://aistudio.google.com/apikey)\n"
                "  - ANTHROPIC_API_KEY"
            )

    if provider == AIProvider.GEMINI:
        return GeminiClient(api_key)
    elif provider == AIProvider.ANTHROPIC:
        return AnthropicClient(api_key)
    else:
        raise ValueError(f"Unknown provider: {provider}")


class SOC1Agent:
    """
    AI-powered agent for processing SOC1 Type II reports.

    Supports multiple AI providers (Gemini, Anthropic).
    """

    def __init__(
        self,
        provider: AIProvider | str | None = None,
        api_key: str | None = None,
    ):
        """
        Initialize the SOC1 Agent.

        Args:
            provider: AI provider to use ('gemini' or 'anthropic'). Auto-detects if not specified.
            api_key: API key for the provider. Uses environment variables if not provided.
        """
        self.client = get_ai_client(provider, api_key)

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

        return f"""You are an expert at analyzing SOC1 Type II audit reports and extracting relevant information to fill management review templates.

## Task
Analyze the following SOC1 Type II report content and extract information to fill the management review Excel template.

## Excel Template Structure
{template_info}

## SOC1 Type II Report Content
{pdf_content.full_text[:50000]}
{tables_info}

## Instructions
1. Analyze the SOC1 Type II report to identify:
   - Control objectives
   - Control activities/descriptions
   - Testing procedures performed
   - Test results (effective/exception)
   - Any exceptions or deviations noted
   - Management responses (if any)
   - Auditor conclusions

2. Map the extracted information to the Excel template columns. Use the EXACT column names from the template.

3. Return a JSON object with the following structure:
{{
    "Sheet Name": [
        {{
            "_row": <row_number starting from 2>,
            "Exact Column Name 1": "value to fill",
            "Exact Column Name 2": "value to fill",
            ...
        }},
        ...
    ]
}}

IMPORTANT:
- Use the EXACT sheet names and column names from the template (case-sensitive)
- Row numbers should start from 2 (row 1 is headers)
- Create one entry per control/finding identified in the report

4. For each piece of information:
   - Be accurate and faithful to the source document
   - Use direct quotes where appropriate for descriptions and findings
   - If information is not found for a column, omit that column from the row
   - Match the format/style of any existing data in the template

5. Focus on extracting these typical SOC1 management review elements:
   - Control ID/Number
   - Control Description/Objective
   - Control Owner/Responsible Party
   - Test Procedure Performed
   - Sample Size
   - Test Result (Effective/Exception/N/A)
   - Exception Details (if any)
   - Management Response
   - Remediation Status
   - Review Date
   - Reviewer Comments/Notes

Return ONLY the JSON object, no additional text or markdown formatting."""

    def _parse_json_response(self, response_text: str) -> dict[str, Any]:
        """Parse JSON from AI response, handling markdown code blocks."""
        # Handle potential markdown code blocks
        json_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", response_text)
        if json_match:
            response_text = json_match.group(1)

        try:
            return json.loads(response_text)
        except json.JSONDecodeError:
            # Try to find JSON object in response
            json_start = response_text.find("{")
            json_end = response_text.rfind("}") + 1
            if json_start >= 0 and json_end > json_start:
                return json.loads(response_text[json_start:json_end])
            else:
                raise ValueError(f"Could not parse AI response as JSON: {response_text[:500]}")

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
        prompt = self._create_extraction_prompt(pdf_content, template)
        response_text = self.client.generate(prompt, max_tokens=8192)
        return self._parse_json_response(response_text)

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

Return a JSON object:
{{
    "total_controls_identified": <number>,
    "controls_with_exceptions": <number>,
    "missing_information": ["list of missing items"],
    "key_findings": ["list of key findings"],
    "recommendations": ["list of recommendations"],
    "summary": "brief summary of the SOC1 report"
}}

Return ONLY the JSON object."""

        response_text = self.client.generate(prompt, max_tokens=4096)

        try:
            return self._parse_json_response(response_text)
        except (json.JSONDecodeError, ValueError):
            return {"error": "Could not parse analysis", "raw": response_text[:1000]}


async def process_soc1_documents(
    type_ii_path: Path,
    management_review_path: Path,
    output_dir: Path | None = None,
    provider: AIProvider | str | None = None,
    api_key: str | None = None,
) -> dict[str, Any]:
    """
    Main processing function for SOC1 Type II documents.

    Args:
        type_ii_path: Path to the Type II report PDF
        management_review_path: Path to the management review Excel template
        output_dir: Directory to save output files (defaults to same as input)
        provider: AI provider to use ('gemini' or 'anthropic'). Auto-detects if not specified.
        api_key: Optional API key for the AI provider

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
    print("Initializing AI agent for content mapping...")
    agent = SOC1Agent(provider=provider, api_key=api_key)

    print("Extracting and mapping content using AI...")
    mappings = agent.extract_and_map(pdf_content, template)

    # Step 4: Fill the template
    output_filename = f"filled_{management_review_path.name}"
    output_path = output_dir / output_filename

    print(f"Filling Excel template...")
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
    provider: AIProvider | str | None = None,
    api_key: str | None = None,
) -> dict[str, Any]:
    """
    Synchronous wrapper for process_soc1_documents.

    Args:
        type_ii_path: Path to the Type II report PDF
        management_review_path: Path to the management review Excel template
        output_dir: Directory to save output files
        provider: AI provider ('gemini' or 'anthropic'). Auto-detects if not specified.
        api_key: Optional API key for the provider

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
            provider,
            api_key,
        )
    )
