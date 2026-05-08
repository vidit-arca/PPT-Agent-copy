"""
Improved prompt templates with few-shot examples for better LLM performance.
"""

import json
from src.models.domain import MappingResponse


def get_mapping_system_prompt() -> str:
    """
    System prompt for LLM mapping task with clear instructions and examples.
    """
    return """You are an expert at mapping Excel data columns to PowerPoint table headers.

Your task is to analyze the PowerPoint template structure and Excel data profile, then generate a precise mapping that connects Excel columns to PPT table headers.

CRITICAL RULES:
1. Output ONLY valid JSON matching the exact schema provided
2. Do NOT include markdown formatting like ```json
3. Do NOT add explanations or commentary
4. Match Excel column names to PPT headers by semantic meaning, not exact text
5. For auto-generated columns like "S. No", use null as the value
6. Combine multiple Excel columns if they map to one PPT header

EXAMPLE 1 - Simple Date Mapping:
Excel column: "IssueDate"
PPT header: "Date of Issue"
Mapping: {"Date of Issue": ["IssueDate"]}

EXAMPLE 2 - Auto-generated Column:
PPT header: "S. No"
Mapping: {"S. No": null}

EXAMPLE 3 - Multiple Columns to One Header:
Excel columns: "SubCategory", "PDF_URL"
PPT header: "Rules / circulars / Notifications / Order"
Mapping: {"Rules / circulars / Notifications / Order": ["SubCategory", "PDF_URL"]}

EXAMPLE 4 - Summary/Gist Mapping:
Excel column: "Summary"
PPT header: "Gist thereof"
Mapping: {"Gist thereof": ["Summary"]}

EXAMPLE 5 - Title/Contents Mapping:
Excel column: "Title"
PPT header: "Contents thereof"
Mapping: {"Contents thereof": ["Title"]}

COMMON PATTERNS:
- "IssueDate" → "Date of Issue"
- "Summary" → "Gist thereof"
- "Title" → "Contents thereof"
- "SubCategory" + "PDF_URL" → "Rules / circulars / Notifications / Order"
- "S. No" → null (auto-generated)

JSON SCHEMA:
The output must match this exact structure:
{
  "tables": [
    {
      "slide_index": <int>,
      "shape_index": <int>,
      "headers": {
        "<PPT_HEADER_1>": null or ["<EXCEL_COL>"],
        "<PPT_HEADER_2>": ["<EXCEL_COL_1>", "<EXCEL_COL_2>"],
        ...
      },
      "source_filter": {"<EXCEL_FILTER_COL>": "<VALUE>"},
      "max_rows": 100
    }
  ]
}

Remember: Output ONLY the JSON object, nothing else."""


def get_mapping_user_prompt(template_spec: dict, data_profile: dict) -> str:
    """
    User prompt with specific template and data information.
    """
    return f"""POWERPOINT TEMPLATE STRUCTURE:
{json.dumps(template_spec, indent=2)}

EXCEL DATA PROFILE:
{json.dumps(data_profile, indent=2)}

Generate the complete mapping JSON for ALL verticals found in the Excel data.

Output the JSON mapping now:"""


def get_json_schema() -> dict:
    """
    Get JSON schema for structured output (OpenAI function calling).
    """
    return MappingResponse.model_json_schema()


def format_few_shot_examples() -> list:
    """
    Format few-shot examples for better LLM understanding.
    Returns list of example mappings.
    """
    return [
        {
            "description": "SEBI vertical mapping with all standard columns",
            "input": {
                "vertical": "SEBI",
                "excel_columns": ["IssueDate", "SubCategory", "PDF_URL", "Title", "Summary"],
                "ppt_headers": ["S. No", "Date of Issue", "Rules / circulars / Notifications / Order", "Contents thereof", "Gist thereof"]
            },
            "output": {
                "tables": [
                    {
                        "slide_index": 4,
                        "shape_index": 4,
                        "headers": {
                            "S. No": None,
                            "Date of Issue": ["IssueDate"],
                            "Rules / circulars / Notifications / Order": ["SubCategory", "PDF_URL"],
                            "Contents thereof": ["Title"],
                            "Gist thereof": ["Summary"]
                        },
                        "source_filter": {"Verticals": "SEBI"},
                        "max_rows": 100
                    }
                ]
            }
        },
        {
            "description": "RBI vertical mapping",
            "input": {
                "vertical": "RBI",
                "excel_columns": ["IssueDate", "SubCategory", "PDF_URL", "Title", "Summary"],
                "ppt_headers": ["S. No", "Date of Issue", "Rules / circulars / Notifications / Order", "Contents thereof", "Gist thereof"]
            },
            "output": {
                "tables": [
                    {
                        "slide_index": 5,
                        "shape_index": 4,
                        "headers": {
                            "S. No": None,
                            "Date of Issue": ["IssueDate"],
                            "Rules / circulars / Notifications / Order": ["SubCategory", "PDF_URL"],
                            "Contents thereof": ["Title"],
                            "Gist thereof": ["Summary"]
                        },
                        "source_filter": {"Verticals": "RBI"},
                        "max_rows": 100
                    }
                ]
            }
        }
    ]
