"""
Pydantic models for structured LLM outputs.
Ensures type safety and automatic validation of LLM responses.
"""

from typing import Optional, Dict, List
from pydantic import BaseModel, Field, validator


class HeaderMapping(BaseModel):
    """Mapping of PPT header to Excel column(s)."""
    header_name: str = Field(..., description="PowerPoint table header name")
    excel_columns: Optional[List[str]] = Field(
        None, 
        description="List of Excel column names to map to this header. None for auto-generated columns like S.No"
    )
    
    class Config:
        json_schema_extra = {
            "example": {
                "header_name": "Date of Issue",
                "excel_columns": ["IssueDate"]
            }
        }


class TableMapping(BaseModel):
    """Complete mapping specification for a single table."""
    slide_index: int = Field(..., ge=0, description="0-based slide index in the presentation")
    shape_index: int = Field(..., ge=0, description="0-based shape index on the slide")
    headers: Dict[str, Optional[List[str]]] = Field(
        ..., 
        description="Mapping of PPT headers to Excel columns"
    )
    source_filter: Dict[str, str] = Field(
        ..., 
        description="Filter to apply on Excel data (e.g., {'Verticals': 'SEBI'})"
    )
    max_rows: int = Field(default=100, ge=1, le=1000, description="Maximum rows to populate")
    
    @validator('headers')
    def validate_headers(cls, v):
        """Ensure headers dict is not empty."""
        if not v:
            raise ValueError("Headers mapping cannot be empty")
        return v
    
    class Config:
        json_schema_extra = {
            "example": {
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
        }


class MappingResponse(BaseModel):
    """Complete mapping response from LLM."""
    tables: List[TableMapping] = Field(
        ..., 
        min_items=1,
        description="List of table mappings"
    )
    
    @validator('tables')
    def validate_tables(cls, v):
        """Ensure at least one table mapping exists."""
        if not v:
            raise ValueError("At least one table mapping is required")
        return v
    
    def to_dict(self) -> dict:
        """Convert to dictionary format compatible with existing code."""
        return {
            "tables": [
                {
                    "slide_index": table.slide_index,
                    "shape_index": table.shape_index,
                    "headers": table.headers,
                    "source_filter": table.source_filter,
                    "max_rows": table.max_rows
                }
                for table in self.tables
            ]
        }
    
    class Config:
        json_schema_extra = {
            "example": {
                "tables": [
                    {
                        "slide_index": 4,
                        "shape_index": 4,
                        "headers": {
                            "S. No": None,
                            "Date of Issue": ["IssueDate"],
                            "Gist thereof": ["Summary"]
                        },
                        "source_filter": {"Verticals": "SEBI"},
                        "max_rows": 100
                    }
                ]
            }
        }
