"""
Pydantic models for AI text processing in Phase 2.
Ensures type safety and validation for bullet generation, formatting, and text cleaning.
"""

from typing import List, Optional
from pydantic import BaseModel, Field, validator


class BulletPoint(BaseModel):
    """Individual bullet point with metadata."""
    text: str = Field(..., min_length=5, max_length=200)
    importance: int = Field(default=3, ge=1, le=5, description="Importance score 1-5")
    length: int = Field(..., ge=0, description="Character count")
    
    @validator('length', always=True)
    def set_length(cls, v, values):
        """Auto-calculate length from text."""
        if 'text' in values:
            return len(values['text'])
        return v


class BulletGenerationResponse(BaseModel):
    """Response from AI bullet generation."""
    bullets: List[str] = Field(
        ..., 
        min_items=1, 
        max_items=7,
        description="Generated bullet points"
    )
    summary_quality: int = Field(
        ..., 
        ge=1, 
        le=10,
        description="AI's self-assessment of quality (1-10)"
    )
    removed_quotes: bool = Field(
        default=False,
        description="Whether quotes were removed from text"
    )
    original_length: int = Field(
        default=0,
        description="Original text length in characters"
    )
    
    @validator('bullets')
    def validate_bullets(cls, v):
        """Ensure bullets are not empty and properly formatted."""
        cleaned = []
        for bullet in v:
            bullet = bullet.strip()
            # Remove bullet symbols if AI added them
            bullet = bullet.lstrip('•-*').strip()
            if bullet:
                cleaned.append(bullet)
        
        if not cleaned:
            raise ValueError("At least one valid bullet point required")
        
        return cleaned
    
    class Config:
        json_schema_extra = {
            "example": {
                "bullets": [
                    "Company to launch new product in Q2",
                    "Revenue grew 15% year-over-year"
                ],
                "summary_quality": 8,
                "removed_quotes": True,
                "original_length": 120
            }
        }


class FormattingDecision(BaseModel):
    """AI decision for text formatting parameters."""
    font_size: int = Field(
        ..., 
        ge=8, 
        le=14,
        description="Font size in points"
    )
    line_spacing_before: int = Field(
        ..., 
        ge=3, 
        le=10,
        description="Space before paragraph in points"
    )
    line_spacing_after: int = Field(
        ..., 
        ge=2, 
        le=8,
        description="Space after paragraph in points"
    )
    margin_left: int = Field(
        default=5,
        ge=3, 
        le=15,
        description="Left margin in points"
    )
    margin_right: int = Field(
        default=5,
        ge=3, 
        le=15,
        description="Right margin in points"
    )
    reasoning: str = Field(
        ...,
        max_length=200,
        description="Brief explanation of formatting choices"
    )
    
    @validator('font_size')
    def validate_font_size(cls, v):
        """Prefer 10-12pt range for readability."""
        if v < 9:
            raise ValueError("Font size too small for readability")
        return v
    
    class Config:
        json_schema_extra = {
            "example": {
                "font_size": 10,
                "line_spacing_before": 6,
                "line_spacing_after": 3,
                "margin_left": 5,
                "margin_right": 5,
                "reasoning": "Moderate content density allows 10pt font with standard spacing"
            }
        }


class TextCleaningResponse(BaseModel):
    """Response from AI text cleaning."""
    cleaned_text: str = Field(..., description="Text after cleaning")
    removed_quotes: List[str] = Field(
        default_factory=list,
        description="List of quote characters removed"
    )
    removed_artifacts: List[str] = Field(
        default_factory=list,
        description="List of other artifacts removed (e.g., extra spaces)"
    )
    confidence: float = Field(
        ..., 
        ge=0.0, 
        le=1.0,
        description="Confidence in cleaning quality (0-1)"
    )
    changes_made: int = Field(
        default=0,
        ge=0,
        description="Number of changes made to text"
    )
    
    class Config:
        json_schema_extra = {
            "example": {
                "cleaned_text": "Regulatory changes require compliance updates",
                "removed_quotes": ['"', '"'],
                "removed_artifacts": ["  ", "\\n"],
                "confidence": 0.95,
                "changes_made": 3
            }
        }
