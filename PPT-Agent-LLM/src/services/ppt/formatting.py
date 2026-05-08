"""
AI-powered formatting engine for Phase 2.
Makes intelligent decisions about font size, spacing, and layout based on content.
"""

import logging
from tenacity import retry, stop_after_attempt, wait_exponential

from langchain_core.messages import SystemMessage, HumanMessage

from src.services.llm.provider import get_llm_provider
from src.models.text import FormattingDecision


class AIFormattingEngine:
    """AI-powered formatting decisions based on content analysis."""
    
    def __init__(self):
        self.llm = get_llm_provider()
        self.logger = logging.getLogger(__name__)
    
    @retry(
        stop=stop_after_attempt(2),
        wait=wait_exponential(multiplier=1, min=1, max=5)
    )
    def determine_formatting(
        self,
        text: str,
        num_bullets: int,
        cell_width: float,
        cell_height: float
    ) -> FormattingDecision:
        """
        AI decides optimal formatting based on content and constraints.
        
        Args:
            text: Full text content
            num_bullets: Number of bullet points
            cell_width: Cell width in inches
            cell_height: Cell height in inches
            
        Returns:
            FormattingDecision with font size, spacing, and reasoning
        """
        # Calculate content metrics
        total_chars = len(text)
        avg_bullet_length = total_chars / max(num_bullets, 1)
        cell_area = cell_width * cell_height
        
        self.logger.debug(
            f"Formatting analysis: {num_bullets} bullets, "
            f"{total_chars} chars, {cell_area:.2f} sq in"
        )
        
        messages = self._build_formatting_prompt(
            total_chars=total_chars,
            num_bullets=num_bullets,
            avg_bullet_length=avg_bullet_length,
            cell_width=cell_width,
            cell_height=cell_height,
            cell_area=cell_area
        )
        
        try:
            # Try structured output
            response = self.llm.invoke_structured(messages, FormattingDecision)
            
            self.logger.info(
                f"✓ AI formatting: {response.font_size}pt, "
                f"spacing {response.line_spacing_before}/{response.line_spacing_after}pt"
            )
            
            return response
            
        except Exception as e:
            self.logger.warning(f"AI formatting failed: {e}, using smart defaults")
            
            # Fallback to smart defaults based on content density
            return self._calculate_fallback_formatting(
                total_chars, num_bullets, cell_area
            )
    
    def _calculate_fallback_formatting(
        self,
        total_chars: int,
        num_bullets: int,
        cell_area: float
    ) -> FormattingDecision:
        """Calculate reasonable defaults based on content density."""
        
        # Content density: chars per square inch
        density = total_chars / max(cell_area, 1.0)
        
        # Adjust font size based on density
        if density > 100:  # High density
            font_size = 9
            spacing_before = 4
            spacing_after = 2
            reasoning = "High content density requires smaller font"
        elif density > 50:  # Medium density
            font_size = 10
            spacing_before = 5
            spacing_after = 3
            reasoning = "Moderate density allows standard formatting"
        else:  # Low density
            font_size = 11
            spacing_before = 6
            spacing_after = 3
            reasoning = "Low density allows larger, more readable font"
        
        return FormattingDecision(
            font_size=font_size,
            line_spacing_before=spacing_before,
            line_spacing_after=spacing_after,
            margin_left=5,
            margin_right=5,
            reasoning=reasoning
        )
    
    def _build_formatting_prompt(
        self,
        total_chars: int,
        num_bullets: int,
        avg_bullet_length: float,
        cell_width: float,
        cell_height: float,
        cell_area: float
    ) -> list:
        """Build prompt for formatting decisions."""
        
        system = """You are an expert at PowerPoint formatting and typography.

TASK: Recommend optimal formatting to ensure content fits well and is readable.

PRINCIPLES:
1. **Readability First**: Text must be easy to read (prefer 10-12pt)
2. **Fit Content**: All content must fit comfortably in the cell
3. **Visual Balance**: Avoid cramped or sparse appearance
4. **Professional**: Maintain professional presentation standards

GUIDELINES:

Font Size:
- 12pt: Use for short content (< 100 chars) or large cells
- 11pt: Use for moderate content (100-200 chars)
- 10pt: Standard for most content (200-400 chars)
- 9pt: Use only for dense content (> 400 chars)
- 8pt: Avoid unless absolutely necessary

Line Spacing:
- More bullets → less spacing (avoid wasted space)
- Fewer bullets → more spacing (better visual balance)
- Before: 4-8pt (prefer 5-6pt)
- After: 2-4pt (prefer 3pt)

Margins:
- Standard: 5pt left/right
- Dense content: 3-4pt (maximize space)
- Sparse content: 6-8pt (better framing)

EXAMPLES:

Example 1 - Short Content:
Input: 80 chars, 2 bullets, 3.5" × 2" cell
Output: {
  "font_size": 11,
  "line_spacing_before": 6,
  "line_spacing_after": 3,
  "margin_left": 6,
  "margin_right": 6,
  "reasoning": "Short content in adequate space allows larger font and generous spacing"
}

Example 2 - Dense Content:
Input: 450 chars, 5 bullets, 3" × 1.5" cell
Output: {
  "font_size": 9,
  "line_spacing_before": 4,
  "line_spacing_after": 2,
  "margin_left": 4,
  "margin_right": 4,
  "reasoning": "High content density requires compact formatting to fit comfortably"
}

Example 3 - Balanced Content:
Input: 220 chars, 3 bullets, 3.5" × 2" cell
Output: {
  "font_size": 10,
  "line_spacing_before": 5,
  "line_spacing_after": 3,
  "margin_left": 5,
  "margin_right": 5,
  "reasoning": "Moderate content density allows standard formatting with good readability"
}

Return ONLY valid JSON matching the FormattingDecision schema."""

        user = f"""Recommend formatting for this content:

CONTENT METRICS:
- Total characters: {total_chars}
- Number of bullets: {num_bullets}
- Average bullet length: {avg_bullet_length:.0f} characters
- Cell dimensions: {cell_width:.2f}" × {cell_height:.2f}"
- Cell area: {cell_area:.2f} square inches
- Content density: {total_chars / max(cell_area, 1.0):.1f} chars/sq in

Provide optimal formatting recommendations. Return ONLY JSON."""

        return [
            SystemMessage(content=system),
            HumanMessage(content=user)
        ]
