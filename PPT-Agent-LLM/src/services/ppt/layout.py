import logging
from pptx.util import Inches, Pt
from src.services.llm.provider import get_llm_provider
from langchain_core.messages import SystemMessage, HumanMessage

class LayoutEngine:
    """
    Intelligent layout engine for PowerPoint slides.
    Handles overflow detection, dynamic spacing, and content fitting.
    """
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.llm = get_llm_provider()
        
        # Constants for estimation (calibrated for Arial 10pt with extra spacing)
        # Constants for estimation (calibrated for Arial 10pt with extra spacing)
        # Constants for estimation (calibrated for Arial 10pt with extra spacing)
        # Constants for estimation (calibrated for Arial 10pt with bullet spacing)
        self.CHARS_PER_LINE_PER_INCH = 14  # Conservative for bullet-heavy content
        self.LINE_HEIGHT_INCHES = 0.42     # Accounts for bullet spacing overhead
        self.PADDING_INCHES = 0.02         # Minimal padding

    def estimate_row_height(self, row_data: list, col_widths: list) -> float:
        """
        Estimate the height of a table row based on its content.
        
        Args:
            row_data: List of strings (cell values)
            col_widths: List of column widths in inches
            
        Returns:
            Estimated height in inches
        """
        max_height = 0.0
        
        for text, width in zip(row_data, col_widths):
            if not text:
                continue
                
            # Calculate lines needed
            # Simple heuristic: text_length / (width * chars_per_inch)
            chars_capacity = width * self.CHARS_PER_LINE_PER_INCH
            lines = max(1, len(text) / max(1, chars_capacity))
            
            # Round up for safety
            import math
            lines = math.ceil(lines)
            
            # Calculate height
            height = (lines * self.LINE_HEIGHT_INCHES) + self.PADDING_INCHES
            max_height = max(max_height, height)
            
        return max_height

    def check_overflow(self, current_table_height: float, next_row_height: float, max_slide_height: float, bottom_margin: float = 1.0) -> bool:
        """
        Check if adding the next row will cause overflow.
        
        Args:
            current_table_height: Current height of the table in inches
            next_row_height: Estimated height of the next row
            max_slide_height: Total slide height in inches
            bottom_margin: Minimum margin required at bottom in inches
            
        Returns:
            True if overflow is detected, False otherwise
        """
        available_space = max_slide_height - bottom_margin
        projected_height = current_table_height + next_row_height
        
        self.logger.debug(f"    Check Overflow: Cur={current_table_height:.2f} + Next={next_row_height:.2f} = {projected_height:.2f} vs Max={available_space:.2f}")
        
        if projected_height > available_space:
            self.logger.info(f"  Overflow detected: Projected {projected_height:.2f}\" > Available {available_space:.2f}\"")
            return True
            
        return False

    def check_horizontal_overflow(self, col_widths: list, slide_width: float, side_margins: float = 1.0) -> bool:
        """
        Check if the table columns exceed the slide width.
        
        Args:
            col_widths: List of column widths in inches
            slide_width: Total slide width in inches
            side_margins: Total side margins (left + right) in inches
            
        Returns:
            True if overflow is detected, False otherwise
        """
        total_width = sum(col_widths)
        available_width = slide_width - side_margins
        
        self.logger.debug(f"    Check Horizontal Overflow: Total={total_width:.2f}\" vs Available={available_width:.2f}\"")
        
        if total_width > available_width:
            self.logger.warning(f"  Horizontal Overflow detected: Table width {total_width:.2f}\" > Available {available_width:.2f}\"")
            return True
            
        return False

    def analyze_split_point(self, rows_fitting: list, next_row: str) -> int:
        """
        Analyze if the split point is semantically sound.
        
        Args:
            rows_fitting: List of row texts that fit on the current slide
            next_row: Text of the next row that causes overflow
            
        Returns:
            Index of the last row to keep on the current slide (1-based).
            Usually len(rows_fitting), but might be less if we should push rows to next slide.
        """
        if not rows_fitting:
            return 0
            
        last_fitting_row = rows_fitting[-1]
        
        # Build prompt
        system = """You are an expert at document layout.
TASK: Determine the best place to split a table across two slides.
GOAL: Avoid separating related items or breaking semantic groups.

INPUT:
- Last item fitting on Slide 1
- First item overflowing to Slide 2

DECISION:
- If the items are strongly related (e.g., part of the same sentence, same date group), move the Last Item to Slide 2.
- Otherwise, keep the split as is.

Return JSON: {"keep_last_item_on_slide_1": boolean, "reasoning": string}"""

        user = f"""Analyze this split point:

Last item on Slide 1: "{last_fitting_row}"
First item on Slide 2: "{next_row}"

Should we keep the Last Item on Slide 1?
Return JSON."""

        try:
            messages = [
                SystemMessage(content=system),
                HumanMessage(content=user)
            ]
            
            # Simple structure for response
            from pydantic import BaseModel, Field
            class SplitDecision(BaseModel):
                keep_last_item_on_slide_1: bool = Field(..., description="True to keep split as is, False to move last item to next slide")
                reasoning: str
            
            response = self.llm.invoke_structured(messages, SplitDecision)
            
            if isinstance(response, dict):
                response = SplitDecision(**response)
                
            self.logger.info(f"  AI Split Decision: Keep={response.keep_last_item_on_slide_1}, Reason={response.reasoning}")
            
            if response.keep_last_item_on_slide_1:
                return len(rows_fitting)
            else:
                return len(rows_fitting) - 1
                
        except Exception as e:
            self.logger.warning(f"AI split analysis failed: {e}, keeping default split")
            return len(rows_fitting)
