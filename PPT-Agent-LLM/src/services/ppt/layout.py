import logging
from pptx.util import Inches, Pt
from src.services.llm.provider import get_llm_provider
from langchain_core.messages import SystemMessage, HumanMessage
from pydantic import BaseModel, Field

class LayoutEngine:
    """
    Intelligent layout engine for PowerPoint slides.
    Handles overflow detection, dynamic spacing, and content fitting.
    """
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self._llm = None
        
        # Calibrated constants for 12pt Arial in PPT tables.
        # Moderate values: fits more rows per slide without visual overflow.
        self.CHARS_PER_LINE_PER_INCH = 13    # ~13 chars/inch accounts for bullet indents
        self.LINE_HEIGHT_INCHES = 0.26       # 12pt font + 1.2 spacing ≈ 0.24", add small buffer
        self.PADDING_INCHES = 0.36           # Cell top (20pt) + bottom (3pt) ≈ 0.32", + buffer
        self.PARA_SPACING_INCHES = 0.20      # space_before(3pt) + space_after(10pt) ≈ 0.18"

    @property
    def llm(self):
        if self._llm is None:
            self._llm = get_llm_provider()
        return self._llm

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
                
            # Split text by newlines so bullet points are calculated individually
            paragraphs = str(text).split('\n')
            valid_paras = [p.strip() for p in paragraphs if p.strip()]
            if not valid_paras:
                continue
                
            total_lines = 0
            chars_capacity = max(1.0, width * self.CHARS_PER_LINE_PER_INCH)
            
            for para in valid_paras:
                # Calculate lines needed for this specific bullet point / paragraph
                para_lines = len(para) / chars_capacity
                import math
                # Use a small tolerance before rounding up to avoid doubling lines on minor wrap
                total_lines += max(1, math.ceil(para_lines - 0.1))
            
            # Paragraph spacing overhead for multi-bullet cells
            para_spacing = (len(valid_paras) - 1) * self.PARA_SPACING_INCHES if len(valid_paras) > 1 else 0.0
            height = (total_lines * self.LINE_HEIGHT_INCHES) + self.PADDING_INCHES + para_spacing
            max_height = max(max_height, height)
            
        return max_height

    def check_overflow(self, current_table_height: float, next_row_height: float, max_slide_height: float, bottom_margin: float = 0.5) -> bool:
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
            
        # Never backtrack if there is only 1 fitting row (would leave the slide empty)
        if len(rows_fitting) <= 1:
            return len(rows_fitting)
            
        last_fitting_row = rows_fitting[-1]
        
        system = """You are an expert document layout and presentation editor.
TASK: You are deciding how to paginate a table across slides.
Slide 1 currently contains a list of items that fit within the slide height.
The next item does NOT fit on Slide 1 and will start on Slide 2.

RULE:
- By default, keep all items currently fitting on Slide 1 (do NOT backtrack / do NOT move fitting items to Slide 2).
- ONLY move the last item from Slide 1 to Slide 2 if the last item on Slide 1 and the first item on Slide 2 are STRICTLY dependent on each other (e.g., part 1 and part 2 of the same single announcement, or a direct continuation of text that cannot stand alone).
- If the items are independent circulars, separate news, or distinct regulatory updates, keep the split as is (do NOT move).

Return JSON:
{
  "should_move_last_item_to_next_slide": boolean,
  "reasoning": string
}"""

        user = f"""Analyze this split point:

Last item on Slide 1: "{last_fitting_row}"
First item on Slide 2: "{next_row}"

Are these items strictly coupled such that the last item on Slide 1 MUST be moved to Slide 2?
Return JSON."""

        try:
            class SplitDecision(BaseModel):
                should_move_last_item_to_next_slide: bool = Field(
                    ..., 
                    description="True ONLY if the two items are strongly coupled and must stay together on Slide 2. False if they are independent items."
                )
                reasoning: str
            
            messages = [
                SystemMessage(content=system),
                HumanMessage(content=user)
            ]
            
            response = self.llm.invoke_structured(messages, SplitDecision)
            
            if isinstance(response, dict):
                response = SplitDecision(**response)
                
            self.logger.info(f"  AI Split Decision: MoveLastToNext={response.should_move_last_item_to_next_slide}, Reason={response.reasoning}")
            
            if response.should_move_last_item_to_next_slide and len(rows_fitting) > 1:
                return len(rows_fitting) - 1
            else:
                return len(rows_fitting)
                
        except Exception as e:
            self.logger.warning(f"AI split analysis failed: {e}, keeping default split")
            return len(rows_fitting)
