"""
AI-powered text processing for Phase 2.
Handles bullet generation, text cleaning, and semantic understanding.
"""

import logging
from typing import List
from tenacity import retry, stop_after_attempt, wait_exponential

from langchain_core.messages import SystemMessage, HumanMessage

from src.services.llm.provider import get_llm_provider
from src.models.text import BulletGenerationResponse, TextCleaningResponse


class AITextProcessor:
    """AI-powered text processing with semantic understanding."""
    
    def __init__(self):
        self.llm = get_llm_provider()
        self.logger = logging.getLogger(__name__)
    
    @retry(
        stop=stop_after_attempt(2),
        wait=wait_exponential(multiplier=1, min=1, max=5)
    )
    def generate_bullets(self, text: str, max_bullets: int = 5) -> BulletGenerationResponse:
        """
        Use AI to create semantic bullet points from text.
        
        Args:
            text: Source text to convert to bullets
            max_bullets: Maximum number of bullets to generate
            
        Returns:
            BulletGenerationResponse with bullets and quality score
        """
        if not text or not text.strip():
            return BulletGenerationResponse(
                bullets=[""],
                summary_quality=1,
                removed_quotes=False,
                original_length=0
            )
        
        self.logger.debug(f"Generating bullets from {len(text)} chars")
        
        messages = self._build_bullet_prompt(text, max_bullets)
        
        try:
            # Try structured output first
            response = self.llm.invoke_structured(messages, BulletGenerationResponse)
            
            # Handle dict response if Pydantic conversion failed
            if isinstance(response, dict):
                response = BulletGenerationResponse(**response)
            
            # Add original length
            response.original_length = len(text)
            
            self.logger.info(
                f"✓ AI bullets: {len(response.bullets)} points, "
                f"quality {response.summary_quality}/10"
            )
            
            return response
            
        except Exception as e:
            self.logger.warning(f"Structured output failed: {e}, trying manual parse")
            
            # Fallback to regular invoke
            response_text = self.llm.invoke(messages)
            
            # Manual parsing
            import json
            import re
            
            json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
            if json_match:
                data = json.loads(json_match.group(0))
                validated = BulletGenerationResponse(**data)
                validated.original_length = len(text)
                return validated
            
            raise ValueError("Could not parse AI response")
    
    @retry(
        stop=stop_after_attempt(2),
        wait=wait_exponential(multiplier=1, min=1, max=5)
    )
    def clean_text(self, text: str) -> TextCleaningResponse:
        """
        Intelligently remove quotes, artifacts, and clean text.
        
        Args:
            text: Text to clean
            
        Returns:
            TextCleaningResponse with cleaned text and metadata
        """
        if not text or not text.strip():
            return TextCleaningResponse(
                cleaned_text="",
                confidence=1.0,
                changes_made=0
            )
        
        self.logger.debug(f"Cleaning text: {len(text)} chars")
        
        messages = self._build_cleaning_prompt(text)
        
        try:
            response = self.llm.invoke_structured(messages, TextCleaningResponse)
            
            self.logger.info(
                f"✓ Text cleaned: {response.changes_made} changes, "
                f"confidence {response.confidence:.2f}"
            )
            
            return response
            
        except Exception as e:
            self.logger.warning(f"AI cleaning failed: {e}, using basic cleanup")
            
            # Fallback to basic cleaning
            cleaned = text.strip().strip('"\'""''')
            cleaned = ' '.join(cleaned.split())
            
            return TextCleaningResponse(
                cleaned_text=cleaned,
                removed_quotes=['"', "'"],
                confidence=0.7,
                changes_made=1
            )
    
    def _build_bullet_prompt(self, text: str, max_bullets: int) -> List:
        """Build prompt for bullet generation."""
        
        system = f"""You are an expert at formatting text into bullet points for PowerPoint presentations.
        
CRITICAL GOAL: Split the input text into bullet points WITHOUT changing the original wording or summarizing.

RULES:
1. Split long text into {max_bullets} or fewer logical bullet points
2. PRESERVE EXACT WORDING as much as possible
3. Do NOT summarize or shorten the content unless absolutely necessary for grammar
4. Remove ALL quotation marks (", ', ", ", ', ')
5. Do NOT add new information
6. Do NOT change the meaning
7. Just break the text into readable chunks
8. Identify and exclude any vague, generic, or low-information lines that do not add meaningful value.
   Vague lines include statements such as:
   - "Uploaded details found on website"
   - "This information can be accessed on the website"
   - Any similar filler or placeholder text that does not provide specific, actionable details.

QUALITY GUIDELINES:
- Rate your output 1-10 based on fidelity to original text
- 10: Perfect preservation of content, just formatted as bullets
- 1-5: Poor - content was summarized or changed too much

EXAMPLES:

Example 1:
Input: "The company announced a new product launch in Q2 and revenue increased by 15%."
Output: {{
  "bullets": ["The company announced a new product launch in Q2", "Revenue increased by 15%"],
  "summary_quality": 10,
  "removed_quotes": false
}}

Example 2:
Input: "Regulatory changes include stricter compliance requirements. Companies must adapt their operations accordingly."
Output: {{
  "bullets": ["Regulatory changes include stricter compliance requirements", "Companies must adapt their operations accordingly"],
  "summary_quality": 10,
  "removed_quotes": false
}}

Return ONLY valid JSON matching the BulletGenerationResponse schema."""

        user = f"""Convert this text into bullet points:

TEXT: {text}

Generate {max_bullets} or fewer high-quality bullet points. Return ONLY JSON."""

        return [
            SystemMessage(content=system),
            HumanMessage(content=user)
        ]
    
    def _build_cleaning_prompt(self, text: str) -> List:
        """Build prompt for text cleaning."""
        
        system = """You are an expert at cleaning and formatting text for professional presentations.

TASK: Remove quotes, extra spaces, and formatting artifacts while preserving meaning.

RULES:
1. Remove ALL types of quotes: ", ', ", ", ', '
2. Remove extra whitespace (multiple spaces, tabs, newlines)
3. Preserve the core meaning and all important words
4. Don't change capitalization unless fixing obvious errors
5. Track what you removed for transparency

EXAMPLES:

Input: "The company announced new products"
Output: {
  "cleaned_text": "The company announced new products",
  "removed_quotes": ["\\""],
  "confidence": 1.0,
  "changes_made": 1
}

Input: "Revenue  increased   by 15%"
Output: {
  "cleaned_text": "Revenue increased by 15%",
  "removed_artifacts": ["  ", "   "],
  "confidence": 1.0,
  "changes_made": 2
}

Return ONLY valid JSON matching the TextCleaningResponse schema."""

        user = f"""Clean this text:

TEXT: {text}

Remove quotes and artifacts. Return ONLY JSON."""

        return [
            SystemMessage(content=system),
            HumanMessage(content=user)
        ]

    @retry(
        stop=stop_after_attempt(2),
        wait=wait_exponential(multiplier=1, min=1, max=5)
    )
    def generate_bullets_batch(self, texts: List[str], max_bullets: int = 5) -> List[BulletGenerationResponse]:
        """
        Batch process multiple texts to generate bullets in a single API call.
        
        Args:
            texts: List of source texts
            max_bullets: Maximum bullets per item
            
        Returns:
            List of BulletGenerationResponse objects corresponding to input texts
        """
        if not texts:
            return []
            
        # Filter out empty texts but keep track of indices
        valid_indices = [i for i, t in enumerate(texts) if t and t.strip()]
        valid_texts = [texts[i] for i in valid_indices]
        
        if not valid_texts:
            return [
                BulletGenerationResponse(bullets=[""], summary_quality=1, removed_quotes=False, original_length=0)
                for _ in texts
            ]
            
        self.logger.info(f"Batch generating bullets for {len(valid_texts)} items...")
        
        messages = self._build_batch_bullet_prompt(valid_texts, max_bullets)
        
        try:
            # We expect a list of responses
            from pydantic import BaseModel
            class BatchResponse(BaseModel):
                responses: List[BulletGenerationResponse]
            
            response = self.llm.invoke_structured(messages, BatchResponse)
            
            if isinstance(response, dict):
                response = BatchResponse(**response)
                
            results = response.responses
            
            # Map back to original indices
            final_results = []
            result_idx = 0
            
            for i in range(len(texts)):
                if i in valid_indices:
                    if result_idx < len(results):
                        res = results[result_idx]
                        res.original_length = len(texts[i])
                        final_results.append(res)
                        result_idx += 1
                    else:
                        # Should not happen if LLM behaves
                        final_results.append(BulletGenerationResponse(
                            bullets=["Error: Missing batch response"], 
                            summary_quality=1, 
                            removed_quotes=False, 
                            original_length=len(texts[i])
                        ))
                else:
                    final_results.append(BulletGenerationResponse(
                        bullets=[""], 
                        summary_quality=1, 
                        removed_quotes=False, 
                        original_length=0
                    ))
            
            self.logger.info(f"✓ Batch generation complete. Processed {len(valid_texts)} items.")
            return final_results
            
        except Exception as e:
            self.logger.warning(f"Batch generation failed: {e}, falling back to individual processing")
            # Fallback: process individually
            return [self.generate_bullets(t, max_bullets) for t in texts]

    def _build_batch_bullet_prompt(self, texts: List[str], max_bullets: int) -> List:
        """Build prompt for batch bullet generation."""
        
        system = f"""You are an expert at formatting text into bullet points.
        
TASK: Process multiple texts and convert EACH one into bullet points.

RULES:
1. Return a JSON object with a 'responses' array.
2. The 'responses' array must contain exactly {len(texts)} items, corresponding to the input texts in order.
3. For each item, follow the standard bullet generation rules:
   - Split into {max_bullets} or fewer logical bullets
   - PRESERVE EXACT WORDING
   - Remove quotes
   - Rate quality 1-10
   - EXCLUDE vague/generic lines (e.g., "Uploaded details found on website", "This information can be accessed on the website")
   
OUTPUT FORMAT:
{{
  "responses": [
    {{
      "bullets": ["Point 1", "Point 2"],
      "summary_quality": 10,
      "removed_quotes": false
    }},
    ...
  ]
}}"""

        user_content = "Process these texts:\n\n"
        for i, text in enumerate(texts):
            user_content += f"--- ITEM {i+1} ---\n{text}\n\n"
            
        user_content += "Return JSON with 'responses' array."

        return [
            SystemMessage(content=system),
            HumanMessage(content=user_content)
        ]
