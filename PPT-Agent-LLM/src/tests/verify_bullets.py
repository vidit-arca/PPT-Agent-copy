import logging
import sys
import os

# Add src to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "../../../")))

from src.services.text.processor import AITextProcessor
from src.utils.logging import setup_logging

# Setup logging
setup_logging()
logger = logging.getLogger(__name__)

def test_vague_line_exclusion():
    processor = AITextProcessor()
    
    test_text = """
    The company reported a 20% increase in revenue.
    Uploaded details found on website.
    New compliance measures were introduced.
    This information can be accessed on the website.
    """
    
    print(f"Input Text:\n{test_text}")
    
    try:
        response = processor.generate_bullets(test_text)
        print("\nGenerated Bullets:")
        for bullet in response.bullets:
            print(f"- {bullet}")
            
        # Check if vague lines are present
        vague_phrases = ["Uploaded details found on website", "This information can be accessed on the website"]
        found_vague = False
        for bullet in response.bullets:
            for phrase in vague_phrases:
                if phrase.lower() in bullet.lower():
                    found_vague = True
                    print(f"\nFAILURE: Found vague phrase: '{phrase}'")
        
        if not found_vague:
            print("\nSUCCESS: No vague phrases found in output.")
            
    except Exception as e:
        print(f"\nERROR: {e}")

if __name__ == "__main__":
    test_vague_line_exclusion()
