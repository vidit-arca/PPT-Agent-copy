"""
Test script to verify current week calculation.
"""

import sys
import os
from datetime import datetime

# Add src to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "../../../")))

from src.utils.date_utils import (
    get_current_week_range,
    get_current_week_folder_name,
    format_week_folder_name,
    parse_week_folder_name
)


def test_week_calculation():
    """Test the week calculation functions."""
    
    print("=" * 70)
    print("CURRENT WEEK CALCULATION TEST")
    print("=" * 70)
    
    # Get current date
    today = datetime.now()
    print(f"Today: {today.strftime('%Y-%m-%d %A')}")
    print()
    
    # Test get_current_week_range
    monday, sunday = get_current_week_range()
    print(f"Week Start (Monday): {monday.strftime('%Y-%m-%d %A')}")
    print(f"Week End (Sunday): {sunday.strftime('%Y-%m-%d %A')}")
    print()
    
    # Test get_current_week_folder_name
    folder_name = get_current_week_folder_name()
    print(f"Expected Folder Name: {folder_name}")
    print()
    
    # Test format_week_folder_name
    formatted = format_week_folder_name(monday, sunday)
    print(f"Formatted Folder Name: {formatted}")
    print()
    
    # Test parse_week_folder_name
    try:
        parsed_start, parsed_end = parse_week_folder_name(folder_name)
        print(f"Parsed Start Date: {parsed_start.strftime('%Y-%m-%d')}")
        print(f"Parsed End Date: {parsed_end.strftime('%Y-%m-%d')}")
        print()
        
        # Verify parsing is correct
        if parsed_start.date() == monday.date() and parsed_end.date() == sunday.date():
            print("✓ SUCCESS: Parsing matches original dates")
        else:
            print("✗ FAILURE: Parsing does not match original dates")
    except Exception as e:
        print(f"✗ FAILURE: Parsing failed with error: {e}")
    
    print("=" * 70)


if __name__ == "__main__":
    test_week_calculation()
