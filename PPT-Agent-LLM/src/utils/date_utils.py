"""
Date utility functions for week-based folder management.
"""

from datetime import datetime, timedelta
from typing import Tuple
import logging


def get_current_week_range() -> Tuple[datetime, datetime]:
    """
    Calculate the previous completed week's date range (Monday to Sunday).
    
    Note: "This week" refers to the PREVIOUS completed week for reporting purposes.
    For example, if today is 2026-01-21 (in week Jan 19-25),
    this function returns Jan 12-18 (the previous completed week).
    
    Returns:
        Tuple of (start_date, end_date) where start_date is Monday and end_date is Sunday
        of the PREVIOUS completed week
    """
    today = datetime.now()
    
    # Get the weekday (0 = Monday, 6 = Sunday)
    weekday = today.weekday()
    
    # Calculate Monday of current week
    current_monday = today - timedelta(days=weekday)
    
    # Go back 7 days to get Monday of previous week
    previous_monday = current_monday - timedelta(days=7)
    
    # Calculate Sunday of previous week
    previous_sunday = previous_monday + timedelta(days=6)
    
    return (previous_monday, previous_sunday)


def format_week_folder_name(start_date: datetime, end_date: datetime, prefix: str = "weekly_outputs") -> str:
    """
    Format a week folder name based on start and end dates.
    
    Args:
        start_date: Start date of the week (Monday)
        end_date: End date of the week (Sunday)
        prefix: Folder prefix (default: "weekly_outputs")
        
    Returns:
        Folder name string in format: "weekly_outputs/YYYY-MM-DD_to_YYYY-MM-DD/"
    """
    start_str = start_date.strftime("%Y-%m-%d")
    end_str = end_date.strftime("%Y-%m-%d")
    
    return f"{prefix}/{start_str}_to_{end_str}/"


def get_current_week_folder_name(prefix: str = "weekly_outputs") -> str:
    """
    Get the folder name for the previous completed week (referred to as "this week" for reporting).
    
    Args:
        prefix: Folder prefix (default: "weekly_outputs")
        
    Returns:
        Folder name string for the previous completed week
    """
    monday, sunday = get_current_week_range()
    return format_week_folder_name(monday, sunday, prefix)


def parse_week_folder_name(folder_name: str) -> Tuple[datetime, datetime]:
    """
    Parse a week folder name to extract start and end dates.
    
    Args:
        folder_name: Folder name in format "weekly_outputs/YYYY-MM-DD_to_YYYY-MM-DD/"
        
    Returns:
        Tuple of (start_date, end_date)
        
    Raises:
        ValueError: If folder name format is invalid
    """
    try:
        # Remove prefix and trailing slash
        # Example: "weekly_outputs/2026-01-20_to_2026-01-26/" -> "2026-01-20_to_2026-01-26"
        parts = folder_name.strip('/').split('/')
        if len(parts) < 2:
            raise ValueError(f"Invalid folder name format: {folder_name}")
        
        date_part = parts[-1]  # Get the last part (date range)
        
        # Split by "_to_"
        date_strings = date_part.split('_to_')
        if len(date_strings) != 2:
            raise ValueError(f"Invalid date range format: {date_part}")
        
        start_date = datetime.strptime(date_strings[0], "%Y-%m-%d")
        end_date = datetime.strptime(date_strings[1], "%Y-%m-%d")
        
        return (start_date, end_date)
        
    except Exception as e:
        raise ValueError(f"Failed to parse folder name '{folder_name}': {e}")


def log_week_info():
    """
    Log previous completed week information for debugging.
    """
    monday, sunday = get_current_week_range()
    folder_name = get_current_week_folder_name()
    
    logging.info("=" * 70)
    logging.info("WEEK INFORMATION (Previous Completed Week)")
    logging.info("=" * 70)
    logging.info(f"Today: {datetime.now().strftime('%Y-%m-%d %A')}")
    logging.info(f"Previous Week Start (Monday): {monday.strftime('%Y-%m-%d')}")
    logging.info(f"Previous Week End (Sunday): {sunday.strftime('%Y-%m-%d')}")
    logging.info(f"Target Folder: {folder_name}")
    logging.info("=" * 70)
