# PPT-Agent Codebase Overview

I have reviewed your `PPT-Agent` codebase located at `/Users/apple/Desktop/Akshayam/PPT-Agent copy`. This is a sophisticated Python application designed to automatically generate or update PowerPoint presentations by populating data from Excel files into specific slide templates.

Here is a breakdown of the core architecture and what the primary scripts do:

## 1. Core Agent Pipeline (`PPT-Agent-LLM/src/core/agent.py`)
This is the heart of the application. It orchestrates the entire data-to-PPT workflow:
*   **Data Ingestion**: It attempts to connect to a MinIO bucket to download the latest weekly Excel files (e.g., using `get_current_week_folder_name()`). If MinIO fails or the `--local-only` flag is used, it falls back to reading Excel files from a local directory or a single file.
*   **Data Processing**: It combines all Excel sheets into a single pandas DataFrame, tagging each row with its `SourceFile`, `Subdomain` (sheet name), and `Verticals` (mapped via `FILE_VERTICAL_MAPPING`). It filters the data to only include allowed verticals (SEBI, RBI, Companies Act, etc.) and excludes specific subdomains for RBI.
*   **Slide Mapping**: It uses `MANUAL_SLIDE_MAPPING` to map specific verticals to specific slide indexes (e.g., "Companies Act" → Slide 4). It supports subdomain-specific mapping to put different subdomains on different slides.
*   **Pagination & Table Filling**: The `paginate_and_fill` function is highly complex. It calculates table heights dynamically. If the data exceeds the slide's height, it duplicates the slide, moves the "footer/news" shapes correctly, and continues filling data on the newly created overflow slide.

## 2. PPT Manipulation Engine (`PPT-Agent-LLM/src/services/ppt/engine.py`)
This module handles all the raw `python-pptx` operations:
*   **Slide Duplication**: The `duplicate_slide` function safely copies a slide's XML and relationships to create a perfect clone for overflow content.
*   **Table Operations**: Contains functions to add/remove rows dynamically (`add_table_rows`, `remove_table_rows`) so that the template table perfectly fits the data.
*   **Text Formatting & AI Bullets**: The `set_cell_text` function handles the styling. If the column requires it (like a "Gist"), it uses `AITextProcessor` to summarize the text into bullet points, with a rule-based fallback if the AI fails.
*   **Dynamic Subtitles**: `update_slide_subtitle` calculates the counts of different subdomains (e.g., "Notifications - 5; Circulars - 2;") and dynamically updates the subtitle on the slide.

## 3. Services & Configuration
*   **`src/services/data/` & `src/services/text/`**: Handlers for profiling data and AI processing (likely wrapping OpenAI or similar LLM calls for bullet point generation).
*   **`src/services/storage/minio_client.py`**: Handles connecting to the MinIO object storage to retrieve the weekly Excel files.
*   **`src/config/settings.py`**: Manages environment variables and configurations for LLMs and Storage.

## 4. Root level & Utility Scripts
*   **`main.py` & `run_agent.bat`**: The entry points to run the pipeline.
*   **`diagnose_template.py` & `inspect_slides.py`**: Utility scripts to iterate over a `.pptx` file and print out the shapes, text, and structure. Useful for debugging template layouts.
*   **`agent.py` (Root Folder)**: An alternative, simpler version of the script ("SAFE TABLE FILLER (Option B)") that just fills tables without adding rows or duplicating slides.

---

### Key Observations & Recommendations

1.  **AI Mapping is Disabled**: In `agent.py`, `ask_llm_for_mapping` is temporarily disabled in favor of `get_fallback_mapping` (manual slide mapping). If the template structure is rigid, manual mapping is much safer and faster anyway.
2.  **Complex Pagination Logic**: The `paginate_and_fill` logic is very impressive. It estimates row heights based on text length and column width to avoid overflowing the slide. However, `python-pptx` height estimation can be tricky, so if you notice tables overflowing the bottom of the slide, you may need to tweak the `bottom_margin` variable (currently set to `0.3`).
3.  **Hardcoded Slide Indexes**: `MANUAL_SLIDE_MAPPING` relies on absolute slide indexes (0-indexed). If a user adds an introduction slide to the template before Slide 4, the entire mapping will be off by one. Consider finding slides by searching their Title Text instead of hardcoded indexes to make the system more robust.
