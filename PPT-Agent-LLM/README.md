# PPT-Agent-LLM

Smart PowerPoint automation agent that populates presentations with Excel data, featuring automatic table overflow handling.

## Overview

This system automatically fills PowerPoint tables with data from Excel files, mapping verticals (like SEBI, RBI, Companies Act) to specific slides. When tables overflow, it intelligently duplicates slides and restructures content.

## Features

### 🎯 Core Capabilities
- **Automatic Slide Mapping**: Maps Excel verticals to PPT slides using manual configuration or LLM assistance
- **Smart Data Population**: Fills tables while preserving formatting and fonts
- **Date Formatting**: Converts dates to DD-MM-YYYY format
- **Dynamic Row Addition**: Adds table rows as needed

### 🚀 Advanced Features (New)
- **Overflow Detection**: Monitors when content exceeds slide boundaries
- **Automatic Slide Duplication**: Creates continuation slides when overflow detected
- **Content Restructuring**: Splits data across slides intelligently
- **Image Preservation**: Maintains logos and images on duplicate slides
- **Smart Positioning**: Auto-positions content to prevent overflow

## File Structure

```
PPT-Agent-LLM/
├── agent.py                 # Main application with overflow handling
├── template_spec.py         # Extracts PPT structure
├── data_profile.py          # Builds Excel data profile
├── requirements.txt         # Python dependencies
├── verify_ppt.py           # Quick PPT verification script
├── inspect_ppt.py          # Detailed slide inspection tool
├── reproduce_issue.py      # Test script for overflow scenarios
├── verify_overflow.py      # Overflow handling verification
└── test_overflow.pptx      # Test template with two tables
```

## Installation

```bash
# Install dependencies
pip install -r requirements.txt

# Requires Ollama with Mistral model (optional for LLM mapping)
# Download from: https://ollama.ai
```

## Usage

### Basic Usage

```bash
python agent.py --excel "SEBI.xlsx" --ppt "template.pptx"
```

### Auto-Detection (if files are in current directory)

```bash
python agent.py
```

The system auto-detects the first `.xlsx` and `.pptx` files in the directory.

### Expected Data Format

**Excel File**: Must have a sheet named "Press Release" with columns:
- `Verticals` - The regulatory body (SEBI, RBI, etc.)
- `IssueDate` - Date of issue
- `SubCategory` - Type of circular/notification
- `Title` - Main content description
- `Summary` - Detailed description

**PowerPoint Template**: Should have slides with 5-column tables matching headers:
- S. No
- Date of Issue
- Rules / circulars / Notifications / Order
- Contents thereof
- Gist thereof

## Configuration

### Manual Slide Mapping

Edit `MANUAL_SLIDE_MAPPING` in `agent.py`:

```python
MANUAL_SLIDE_MAPPING = {
    "Companies Act": 3,  # Slide 4 (0-indexed)
    "SEBI": 4,           # Slide 5
    "RBI": 5,            # Slide 6
    # ... add more mappings
}
```

### LLM Mapping (Optional)

If Ollama with Mistral is running, the system attempts dynamic mapping via LLM. Falls back to manual mapping on failure.

## How Overflow Handling Works

### Detection
- Monitors table growth after adding rows
- Checks if content exceeds slide height (with 0.5" margin)
- Triggers when any shape would be pushed off-page

### Duplication Process
1. **Detects Overflow**: Content exceeds boundaries
2. **Duplicates Slide**: Creates exact copy positioned right after original
3. **Restructures Content**:
   - **Original Slide**: Keeps filled table, removes second table
   - **Duplicate Slide**: Clears first table (keeps header), positions second table with proper spacing
4. **Preserves Assets**: Copies all images, logos, and formatting
5. **Cleans Up**: Removes ghost placeholders

### Example Output

```
⚠ OVERFLOW DETECTED on Slide 5!
Creating overflow slide...
Removed empty placeholder from duplicate slide (x2)
Moving 1 shapes to new slide...
Cleared 4 data rows from Table 1 on new slide (kept header)
Shifting 1 shapes up by 4.17 inches
✓ SUCCESS! Filled 4 total rows across all sectors
```

## Key Functions

### `agent.py`

- **`duplicate_slide(prs, source_slide_index)`**: Duplicates a slide with all content and relationships
- **`fill_using_mapping(prs, df, mapping)`**: Core data population with overflow handling
- **`add_table_rows(table, num_rows_to_add)`**: Dynamically adds table rows
- **`move_shapes_below(slide, table_shape, added_rows_count)`**: Repositions content
- **`remove_shape(slide, shape)`**: Removes shapes from slides
- **`get_shapes_below(slide, threshold_top)`**: Finds shapes below a position
- **`format_date(value)`**: Formats dates to DD-MM-YYYY
- **`set_cell_text(cell, value)`**: Sets text while preserving formatting

### `template_spec.py`

- **`extract_template_spec(ppt_path, max_slides)`**: Analyzes PPT structure, extracting shapes and tables

### `data_profile.py`

- **`build_data_profile(excel_path, sheet_name, sample_rows)`**: Profiles Excel data structure

## Testing

### Test Overflow Handling

```bash
python reproduce_issue.py
```

### Inspect Generated PPT

```bash
python inspect_ppt.py
```

### Verify Output

```bash
python verify_overflow.py
```

## Output

Generated files are saved with timestamps:
```
tejomaya_filled_YYYYMMDD_HHMMSS.pptx
```

## Troubleshooting

### LLM Returns Invalid JSON
- System automatically falls back to manual mapping
- Ensure `MANUAL_SLIDE_MAPPING` is configured

### Missing Images on Duplicate Slides
- Resolved in current version
- Uses deep relationship copying

### Content Still Overflowing
- Check slide height in template
- Adjust overflow margin in code (currently 0.5")
- Verify table positioning

### No Vertical Found
- Add to `MANUAL_SLIDE_MAPPING`
- Check Excel "Verticals" column spelling

## Dependencies

- `pandas`: Excel data handling
- `python-pptx`: PowerPoint manipulation
- `langchain-ollama`: LLM integration (optional)
- `langchain-core`: LLM framework (optional)

## License

Internal use - Akshayam Corporate

## Recent Updates

### v2.0 - Overflow Handling (2025-11-21)
- ✅ Automatic overflow detection
- ✅ Smart slide duplication
- ✅ Content restructuring across slides
- ✅ Image relationship preservation
- ✅ Placeholder cleanup
- ✅ Duplicate slide positioning fix

### v1.0 - Core Features
- Basic table population
- Manual slide mapping
- Date formatting
- Font preservation

### For running it on local dir 
- python main.py --local-only --excel-dir /Users/apple/Desktop/PPT-Agent\ copy/2026-01-12_to_2026-01-18 


python3 -m src.core.agent --excel-dir "/Users/apple/Desktop/PPT-Agent copy/2026-04-13_to_2026-04-19" --ppt "Akshayam Tejomaya.pptx" --local-only

















