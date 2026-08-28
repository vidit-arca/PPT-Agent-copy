# PPT-Agent-LLM

> **Smart PowerPoint automation agent** — reads weekly regulatory Excel data and automatically fills the Tejomaya presentation template, handling multi-slide pagination, AI-generated bullet points, and dynamic layout.

---

## Table of Contents

1. [What This Does](#1-what-this-does)
2. [How It Works — Big Picture](#2-how-it-works--big-picture)
3. [Project Structure](#3-project-structure)
4. [Prerequisites](#4-prerequisites)
5. [Setup (Step-by-Step)](#5-setup-step-by-step)
6. [Running the Agent](#6-running-the-agent)
7. [Input Format — Excel Files](#7-input-format--excel-files)
8. [Configuration Guide](#8-configuration-guide)
9. [Output](#9-output)
10. [How Overflow / Pagination Works](#10-how-overflow--pagination-works)
11. [Troubleshooting](#11-troubleshooting)
12. [Architecture Reference](#12-architecture-reference)
13. [Changelog](#13-changelog)

---

## 1. What This Does

Every week, the Akshayam team receives regulatory updates from bodies like SEBI, RBI, IFSCA, MCA, IBBI, and AIF. These updates live in multiple Excel files (one per regulatory body). This agent:

1. **Reads all Excel files** from a weekly folder
2. **Generates AI bullet points** for each regulatory update (via Ollama or Azure OpenAI)
3. **Populates the PowerPoint template** (`Akshayam Tejomaya.pptx`) — filling the correct slide's table for each regulatory body
4. **Handles overflow** — if too many rows exist for one slide, it automatically creates continuation slides
5. **Saves a timestamped `.pptx`** ready for presentation

---

## 2. How It Works — Big Picture

```
Weekly Excel folder
  ├── SEBI.xlsx       → Slide 5  (SEBI during the week)
  ├── RBI.xlsx        → Slide 6  (RBI during the week)
  ├── IFSCA.xlsx      → Slide 7  (IFSC during the week)
  ├── Companies Act.xlsx → Slide 4
  ├── IBBI.xlsx       → Slide 8  (Insolvency and Bankruptcy Code)
  ├── AIF.xlsx        → Slide 9  (Alternate Investment Funds)
  └── ICAI.xlsx / Listed Companies.xlsx → Slide 10 (Others)
           │
           ▼
    [PPT-Agent-LLM]
           │
           ├── Reads each Excel sheet (each sheet = one subdomain, e.g. "Circulars", "Press Release")
           ├── Generates AI bullet points for "Gist thereof" column
           ├── Maps each file → slide using FILE_VERTICAL_MAPPING
           ├── Fills table rows on the correct slide
           └── Paginates to new slides if content overflows
           │
           ▼
    tejomaya_filled_YYYYMMDD_HHMMSS.pptx
```

---

## 3. Project Structure

```
PPT-Agent-LLM/
│
├── main.py                         # Entry point (just calls src/core/agent.py)
├── requirements.txt                # Python dependencies
├── .env                            # LLM configuration (API keys, model names)
├── Akshayam Tejomaya.pptx          # The PPT template (DO NOT DELETE)
│
├── src/
│   ├── core/
│   │   └── agent.py                # ⭐ Main logic — mapping, pagination, filling
│   │
│   ├── services/
│   │   ├── ppt/
│   │   │   ├── engine.py           # PPT manipulation (cells, rows, shapes, slides)
│   │   │   ├── layout.py           # ⭐ Overflow detection & row height estimation
│   │   │   └── template.py         # Extracts table structure from the PPT template
│   │   │
│   │   ├── llm/
│   │   │   └── provider.py         # LLM client (Ollama / Azure OpenAI switcher)
│   │   │
│   │   ├── text/
│   │   │   └── processor.py        # AI bullet point generation
│   │   │
│   │   ├── data/
│   │   │   └── profiler.py         # Builds data profile from Excel
│   │   │
│   │   └── storage/
│   │       └── minio_client.py     # MinIO integration (optional cloud storage)
│   │
│   ├── config/
│   │   └── settings.py             # Reads .env config
│   │
│   └── utils/
│       └── logging.py              # Logging setup
│
├── templates/                      # (unused currently, for future template variants)
├── scripts/                        # Utility scripts
├── data/                           # Scratch data folder
└── tests/                          # Test scripts
```

---

## 4. Prerequisites

Before you start, make sure you have:

| Requirement | Why |
|---|---|
| **Python 3.9+** | The agent runs on Python |
| **Ollama** (local) OR **Azure OpenAI** (cloud) | For AI bullet generation and split decisions |
| The `Akshayam Tejomaya.pptx` template | The presentation template to fill |
| Weekly Excel files in a dated folder | Input data |

### Check Python version

```bash
python3 --version
# Should print Python 3.9.x or higher
```

### Install Ollama (for local AI — recommended for offline use)

Download from: https://ollama.com/download

Then pull the model used by this project:

```bash
ollama pull qwen2.5-coder:7b
```

Start Ollama (it runs as a background service after installation on macOS):

```bash
ollama serve
```

---

## 5. Setup (Step-by-Step)

### Step 1 — Clone / navigate to the project

```bash
cd "/Users/apple/Desktop/Akshayam/PPT-Agent copy/PPT-Agent-LLM"
```

### Step 2 — Create a virtual environment (recommended)

```bash
python3 -m venv venv
source venv/bin/activate       # macOS/Linux
# OR on Windows:
# venv\Scripts\activate
```

### Step 3 — Install dependencies

```bash
pip install -r requirements.txt
```

This installs:
- `pandas` — Excel reading
- `python-pptx` — PowerPoint manipulation
- `langchain-ollama` + `langchain-openai` — LLM integration
- `pydantic` — Structured LLM outputs
- `python-dotenv` — Load `.env` config
- `tenacity` — Retry logic for LLM calls

### Step 4 — Configure the `.env` file

Open `.env` and set your LLM provider:

```env
# Choose: "ollama" for local, "azure" for Azure OpenAI
LLM_PROVIDER=ollama

# Model name for Ollama
LLM_MODEL=qwen2.5-coder:7b

# Ollama server URL (default if running locally)
OLLAMA_BASE_URL=http://localhost:11434

# --- OR if using Azure OpenAI ---
LLM_PROVIDER=azure
AZURE_OPENAI_API_KEY=your_key_here
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/
AZURE_API_VERSION=2024-12-01-preview
AZURE_DEPLOYMENT_NAME=gpt-4.1-mini
```

> **Note:** If Ollama is running on a different machine on the network, update `OLLAMA_BASE_URL` to that machine's IP (e.g., `http://192.168.1.100:11434`).

### Step 5 — Prepare your weekly Excel folder

Create a folder named after the week's date range and place your Excel files inside:

```
2026-08-17_to_2026-08-22/
  ├── SEBI.xlsx
  ├── RBI.xlsx
  ├── IFSCA.xlsx
  ├── IBBI.xlsx
  ├── AIF.xlsx
  ├── Companies Act.xlsx
  ├── ICAI.xlsx
  └── Listed Companies.xlsx
```

> Each Excel file can have **multiple sheets** — one per subdomain (e.g., `Circulars`, `Press Release`, `Notifications`).

### Step 6 — Verify the PPT template is present

```bash
ls "Akshayam Tejomaya.pptx"
# Should print the filename. If missing, get it from the team.
```

---

## 6. Running the Agent

### ✅ Standard Run (most common)

```bash
python3 -m src.core.agent \
  --excel-dir "/path/to/your/weekly-folder" \
  --local-only
```

**Real example:**

```bash
python3 -m src.core.agent \
  --excel-dir "/Users/apple/Desktop/Akshayam/PPT-Agent copy/2026-08-17_to_2026-08-22" \
  --local-only
```

### All CLI Options

| Flag | Description | Example |
|---|---|---|
| `--excel-dir PATH` | Path to folder containing all weekly Excel files | `--excel-dir ./2026-08-17_to_2026-08-22` |
| `--excel PATH` | Path to a single Excel file (single-file mode) | `--excel ./SEBI.xlsx` |
| `--ppt PATH` | Path to PPT template (auto-detected if not set) | `--ppt "Akshayam Tejomaya.pptx"` |
| `--local-only` | Skip MinIO cloud storage, use local files only | `--local-only` |

### Using main.py (alternative)

```bash
python3 main.py \
  --excel-dir "/path/to/weekly-folder" \
  --local-only
```

---

## 7. Input Format — Excel Files

### File Naming → Vertical Mapping

The filename (without `.xlsx`) determines which slide the data goes to:

| File Name | Vertical | PPT Slide |
|---|---|---|
| `SEBI.xlsx` | SEBI | Slide 5 |
| `RBI.xlsx` | RBI | Slide 6 |
| `IFSCA.xlsx` | IFSCA | Slide 7 |
| `Companies Act.xlsx` | Companies Act | Slide 4 |
| `IBBI.xlsx` | Insolvency and Bankruptcy Code | Slide 8 |
| `AIF.xlsx` | Alternate Investment Funds | Slide 9 |
| `ICAI.xlsx` | Others during the week | Slide 10 |
| `Listed Companies.xlsx` | Others during the week | Slide 10 |

> To add a new file-to-vertical mapping, edit `FILE_VERTICAL_MAPPING` in [`src/core/agent.py`](src/core/agent.py).

### Sheet Names → Subdomains

Each sheet in an Excel file becomes a **subdomain** row category. Common sheet names:

- `Circulars`
- `Press Release`
- `Notifications`
- `Master Directions`
- `Consultation Paper`
- `Public Consultation`
- `FAQs`

### Required Columns (per sheet)

Each sheet must contain these columns:

| Column | Description | Example |
|---|---|---|
| `IssueDate` | Date of the circular/notification | `2026-08-20` |
| `SubCategory` | Type of document | `Circular`, `Press Release` |
| `PDF_URL` | Link to the PDF document | `https://sebi.gov.in/...pdf` |
| `Title` | Short title / subject line | `Ease of onboarding for FPIs` |
| `Summary` | Full description (AI uses this for bullet points) | Long text paragraph |

> The `Verticals` column is **automatically added** by the agent based on the filename — you don't need it in the Excel.

---

## 8. Configuration Guide

### 8.1 Slide Mapping

Edit `MANUAL_SLIDE_MAPPING` in [`src/core/agent.py`](src/core/agent.py) (line ~32):

```python
MANUAL_SLIDE_MAPPING = {
    "Companies Act": 3,                   # Slide 4 (0-indexed)
    "SEBI": 4,                            # Slide 5
    "RBI": 5,                             # Slide 6
    "IFSC": 6,                            # Slide 7
    "Insolvency and Bankruptcy Code": 7,  # Slide 8
    "Alternate Investment Funds": 8,      # Slide 9
    "Others during the week": 9,          # Slide 10
}
```

> **Note:** Slide indices are **0-based** (so Slide 5 in PowerPoint = index 4 here).

### 8.2 Adding a New Regulatory Body

1. Add its Excel filename → vertical mapping in `FILE_VERTICAL_MAPPING`:
   ```python
   FILE_VERTICAL_MAPPING = {
       "NewBody": "New Regulatory Body",  # Add this line
       ...
   }
   ```

2. Add the vertical → slide mapping in `MANUAL_SLIDE_MAPPING`:
   ```python
   MANUAL_SLIDE_MAPPING = {
       "New Regulatory Body": 11,  # Slide 12 in PPT
       ...
   }
   ```

3. Add it to the allowed list in `ALLOWED_VERTICALS` (line ~1134 in agent.py):
   ```python
   ALLOWED_VERTICALS = [..., "New Regulatory Body"]
   ```

### 8.3 Overflow / Layout Tuning

Edit constants in [`src/services/ppt/layout.py`](src/services/ppt/layout.py):

```python
self.CHARS_PER_LINE_PER_INCH = 13   # Lower = more lines per bullet → rows estimated taller
self.LINE_HEIGHT_INCHES = 0.26      # Height per text line in inches
self.PADDING_INCHES = 0.36          # Cell top/bottom padding
self.PARA_SPACING_INCHES = 0.20     # Space between bullet points
```

The overflow check margin (in `agent.py`, line ~386):

```python
# Default margin for all slides (0.5" from slide bottom)
bottom_margin=0.5

# IFSCA uses 1.5" because its URLs wrap heavily and overestimate fitting
ifsca_margin = 1.5 if is_ifsca else 0.5
```

> **Rule of thumb:** Increase `bottom_margin` if a slide's table visually overflows the white background. Decrease if the table is paginating too early (too many slides created).

### 8.4 LLM Provider

Switch between local Ollama and Azure OpenAI in `.env`:

```env
# For local (no internet required):
LLM_PROVIDER=ollama
LLM_MODEL=qwen2.5-coder:7b

# For Azure OpenAI:
LLM_PROVIDER=azure
AZURE_DEPLOYMENT_NAME=gpt-4.1-mini
```

---

## 9. Output

The agent saves the filled presentation in the project root:

```
tejomaya_filled_20260822_125134.pptx
```

The file is automatically opened after generation (macOS). You can also open it manually in PowerPoint.

**What gets filled:**
- ✅ Table rows with S.No, Date, Rules/Circular link, Title, and AI bullet points
- ✅ Slide subtitle (e.g., `Circulars – 1; Press Release – 4; ...`)
- ✅ Continuation slides when content overflows
- ✅ Footer preserved only on the last slide of each section

---

## 10. How Overflow / Pagination Works

When a regulatory body has more data than fits on one slide, the agent:

1. **Estimates row heights** using character count, column widths, and bullet point count
2. **Detects overflow** when `current_height + next_row_height > slide_height - bottom_margin`
3. **Duplicates the slide** (copying all shapes, logos, backgrounds, and relationships)
4. **Fills the current slide** with as many rows as fit
5. **Clears the duplicate** (keeps only header row) and **recurses** with remaining data
6. **Removes footer** (news table, bottom decorators) from non-last slides
7. **AI decides split points** — if two adjacent rows are strongly related, they stay together

```
Slide 5 (SEBI)          Slide 6 (SEBI cont.)      Slide 7 (SEBI cont.)
┌──────────────┐        ┌──────────────┐           ┌──────────────┐
│ Row 1        │        │ Row 4        │           │ Row 7        │
│ Row 2        │   →    │ Row 5        │   →       │ [footer]     │
│ Row 3        │        │ Row 6        │           └──────────────┘
│ [no footer]  │        │ [no footer]  │
└──────────────┘        └──────────────┘
```

---

## 11. Troubleshooting

### ❌ `command not found: python3`

Install Python from https://python.org or use `python` instead of `python3`.

### ❌ `ModuleNotFoundError: No module named 'pptx'`

You forgot to install dependencies or activate your virtual environment:

```bash
source venv/bin/activate
pip install -r requirements.txt
```

### ❌ Ollama connection refused

Make sure Ollama is running:

```bash
ollama serve
# Then in a new terminal:
ollama list   # Should show qwen2.5-coder:7b
```

If Ollama is on another machine, update `OLLAMA_BASE_URL` in `.env` to its IP address.

### ❌ Table overflowing the white background on a slide

The height estimator is too optimistic for that vertical. Increase `bottom_margin` for it in `agent.py` (see §8.3). IFSCA already has a dedicated fix with `ifsca_margin = 1.5`.

### ❌ Too many continuation slides being created

The estimator is too conservative. Try slightly reducing constants in `layout.py`:

```python
self.LINE_HEIGHT_INCHES = 0.22   # Try smaller value
self.PADDING_INCHES = 0.30
```

### ❌ `No data remaining after filtering`

The `Verticals` column value in your data doesn't match any entry in `ALLOWED_VERTICALS`. Check:

1. What vertical name did the agent assign? Check the log for `Mapped '...' to Vertical: '...'`
2. Add it to `ALLOWED_VERTICALS` in `agent.py` if it's valid

### ❌ `⚠ Unrecognized subdomain for subtitle: 'Public Consultation'`

This is a **warning, not an error** — the subtitle counter doesn't know this subdomain. The slide still fills correctly. To fix it, add the subdomain alias to `update_slide_subtitle()` in [`src/services/ppt/engine.py`](src/services/ppt/engine.py) under the appropriate `categories` dict.

### ❌ Missing vertical (no mapping found)

```
✗ No mapping found for 'XYZ' – add it to MANUAL_SLIDE_MAPPING
```

Follow §8.2 to add the new body.

---

## 12. Architecture Reference

```
┌─────────────────────────────────────────────────────────┐
│                    main() in agent.py                    │
│                                                         │
│  1. Parse CLI args (--excel-dir, --local-only, etc.)    │
│  2. Read all Excel files from directory                  │
│  3. Filter to ALLOWED_VERTICALS                          │
│  4. Load PPT template                                    │
│  5. Build mapping (vertical → slide, shape)             │
│  6. For each target slide:                               │
│     a. Precompute AI bullets (concurrent)               │
│     b. Call paginate_and_fill()                          │
│  7. Save output .pptx                                    │
└─────────────────────────────────────────────────────────┘
          │
          ▼
┌─────────────────────────────────────────────────────────┐
│               paginate_and_fill()                        │
│                                                         │
│  - Estimates how many rows fit using LayoutEngine       │
│  - If overflow: duplicate_slide() → recurse             │
│  - Fills rows using set_cell_text() from engine.py      │
│  - Updates subtitle using update_slide_subtitle()       │
└─────────────────────────────────────────────────────────┘
          │
          ▼
┌─────────────────────────────────────────────────────────┐
│               LayoutEngine (layout.py)                   │
│                                                         │
│  - estimate_row_height(): chars/line × height/line      │
│  - check_overflow(): projected > available - margin     │
│  - analyze_split_point(): AI decides if rows must stay  │
│    together across the page break                       │
└─────────────────────────────────────────────────────────┘
```

### Key Files at a Glance

| File | Purpose |
|---|---|
| [`src/core/agent.py`](src/core/agent.py) | Main pipeline — CLI, data loading, mapping, pagination |
| [`src/services/ppt/layout.py`](src/services/ppt/layout.py) | Row height estimation, overflow detection |
| [`src/services/ppt/engine.py`](src/services/ppt/engine.py) | Low-level PPT ops (cells, rows, shapes, dates) |
| [`src/services/ppt/template.py`](src/services/ppt/template.py) | Reads PPT template structure |
| [`src/services/llm/provider.py`](src/services/llm/provider.py) | LLM client factory (Ollama / Azure) |
| [`src/services/text/processor.py`](src/services/text/processor.py) | Generates AI bullet points |
| [`.env`](.env) | LLM API keys and provider selection |

---

## 13. Changelog

### v3.0 — Layout & IFSCA Fix (2026-08)
- ✅ Tuned `LayoutEngine` constants for more accurate row height estimation
- ✅ `paginate_and_fill` accepts per-slide `bottom_margin` override
- ✅ IFSCA slides use `bottom_margin=1.5"` to handle long URL wrapping
- ✅ Default `bottom_margin` reduced from `0.8"` to `0.5"` for all other slides

### v2.0 — Overflow Handling (2025-11)
- ✅ Automatic overflow detection using `LayoutEngine`
- ✅ Smart slide duplication with image/relationship preservation
- ✅ Recursive pagination across multiple continuation slides
- ✅ AI-powered split point decision (keeps related rows together)
- ✅ Footer removed from non-last overflow slides
- ✅ Concurrent AI bullet precomputation

### v1.0 — Core Features (2025-05)
- Basic table population from Excel
- Manual slide mapping
- Date formatting (DD-MM-YY)
- Font and style preservation

---

## License

Internal use — Akshayam Corporate
