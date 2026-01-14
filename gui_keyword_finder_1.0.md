# GUI Keyword Finder for PDFs (Kreuzberg)

A small Tkinter desktop app that uses **Kreuzberg** to extract PDF text and **search keywords**, then reports matches with:

- **File path**
- **Keyword**
- **Page number**
- **Line number within that page**
- **Matched text**
- Optional **Export to CSV**

This app is a GUI wrapper around `batch_extract_files_sync` + keyword scanning.

---

## Features

- ✅ Add multiple PDF files (multi-select)
- ✅ Keywords input supports **comma-separated** or **newline-separated**
- ✅ Case-insensitive matching for ASCII keywords (e.g., `history` matches `History`)
- ✅ Shows **Page + Line (within page)** for each keyword hit
- ✅ Export results to **CSV**

---

## Requirements

- Python **3.9+** recommended
- `kreuzberg` installed
- Tkinter (usually bundled with Python on macOS/Windows; on some Linux distros you may need to install it separately)

---

## Installation

### 1) Create and activate a virtual environment (recommended)

```bash
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
# .venv\Scripts\activate   # Windows PowerShell
```

### 2) Install dependencies

```bash
python -m pip install -U pip
python -m pip install kreuzberg
```

---

## Usage

Run the GUI:

```bash
python gui_keyword_finder.py
```

### In the app

1. Click **Add PDFs...** to select one or more `.pdf` files  
2. Enter keywords in the keyword box:
   - one per line, or
   - separated by commas
3. Click **Run**
4. Matches appear in the results table
5. Click **Export CSV...** to save results

---

## How Page + Line Mapping Works

Kreuzberg can insert **page markers** into the extracted full document text:

```text
--- Page 1 ---
...page text...
--- Page 2 ---
...page text...
```

This script enables that behavior via:

- `PageConfig(extract_pages=True, insert_page_markers=True, marker_format=...)`

Then it parses those markers and computes **line numbers within each page**, so every hit is reported as:

- `Page X, Line Y (within that page)`

---

## Keyword Matching Rules

- For ASCII keywords (e.g., `history`): **case-insensitive** match  
  - `history` will match `History`, `HISTORY`, etc.
- For non-ASCII keywords (e.g., Chinese `签署`): exact substring match

You can modify this logic in `hit()` inside the script if you want all keywords to be case-sensitive.

---

## Output

Results table columns:

- `File` — absolute path  
- `Keyword` — the keyword that matched  
- `Page` — 1-based page number  
- `Line (within page)` — 1-based line index within that page  
- `Matched text` — the matched line content  

CSV export uses the same columns.

---

## Troubleshooting

### 1) “No matches found” but you expect matches

- PDF text extraction may vary (scanned PDFs, encoding, layout).
- Try enabling OCR options if your Kreuzberg setup supports it (depends on your environment and Kreuzberg version).
- Verify the keyword exists in extracted text by printing `result.content` for a file.

### 2) Page is “unknown” (should not happen with current script)

- Ensure `insert_page_markers=True` and `marker_format` includes newlines:
  - `"\n--- Page {page_num} ---\n"`

### 3) Tkinter not found (Linux)

On Debian/Ubuntu:

```bash
sudo apt-get install python3-tk
```

---

## Customization Ideas

- Add a checkbox for “Case sensitive”
- Add “Context lines ±N” around each hit
- Save results as JSON
- Add a “Open PDF at page” button (Preview/Skim on macOS)

 

 
