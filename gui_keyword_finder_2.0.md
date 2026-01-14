# GUI Keyword Finder for PDFs Version 2.0

Batch key word search across multiple PDFs with in-app page preview and visual highlighting. Designed to reduce the overhead of verification and nevigation when working with multiple documents.

# Background

Standard PDF readers provide reliable keyword search, but the workflow becomes inefficient when you need to:

- Search multiple PDFs and/or multiple keywords in one run
- repeatedly verify hits by switching between a result list and a seperate PDF viewer
- keep an exportable record of matches for follow up, reporting, or review

This tool focuses on shortening th "search -> verify -> extract" loop by integrating navigation and visualization into the result review step.

# Key Capabilities 

- Batch Search: search across different PDFs and different keywords in a single ru 
- Interactive verification: selecting a result row jumps to the corresponding page in the app
- Visual localization: keyword matches are highlighted with bounding boxes on the rendered page
- Context inspection: shows nearby context to confirm the match without leaving the app
- Export: CSV/JSON export including page + bounding box coordinates

# Typical Use Cases

- Reviewing policies/specifications across versions
- Scanning large sets of lecture notes / search readings
- Spot-checking compliance keywords in document packs
- Building a checklist of occurrences for downstream processing

# Architecture Overview

- Search + coordinates: PyMuPDF (fitz) provides text search and hit bounding boxes
- Rendering: pages are rendered to images for preview, with overlays for highlight
- UI: PySide6 (Qt) provides a responsive, cross-platform desktop interface

# Installation 

## Requirements
- Python 3.10+ recommended

## Dependencies
``
pip install PySide6 pymupdf pillow
``
## Run

``
python integrated_keyword_finder.py 
``

# Usage

1. Click Add PDFs and select one or more PDF files
2. Enter keywords:
   - keyword input supports comma-seperated values
   - the multi-line box supports one keyword per line (commas also supported)
3. Click Search
4. Click a row in the result table to:
   - preview the target page
   - see highlighted occurrences
   - read surrounding context
5. Export results if needed:
   - Export CSV
   - Export JSON

# Output Format

## CSV

Columns:
- file
- keyword
- page
- x0, y0, x1, y1 (bounding box)
- snippet

## JSON 

Each records:
- file, keyword, page
- rect: {x0, y0, x1, y1}
- snippet

# Limitations

- PDFs without a text layer (scanned documents) may return no hits unless OCR is enabled (not included by default).
- Current highlighting marks all occurrences of the selected keyword on that page.

# Roadmap

- Search options: case sensitivity, whole-word match, regex
- Result management: filtering, grouping (by file/leyword/page), duplication
- Performance: caching and incremental search for repeated runs
- Optional OCR fallback for scanned PDFs
- "Highlight only this hit" mode for dense pages

# Screenshots

<img width="2512" height="1396" alt="image" src="https://github.com/user-attachments/assets/3ae94866-f517-457c-9f04-494ecb021e94" />

- Results table with multi-keyword search
- Page preview with bounding-box highlights


