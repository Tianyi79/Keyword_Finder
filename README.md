# Kreuzberg Document Keyword Search Tool

A Python tool based on the Kreuzberg document intelligence framework for extracting text from 56+ document formats (PDF, Word, Excel, etc.) and searching for keywords.

https://docs.kreuzberg.dev/getting-started/quickstart/

## Features
- ✅ Support for 56+ document formats (PDF, Word, Excel, PowerPoint, etc.)
- ✅ Multi-keyword search with line numbers and content display
- ✅ Batch processing of multiple documents
- ✅ OCR text recognition (supports scanned documents and images)
- ✅ Table data extraction
- ✅ Metadata extraction (author, creation time, page count, etc.)

## Supported Document Forms

Office Documents

- PDF (.pdf)
- Word (.doc, .docx)
- Excel (.xls, .xlsx)
- PowerPoint (.ppt, .pptx)

Image Files (OCR required)

- PNG, JPEG, TIFF, BMP

Web and Data

- HTML, XML, JSON, Markdown

Email and Archives
- EML, MSG, ZIP, TAR

Academic Formats

- LaTeX, BibTeX, Jupyter Notebook

## System Requirements
Python 3.8+
Windows / macOS / Linux

## Installation
1. Install Kreuzberg
```
pip install kreuzberg
```

2. (Optional) Install OCR Support
To recognize text in scanned documents, install Tesseract:

``Windows``
- Download and install from github.com/UB-Mannheim/tesseract/wiki

``macOs``
```
brew install tesseract tesseract-lang
```

## Advanced Features

OCR Text Recognition (Scanned Documnent and Images)
```
from kreuzberg import extract_file_sync, ExtractionConfig, OcrConfig

# Configure OCR (supports Chinese and English)
config = ExtractionConfig(
    ocr=OcrConfig(backend="tesseract", language="chi_sim+eng")
)

# Extract scanned document
result = extract_file_sync("scanned_document.pdf", config=config)
print(result.content)
```

# Versions

## v1.0
gui_keyword_finder_1.0.py

## v2.0
gui_keyword_finder_2.0.py

