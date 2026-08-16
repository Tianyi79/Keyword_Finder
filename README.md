# Keyword Finder

[![Build Windows](https://github.com/Tianyi79/Keyword_Finder/actions/workflows/build-windows.yml/badge.svg)](https://github.com/Tianyi79/Keyword_Finder/actions/workflows/build-windows.yml)

**Keyword Finder** is a local-first desktop app for searching keywords across multiple documents, reviewing matches, previewing PDF results, saving useful quotes, and exporting results.

一个在本地运行的桌面关键词搜索工具。文档搜索在你的电脑上完成，无需上传文件。

## Download

### Windows test build

1. Open the [Build Windows workflow](https://github.com/Tianyi79/Keyword_Finder/actions/workflows/build-windows.yml).
2. Select the latest successful run.
3. Under **Artifacts**, download **Keyword-Finder-Windows**.
4. Extract the downloaded ZIP and run `Keyword Finder.exe`.

Python is not required to run the packaged app. The current test build is unsigned, so Windows may display a SmartScreen warning. Only run builds downloaded from this repository.

### macOS

A macOS installer has not been published to GitHub Releases yet. You can run the app from source using the instructions below.

## Features

- Search multiple files and multiple keywords in one run.
- Review individual hits or group results by page, line, or row.
- Preview PDF pages with highlighted matches and surrounding context.
- Use plain, regular-expression, or fuzzy search modes.
- Control case sensitivity and whole-word matching.
- Save selected text as clips with optional notes.
- Export results to CSV or XLSX and clips to Markdown or CSV.
- Cache PDF and extracted text locally for faster repeat searches.
- Run in normal or portable mode.

## Supported files

| File type | Support |
| --- | --- |
| PDF (`.pdf`) | Direct search and highlighted page preview with PyMuPDF |
| CSV (`.csv`) | Direct row-based search |
| Excel (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`) | Direct row-based search with openpyxl |
| Other documents, such as Word or PowerPoint | Text extraction through Kreuzberg when supported by the installed version |

PDF, CSV, and supported Excel files do not require Kreuzberg. Fuzzy matching applies to CSV, Excel, and Kreuzberg-extracted text; PDF fuzzy mode currently uses literal PDF matching.

## Quick start

1. Select **Files…** and add the documents you want to search.
2. Select **Keywords…** and enter one keyword per line or separate keywords with commas.
3. Choose your search settings, then select **Run search**.
4. Select a result and open **Preview** to inspect the match.
5. Use the **Export** menu to save results or clips.

## Run from source

Python 3.12 is recommended.

```bash
git clone https://github.com/Tianyi79/Keyword_Finder.git
cd Keyword_Finder
python -m venv .venv
```

Activate the virtual environment:

```powershell
# Windows PowerShell
.venv\Scripts\Activate.ps1
```

```bash
# macOS or Linux
source .venv/bin/activate
```

Install dependencies and start the app:

```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python gui_keyword_finder_4.0.1_fix_unpack.py
```

Portable mode stores configuration, cache files, and logs beside the program:

```bash
python gui_keyword_finder_4.0.1_fix_unpack.py --portable
```

## Build the Windows executable

The repository includes a GitHub Actions workflow that builds the app on `windows-latest` with Python 3.12 and PyInstaller.

To start a build, open **Actions → Build Windows → Run workflow**. The finished `Keyword Finder.exe` is uploaded as the **Keyword-Finder-Windows** artifact.

## Privacy and local data

Keyword Finder processes documents locally. In normal mode, settings, cache files, and rotating logs are stored under `~/.keyword_finder/`. Portable mode stores them in `.keyword_finder_data/` beside the program.

## Notes

- Excel searches stored cell values and cached formula results, not formula expressions.
- Non-PDF formats beyond CSV and supported Excel files depend on Kreuzberg.
- Packaged Windows builds currently use one-file mode and may take a few seconds to open on first launch.

## License

Released under the [MIT License](LICENSE.md).
