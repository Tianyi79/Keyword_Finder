# Keyword Finder - Executable Build

Build standalone executables of the GUI Keyword Finder application for Windows, Linux, and Mac.

## Quick Start:

### Windows:
```cmd
build_onedir.bat
```

### Linux/Mac:
```bash
chmod +x build.sh
./build_onedir.sh
```

## Output Files:
This project uses onedir packaging for stability (recommended for PySide6/Qt apps).
- **Windows**: `dist/KeywordFinder.exe` (~182MB)
- **Linux**: `dist/KeywordFinder/KeywordFinder`
- **Mac**: `dist/KeywordFinder/KeywordFinder`
Onedir outputs a folder containing the executable + bundled dependencies.

## Manual Build:
```bash
python -m pip install -r requirements.txt
python -m pyinstaller keyword_finder_onedir.spec
```

## Distribution:
- Completely standalone - no Python required
- Onedir is typically more reliable than onefile for Qt apps.
- Build separately for each platform you want to distribute to.

## Files:
- `requirements.txt` - Dependencies
- `keyword_finder.spec` - PyInstaller config
- `build.sh` / `build.bat` - Build scripts

## Troubleshooting:
- Python 3.10+ is recommended
- Update pip: `python -m pip install --U pip`
- Install PyInstaller: `python -m pip install pyinstaller`
- App launches but immediately exits (Windows): build once with console to see the error output:
  ` pyinstaller --onedir keyword_finder.spec`
- Then run the executable from dist/KeywordFinder/ and read the console output.
```cmd
cd dist\KeywordFinder
.\KeywordFinder.exe
```
