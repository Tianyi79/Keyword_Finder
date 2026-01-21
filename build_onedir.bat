@echo off
setlocal

REM Build script (onedir) for Keyword Finder
REM Output: dist\KeywordFinder\KeywordFinder.exe

echo Building Keyword Finder (onedir)...

REM Use python -m to avoid PATH / multiple-Python issues
python -m pip install -U pip
python -m pip install -r requirements.txt

REM Clean build artifacts for reproducibility
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

python -m PyInstaller --clean -y keyword_finder_onedir.spec

echo.
echo Build complete!
echo Output folder: dist\KeywordFinder
pause
