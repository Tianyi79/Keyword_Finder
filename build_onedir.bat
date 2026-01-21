@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM Build script (onedir) for Keyword Finder
REM Output: dist\KeywordFinder\KeywordFinder.exe

echo Building Keyword Finder (onedir)...

REM Choose ONE Qt binding. Default is PySide6.
REM Set QT_BINDING=PyQt5 before running this script if your app uses PyQt5.
if "%QT_BINDING%"=="" set "QT_BINDING=PySide6"

echo Using QT_BINDING=%QT_BINDING%

REM Create an isolated venv to avoid picking up globally-installed Qt bindings
if exist .venv-build (
  echo Removing existing .venv-build...
  rmdir /s /q .venv-build
)

python -m venv .venv-build
if errorlevel 1 (
  echo Failed to create venv. Ensure you are using Python 3.10+ with venv available.
  exit /b 1
)

call .venv-build\Scripts\activate.bat

REM Install dependencies
python -m pip install -U pip
python -m pip install -r requirements.txt

REM PyInstaller cannot freeze multiple Qt bindings at once. Keep ONLY one.
if /I "%QT_BINDING%"=="PyQt5" (
  python -m pip uninstall -y PySide6 PySide2 PyQt6 2>nul
) else (
  REM Default: PySide6
  python -m pip uninstall -y PyQt5 PyQt6 PySide2 2>nul
)

REM Clean build artifacts for reproducibility
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM NOTE: when building from a .spec, do not pass makespec-only options.
python -m PyInstaller --clean -y keyword_finder_onedir.spec

echo.
echo Build complete!
echo Output folder: dist\KeywordFinder
pause
