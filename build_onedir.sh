#!/usr/bin/env bash
set -e

# Build script (onedir) for Keyword Finder
# Output: dist/KeywordFinder/

echo "Building Keyword Finder (onedir)..."

# Build in an isolated venv to avoid picking up globally-installed Qt bindings
QT_BINDING=${QT_BINDING:-PySide6}  # set to PyQt5 if your app uses PyQt5

python3 -m venv .venv-build
source .venv-build/bin/activate

python -m pip install -U pip
python -m pip install -r requirements.txt

# PyInstaller cannot freeze multiple Qt bindings at once. Keep ONLY one.
if [ "$QT_BINDING" = "PyQt5" ]; then
  python -m pip uninstall -y PySide6 PySide2 PyQt6 || true
  QT_EXCLUDES=(--exclude-module PySide6 --exclude-module PySide2 --exclude-module PyQt6)
else
  # Default: PySide6
  python -m pip uninstall -y PyQt5 PyQt6 PySide2 || true
  QT_EXCLUDES=(--exclude-module PyQt5 --exclude-module PyQt6 --exclude-module PySide2)
fi

rm -rf build dist

python -m PyInstaller --clean -y keyword_finder_onedir.spec

echo
echo "Build complete!"
echo "Output folder: dist/KeywordFinder"
