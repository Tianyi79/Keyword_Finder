#!/usr/bin/env bash
set -e

# Build script (onedir) for Keyword Finder
# Output: dist/KeywordFinder/

echo "Building Keyword Finder (onedir)..."

python3 -m pip install -U pip
python3 -m pip install -r requirements.txt

rm -rf build dist

python3 -m PyInstaller --clean -y keyword_finder_onedir.spec

echo
echo "Build complete!"
echo "Output folder: dist/KeywordFinder"
