# -*- mode: python ; coding: utf-8 -*-

"""PyInstaller spec (onedir) for Keyword Finder.

Outputs (examples):
  - Windows: dist/KeywordFinder/KeywordFinder.exe
  - Linux:   dist/KeywordFinder/KeywordFinder
  - macOS:   dist/KeywordFinder/KeywordFinder (wrap into .app optionally)

Notes:
  - Build on the same OS you want to run on.
  - onedir is recommended for Qt (PySide6) GUI apps.
"""

from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# Collect PySide6 runtime resources (Qt plugins, translations, etc.)
datas = []
datas += collect_data_files('PySide6', include_py_files=False)

hiddenimports = [
    'PySide6.QtCore',
    'PySide6.QtGui',
    'PySide6.QtWidgets',
    # Optional Qt modules if your app ever uses them; harmless if unused.
    'PySide6.QtNetwork',
    'PySide6.QtPrintSupport',
    # App deps
    'kreuzberg',
    'fitz',
    'PIL',
    'openpyxl',
]

# Keep excludes conservative; do not exclude common libs unless you are sure.
excludes = [
    'tensorflow',
    'torch',
]

a = Analysis(
    ['gui_keyword_finder_4.0.1_fix_unpack.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    exclude_binaries=True,
    name='KeywordFinder',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='keywordfinder.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name='KeywordFinder',
)
