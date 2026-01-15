"""
Hybrid Keyword Workbench (Kreuzberg + PDF Preview)

Goal (合体版)
- One GUI for multi-format keyword search (PDF / Word / PPT / Excel / …)
- Uses kreuzberg to extract text for non-PDF files (fast, unified)
- Uses PyMuPDF (fitz) for PDFs to provide:
  - page number
  - bbox-based highlight preview (click result -> jump to page image + red box)
  - "Detected" word + snippet aligned to the hit line

Install
    python -m pip install PySide6 kreuzberg pymupdf pillow openpyxl

Run
    python hybrid_keyword_workbench.py

Export
- CSV / XLSX include page & bbox for PDF hits, and line_no for non-PDF hits.

Notes
- Non-PDF “line number” is based on extracted plain text (content.splitlines()).
- Scanned PDFs without a text layer may return no hits unless OCR is enabled elsewhere.
"""

from __future__ import annotations

import sys
import csv
import re
from collections import OrderedDict
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple

from pathlib import Path
from datetime import datetime

import fitz  # PyMuPDF
from PIL import Image, ImageQt, ImageDraw
from openpyxl import Workbook

from kreuzberg import batch_extract_files_sync, ExtractionConfig

from PySide6.QtCore import Qt, QThread, Signal, QTimer, QUrl
from PySide6.QtGui import QPixmap, QDesktopServices
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog,
    QHBoxLayout, QVBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView,
    QListWidget, QListWidgetItem, QTextEdit, QSplitter,
    QAbstractItemView, QCheckBox, QScrollArea, QGroupBox
)


# -------------------- Data model --------------------

@dataclass
class Hit:
    file_path: str
    file_type: str              # "PDF" or "DOC/PPT/XLS/…"
    keyword: str
    detected_word: str
    page: Optional[int]         # PDF only (1-based)
    line_no: Optional[int]      # non-PDF (1-based)
    rect: Optional[Tuple[float, float, float, float]]  # PDF bbox in page coords
    snippet: str                # PDF: hit line, non-PDF: hit line
    context: str                # PDF: prev/hit/next lines (best-effort), non-PDF same


# -------------------- Worker thread --------------------

class SearchWorker(QThread):
    progress = Signal(str)
    finished_ok = Signal(list)   # List[Hit]
    finished_err = Signal(str)

    def __init__(self, files: List[str], keywords: List[str], case_sensitive: bool = True, whole_word: bool = False):
        super().__init__()
        self.files = files
        self.keywords = keywords
        self.case_sensitive = case_sensitive
        self.whole_word = whole_word

    def run(self):
        try:
            if not self.files:
                self.finished_err.emit("No files selected.")
                return
            if not self.keywords:
                self.finished_err.emit("No keywords provided.")
                return

            pdfs = [f for f in self.files if Path(f).suffix.lower() == ".pdf"]
            others = [f for f in self.files if Path(f).suffix.lower() != ".pdf"]

            hits: List[Hit] = []

            # --- PDFs: PyMuPDF for page + bbox preview ---
            for idx, fpath in enumerate(pdfs, 1):
                self.progress.emit(f"PDF search: {Path(fpath).name} ({idx}/{len(pdfs)})")
                try:
                    doc = fitz.open(fpath)
                except Exception as e:
                    self.progress.emit(f"Failed to open PDF: {Path(fpath).name} — {e}")
                    continue

                for pno in range(len(doc)):
                    page = doc.load_page(pno)
                    words = page.get_text("words") or []  # [x0,y0,x1,y1,word,block,line,wordno]
                    line_text, ordered_lines = build_line_index(words)

                    # search_for is case-sensitive by default; for most CN keywords it's fine.
                    # For English, we implement an optional case-insensitive mode by doing two passes.
                    if self.case_sensitive:
                        kws = self.keywords
                    else:
                        kws = self.keywords  # still use kw; we will compare later at word/line stage

                    for kw in kws:
                        if not kw:
                            continue
                        rects = page.search_for(kw)
                        if not rects and (not self.case_sensitive):
                            # crude fallback: if case-insensitive for ASCII, try both cases
                            rects = page.search_for(kw.lower()) + page.search_for(kw.upper())

                        for r in rects:
                            best = best_word_entry_from_rect(words, r)
                            if best is None:
                                detected, blk, ln = kw, -1, -1
                                snippet = ""
                                ctx = ""
                            else:
                                detected, blk, ln = best
                                snippet = line_text.get((blk, ln), "")
                                ctx = context_from_line_key(line_text, ordered_lines, (blk, ln), window=1)

                            if not _whole_word_accept(kw, detected, self.case_sensitive, self.whole_word):
                                continue
                            hits.append(Hit(
                                file_path=fpath,
                                file_type="PDF",
                                keyword=kw,
                                detected_word=detected or kw,
                                page=pno + 1,
                                line_no=None,
                                rect=(r.x0, r.y0, r.x1, r.y1),
                                snippet=snippet,
                                context=ctx or snippet
                            ))

                try:
                    doc.close()
                except Exception:
                    pass

            # --- Other formats: kreuzberg extraction + line search ---
            if others:
                self.progress.emit("Extracting non-PDF files with kreuzberg…")
                cfg = ExtractionConfig()
                results = batch_extract_files_sync(others, config=cfg)

                for i, result in enumerate(results):
                    fpath = others[i]
                    self.progress.emit(f"Text search: {Path(fpath).name} ({i+1}/{len(others)})")

                    content = getattr(result, "content", "") or ""
                    lines = content.splitlines()

                    for ln, raw_line in enumerate(lines, 1):
                        line = raw_line.strip()
                        if not line:
                            continue

                        if self.case_sensitive:
                            hay = line
                            for kw in self.keywords:
                                if kw and (kw in hay):
                                    detected = detect_token(line, kw) or kw
                                    if not _whole_word_accept(kw, detected, self.case_sensitive, self.whole_word):
                                        continue
                                    ctx = nonpdf_context(lines, ln, window=1)
                                    hits.append(Hit(
                                        file_path=fpath,
                                        file_type="DOC/PPT/XLS/…",
                                        keyword=kw,
                                        detected_word=detected,
                                        page=None,
                                        line_no=ln,
                                        rect=None,
                                        snippet=line,
                                        context=ctx
                                    ))
                        else:
                            hay = line.lower()
                            for kw in self.keywords:
                                if not kw:
                                    continue
                                if kw.lower() in hay:
                                    detected = detect_token(line, kw) or kw
                                    if not _whole_word_accept(kw, detected, self.case_sensitive, self.whole_word):
                                        continue
                                    ctx = nonpdf_context(lines, ln, window=1)
                                    hits.append(Hit(
                                        file_path=fpath,
                                        file_type="DOC/PPT/XLS/…",
                                        keyword=kw,
                                        detected_word=detected,
                                        page=None,
                                        line_no=ln,
                                        rect=None,
                                        snippet=line,
                                        context=ctx
                                    ))

            self.progress.emit(f"Done. {len(hits)} hit(s).")
            self.finished_ok.emit(hits)

        except Exception as e:
            self.finished_err.emit(str(e))


# -------------------- Helpers (PDF line alignment) --------------------

def build_line_index(words) -> Tuple[Dict[Tuple[int, int], str], List[Tuple[int, int]]]:
    if not words:
        return {}, []

    tmp: Dict[Tuple[int, int], List[Tuple[int, float, float, str]]] = {}
    line_pos: Dict[Tuple[int, int], Tuple[float, float]] = {}

    for w in words:
        if len(w) < 8:
            continue
        x0, y0, x1, y1, txt, blk, ln, wordno = w[0], w[1], w[2], w[3], w[4], int(w[5]), int(w[6]), int(w[7])
        key = (blk, ln)
        tmp.setdefault(key, []).append((wordno, float(x0), float(y0), str(txt)))
        if key not in line_pos:
            line_pos[key] = (float(y0), float(x0))
        else:
            cy, cx = line_pos[key]
            line_pos[key] = (min(cy, float(y0)), min(cx, float(x0)))

    line_text: Dict[Tuple[int, int], str] = {}
    for key, items in tmp.items():
        items.sort(key=lambda t: (t[0], t[1]))
        line_text[key] = " ".join([t[3] for t in items]).strip()

    ordered_lines = sorted(
        line_text.keys(),
        key=lambda k: (line_pos.get(k, (1e9, 1e9))[0], line_pos.get(k, (1e9, 1e9))[1], k[0], k[1])
    )
    return line_text, ordered_lines


def best_word_entry_from_rect(words, rect: fitz.Rect) -> Optional[Tuple[str, int, int]]:
    if not words:
        return None
    best = None
    best_score = 0.0
    for w in words:
        if len(w) < 8:
            continue
        x0, y0, x1, y1, txt, blk, ln = w[0], w[1], w[2], w[3], w[4], int(w[5]), int(w[6])
        wr = fitz.Rect(x0, y0, x1, y1)
        inter = rect & wr
        if inter.is_empty:
            continue
        score = max(0.0, inter.get_area())
        if score > best_score:
            best_score = score
            best = (str(txt), blk, ln)
    return best


def context_from_line_key(line_text: Dict[Tuple[int, int], str],
                          ordered_lines: List[Tuple[int, int]],
                          key: Tuple[int, int],
                          window: int = 1) -> str:
    if not line_text or not ordered_lines:
        return ""
    if key not in line_text:
        return ""
    try:
        idx = ordered_lines.index(key)
    except ValueError:
        return ""
    start = max(0, idx - window)
    end = min(len(ordered_lines), idx + window + 1)
    out = []
    for i in range(start, end):
        k = ordered_lines[i]
        prefix = ">> " if k == key else "   "
        out.append(prefix + line_text.get(k, ""))
    return "\n".join(out).strip()


# -------------------- Helpers (non-PDF detected token + context) --------------------

_word_re = re.compile(r"[A-Za-z0-9_]+|[\u4e00-\u9fff]+")

def detect_token(line: str, kw: str) -> str:
    """
    Best-effort "detected word" for non-PDF lines:
    - If the line contains the keyword, return the token (word/Chinese chunk) that contains it.
    - Falls back to the keyword itself.
    """
    if not line or not kw:
        return kw
    idx = line.lower().find(kw.lower())
    if idx < 0:
        return kw

    # Tokenize and find token containing the match span
    tokens = []
    for m in _word_re.finditer(line):
        tokens.append((m.start(), m.end(), m.group(0)))
    span = (idx, idx + len(kw))
    for s, e, tok in tokens:
        if s <= span[0] < e or s < span[1] <= e or (span[0] <= s and e <= span[1]):
            return tok
    return kw



def _is_ascii_word(s: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z0-9_]+", s or ""))

def _eq(a: str, b: str, case_sensitive: bool) -> bool:
    if case_sensitive:
        return (a or "") == (b or "")
    return (a or "").lower() == (b or "").lower()

def _whole_word_accept(keyword: str, detected_word: str, case_sensitive: bool, enabled: bool) -> bool:
    """
    Whole-word filter (English/ASCII only).
    If enabled and keyword is ASCII word-like, only accept hits where the token we detected equals the keyword.
    This prevents cases like searching 'art' but hitting 'departure'.
    """
    if not enabled:
        return True
    if not _is_ascii_word(keyword):
        return True
    if not detected_word:
        return False
    return _eq(keyword, detected_word, case_sensitive)

def nonpdf_context(lines: List[str], line_no_1based: int, window: int = 1) -> str:
    if not lines:
        return ""
    i = max(1, min(line_no_1based, len(lines)))
    start = max(1, i - window)
    end = min(len(lines), i + window)
    out = []
    for ln in range(start, end + 1):
        prefix = ">> " if ln == i else "   "
        out.append(prefix + (lines[ln - 1].strip()))
    return "\n".join(out).strip()


# -------------------- Main window --------------------

class HybridKeywordWorkbench(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Hybrid Keyword Workbench (Multi-format)")
        self.setMinimumSize(1200, 720)
        # Lightweight style polish (no hard-coded colors beyond subtle borders)
        self.setStyleSheet(
            "QGroupBox { font-weight: 600; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }"
            "QPushButton { padding: 6px 10px; }"
            "QTableWidget { selection-behavior: select; }"
            "QHeaderView::section { font-weight: 600; padding: 6px; }"
        )

        self.files: List[str] = []
        self.hits: List[Hit] = []
        # Caches for smoother PDF preview
        self._doc_cache: OrderedDict[str, fitz.Document] = OrderedDict()
        self._page_cache: OrderedDict[tuple, Image.Image] = OrderedDict()  # (path, page, zoom_key) -> PIL Image
        self._doc_cache_max = 6
        self._page_cache_max = 24
        self.worker: Optional[SearchWorker] = None

        # preview state
        self.zoom: float = 2.0
        self._panning: bool = False
        self._pan_start = None

        self._build_ui()

        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.timeout.connect(self._rerender_current_selection)

        self.preview_label.installEventFilter(self)

    # ---------- UI ----------
    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)

        layout = QHBoxLayout(root)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        layout.addWidget(splitter)

        # ---------------- Left: inputs ----------------
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(10)
        splitter.addWidget(left)

        # Files group
        files_group = QGroupBox("Files")
        fg = QVBoxLayout(files_group)
        fg.setSpacing(8)

        btn_row = QHBoxLayout()
        self.btn_add = QPushButton("Add files")
        self.btn_remove = QPushButton("Remove")
        self.btn_clear = QPushButton("Clear")
        self.btn_add.setToolTip("Select multiple files (PDF/DOCX/PPTX/XLSX/…)")
        self.btn_remove.setToolTip("Remove selected files from the list")
        self.btn_clear.setToolTip("Clear file list and results")
        btn_row.addWidget(self.btn_add)
        btn_row.addWidget(self.btn_remove)
        btn_row.addWidget(self.btn_clear)
        btn_row.addStretch(1)
        fg.addLayout(btn_row)

        self.btn_add.clicked.connect(self.add_files)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear.clicked.connect(self.clear_files)

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.file_list.setToolTip("Selected input files")
        fg.addWidget(self.file_list, 1)

        left_layout.addWidget(files_group, 2)

        # Search group
        search_group = QGroupBox("Search")
        sg = QVBoxLayout(search_group)
        sg.setSpacing(8)

        sg.addWidget(QLabel("Keywords (one per line or comma-separated):"))
        self.kw_box = QTextEdit()
        self.kw_box.setPlaceholderText("合同\n协议\n签署\nor: 合同, 协议, 签署")
        self.kw_box.setFixedHeight(120)
        sg.addWidget(self.kw_box)

        opt_row = QHBoxLayout()
        self.case_cb = QCheckBox("Case sensitive (English)")
        self.case_cb.setChecked(True)
        self.whole_word_cb = QCheckBox("Whole word (English)")
        self.whole_word_cb.setChecked(True)
        self.whole_word_cb.setToolTip("When enabled, English/ASCII keywords only match whole tokens (reduces false hits like 'art' in 'departure').")
        opt_row.addWidget(self.case_cb)
        opt_row.addWidget(self.whole_word_cb)
        opt_row.addStretch(1)
        sg.addLayout(opt_row)

        run_row = QHBoxLayout()
        self.btn_run = QPushButton("Search")
        self.btn_export_csv = QPushButton("Export CSV")
        self.btn_export_xlsx = QPushButton("Export XLSX")
        self.btn_run.setToolTip("Run keyword search on all selected files")
        self.btn_export_csv.setToolTip("Export current results as CSV")
        self.btn_export_xlsx.setToolTip("Export current results as XLSX")
        run_row.addWidget(self.btn_run)
        run_row.addStretch(1)
        run_row.addWidget(self.btn_export_csv)
        run_row.addWidget(self.btn_export_xlsx)
        sg.addLayout(run_row)

        self.btn_run.clicked.connect(self.run_search)
        self.btn_export_csv.clicked.connect(self.export_csv)
        self.btn_export_xlsx.clicked.connect(self.export_xlsx)

        left_layout.addWidget(search_group, 1)

        # Status (kept compact)
        status_group = QGroupBox("Status")
        st = QVBoxLayout(status_group)
        st.setContentsMargins(10, 10, 10, 10)
        self.status = QLabel("Ready.")
        self.status.setWordWrap(True)
        st.addWidget(self.status)
        left_layout.addWidget(status_group, 0)

        # ---------------- Right: results + preview ----------------
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)
        splitter.addWidget(right)

        # Results table
        results_group = QGroupBox("Results")
        rg = QVBoxLayout(results_group)
        rg.setSpacing(8)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["Type", "File", "Keyword", "Detected", "Page", "Line #", "Snippet"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.cellClicked.connect(self.on_row_clicked)
        self.table.cellDoubleClicked.connect(lambda r, c: self.open_hit_in_viewer(r))
        rg.addWidget(self.table)

        right_layout.addWidget(results_group, 2)

        # Preview area
        preview_group = QGroupBox("Preview")
        pg = QVBoxLayout(preview_group)
        pg.setSpacing(8)

        toolbar = QHBoxLayout()
        toolbar.addWidget(QLabel("Hits:"))
        self.btn_prev_hit = QPushButton("Prev")
        self.btn_next_hit = QPushButton("Next")
        for b in (self.btn_prev_hit, self.btn_next_hit):
            b.setFixedHeight(32)
            b.setFixedWidth(72)
        
        toolbar.addWidget(self.btn_prev_hit)
        toolbar.addWidget(self.btn_next_hit)
        toolbar.addSpacing(12)
        toolbar.addWidget(QLabel("Zoom:"))
        self.btn_zoom_out = QPushButton("−")
        self.btn_zoom_in = QPushButton("+")
        self.btn_zoom_reset = QPushButton("100%")
        for b in (self.btn_zoom_out, self.btn_zoom_in, self.btn_zoom_reset):
            b.setFixedHeight(32)
        self.btn_zoom_out.setFixedWidth(44)
        self.btn_zoom_in.setFixedWidth(44)
        self.btn_zoom_reset.setFixedWidth(64)

        self.btn_zoom_out.setToolTip("Zoom out (Ctrl + Mouse Wheel)")
        self.btn_zoom_in.setToolTip("Zoom in (Ctrl + Mouse Wheel)")
        self.btn_zoom_reset.setToolTip("Reset zoom")

        toolbar.addWidget(self.btn_zoom_out)
        toolbar.addWidget(self.btn_zoom_in)
        toolbar.addWidget(self.btn_zoom_reset)
        toolbar.addStretch(1)

        self.btn_continue_reading = QPushButton("Continue reading")
        self.btn_continue_reading.setEnabled(False)
        self.btn_continue_reading.setToolTip("Open the PDF in your default viewer at the hit page, so you can keep reading")
        self.btn_continue_reading.setFixedHeight(32)
        toolbar.addWidget(self.btn_continue_reading)

        pg.addLayout(toolbar)

        self.btn_prev_hit.clicked.connect(self.select_prev_hit)
        self.btn_next_hit.clicked.connect(self.select_next_hit)

        self.btn_zoom_in.clicked.connect(lambda: self._set_zoom(self.zoom * 1.25))
        self.btn_zoom_out.clicked.connect(lambda: self._set_zoom(self.zoom / 1.25))
        self.btn_zoom_reset.clicked.connect(lambda: self._set_zoom(2.0))
        self.btn_continue_reading.clicked.connect(self.open_selected_in_viewer)

        # Scrollable preview canvas
        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.preview_label.setStyleSheet("background: #111;")
        self.preview_label.setMinimumSize(680, 420)

        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidgetResizable(False)
        self.preview_scroll.setStyleSheet("QScrollArea { border: 1px solid rgba(0,0,0,0.15); }")
        self.preview_scroll.setWidget(self.preview_label)
        pg.addWidget(self.preview_scroll, 1)

        right_layout.addWidget(preview_group, 3)

        # Context area
        context_group = QGroupBox("Context")
        cg = QVBoxLayout(context_group)
        cg.setSpacing(6)
        self.context_box = QTextEdit()
        self.context_box.setReadOnly(True)
        self.context_box.setPlaceholderText("Context lines will appear here.")
        self.context_box.setFixedHeight(170)
        cg.addWidget(self.context_box)
        right_layout.addWidget(context_group, 1)

        splitter.setSizes([420, 1080])

    # ---------- Events (zoom/pan on PDF preview) ----------
    def _set_zoom(self, new_zoom: float):
        self.zoom = max(0.5, min(8.0, float(new_zoom)))
        self._rerender_current_selection()

    def eventFilter(self, obj, event):
        if obj is self.preview_label:
            et = event.type()

            # Ctrl+wheel zoom
            if et == event.Type.Wheel and (event.modifiers() & Qt.ControlModifier):
                delta = event.angleDelta().y()
                factor = 1.15 if delta > 0 else 1 / 1.15
                self._set_zoom(self.zoom * factor)
                return True

            # drag-to-pan
            if et == event.Type.MouseButtonPress and event.button() == Qt.LeftButton:
                self._panning = True
                self._pan_start = event.globalPosition().toPoint()
                self.preview_label.setCursor(Qt.ClosedHandCursor)
                return True

            if et == event.Type.MouseMove and self._panning:
                cur = event.globalPosition().toPoint()
                delta = cur - self._pan_start
                self._pan_start = cur
                hbar = self.preview_scroll.horizontalScrollBar()
                vbar = self.preview_scroll.verticalScrollBar()
                hbar.setValue(hbar.value() - delta.x())
                vbar.setValue(vbar.value() - delta.y())
                return True

            if et == event.Type.MouseButtonRelease and event.button() == Qt.LeftButton:
                self._panning = False
                self._pan_start = None
                self.preview_label.setCursor(Qt.ArrowCursor)
                return True

        return super().eventFilter(obj, event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._resize_timer.start(120)

    # ---------- File handling ----------
    def add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select files",
            "",
            "All files (*.*);;PDF (*.pdf);;Word (*.docx);;PowerPoint (*.pptx);;Excel (*.xlsx *.xlsm)"
        )
        if not paths:
            return
        added = 0
        for p in paths:
            p = str(Path(p))
            if p not in self.files:
                self.files.append(p)
                self.file_list.addItem(QListWidgetItem(p))
                added += 1
        if added:
            self.status.setText(f"Selected {len(self.files)} file(s).")

    def remove_selected(self):
        items = self.file_list.selectedItems()
        if not items:
            return
        to_remove = {it.text() for it in items}
        self.files = [f for f in self.files if f not in to_remove]
        self.file_list.clear()
        for f in self.files:
            self.file_list.addItem(QListWidgetItem(f))
        self.status.setText(f"Selected {len(self.files)} file(s).")

    def clear_files(self):
        self.files = []
        self.file_list.clear()
        self.status.setText("Ready.")
        self.hits = []
        self.table.setRowCount(0)
        self.preview_label.clear()
        self.context_box.clear()
        self.btn_continue_reading.setEnabled(False)
        self._clear_pdf_caches()

    # ---------- Keywords ----------
    def _parse_keywords(self) -> List[str]:
        raw = self.kw_box.toPlainText().strip()
        if not raw:
            return []
        parts: List[str] = []
        for line in raw.splitlines():
            for p in line.split(","):
                k = p.strip()
                if k:
                    parts.append(k)

        # de-dup
        seen = set()
        out = []
        for k in parts:
            key = k if self.case_cb.isChecked() else k.lower()
            if key not in seen:
                seen.add(key)
                out.append(k)
        return out

    # ---------- Search ----------
    def run_search(self):
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "Running", "Search is already running.")
            return

        if not self.files:
            QMessageBox.information(self, "Tip", "Please add files first.")
            return

        keywords = self._parse_keywords()
        if not keywords:
            QMessageBox.information(self, "Tip", "Please enter at least one keyword.")
            return

        self.table.setRowCount(0)
        self.hits = []
        self.preview_label.clear()
        self.context_box.clear()

        self.status.setText("Starting…")
        self.btn_run.setEnabled(False)

        self.worker = SearchWorker(self.files, keywords, case_sensitive=self.case_cb.isChecked(), whole_word=self.whole_word_cb.isChecked())
        self.worker.progress.connect(self.status.setText)
        self.worker.finished_ok.connect(self._on_finished_ok)
        self.worker.finished_err.connect(self._on_finished_err)
        self.worker.start()

    def _on_finished_ok(self, hits: List[Hit]):
        self.hits = hits
        self._populate_table(hits)
        self.btn_run.setEnabled(True)

        if hits:
            self.table.selectRow(0)
            self.on_row_clicked(0, 0)

    def _on_finished_err(self, msg: str):
        self.btn_run.setEnabled(True)
        QMessageBox.critical(self, "Error", msg)
        self.status.setText("Error.")

    def _populate_table(self, hits: List[Hit]):
        self.table.setRowCount(len(hits))
        for i, h in enumerate(hits):
            self.table.setItem(i, 0, QTableWidgetItem(h.file_type))
            self.table.setItem(i, 1, QTableWidgetItem(Path(h.file_path).name))
            self.table.setItem(i, 2, QTableWidgetItem(h.keyword))
            self.table.setItem(i, 3, QTableWidgetItem(h.detected_word))
            self.table.setItem(i, 4, QTableWidgetItem("" if h.page is None else str(h.page)))
            self.table.setItem(i, 5, QTableWidgetItem("" if h.line_no is None else str(h.line_no)))
            self.table.setItem(i, 6, QTableWidgetItem(h.snippet))
        self.status.setText(f"Done. {len(hits)} hit(s).")
        self.setWindowTitle(f"Hybrid Keyword Workbench — {len(hits)} hit(s)")

    # ---------- Preview ----------
    def on_row_clicked(self, row: int, col: int):
        if row < 0 or row >= len(self.hits):
            return
        hit = self.hits[row]
        self.btn_continue_reading.setEnabled(hit.file_type == "PDF" and hit.page is not None)
        self.context_box.setPlainText(hit.context or hit.snippet or "")

        if hit.file_type != "PDF" or hit.page is None or hit.rect is None:
            # Non-PDF: no image preview
            self.preview_label.clear()
            self.preview_label.setText("Preview: (non-PDF)\n\nNo page image available.\nUse the context box below.")
            self.preview_label.setStyleSheet("background: #111; color: #ddd; padding: 12px;")
            self.preview_label.adjustSize()
            return

        self._render_pdf_hit(hit)


    # ---------- Hit navigation ----------
    def select_next_hit(self):
        if not self.hits:
            return
        row = self.table.currentRow()
        next_row = 0 if row < 0 else min(len(self.hits) - 1, row + 1)
        self.table.selectRow(next_row)
        self.table.scrollToItem(self.table.item(next_row, 0), QAbstractItemView.PositionAtCenter)
        self.on_row_clicked(next_row, 0)

    def select_prev_hit(self):
        if not self.hits:
            return
        row = self.table.currentRow()
        prev_row = 0 if row < 0 else max(0, row - 1)
        self.table.selectRow(prev_row)
        self.table.scrollToItem(self.table.item(prev_row, 0), QAbstractItemView.PositionAtCenter)
        self.on_row_clicked(prev_row, 0)

    def keyPressEvent(self, event):
        # Avoid hijacking typing when focus is inside text boxes
        fw = QApplication.focusWidget()
        if fw is self.kw_box: 
            return super().keyPressEvent(event)

        k = event.key()
        if k == Qt.Key_J:
            self.select_next_hit()
            return
        if k == Qt.Key_K:
            self.select_prev_hit()
            return
        if k in (Qt.Key_Return, Qt.Key_Enter):
            if self.btn_continue_reading.isEnabled():
                self.open_selected_in_viewer()
                return
        return super().keyPressEvent(event)


    # ---------- Continue reading (open in external PDF viewer) ----------
    def open_selected_in_viewer(self):
        """Open the currently selected PDF hit in the system default PDF viewer at its page."""
        row = self.table.currentRow()
        if row < 0 or row >= len(self.hits):
            QMessageBox.information(self, "Tip", "Select a result row first.")
            return
        self.open_hit_in_viewer(row)

    def open_hit_in_viewer(self, row: int):
        """Open a specific hit row in the system PDF viewer (best-effort jump to page)."""
        if row < 0 or row >= len(self.hits):
            return

        hit = self.hits[row]
        if hit.file_type != "PDF" or not hit.page:
            QMessageBox.information(self, "Tip", "Continue reading is available for PDF hits only.")
            return

        pdf_path = str(Path(hit.file_path).resolve())

        # Many PDF viewers/browsers accept URL fragments like #page=12.
        url = QUrl.fromLocalFile(pdf_path)
        url.setFragment(f"page={hit.page}")

        if not QDesktopServices.openUrl(url):
            # Fallback: open without page fragment
            if not QDesktopServices.openUrl(QUrl.fromLocalFile(pdf_path)):
                QMessageBox.warning(
                    self,
                    "Open failed",
                    "Could not open the PDF in a viewer.\n\n"
                    "Try opening the file manually:\n" + pdf_path
                )

    def _rerender_current_selection(self):
        row = self.table.currentRow()
        if 0 <= row < len(self.hits):
            hit = self.hits[row]
            if hit.file_type == "PDF" and hit.page and hit.rect:
                self._render_pdf_hit(hit)

    
    # ---------- PDF caching (smooth preview) ----------
    def _clear_pdf_caches(self):
        # close cached docs
        try:
            for _, d in list(self._doc_cache.items()):
                try:
                    d.close()
                except Exception:
                    pass
        finally:
            self._doc_cache.clear()
            self._page_cache.clear()

    def _get_doc_cached(self, pdf_path: str) -> fitz.Document:
        key = str(Path(pdf_path).resolve())
        if key in self._doc_cache:
            d = self._doc_cache.pop(key)
            self._doc_cache[key] = d
            return d

        d = fitz.open(key)
        self._doc_cache[key] = d
        while len(self._doc_cache) > self._doc_cache_max:
            old_key, old_doc = self._doc_cache.popitem(last=False)
            try:
                old_doc.close()
            except Exception:
                pass
        return d

    def _get_page_image_cached(self, pdf_path: str, page_1based: int, zoom: float) -> Image.Image:
        key = (str(Path(pdf_path).resolve()), int(page_1based), int(round(zoom * 100)))
        if key in self._page_cache:
            img = self._page_cache.pop(key)
            self._page_cache[key] = img
            return img

        doc = self._get_doc_cached(pdf_path)
        page = doc.load_page(page_1based - 1)
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        self._page_cache[key] = img
        while len(self._page_cache) > self._page_cache_max:
            self._page_cache.popitem(last=False)
        return img

    def _render_pdf_hit(self, hit: Hit):
        try:
            base = self._get_page_image_cached(hit.file_path, hit.page, self.zoom)
            img = base.copy()  # keep cache immutable

            draw = ImageDraw.Draw(img)
            x0, y0, x1, y1 = hit.rect
            rr = fitz.Rect(x0 * self.zoom, y0 * self.zoom, x1 * self.zoom, y1 * self.zoom)
            draw.rectangle([rr.x0, rr.y0, rr.x1, rr.y1], outline="red", width=max(2, int(3 * self.zoom)))

            qimg = ImageQt.ImageQt(img)
            pm = QPixmap.fromImage(qimg)
            self.preview_label.setStyleSheet("background: #111;")
            self.preview_label.setPixmap(pm)
            self.preview_label.setFixedSize(pm.size())

        except Exception as e:
            self.preview_label.clear()
            self.preview_label.setText(f"Failed to render PDF preview:\n{e}")
            self.preview_label.adjustSize()

    # ---------- Export ----------
    def export_csv(self):
        if not self.hits:
            QMessageBox.information(self, "Tip", "No results to export.")
            return
        default = f"keyword_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        out, _ = QFileDialog.getSaveFileName(self, "Save CSV", default, "CSV (*.csv)")
        if not out:
            return
        try:
            with open(out, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["file", "type", "keyword", "detected_word", "page", "line_no", "x0", "y0", "x1", "y1", "snippet"])
                for h in self.hits:
                    x0 = y0 = x1 = y1 = ""
                    if h.rect:
                        x0, y0, x1, y1 = h.rect
                    w.writerow([
                        h.file_path, h.file_type, h.keyword, h.detected_word,
                        "" if h.page is None else h.page,
                        "" if h.line_no is None else h.line_no,
                        x0, y0, x1, y1,
                        h.snippet
                    ])
            QMessageBox.information(self, "Export", f"Saved:\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "Export failed", str(e))

    def export_xlsx(self):
        if not self.hits:
            QMessageBox.information(self, "Tip", "No results to export.")
            return
        default = f"keyword_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        out, _ = QFileDialog.getSaveFileName(self, "Save XLSX", default, "Excel (*.xlsx)")
        if not out:
            return
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "results"
            ws.append(["file", "type", "keyword", "detected_word", "page", "line_no", "x0", "y0", "x1", "y1", "snippet"])
            for h in self.hits:
                x0 = y0 = x1 = y1 = None
                if h.rect:
                    x0, y0, x1, y1 = h.rect
                ws.append([
                    h.file_path, h.file_type, h.keyword, h.detected_word,
                    h.page, h.line_no, x0, y0, x1, y1, h.snippet
                ])

            ws.column_dimensions["A"].width = 46
            ws.column_dimensions["B"].width = 14
            ws.column_dimensions["C"].width = 18
            ws.column_dimensions["D"].width = 20
            ws.column_dimensions["E"].width = 8
            ws.column_dimensions["F"].width = 8
            ws.column_dimensions["G"].width = 10
            ws.column_dimensions["H"].width = 10
            ws.column_dimensions["I"].width = 10
            ws.column_dimensions["J"].width = 10
            ws.column_dimensions["K"].width = 90

            wb.save(out)
            QMessageBox.information(self, "Export", f"Saved:\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "Export failed", str(e))


def main():
    app = QApplication(sys.argv)
    w = HybridKeywordWorkbench()
    w.resize(1500, 820)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
