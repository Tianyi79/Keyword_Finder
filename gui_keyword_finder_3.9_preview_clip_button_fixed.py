"""GUI Keyword Finder 3.9

Menu-first UX (no bulky left panels)
- Files and Keywords are managed via dialogs opened from the menubar.
- Preview is a popup window (non-modal). Selecting a result updates the popup.
- Control actions (Run search / Preview nav / Zoom / Continue reading / Clips export)
  live in the menubar.

Install
    python -m pip install PySide6 kreuzberg pymupdf pillow openpyxl

Run
    python gui_keyword_finder_3.7_menu_dialog_preview_popup.py

Export
- Results: CSV / XLSX
- Clips: Markdown (Notion-friendly blockquotes) / CSV

Notes
- PDF preview uses PyMuPDF (fitz) for page rendering + bbox highlight.
- Non-PDF search uses kreuzberg extraction and line-based matching.
"""

from __future__ import annotations

import sys
import csv
import re
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple
from pathlib import Path
from datetime import datetime
from collections import OrderedDict

import fitz  # PyMuPDF
from PIL import Image, ImageQt, ImageDraw
from openpyxl import Workbook

from kreuzberg import batch_extract_files_sync, ExtractionConfig

from PySide6.QtCore import Qt, QThread, Signal, QTimer, QUrl
from PySide6.QtGui import QPixmap, QDesktopServices, QAction
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog,
    QHBoxLayout, QVBoxLayout, QLabel,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView,
    QListWidget, QListWidgetItem, QTextEdit, QSplitter,
    QAbstractItemView, QScrollArea, QGroupBox,
    QDialog, QDialogButtonBox, QPushButton
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


@dataclass
class Clip:
    created_at: str
    file_path: str
    file_type: str
    page: Optional[int]
    line_no: Optional[int]
    keyword: str
    detected_word: str
    selected_text: str
    note: str = ""


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
                    words = page.get_text("words") or []
                    line_text, ordered_lines = build_line_index(words)

                    kws = self.keywords

                    for kw in kws:
                        if not kw:
                            continue
                        rects = page.search_for(kw)
                        if not rects and (not self.case_sensitive):
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
    if not line or not kw:
        return kw
    idx = line.lower().find(kw.lower())
    if idx < 0:
        return kw

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


# -------------------- Dialogs --------------------

class FileManagerDialog(QDialog):
    def __init__(self, parent, files: List[str]):
        super().__init__(parent)
        self.setWindowTitle("Files")
        self.setMinimumWidth(760)
        self._files = list(files)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        self.list = QListWidget()
        self.list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        layout.addWidget(self.list, 1)

        btn_row = QHBoxLayout()
        self.btn_add = QAction("Add")
        self.btn_remove = QAction("Remove")
        self.btn_clear = QAction("Clear")

        # Use simple buttons (QDialog is fine with plain labels)
        from PySide6.QtWidgets import QPushButton
        b_add = QPushButton("Add…")
        b_remove = QPushButton("Remove selected")
        b_clear = QPushButton("Clear all")
        btn_row.addWidget(b_add)
        btn_row.addWidget(b_remove)
        btn_row.addWidget(b_clear)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)

        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(self.buttons)

        b_add.clicked.connect(self._add)
        b_remove.clicked.connect(self._remove)
        b_clear.clicked.connect(self._clear)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)

        self._refresh()

    def _refresh(self):
        self.list.clear()
        for p in self._files:
            self.list.addItem(QListWidgetItem(p))

    def _add(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select files",
            "",
            "All files (*.*);;PDF (*.pdf);;Word (*.docx);;PowerPoint (*.pptx);;Excel (*.xlsx *.xlsm)"
        )
        if not paths:
            return
        for p in paths:
            p = str(Path(p))
            if p not in self._files:
                self._files.append(p)
        self._refresh()

    def _remove(self):
        items = self.list.selectedItems()
        if not items:
            return
        remove = {it.text() for it in items}
        self._files = [f for f in self._files if f not in remove]
        self._refresh()

    def _clear(self):
        self._files = []
        self._refresh()

    def get_files(self) -> List[str]:
        return list(self._files)


class KeywordManagerDialog(QDialog):
    run_search_requested = Signal()

    def __init__(self, parent, raw_text: str, case_sensitive: bool = True, whole_word: bool = False):
        super().__init__(parent)
        self.setWindowTitle("Input Keywords")
        self.setMinimumWidth(680)
        
        self.case_sensitive = case_sensitive
        self.whole_word = whole_word

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        hint = QLabel("One per line (recommended) or comma-separated.\nTip: keep keywords clean (no bullets / prefixes).")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.box = QTextEdit()
        self.box.setPlaceholderText("a\nb\nc\n or: a, b, c")
        self.box.setPlainText(raw_text or "")
        self.box.setMinimumHeight(260)
        layout.addWidget(self.box, 1)

        # Add options from Keywords menu
        options_layout = QHBoxLayout()
        
        from PySide6.QtWidgets import QCheckBox
        self.clear_btn = QPushButton("Clear keywords")
        self.clear_btn.clicked.connect(self._clear_keywords)
        
        self.case_sensitive_cb = QCheckBox("Case sensitive (English)")
        self.case_sensitive_cb.setChecked(self.case_sensitive)
        
        self.whole_word_cb = QCheckBox("Whole word (English)")
        self.whole_word_cb.setChecked(self.whole_word)
        self.whole_word_cb.setToolTip("When enabled, English/ASCII keywords only match whole tokens (reduces false hits like 'art' in 'departure').")
        
        options_layout.addWidget(self.clear_btn)
        options_layout.addStretch()
        options_layout.addWidget(self.case_sensitive_cb)
        options_layout.addWidget(self.whole_word_cb)
        layout.addLayout(options_layout)

        # Add Run Search button
        search_layout = QHBoxLayout()
        self.run_search_btn = QPushButton("▶️ Run Search")
        self.run_search_btn.setStyleSheet("QPushButton { background-color: #0066cc; color: white; font-weight: bold; padding: 8px 16px; }")
        self.run_search_btn.clicked.connect(self._on_run_search)
        search_layout.addStretch()
        search_layout.addWidget(self.run_search_btn)
        layout.addLayout(search_layout)

        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(self.buttons)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)

    def _clear_keywords(self):
        self.box.clear()

    def _on_run_search(self):
        self.run_search_requested.emit()
        self.accept()

    def get_text(self) -> str:
        return self.box.toPlainText()
        
    def get_case_sensitive(self) -> bool:
        return self.case_sensitive_cb.isChecked()
        
    def get_whole_word(self) -> bool:
        return self.whole_word_cb.isChecked()


class PreviewPopup(QDialog):
    """Popup preview window."""

    request_save_clip = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Preview")
        self.setMinimumSize(980, 760)
        self.setWindowFlag(Qt.Window)

        self.current_hit: Optional[Hit] = None
        self.zoom: float = 2.0

        # caches
        self._doc_cache: OrderedDict[str, fitz.Document] = OrderedDict()
        self._page_cache: OrderedDict[tuple, Image.Image] = OrderedDict()
        self._doc_cache_max = 6
        self._page_cache_max = 24

        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.timeout.connect(self._rerender)

        root = QWidget()
        self.setLayout(QVBoxLayout())
        self.layout().setContentsMargins(10, 10, 10, 10)
        self.layout().addWidget(root)

        outer = QVBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(10)

        # Add preview controls from Preview menu
        controls_layout = QHBoxLayout()
        
        self.btn_continue = QPushButton("Continue reading (open in viewer)")
        self.btn_continue.setToolTip("Open in external viewer (Return)")
        self.btn_continue.clicked.connect(self._continue_reading)
        
        self.btn_prev_hit = QPushButton("◀ Previous hit")
        self.btn_prev_hit.setToolTip("Previous hit (K)")
        self.btn_prev_hit.clicked.connect(self._prev_hit)
        
        self.btn_next_hit = QPushButton("Next hit ▶")
        self.btn_next_hit.setToolTip("Next hit (J)")
        self.btn_next_hit.clicked.connect(self._next_hit)
        
        controls_layout.addWidget(self.btn_continue)
        controls_layout.addStretch()
        controls_layout.addWidget(self.btn_prev_hit)
        controls_layout.addWidget(self.btn_next_hit)
        outer.addLayout(controls_layout)

        # Add zoom controls
        zoom_layout = QHBoxLayout()
        self.btn_zoom_out = QPushButton("Zoom out")
        self.btn_zoom_out.setToolTip("Zoom out (Ctrl+-)")
        self.btn_zoom_out.clicked.connect(self.zoom_out)
        
        self.btn_zoom_reset = QPushButton("Zoom reset")
        self.btn_zoom_reset.setToolTip("Zoom reset (Ctrl+0)")
        self.btn_zoom_reset.clicked.connect(self.zoom_reset)
        
        self.btn_zoom_in = QPushButton("Zoom in")
        self.btn_zoom_in.setToolTip("Zoom in (Ctrl++)")
        self.btn_zoom_in.clicked.connect(self.zoom_in)
        
        zoom_layout.addStretch()
        zoom_layout.addWidget(self.btn_zoom_out)
        zoom_layout.addWidget(self.btn_zoom_reset)
        zoom_layout.addWidget(self.btn_zoom_in)
        outer.addLayout(zoom_layout)

        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)
        outer.addWidget(splitter, 1)

        # Top: image preview
        preview_group = QGroupBox("PDF page preview")
        pg = QVBoxLayout(preview_group)
        pg.setSpacing(8)

        self.preview_label = QLabel("Select a PDF hit to preview.")
        self.preview_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.preview_label.setStyleSheet("background: #111; color: #ddd; padding: 10px;")

        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidgetResizable(False)
        self.preview_scroll.setStyleSheet("QScrollArea { border: 1px solid rgba(0,0,0,0.15); }")
        self.preview_scroll.setWidget(self.preview_label)
        pg.addWidget(self.preview_scroll, 1)

        splitter.addWidget(preview_group)

        # Bottom: context + page text + clips
        bottom = QWidget()
        bl = QVBoxLayout(bottom)
        bl.setContentsMargins(0, 0, 0, 0)
        bl.setSpacing(10)

        ctx_group = QGroupBox("Context")
        cg = QVBoxLayout(ctx_group)
        self.context_box = QTextEdit()
        self.context_box.setReadOnly(True)
        self.context_box.setMinimumHeight(120)
        cg.addWidget(self.context_box)
        bl.addWidget(ctx_group)

        txt_group = QGroupBox("Page text (select to clip)")
        tg = QVBoxLayout(txt_group)

        # Inline clip button (requested): keep "Save selection as clip" in the preview window,
        # not in the Clips menu. The actual save logic lives in the main window.
        top_row = QHBoxLayout()
        top_row.addStretch(1)
        self.btn_save_clip = QPushButton("Save selection as clip")
        self.btn_save_clip.setToolTip("Select text in Context or Page text, then save as a clip")
        self.btn_save_clip.clicked.connect(lambda: self.request_save_clip.emit())
        top_row.addWidget(self.btn_save_clip)
        tg.addLayout(top_row)

        self.page_text_box = QTextEdit()
        self.page_text_box.setReadOnly(True)
        self.page_text_box.setMinimumHeight(160)
        tg.addWidget(self.page_text_box)
        bl.addWidget(txt_group)

        splitter.addWidget(bottom)
        splitter.setSizes([520, 240])

    def set_main_window(self, main_window):
        """Set reference to main window for navigation callbacks"""
        self.main_window = main_window

    def _continue_reading(self):
        if hasattr(self, 'main_window') and self.main_window:
            self.main_window.open_selected_in_viewer()

    def _prev_hit(self):
        if hasattr(self, 'main_window') and self.main_window:
            self.main_window.select_prev_hit()

    def _next_hit(self):
        if hasattr(self, 'main_window') and self.main_window:
            self.main_window.select_next_hit()

    # ----- cache helpers -----
    def closeEvent(self, event):
        self._clear_pdf_caches()
        super().closeEvent(event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._resize_timer.start(120)

    def _clear_pdf_caches(self):
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

    # ----- public controls used by main window -----
    def set_zoom(self, zoom: float):
        self.zoom = max(0.5, min(8.0, float(zoom)))
        self._rerender()

    def zoom_in(self):
        self.set_zoom(self.zoom * 1.25)

    def zoom_out(self):
        self.set_zoom(self.zoom / 1.25)

    def zoom_reset(self):
        self.set_zoom(2.0)

    def show_hit(self, hit: Hit):
        self.current_hit = hit
        self.context_box.setPlainText(hit.context or hit.snippet or "")

        # selectable text for clipping
        if hit.file_type == "PDF" and hit.page is not None:
            try:
                doc = self._get_doc_cached(hit.file_path)
                page = doc.load_page(hit.page - 1)
                txt = page.get_text("text") or ""
                self.page_text_box.setPlainText(txt.strip())
            except Exception:
                self.page_text_box.setPlainText(hit.context or hit.snippet or "")
        else:
            self.page_text_box.setPlainText(hit.context or hit.snippet or "")

        self._rerender()

    def _rerender(self):
        hit = self.current_hit
        if not hit:
            return

        if hit.file_type != "PDF" or hit.page is None or hit.rect is None:
            self.preview_label.setPixmap(QPixmap())
            self.preview_label.setText("(No PDF image preview for this hit.)")
            self.preview_label.setStyleSheet("background: #111; color: #ddd; padding: 10px;")
            self.preview_label.adjustSize()
            return

        try:
            base = self._get_page_image_cached(hit.file_path, hit.page, self.zoom)
            img = base.copy()

            draw = ImageDraw.Draw(img)
            x0, y0, x1, y1 = hit.rect
            rr = fitz.Rect(x0 * self.zoom, y0 * self.zoom, x1 * self.zoom, y1 * self.zoom)
            draw.rectangle([rr.x0, rr.y0, rr.x1, rr.y1], outline="red", width=max(2, int(3 * self.zoom)))

            qimg = ImageQt.ImageQt(img)
            pm = QPixmap.fromImage(qimg)
            self.preview_label.setStyleSheet("background: #111;")
            self.preview_label.setText("")
            self.preview_label.setPixmap(pm)
            self.preview_label.setFixedSize(pm.size())
        except Exception as e:
            self.preview_label.setPixmap(QPixmap())
            self.preview_label.setText(f"Failed to render PDF preview:\n{e}")
            self.preview_label.adjustSize()

    def get_selected_text(self) -> str:
        for w in (self.page_text_box, self.context_box):
            try:
                cur = w.textCursor()
                if cur is not None and cur.hasSelection():
                    return (cur.selectedText() or "").replace("\u2029", "\n").strip()
            except Exception:
                pass
        return ""




class ClipsDialog(QDialog):
    """Clips viewer window (opened from Clips menu)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Clips")
        self.setMinimumSize(860, 560)
        self.setWindowFlag(Qt.Window)

        self.clips: List[Clip] = []

        root = QWidget()
        self.setLayout(QVBoxLayout())
        self.layout().setContentsMargins(10, 10, 10, 10)
        self.layout().addWidget(root)

        outer = QVBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(10)

        self.clip_table = QTableWidget(0, 5)
        self.clip_table.setHorizontalHeaderLabels(["Created", "File", "Page/Line", "Keyword", "Preview"])
        self.clip_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.clip_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.clip_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.clip_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.clip_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.clip_table.verticalHeader().setVisible(False)
        self.clip_table.setAlternatingRowColors(True)
        self.clip_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.clip_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.clip_table.setSelectionMode(QAbstractItemView.SingleSelection)
        outer.addWidget(self.clip_table, 1)

        self.clip_detail = QTextEdit()
        self.clip_detail.setReadOnly(True)
        self.clip_detail.setMinimumHeight(150)
        outer.addWidget(self.clip_detail)

        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.rejected.connect(self.close)
        bb.accepted.connect(self.close)
        outer.addWidget(bb)

        self.clip_table.cellClicked.connect(self._on_row_clicked)

    def refresh(self, clips: List[Clip]):
        self.clips = list(clips)
        self.clip_table.setRowCount(len(self.clips))
        for i, c in enumerate(self.clips):
            page_or_line = ""
            if c.file_type == "PDF" and c.page is not None:
                page_or_line = f"p.{c.page}"
            elif c.line_no is not None:
                page_or_line = f"ln.{c.line_no}"

            preview = (c.selected_text or "").replace("\n", " ").strip()
            if len(preview) > 160:
                preview = preview[:160] + "…"

            self.clip_table.setItem(i, 0, QTableWidgetItem(c.created_at))
            self.clip_table.setItem(i, 1, QTableWidgetItem(Path(c.file_path).name))
            self.clip_table.setItem(i, 2, QTableWidgetItem(page_or_line))
            self.clip_table.setItem(i, 3, QTableWidgetItem(c.keyword))
            self.clip_table.setItem(i, 4, QTableWidgetItem(preview))

        if not self.clips:
            self.clip_detail.setPlainText("(No clips yet. Open Preview, select text in Context/Page text, then click 'Save selection as clip'.)")

    def select_last(self):
        if not self.clips:
            return
        last = len(self.clips) - 1
        self.clip_table.selectRow(last)
        self.clip_table.scrollToItem(self.clip_table.item(last, 0), QAbstractItemView.PositionAtCenter)
        self._on_row_clicked(last, 0)

    def _on_row_clicked(self, row: int, col: int):
        if row < 0 or row >= len(self.clips):
            return
        c = self.clips[row]
        where = ""
        if c.file_type == "PDF" and c.page is not None:
            where = f"Page {c.page}"
        elif c.line_no is not None:
            where = f"Line {c.line_no}"
        self.clip_detail.setPlainText(
            f"{Path(c.file_path).name} — {where}\n"
            f"Keyword: {c.keyword} (detected: {c.detected_word})\n"
            f"Saved: {c.created_at}\n\n"
            f"{c.selected_text}"
        )


# -------------------- Main window --------------------

class HybridKeywordWorkbench(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Keyword Finder")
        self.setMinimumSize(1100, 680)

        self.files: List[str] = []
        self.keywords_text: str = ""
        self.hits: List[Hit] = []
        self.clips: List[Clip] = []
        self.worker: Optional[SearchWorker] = None

        self.case_sensitive: bool = True
        self.whole_word: bool = True

        self.preview = PreviewPopup(self)
        self.preview.request_save_clip.connect(self.save_current_selection_as_clip)
        self.preview.set_main_window(self)
        self.clips_dialog = ClipsDialog(self)

        self._build_ui()
        self._build_menus()

        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.timeout.connect(self._rerender_preview)

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Add toolbar for quick actions
        toolbar_layout = QHBoxLayout()
        
        self.btn_manage_files = QPushButton("📁 Files")
        self.btn_manage_files.setToolTip("Manage files (Ctrl+O)")
        self.btn_manage_files.clicked.connect(self.manage_files)
        
        self.btn_manage_keywords = QPushButton("🔍 Input keywords")
        self.btn_manage_keywords.setToolTip("Manage keywords (Ctrl+K)")
        self.btn_manage_keywords.clicked.connect(self.manage_keywords)
               
        self.btn_show_preview = QPushButton("👁️ Preview")
        self.btn_show_preview.setToolTip("Show preview (Ctrl+P)")
        self.btn_show_preview.clicked.connect(self.toggle_preview)
        
        toolbar_layout.addWidget(self.btn_manage_files)
        toolbar_layout.addWidget(self.btn_manage_keywords)
        toolbar_layout.addWidget(self.btn_show_preview)
        toolbar_layout.addStretch()
        
        layout.addLayout(toolbar_layout)

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

        layout.addWidget(results_group, 1)

        self.status = QLabel("Ready.")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)

    def _build_menus(self):
        mb = self.menuBar()
        mb.clear()

        # Export
        m_files = mb.addMenu("Export")
        self.act_export_csv = QAction("Results as CSV…", self)
        self.act_export_csv.setShortcut("Ctrl+Shift+C")
        self.act_export_csv.triggered.connect(self.export_csv)

        self.act_export_xlsx = QAction("Results as XLSX…", self)
        self.act_export_xlsx.setShortcut("Ctrl+Shift+E")
        self.act_export_xlsx.triggered.connect(self.export_xlsx)

        self.act_export_clips_md = QAction("Clips as Markdown…", self)
        self.act_export_clips_md.setShortcut("Ctrl+Shift+M")
        self.act_export_clips_md.triggered.connect(self.export_clips_markdown)

        self.act_export_clips_csv = QAction("Clips as CSV…", self)
        self.act_export_clips_csv.setShortcut("Ctrl+Shift+L")
        self.act_export_clips_csv.triggered.connect(self.export_clips_csv)

        m_files.addAction(self.act_export_csv)
        m_files.addAction(self.act_export_xlsx)
        m_files.addAction(self.act_export_clips_md)
        m_files.addAction(self.act_export_clips_csv)

        # Clips
        m_clips = mb.addMenu("Clips")
        act_show_clips = QAction("Show clips…", self)
        act_show_clips.setShortcut("Ctrl+Alt+K")
        act_show_clips.triggered.connect(self.show_clips_dialog)
        m_clips.addAction(act_show_clips)
        m_clips.addSeparator()

        act_clear_clips = QAction("Clear clips", self)
        act_clear_clips.triggered.connect(self.clear_clips)
        m_clips.addAction(act_clear_clips)

        # initial enabling
        self._set_run_enabled(True)
        self._refresh_status_line()

    # ----- menu helpers -----
    def _refresh_status_line(self):
        cs = "ON" if self.case_sensitive else "OFF"
        ww = "ON" if self.whole_word else "OFF"
        self.status.setText(
            f"Files: {len(self.files)} | Keywords: {len(self._parse_keywords())} | Options: Case-sensitive {cs} | Whole-word {ww}"
        )

    def _set_run_enabled(self, enabled: bool):
        try:
            self.act_run_search.setEnabled(enabled)
            self.btn_run_search.setEnabled(enabled)
        except Exception:
            pass
        try:
            self.act_export_csv.setEnabled(enabled)
            self.act_export_xlsx.setEnabled(enabled)
        except Exception:
            pass

    # ----- Files / Keywords dialogs -----
    def manage_files(self):
        dlg = FileManagerDialog(self, self.files)
        if dlg.exec() == QDialog.Accepted:
            self.files = dlg.get_files()
            self._refresh_status_line()

    def clear_files(self):
        self.files = []
        self.hits = []
        self.table.setRowCount(0)
        self._refresh_status_line()
        self.status.setText("Files cleared.")

    def manage_keywords(self):
        dlg = KeywordManagerDialog(self, self.keywords_text, self.case_sensitive, self.whole_word)
        dlg.run_search_requested.connect(self._on_keywords_run_search)
        if dlg.exec() == QDialog.Accepted:
            self.keywords_text = dlg.get_text()
            self.case_sensitive = dlg.get_case_sensitive()
            self.whole_word = dlg.get_whole_word()
            self._refresh_status_line()

    def _on_keywords_run_search(self):
        # Save keywords and run search immediately
        dlg = self.sender()
        self.keywords_text = dlg.get_text()
        self.case_sensitive = dlg.get_case_sensitive()
        self.whole_word = dlg.get_whole_word()
        self._refresh_status_line()
        self.run_search()

    def clear_keywords(self):
        self.keywords_text = ""
        self._refresh_status_line()
        self.status.setText("Keywords cleared.")

    # ----- Search -----
    def _parse_keywords(self) -> List[str]:
        raw = (self.keywords_text or "").strip()
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
            key = k if self.case_sensitive else k.lower()
            if key not in seen:
                seen.add(key)
                out.append(k)
        return out

    def run_search(self):
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "Running", "Search is already running.")
            return

        if not self.files:
            QMessageBox.information(self, "Tip", "Add files first: Files → Manage files…")
            return

        keywords = self._parse_keywords()
        if not keywords:
            QMessageBox.information(self, "Tip", "Add keywords first: Keywords → Manage keywords…")
            return

        self.table.setRowCount(0)
        self.hits = []
        self.status.setText("Starting…")
        self._set_run_enabled(False)

        self.worker = SearchWorker(self.files, keywords, case_sensitive=self.case_sensitive, whole_word=self.whole_word)
        self.worker.progress.connect(self.status.setText)
        self.worker.finished_ok.connect(self._on_finished_ok)
        self.worker.finished_err.connect(self._on_finished_err)
        self.worker.start()

    def _on_finished_ok(self, hits: List[Hit]):
        self.hits = hits
        self._populate_table(hits)
        self._set_run_enabled(True)

        if hits:
            self.table.selectRow(0)
            self.on_row_clicked(0, 0)

    def _on_finished_err(self, msg: str):
        self._set_run_enabled(True)
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
        self.setWindowTitle(f"Keyword Finder — {len(hits)} hit(s)")

    # ----- Preview -----
    def toggle_preview(self):
        if self.preview.isVisible():
            self.preview.hide()
        else:
            self.preview.show()
            self._rerender_preview()

    def _rerender_preview(self):
        row = self.table.currentRow()
        if 0 <= row < len(self.hits):
            self.preview.show_hit(self.hits[row])

    def on_row_clicked(self, row: int, col: int):
        if row < 0 or row >= len(self.hits):
            return
        if self.preview.isVisible():
            self.preview.show_hit(self.hits[row])

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._resize_timer.start(120)

    # ----- Hit navigation -----
    def select_next_hit(self):
        if not self.hits:
            return
        row = self.table.currentRow()
        next_row = 0 if row < 0 else min(len(self.hits) - 1, row + 1)
        self.table.selectRow(next_row)
        self.table.scrollToItem(self.table.item(next_row, 0), QAbstractItemView.PositionAtCenter)
        if self.preview.isVisible():
            self.preview.show_hit(self.hits[next_row])

    def select_prev_hit(self):
        if not self.hits:
            return
        row = self.table.currentRow()
        prev_row = 0 if row < 0 else max(0, row - 1)
        self.table.selectRow(prev_row)
        self.table.scrollToItem(self.table.item(prev_row, 0), QAbstractItemView.PositionAtCenter)
        if self.preview.isVisible():
            self.preview.show_hit(self.hits[prev_row])

    def keyPressEvent(self, event):
        # Avoid hijacking typing when focus is inside text boxes (preview)
        fw = QApplication.focusWidget()
        if isinstance(fw, QTextEdit):
            return super().keyPressEvent(event)

        k = event.key()
        if k == Qt.Key_J:
            self.select_next_hit()
            return
        if k == Qt.Key_K:
            self.select_prev_hit()
            return
        return super().keyPressEvent(event)

    # ----- Continue reading -----
    def open_selected_in_viewer(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self.hits):
            QMessageBox.information(self, "Tip", "Select a result row first.")
            return
        self.open_hit_in_viewer(row)

    def open_hit_in_viewer(self, row: int):
        if row < 0 or row >= len(self.hits):
            return
        hit = self.hits[row]
        if hit.file_type != "PDF" or not hit.page:
            QMessageBox.information(self, "Tip", "Continue reading is available for PDF hits only.")
            return

        pdf_path = str(Path(hit.file_path).resolve())
        url = QUrl.fromLocalFile(pdf_path)
        url.setFragment(f"page={hit.page}")

        if not QDesktopServices.openUrl(url):
            if not QDesktopServices.openUrl(QUrl.fromLocalFile(pdf_path)):
                QMessageBox.warning(self, "Open failed", "Could not open the PDF in a viewer.\n\n" + pdf_path)

    # ----- Clips -----
    def show_clips_dialog(self):
        self.clips_dialog.refresh(self.clips)
        self.clips_dialog.show()
        self.clips_dialog.raise_()
        self.clips_dialog.activateWindow()

    def save_current_selection_as_clip(self):
        if not self.hits:
            QMessageBox.information(self, "Tip", "Run a search and select a result first.")
            return
        row = self.table.currentRow()
        if row < 0 or row >= len(self.hits):
            QMessageBox.information(self, "Tip", "Select a result row first.")
            return
        if not self.preview.isVisible():
            QMessageBox.information(self, "Tip", "Open the preview first: Preview → Show preview")
            return

        selected = self.preview.get_selected_text()
        if not selected:
            QMessageBox.information(self, "Tip", "Select some text in the Preview window first (Page text / Context).")
            return

        hit = self.hits[row]
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        clip = Clip(
            created_at=stamp,
            file_path=hit.file_path,
            file_type=hit.file_type,
            page=hit.page,
            line_no=hit.line_no,
            keyword=hit.keyword,
            detected_word=hit.detected_word,
            selected_text=selected,
        )
        self.clips.append(clip)
        # refresh clips viewer (if open)
        try:
            self.clips_dialog.refresh(self.clips)
        except Exception:
            pass

        self.status.setText(f"Saved clip. Total: {len(self.clips)}")
        # auto-select newest in Clips window
        try:
            self.clips_dialog.select_last()
        except Exception:
            pass

    def clear_clips(self):
        self.clips = []
        try:
            self.clips_dialog.refresh(self.clips)
        except Exception:
            pass
        self.status.setText("Clips cleared.")

    def export_clips_markdown(self):
        if not self.clips:
            QMessageBox.information(self, "Tip", "No clips to export.")
            return
        default = f"quote_clips_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
        out, _ = QFileDialog.getSaveFileName(self, "Save Markdown", default, "Markdown (*.md)")
        if not out:
            return
        try:
            lines = ["# Quote Clips\n", f"Exported: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n", "---\n"]
            for c in self.clips:
                where = ""
                if c.file_type == "PDF" and c.page is not None:
                    where = f"Page {c.page}"
                elif c.line_no is not None:
                    where = f"Line {c.line_no}"
                lines.append(f"## {Path(c.file_path).name} — {where}\n")
                lines.append(f"- Saved: {c.created_at}")
                lines.append(f"- Keyword: {c.keyword}")
                if c.detected_word:
                    lines.append(f"- Detected: {c.detected_word}")
                lines.append(f"- Source path: {c.file_path}\n")
                for ln in (c.selected_text or "").splitlines():
                    lines.append(f"> {ln}")
                lines.append("\n")
            Path(out).write_text("\n".join(lines), encoding="utf-8")
            QMessageBox.information(self, "Export", f"Saved:\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "Export failed", str(e))

    def export_clips_csv(self):
        if not self.clips:
            QMessageBox.information(self, "Tip", "No clips to export.")
            return
        default = f"quote_clips_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        out, _ = QFileDialog.getSaveFileName(self, "Save CSV", default, "CSV (*.csv)")
        if not out:
            return
        try:
            with open(out, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["created_at", "file", "file_type", "page", "line_no", "keyword", "detected_word", "selected_text"])
                for c in self.clips:
                    w.writerow([
                        c.created_at,
                        c.file_path,
                        c.file_type,
                        "" if c.page is None else c.page,
                        "" if c.line_no is None else c.line_no,
                        c.keyword,
                        c.detected_word,
                        c.selected_text,
                    ])
            QMessageBox.information(self, "Export", f"Saved:\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "Export failed", str(e))

    # ----- Export results -----
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
    w.resize(1400, 820)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
