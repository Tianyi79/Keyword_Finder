"""
Integrated Keyword Finder (PDF) - single-file app

v2.2 update (Aligned snippet)
- Adds "Detected" column (actual word at hit location).
- Snippet is now extracted from the SAME LINE as the detected word (word-level bbox -> line grouping),
  so it won't always show the first occurrence on the page.
- Right-side context now shows previous/current/next line around the hit line.

Install
    python -m pip install PySide6 pymupdf pillow

Run
    python integrated_keyword_finder_v2_2_aligned_snippet.py
"""

from __future__ import annotations

import sys
import json
import csv
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional
from pathlib import Path
from datetime import datetime

import fitz  # PyMuPDF
from PIL import Image, ImageQt, ImageDraw

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog,
    QHBoxLayout, QVBoxLayout, QPushButton, QLineEdit, QLabel,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView,
    QListWidget, QListWidgetItem, QTextEdit, QSplitter,
    QAbstractItemView, QScrollArea
)


@dataclass
class Hit:
    file_path: str
    keyword: str
    detected_word: str
    page: int            # 1-based
    rect: fitz.Rect      # PDF coordinates
    block: int           # from get_text("words")
    line: int            # from get_text("words")
    snippet: str         # the full line text containing the detected word


class IntegratedKeywordFinder(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Integrated Keyword Finder (PDF)")

        self.pdfs: List[str] = []
        self.docs: Dict[str, fitz.Document] = {}
        self.hits: List[Hit] = []
        self.zoom: float = 2.0

        # drag-to-pan
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
        layout = QHBoxLayout(root)
        layout.addWidget(splitter)

        # Left panel
        left = QWidget()
        left_layout = QVBoxLayout(left)
        splitter.addWidget(left)

        pdf_row = QHBoxLayout()
        self.btn_add = QPushButton("Add PDFs")
        self.btn_add.clicked.connect(self.add_pdfs)
        self.btn_remove = QPushButton("Remove Selected")
        self.btn_remove.clicked.connect(self.remove_selected_pdfs)
        self.btn_clear = QPushButton("Clear")
        self.btn_clear.clicked.connect(self.clear_pdfs)
        pdf_row.addWidget(self.btn_add)
        pdf_row.addWidget(self.btn_remove)
        pdf_row.addWidget(self.btn_clear)
        left_layout.addLayout(pdf_row)

        self.pdf_list = QListWidget()
        self.pdf_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        left_layout.addWidget(self.pdf_list, 1)

        kw_row = QHBoxLayout()
        self.kw_input = QLineEdit()
        self.kw_input.setPlaceholderText("Quick keyword (comma-separated ok)")
        kw_row.addWidget(QLabel("Keyword:"))
        kw_row.addWidget(self.kw_input, 1)
        left_layout.addLayout(kw_row)

        self.kw_box = QTextEdit()
        self.kw_box.setPlaceholderText("More keywords (one per line). Example:\nhistory\neligibility\ncertificate\n")
        self.kw_box.setFixedHeight(110)
        left_layout.addWidget(self.kw_box)

        action_row = QHBoxLayout()
        self.btn_search = QPushButton("Search")
        self.btn_search.clicked.connect(self.search_all)
        self.btn_export_csv = QPushButton("Export CSV")
        self.btn_export_csv.clicked.connect(self.export_csv)
        self.btn_export_json = QPushButton("Export JSON")
        self.btn_export_json.clicked.connect(self.export_json)
        action_row.addWidget(self.btn_search)
        action_row.addWidget(self.btn_export_csv)
        action_row.addWidget(self.btn_export_json)
        left_layout.addLayout(action_row)

        # Results table (Detected + aligned snippet)
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["#", "File", "Keyword", "Detected", "Page", "Snippet (hit line)"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.cellClicked.connect(self.on_row_clicked)
        left_layout.addWidget(self.table, 2)

        # Right panel
        right = QWidget()
        right_layout = QVBoxLayout(right)
        splitter.addWidget(right)

        # Zoom controls
        zoom_row = QHBoxLayout()
        self.btn_zoom_out = QPushButton("−")
        self.btn_zoom_in = QPushButton("+")
        self.btn_zoom_reset = QPushButton("100%")
        zoom_row.addWidget(QLabel("Zoom"))
        zoom_row.addWidget(self.btn_zoom_out)
        zoom_row.addWidget(self.btn_zoom_in)
        zoom_row.addWidget(self.btn_zoom_reset)
        zoom_row.addStretch(1)
        right_layout.addLayout(zoom_row)

        self.btn_zoom_in.clicked.connect(lambda: self._set_zoom(self.zoom * 1.25))
        self.btn_zoom_out.clicked.connect(lambda: self._set_zoom(self.zoom / 1.25))
        self.btn_zoom_reset.clicked.connect(lambda: self._set_zoom(2.0))

        # Scrollable preview
        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.preview_label.setStyleSheet("background: #111;")
        self.preview_label.setMinimumSize(600, 600)

        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidgetResizable(False)
        self.preview_scroll.setWidget(self.preview_label)
        right_layout.addWidget(self.preview_scroll, 3)

        self.context_box = QTextEdit()
        self.context_box.setReadOnly(True)
        self.context_box.setPlaceholderText("Context around the hit line will appear here.")
        self.context_box.setFixedHeight(180)
        right_layout.addWidget(self.context_box, 1)

        splitter.setSizes([780, 520])

    # ---------- Zoom / Pan ----------
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

    # ---------- PDF list ----------
    def add_pdfs(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select PDFs", "", "PDF Files (*.pdf)")
        if not paths:
            return
        added = 0
        for p in paths:
            p = str(Path(p))
            if p not in self.pdfs:
                self.pdfs.append(p)
                self.pdf_list.addItem(QListWidgetItem(p))
                added += 1
        if added:
            self._ensure_docs_loaded()

    def remove_selected_pdfs(self):
        items = self.pdf_list.selectedItems()
        if not items:
            return
        to_remove = {it.text() for it in items}
        self.pdfs = [p for p in self.pdfs if p not in to_remove]
        self._close_docs(to_remove)
        self._refresh_pdf_list()
        self._clear_results_if_files_removed(to_remove)

    def clear_pdfs(self):
        self._close_docs(set(self.pdfs))
        self.pdfs = []
        self._refresh_pdf_list()
        self.hits = []
        self.table.setRowCount(0)
        self.preview_label.clear()
        self.context_box.clear()

    def _refresh_pdf_list(self):
        self.pdf_list.clear()
        for p in self.pdfs:
            self.pdf_list.addItem(QListWidgetItem(p))

    def _ensure_docs_loaded(self):
        for p in self.pdfs:
            if p in self.docs:
                continue
            try:
                self.docs[p] = fitz.open(p)
            except Exception as e:
                QMessageBox.warning(self, "Open PDF failed", f"Failed to open:\n{p}\n\n{e}")

    def _close_docs(self, paths: set[str]):
        for p in list(paths):
            doc = self.docs.pop(p, None)
            try:
                if doc is not None:
                    doc.close()
            except Exception:
                pass

    def _clear_results_if_files_removed(self, removed: set[str]):
        if not self.hits:
            return
        self.hits = [h for h in self.hits if h.file_path not in removed]
        self._rebuild_table()

    # ---------- Search ----------
    def _get_keywords(self) -> List[str]:
        keywords: List[str] = []
        quick = self.kw_input.text().strip()
        if quick:
            for part in quick.split(","):
                k = part.strip()
                if k:
                    keywords.append(k)

        box = self.kw_box.toPlainText()
        if box.strip():
            for line in box.splitlines():
                for part in line.split(","):
                    k = part.strip()
                    if k:
                        keywords.append(k)

        seen = set()
        uniq = []
        for k in keywords:
            kl = k.lower()
            if kl not in seen:
                seen.add(kl)
                uniq.append(k)
        return uniq

    def search_all(self):
        if not self.pdfs:
            QMessageBox.information(self, "Tip", "Add at least one PDF.")
            return

        keywords = self._get_keywords()
        if not keywords:
            QMessageBox.information(self, "Tip", "Enter at least one keyword.")
            return

        self._ensure_docs_loaded()

        self.hits = []
        self.table.setRowCount(0)
        self.preview_label.clear()
        self.context_box.clear()
        QApplication.processEvents()

        for pdf_path in self.pdfs:
            doc = self.docs.get(pdf_path)
            if not doc:
                continue

            for pno in range(len(doc)):
                page = doc.load_page(pno)

                # word-level data (bbox + block/line ids)
                words = page.get_text("words") or []  # [x0,y0,x1,y1,word,block,line,wordno]
                line_text, ordered_lines = self._build_line_index(words)

                for kw in keywords:
                    rects = page.search_for(kw)
                    if not rects:
                        continue

                    for r in rects:
                        best = self._best_word_entry_from_rect(words, r)
                        if best is None:
                            detected, blk, ln = kw, -1, -1
                            snippet = ""
                        else:
                            detected, blk, ln = best
                            snippet = line_text.get((blk, ln), "")

                        self.hits.append(Hit(
                            file_path=pdf_path,
                            keyword=kw,
                            detected_word=detected,
                            page=pno + 1,
                            rect=r,
                            block=blk,
                            line=ln,
                            snippet=snippet
                        ))

        self._rebuild_table()

        if self.hits:
            self.table.selectRow(0)
            self.on_row_clicked(0, 0)
        else:
            self.context_box.setPlainText("No hits.\n\nTip: If the PDF is scanned (no text layer), text search will not work.")

    def _rebuild_table(self):
        self.table.setRowCount(len(self.hits))
        for i, h in enumerate(self.hits):
            self.table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.table.setItem(i, 1, QTableWidgetItem(Path(h.file_path).name))
            self.table.setItem(i, 2, QTableWidgetItem(h.keyword))
            self.table.setItem(i, 3, QTableWidgetItem(h.detected_word))
            self.table.setItem(i, 4, QTableWidgetItem(str(h.page)))
            self.table.setItem(i, 5, QTableWidgetItem(h.snippet))
        self.setWindowTitle(f"Integrated Keyword Finder (PDF) — {len(self.hits)} hit(s)")

    # ---------- Preview ----------
    def on_row_clicked(self, row: int, col: int):
        if row < 0 or row >= len(self.hits):
            return
        self.render_hit(self.hits[row])

    def render_hit(self, hit: Hit):
        self._ensure_docs_loaded()
        doc = self.docs.get(hit.file_path)
        if not doc:
            QMessageBox.warning(self, "Error", f"PDF not loaded:\n{hit.file_path}")
            return

        try:
            page = doc.load_page(hit.page - 1)

            # Render page image
            mat = fitz.Matrix(self.zoom, self.zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

            # Highlight ONLY the selected hit rect
            draw = ImageDraw.Draw(img)
            r = hit.rect
            rr = fitz.Rect(r.x0 * self.zoom, r.y0 * self.zoom, r.x1 * self.zoom, r.y1 * self.zoom)
            draw.rectangle([rr.x0, rr.y0, rr.x1, rr.y1], outline="red", width=max(2, int(3 * self.zoom)))

            qimg = ImageQt.ImageQt(img)
            pm = QPixmap.fromImage(qimg)
            self.preview_label.setPixmap(pm)
            self.preview_label.setFixedSize(pm.size())

            # Context: prev/hit/next lines around the hit's (block,line)
            words = page.get_text("words") or []
            line_text, ordered_lines = self._build_line_index(words)
            ctx = self._context_from_line_key(line_text, ordered_lines, (hit.block, hit.line))
            if not ctx:
                # fallback
                page_text = page.get_text("text") or ""
                ctx = page_text[:3000] if page_text else "No text layer detected on this page."
            self.context_box.setPlainText(ctx)

        except Exception as e:
            QMessageBox.critical(self, "Render error", f"Failed to render page:\n{e}")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._resize_timer.start(120)

    def _rerender_current_selection(self):
        row = self.table.currentRow()
        if 0 <= row < len(self.hits):
            self.render_hit(self.hits[row])

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
                w.writerow(["file", "keyword", "detected_word", "page", "block", "line", "x0", "y0", "x1", "y1", "snippet"])
                for h in self.hits:
                    r = h.rect
                    w.writerow([h.file_path, h.keyword, h.detected_word, h.page, h.block, h.line, r.x0, r.y0, r.x1, r.y1, h.snippet])
            QMessageBox.information(self, "Export", f"Saved:\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "Export failed", str(e))

    def export_json(self):
        if not self.hits:
            QMessageBox.information(self, "Tip", "No results to export.")
            return
        default = f"keyword_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        out, _ = QFileDialog.getSaveFileName(self, "Save JSON", default, "JSON (*.json)")
        if not out:
            return
        try:
            data = []
            for h in self.hits:
                r = h.rect
                data.append({
                    "file": h.file_path,
                    "keyword": h.keyword,
                    "detected_word": h.detected_word,
                    "page": h.page,
                    "block": h.block,
                    "line": h.line,
                    "rect": {"x0": r.x0, "y0": r.y0, "x1": r.x1, "y1": r.y1},
                    "snippet": h.snippet
                })
            with open(out, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "Export", f"Saved:\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "Export failed", str(e))

    # ---------- Helpers ----------
    @staticmethod
    def _build_line_index(words) -> Tuple[Dict[Tuple[int, int], str], List[Tuple[int, int]]]:
        """
        Build:
          - line_text: (block,line) -> full line text (words joined by wordno)
          - ordered_lines: list of (block,line) sorted by (top y, then x) for context
        """
        if not words:
            return {}, []

        # (block,line) -> list of (wordno, x0, y0, word)
        tmp: Dict[Tuple[int, int], List[Tuple[int, float, float, str]]] = {}
        # for ordering
        line_pos: Dict[Tuple[int, int], Tuple[float, float]] = {}

        for w in words:
            if len(w) < 8:
                continue
            x0, y0, x1, y1, txt, blk, ln, wordno = w[0], w[1], w[2], w[3], w[4], int(w[5]), int(w[6]), int(w[7])
            key = (blk, ln)
            tmp.setdefault(key, []).append((wordno, float(x0), float(y0), str(txt)))
            # use min y then min x as representative position
            if key not in line_pos:
                line_pos[key] = (float(y0), float(x0))
            else:
                cy, cx = line_pos[key]
                line_pos[key] = (min(cy, float(y0)), min(cx, float(x0)))

        line_text: Dict[Tuple[int, int], str] = {}
        for key, items in tmp.items():
            items.sort(key=lambda t: (t[0], t[1]))
            # join with spaces; this is best-effort (PDFs sometimes need custom spacing)
            line_text[key] = " ".join([t[3] for t in items]).strip()

        ordered_lines = sorted(line_text.keys(), key=lambda k: (line_pos.get(k, (1e9, 1e9))[0], line_pos.get(k, (1e9, 1e9))[1], k[0], k[1]))
        return line_text, ordered_lines

    @staticmethod
    def _best_word_entry_from_rect(words, rect: fitz.Rect) -> Optional[Tuple[str, int, int]]:
        """
        Return (word, block, line) whose word-rect overlaps the hit-rect the most (by intersection area).
        """
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

    @staticmethod
    def _context_from_line_key(line_text: Dict[Tuple[int, int], str],
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


def main():
    app = QApplication(sys.argv)
    w = IntegratedKeywordFinder()
    w.resize(1320, 780)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
