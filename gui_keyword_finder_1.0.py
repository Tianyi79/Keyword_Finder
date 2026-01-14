import re
import threading
import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from kreuzberg import (
    ExtractionConfig, PdfConfig, HierarchyConfig, PageConfig,
    batch_extract_files_sync
)

# ----------------------------
# Core logic: extract + map to (page, line_in_page) via page markers
# ----------------------------

MARKER_FORMAT = "\n--- Page {page_num} ---\n"
MARKER_RE = re.compile(r"--- Page\s+(\d+)\s+---")

def build_config():
    return ExtractionConfig(
        pages=PageConfig(
            extract_pages=True,
            insert_page_markers=True,
            marker_format=MARKER_FORMAT
        ),
        pdf_options=PdfConfig(
            hierarchy=HierarchyConfig(
                enabled=True,
                k_clusters=6,
                include_bbox=True,
                ocr_coverage_threshold=None
            )
        )
    )

def parse_keywords(raw: str):
    # Accept comma-separated OR newline-separated
    items = []
    for chunk in raw.replace(",", "\n").splitlines():
        s = chunk.strip()
        if s:
            items.append(s)
    # de-dup while preserving order
    seen = set()
    out = []
    for k in items:
        if k not in seen:
            seen.add(k)
            out.append(k)
    return out

def hit(keyword: str, line: str) -> bool:
    if keyword.isascii():
        return keyword.lower() in line.lower()
    return keyword in line

def find_keywords_in_result_content(result_content: str, keywords):
    """
    Returns list of dict rows:
      {page, line_in_page, keyword, text}
    """
    rows = []
    content = result_content or ""

    # Split content into: [before_first_marker, page_num, page_text, page_num, page_text, ...]
    parts = MARKER_RE.split(content)

    # iterate page chunks
    for idx in range(1, len(parts), 2):
        page_num_str = parts[idx]
        page_text = parts[idx + 1] if idx + 1 < len(parts) else ""
        if not page_num_str.isdigit():
            continue
        page_num = int(page_num_str)

        lines = page_text.splitlines()
        for line_in_page, line in enumerate(lines, 1):
            s = line.strip()
            if not s:
                continue
            for kw in keywords:
                if hit(kw, s):
                    rows.append({
                        "page": page_num,
                        "line_in_page": line_in_page,
                        "keyword": kw,
                        "text": s
                    })

    return rows


# ----------------------------
# GUI
# ----------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Kreuzberg PDF Keyword Finder (page + line)")
        self.geometry("1100x650")

        self.files = []
        self.results_rows = []

        self._build_ui()

    def _build_ui(self):
        # Top: file controls + keyword entry + run
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        file_box = ttk.LabelFrame(top, text="PDF files", padding=10)
        file_box.pack(side="left", fill="both", expand=True)

        self.file_list = tk.Listbox(file_box, height=8, selectmode=tk.EXTENDED)
        self.file_list.pack(side="left", fill="both", expand=True)

        file_btns = ttk.Frame(file_box)
        file_btns.pack(side="left", fill="y", padx=(10, 0))

        ttk.Button(file_btns, text="Add PDFs...", command=self.add_files).pack(fill="x", pady=2)
        ttk.Button(file_btns, text="Remove selected", command=self.remove_selected).pack(fill="x", pady=2)
        ttk.Button(file_btns, text="Clear", command=self.clear_files).pack(fill="x", pady=2)

        kw_box = ttk.LabelFrame(top, text="Keywords (comma or newline separated)", padding=10)
        kw_box.pack(side="left", fill="both", expand=True, padx=(10, 0))

        self.kw_text = tk.Text(kw_box, height=8, width=40)
        self.kw_text.pack(fill="both", expand=True)
        self.kw_text.insert("1.0", "history\n签署")

        run_box = ttk.Frame(top)
        run_box.pack(side="left", fill="y", padx=(10, 0))

        self.run_btn = ttk.Button(run_box, text="Run", command=self.run_search)
        self.run_btn.pack(fill="x", pady=2)

        self.export_btn = ttk.Button(run_box, text="Export CSV...", command=self.export_csv, state="disabled")
        self.export_btn.pack(fill="x", pady=2)

        self.status = ttk.Label(run_box, text="Ready")
        self.status.pack(fill="x", pady=(12, 2))

        # Bottom: results table
        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill="both", expand=True)

        cols = ("file", "keyword", "page", "line_in_page", "text")
        self.tree = ttk.Treeview(bottom, columns=cols, show="headings")
        self.tree.heading("file", text="File")
        self.tree.heading("keyword", text="Keyword")
        self.tree.heading("page", text="Page")
        self.tree.heading("line_in_page", text="Line (within page)")
        self.tree.heading("text", text="Matched text")

        self.tree.column("file", width=260, anchor="w")
        self.tree.column("keyword", width=120, anchor="w")
        self.tree.column("page", width=70, anchor="center")
        self.tree.column("line_in_page", width=120, anchor="center")
        self.tree.column("text", width=520, anchor="w")

        yscroll = ttk.Scrollbar(bottom, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)

        self.tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not paths:
            return
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.file_list.insert(tk.END, p)

    def remove_selected(self):
        sel = list(self.file_list.curselection())
        if not sel:
            return
        for idx in reversed(sel):
            path = self.file_list.get(idx)
            self.file_list.delete(idx)
            if path in self.files:
                self.files.remove(path)

    def clear_files(self):
        self.files.clear()
        self.file_list.delete(0, tk.END)

    def set_status(self, text):
        self.status.config(text=text)
        self.update_idletasks()

    def clear_results_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.results_rows = []
        self.export_btn.config(state="disabled")

    def run_search(self):
        if not self.files:
            messagebox.showwarning("No files", "Please add at least one PDF.")
            return

        keywords = parse_keywords(self.kw_text.get("1.0", tk.END))
        if not keywords:
            messagebox.showwarning("No keywords", "Please input at least one keyword.")
            return

        self.clear_results_table()
        self.run_btn.config(state="disabled")
        self.set_status("Running...")

        def worker():
            try:
                config = build_config()
                results = batch_extract_files_sync(self.files, config=config)

                all_rows = []
                for file_path, res in zip(self.files, results):
                    rows = find_keywords_in_result_content(res.content or "", keywords)
                    for r in rows:
                        all_rows.append({
                            "file": file_path,
                            "keyword": r["keyword"],
                            "page": r["page"],
                            "line_in_page": r["line_in_page"],
                            "text": r["text"]
                        })

                self.after(0, lambda: self._render_results(all_rows))

            except Exception as e:
                self.after(0, lambda: self._on_error(e))

        threading.Thread(target=worker, daemon=True).start()

    def _render_results(self, rows):
        self.results_rows = rows

        for r in rows:
            self.tree.insert(
                "",
                tk.END,
                values=(r["file"], r["keyword"], r["page"], r["line_in_page"], r["text"])
            )

        self.run_btn.config(state="normal")
        self.export_btn.config(state=("normal" if rows else "disabled"))

        self.set_status(f"Done. Hits: {len(rows)}")

        if not rows:
            messagebox.showinfo("No matches", "No keywords were found in the selected PDFs.")

    def _on_error(self, e: Exception):
        self.run_btn.config(state="normal")
        self.export_btn.config(state="disabled")
        self.set_status("Error")
        messagebox.showerror("Error", str(e))

    def export_csv(self):
        if not self.results_rows:
            return

        path = filedialog.asksaveasfilename(
            title="Save results as CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")]
        )
        if not path:
            return

        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=["file", "keyword", "page", "line_in_page", "text"])
                writer.writeheader()
                writer.writerows(self.results_rows)

            messagebox.showinfo("Saved", f"Exported {len(self.results_rows)} rows to:\n{path}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))


if __name__ == "__main__":
    app = App()
    app.mainloop()
