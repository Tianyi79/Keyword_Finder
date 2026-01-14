import re
from kreuzberg import (
    ExtractionConfig, PdfConfig, HierarchyConfig, PageConfig,
    batch_extract_files_sync
)

# ✅ Put markers on their own lines
MARKER_FORMAT = "\n--- Page {page_num} ---\n"  # {page_num} is the supported placeholder :contentReference[oaicite:1]{index=1}
MARKER_RE = re.compile(r"--- Page\s+(\d+)\s+---")

config = ExtractionConfig(
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

文件列表 = ["file1.pdf", "file2.pdf"]
关键词列表 = ["history", "签署"]  # history: case-insensitive; Chinese: exact

results = batch_extract_files_sync(文件列表, config=config)

def hit(keyword: str, line: str) -> bool:
    if keyword.isascii():
        return keyword.lower() in line.lower()
    return keyword in line

for i, result in enumerate(results):
    print("=" * 80)
    print(f"文件: {文件列表[i]}")
    print("=" * 80)

    content = result.content or ""

    # Split content into: [text_before_first_marker, page_num, page_text, page_num, page_text, ...]
    parts = MARKER_RE.split(content)

    # parts[0] is anything before the first marker (often empty)
    found = {kw: [] for kw in 关键词列表}

    # Iterate page chunks
    # page numbers will be in parts[1], parts[3], ... and page text in parts[2], parts[4], ...
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
            for kw in 关键词列表:
                if hit(kw, s):
                    found[kw].append({
                        "页码": page_num,
                        "页内行号": line_in_page,
                        "内容": s
                    })

    hit_keywords = [kw for kw in 关键词列表 if found[kw]]

    if not hit_keywords:
        print("✗ 未找到任何关键词\n")
        continue

    print(f"✓ 找到 {len(hit_keywords)} 个关键词\n")
    for kw in hit_keywords:
        print(f"关键词 '{kw}' 出现 {len(found[kw])} 次:")
        for item in found[kw]:
            print(f"  第 {item['页码']} 页，第 {item['页内行号']} 行: {item['内容']}")
        print()
