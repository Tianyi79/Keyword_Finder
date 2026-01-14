# 搜索同时包含所有关键词的内容
from kreuzberg import extract_file_sync, ExtractionConfig

result = extract_file_sync("file.pdf", config=ExtractionConfig())

关键词列表 = ["keyword_1", "keyword_2", "keyword_3"]
# 按行分割文本（使用 splitlines() 方法更可靠）
行列表 = result.content.splitlines()

# 搜索包含关键词的行
找到的关键词 = {}

for 词 in 关键词列表:
    找到的关键词[词] = []
    for 行号, 行内容 in enumerate(行列表, 1):
        if 词 in 行内容:
            找到的关键词[词].append({
                '行号': 行号,
                '内容': 行内容.strip()
            })

# 输出结果
包含所有关键词 = all(len(找到的关键词[词]) > 0 for 词 in 关键词列表)

if 包含所有关键词:
    print(f"✓ 文档包含所有关键词: {', '.join(关键词列表)}")
else:
    缺失的词 = [词 for 词 in 关键词列表 if len(找到的关键词[词]) == 0]
    print(f"✗ 缺失关键词: {', '.join(缺失的词)}")

# 打印每个关键词所在的行
for 词 in 关键词列表:
    if 找到的关键词[词]:
        print(f"关键词 '{词}' 出现 {len(找到的关键词[词])} 次:")
        for 项 in 找到的关键词[词]:
            print(f"  第 {项['行号']} 行: {项['内容']}")
        print()
    else:
        print(f"关键词 '{词}': 未找到")
