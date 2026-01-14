# 批量搜索多个文件中的关键字

from kreuzberg import batch_extract_files_sync, ExtractionConfig

文件列表 = ["file.pdf", "file2.pdf"]
关键词列表 = ["合同", "协议", "签署"]

results = batch_extract_files_sync(文件列表, config=ExtractionConfig())

# 统计每个文件包含的关键词
for i, result in enumerate(results):
    print(f"{'='*80}")
    print(f"文件: {文件列表[i]}")
    print(f"{'='*80}")

    # 按行分割文本
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

    # 检查是否找到关键词
    找到的词 = [词 for 词 in 关键词列表 if len(找到的关键词[词]) > 0]

    if 找到的词:
        print(f"✓ 找到 {len(找到的词)} 个关键词")

        # 打印每个关键词所在的行
        for 词 in 找到的词:
            print(f"关键词 '{词}' 出现 {len(找到的关键词[词])} 次:")
            for 项 in 找到的关键词[词]:
                print(f"  第 {项['行号']} 行: {项['内容']}")
            print()
    else:
        print(f"✗ 未找到任何关键词")

    print()
