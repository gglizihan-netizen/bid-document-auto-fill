# -*- coding: utf-8 -*-
"""
分析模板中的表格结构
"""

import os
import sys
from pathlib import Path

if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document

TEMPLATE_PATH = r"D:\桌面\投标文件组成模版01.docx"

doc = Document(TEMPLATE_PATH)

print("=" * 70)
print("模板表格结构分析")
print("=" * 70)

for table_idx, table in enumerate(doc.tables):
    print(f"\n【表格 {table_idx + 1}】")
    print(f"行数: {len(table.rows)}, 列数: {len(table.rows[0].cells) if table.rows else 0}")
    print("-" * 70)

    # 显示表格内容
    for row_idx, row in enumerate(table.rows):
        cells_text = []
        for cell in row.cells:
            text = cell.text.strip()[:30]  # 截取前30字符
            if not text:
                text = "[空]"
            cells_text.append(text)

        # 判断是否是重复行（可能是动态列表）
        if row_idx > 0:
            prev_cells = []
            for cell in table.rows[row_idx - 1].cells:
                prev_text = cell.text.strip()[:30]
                if not prev_text:
                    prev_text = "[空]"
                prev_cells.append(prev_text)

            # 如果连续多行结构相似（都有空白），可能是动态列表
            empty_count = sum(1 for c in cells_text if c == "[空]")
            prev_empty_count = sum(1 for c in prev_cells if c == "[空]")

            if empty_count > 3 and prev_empty_count > 3:
                print(f"  行{row_idx}: {cells_text}  <-- 可能是动态列表行（多空白）")
            else:
                print(f"  行{row_idx}: {cells_text}")
        else:
            print(f"  行{row_idx}: {cells_text}")

    print()