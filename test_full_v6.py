# -*- coding: utf-8 -*-
"""
资信标自动填充系统 - 严格按技术路线文档实现
"""

import os
import sys
import re
import json
import time
from pathlib import Path
from datetime import datetime
from collections import deque

if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

sys.path.insert(0, str(Path(__file__).parent))

from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docxtpl import DocxTemplate
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

KIMI_API_KEY = os.getenv("KIMI_API_KEY", "")
KIMI_BASE_URL = os.getenv("KIMI_BASE_URL", "https://api.moonshot.cn/v1")
# 使用最强的模型
KIMI_MODEL = "moonshot-v1-128k"

TEMPLATE_PATH = r"D:\桌面\投标文件组成模版01.docx"
COMPANY_INFO_PATH = r"D:\桌面\企业信息01.txt"
OUTPUT_PATH = r"D:\桌面\最终输出_完整填充_v6.docx"

print("=" * 70)
print("资信标自动填充系统 - 严格按技术路线实现")
print(f"使用模型: {KIMI_MODEL}")
print("=" * 70)
print(f"开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

client = OpenAI(api_key=KIMI_API_KEY, base_url=KIMI_BASE_URL)


# ============== 辅助函数 ==============

def get_run_format(run):
    font = run.font
    return (
        font.name,
        font.size.pt if font.size else None,
        font.bold,
        font.italic,
        font.underline,
        run._element.rPr
    )


def merge_adjacent_runs(doc: Document) -> int:
    merged_count = 0
    for para in doc.paragraphs:
        merged_count += _merge_runs_in_para(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    merged_count += _merge_runs_in_para(para)
    return merged_count


def _merge_runs_in_para(para) -> int:
    merged_count = 0
    if len(para.runs) <= 1:
        return merged_count
    i = 0
    while i < len(para.runs) - 1:
        current_run = para.runs[i]
        next_run = para.runs[i + 1]
        if get_run_format(current_run) == get_run_format(next_run):
            current_run.text += next_run.text
            next_run._element.getparent().remove(next_run._element)
            merged_count += 1
        else:
            i += 1
    return merged_count


def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("不支持的文档对象")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def get_tables_with_upper_context(doc: Document, window_size: int = 5) -> list:
    upper_context_window = deque(maxlen=window_size)
    table_context_list = []
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if text:
                upper_context_window.append(text)
        elif isinstance(block, Table):
            table_upper_context = "\n".join(upper_context_window)
            serialized_table, unique_cells_info = serialize_table(block)
            table_context_list.append({
                "table": block,
                "upper_context": table_upper_context,
                "serialized_table": serialized_table,
                "unique_cells_info": unique_cells_info
            })
            upper_context_window.clear()
    return table_context_list


def is_blank_cell(text: str) -> bool:
    if not text or not text.strip():
        return True
    text = text.strip()
    if text in ['[空]', '[合并]', '待填写', '待补全']:
        return True
    if re.match(r'^_{2,}$', text):
        return True
    cleaned = re.sub(r'[\s\_\-:\(\)（）\[\]【】]', '', text)
    if not cleaned:
        return True
    return False


def serialize_table(table) -> tuple:
    """按技术路线：唯一单元格去重 + |分隔符序列化"""
    rows_text = []
    unique_cells_info = []
    seen_cell_ids = set()

    for row_idx, row in enumerate(table.rows):
        cells_text = []
        row_cells_info = []

        for cell_idx, cell in enumerate(row.cells):
            cell_id = id(cell._tc)

            if cell_id not in seen_cell_ids:
                seen_cell_ids.add(cell_id)
                text = cell.text.strip()

                if is_blank_cell(text):
                    cells_text.append('[空]')
                    row_cells_info.append({
                        'row_idx': row_idx,
                        'cell_idx': cell_idx,
                        'cell': cell,
                        'original_text': '',
                        'is_empty': True
                    })
                else:
                    text = text.replace('|', '｜').replace('\n', ' ')
                    cells_text.append(text)
                    row_cells_info.append({
                        'row_idx': row_idx,
                        'cell_idx': cell_idx,
                        'cell': cell,
                        'original_text': text,
                        'is_empty': False
                    })
            else:
                cells_text.append('[合并]')
                row_cells_info.append({
                    'row_idx': row_idx,
                    'cell_idx': cell_idx,
                    'cell': cell,
                    'is_merged': True
                })

        rows_text.append('| ' + ' | '.join(cells_text) + ' |')
        unique_cells_info.append(row_cells_info)

    return '\n'.join(rows_text), unique_cells_info


def parse_serialized_table(text: str) -> list:
    rows = []
    for line in text.strip().split('\n'):
        line = line.strip()
        if not line:
            continue
        if '|' in line:
            cells = [c.strip() for c in line.split('|')]
            cells = [c for c in cells if c]
            if cells:
                rows.append(cells)
    return rows


def detect_dynamic_list(table_context: dict) -> bool:
    """检测动态列表：连续多行空白超过70%"""
    serialized = table_context['serialized_table']
    rows = parse_serialized_table(serialized)

    if len(rows) < 3:
        return False

    empty_streak = 0
    max_streak = 0

    for row in rows:
        empty_count = sum(1 for cell in row if cell in ['[空]', '[合并]'])
        total_cells = len(row)

        if total_cells > 0 and empty_count / total_cells > 0.7:
            empty_streak += 1
            max_streak = max(max_streak, empty_streak)
        else:
            empty_streak = 0

    return max_streak >= 3


def call_kimi_api(prompt: str, system_prompt: str, max_tokens: int = 4096) -> str:
    """调用Kimi API"""
    try:
        response = client.chat.completions.create(
            model=KIMI_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=max_tokens
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"  [ERROR] API调用失败: {e}")
        raise


# ============== 按技术路线文档的打标指令 ==============

def call_kimi_tag_paragraphs(text_block: str) -> str:
    """
    普通文本打标指令（技术路线 2.2.6节）
    """
    prompt = f"""【角色】你是一名专业的投标助理,精通各类投标文件模板的语义理解。
【任务】识别文本中所有需要投标人填写、补充或提供的空白处。

【空白处类型包括但不限于】:
1. 下划线:________
2. 括号提示:(项目名称)、[请填写]、（单位盖章）等
3. 散落空格:  年  月  日
4. 冒号留白:投标人:
5. 表格中的空单元格

【打标规则】:
1. 根据上下文语义,为每个空白处生成一个简洁的中文标签
2. 标签格式统一为 {{{{标签名}}}}
3. 必须返回原句,只替换空白处为标签,严禁删减或改动原句中的其他任何字符,包括标点符号
4. 绝对禁止生成如 {{{{地址}}}}、{{{{姓名}}}} 这样指代不明的标签！结合整块文档上下文，生成高度具体的标签
5. 如果空白处原有下划线,请用下划线包裹标签:_{{{{标签}}}}_
6. 输入文本每一行开头都带有固定段落索引格式 [数字]，你必须完整保留每一行的 [数字] 索引，不得删除、修改、打乱顺序、合并多行。

【示例】
输入：
[15] 单位名称：________________
[16] 单位地址：________________
[17] 法定代表人：    年  月  日

输出：
[15] 单位名称：{{{{投标人单位名称}}}}
[16] 单位地址：{{{{投标人单位地址}}}}
[17] 法定代表人：{{{{年}}}}年{{{{月}}}}月{{{{日}}}}日

【待打标文本块】
{text_block}"""

    system_prompt = "你是标签生成器，只输出打标后的内容，完整保留所有[数字]索引。"

    return call_kimi_api(prompt, system_prompt, max_tokens=8192)


def call_kimi_tag_table(serialized_table: str, upper_context: str, is_dynamic: bool) -> str:
    """
    表格打标指令（技术路线 2.3.3节）
    """
    if is_dynamic:
        prompt = f"""【角色】标书模板表格智能打标专家

【任务】接收按|分隔符序列化的表格文本，这是动态列表表格。参考【表格前文背景】，为表格打标，仅处理表格内容。

【判断规则】
这是动态列表：业绩清单、人员名录、设备清单，特征为顶部统一表头、下方多行重复空白行。

【输入】
=== 表格前文背景（仅参考，禁止修改、禁止输出） ===
{upper_context}
=== 待打标表格（|分隔序列化） ===
{serialized_table}

【打标规则】
1. 动态列表：仅保留表头和第一行空白位，删除后续所有重复空行；
2. 在首行第一个单元格开头添加Jinja2循环语法（注意是百分号花括号）：{{%tr for item in 列表名称%}}
3. 单元格使用双花括号标签：{{{{item.字段名}}}}
4. 列表名称根据上下文命名，如"历史业绩列表"、"项目负责人列表"等。
5. 绝对禁止生成如 {{{{地址}}}}、{{{{姓名}}}} 这样指代不明的标签！

【重要语法】
- 循环开头：{{%tr for item in 列表名%}}
- 字段引用：{{{{item.字段名}}}}

【示例】
输入：
=== 表格前文背景 ===
历史业绩表格
=== 待打标表格 ===
| 项目名称 | 建设单位 | 合同金额 |
| [空] | [空] | [空] |
| [空] | [空] | [空] |

输出：
| 项目名称 | 建设单位 | 合同金额 |
| {{%tr for item in 历史业绩列表%}}{{{{item.项目名称}}}} | {{{{item.建设单位}}}} | {{{{item.合同金额}}}} |

【输出要求】
1. 严格保留原有|分隔符，仅替换[空]、增减空行，不改动原有文字和顺序。
2. 只输出打标后的表格文本，禁止输出前文、解释、代码块！"""
    else:
        prompt = f"""【角色】标书模板表格智能打标专家

【任务】接收按|分隔符序列化的表格文本，这是静态表单。参考【表格前文背景】，为表格打标。

【判断规则】
这是静态表单：人员信息、企业资质、固定填报项，特征为属性名+单一项填空。

【输入】
=== 表格前文背景（仅参考，禁止修改、禁止输出） ===
{upper_context}
=== 待打标表格（|分隔序列化） ===
{serialized_table}

【打标规则】
1. 静态表单：所有[空]逐一替换为独立业务标签{{{{字段名}}}}，保留所有行列结构；
2. 标签名必须高度具体，结合前文背景生成。
3. 禁止使用泛标签。

【输出要求】
1. 严格保留原有|分隔符，仅替换[空]。
2. 只输出打标后的表格文本，禁止输出前文、解释、代码块！"""

    system_prompt = "你是表格打标工具，只输出打标后的表格文本。"

    return call_kimi_api(prompt, system_prompt, max_tokens=8192)


def call_kimi_match_info(tags: list, company_info: str) -> dict:
    """
    信息匹配指令（技术路线第3.2节）
    """
    if not tags:
        return {}

    # 分离普通标签和列表标签
    simple_tags = []
    list_tags = []

    for tag in tags:
        if tag.startswith('item.') or tag.startswith('/item.'):
            continue
        elif '列表' in tag or tag.startswith('tr for'):
            list_tags.append(tag)
        else:
            simple_tags.append(tag)

    all_results = {}

    # 处理普通标签
    if simple_tags:
        prompt = f"""【角色】你是一名专业的企业信息分析师。

【输入】
1. 标签列表:{json.dumps(simple_tags, ensure_ascii=False, indent=2)}
2. 企业基础信息文本(长文本)

【企业信息】
{company_info[:5000]}

【任务】请在"企业基础信息文本"中,为"标签列表"里的每一个标签寻找最匹配的具体信息。

【请严格按照以下 JSON 骨架输出，填入真实数据】
{{{{
  "企业名称": "XX科技有限公司",
  "注册资本": "1000万元",
  "法定代表人姓名": "张三",
  "企业网址": "待补全",
  "资质证书号": "待补全"
}}}}

【规则】
1. 如果企业信息中找不到该标签对应的内容,Value 请返回"待补全"
2. 不要编造信息,必须基于提供的企业信息文本
3. 金额、日期等数据保持原文格式
4. 直接返回JSON对象，不要输出任何说明或解释性内容"""

        system_prompt = "你是数据匹配工具，只输出JSON。"

        try:
            result = call_kimi_api(prompt, system_prompt, max_tokens=8192)

            # 清理并解析JSON
            result = re.sub(r'^```json\s*', '', result)
            result = re.sub(r'^```\s*', '', result)
            result = re.sub(r'```\s*$', '', result)

            json_match = re.search(r'\{[\s\S]*\}', result)
            if json_match:
                result = json_match.group()

            parsed = json.loads(result)
            all_results.update(parsed)

            # 确保所有标签都有值
            for tag in simple_tags:
                if tag not in all_results:
                    all_results[tag] = "待补全"

        except Exception as e:
            print(f"  [WARN] 普通标签匹配失败: {e}")
            for tag in simple_tags:
                all_results[tag] = "待补全"

    # 处理列表标签
    for list_tag in list_tags:
        # 提取列表名称
        match = re.search(r'for\s+item\s+in\s+(\S+)', list_tag)
        if match:
            list_name = match.group(1)
        else:
            list_name = list_tag.replace('列表', '').replace('tr for item in ', '')

        prompt = f"""【角色】你是一名专业的企业信息分析师。

【特殊说明：列表类型】
列表名称：{list_name}

【企业信息】
{company_info}

【任务】请在企业信息中寻找所有符合"{list_name}"的项目，提取完整的列表。

【请输出成 JSON 数组形式】
[
  {{{{"建设单位": "待填入", "工程名称": "待填入", "建设规模": "待填入", "质量": "待填入"}}}},
  {{{{"建设单位": "待填入", "工程名称": "待填入", "建设规模": "待填入", "quality": "待填入"}}}}
]

【规则】
1. 找到所有符合条件的项目
2. 如果找不到，返回空数组 []
3. 直接返回JSON数组"""

        system_prompt = "你是数据匹配工具，只输出JSON数组。"

        try:
            result = call_kimi_api(prompt, system_prompt, max_tokens=8192)

            result = re.sub(r'^```json\s*', '', result)
            result = re.sub(r'^```\s*', '', result)
            result = re.sub(r'```\s*$', '', result)

            array_match = re.search(r'\[[\s\S]*\]', result)
            if array_match:
                result = array_match.group()
                parsed = json.loads(result)
                all_results[list_name] = parsed
            else:
                all_results[list_name] = []

        except Exception as e:
            print(f"  [WARN] 列表{list_name}匹配失败: {e}")
            all_results[list_name] = []

    # 后处理签字盖章类标签
    for tag in all_results:
        if ('签字' in tag or '盖章' in tag) and all_results[tag] in ["待补全", ""]:
            all_results[tag] = f"[需{tag}]"

    return all_results


def extract_tags_from_doc(doc: Document) -> list:
    """提取所有标签"""
    tags = set()
    pattern = r'\{\{(.*?)\}\}'

    for para in doc.paragraphs:
        matches = re.findall(pattern, para.text)
        tags.update(matches)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    matches = re.findall(pattern, para.text)
                    tags.update(matches)

    return sorted(list(tags))


def replace_text_in_paragraph(para, old_text: str, new_text) -> bool:
    if not old_text:
        return False
    new_text = str(new_text) if new_text is not None else ""
    full_text = para.text
    if old_text not in full_text:
        return False

    start_idx = full_text.find(old_text)
    end_idx = start_idx + len(old_text)

    run_start_idx = 0
    matching_runs = []

    for run in para.runs:
        run_end_idx = run_start_idx + len(run.text)
        if (run_start_idx <= start_idx < run_end_idx or
            run_start_idx < end_idx <= run_end_idx or
            (start_idx <= run_start_idx and run_end_idx <= end_idx)):
            matching_runs.append({
                'run': run,
                'run_start': run_start_idx,
                'run_end': run_end_idx,
                'text_start': max(start_idx, run_start_idx),
                'text_end': min(end_idx, run_end_idx)
            })
        run_start_idx = run_end_idx

    if not matching_runs:
        return False

    if len(matching_runs) == 1:
        run_info = matching_runs[0]
        run = run_info['run']
        run_local_start = run_info['text_start'] - run_info['run_start']
        run_local_end = run_info['text_end'] - run_info['run_start']
        run.text = run.text[:run_local_start] + new_text + run.text[run_local_end:]
        return True

    first_run_info = matching_runs[0]
    first_run = first_run_info['run']
    first_run_local_start = first_run_info['text_start'] - first_run_info['run_start']
    first_run.text = first_run.text[:first_run_local_start] + new_text

    for i, run_info in enumerate(matching_runs):
        if i == 0:
            continue
        run = run_info['run']
        run_local_end = run_info['text_end'] - run_info['run_start']
        if run_local_end < len(run.text):
            run.text = run.text[run_local_end:]
        else:
            run.text = ""
    return True


def apply_table_tags(table: Table, tagged_rows: list, unique_cells_info: list, is_dynamic: bool):
    """将打标结果应用回表格"""
    if is_dynamic:
        # 删除多余行
        rows_to_keep = len(tagged_rows)
        while len(table.rows) > rows_to_keep:
            tr = table.rows[-1]._tr
            tr.getparent().remove(tr)

    # 应用标签
    for row_idx, row in enumerate(table.rows):
        if row_idx >= len(tagged_rows):
            break

        tagged_cells = tagged_rows[row_idx]
        cells_info = unique_cells_info[row_idx] if row_idx < len(unique_cells_info) else []

        for cell_idx, cell in enumerate(row.cells):
            if cell_idx >= len(tagged_cells):
                break

            cell_info = cells_info[cell_idx] if cell_idx < len(cells_info) else {}

            if cell_info.get('is_merged'):
                continue

            tagged_text = tagged_cells[cell_idx]

            # 清空单元格
            for para in cell.paragraphs:
                for run in para.runs:
                    run.text = ""

            # 写入新内容
            if cell.paragraphs:
                if cell.paragraphs[0].runs:
                    cell.paragraphs[0].runs[0].text = tagged_text
                else:
                    cell.paragraphs[0].add_run(tagged_text)


# ============== 主流程 ==============

print("\n【Step 1】加载文档")
print("-" * 70)

doc = Document(TEMPLATE_PATH)
print(f"[OK] 文档加载成功")
print(f"  - 段落数量: {len(doc.paragraphs)}")
print(f"  - 表格数量: {len(doc.tables)}")

with open(COMPANY_INFO_PATH, 'r', encoding='utf-8') as f:
    company_info = f.read()
print(f"[OK] 企业信息已加载，共 {len(company_info)} 字符")

print("\n【Step 2】合并碎片化Run")
print("-" * 70)

merged = merge_adjacent_runs(doc)
print(f"[OK] Run合并完成，共合并 {merged} 个")

print("\n【Step 3】AI辅助打标 - 段落（按技术路线2.2.6节指令）")
print("-" * 70)

# 获取有效段落
valid_paragraphs = [(i, p) for i, p in enumerate(doc.paragraphs) if p.text.strip()]
print(f"发现 {len(valid_paragraphs)} 个有效段落")

# 构建带索引的文本块
paragraph_processed = 0

for batch_start in range(0, len(valid_paragraphs), 20):
    batch_paras = valid_paragraphs[batch_start:batch_start + 20]

    # 构建输入
    text_block = "\n".join([f"[{idx}] {p.text.strip()}" for idx, p in batch_paras])

    print(f"  处理批次 {batch_start//20 + 1}...")

    try:
        tagged_text = call_kimi_tag_paragraphs(text_block)

        # 解析结果
        pattern = r'^\[(\d+)\]\s*(.*)$'
        tagged_map = {}

        for line in tagged_text.splitlines():
            line = line.strip()
            if not line:
                continue
            m = re.match(pattern, line)
            if m:
                idx = int(m.group(1))
                content = m.group(2)
                tagged_map[idx] = content

        # 应用
        for orig_idx, para in batch_paras:
            if orig_idx in tagged_map:
                new_text = tagged_map[orig_idx]
                if "{{" in new_text:
                    # 替换段落内容
                    full_text = para.text
                    replace_text_in_paragraph(para, full_text, new_text)
                    paragraph_processed += 1

        print(f"    打标成功 {len(tagged_map)} 个")
        time.sleep(1)

    except Exception as e:
        print(f"    [WARN] 失败: {e}")

print(f"[OK] 段落打标完成，成功处理 {paragraph_processed} 个")

print("\n【Step 4】AI辅助打标 - 表格（按技术路线2.3.3节指令）")
print("-" * 70)

tables_with_context = get_tables_with_upper_context(doc)
print(f"发现 {len(tables_with_context)} 个表格")

table_processed = 0
list_names = []

for table_idx, table_ctx in enumerate(tables_with_context):
    table = table_ctx['table']
    upper_context = table_ctx['upper_context']
    serialized_table = table_ctx['serialized_table']
    unique_cells_info = table_ctx['unique_cells_info']

    is_dynamic = detect_dynamic_list(table_ctx)
    print(f"  处理表格 {table_idx + 1} (动态列表: {is_dynamic})...")

    try:
        tagged_table = call_kimi_tag_table(serialized_table, upper_context, is_dynamic)
        tagged_rows = parse_serialized_table(tagged_table)

        apply_table_tags(table, tagged_rows, unique_cells_info, is_dynamic)

        # 记录列表名称
        if is_dynamic and tagged_rows:
            first_cell = tagged_rows[0][0] if tagged_rows[0] else ""
            match = re.search(r'for\s+item\s+in\s+(\S+)', first_cell)
            if match:
                list_names.append(match.group(1))

        table_processed += 1
        print(f"    完成")
        time.sleep(1)

    except Exception as e:
        print(f"    [WARN] 失败: {e}")

print(f"[OK] 表格打标完成，成功处理 {table_processed} 个")
print(f"发现动态列表: {list_names}")

print("\n【Step 5】标签提取")
print("-" * 70)

tags = extract_tags_from_doc(doc)
print(f"[OK] 提取到 {len(tags)} 个标签")
print(f"  前20个: {tags[:20]}")

print("\n【Step 6】企业信息匹配（按技术路线第3.2节指令）")
print("-" * 70)

start_time = time.time()
matched_data = call_kimi_match_info(tags, company_info)
elapsed = time.time() - start_time

matched_count = sum(1 for v in matched_data.values() if v not in ["待补全", [], ""])
print(f"[OK] 匹配完成，耗时 {elapsed:.2f}秒")
print(f"  成功匹配: {matched_count}/{len(matched_data)} 个")

# 显示部分匹配结果
print("\n  匹配结果预览:")
for i, (k, v) in enumerate(list(matched_data.items())[:15]):
    v_preview = str(v)[:40] if isinstance(v, str) else str(v)[:60]
    print(f"    {k}: {v_preview}")

print("\n【Step 7】使用docxtpl渲染")
print("-" * 70)

# 保存临时模板
temp_template = r"D:\桌面\temp_template_v6.docx"
doc.save(temp_template)

try:
    tpl = DocxTemplate(temp_template)
    tpl.render(matched_data)
    tpl.save(OUTPUT_PATH)
    print(f"[OK] docxtpl渲染完成")
except Exception as e:
    print(f"[WARN] docxtpl失败，使用普通保存: {e}")
    doc.save(OUTPUT_PATH)
finally:
    try:
        os.remove(temp_template)
    except:
        pass

print("\n" + "=" * 70)
print("处理完成!")
print("=" * 70)
print(f"结束时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print(f"\n处理结果摘要:")
print(f"  - 段落打标: {paragraph_processed} 个")
print(f"  - 表格打标: {table_processed} 个")
print(f"  - 标签提取: {len(tags)} 个")
print(f"  - 信息匹配: {matched_count}/{len(matched_data)} 个成功")
print(f"  - 动态列表: {len(list_names)} 个")
print(f"\n输出文件: {OUTPUT_PATH}")
