"""
资信标自动填充系统 - 后端服务
基于 FastAPI + python-docx + Kimi API 实现
按照技术路线文档完整实现所有功能
"""

import os
import re
import json
import time
import uuid
import logging
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from collections import deque, OrderedDict

from fastapi import FastAPI, File, UploadFile, Form, HTTPException, Body, Request
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from dotenv import load_dotenv
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from openai import OpenAI
import aiofiles

# 加载环境变量
load_dotenv()

# 配置
KIMI_API_KEY = os.getenv("KIMI_API_KEY", "")
KIMI_BASE_URL = os.getenv("KIMI_BASE_URL", "https://api.moonshot.cn/v1")
KIMI_MODEL = os.getenv("KIMI_MODEL", "moonshot-v1-32k")
MAX_FILE_SIZE = int(os.getenv("MAX_FILE_SIZE", "10485760"))
UPLOAD_DIR = Path(os.getenv("UPLOAD_DIR", "uploads"))
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "outputs"))
LOG_DIR = Path("logs")

# 确保目录存在
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

# 初始化 Kimi 客户端
client = OpenAI(api_key=KIMI_API_KEY, base_url=KIMI_BASE_URL)

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / 'app.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


# ============== 数据模型 ==============

class ProcessStep(BaseModel):
    step: int
    name: str
    status: str
    time_ms: Optional[int] = None
    message: Optional[str] = None


class ProcessResponse(BaseModel):
    success: bool
    task_id: str
    steps: List[ProcessStep]
    tags: Optional[List[str]] = None
    matched_data: Optional[Dict[str, Any]] = None
    download_url: Optional[str] = None
    error: Optional[str] = None


class ProcessRequest(BaseModel):
    filename: str
    company_info: str


# ============== 核心算法：跨Run替换 ==============

def replace_text_in_paragraph(para, old_text: str, new_text) -> bool:
    """
    跨 Run 替换文本，保留格式
    核心原则：每次替换前重新读取段落文本，解决索引偏移问题
    """
    if not old_text:
        return False

    new_text = str(new_text) if new_text is not None else ""

    # 关键：每次都重新获取最新文本，解决多标签索引偏移
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

    # 跨Run处理：保留首个Run格式
    first_run_info = matching_runs[0]
    first_run = first_run_info['run']
    first_run_local_start = first_run_info['text_start'] - first_run_info['run_start']

    new_run_text = first_run.text[:first_run_local_start] + new_text
    first_run.text = new_run_text

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


def replace_all_in_paragraph(para, replacements: Dict[str, str]) -> int:
    """
    在单个段落中执行多次替换
    每次替换后重新读取文本，避免索引偏移
    """
    count = 0
    for old_text, new_text in replacements.items():
        if old_text in para.text:
            if replace_text_in_paragraph(para, old_text, new_text):
                count += 1
    return count


# ============== 合并碎片化Run ==============

def get_run_format(run):
    """获取Run的格式特征"""
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
    """合并格式相同的相邻Run，返回合并数量"""
    merged_count = 0

    # 处理正文段落
    for para in doc.paragraphs:
        merged_count += _merge_runs_in_para(para)

    # 处理表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    merged_count += _merge_runs_in_para(para)

    return merged_count


def _merge_runs_in_para(para) -> int:
    """合并单个段落中格式相同的相邻Run"""
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


# ============== 文档顺序遍历（解决表格孤岛问题） ==============

def iter_block_items(parent):
    """
    按文档物理顺序，逐块返回 Paragraph 和 Table 对象
    解决段落与表格顺序丢失问题
    """
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


# ============== 表格上下文增强处理 ==============

def get_tables_with_upper_context(doc: Document, window_size: int = 5) -> List[Dict]:
    """
    遍历文档，为每个表格收集上方N段非空上文
    解决表格孤岛问题
    """
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

            # 关键：处理完表格清空窗口，避免污染下一个表格
            upper_context_window.clear()

    return table_context_list


# ============== 表格序列化（唯一单元格去重 + |分隔符） ==============

def serialize_table(table) -> Tuple[str, List[Dict]]:
    """
    将Word表格序列化为 | 分隔符文本
    返回：(序列化文本, 唯一单元格信息列表)

    核心功能：
    1. 唯一单元格去重提取（解决合并单元格重复问题）
    2. 使用 | 分隔符伪Markdown（不包含分隔行）
    3. 空单元格统一替换为 [空]
    """
    rows_text = []
    unique_cells_info = []  # 存储每个唯一单元格的引用信息
    seen_cell_ids = set()  # 用于去重

    for row_idx, row in enumerate(table.rows):
        cells_text = []
        row_cells_info = []

        for cell_idx, cell in enumerate(row.cells):
            # 使用内存地址作为唯一标识
            cell_id = id(cell._tc)

            if cell_id not in seen_cell_ids:
                seen_cell_ids.add(cell_id)
                text = cell.text.strip()

                # 处理空白单元格
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
                    # 转义分隔符
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
                # 跳过重复单元格，但仍需要占位
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


def is_blank_cell(text: str) -> bool:
    """判断单元格是否为空白"""
    if not text or not text.strip():
        return True

    blank_patterns = [
        r'^_{2,}$',  # 下划线
        r'^\s*$',    # 纯空格
        r'^\[空\]$',
        r'^待填写$',
        r'^待补全$',
        r'^□+$',     # 方框
    ]

    text_clean = text.strip()
    for pattern in blank_patterns:
        if re.match(pattern, text_clean):
            return True

    # 移除空白字符后判断
    cleaned = re.sub(r'[\s\_\-:\(\)（）\[\]【】□]', '', text_clean)
    if not cleaned:
        return True

    return False


# ============== 解析AI返回的表格 ==============

def parse_serialized_table(text: str) -> List[List[str]]:
    """
    解析 | 分隔符表格文本为二维数组
    """
    rows = []
    for line in text.strip().split('\n'):
        line = line.strip()
        if not line:
            continue
        if '|' in line:
            cells = [c.strip() for c in line.split('|')]
            cells = [c for c in cells if c]  # 过滤空字符串
            if cells:
                rows.append(cells)
    return rows


# ============== 动态列表识别 ==============

def detect_dynamic_list(table_context: Dict) -> Tuple[bool, Optional[int]]:
    """
    判断表格是否为动态列表
    返回：(是否动态列表, 数据起始行索引)
    """
    serialized = table_context['serialized_table']
    rows = parse_serialized_table(serialized)

    if len(rows) < 3:
        return False, None

    # 统计空白行的连续数量
    empty_streak = 0
    max_streak = 0
    streak_start = -1
    data_start_row = -1

    for i, row in enumerate(rows):
        empty_count = sum(1 for cell in row if cell in ['[空]', '[合并]'])
        total_cells = len(row)

        # 如果一行中超过70%是空白
        if total_cells > 0 and empty_count / total_cells > 0.7:
            if empty_streak == 0:
                streak_start = i
            empty_streak += 1
            max_streak = max(max_streak, empty_streak)
        else:
            if empty_streak >= 3:
                # 找到了连续的空白行，可能是动态列表
                data_start_row = streak_start
            empty_streak = 0

    # 如果有超过3个连续空白行，判定为动态列表
    is_dynamic = max_streak >= 3
    return is_dynamic, data_start_row


# ============== 日志记录器 ==============

class StepLogger:
    def __init__(self, task_id: str):
        self.task_id = task_id
        self.log_file = LOG_DIR / f"{task_id}_steps.json"
        self.logs = {
            "task_id": task_id,
            "start_time": datetime.now().isoformat(),
            "steps": []
        }

    def log_step(self, step_name: str, status: str, details: Dict = None):
        step_data = {
            "timestamp": datetime.now().isoformat(),
            "step_name": step_name,
            "status": status
        }
        if details:
            step_data["details"] = details
        self.logs["steps"].append(step_data)
        self._save()
        logger.info(f"[{self.task_id}] {step_name}: {status}")

    def log_api_call(self, step: str, prompt: str, response: str = None, error: str = None):
        api_log = {
            "timestamp": datetime.now().isoformat(),
            "step": step,
            "prompt_preview": prompt[:500] + "..." if len(prompt) > 500 else prompt,
            "response_preview": response[:500] + "..." if response and len(response) > 500 else response,
            "error": error
        }
        api_log_file = LOG_DIR / f"{self.task_id}_api_calls.jsonl"
        with open(api_log_file, 'a', encoding='utf-8') as f:
            f.write(json.dumps(api_log, ensure_ascii=False) + '\n')

    def _save(self):
        with open(self.log_file, 'w', encoding='utf-8') as f:
            json.dump(self.logs, f, ensure_ascii=False, indent=2)


# ============== Kimi API 调用 ==============

def call_kimi_api(prompt: str, system_prompt: str, max_tokens: int = 2048) -> str:
    """通用Kimi API调用"""
    try:
        start_time = time.time()
        response = client.chat.completions.create(
            model=KIMI_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=max_tokens
        )
        elapsed = time.time() - start_time
        result = response.choices[0].message.content.strip()
        logger.info(f"API调用成功，耗时{elapsed:.2f}秒")
        return result
    except Exception as e:
        logger.error(f"API调用失败: {e}")
        raise


def call_kimi_semantic_chunk(paragraphs_text: str, total_paragraphs: int) -> List[int]:
    """
    调用Kimi API进行语义分块
    返回语义块起始索引数组
    """
    prompt = f"""【角色】文档结构分析专家
【任务】我将提供一份标书文档纯文本，每一行开头有中括号`[数字]`代表段落物理序号。
请根据语义识别所有全新业务模块、核心章节的**起始段落序号**。

【判断依据】
1. 标准章节标题：一、xxx、1.1 xxx、第x章等；
2. 无标准标号但独立成行、开启全新业务主题的段落；
3. 忽略普通正文换行、无关短句。

【输出要求】
1. 仅输出语义块起始段落序号，以纯JSON一维数组返回；
2. 禁止输出解释、多余文字、代码块；
3. 严格只返回数字数组，示例：[0,15,34]

【文档内容】
{paragraphs_text[:6000]}"""

    system_prompt = "你是文档结构分析工具，只输出JSON数组。"

    try:
        result = call_kimi_api(prompt, system_prompt, max_tokens=1024)
        return parse_block_indices(result, total_paragraphs)
    except Exception as e:
        logger.warning(f"语义分块失败，使用兜底方案: {e}")
        return fallback_chunk_indices(total_paragraphs)


def parse_block_indices(response_text: str, total_paragraphs: int, chunk_size: int = 20) -> List[int]:
    """解析AI返回的索引数组，失败自动兜底"""
    try:
        match = re.search(r'\[[\d,\s]*\]', response_text)
        if not match:
            raise ValueError("未识别到索引数组")

        indices = json.loads(match.group(0))

        if not isinstance(indices, list) or not all(isinstance(x, int) for x in indices):
            raise ValueError("索引包含非整数")

        # 过滤越界索引、去重、补首尾边界
        valid_indices = [idx for idx in indices if 0 <= idx < total_paragraphs]
        if 0 not in valid_indices:
            valid_indices.insert(0, 0)
        valid_indices.append(total_paragraphs)
        final_indices = sorted(list(set(valid_indices)))

        # 校验是否有效分块
        if len(final_indices) <= 2 and total_paragraphs > chunk_size:
            raise ValueError("AI未完成有效语义分块")

        return final_indices
    except Exception as e:
        logger.warning(f"索引解析失败: {e}")
        return fallback_chunk_indices(total_paragraphs, chunk_size)


def fallback_chunk_indices(total_paragraphs: int, chunk_size: int = 20) -> List[int]:
    """兜底：固定段落分块"""
    indices = list(range(0, total_paragraphs, chunk_size))
    if indices[-1] != total_paragraphs:
        indices.append(total_paragraphs)
    return indices


def call_kimi_tag_paragraphs(text_block: str) -> str:
    """调用Kimi API为段落块打标"""
    prompt = f"""【角色】你是一名专业的投标助理，精通各类投标文件模板的语义理解。
【任务】识别文本中所有需要投标人填写、补充或提供的空白处。

【空白处类型包括但不限于】:
1. 下划线:________
2. 括号提示:(项目名称)、[请填写]、（单位盖章）
3. 散落空格:  年  月  日
4. 冒号留白:投标人:
5. 表格中的空单元格

【打标规则】:
1. 根据上下文语义，为每个空白处生成一个简洁的中文标签
2. 标签格式统一为 {{{{标签名}}}}
3. 必须返回原句，只替换空白处为标签，严禁删减或改动原句中的其他任何字符
4. 绝对禁止生成如 {{{{地址}}}}、{{{{姓名}}}} 这样指代不明的标签！结合整块文档上下文，生成高度具体的标签
5. 如果空白处原有下划线，请用下划线包裹标签:_{{{{标签}}}}_
6. 输入文本每一行开头都带有固定段落索引格式 [数字]，你必须完整保留每一行的 [数字] 索引

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
    return call_kimi_api(prompt, system_prompt, max_tokens=4096)


def call_kimi_tag_table(serialized_table: str, upper_context: str, is_dynamic: bool) -> str:
    """调用Kimi API为表格打标"""
    if is_dynamic:
        prompt = f"""【角色】标书模板表格智能打标专家

【任务】接收按|分隔符序列化的表格文本，这是动态列表表格（业绩清单、人员名录等）。
参考【表格前文背景】，为表格打标。

【输入】
=== 表格前文背景（仅参考，禁止修改、禁止输出） ===
{upper_context}
=== 待打标表格（|分隔序列化） ===
{serialized_table}

【打标规则】
1. 仅保留表头和第一行空白位，删除后续所有重复空行
2. 在首行添加docxtpl循环语法：{{{{tr for item in 列表名称}}}}
3. 单元格使用泛型标签：{{{{item.字段名}}}}
4. 列表名称根据上下文命名（如"项目负责人业绩列表"、"历史业绩列表"等）
5. 标签名必须高度具体，禁止使用{{{{姓名}}}}、{{{{地址}}}}等泛标签

【示例】
输入：
| 项目名称 | 建设单位 | 合同金额 |
| [空] | [空] | [空] |
| [空] | [空] | [空] |

输出：
| 项目名称 | 建设单位 | 合同金额 |
| {{{{tr for item in 历史业绩列表}}}}{{{{item.项目名称}}}} | {{{{item.建设单位}}}} | {{{{item.合同金额}}}} |

【输出要求】
1. 严格保留原有|分隔符，仅替换[空]、增减空行
2. 只输出打标后的表格文本"""
    else:
        prompt = f"""【角色】标书模板表格智能打标专家

【任务】接收按|分隔符序列化的表格文本，这是静态表单（人员信息、企业资质等）。
参考【表格前文背景】，为表格打标。

【输入】
=== 表格前文背景（仅参考，禁止修改、禁止输出） ===
{upper_context}
=== 待打标表格（|分隔序列化） ===
{serialized_table}

【打标规则】
1. 所有[空]逐一替换为独立业务标签{{{{字段名}}}}
2. 保留所有行列结构，不删除任何行
3. 标签名必须高度具体，结合前文背景生成
4. 禁止使用{{{{姓名}}}}、{{{{地址}}}}等泛标签，应使用{{{{投标人单位名称}}}}、{{{{项目负责人姓名}}}}等

【输出要求】
1. 严格保留原有|分隔符
2. 只输出打标后的表格文本"""

    system_prompt = "你是表格打标工具，只输出打标后的表格，不输出任何解释。"
    return call_kimi_api(prompt, system_prompt, max_tokens=4096)


def call_kimi_match_info(tags: List[str], company_info: str) -> Dict[str, Any]:
    """调用Kimi API进行信息匹配（支持列表类型）"""
    if not tags:
        return {}

    batch_size = 60
    all_results = {}

    for batch_start in range(0, len(tags), batch_size):
        batch_tags = tags[batch_start:batch_start + batch_size]

        prompt = f"""【角色】你是一名专业的企业信息匹配专家。

【输入】
1. 标签列表:{json.dumps(batch_tags, ensure_ascii=False, indent=2)}
2. 企业基础信息文本(长文本)

【企业信息】：
{company_info[:4000]}

【任务】请在"企业基础信息文本"中，为"标签列表"里的每一个标签寻找最匹配的具体信息。

【特殊说明：列表类型】
当你看到标签包含"列表"、"业绩"、"人员"等关键字时，说明这是一组列表数据。
请在企业信息中寻找所有符合条件的项目，提取完整的列表，输出成 JSON 数组形式。

【请严格按照以下 JSON 骨架输出】
{{
  "企业名称": "XX科技有限公司",
  "注册资本": "1000万元",
  "历史业绩列表": [
    {{"建设单位": "XX", "工程名称": "XX", "建设规模": "XX", "质量": "合格"}},
    {{"建设单位": "YY", "工程名称": "YY", "建设规模": "YY", "quality": "合格"}}
  ]
}}

【规则】
1. 如果企业信息中找不到该标签对应的内容，Value 请返回"待补全"
2. 不要编造信息，必须基于提供的企业信息文本
3. 金额、日期等数据保持原文格式
4. 直接返回JSON对象，不要输出任何说明"""

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

            # 确保所有标签都在结果中
            for tag in batch_tags:
                if tag not in parsed:
                    parsed[tag] = "待补全"
                    logger.warning(f"标签 '{tag}' 未在AI返回结果中")

            all_results.update(parsed)

        except Exception as e:
            logger.error(f"信息匹配失败: {e}")
            for tag in batch_tags:
                all_results[tag] = "待补全"

    # 后处理：签字盖章类标签
    for tag in all_results:
        if ('签字' in tag or '盖章' in tag) and all_results[tag] == "待补全":
            all_results[tag] = f"[需{tag}]"

    return all_results


# ============== 解析打标结果 ==============

def parse_tagged_paragraphs(text: str) -> Dict[int, str]:
    """解析带索引的打标段落文本"""
    line_map = {}
    pattern = r'^\[(\d+)\]\s*(.*)$'

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        m = re.match(pattern, line)
        if m:
            idx = int(m.group(1))
            content = m.group(2)
            line_map[idx] = content

    return line_map


# ============== 提取标签 ==============

def extract_tags_from_doc(doc: Document) -> List[str]:
    """从文档中提取所有 {{标签}}"""
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


# ============== 文档处理主流程 ==============

def process_document(
    doc: Document,
    company_info: str,
    step_logger: StepLogger,
    progress_callback=None
) -> tuple:
    """处理文档的主流程"""

    # Step 1: 合并碎片化 Run
    step_logger.log_step("文档预处理", "running")
    if progress_callback:
        progress_callback(1, "合并文档格式...")

    start_time = time.time()
    merged_count = merge_adjacent_runs(doc)
    elapsed = int((time.time() - start_time) * 1000)
    step_logger.log_step("文档预处理", "completed", {"time_ms": elapsed, "merged_runs": merged_count})

    # Step 2: AI 辅助打标 - 正文段落（索引式语义分块）
    step_logger.log_step("AI辅助打标-段落", "running")
    if progress_callback:
        progress_callback(2, "AI正在识别空白处...")

    start_time = time.time()

    # 2.1 获取有效段落
    valid_paragraphs = []
    for para in doc.paragraphs:
        txt = para.text.strip()
        if txt:
            valid_paragraphs.append(para)

    # 2.2 调用AI进行语义分块
    numbered_lines = [f"[{idx}] {p.text.strip()}" for idx, p in enumerate(valid_paragraphs)]
    numbered_text = "\n".join(numberd_lines[:100])  # 限制长度

    try:
        block_indices = call_kimi_semantic_chunk(numberd_text, len(valid_paragraphs))
        logger.info(f"语义分块结果: {block_indices}")
    except Exception as e:
        logger.warning(f"语义分块失败，使用全文处理: {e}")
        block_indices = [0, len(valid_paragraphs)]

    # 2.3 按块打标
    paragraph_processed = 0
    for i in range(len(block_indices) - 1):
        start_idx = block_indices[i]
        end_idx = block_indices[i + 1]

        # 提取块内段落
        block_paras = valid_paragraphs[start_idx:end_idx]
        if not block_paras:
            continue

        block_text = "\n".join([f"[{start_idx + j}] {p.text.strip()}" for j, p in enumerate(block_paras)])

        # 调用AI打标
        try:
            tagged_text = call_kimi_tag_paragraphs(block_text)
            tagged_map = parse_tagged_paragraphs(tagged_text)

            # 应用打标结果
            for idx_offset, para in enumerate(block_paras):
                global_idx = start_idx + idx_offset
                if global_idx in tagged_map:
                    new_text = tagged_map[global_idx]
                    if "{{" in new_text and new_text != para.text:
                        replace_text_in_paragraph(para, para.text, new_text)
                        paragraph_processed += 1

            step_logger.log_api_call(f"paragraph_block_{i}", block_text[:300], tagged_text[:300])

        except Exception as e:
            logger.error(f"段落块 {i} 打标失败: {e}")

    elapsed = int((time.time() - start_time) * 1000)
    step_logger.log_step("AI辅助打标-段落", "completed", {"time_ms": elapsed, "processed": paragraph_processed})

    # Step 3: AI 辅助打标 - 表格（上下文增强 + 唯一单元格）
    step_logger.log_step("AI辅助打标-表格", "running")
    if progress_callback:
        progress_callback(3, "处理表格...")

    start_time = time.time()
    table_processed = 0

    # 获取带上下文的表格
    tables_with_context = get_tables_with_upper_context(doc)
    logger.info(f"发现 {len(tables_with_context)} 个表格")

    for table_idx, table_ctx in enumerate(tables_with_context[:10]):
        table = table_ctx['table']
        upper_context = table_ctx['upper_context']
        serialized_table = table_ctx['serialized_table']
        unique_cells_info = table_ctx['unique_cells_info']

        # 检测是否为动态列表
        is_dynamic, data_start_row = detect_dynamic_list(table_ctx)

        try:
            # 调用AI打标
            tagged_table = call_kimi_tag_table(serialized_table, upper_context, is_dynamic)

            step_logger.log_api_call(f"table_{table_idx}", serialized_table[:500], tagged_table[:500])

            # 解析打标结果
            tagged_rows = parse_serialized_table(tagged_table)

            # 应用回表格
            apply_table_tags(table, tagged_rows, unique_cells_info, is_dynamic)

            table_processed += 1
            logger.info(f"表格 {table_idx} 处理完成 (动态列表: {is_dynamic})")

        except Exception as e:
            logger.error(f"处理表格 {table_idx} 失败: {e}")
            step_logger.log_step(f"表格_{table_idx}_错误", "error", {"error": str(e)})

    elapsed = int((time.time() - start_time) * 1000)
    step_logger.log_step("AI辅助打标-表格", "completed", {"time_ms": elapsed, "processed": table_processed})

    # Step 4: 标签提取
    step_logger.log_step("标签提取", "running")
    if progress_callback:
        progress_callback(4, "提取标签...")

    start_time = time.time()
    tags = extract_tags_from_doc(doc)
    elapsed = int((time.time() - start_time) * 1000)
    step_logger.log_step("标签提取", "completed", {"time_ms": elapsed, "count": len(tags), "tags": tags})

    # Step 5: 企业信息匹配
    step_logger.log_step("企业信息匹配", "running")
    if progress_callback:
        progress_callback(5, "匹配企业信息...")

    start_time = time.time()
    matched_data = call_kimi_match_info(tags, company_info)
    elapsed = int((time.time() - start_time) * 1000)
    step_logger.log_step("企业信息匹配", "completed", {"time_ms": elapsed, "matched_count": len(matched_data)})

    # Step 6: 内容填充
    step_logger.log_step("内容填充", "running")
    if progress_callback:
        progress_callback(6, "填充内容...")

    start_time = time.time()
    replacement_count = 0

    # 处理正文
    for para in doc.paragraphs:
        for tag, value in matched_data.items():
            placeholder = f"{{{{{tag}}}}}"
            if placeholder in para.text:
                if replace_text_in_paragraph(para, placeholder, value):
                    replacement_count += 1

    # 处理表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for tag, value in matched_data.items():
                        placeholder = f"{{{{{tag}}}}}"
                        if placeholder in para.text:
                            if replace_text_in_paragraph(para, placeholder, value):
                                replacement_count += 1

    elapsed = int((time.time() - start_time) * 1000)
    step_logger.log_step("内容填充", "completed", {"time_ms": elapsed, "count": replacement_count})

    details = {
        "steps": [
            {"step": 1, "name": "文档预处理", "time_ms": 0, "merged_runs": merged_count},
            {"step": 2, "name": "AI辅助打标-段落", "time_ms": 0, "processed": paragraph_processed},
            {"step": 3, "name": "AI辅助打标-表格", "time_ms": 0, "processed": table_processed},
            {"step": 4, "name": "标签提取", "time_ms": 0, "count": len(tags)},
            {"step": 5, "name": "企业信息匹配", "time_ms": 0},
            {"step": 6, "name": "内容填充", "time_ms": 0, "count": replacement_count}
        ],
        "tags": tags,
        "matched_data": matched_data,
        "replacements": replacement_count
    }

    return doc, details


def apply_table_tags(table: Table, tagged_rows: List[List[str]], unique_cells_info: List, is_dynamic: bool):
    """将打标结果应用回Word表格"""

    # 处理动态列表：删除多余空行
    if is_dynamic and len(tagged_rows) < len(table.rows):
        # 保留表头和第一数据行，删除其余空行
        rows_to_keep = len(tagged_rows)
        rows_to_delete = len(table.rows) - rows_to_keep

        for _ in range(rows_to_delete):
            if len(table.rows) > rows_to_keep:
                # 删除最后一行
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

            # 跳过合并单元格
            if cell_info.get('is_merged'):
                continue

            tagged_text = tagged_cells[cell_idx]

            # 如果包含标签，更新单元格
            if '{{' in tagged_text and '}}' in tagged_text:
                for para in cell.paragraphs:
                    if para.text.strip():
                        replace_text_in_paragraph(para, para.text, tagged_text)
                    else:
                        if para.runs:
                            para.runs[0].text = tagged_text
                        else:
                            para.add_run(tagged_text)


# ============== FastAPI 应用 ==============

app = FastAPI(
    title="资信标自动填充系统",
    description="基于 AI 的 Word 文档自动填充服务",
    version="3.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

try:
    app.mount("/static", StaticFiles(directory="static"), name="static")
except:
    pass


@app.get("/")
async def root():
    return FileResponse("index.html")


@app.post("/api/upload", response_model=ProcessResponse)
async def upload_file(file: UploadFile = File(...)):
    if not file.filename.endswith('.docx'):
        raise HTTPException(400, "只支持 .docx 格式的 Word 文件")

    task_id = str(uuid.uuid4())
    file_path = UPLOAD_DIR / f"{task_id}_{file.filename}"

    try:
        content = await file.read()
        if len(content) > MAX_FILE_SIZE:
            raise HTTPException(400, f"文件大小超过限制 ({MAX_FILE_SIZE / 1024 / 1024:.1f}MB)")

        async with aiofiles.open(file_path, 'wb') as f:
            await f.write(content)

        logger.info(f"文件上传成功: {task_id}")
        return ProcessResponse(
            success=True,
            task_id=task_id,
            steps=[
                ProcessStep(step=1, name="文件上传与解析", status="completed", time_ms=100)
            ]
        )
    except Exception as e:
        logger.error(f"文件上传失败: {e}")
        raise HTTPException(500, f"文件上传失败: {str(e)}")


@app.post("/api/process/{task_id}", response_model=ProcessResponse)
async def process_file(
    task_id: str,
    data: ProcessRequest = Body(...)
):
    filename = data.filename
    company_info = data.company_info

    input_path = UPLOAD_DIR / f"{task_id}_{filename}"
    output_filename = filename.replace('.docx', '_filled.docx')
    output_path = OUTPUT_DIR / f"{task_id}_{output_filename}"

    if not input_path.exists():
        raise HTTPException(404, "文件不存在，请先上传")

    step_logger = StepLogger(task_id)
    step_logger.log_step("开始处理", "running", {"filename": filename})

    try:
        doc = Document(str(input_path))
        step_logger.log_step("文档加载", "completed", {"paragraphs": len(doc.paragraphs), "tables": len(doc.tables)})

        def progress_callback(step, message):
            logger.info(f"[{task_id}] Step {step}: {message}")

        processed_doc, details = process_document(doc, company_info, step_logger, progress_callback)
        processed_doc.save(str(output_path))
        step_logger.log_step("保存结果", "completed")

        steps = []
        for step_detail in details["steps"]:
            steps.append(ProcessStep(
                step=step_detail["step"],
                name=step_detail["name"],
                status="completed",
                time_ms=step_detail.get("time_ms", 0)
            ))

        logger.info(f"处理完成: {task_id}")
        return ProcessResponse(
            success=True,
            task_id=task_id,
            steps=steps,
            tags=details["tags"],
            matched_data=details["matched_data"],
            download_url=f"/api/download/{task_id}/{output_filename}"
        )

    except Exception as e:
        logger.error(f"处理失败: {e}", exc_info=True)
        step_logger.log_step("处理失败", "error", {"error": str(e)})
        raise HTTPException(500, f"处理失败: {str(e)}")


@app.get("/api/download/{task_id}/{filename}")
async def download_file(task_id: str, filename: str):
    file_path = OUTPUT_DIR / f"{task_id}_{filename}"

    if not file_path.exists():
        raise HTTPException(404, "文件不存在")

    return FileResponse(
        str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@app.get("/api/logs/{task_id}")
async def get_logs(task_id: str):
    log_file = LOG_DIR / f"{task_id}_steps.json"
    if log_file.exists():
        with open(log_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"error": "日志不存在"}


@app.delete("/api/cleanup/{task_id}")
async def cleanup(task_id: str):
    for file in UPLOAD_DIR.glob(f"{task_id}_*"):
        file.unlink()
    for file in OUTPUT_DIR.glob(f"{task_id}_*"):
        file.unlink()
    return {"success": True}


if __name__ == "__main__":
    import uvicorn
    print("=" * 50)
    print("资信标自动填充系统 v3.0")
    print("=" * 50)
    print("服务地址: http://localhost:8000")
    print("=" * 50)
    uvicorn.run(app, host="0.0.0.0", port=8000)
