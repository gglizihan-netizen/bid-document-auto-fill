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
KIMI_MODEL = os.getenv("KIMI_MODEL", "moonshot-v1-128k")
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

def get_paragraph_text_with_underline(para) -> str:
    """
    提取段落文本，保留下划线格式

    规则：
    - 如果下划线文字是纯空白（空格、下划线字符），转换为下划线
    - 如果下划线文字包含括号提示如（招标人名称），保留原文，让模型理解空白含义
    - 不影响跨Run替换算法（替换时仍使用 para.text）
    """
    from docx.enum.text import WD_UNDERLINE

    result = []
    for run in para.runs:
        text = run.text
        if not text:
            continue

        # 检查是否有下划线
        if run.underline and run.underline != WD_UNDERLINE.NONE:
            # 判断是否为纯空白/下划线字符
            stripped = text.strip()
            # 如果是纯空白或纯下划线，转换为下划线
            if not stripped or stripped == '_' * len(stripped) or re.match(r'^[\s_]+$', stripped):
                result.append('_' * len(text))
            else:
                # 包含实际文字（如括号提示），保留原文
                result.append(text)
        else:
            result.append(text)

    return ''.join(result)


def replace_text_in_paragraph(para, old_text: str, new_text) -> bool:
    """
    跨 Run 替换文本，保留格式（包括下划线）
    核心原则：每次替换前重新读取段落文本，解决索引偏移问题
    修复：只给新填充的文字添加下划线，不影响其他文字
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

    # 检测原位置是否有下划线格式
    from docx.enum.text import WD_UNDERLINE
    has_underline = False
    for run_info in matching_runs:
        run = run_info['run']
        if run.underline and run.underline != WD_UNDERLINE.NONE:
            has_underline = True
            break

    if len(matching_runs) == 1:
        run_info = matching_runs[0]
        run = run_info['run']
        run_local_start = run_info['text_start'] - run_info['run_start']
        run_local_end = run_info['text_end'] - run_info['run_start']

        # 检查是否需要拆分 Run（替换只发生在部分位置）
        text_before = run.text[:run_local_start]
        text_after = run.text[run_local_end:]
        needs_split = bool(text_before or text_after)

        if needs_split:
            # 需要拆分：保留前后文字，只给新文字加下划线
            # 先保存原 Run 的格式信息
            original_underline = run.underline
            original_font = run.font

            # 清空原 Run，只保留前面的文字
            if text_before:
                run.text = text_before
                # 前面的文字保持原格式（去掉下划线）
                run.underline = WD_UNDERLINE.NONE
            else:
                run.text = ""

            # 创建新 Run 用于新文字（带下划线）
            new_run = para.add_run(new_text)
            if has_underline:
                new_run.underline = WD_UNDERLINE.SINGLE
            # 复制其他格式
            if original_font:
                try:
                    new_run.font.name = original_font.name
                    new_run.font.size = original_font.size
                    new_run.font.bold = original_font.bold
                except:
                    pass

            # 创建新 Run 用于后面的文字（不带下划线）
            if text_after:
                after_run = para.add_run(text_after)
                after_run.underline = WD_UNDERLINE.NONE
                if original_font:
                    try:
                        after_run.font.name = original_font.name
                        after_run.font.size = original_font.size
                        after_run.font.bold = original_font.bold
                    except:
                        pass
        else:
            # 整个 Run 都是被替换的内容，直接替换并保留下划线
            run.text = new_text
            if has_underline:
                run.underline = WD_UNDERLINE.SINGLE

        return True

    # 跨Run处理：保留首个Run格式
    first_run_info = matching_runs[0]
    first_run = first_run_info['run']
    first_run_local_start = first_run_info['text_start'] - first_run_info['run_start']

    # 检查首个 Run 是否有前置文字
    text_before_first = first_run.text[:first_run_local_start]

    if text_before_first:
        # 有前置文字，需要保留并去掉下划线
        first_run.text = text_before_first
        first_run.underline = WD_UNDERLINE.NONE

        # 创建新 Run 用于新文字（带下划线）
        new_run = para.add_run(new_text)
        if has_underline:
            new_run.underline = WD_UNDERLINE.SINGLE
    else:
        # 没有前置文字，直接替换
        first_run.text = new_text
        if has_underline:
            first_run.underline = WD_UNDERLINE.SINGLE

    # 处理后续的 Run：删除它们中包含的目标文本部分
    for i, run_info in enumerate(matching_runs):
        if i == 0:
            continue
        run = run_info['run']
        run_local_start = run_info['text_start'] - run_info['run_start']
        run_local_end = run_info['text_end'] - run_info['run_start']

        if run_local_end < len(run.text):
            # 这个 Run 除了目标文本外还有其他内容，保留后面的内容
            run.text = run.text[run_local_end:]
            # 后面的内容去掉下划线
            run.underline = WD_UNDERLINE.NONE
        else:
            # 整个 Run 都是目标文本的一部分，清空它
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
    logger.info(f"[表格上下文增强] 开始处理，窗口大小: {window_size}")

    upper_context_window = deque(maxlen=window_size)
    table_context_list = []

    table_count = 0
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if text:
                upper_context_window.append(text)

        elif isinstance(block, Table):
            table_count += 1
            table_upper_context = "\n".join(upper_context_window)

            logger.info(f"  [表格 {table_count}] 收集上文:")
            logger.info(f"    上文段落数: {len(upper_context_window)}")
            logger.info(f"    上文内容: {table_upper_context[:150]}...")

            serialized_table, unique_cells_info = serialize_table(block)

            table_context_list.append({
                "table": block,
                "upper_context": table_upper_context,
                "serialized_table": serialized_table,
                "unique_cells_info": unique_cells_info
            })

            # 关键：处理完表格清空窗口，避免污染下一个表格
            upper_context_window.clear()

    logger.info(f"[表格上下文增强] 完成，共处理 {len(table_context_list)} 个表格")
    return table_context_list


# ============== 表格序列化（合并单元格每行显示内容） ==============

def serialize_table(table) -> Tuple[str, List[Dict]]:
    """
    将Word表格序列化为 | 分隔符文本
    返回：(序列化文本, 唯一单元格信息列表)

    核心功能：
    1. 使用 | 分隔符伪Markdown（不包含分隔行）
    2. 空单元格统一替换为 [空]
    3. 合并单元格在每一行都显示其内容（让模型更好理解表格结构）
    4. 记录唯一单元格信息，用于后续回写

    序列化规则：
    - 每行的列数与原始表格一致
    - 合并单元格的内容在每一行都显示
    - 这样模型能理解表格的完整结构
    """
    logger.info(f"    [表格序列化] 开始处理表格: {len(table.rows)}行 x {len(table.columns)}列")

    rows_text = []
    unique_cells_info = []  # 存储每个唯一单元格的引用信息
    seen_cell_ids = set()  # 用于标记已记录的唯一单元格

    empty_cell_count = 0
    merged_cell_count = 0
    content_cell_count = 0

    for row_idx, row in enumerate(table.rows):
        cells_text = []
        row_cells_info = []

        for cell_idx, cell in enumerate(row.cells):
            # 使用内存地址作为唯一标识
            cell_id = id(cell._tc)
            text = cell.text.strip()

            # 判断是否为唯一单元格（第一次遇到）
            is_unique = cell_id not in seen_cell_ids
            if is_unique:
                seen_cell_ids.add(cell_id)

            # 处理空白单元格
            if is_blank_cell(text):
                cells_text.append('[空]')
                empty_cell_count += 1
                row_cells_info.append({
                    'row_idx': row_idx,
                    'cell_idx': cell_idx,
                    'cell': cell,
                    'original_text': '',
                    'is_empty': True,
                    'is_unique': is_unique
                })
            else:
                # 转义分隔符
                text_escaped = text.replace('|', '｜').replace('\n', ' ')
                cells_text.append(text_escaped)

                if is_unique:
                    content_cell_count += 1
                else:
                    merged_cell_count += 1

                row_cells_info.append({
                    'row_idx': row_idx,
                    'cell_idx': cell_idx,
                    'cell': cell,
                    'original_text': text_escaped,
                    'is_empty': False,
                    'is_unique': is_unique
                })

        rows_text.append('| ' + ' | '.join(cells_text) + ' |')
        unique_cells_info.append(row_cells_info)

    result = '\n'.join(rows_text)

    logger.info(f"    [表格序列化] 完成:")
    logger.info(f"      - 空单元格: {empty_cell_count}")
    logger.info(f"      - 唯一内容单元格: {content_cell_count}")
    logger.info(f"      - 合并单元格(重复显示): {merged_cell_count}")
    logger.info(f"      - 序列化行数: {len(rows_text)}")
    logger.info(f"      - 序列化长度: {len(result)} 字符")
    logger.info(f"      - 序列化预览: {result[:300]}...")

    return result, unique_cells_info


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

    动态列表特征：
    - 顶部有表头行（非空单元格）
    - 后续有多行连续空白行（供用户填写数据）
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
        # 统计空单元格数量（[空] 表示空白单元格）
        empty_count = sum(1 for cell in row if cell == '[空]')
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

def call_kimi_api(prompt: str, system_prompt: str, max_tokens: int = 2048, timeout: int = 120) -> str:
    """通用Kimi API调用 - 增加超时和详细日志"""
    try:
        start_time = time.time()

        # 记录请求信息
        logger.info(f"=== API请求开始 ===")
        logger.info(f"模型: {KIMI_MODEL}")
        logger.info(f"System Prompt: {system_prompt[:100]}...")
        logger.info(f"User Prompt长度: {len(prompt)} 字符")
        logger.info(f"Max Tokens: {max_tokens}, Timeout: {timeout}秒")

        # 使用带超时的客户端
        from openai import OpenAI
        timeout_client = OpenAI(
            api_key=KIMI_API_KEY,
            base_url=KIMI_BASE_URL,
            timeout=timeout
        )

        response = timeout_client.chat.completions.create(
            model=KIMI_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ],
            max_tokens=max_tokens,
            temperature=0.1  # moonshot模型支持temperature
        )

        elapsed = time.time() - start_time
        result = response.choices[0].message.content.strip()

        # 记录响应信息
        logger.info(f"=== API响应成功 ===")
        logger.info(f"耗时: {elapsed:.2f}秒")
        logger.info(f"响应长度: {len(result)} 字符")
        logger.info(f"响应预览: {result[:200]}...")
        logger.info(f"使用Tokens: {response.usage.total_tokens if response.usage else 'N/A'}")

        return result
    except Exception as e:
        elapsed = time.time() - start_time
        logger.error(f"=== API调用失败 ===")
        logger.error(f"耗时: {elapsed:.2f}秒")
        logger.error(f"错误类型: {type(e).__name__}")
        logger.error(f"错误详情: {e}")
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
    """调用Kimi API为段落块打标 - 严格版"""
    prompt = f"""【角色】你是投标文件空白处标签生成器。

【核心任务】
找出文本中的所有空白处（下划线、空格），用{{{{标签名}}}}替换整个空白区域。

【绝对禁止】
1. 禁止在标签后面保留下划线字符"_"
2. 禁止在标签后面保留空格
3. 空白处必须100%被标签替换，不能有任何残留

【空白处识别】
- 连续下划线：________（整段都要替换）
- 连续空格：          （整段都要替换）
- 括号提示：（项目名称）

【正确示例】
原文：联系人：____________________________
正确：联系人：{{{{联系人姓名}}}}
错误：联系人：{{{{联系人姓名}}}}____________________________

原文：电话：______________________________
正确：电话：{{{{联系电话}}}}
错误：电话：{{{{联系电话}}}}______________________________

原文：招标编号为__________的项目
正确：招标编号为{{{{招标编号}}}}的项目
错误：招标编号为{{{{招标编号}}}}__________的项目

【待处理文本】
{text_block}

【输出要求】
只输出打标后的文本，保留[数字]索引。空白处必须完全替换，不能有任何下划线或空格残留！"""

    system_prompt = "你是空白处标签生成器。必须完全替换空白处，绝对禁止保留下划线或空格。"
    return call_kimi_api(prompt, system_prompt, max_tokens=4096)


def call_kimi_tag_table(serialized_table: str, upper_context: str) -> str:
    """调用Kimi API为表格打标 - 统一版本

    按照技术文档设计：由模型自行区分静态表单和动态列表，不由后端判断
    """
    prompt = f"""【角色】标书模板表格智能打标专家

【任务】接收按|分隔符序列化的表格文本，自主判断表格类型：静态表单 / 动态列表。参考【表格前文背景】，为表格打标，仅处理表格内容。

【判断规则】
1. 静态表单：人员信息、企业资质、固定填报项，特征为属性名+单一项填空；
2. 动态列表：业绩清单、人员名录、设备清单，特征为顶部统一表头、下方多行重复空白行。

【输入】
=== 表格前文背景（仅参考，禁止修改、禁止输出） ===
{upper_context}
=== 待打标表格（|分隔序列化） ===
{serialized_table}

【打标规则】
1. 静态表单：所有[空]逐一替换为独立业务标签{{{{字段名}}}}，保留所有行列结构；
2. 动态列表：仅保留表头和第一行空白位，删除后续所有重复空行；
   在首行添加docxtpl循环语法：{{% tr for item in 列表名称 %}}
   单元格使用泛型标签：{{{{item.字段名}}}}。
3. 绝对禁止生成如 {{{{地址}}}}、{{{{姓名}}}} 这样指代不明的标签！你必须结合前文背景（如标题、业务场景），生成高度具体的标签（如 {{{{投标人单位地址}}}}、{{{{项目负责人姓名}}}}）。

【示例】
输入：
=== 表格前文背景（仅参考，禁止修改、禁止输出） ===
历史业绩表格
=== 待打标表格（|分隔序列化） ===
| 项目名称 | 建设单位 | 合同金额 |
| [空] | [空] | [空] |
| [空] | [空] | [空] |

输出：
| 项目名称 | 建设单位 | 合同金额 |
| {{% tr for item in 历史业绩列表 %}}{{{{item.项目名称}}}} | {{{{item.建设单位}}}} | {{{{item.合同金额}}}} |

【输出要求】
1. 严格保留原有|分隔符，仅替换[空]、增减空行，不改动原有文字和顺序。
2. 只输出打标后的表格文本，禁止输出前文、解释、代码块！"""

    system_prompt = "你是表格打标工具，只输出打标后的表格文本。标签必须高度具体，禁止使用泛标签。"
    return call_kimi_api(prompt, system_prompt, max_tokens=4096)


def call_kimi_match_info(tags: List[str], company_info: str) -> Dict[str, Any]:
    """调用Kimi API进行信息匹配（支持列表类型）- 优化版"""
    if not tags:
        return {}

    batch_size = 60
    all_results = {}

    for batch_start in range(0, len(tags), batch_size):
        batch_tags = tags[batch_start:batch_start + batch_size]

        prompt = f"""【角色】你是一名专业的企业信息匹配专家，精通投标文件信息提取。

【输入】
1. 标签列表:{json.dumps(batch_tags, ensure_ascii=False, indent=2)}
2. 企业基础信息文本

【企业信息】：
{company_info[:5000]}

【任务】请在"企业基础信息文本"中，为"标签列表"里的每一个标签寻找最匹配的具体信息。

【匹配规则 - 必须严格遵守】
1. 标签名称包含完整语义，需要精确匹配：
   - "投标人单位名称" → 找"企业名称"对应的值
   - "法定代表人姓名" → 找"法定代表人"下的"姓名"值
   - "项目负责人姓名" → 找"项目负责人"下的"姓名"值
   - "联系人电话" → 找"授权委托代理人"下的"联系电话"值
   - "投标日期年/月/日" → 找"投标日期"对应的年/月/日

2. 对于日期类标签，提取对应日期的数字部分：
   - "投标日期年" → "2026"
   - "投标日期月" → "05"
   - "投标日期日" → "09"

3. 对于签字盖章类标签，返回"[需签字]"或"[需盖章]"

4. 如果企业信息中找不到该标签对应的内容，Value 请返回"待补全"

5. 不要编造信息，必须基于提供的企业信息文本

【特殊说明：列表类型】
当你看到标签包含"列表"、"业绩"、"人员"等关键字时，说明这是一组列表数据。
请在企业信息中寻找所有符合条件的项目，提取完整的列表，输出成 JSON 数组形式。

【输出格式】
直接返回JSON对象，不要输出任何说明或代码块标记：
{{
  "投标人单位名称": "江苏华宇市政建设集团有限公司",
  "法定代表人姓名": "陈铭宇",
  "项目负责人姓名": "周建峰",
  "投标日期年": "2026",
  "投标日期月": "05",
  "投标日期日": "09"
}}"""

        system_prompt = "你是数据匹配工具，只输出JSON，不要输出代码块标记。"

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
    """处理文档的主流程 - 增强日志版"""

    # Step 1: 合并碎片化 Run
    logger.info("=" * 60)
    logger.info("【Step 1】文档预处理 - 合并碎片化Run")
    logger.info("=" * 60)

    step_logger.log_step("文档预处理", "running")
    if progress_callback:
        progress_callback(1, "合并文档格式...")

    start_time = time.time()

    # 记录处理前的Run数量
    total_runs_before = sum(len(para.runs) for para in doc.paragraphs)
    logger.info(f"处理前总Run数: {total_runs_before}")

    merged_count = merge_adjacent_runs(doc)

    total_runs_after = sum(len(para.runs) for para in doc.paragraphs)
    logger.info(f"处理后总Run数: {total_runs_after}")
    logger.info(f"合并Run数: {merged_count}")

    elapsed = int((time.time() - start_time) * 1000)
    step_logger.log_step("文档预处理", "completed", {"time_ms": elapsed, "merged_runs": merged_count})

    # Step 2: AI 辅助打标 - 正文段落（索引式语义分块）
    logger.info("")
    logger.info("=" * 60)
    logger.info("【Step 2】AI辅助打标 - 正文段落")
    logger.info("=" * 60)

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

    logger.info(f"有效段落数: {len(valid_paragraphs)}")

    # 2.2 调用AI进行语义分块（使用带下划线的文本）
    numbered_lines = [f"[{idx}] {get_paragraph_text_with_underline(p).strip()}" for idx, p in enumerate(valid_paragraphs)]
    numbered_text = "\n".join(numbered_lines[:100])  # 限制长度

    logger.info(f"语义分块输入文本长度: {len(numbered_text)} 字符")
    logger.info(f"语义分块输入预览: {numbered_text[:300]}...")

    try:
        block_indices = call_kimi_semantic_chunk(numbered_text, len(valid_paragraphs))
        logger.info(f"语义分块结果: {block_indices}")
        logger.info(f"分块数量: {len(block_indices) - 1}")
    except Exception as e:
        logger.warning(f"语义分块失败，使用兜底方案: {e}")
        block_indices = fallback_chunk_indices(len(valid_paragraphs))
        logger.info(f"使用兜底分块: {block_indices}")

    # 2.3 按块打标
    paragraph_processed = 0
    all_tags_from_paragraphs = []

    for i in range(len(block_indices) - 1):
        start_idx = block_indices[i]
        end_idx = block_indices[i + 1]

        # 提取块内段落
        block_paras = valid_paragraphs[start_idx:end_idx]
        if not block_paras:
            continue

        block_text = "\n".join([f"[{start_idx + j}] {get_paragraph_text_with_underline(p).strip()}" for j, p in enumerate(block_paras)])

        logger.info("")
        logger.info(f"--- 处理段落块 {i + 1}/{len(block_indices) - 1} ---")
        logger.info(f"段落范围: [{start_idx}] - [{end_idx - 1}]")
        logger.info(f"块内段落数: {len(block_paras)}")
        logger.info(f"块文本长度: {len(block_text)} 字符")
        logger.info(f"块文本预览: {block_text[:200]}...")

        # 调用AI打标
        try:
            logger.info(f"开始调用AI打标...")
            tagged_text = call_kimi_tag_paragraphs(block_text)

            logger.info(f"AI打标结果长度: {len(tagged_text)} 字符")
            logger.info(f"AI打标结果预览: {tagged_text[:300]}...")

            tagged_map = parse_tagged_paragraphs(tagged_text)
            logger.info(f"解析出的段落映射数: {len(tagged_map)}")

            # 应用打标结果
            applied_count = 0
            for idx_offset, para in enumerate(block_paras):
                global_idx = start_idx + idx_offset
                if global_idx in tagged_map:
                    new_text = tagged_map[global_idx]
                    if "{{" in new_text and new_text != para.text:
                        old_text = para.text[:50]
                        replace_text_in_paragraph(para, para.text, new_text)
                        logger.info(f"  段落[{global_idx}] 已打标: '{old_text}...' -> '{new_text[:50]}...'")
                        paragraph_processed += 1
                        applied_count += 1

                        # 提取标签
                        tags_in_para = re.findall(r'\{\{([^}]+)\}\}', new_text)
                        all_tags_from_paragraphs.extend(tags_in_para)

            logger.info(f"块内应用打标数: {applied_count}")
            step_logger.log_api_call(f"paragraph_block_{i}", block_text[:300], tagged_text[:300])

        except Exception as e:
            logger.error(f"段落块 {i} 打标失败: {e}")
            import traceback
            logger.error(traceback.format_exc())

    logger.info("")
    logger.info(f"段落打标完成统计:")
    logger.info(f"  - 处理段落块数: {len(block_indices) - 1}")
    logger.info(f"  - 成功打标段落数: {paragraph_processed}")
    logger.info(f"  - 提取标签数: {len(all_tags_from_paragraphs)}")
    logger.info(f"  - 标签列表: {all_tags_from_paragraphs[:20]}...")

    elapsed = int((time.time() - start_time) * 1000)
    step_logger.log_step("AI辅助打标-段落", "completed", {"time_ms": elapsed, "processed": paragraph_processed})

    # Step 3: AI 辅助打标 - 表格（上下文增强 + 唯一单元格）
    logger.info("")
    logger.info("=" * 60)
    logger.info("【Step 3】AI辅助打标 - 表格")
    logger.info("=" * 60)

    step_logger.log_step("AI辅助打标-表格", "running")
    if progress_callback:
        progress_callback(3, "处理表格...")

    start_time = time.time()
    table_processed = 0
    all_tags_from_tables = []

    # 获取带上下文的表格
    tables_with_context = get_tables_with_upper_context(doc)
    logger.info(f"发现表格数: {len(tables_with_context)}")

    for table_idx, table_ctx in enumerate(tables_with_context[:10]):
        table = table_ctx['table']
        upper_context = table_ctx['upper_context']
        serialized_table = table_ctx['serialized_table']
        unique_cells_info = table_ctx['unique_cells_info']

        logger.info("")
        logger.info(f"--- 处理表格 {table_idx + 1}/{len(tables_with_context)} ---")
        logger.info(f"表格尺寸: {len(table.rows)} 行 x {len(table.columns)} 列")
        logger.info(f"表格上文: {upper_context[:100]}...")
        logger.info(f"序列化表格长度: {len(serialized_table)} 字符")
        logger.info(f"序列化表格预览: {serialized_table[:300]}...")

        try:
            logger.info(f"开始调用AI表格打标...")
            tagged_table = call_kimi_tag_table(serialized_table, upper_context)

            logger.info(f"AI表格打标结果长度: {len(tagged_table)} 字符")
            logger.info(f"AI表格打标结果: {tagged_table[:300]}...")

            step_logger.log_api_call(f"table_{table_idx}", serialized_table[:500], tagged_table[:500])

            # 解析打标结果
            tagged_rows = parse_serialized_table(tagged_table)
            logger.info(f"解析出行数: {len(tagged_rows)}")

            # 提取标签
            for row in tagged_rows:
                for cell in row:
                    tags_in_cell = re.findall(r'\{\{([^}]+)\}\}', cell)
                    all_tags_from_tables.extend(tags_in_cell)

            # 应用回表格
            apply_table_tags(table, tagged_rows, unique_cells_info)

            table_processed += 1
            logger.info(f"表格 {table_idx} 处理完成")
            logger.info(f"表格标签: {all_tags_from_tables[-10:] if all_tags_from_tables else '无'}")

        except Exception as e:
            logger.error(f"处理表格 {table_idx} 失败: {e}")
            import traceback
            logger.error(traceback.format_exc())
            step_logger.log_step(f"表格_{table_idx}_错误", "error", {"error": str(e)})

    logger.info("")
    logger.info(f"表格打标完成统计:")
    logger.info(f"  - 处理表格数: {table_processed}")
    logger.info(f"  - 提取标签数: {len(all_tags_from_tables)}")

    elapsed = int((time.time() - start_time) * 1000)
    step_logger.log_step("AI辅助打标-表格", "completed", {"time_ms": elapsed, "processed": table_processed})

    # Step 4: 标签提取
    logger.info("")
    logger.info("=" * 60)
    logger.info("【Step 4】标签提取")
    logger.info("=" * 60)

    step_logger.log_step("标签提取", "running")
    if progress_callback:
        progress_callback(4, "提取标签...")

    start_time = time.time()
    tags = extract_tags_from_doc(doc)

    logger.info(f"提取标签总数: {len(tags)}")
    logger.info(f"标签列表: {tags[:30]}...")

    elapsed = int((time.time() - start_time) * 1000)
    step_logger.log_step("标签提取", "completed", {"time_ms": elapsed, "count": len(tags), "tags": tags})

    # Step 5: 企业信息匹配
    logger.info("")
    logger.info("=" * 60)
    logger.info("【Step 5】企业信息匹配")
    logger.info("=" * 60)

    step_logger.log_step("企业信息匹配", "running")
    if progress_callback:
        progress_callback(5, "匹配企业信息...")

    start_time = time.time()

    logger.info(f"待匹配标签数: {len(tags)}")
    logger.info(f"企业信息长度: {len(company_info)} 字符")
    logger.info(f"企业信息预览: {company_info[:300]}...")

    matched_data = call_kimi_match_info(tags, company_info)

    # 统计匹配结果
    matched_count = sum(1 for v in matched_data.values() if v and v != "待补全")
    unmatched_count = len(matched_data) - matched_count

    logger.info(f"匹配完成统计:")
    logger.info(f"  - 成功匹配: {matched_count}")
    logger.info(f"  - 待补全: {unmatched_count}")
    logger.info(f"  - 匹配率: {matched_count * 100 // len(matched_data) if matched_data else 0}%")

    # 显示匹配结果详情
    logger.info("匹配结果详情:")
    for tag, value in list(matched_data.items())[:20]:
        logger.info(f"  {tag}: {value}")

    elapsed = int((time.time() - start_time) * 1000)
    step_logger.log_step("企业信息匹配", "completed", {"time_ms": elapsed, "matched_count": len(matched_data)})

    # Step 6: 内容填充
    logger.info("")
    logger.info("=" * 60)
    logger.info("【Step 6】内容填充")
    logger.info("=" * 60)

    step_logger.log_step("内容填充", "running")
    if progress_callback:
        progress_callback(6, "填充内容...")

    start_time = time.time()
    replacement_count = 0

    # 处理正文
    logger.info("填充正文段落...")
    para_replacement_count = 0
    for para in doc.paragraphs:
        for tag, value in matched_data.items():
            placeholder = f"{{{{{tag}}}}}"
            if placeholder in para.text:
                old_text = para.text[:50]
                if replace_text_in_paragraph(para, placeholder, value):
                    new_text = para.text[:50]
                    logger.info(f"  替换: '{old_text}...' -> '{new_text}...'")
                    replacement_count += 1
                    para_replacement_count += 1

    logger.info(f"正文填充数: {para_replacement_count}")

    # 处理表格
    logger.info("填充表格...")
    table_replacement_count = 0
    for table_idx, table in enumerate(doc.tables):
        table_count = 0
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for tag, value in matched_data.items():
                        placeholder = f"{{{{{tag}}}}}"
                        if placeholder in para.text:
                            if replace_text_in_paragraph(para, placeholder, value):
                                replacement_count += 1
                                table_replacement_count += 1
                                table_count += 1
        if table_count > 0:
            logger.info(f"  表格 {table_idx} 填充数: {table_count}")

    logger.info(f"表格填充数: {table_replacement_count}")
    logger.info(f"总填充数: {replacement_count}")

    elapsed = int((time.time() - start_time) * 1000)
    step_logger.log_step("内容填充", "completed", {"time_ms": elapsed, "count": replacement_count})

    logger.info("")
    logger.info("=" * 60)
    logger.info("【处理完成】最终统计")
    logger.info("=" * 60)
    logger.info(f"合并Run数: {merged_count}")
    logger.info(f"段落打标数: {paragraph_processed}")
    logger.info(f"表格打标数: {table_processed}")
    logger.info(f"标签总数: {len(tags)}")
    logger.info(f"匹配成功数: {matched_count}")
    logger.info(f"填充总数: {replacement_count}")

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


def apply_table_tags(table: Table, tagged_rows: List[List[str]], unique_cells_info: List):
    """将打标结果应用回Word表格

    说明：
    - 模型自行判断表格类型并输出打标结果
    - 序列化时合并单元格在每行都显示内容，所以列数一致
    - 只更新唯一单元格（is_unique=True），合并单元格不重复更新
    """

    for row_idx, row in enumerate(table.rows):
        if row_idx >= len(tagged_rows):
            break

        tagged_cells = tagged_rows[row_idx]
        cells_info = unique_cells_info[row_idx] if row_idx < len(unique_cells_info) else []

        for cell_idx, cell in enumerate(row.cells):
            if cell_idx >= len(tagged_cells):
                break

            cell_info = cells_info[cell_idx] if cell_idx < len(cells_info) else {}

            # 只更新唯一单元格，合并单元格不重复更新
            if not cell_info.get('is_unique', True):
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
