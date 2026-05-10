"""
资信标自动填充系统 - 后端服务
基于 FastAPI + python-docx + Kimi API 实现
采用 Markdown 中间语言处理表格
"""

import os
import re
import json
import time
import uuid
import logging
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional

from fastapi import FastAPI, File, UploadFile, Form, HTTPException, Body, Request
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from dotenv import load_dotenv
from docx import Document
from openai import OpenAI
import aiofiles

# 加载环境变量
load_dotenv()

# 配置
KIMI_API_KEY = os.getenv("KIMI_API_KEY", "")
KIMI_BASE_URL = os.getenv("KIMI_BASE_URL", "https://api.moonshot.cn/v1")
KIMI_MODEL = os.getenv("KIMI_MODEL", "moonshot-v1-8k")
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
    matched_data: Optional[Dict[str, Any]] = None  # 改为 Any，允许任意类型
    download_url: Optional[str] = None
    error: Optional[str] = None


class ProcessRequest(BaseModel):
    filename: str
    company_info: str


# ============== 核心算法 ==============

def replace_text_in_paragraph(para, old_text: str, new_text) -> bool:
    """跨 Run 替换文本，保留格式"""
    if not old_text:
        return False

    # 确保 new_text 是字符串
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

    new_run_text = first_run.text[:first_run_local_start] + new_text
    first_run.text = new_run_text

    for i, run_info in enumerate(matching_runs):
        if i == 0:
            continue
        run = run_info['run']
        run_local_start = run_info['text_start'] - run_info['run_start']
        run_local_end = run_info['text_end'] - run_info['run_start']

        if run_local_end < len(run.text):
            run.text = run.text[run_local_end:]
        else:
            run.text = ""

    return True


def merge_adjacent_runs(doc: Document) -> None:
    """合并格式相同的相邻 Run"""
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

    for para in doc.paragraphs:
        if len(para.runs) <= 1:
            continue
        i = 0
        while i < len(para.runs) - 1:
            current_run = para.runs[i]
            next_run = para.runs[i + 1]
            if get_run_format(current_run) == get_run_format(next_run):
                current_run.text += next_run.text
                next_run._element.getparent().remove(next_run._element)
            else:
                i += 1

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if len(para.runs) <= 1:
                        continue
                    i = 0
                    while i < len(para.runs) - 1:
                        current_run = para.runs[i]
                        next_run = para.runs[i + 1]
                        if get_run_format(current_run) == get_run_format(next_run):
                            current_run.text += next_run.text
                            next_run._element.getparent().remove(next_run._element)
                        else:
                            i += 1


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
            "prompt_preview": prompt[:300] + "..." if len(prompt) > 300 else prompt,
            "response_preview": response[:300] + "..." if response and len(response) > 300 else response,
            "error": error
        }
        api_log_file = LOG_DIR / f"{self.task_id}_api_calls.jsonl"
        with open(api_log_file, 'a', encoding='utf-8') as f:
            f.write(json.dumps(api_log, ensure_ascii=False) + '\n')

    def _save(self):
        with open(self.log_file, 'w', encoding='utf-8') as f:
            json.dump(self.logs, f, ensure_ascii=False, indent=2)


# ============== Kimi API 调用 ==============

def call_kimi_tagging(text: str, is_table: bool = False) -> str:
    """调用 Kimi API 进行打标"""
    if is_table:
        prompt = f"""【角色】你是专业的投标文件标签生成器。

【任务】为Markdown表格中的空白单元格打标签。

【输入表格】：
{text}

【打标规则】
1. 只在 [空] 或空白单元格处填入 {{{{标签名}}}}
2. 标签名根据该列表头语义生成，简洁明确
3. 如果同一列有多个空白，使用相同标签（如多行重复）
4. 保持表格格式完整，包括 | 分隔符和 |---|---| 分隔行
5. 非空白单元格保持原样不变

【标签命名规范】
- 人员信息：姓名、职务、职称、学历、年龄、性别、身份证号
- 联系方式：联系电话、传真、邮编
- 企业信息：单位名称、注册资金、成立时间、注册地址
- 项目信息：工程名称、建设单位、建设规模、开竣工日期
- 时间日期：年、月、日（如"  年  月  日" → "{{{{年}}}}年{{{{月}}}}月{{{{日}}}}日"）

【示例】
输入：
| 姓名 |  | 性别 |  | 年龄 |  |
|---|---|---|---|---|---|
| 张三 |  | 男 |  |  |  |

输出：
| 姓名 | {{{{姓名}}}} | 性别 | {{{{性别}}}} | 年龄 | {{{{年龄}}}} |
|---|---|---|---|---|---|
| 张三 | {{{{姓名}}}} | 男 | {{{{性别}}}} |  | {{{{年龄}}}} |

【重要】只输出打标后的Markdown表格，不要输出任何说明文字！"""
    else:
        prompt = f"""【角色】你是专业的投标文件标签生成器。

【任务】识别文本中的空白处并打上标签。

【输入文本】：
{text}

【空白处识别】
1. 下划线：________、____、_等
2. 多个空格：      （连续空格）
3. 冒号后留白：投标人：  （冒号后无内容）
4. 括号提示：(项目名称)、[请填写]、（单位盖章）等
5. 日期格式空白：  年  月  日

【打标规则】
1. 根据上下文语义生成简洁的中文标签名
2. 标签格式：{{{{标签名}}}}
3. 必须原样返回整句，只替换空白处为标签
4. 严禁删减、增加或改动任何原有字符
5. 如果找不到空白处，原样返回输入文本

【标签命名规范】
- 企业相关：单位名称、注册资金、成立时间、注册地址、法定代表人姓名
- 项目相关：工程名称、招标人名称、招标编号、项目负责人
- 联系方式：联系电话、传真、邮编、联系人姓名
- 日期时间：年、月、日（分别打标）
- 签字盖章：签字或盖章、单位盖章、法定代表人签字

【示例】
输入："投标人：        （单位盖章）"
输出："投标人：{{{{单位名称}}}}（单位盖章）"

输入："  年    月    日"
输出："{{{{年}}}}年{{{{月}}}}月{{{{日}}}}日"

输入："法定代表人（签字）：____"
输出："法定代表人（签字）：{{{{法定代表人签字}}}}"

【重要】只输出打标后的文本，不要输出任何说明！"""

    try:
        start_time = time.time()
        response = client.chat.completions.create(
            model=KIMI_MODEL,
            messages=[
                {"role": "system", "content": "你是标签生成器，只输出打标后的内容，不输出任何解释、说明或示例。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=2048
        )
        elapsed = time.time() - start_time
        result = response.choices[0].message.content.strip()

        # 清理AI返回的多余内容
        result = re.sub(r'\n*【空白处识别】.*', '', result, flags=re.DOTALL)
        result = re.sub(r'\n*【打标规则】.*', '', result, flags=re.DOTALL)
        result = re.sub(r'\n*【标签命名规范】.*', '', result, flags=re.DOTALL)
        result = re.sub(r'\n*【示例】.*', '', result, flags=re.DOTALL)
        result = re.sub(r'\n*【重要】.*', '', result, flags=re.DOTALL)
        result = re.sub(r'\n*【输入.*', '', result, flags=re.DOTALL)
        result = re.sub(r'\n*【输出.*', '', result, flags=re.DOTALL)

        # 移除代码块标记
        result = re.sub(r'^```markdown\s*', '', result)
        result = re.sub(r'^```\s*', '', result)
        result = re.sub(r'```\s*$', '', result)

        logger.info(f"API调用成功，耗时{elapsed:.2f}秒")
        return result.strip()
    except Exception as e:
        logger.error(f"API调用失败: {e}")
        return text


def call_kimi_match_info(tags: List[str], company_info: str) -> Dict[str, str]:
    """调用 Kimi API 进行信息匹配"""
    if not tags:
        return {}

    # 分批处理标签，每批最多60个
    batch_size = 60
    all_results = {}

    for batch_start in range(0, len(tags), batch_size):
        batch_tags = tags[batch_start:batch_start + batch_size]
        batch_num = batch_start // batch_size + 1
        total_batches = (len(tags) + batch_size - 1) // batch_size
        logger.info(f"处理标签批次 {batch_num}/{total_batches}，共 {len(batch_tags)} 个标签")

        prompt = f"""【角色】你是一名专业的企业信息匹配专家，擅长从企业信息中精确提取数据。

【重要规则】
1. 输出的JSON必须包含【标签列表】中的所有标签作为key
2. 找不到对应信息的标签，value填"待补全"
3. 不要遗漏任何标签，不要添加额外标签

【标签列表】（共{len(batch_tags)}个，必须全部出现在输出JSON的key中）：
{json.dumps(batch_tags, ensure_ascii=False, indent=2)}

【企业信息】：
{company_info[:3000]}

【字段名映射指南】（帮助理解标签含义）：
- "投标单位名称"="企业名称"="单位名称"="投标人名称"
- "邮编"="邮政编码"
- "联系地址"="注册地址"="办公地址"
- "注册资金"="注册资本"
- "法定代表人姓名"="法人姓名"
- "联系人姓名"="联系人"
- "联系电话"="电话"="固定电话"
- "身份证号"="身份证号码"
- "从事项目负责人年限"="项目经理年限"
- "从事技术负责人年限"="技术负责人年限"

【特殊标签处理】：
- 包含"签字"、"盖章"的标签 → value填"[需手动签字/盖章]"
- 包含"日期"、"时间"且无具体值的 → 从企业信息中推断合理日期

【输出要求】：
只输出一个JSON对象，格式如下：
{{"标签1": "对应值1", "标签2": "对应值2", ...}}

请确保输出的JSON中包含所有{len(batch_tags)}个标签！"""

        try:
            response = client.chat.completions.create(
                model=KIMI_MODEL,
                messages=[
                    {"role": "system", "content": "你是一个精确的数据匹配工具，只输出JSON，不要输出任何解释。输出必须包含所有输入标签作为key。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                max_tokens=4096
            )

            content = response.choices[0].message.content.strip()

            # 清理可能的markdown代码块
            content = re.sub(r'^```json\s*', '', content)
            content = re.sub(r'^```\s*', '', content)
            content = re.sub(r'```\s*$', '', content)

            json_match = re.search(r'\{[\s\S]*\}', content)
            if json_match:
                content = json_match.group()

            result = json.loads(content)

            # 确保所有标签都在结果中
            for tag in batch_tags:
                if tag not in result:
                    result[tag] = "待补全"
                    logger.warning(f"标签 '{tag}' 未在AI返回结果中，已设为待补全")

            all_results.update(result)
            logger.info(f"批次 {batch_num} 匹配完成，成功匹配 {sum(1 for v in result.values() if v != '待补全')} 个标签")

        except json.JSONDecodeError as e:
            logger.error(f"JSON解析失败: {e}, content: {content[:300]}")
            for tag in batch_tags:
                all_results[tag] = "待补全"
        except Exception as e:
            logger.error(f"信息匹配API调用失败: {e}")
            for tag in batch_tags:
                all_results[tag] = "待补全"

    # 后处理：处理签字盖章类标签
    for tag in all_results:
        if '签字' in tag or '盖章' in tag:
            if all_results[tag] == "待补全":
                all_results[tag] = f"[需{tag}]"

    return all_results


# ============== Markdown 表格处理 ==============

def table_to_markdown(table) -> str:
    """将Word表格转换为Markdown格式"""
    rows = []
    for row in table.rows:
        cells = []
        for cell in row.cells:
            text = cell.text.strip()
            # 将空内容或占位符替换为[空]便于AI识别
            if not text or text in ['待补全', '待填写', '[空]']:
                cells.append('[空]')
            else:
                # 转义Markdown特殊字符
                text = text.replace('|', '｜').replace('\n', ' ')
                cells.append(text)
        rows.append('| ' + ' | '.join(cells) + ' |')

    # 添加表头分隔行（如果表格有内容）
    if rows:
        col_count = len(table.rows[0].cells) if table.rows else 0
        separator = '|' + '|'.join(['---'] * col_count) + '|'
        rows.insert(1, separator)

    return '\n'.join(rows)


def parse_markdown_table(md_text: str) -> List[List[str]]:
    """解析Markdown表格为二维数组"""
    lines = [line.strip() for line in md_text.strip().split('\n') if line.strip()]
    rows = []
    for line in lines:
        # 跳过分隔行
        if '|' in line and '---' in line:
            continue
        # 解析单元格
        cells = [cell.strip() for cell in line.split('|')]
        # 过滤空单元格（首尾可能是空的）
        cells = [cell for cell in cells if cell]
        if cells:
            rows.append(cells)
    return rows


def apply_markdown_tags_to_table(table, md_rows: List[List[str]], step_logger: StepLogger = None):
    """将Markdown打标结果应用回Word表格"""
    # 跳过表头，从数据行开始（md_rows已经去掉了分隔行）
    for row_idx, row in enumerate(table.rows):
        if row_idx >= len(md_rows):
            break

        md_cells = md_rows[row_idx]
        for cell_idx, cell in enumerate(row.cells):
            if cell_idx >= len(md_cells):
                break

            md_text = md_cells[cell_idx]
            original_text = cell.text.strip()

            # 检查是否有标签
            if '{{' in md_text and '}}' in md_text:
                # 需要替换
                for para in cell.paragraphs:
                    if para.text.strip():
                        replace_text_in_paragraph(para, para.text, md_text)
                    else:
                        # 空段落，添加run
                        if para.runs:
                            para.runs[0].text = md_text
                        else:
                            para.add_run(md_text)

                if step_logger:
                    step_logger.log_api_call(
                        f"table_cell_r{row_idx}_c{cell_idx}",
                        f"original: '{original_text}'",
                        md_text
                    )


def is_blank_cell(text: str) -> bool:
    """判断单元格是否为空白"""
    if not text or not text.strip():
        return True

    blank_patterns = [
        r'^_{2,}$',
        r'^\s*$',
        r'^\[空\]$',
        r'^待填写$',
        r'^待补全$',
        r'^□+$',
    ]

    text_clean = text.strip()
    for pattern in blank_patterns:
        if re.match(pattern, text_clean):
            return True

    cleaned = re.sub(r'[\s\_\-:\(\)（）\[\]【】□]', '', text_clean)
    if not cleaned:
        return True

    return False


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
    details = {
        "steps": [],
        "tags": [],
        "matched_data": {},
        "replacements": 0
    }

    # Step 1: 合并碎片化 Run
    step_logger.log_step("文档预处理", "running")
    if progress_callback:
        progress_callback(1, "合并文档格式...")

    start_time = time.time()
    merge_adjacent_runs(doc)
    elapsed = int((time.time() - start_time) * 1000)
    details["steps"].append({
        "step": 1,
        "name": "文档解析与格式合并",
        "time_ms": elapsed
    })
    step_logger.log_step("文档预处理", "completed", {"time_ms": elapsed})

    # Step 2: AI 辅助打标
    step_logger.log_step("AI辅助打标", "running")
    if progress_callback:
        progress_callback(2, "AI正在识别空白处...")

    start_time = time.time()

    # 2.1 处理正文段落
    paragraphs_to_process = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if text and len(text) < 500:
            if any(c in text for c in ['_', '：', ':', '[', '□', '空', '待填写', '待补全']):
                paragraphs_to_process.append((i, para))

    logger.info(f"发现 {len(paragraphs_to_process)} 个可能包含空白的段落")
    step_logger.log_step("AI辅助打标", "running", {"total_paragraphs": len(paragraphs_to_process)})

    max_paragraphs = 25
    processed_count = 0
    skipped_count = 0

    for idx, (para_idx, para) in enumerate(paragraphs_to_process[:max_paragraphs]):
        try:
            original_text = para.text
            tagged_text = call_kimi_tagging(original_text, is_table=False)

            if "{{" in tagged_text and tagged_text != original_text:
                step_logger.log_api_call(
                    f"paragraph_{para_idx}",
                    original_text[:200],
                    tagged_text[:200]
                )
                if replace_text_in_paragraph(para, original_text, tagged_text):
                    processed_count += 1
            else:
                skipped_count += 1

            if (idx + 1) % 5 == 0 and progress_callback:
                progress_callback(2, f"已处理 {idx + 1}/{min(len(paragraphs_to_process), max_paragraphs)} 个段落...")

        except Exception as e:
            logger.error(f"打标段落 {para_idx} 失败: {e}")

    # 2.2 处理表格 - Markdown方案
    table_stats = {"processed_tables": 0, "total_tables": len(doc.tables)}

    for table_idx, table in enumerate(doc.tables[:10]):
        try:
            # 转换为Markdown
            md_table = table_to_markdown(table)
            step_logger.log_step(f"表格_{table_idx}_转换", "info", {"markdown": md_table[:500]})

            # 调用AI打标
            tagged_md = call_kimi_tagging(md_table, is_table=True)
            step_logger.log_api_call(
                f"table_{table_idx}_tagging",
                md_table[:500],
                tagged_md[:500]
            )

            # 解析Markdown结果
            md_rows = parse_markdown_table(tagged_md)

            # 应用回Word表格
            apply_markdown_tags_to_table(table, md_rows, step_logger)

            table_stats["processed_tables"] += 1

        except Exception as e:
            logger.error(f"处理表格 {table_idx} 失败: {e}")
            step_logger.log_step(f"表格_{table_idx}_错误", "error", {"error": str(e)})

    elapsed = int((time.time() - start_time) * 1000)
    details["steps"].append({
        "step": 2,
        "name": "AI辅助打标",
        "time_ms": elapsed,
        "stats": {
            "processed_paragraphs": processed_count,
            "skipped_paragraphs": skipped_count,
            "table_stats": table_stats
        }
    })
    step_logger.log_step("AI辅助打标", "completed", {"time_ms": elapsed, "stats": details["steps"][-1]["stats"]})

    # Step 3: 标签提取
    step_logger.log_step("标签提取", "running")
    if progress_callback:
        progress_callback(3, "提取标签...")

    start_time = time.time()
    tags = extract_tags_from_doc(doc)
    details["tags"] = tags
    elapsed = int((time.time() - start_time) * 1000)
    details["steps"].append({
        "step": 3,
        "name": "标签收集与去重",
        "time_ms": elapsed,
        "count": len(tags)
    })
    step_logger.log_step("标签提取", "completed", {"time_ms": elapsed, "count": len(tags), "tags": tags})

    # Step 4: 企业信息匹配
    step_logger.log_step("企业信息匹配", "running")
    if progress_callback:
        progress_callback(4, "匹配企业信息...")

    start_time = time.time()
    matched_data = call_kimi_match_info(tags, company_info)
    details["matched_data"] = matched_data
    elapsed = int((time.time() - start_time) * 1000)
    details["steps"].append({
        "step": 4,
        "name": "企业信息匹配",
        "time_ms": elapsed
    })
    step_logger.log_step("企业信息匹配", "completed", {"time_ms": elapsed, "matched": matched_data})

    # Step 5: 内容填充
    step_logger.log_step("内容填充", "running")
    if progress_callback:
        progress_callback(5, "填充内容...")

    start_time = time.time()
    replacement_count = 0
    failed_replacements = []

    # 处理正文
    for para_idx, para in enumerate(doc.paragraphs):
        for tag, value in matched_data.items():
            placeholder = f"{{{{{tag}}}}}"
            if placeholder in para.text:
                if replace_text_in_paragraph(para, placeholder, value):
                    replacement_count += 1
                    logger.debug(f"段落{para_idx}: 替换成功 {tag} -> {value}")
                else:
                    failed_replacements.append({"location": f"段落{para_idx}", "tag": tag, "text": para.text[:100]})
                    logger.warning(f"段落{para_idx}: 替换失败 {tag}")

    # 处理表格
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                for para in cell.paragraphs:
                    for tag, value in matched_data.items():
                        placeholder = f"{{{{{tag}}}}}"
                        if placeholder in para.text:
                            if replace_text_in_paragraph(para, placeholder, value):
                                replacement_count += 1
                                logger.debug(f"表格{table_idx}行{row_idx}单元格{cell_idx}: 替换成功 {tag}")
                            else:
                                failed_replacements.append({"location": f"表格{table_idx}行{row_idx}单元格{cell_idx}", "tag": tag, "text": para.text[:100]})
                                logger.warning(f"表格{table_idx}行{row_idx}单元格{cell_idx}: 替换失败 {tag}")

    if failed_replacements:
        step_logger.log_step("内容填充_失败项", "warning", {"failed": failed_replacements[:10]})

    elapsed = int((time.time() - start_time) * 1000)
    details["replacements"] = replacement_count
    details["steps"].append({
        "step": 5,
        "name": "内容填充与格式保留",
        "time_ms": elapsed,
        "count": replacement_count
    })
    step_logger.log_step("内容填充", "completed", {"time_ms": elapsed, "count": replacement_count})

    return doc, details


# ============== FastAPI 应用 ==============

app = FastAPI(
    title="资信标自动填充系统",
    description="基于 AI 的 Word 文档自动填充服务",
    version="2.0.0"
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
                time_ms=step_detail["time_ms"]
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
    print("资信标自动填充系统")
    print("=" * 50)
    print("服务地址: http://localhost:8001")
    print("=" * 50)
    uvicorn.run(app, host="0.0.0.0", port=8000)
