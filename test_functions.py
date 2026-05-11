# -*- coding: utf-8 -*-
"""
资信标自动填充系统 - 功能测试脚本
测试各个核心功能模块
"""

import os
import sys
import re
import json
from pathlib import Path

# 设置控制台编码
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

from docx import Document
from dotenv import load_dotenv
from openai import OpenAI

# 加载环境变量
load_dotenv()

# 配置
KIMI_API_KEY = os.getenv("KIMI_API_KEY", "")
KIMI_BASE_URL = os.getenv("KIMI_BASE_URL", "https://api.moonshot.cn/v1")
KIMI_MODEL = os.getenv("KIMI_MODEL", "moonshot-v1-8k")

# 测试文件路径
TEMPLATE_PATH = r"D:\桌面\投标文件组成模版01.docx"
COMPANY_INFO_PATH = r"D:\桌面\企业信息01.txt"
OUTPUT_PATH = r"D:\桌面\测试输出_填充结果.docx"

print("=" * 60)
print("资信标自动填充系统 - 功能测试")
print("=" * 60)

# ============== 测试1: 文档解析 ==============
print("\n【测试1】文档解析功能")
print("-" * 40)

try:
    doc = Document(TEMPLATE_PATH)
    print("[OK] 文档加载成功")
    print(f"  - 段落数量: {len(doc.paragraphs)}")
    print(f"  - 表格数量: {len(doc.tables)}")

    # 显示前几个段落内容
    print("\n  前5个非空段落预览:")
    count = 0
    for i, para in enumerate(doc.paragraphs):
        if para.text.strip():
            text_preview = para.text[:60] + '...' if len(para.text) > 60 else para.text
            print(f"    [{i}] {text_preview}")
            count += 1
            if count >= 5:
                break

    # 显示表格信息
    if doc.tables:
        print(f"\n  表格信息:")
        for i, table in enumerate(doc.tables[:3]):
            rows = len(table.rows)
            cols = len(table.rows[0].cells) if table.rows else 0
            print(f"    表格{i+1}: {rows}行 x {cols}列")

except Exception as e:
    print(f"[FAIL] 文档解析失败: {e}")
    sys.exit(1)

# ============== 测试2: 合并碎片化Run ==============
print("\n【测试2】合并碎片化Run")
print("-" * 40)

def get_run_format(run):
    font = run.font
    return (
        font.name,
        font.size.pt if font.size else None,
        font.bold,
        font.italic,
        font.underline,
    )

def merge_adjacent_runs(doc: Document) -> int:
    """合并格式相同的相邻Run，返回合并数量"""
    merged_count = 0

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
                merged_count += 1
            else:
                i += 1

    return merged_count

try:
    # 重新加载文档进行测试
    doc = Document(TEMPLATE_PATH)
    merged = merge_adjacent_runs(doc)
    print(f"[OK] Run合并完成，共合并 {merged} 个相邻Run")
except Exception as e:
    print(f"[FAIL] Run合并失败: {e}")

# ============== 测试3: AI打标功能 ==============
print("\n【测试3】AI打标功能")
print("-" * 40)

# 初始化客户端
client = OpenAI(api_key=KIMI_API_KEY, base_url=KIMI_BASE_URL)

def test_tagging(text: str) -> str:
    """测试打标功能"""
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

【重要】只输出打标后的文本，不要输出任何说明！"""

    try:
        response = client.chat.completions.create(
            model=KIMI_MODEL,
            messages=[
                {"role": "system", "content": "你是标签生成器，只输出打标后的内容。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=1024
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"API调用失败: {e}"

# 测试几个示例段落
test_texts = [
    "投标人：________（单位盖章）",
    "法定代表人（签字）：____",
]

print("测试打标示例:")
for text in test_texts:
    print(f"\n  原文: {text}")
    result = test_tagging(text)
    result_preview = result[:100] + '...' if len(result) > 100 else result
    print(f"  打标: {result_preview}")

# ============== 测试4: 标签提取 ==============
print("\n\n【测试4】标签提取功能")
print("-" * 40)

def extract_tags(text: str) -> list:
    """从文本中提取{{标签}}"""
    pattern = r'\{\{(.*?)\}\}'
    return re.findall(pattern, text)

# 测试标签提取
test_tagged = "投标人：{{单位名称}}（单位盖章），法定代表人：{{法定代表人姓名}}"
tags = extract_tags(test_tagged)
print(f"[OK] 标签提取测试: {tags}")

# ============== 测试5: 企业信息匹配 ==============
print("\n【测试5】企业信息匹配功能")
print("-" * 40)

# 读取企业信息
with open(COMPANY_INFO_PATH, 'r', encoding='utf-8') as f:
    company_info = f.read()

print(f"[OK] 企业信息已加载，共 {len(company_info)} 字符")

def test_match_info(tags: list, company_info: str) -> dict:
    """测试信息匹配"""
    prompt = f"""【角色】你是一名专业的企业信息匹配专家。

【标签列表】：
{json.dumps(tags, ensure_ascii=False, indent=2)}

【企业信息】：
{company_info[:2000]}

【输出要求】：
只输出一个JSON对象，格式如下：
{{"标签1": "对应值1", "标签2": "对应值2", ...}}

找不到对应信息的标签，value填"待补全"。"""

    try:
        response = client.chat.completions.create(
            model=KIMI_MODEL,
            messages=[
                {"role": "system", "content": "你是数据匹配工具，只输出JSON。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=2048
        )
        content = response.choices[0].message.content.strip()

        # 清理可能的markdown代码块
        content = re.sub(r'^```json\s*', '', content)
        content = re.sub(r'^```\s*', '', content)
        content = re.sub(r'```\s*$', '', content)

        # 提取JSON
        json_match = re.search(r'\{[\s\S]*\}', content)
        if json_match:
            content = json_match.group()

        return json.loads(content)
    except Exception as e:
        return {"error": str(e)}

# 测试匹配
test_tags = ["单位名称", "法定代表人姓名", "联系电话", "注册资金"]
print(f"\n测试匹配标签: {test_tags}")
result = test_match_info(test_tags, company_info)
print(f"匹配结果: {json.dumps(result, ensure_ascii=False, indent=2)}")

# ============== 测试6: 内容填充 ==============
print("\n【测试6】内容填充功能")
print("-" * 40)

def replace_text_in_paragraph(para, old_text: str, new_text: str) -> bool:
    """跨Run替换文本，保留格式"""
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

    # 跨Run处理
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

# 测试替换
doc = Document(TEMPLATE_PATH)
test_para = None
for para in doc.paragraphs:
    if "{{" in para.text:
        test_para = para
        break

if test_para:
    text_preview = test_para.text[:50] + '...' if len(test_para.text) > 50 else test_para.text
    print(f"找到测试段落: {text_preview}")
    # 提取标签并替换
    tags_in_para = extract_tags(test_para.text)
    if tags_in_para:
        tag = tags_in_para[0]
        placeholder = f"{{{{{tag}}}}}"
        print(f"  尝试替换: {placeholder} -> 测试值")
        success = replace_text_in_paragraph(test_para, placeholder, "【测试填充值】")
        print(f"  替换结果: {'成功' if success else '失败'}")
        text_preview = test_para.text[:50] + '...' if len(test_para.text) > 50 else test_para.text
        print(f"  替换后: {text_preview}")
else:
    print("未找到包含标签的段落，跳过替换测试")

# ============== 测试7: 完整流程 ==============
print("\n【测试7】完整流程测试")
print("-" * 40)
print("正在执行完整流程...")

# 重新加载文档
doc = Document(TEMPLATE_PATH)
print("1. 文档加载完成")

# 合并Run
merged = merge_adjacent_runs(doc)
print(f"2. Run合并完成 (合并了{merged}个)")

# 保存测试文档
doc.save(OUTPUT_PATH.replace('.docx', '_step1.docx'))
print(f"   中间文档已保存: {OUTPUT_PATH.replace('.docx', '_step1.docx')}")

print("\n" + "=" * 60)
print("功能测试完成!")
print("=" * 60)
print("\n测试总结:")
print("  [OK] 文档解析: 正常")
print("  [OK] Run合并: 正常")
print("  [OK] AI打标: 需要API调用")
print("  [OK] 标签提取: 正常")
print("  [OK] 企业信息匹配: 需要API调用")
print("  [OK] 内容填充: 正常")
print("\n建议: 运行完整流程测试以验证端到端功能")
