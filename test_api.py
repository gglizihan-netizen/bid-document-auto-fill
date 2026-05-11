# -*- coding: utf-8 -*-
"""
测试 Web API 接口
"""

import os
import sys
import json
import time

if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import httpx

API_BASE = "http://localhost:8000"

def test_api():
    print("=" * 60)
    print("测试 Web API 接口")
    print("=" * 60)

    # 读取测试文件
    template_path = r"D:\桌面\投标文件组成模版01.docx"
    company_info_path = r"D:\桌面\企业信息01.txt"

    with open(company_info_path, 'r', encoding='utf-8') as f:
        company_info = f.read()

    print(f"\n企业信息长度: {len(company_info)} 字符")

    # Step 1: 上传文件
    print("\n【Step 1】上传文件...")
    with open(template_path, 'rb') as f:
        files = {'file': ('document.docx', f, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')}
        response = httpx.post(f"{API_BASE}/api/upload", files=files, timeout=30)

    print(f"状态码: {response.status_code}")
    if response.status_code != 200:
        print(f"上传失败: {response.text}")
        return

    upload_result = response.json()
    task_id = upload_result.get('task_id')
    print(f"任务ID: {task_id}")

    # Step 2: 处理文档
    print("\n【Step 2】处理文档...")
    data = {
        "filename": "document.docx",
        "company_info": company_info
    }

    start_time = time.time()
    response = httpx.post(
        f"{API_BASE}/api/process/{task_id}",
        json=data,
        timeout=300  # 5分钟超时
    )
    elapsed = time.time() - start_time

    print(f"状态码: {response.status_code}")
    print(f"处理耗时: {elapsed:.2f}秒")

    if response.status_code != 200:
        print(f"处理失败: {response.text}")
        return

    result = response.json()
    print(f"\n处理结果:")
    print(f"  - 成功: {result.get('success')}")
    print(f"  - 标签数量: {len(result.get('tags', []))}")
    print(f"  - 匹配数据数量: {len(result.get('matched_data', {}))}")

    if result.get('tags'):
        print(f"\n标签列表 (前10个):")
        for tag in result.get('tags', [])[:10]:
            print(f"    - {tag}")

    if result.get('matched_data'):
        print(f"\n匹配数据 (前10个):")
        for i, (k, v) in enumerate(list(result.get('matched_data', {}).items())[:10]):
            print(f"    - {k}: {v}")

    if result.get('download_url'):
        print(f"\n下载地址: {API_BASE}{result.get('download_url')}")

    print("\n" + "=" * 60)
    print("API测试完成!")
    print("=" * 60)

if __name__ == "__main__":
    test_api()
