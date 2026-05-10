# 资信标自动填充系统

基于 FastAPI + Kimi API 的 Word 文档自动填充工具，完整实现技术路线文档中的五步流程。

## 功能特性

- **文档解析**：智能合并碎片化 Run，保留原始格式
- **AI 辅助打标**：调用 Kimi API 识别文档空白处并打上标签
- **标签提取**：自动提取所有 `{{标签}}`
- **企业信息匹配**：AI 智能匹配标签与企业信息
- **精准填充**：跨 Run 替换算法，100% 保留格式

## 快速开始

### 1. 安装依赖

```bash
# 创建虚拟环境
python -m venv venv

# 激活虚拟环境
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# 安装依赖
pip install -r requirements.txt
```

### 2. 配置环境变量

复制 `.env.example` 为 `.env` 并填入你的 Kimi API 密钥：

```bash
cp .env.example .env
```

编辑 `.env` 文件：

```env
KIMI_API_KEY=sk-your-api-key-here
KIMI_BASE_URL=https://api.moonshot.cn/v1
KIMI_MODEL=moonshot-v1-8k
```

### 3. 启动服务

**Windows:**
```bash
start.bat
```

**macOS/Linux:**
```bash
chmod +x start.sh
./start.sh
```

或直接使用 Python：
```bash
python main.py
```

### 4. 访问系统

打开浏览器访问：`http://localhost:8000`

## 使用流程

1. **上传 Word 模板文件**（仅支持 .docx 格式）
2. **输入企业信息**（自由文本格式）
3. **点击"开始填充"**
4. **查看处理进度**和调试信息
5. **下载填充后的文档**

## API 接口

### 上传文件
```http
POST /api/upload
Content-Type: multipart/form-data

file: <Word文件>
```

### 处理文档
```http
POST /api/process/{task_id}
Content-Type: multipart/form-data

company_info: 企业信息文本
filename: 原始文件名
```

### 下载结果
```http
GET /api/download/{task_id}/{filename}
```

## 项目结构

```
.
├── main.py              # 后端主程序
├── index.html           # 前端页面
├── requirements.txt     # Python 依赖
├── .env                 # 环境变量配置
├── .env.example         # 环境变量模板
├── uploads/             # 上传文件目录
├── outputs/             # 输出文件目录
└── README.md            # 本文件
```

## 核心算法说明

### 跨 Run 替换算法

Word 文档底层由多个 Run（文字块）组成，每个 Run 具有独立的格式。当标签跨越多个 Run 时，采用以下策略：

1. **字符索引映射**：计算目标文本在整个段落中的起始/结束位置
2. **Run 范围识别**：确定哪些 Run 包含目标文本
3. **首个 Run 保留**：保留第一个 Run 的格式，将替换文本放入其中
4. **后续 Run 清空**：清空其他包含目标文本的 Run（不删除对象，只清空文本）

这样可以确保格式 100% 保留，不会产生误删。

## 注意事项

1. **API 密钥安全**：请勿将包含真实密钥的 `.env` 文件提交到代码仓库
2. **文件大小限制**：默认最大 10MB，可在 `.env` 中修改 `MAX_FILE_SIZE`
3. **并发处理**：当前版本为简化 MVP，建议单用户单任务使用
4. **模型选择**：默认使用 `moonshot-v1-8k`，长文档可切换到 `moonshot-v1-32k`

## License

MIT
