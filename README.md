# PDF-AI-Gen_PPT

使用AI总结PDF文件内容并生成对应题目，然后生成PPT文件。

## 项目背景

本项目起因是学校老师喊帮忙，需要将PDF教材转换为试题和PPT。经过实测，使用DeepSeek的普通chat模型（非思考模型），生成一整本书、每章25道题，成本仅需约1元人民币。

## 功能特性

### 1. PDF内容解析与理解
- 深度分析PDF文件内容，提取文本信息
- 自动识别PDF书签结构，建立清晰的内容层级关系
- 自动检测并跳过目录页
- 支持多种章节格式识别（中英文）

### 2. 试题自动生成
- 为每个章节生成20+道单选题
- 每道题目包含4个选项，1个正确答案
- 题目覆盖核心知识点，难度分布合理
- 题目与答案分开存储，便于管理
- 支持指定章节生成、失败重试

### 3. PPT自动生成
- 将PDF内容转换为PPT格式
- 按章节划分独立PPT文件
- 包含标题页、内容页、总结页
- 支持自定义模板

### 4. 输出管理
- 支持JSON和Excel格式输出
- 自动生成处理报告
- 文件按章节命名，便于识别
- 增量保存，中断不丢失

## 安装

### 环境要求
- Python 3.9+
- pip 包管理器

### 安装步骤

1. 克隆项目
```bash
git clone https://github.com/Mo-cn/PDF-AI-Gen_PPT.git
cd PDF-AI-Gen_PPT
```

2. 创建虚拟环境（推荐）
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
# 或
venv\Scripts\activate  # Windows
```

3. 安装依赖
```bash
pip install -r requirements.txt
```

4. 配置环境变量
```bash
cp .env.example .env
# 编辑.env文件，填入你的AI API密钥
```

## 配置说明

编辑 `.env` 文件进行配置：

```ini
# AI配置
AI_PROVIDER=deepseek        # AI提供商: openai, anthropic, deepseek, custom
AI_API_KEY=your_api_key     # API密钥
AI_BASE_URL=                # 自定义API地址（可选）
AI_MODEL=deepseek-chat      # 模型名称

# 生成配置
QUESTIONS_PER_SECTION=25    # 每节题目数量
MAX_TOKENS_PER_REQUEST=8192 # 每次请求最大token
TEMPERATURE=0.5             # 生成温度（越低越稳定）

# 输出配置
OUTPUT_DIR=output           # 输出目录
```

### 支持的AI提供商

| 提供商 | AI_PROVIDER值 | 默认模型 | 说明 |
|--------|---------------|----------|------|
| DeepSeek | deepseek | deepseek-chat | 推荐，性价比高 |
| OpenAI | openai | gpt-4o | |
| Anthropic | anthropic | claude-3-5-sonnet-20241022 | |
| 自定义 | custom | 需配置AI_BASE_URL | |

## 使用方法

### 命令行界面

#### 查看章节列表
```bash
python run.py parse your_document.pdf
```

#### 生成试题
```bash
# 处理所有章节
python run.py questions your_document.pdf

# 指定章节生成
python run.py questions your_document.pdf -s 1,3,5-8

# 失败重试
python run.py questions your_document.pdf -s 10-15 -r 2
```

#### 生成PPT
```bash
python run.py ppt your_document.pdf --combined
```

#### 完整处理流程
```bash
python run.py process your_document.pdf
```

### 命令参数说明

#### questions 命令
| 参数 | 说明 | 默认值 |
|------|------|--------|
| pdf_path | PDF文件路径 | 必填 |
| -o, --output | 输出目录 | output |
| -q, --questions | 每节题目数量 | 25 |
| -s, --section | 指定章节编号(如: 1,3,5-8) | 全部 |
| -r, --retry | 失败重试次数 | 0 |

## 输出文件说明

### 目录结构
```
output/
├── questions_20240101_120000.json      # 试题JSON文件
├── questions_20240101_120000_题目.xlsx  # 题目Excel
├── questions_20240101_120000_答案.xlsx  # 答案Excel
├── structure_20240101_120000.json       # 文档结构
└── ppt/                                  # PPT文件目录
    ├── 第一章.pptx
    └── ...
```

## 注意事项

1. **API密钥安全**：请勿将API密钥提交到版本控制系统
2. **费用控制**：使用DeepSeek生成一整本书约1元，其他模型请参考官方定价
3. **内容审核**：生成的内容建议人工审核后使用

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！

项目地址: https://github.com/Mo-cn/PDF-AI-Gen_PPT
