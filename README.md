# AI智能文档润色/翻译助手

> 基于大模型API的专业文本处理工具，集成学术润色、技术文档优化、商务改写、中英翻译等多功能模块。

## 项目概述

本项目是"启航科创"工程实践项目，旨在通过调用国内外大模型（通义千问/DeepSeek/OpenAI）API，构建一个轻量级、易用的智能文档处理Web应用。

### 核心功能

| 功能模块 | 说明 |
|---------|------|
| 严谨学术风 | 学术论文语言润色、语法修正、表达优化 |
| 硅谷极客风 | 技术文档、代码注释优化，符合业界规范 |
| 简洁商务风 | 商务邮件、工作报告的专业化改写 |
| 中英翻译 | 专业学术级中英互译，保持术语一致 |
| 代码解释 | 为代码添加详细中文注释和解释 |

### 技术栈

- **Python 3.8+**: 核心编程语言
- **Streamlit**: 极简Web应用框架
- **LLM API**: 大语言模型接口调用
- **Requests**: HTTP通信库

## 快速开始

### 1. 环境准备

```bash
# 克隆项目（或直接下载）
cd ai-writing-assistant

# 创建虚拟环境（推荐）
python -m venv venv

# 激活虚拟环境
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# 安装依赖
pip install -r requirements.txt
```

### 2. 获取API Key

推荐使用**通义千问**（免费额度充足）：

1. 访问 [通义千问控制台](https://dashscope.aliyun.com)
2. 注册/登录阿里云账号
3. 创建API Key（首次使用赠送100万Token免费额度）
4. 复制API Key备用

其他可选平台：
- [DeepSeek](https://platform.deepseek.com) 
- [OpenAI](https://platform.openai.com)

### 3. 启动应用

```bash
streamlit run app.py
```

浏览器将自动打开 `http://localhost:8501`

### 4. 使用流程

1. 在左侧边栏粘贴API Key
2. 点击"测试连接"验证API Key有效
3. 选择适合的角色模板（学术/技术/商务/翻译）
4. 输入或上传待处理文本
5. 点击"开始处理"按钮
6. 查看结果并导出（支持Word/TXT/Markdown格式）

## 部署上线

### Streamlit Cloud（推荐，免费）

5分钟完成部署，自动分配域名，每次 `git push` 自动更新：

```bash
# 1. 推送代码到GitHub
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/你的用户名/ai-writing-assistant.git
git push -u origin main

# 2. 访问 https://streamlit.io/cloud 登录
# 3. 点击 "New app" → 选择GitHub仓库
# 4. 在 Secrets 中设置 DASHSCOPE_API_KEY
# 5. 点击 Deploy! 完成部署
```

详细步骤请参阅 [DEPLOY.md](DEPLOY.md)

### Docker 部署

```bash
# 构建镜像
docker build -t ai-writing-assistant .

# 运行容器
docker run -d -p 8501:8501 -e DASHSCOPE_API_KEY=sk-xxx ai-writing-assistant
```

## 项目结构

```
ai-writing-assistant/
├── app.py                      # 主程序（Streamlit界面和交互逻辑）
├── config.py                   # 配置文件（API配置、角色模板、UI文本）
├── requirements.txt            # Python依赖清单
├── README.md                   # 项目说明文档
├── utils/
│   ├── __init__.py
│   ├── api_client.py           # 大模型API客户端（核心调用逻辑）
│   └── document_processor.py   # 文档处理工具（读写、统计、导出）
└── assets/
    └── demo/                   # 演示数据
```

## 核心模块说明

### api_client.py - API客户端

封装了对各大模型API的调用，提供统一的接口：

```python
from utils.api_client import LLMClient

# 初始化客户端
client = LLMClient(
    provider="通义千问",      # API提供商
    api_key="sk-xxx",         # API密钥
    model="qwen-turbo"        # 模型名称
)

# 非流式调用（适合短文本）
result = client.polish_text("待润色文本", "系统提示词")

# 流式调用（实时显示，体验更好）
for chunk in client.polish_text_stream("待润色文本", "系统提示词"):
    print(chunk, end="", flush=True)
```

**关键技术点：**
- 兼容OpenAI标准格式（通义千问/DeepSeek/OpenAI通用）
- 支持流式输出（SSE协议），实现打字机效果
- 完善的错误处理和日志记录
- 自动重试和超时控制

### document_processor.py - 文档处理器

提供文本导入导出和统计功能：

```python
from utils.document_processor import (
    read_docx, save_docx,      # Word读写
    read_txt, save_txt,        # 文本读写
    count_text_stats,          # 字数统计
    detect_language,           # 语言检测
)
```

### config.py - 配置中心

集中管理所有配置：
- `API_PROVIDERS`: 支持的API平台配置
- `ROLE_TEMPLATES`: 5种角色的Prompt Engineering模板
- `APP_CONFIG`: 应用参数（超时时间、最大长度等）

## Prompt Engineering设计

本项目核心亮点之一是对Prompt的精心设计，每个角色都有独立的系统提示词：

**严谨学术风示例：**
```
你是一位拥有20年经验的资深学术编辑，精通中英文学术写作规范。

【任务要求】
1. 对输入文本进行深度语言润色，提升学术规范性
2. 优化句子结构，增强逻辑连贯性
3. 修正语法错误、拼写错误和标点使用
4. 使用恰当的学术词汇和表达方式
5. 保持原文的学术观点和核心结论不变
```

提示词设计遵循原则：角色定义→任务要求→输出规范的三段式结构，确保输出质量稳定可控。

## 简历亮点

> "熟悉提示词工程（Prompt Engineering），掌握大模型API调用逻辑及Web应用轻量化开发"

- **LLM应用开发**: 深入理解大模型API调用机制，包括流式输出、错误处理、多平台兼容
- **Prompt Engineering**: 设计高质量的角色提示词模板，实现精准的文本定向处理
- **Web应用开发**: 使用Streamlit快速构建交互式Web应用，掌握前后端一体化开发
- **文档处理**: 实现Word/TXT/Markdown多格式导入导出，文本统计与分析

## 常见问题

**Q: 为什么推荐通义千问？**
A: 通义千问对中文支持优秀，且新用户赠送100万Token免费额度，足够项目演示使用。

**Q: 可以离线使用吗？**
A: 不可以，本项目需要调用云端大模型API，使用时需要网络连接。

**Q: 支持哪些文件格式？**
A: 支持.docx（Word）和.txt（纯文本）格式的导入，支持Word/TXT/Markdown三种格式导出。

**Q: 文本有长度限制吗？**
A: 默认最大输入5000字符，可在config.py中修改`max_input_length`调整。

## 许可

本项目仅供学习交流使用。

---

**东北大学秦皇岛分校 · 计算机与通信工程学院 · 启航科创项目**
