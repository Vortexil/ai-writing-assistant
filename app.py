"""
AI智能文档润色/翻译助手 v3.0
功能完备版 · 自动语言检测 · 双向翻译 · Word进Word出 · PDF摘要
"""

import io
import json
import logging
import os
import re
import time
from typing import Generator, Optional, List, Dict

import requests
import streamlit as st

# ============================================================
# 配置
# ============================================================

API_PROVIDERS = {
    "通义千问": {
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
        "models": ["qwen-turbo", "qwen-plus", "qwen-max"],
        "default_model": "qwen-turbo",
    },
    "DeepSeek": {
        "base_url": "https://api.deepseek.com/chat/completions",
        "models": ["deepseek-chat", "deepseek-reasoner"],
        "default_model": "deepseek-chat",
    },
}

# 润色角色（5种）
POLISH_ROLES = {
    "严谨学术风": {
        "icon": "🎓",
        "color": "#3b82f6",
        "desc": "学术论文润色",
        "prompt": "你是一位资深学术编辑。对以下文本进行深度润色：提升学术规范性，优化句子结构，修正语法错误，使用恰当的学术词汇，保持原文核心观点和结论不变。直接返回润色后的完整文本，不要添加任何解释说明或额外评论。",
    },
    "硅谷极客风": {
        "icon": "⚡",
        "color": "#06b6d4",
        "desc": "技术文档优化",
        "prompt": "你是一位在硅谷顶级科技公司工作的技术文档专家。优化以下技术文档的表达：使其简洁、专业、易读，统一技术术语的使用，删除冗余表达，遵循Google/Apple技术文档风格。直接返回优化后的完整文本，不要添加任何解释说明。",
    },
    "简洁商务风": {
        "icon": "💼",
        "color": "#8b5cf6",
        "desc": "商务报告改写",
        "prompt": "你是一位资深商务沟通顾问。将以下内容改写为简洁、专业的商务表达：突出关键信息和行动项，使用商务场景惯用的礼貌用语，删除口语化表达，使用数据驱动的方式。直接返回改写后的完整文本，不要添加任何解释说明。",
    },
    "中英互译": {
        "icon": "🌏",
        "color": "#10b981",
        "desc": "自动检测语言双向翻译",
        "prompt_cn_to_en": "你是一位专业的英文学术翻译专家。请将以下中文翻译成地道、专业的学术英文：保持专业术语准确，使用符合英文学术论文表达习惯的句式，确保翻译后的文本流畅自然。直接返回英文翻译结果，不要添加任何解释。",
        "prompt_en_to_cn": "你是一位专业的学术翻译专家。请将以下英文翻译成准确、流畅的中文学术表达：保持专业术语的翻译一致性，确保符合中文学术论文的表达习惯。直接返回中文翻译结果，不要添加任何解释。",
    },
    "代码解释": {
        "icon": "💻",
        "color": "#f59e0b",
        "desc": "为代码添加详细中文注释",
        "prompt": "你是一位经验丰富的技术导师。请为以下代码添加清晰、详尽的中文注释：为每一行关键代码添加注释，解释设计思路和实现逻辑，说明函数参数的含义和返回值，标注潜在的性能优化点。直接返回带有详细注释的完整代码，不要添加任何解释说明。",
    },
}

# PDF摘要配置
PDF_CONFIG = {
    "chunk_size": 3000,
    "chunk_overlap": 200,
    "max_chunks": 10,
}

APP_CONFIG = {
    "max_input_length": 8000,
    "default_temperature": 0.3,
    "timeout_seconds": 120,
}

# ============================================================
# API 客户端
# ============================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


class LLMClient:
    def __init__(self, provider: str, api_key: str, model: Optional[str] = None):
        if provider not in API_PROVIDERS:
            raise ValueError(f"不支持的API提供商: {provider}")
        self.provider = provider
        self.api_key = api_key
        self.config = API_PROVIDERS[provider]
        self.model = model or self.config["default_model"]
        self.base_url = self.config["base_url"]

    def process_sync(self, text: str, system_prompt: str) -> str:
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
        payload = {
            "model": self.model,
            "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": text}],
            "temperature": APP_CONFIG["default_temperature"],
            "stream": False,
        }
        response = requests.post(self.base_url, headers=headers, json=payload,
                                 timeout=APP_CONFIG["timeout_seconds"])
        if response.status_code != 200:
            raise ValueError(f"API请求失败 [{response.status_code}]")
        return response.json()["choices"][0]["message"]["content"]

    def stream_process(self, text: str, system_prompt: str) -> Generator[str, None, None]:
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
        payload = {
            "model": self.model,
            "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": text}],
            "temperature": APP_CONFIG["default_temperature"],
            "stream": True,
        }
        try:
            response = requests.post(self.base_url, headers=headers, json=payload,
                                     timeout=APP_CONFIG["timeout_seconds"], stream=True)
        except requests.exceptions.Timeout:
            raise TimeoutError("请求超时")
        except requests.exceptions.ConnectionError:
            raise ConnectionError("网络连接失败")
        if response.status_code != 200:
            raise ValueError(f"API请求失败 [{response.status_code}]")
        for line in response.iter_lines(decode_unicode=True):
            if not line or not line.startswith("data: "):
                continue
            data = line[6:]
            if data == "[DONE]":
                break
            try:
                chunk = json.loads(data)
                content = chunk.get("choices", [{}])[0].get("delta", {}).get("content", "")
                if content:
                    yield content
            except:
                continue


# ============================================================
# 语言检测
# ============================================================

def detect_language(text: str) -> str:
    """检测文本主要语言：zh(中文), en(英文), mixed(混合)"""
    cn_chars = len(re.findall(r"[\u4e00-\u9fff]", text))
    en_chars = len(re.findall(r"[a-zA-Z]", text))
    total = len(text.strip())
    if total == 0:
        return "unknown"
    if cn_chars / total > 0.3:
        return "zh"
    if en_chars / total > 0.3:
        return "en"
    return "mixed"


def get_translation_prompt(role_info: dict, text: str) -> str:
    """根据输入语言自动选择翻译方向的提示词"""
    lang = detect_language(text)
    if lang == "zh":
        return role_info["prompt_cn_to_en"]
    else:
        return role_info["prompt_en_to_cn"]


def get_translation_label(text: str) -> str:
    """获取翻译方向的显示标签"""
    lang = detect_language(text)
    if lang == "zh":
        return "🇨🇳 中文 → 🇬🇧 英文"
    elif lang == "en":
        return "🇬🇧 英文 → 🇨🇳 中文"
    else:
        return "🌐 自动检测翻译方向"


# ============================================================
# Word 文档处理
# ============================================================

def read_docx_structured(file_bytes: bytes) -> List[Dict]:
    from docx import Document
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = []
    for p in doc.paragraphs:
        if p.text.strip():
            paragraphs.append({"text": p.text.strip(), "style": p.style.name if p.style else "Normal"})
    return paragraphs


def save_docx_structured(paragraphs: List[Dict], title: str = "润色结果") -> bytes:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from datetime import datetime

    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Microsoft YaHei'
    font.size = Pt(11)

    heading = doc.add_heading(title, level=0)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1e, 0x3a, 0x5f)

    doc.add_paragraph("")
    info = doc.add_paragraph(f"处理时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info.runs[0].font.size = Pt(9)
    info.runs[0].font.color.rgb = RGBColor(0x94, 0xa3, 0xb8)
    doc.add_paragraph("")
    doc.add_paragraph("─" * 40)
    doc.add_paragraph("")

    for para in paragraphs:
        text = para.get("text", "").strip()
        style_name = para.get("style", "Normal")
        if not text:
            continue
        if "Heading" in style_name or "标题" in style_name:
            level = 1 if "1" in style_name else 2 if "2" in style_name else 3 if "3" in style_name else 1
            doc.add_heading(text, level=level)
        else:
            doc.add_paragraph(text)

    doc.add_paragraph("")
    doc.add_paragraph("─" * 40)
    footer = doc.add_paragraph("由 AI智能文档润色助手 生成")
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.runs[0].font.size = Pt(9)
    footer.runs[0].font.color.rgb = RGBColor(0x94, 0xa3, 0xb8)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


def merge_short_paragraphs(paragraphs: List[Dict], max_chars: int = 1500) -> List[Dict]:
    """合并短段落减少API调用"""
    if not paragraphs:
        return []
    merged = []
    buffer_text = ""
    buffer_style = "Normal"
    for para in paragraphs:
        text = para["text"]
        style = para["style"]
        if "Heading" in style or "标题" in style:
            if buffer_text:
                merged.append({"text": buffer_text.strip(), "style": buffer_style})
                buffer_text = ""
            merged.append(para)
            continue
        if len(buffer_text) + len(text) < max_chars:
            buffer_text += "\n\n" + text if buffer_text else text
            buffer_style = style
        else:
            if buffer_text:
                merged.append({"text": buffer_text.strip(), "style": buffer_style})
            buffer_text = text
            buffer_style = style
    if buffer_text:
        merged.append({"text": buffer_text.strip(), "style": buffer_style})
    return merged


# ============================================================
# PDF 处理
# ============================================================

def read_pdf(file_bytes: bytes) -> str:
    import PyPDF2
    reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
    texts = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            texts.append(text.strip())
    return "\n\n".join(texts)


def split_text_to_chunks(text: str, chunk_size: int = 3000, overlap: int = 200) -> List[str]:
    if len(text) <= chunk_size:
        return [text]
    chunks = []
    start = 0
    while start < len(text):
        end = start + chunk_size
        if end >= len(text):
            chunks.append(text[start:])
            break
        cut_pos = text.rfind("。", start, end)
        if cut_pos == -1:
            cut_pos = text.rfind("\n", start, end)
        if cut_pos == -1:
            cut_pos = end
        chunks.append(text[start:cut_pos + 1])
        start = cut_pos + 1 - overlap
    return chunks[:PDF_CONFIG["max_chunks"]]


def summarize_chunk(client: LLMClient, chunk: str, idx: int, total: int) -> str:
    prompt = f"""你是一位学术摘要专家。请对以下文档片段（第{idx}/{total}部分）进行摘要：

要求：
1. 提取核心观点、关键数据、重要结论
2. 用中文输出，分点列出
3. 简明扼要，每点不超过30字

文档片段：
{chunk}"""
    return client.process_sync(chunk, prompt)


def generate_final_summary(client: LLMClient, merged: str) -> str:
    prompt = f"""你是一位学术摘要专家。请根据以下分块摘要，生成一份完整的文档摘要报告：

要求：
1. 包含：【文档主题】→【核心观点】→【研究方法】→【关键结论】→【主要贡献】
2. 每个部分用中文输出，清晰分点
3. 总字数控制在300-500字
4. 语言严谨、专业

分块摘要：
{merged}"""
    return client.process_sync(merged, prompt)


# ============================================================
# 工具函数
# ============================================================

def count_stats(text: str) -> dict:
    clean = text.strip()
    cn = len(re.findall(r"[\u4e00-\u9fff]", clean))
    return {"总字符": len(clean), "中文": cn, "段落": max(len([p for p in clean.split("\n\n") if p.strip()]), 1)}


# ============================================================
# Streamlit 界面
# ============================================================

st.set_page_config(page_title="AI智能文档润色助手", page_icon="📝", layout="wide")

st.markdown("""
<style>
.block-container { padding: 1rem 2rem !important; max-width: 100% !important; }
.hero-section {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 50%, #0d9488 100%);
    border-radius: 16px; padding: 2rem 2.5rem; margin-bottom: 1.5rem;
    color: white; box-shadow: 0 8px 32px rgba(15,23,42,0.3);
}
.hero-title { font-size: 2rem !important; font-weight: 700; margin: 0; }
.hero-subtitle { font-size: 0.95rem; color: #94a3b8; margin-top: 0.3rem; }
.role-card {
    border: 2px solid #e2e8f0; border-radius: 14px; padding: 1rem 0.8rem;
    text-align: center; background: white; transition: all 0.3s;
}
.role-card:hover { transform: translateY(-3px); box-shadow: 0 8px 24px rgba(0,0,0,0.08); }
.role-card.active { border-color: #3b82f6; background: #eff6ff; }
.role-icon { font-size: 1.8rem; margin-bottom: 0.3rem; }
.role-name { font-weight: 600; font-size: 0.85rem; color: #1e293b; }
.role-desc { font-size: 0.72rem; color: #94a3b8; margin-top: 2px; }
.pdf-card {
    background: linear-gradient(135deg, #fef2f2, #fee2e2);
    border: 2px solid #ef4444; border-radius: 14px; padding: 1.2rem;
    text-align: center; cursor: pointer; transition: all 0.3s;
}
.pdf-card:hover { transform: translateY(-3px); box-shadow: 0 8px 24px rgba(239,68,68,0.15); }
.stat-box {
    background: linear-gradient(135deg, #f0f9ff, #e0f2fe);
    border-radius: 10px; padding: 8px 12px; text-align: center;
}
.stat-num { font-size: 1.3rem; font-weight: 700; color: #0369a1; }
.stat-label { font-size: 0.7rem; color: #64748b; }
.result-area {
    background: #f8fafc; border-left: 4px solid #0d9488;
    border-radius: 0 10px 10px 0; padding: 1.2rem;
    min-height: 380px; max-height: 600px; overflow-y: auto;
    font-size: 0.95rem; line-height: 1.8; color: #334155;
}
.code-result {
    background: #1e293b; border-radius: 10px; padding: 1.2rem;
    min-height: 380px; max-height: 600px; overflow-y: auto;
    font-family: 'Consolas', 'Monaco', 'Courier New', monospace !important;
    font-size: 0.8rem !important; line-height: 1.6 !important;
    color: #e2e8f0 !important;
}
.lang-badge {
    display: inline-block;
    background: linear-gradient(135deg, #dbeafe, #d1fae5);
    border-radius: 20px;
    padding: 4px 14px;
    font-size: 0.8rem;
    color: #1e40af;
    font-weight: 500;
}
#MainMenu, footer, header, .stDeployButton { display: none !important; }
section[data-testid="stSidebar"] { min-width: 280px !important; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 会话状态
# ============================================================

if "result" not in st.session_state:
    st.session_state.result = ""
if "structured_paras" not in st.session_state:
    st.session_state.structured_paras = []
if "is_word_mode" not in st.session_state:
    st.session_state.is_word_mode = False
if "is_pdf_mode" not in st.session_state:
    st.session_state.is_pdf_mode = False
if "count" not in st.session_state:
    st.session_state.count = 0
if "role" not in st.session_state:
    st.session_state.role = "严谨学术风"
if "stop_flag" not in st.session_state:
    st.session_state.stop_flag = False
if "processing" not in st.session_state:
    st.session_state.processing = False

api_key = ""

# ============================================================
# 侧边栏
# ============================================================

with st.sidebar:
    st.header("⚙️ 设置")
    provider = st.selectbox("API提供商", list(API_PROVIDERS.keys()), index=0)

    default_key = ""
    if "DASHSCOPE_API_KEY" in st.secrets:
        default_key = st.secrets["DASHSCOPE_API_KEY"]
    elif "DASHSCOPE_API_KEY" in os.environ:
        default_key = os.environ["DASHSCOPE_API_KEY"]

    if default_key:
        api_key = default_key
        st.success("✅ API Key已加载")
    else:
        api_key = st.text_input("API Key", type="password", placeholder="粘贴你的Key")

    model = st.selectbox("模型", API_PROVIDERS[provider]["models"], index=0)
    st.markdown("---")
    st.caption("💡 推荐通义千问，免费额度充足")

# ============================================================
# Hero 区域
# ============================================================

st.markdown("""
<div class="hero-section">
    <div style="display:flex;justify-content:space-between;align-items:center;">
        <div>
            <div class="hero-title">📝 AI 智能文档润色助手</div>
            <div class="hero-subtitle">Word进Word出 · PDF智能摘要 · 中英双向翻译 · 保留原文格式</div>
        </div>
        <div style="font-size:2.5rem;font-weight:700;opacity:0.9;">v3.0</div>
    </div>
</div>
""", unsafe_allow_html=True)

if not api_key:
    st.warning("⚠️ 请在左侧边栏输入API Key")
    st.stop()

# ============================================================
# 功能选择区
# ============================================================

st.subheader("🎭 选择功能")

cols = st.columns(len(POLISH_ROLES))
for col, (key, info) in zip(cols, POLISH_ROLES.items()):
    with col:
        active = "active" if st.session_state.role == key and not st.session_state.is_pdf_mode else ""
        st.markdown(f"""
        <div class="role-card {active}">
            <div class="role-icon">{info['icon']}</div>
            <div class="role-name">{key}</div>
            <div class="role-desc">{info['desc']}</div>
        </div>""", unsafe_allow_html=True)
        if st.button("选择", key=f"sel_{key}", use_container_width=True):
            st.session_state.role = key
            st.session_state.is_word_mode = False
            st.session_state.is_pdf_mode = False
            st.session_state.structured_paras = []
            st.session_state.result = ""
            st.rerun()

# PDF摘要独立入口
st.markdown("<div style='margin-top:0.8rem'></div>", unsafe_allow_html=True)
pdf_cols = st.columns([1, 3, 1])
with pdf_cols[1]:
    pdf_active = "style='border-color:#ef4444;background:#fef2f2;'" if st.session_state.is_pdf_mode else ""
    st.markdown(f"""
    <div class="pdf-card" {pdf_active}>
        <div style="font-size:2rem;margin-bottom:0.3rem">📄</div>
        <div style="font-weight:700;font-size:1.1rem;color:#991b1b">PDF 智能摘要</div>
        <div style="font-size:0.8rem;color:#b91c1c;margin-top:0.3rem">上传PDF文档 → 自动分段 → 生成结构化摘要报告</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("📄 进入PDF摘要模式", key="pdf_mode_btn", use_container_width=True):
        st.session_state.is_pdf_mode = True
        st.session_state.is_word_mode = False
        st.session_state.role = ""
        st.session_state.structured_paras = []
        st.session_state.result = ""
        st.rerun()

# 当前模式显示
if st.session_state.is_pdf_mode:
    st.info("📄 **PDF智能摘要模式** — 上传PDF文档，AI自动分段生成结构化摘要报告")
else:
    info = POLISH_ROLES[st.session_state.role]
    st.info(f"{info['icon']} **{st.session_state.role}** — {info['desc']}")

st.markdown("---")

# ============================================================
# 输入区域
# ============================================================

st.subheader("📥 输入")

input_text = ""
source_paras = []

if st.session_state.is_pdf_mode:
    file = st.file_uploader("📄 上传PDF文档", type=["pdf"], label_visibility="collapsed")
    if file:
        try:
            with st.spinner("🔍 正在读取PDF内容..."):
                input_text = read_pdf(file.getvalue())
            stats = count_stats(input_text)
            st.success(f"✅ PDF读取成功: {file.name} · {stats['总字符']}字符")
            with st.expander("📋 查看PDF提取内容"):
                st.text(input_text[:2000] + ("..." if len(input_text) > 2000 else ""))
        except Exception as e:
            st.error(f"❌ PDF读取失败: {e}")

else:
    method = st.radio("方式", ["直接输入", "上传文件 (txt/docx)"], horizontal=True, label_visibility="collapsed")

    if method.startswith("直接输入"):
        input_text = st.text_area("文本", placeholder="粘贴需要处理的文本...", height=280,
                                   max_chars=APP_CONFIG["max_input_length"], label_visibility="collapsed")
        # 中英翻译模式下显示语言检测
        if st.session_state.role == "中英互译" and input_text.strip():
            lang_label = get_translation_label(input_text)
            st.markdown(f"<span class='lang-badge'>{lang_label}</span>", unsafe_allow_html=True)
    else:
        file = st.file_uploader("上传", type=["txt", "docx"], label_visibility="collapsed")
        if file:
            try:
                if file.name.endswith(".docx"):
                    source_paras = read_docx_structured(file.getvalue())
                    input_text = "\n\n".join([p["text"] for p in source_paras])
                    st.session_state.is_word_mode = True
                    st.success(f"✅ Word: {file.name} · {len(source_paras)}个段落")
                    # 中英翻译模式下显示语言检测
                    if st.session_state.role == "中英互译":
                        lang_label = get_translation_label(input_text)
                        st.markdown(f"<span class='lang-badge'>{lang_label}</span>", unsafe_allow_html=True)
                else:
                    input_text = file.getvalue().decode("utf-8")
                    st.success(f"✅ 文本: {file.name}")
            except Exception as e:
                st.error(f"❌ 读取失败: {e}")

if input_text.strip() and not st.session_state.is_pdf_mode:
    stats = count_stats(input_text)
    cols = st.columns(4)
    for c, (k, v) in zip(cols, stats.items()):
        c.markdown(f"<div class='stat-box'><div class='stat-num'>{v}</div><div class='stat-label'>{k}</div></div>", unsafe_allow_html=True)

# 开始/停止按钮
col_btn1, col_btn2 = st.columns([1, 1])
with col_btn1:
    process = st.button("🚀 开始处理", type="primary", use_container_width=True,
                        disabled=not (api_key and input_text.strip()) or st.session_state.processing)
with col_btn2:
    if st.session_state.processing:
        if st.button("⏹️ 停止处理", type="secondary", use_container_width=True):
            st.session_state.stop_flag = True
            st.session_state.processing = False
            st.rerun()

st.markdown("---")

# ============================================================
# 处理逻辑
# ============================================================

st.subheader("📤 处理结果")

result_placeholder = st.empty()

if process and api_key and input_text.strip():
    try:
        client = LLMClient(provider=provider, api_key=api_key, model=model)

        # ========== PDF摘要模式 ==========
        if st.session_state.is_pdf_mode:
            st.session_state.stop_flag = False
            st.session_state.processing = True
            chunks = split_text_to_chunks(input_text)
            st.info(f"📑 文档已分割为 {len(chunks)} 个片段")
            progress = st.progress(0)
            status = st.empty()

            chunk_summaries = []
            for i, chunk in enumerate(chunks):
                if st.session_state.stop_flag:
                    status.warning("⏹️ 处理已停止")
                    break
                status.info(f"🤖 正在处理第 {i+1}/{len(chunks)} 个片段...")
                try:
                    summary = summarize_chunk(client, chunk, i+1, len(chunks))
                    chunk_summaries.append(summary)
                except Exception as e:
                    chunk_summaries.append(f"[第{i+1}片段失败: {e}]")
                progress.progress((i + 1) / len(chunks))

            if chunk_summaries and not st.session_state.stop_flag:
                status.info("📝 正在生成最终摘要...")
                merged = "\n\n".join(chunk_summaries)
                final_summary = generate_final_summary(client, merged)
                st.session_state.result = final_summary
            elif chunk_summaries:
                st.session_state.result = "\n\n".join(chunk_summaries)

            st.session_state.count += 1
            st.session_state.processing = False
            progress.empty()
            status.empty()

            if st.session_state.stop_flag:
                st.warning(f"⏹️ 已停止，完成 {len(chunk_summaries)}/{len(chunks)} 个片段")
            else:
                st.success(f"✅ PDF摘要完成！处理了 {len(chunks)} 个片段")

        # ========== Word润色模式 ==========
        elif st.session_state.is_word_mode and source_paras:
            merged_paras = merge_short_paragraphs(source_paras)
            st.info(f"📊 原文 {len(source_paras)} 个段落 → 合并为 {len(merged_paras)} 个处理块")
            st.session_state.stop_flag = False
            st.session_state.processing = True

            progress = st.progress(0)
            status = st.empty()

            # 获取提示词（中英翻译特殊处理）
            role_info = POLISH_ROLES[st.session_state.role]
            if st.session_state.role == "中英互译":
                base_prompt = get_translation_prompt(role_info, input_text)
            else:
                base_prompt = role_info["prompt"]

            polished_paras = []
            for i, para in enumerate(merged_paras):
                if st.session_state.stop_flag:
                    status.warning("⏹️ 处理已停止")
                    break
                status.info(f"🤖 正在处理第 {i+1}/{len(merged_paras)} 块...")
                if para["text"].strip():
                    try:
                        polished = client.process_sync(para["text"], base_prompt)
                        polished_paras.append({"text": polished.strip(), "style": para["style"]})
                    except Exception as e:
                        st.warning(f"第{i+1}块失败: {e}")
                        polished_paras.append(para)
                else:
                    polished_paras.append(para)
                progress.progress((i + 1) / len(merged_paras))

            st.session_state.structured_paras = polished_paras
            st.session_state.result = "\n\n".join([p["text"] for p in polished_paras])
            st.session_state.count += 1
            st.session_state.processing = False
            progress.empty()
            status.empty()

            if st.session_state.stop_flag:
                st.warning(f"⏹️ 已停止，完成 {len(polished_paras)}/{len(merged_paras)} 块")
            else:
                st.success(f"✅ Word润色完成！处理了 {len(merged_paras)} 个块")

        # ========== 普通模式（直接输入） ==========
        else:
            role_info = POLISH_ROLES[st.session_state.role]
            # 中英翻译特殊处理
            if st.session_state.role == "中英互译":
                prompt = get_translation_prompt(role_info, input_text)
                lang_label = get_translation_label(input_text)
                result_placeholder.info(f"🌐 {lang_label} · 处理中...")
            else:
                prompt = role_info["prompt"]
                result_placeholder.info("🤖 AI处理中...")

            full = ""
            t0 = time.time()
            is_code_mode = st.session_state.role == "代码解释"
            style_class = "code-result" if is_code_mode else "result-area"
            for chunk in client.stream_process(input_text, prompt):
                full += chunk
                safe = full.replace("<", "&lt;").replace(">", "&gt;")
                result_placeholder.markdown(f"<div class='{style_class}'>{safe}</div>", unsafe_allow_html=True)
            st.session_state.result = full
            st.session_state.structured_paras = []
            st.session_state.count += 1
            st.success(f"✅ 完成！耗时 {time.time()-t0:.1f}秒")

    except Exception as e:
        st.session_state.processing = False
        result_placeholder.error(f"❌ 失败: {e}")

elif st.session_state.result:
    safe = st.session_state.result.replace("<", "&lt;").replace(">", "&gt;")
    style_class = "code-result" if st.session_state.role == "代码解释" else "result-area"
    result_placeholder.markdown(f"<div class='{style_class}'>{safe}</div>", unsafe_allow_html=True)
else:
    result_placeholder.info("👆 输入后点击「开始处理」")

# ============================================================
# 导出区域
# ============================================================

if st.session_state.result:
    st.markdown("---")
    st.subheader("💾 导出结果")

    # Word润色模式：导出完整格式Word
    if st.session_state.is_word_mode and st.session_state.structured_paras:
        try:
            docx_data = save_docx_structured(st.session_state.structured_paras,
                                              title=f"{st.session_state.role}结果")
            st.download_button("📥 下载完整Word文档", docx_data,
                               file_name=f"润色结果_{st.session_state.role}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True)
        except Exception as e:
            st.error(f"Word导出失败: {e}")

    # 其他模式：文本+Word导出
    else:
        c1, c2 = st.columns(2)
        with c1:
            st.text_area("📋 复制文本", st.session_state.result, height=80, label_visibility="collapsed")
            st.caption("👆 Ctrl+A 全选 → Ctrl+C 复制")
        with c2:
            try:
                paras = [{"text": p, "style": "Normal"} for p in st.session_state.result.split("\n\n") if p.strip()]
                docx_data = save_docx_structured(paras, title=f"{st.session_state.role if not st.session_state.is_pdf_mode else 'PDF摘要'}结果")
                st.download_button("📄 导出Word文档", docx_data,
                                   file_name=f"结果_{st.session_state.role if not st.session_state.is_pdf_mode else 'PDF摘要'}.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                   use_container_width=True)
            except:
                txt_data = st.session_state.result.encode("utf-8")
                st.download_button("📝 导出TXT", txt_data, "结果.txt", "text/plain", use_container_width=True)

st.markdown("""
<div style="text-align:center;color:#94a3b8;font-size:0.8rem;margin-top:2rem;padding:1rem;border-top:1px solid #e2e8f0;">
    🎓 东北大学秦皇岛分校 · 计算机与通信工程学院 · 启航科创项目 | v3.0
</div>
""", unsafe_allow_html=True)
