"""
AI智能文档润色/翻译助手 v2.0
基于Streamlit + 大模型API
支持: Word润色导出、PDF摘要、现代UI
"""

import io
import json
import logging
import os
import re
import time
from typing import Generator, Optional

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

ROLE_TEMPLATES = {
    "严谨学术风": {
        "icon": "🎓",
        "color": "#3b82f6",
        "desc": "学术论文、期刊投稿润色",
        "prompt": "你是一位资深学术编辑。对文本进行深度润色：提升学术规范性，优化句子结构，修正语法错误，使用恰当学术词汇，保持原意不变。直接返回润色结果，不要解释。",
    },
    "硅谷极客风": {
        "icon": "⚡",
        "color": "#06b6d4",
        "desc": "技术文档、代码注释优化",
        "prompt": "你是一位技术文档专家。优化表达使其简洁专业，统一技术术语，删除冗余，遵循Google技术文档风格。直接返回优化结果，不要解释。",
    },
    "简洁商务风": {
        "icon": "💼",
        "color": "#8b5cf6",
        "desc": "商务邮件、工作报告改写",
        "prompt": "你是一位商务沟通顾问。改写为简洁专业的商务表达，突出关键信息，删除口语化内容，使用数据驱动表达。直接返回改写结果，不要解释。",
    },
    "中英翻译": {
        "icon": "🌏",
        "color": "#10b981",
        "desc": "专业学术中英互译",
        "prompt": "你是一位学术翻译专家。进行中英或英中专业翻译，保持术语一致性，确保符合目标语言学术表达习惯。直接返回翻译结果，不要解释。",
    },
    "代码解释": {
        "icon": "💻",
        "color": "#f59e0b",
        "desc": "为代码添加详细注释",
        "prompt": "你是一位技术导师。为代码添加清晰详尽的中文注释，解释设计思路和实现逻辑，说明函数参数和返回值。直接返回带注释的代码，不要解释。",
    },
    "PDF摘要": {
        "icon": "📄",
        "color": "#ef4444",
        "desc": "PDF文档内容摘要生成",
        "prompt": "你是一位学术摘要专家。请对以下PDF文档内容进行摘要：提取核心观点、研究方法、关键结论。用中文输出，分点列出，每点简明扼要。格式：【核心主题】→【主要观点】→【关键结论】。",
    },
}

APP_CONFIG = {
    "max_input_length": 8000,
    "default_temperature": 0.3,
    "timeout_seconds": 60,
}

# ============================================================
# API客户端
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


# ============================================================
# 文档处理
# ============================================================

def read_docx(file_bytes: bytes) -> list:
    """读取Word文档，返回段落列表（保留结构）"""
    from docx import Document
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = []
    for p in doc.paragraphs:
        if p.text.strip():
            paragraphs.append({"text": p.text.strip(), "style": p.style.name if p.style else "Normal"})
    return paragraphs


def save_docx_from_paragraphs(paragraphs: list, title: str = "润色结果") -> bytes:
    """生成格式化的Word文档"""
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()

    # 设置默认字体
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Microsoft YaHei'
    font.size = Pt(11)

    # 添加标题
    heading = doc.add_heading(title, level=0)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x1e, 0x3a, 0x5f)

    doc.add_paragraph("")

    # 添加处理时间
    from datetime import datetime
    info = doc.add_paragraph(f"处理时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info.runs[0].font.size = Pt(9)
    info.runs[0].font.color.rgb = RGBColor(0x94, 0xa3, 0xb8)

    doc.add_paragraph("")
    doc.add_paragraph("─" * 40)
    doc.add_paragraph("")

    # 按段落添加内容
    for para in paragraphs:
        text = para.get("text", "").strip()
        style_name = para.get("style", "Normal")
        if not text:
            continue

        if "Heading" in style_name or "标题" in style_name:
            level = 1 if "1" in style_name else 2 if "2" in style_name else 1
            doc.add_heading(text, level=level)
        else:
            doc.add_paragraph(text)

    # 添加页脚
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


def read_pdf(file_bytes: bytes) -> str:
    """读取PDF文档文本"""
    try:
        import PyPDF2
        reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
        texts = []
        for page in reader.pages:
            text = page.extract_text()
            if text:
                texts.append(text.strip())
        return "\n\n".join(texts)
    except ImportError:
        raise ImportError("PyPDF2未安装，无法读取PDF文件")


def count_stats(text: str) -> dict:
    clean = text.strip()
    cn = len(re.findall(r"[\u4e00-\u9fff]", clean))
    return {"总字符": len(clean), "中文": cn, "段落": max(len([p for p in clean.split("\n\n") if p.strip()]), 1)}


# ============================================================
# Streamlit 界面
# ============================================================

st.set_page_config(page_title="AI智能文档润色助手", page_icon="📝", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
/* ===== 全局样式 ===== */
.block-container { padding: 1rem 2rem !important; max-width: 100% !important; }
.main > div { padding: 0 !important; }

/* ===== 顶部标题区 ===== */
.hero-section {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 50%, #0d9488 100%);
    border-radius: 16px;
    padding: 2rem 2.5rem;
    margin-bottom: 1.5rem;
    color: white;
    box-shadow: 0 8px 32px rgba(15, 23, 42, 0.3);
}
.hero-title { font-size: 2rem !important; font-weight: 700; margin: 0; letter-spacing: -0.5px; }
.hero-subtitle { font-size: 0.95rem; color: #94a3b8; margin-top: 0.3rem; }
.hero-badge {
    display: inline-block;
    background: rgba(255,255,255,0.15);
    backdrop-filter: blur(10px);
    border-radius: 20px;
    padding: 4px 14px;
    font-size: 0.75rem;
    margin-top: 0.5rem;
    color: #e2e8f0;
}

/* ===== API配置栏 ===== */
.config-bar {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 0.8rem 1.2rem;
    margin-bottom: 1.5rem;
}

/* ===== 角色卡片 ===== */
.role-card {
    border: 2px solid #e2e8f0;
    border-radius: 14px;
    padding: 1rem 0.8rem;
    text-align: center;
    cursor: pointer;
    transition: all 0.3s ease;
    background: white;
}
.role-card:hover { transform: translateY(-3px); box-shadow: 0 8px 24px rgba(0,0,0,0.08); }
.role-card.active { border-color: #3b82f6; background: #eff6ff; box-shadow: 0 4px 16px rgba(59,130,246,0.15); }
.role-icon { font-size: 1.8rem; margin-bottom: 0.3rem; }
.role-name { font-weight: 600; font-size: 0.85rem; color: #1e293b; }
.role-desc { font-size: 0.72rem; color: #94a3b8; margin-top: 2px; }

/* ===== 输入输出区域 ===== */
.input-zone {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 1.2rem;
    height: 100%;
}
.output-zone {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 1.2rem;
    height: 100%;
    min-height: 500px;
}

/* ===== 统计 ===== */
.stat-box {
    background: linear-gradient(135deg, #f0f9ff, #e0f2fe);
    border-radius: 10px;
    padding: 8px 12px;
    text-align: center;
}
.stat-num { font-size: 1.3rem; font-weight: 700; color: #0369a1; }
.stat-label { font-size: 0.7rem; color: #64748b; }

/* ===== 结果展示 ===== */
.result-area {
    background: #f8fafc;
    border-left: 4px solid #0d9488;
    border-radius: 0 10px 10px 0;
    padding: 1.2rem;
    min-height: 380px;
    max-height: 600px;
    overflow-y: auto;
    font-size: 0.95rem;
    line-height: 1.8;
    color: #334155;
}

/* ===== 导出按钮 ===== */
.export-btn {
    background: linear-gradient(135deg, #3b82f6, #0d9488) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.6rem 1.5rem !important;
}

/* ===== 隐藏Streamlit默认元素 ===== */
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none !important; }

/* ===== 侧边栏优化 ===== */
section[data-testid="stSidebar"] { min-width: 280px !important; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 会话状态
# ============================================================

if "result" not in st.session_state:
    st.session_state.result = ""
if "result_paras" not in st.session_state:
    st.session_state.result_paras = []
if "count" not in st.session_state:
    st.session_state.count = 0
if "role" not in st.session_state:
    st.session_state.role = "严谨学术风"

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
    st.caption("📌 新用户送100万Token")

# ============================================================
# Hero 区域
# ============================================================

st.markdown("""
<div class="hero-section">
    <div style="display:flex;justify-content:space-between;align-items:center;">
        <div>
            <div class="hero-title">📝 AI 智能文档润色助手</div>
            <div class="hero-subtitle">基于大模型API的专业文本处理工具 · 支持Word润色导出 / PDF摘要</div>
            <span class="hero-badge">✦ 启航科创项目</span>
        </div>
        <div style="text-align:right;">
            <div style="font-size:2.5rem;font-weight:700;opacity:0.9;">v2.0</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# API Key 检查
# ============================================================

if not api_key:
    st.warning("⚠️ 请在左侧边栏输入API Key，或联系管理员配置Secrets")
    st.stop()

# ============================================================
# 角色选择
# ============================================================

st.subheader("🎭 选择处理角色")
role_cols = st.columns(len(ROLE_TEMPLATES))
for col, (key, info) in zip(role_cols, ROLE_TEMPLATES.items()):
    with col:
        active_class = "active" if st.session_state.role == key else ""
        st.markdown(f"""
        <div class="role-card {active_class}" style="border-color:{info['color']}40;">
            <div class="role-icon">{info['icon']}</div>
            <div class="role-name">{key}</div>
            <div class="role-desc">{info['desc']}</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("选择", key=f"sel_{key}", use_container_width=True):
            st.session_state.role = key
            st.rerun()

info = ROLE_TEMPLATES[st.session_state.role]
st.info(f"当前模式: {info['icon']} **{st.session_state.role}** — {info['desc']}")

st.markdown("---")

# ============================================================
# 输入区域
# ============================================================

st.subheader("📥 输入原文")

# 判断是否是PDF摘要模式
is_pdf_mode = st.session_state.role == "PDF摘要"

upload_types = ["txt", "docx"]
if is_pdf_mode:
    upload_types = ["pdf"]

method = st.radio("输入方式", ["直接输入", f"上传文件 ({'/'.join(upload_types)})"],
                  horizontal=True, label_visibility="collapsed")

input_text = ""
file_uploaded = False

if method.startswith("直接输入"):
    placeholder_text = "粘贴需要处理的文本...\n\n支持学术论文、技术文档、商务邮件、代码注释等"
    if is_pdf_mode:
        placeholder_text = "⚠️ PDF摘要模式请使用右侧上传功能"
    input_text = st.text_area("文本", placeholder=placeholder_text, height=280,
                               max_chars=APP_CONFIG["max_input_length"], label_visibility="collapsed")
else:
    file = st.file_uploader("上传文件", type=upload_types, label_visibility="collapsed")
    if file:
        try:
            if file.name.endswith(".docx"):
                paras = read_docx(file.getvalue())
                input_text = "\n\n".join([p["text"] for p in paras])
                st.session_state.upload_paras = paras
                file_uploaded = True
                st.success(f"✅ 已读取Word文档: {file.name} · {len(paras)}个段落")
            elif file.name.endswith(".pdf"):
                input_text = read_pdf(file.getvalue())
                st.success(f"✅ 已读取PDF: {file.name}")
            else:
                input_text = file.getvalue().decode("utf-8")
                st.success(f"✅ 已读取: {file.name}")
        except Exception as e:
            st.error(f"❌ 读取失败: {e}")

if input_text.strip():
    stats = count_stats(input_text)
    stat_cols = st.columns(4)
    for c, (k, v) in zip(stat_cols, stats.items()):
        c.markdown(f"""
        <div class="stat-box">
            <div class="stat-num">{v}</div>
            <div class="stat-label">{k}</div>
        </div>""", unsafe_allow_html=True)

# 处理按钮
col_btn, col_info = st.columns([1, 4])
with col_btn:
    process = st.button("🚀 开始处理", type="primary", use_container_width=True,
                        disabled=not (api_key and input_text.strip()))

st.markdown("---")

# ============================================================
# 结果展示（大区域）
# ============================================================

st.subheader("📤 处理结果")

result_placeholder = st.empty()

if process and api_key and input_text.strip():
    try:
        client = LLMClient(provider=provider, api_key=api_key, model=model)
        prompt = ROLE_TEMPLATES[st.session_state.role]["prompt"]

        result_placeholder.info("🤖 AI正在处理中，请稍候...")

        full_result = ""
        t0 = time.time()

        for chunk in client.stream_process(input_text, prompt):
            full_result += chunk
            safe = full_result.replace("<", "&lt;").replace(">", "&gt;")
            result_placeholder.markdown(f"""
            <div class="result-area">{safe}</div>
            """, unsafe_allow_html=True)

        elapsed = time.time() - t0
        st.session_state.result = full_result
        st.session_state.count += 1

        # 如果是Word模式，保存段落结构
        if file_uploaded and st.session_state.role != "PDF摘要":
            st.session_state.result_paras = [{"text": p, "style": "Normal"}
                                              for p in full_result.split("\n\n") if p.strip()]

        st.success(f"✅ 处理完成！耗时 {elapsed:.1f}秒")

    except Exception as e:
        result_placeholder.error(f"❌ 处理失败: {e}")

elif st.session_state.result:
    safe = st.session_state.result.replace("<", "&lt;").replace(">", "&gt;")
    result_placeholder.markdown(f"""
    <div class="result-area">{safe}</div>
    """, unsafe_allow_html=True)
else:
    result_placeholder.info("👆 输入文本后点击「开始处理」按钮，结果将在这里显示")

# ============================================================
# 导出区域
# ============================================================

if st.session_state.result:
    st.markdown("---")
    st.subheader("💾 导出结果")

    exp_col1, exp_col2, exp_col3 = st.columns(3)

    with exp_col1:
        st.text_area("📋 复制文本", st.session_state.result, height=80, label_visibility="collapsed")
        st.caption("👆 Ctrl+A 全选 → Ctrl+C 复制")

    with exp_col2:
        # 导出Word（保留格式）
        try:
            if st.session_state.result_paras:
                paras = st.session_state.result_paras
            else:
                paras = [{"text": p, "style": "Normal"}
                         for p in st.session_state.result.split("\n\n") if p.strip()]

            docx_data = save_docx_from_paragraphs(paras, title=f"{st.session_state.role}结果")
            st.download_button("📄 导出Word文档", docx_data,
                               file_name=f"润色结果_{st.session_state.role}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True)
        except Exception as e:
            st.error(f"Word导出失败: {e}")

    with exp_col3:
        # 导出TXT
        txt_data = st.session_state.result.encode("utf-8")
        st.download_button("📝 导出TXT文本", txt_data,
                           file_name=f"润色结果_{st.session_state.role}.txt",
                           mime="text/plain", use_container_width=True)

# ============================================================
# 页脚
# ============================================================

st.markdown("""
<div style="text-align:center;color:#94a3b8;font-size:0.8rem;margin-top:2rem;padding:1rem;border-top:1px solid #e2e8f0;">
    🎓 东北大学秦皇岛分校 · 计算机与通信工程学院 · 启航科创项目 | v2.0
</div>
""", unsafe_allow_html=True)
