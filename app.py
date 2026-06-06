"""
AI智能文档润色/翻译助手
基于Streamlit + 大模型API
部署: streamlit run app.py
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
    "OpenAI": {
        "base_url": "https://api.openai.com/v1/chat/completions",
        "models": ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"],
        "default_model": "gpt-4o-mini",
    },
}

ROLE_TEMPLATES = {
    "严谨学术风": {
        "icon": "📚",
        "description": "学术论文润色",
        "system_prompt": """你是一位资深学术编辑。对输入文本进行深度语言润色，提升学术规范性，优化句子结构，修正语法错误，使用恰当的学术词汇，保持原文核心结论不变。直接返回润色后的完整文本，不要添加解释。""",
    },
    "硅谷极客风": {
        "icon": "💻",
        "description": "技术文档优化",
        "system_prompt": """你是一位技术文档专家。优化技术文档表达，使其简洁专业，统一技术术语，删除冗余表达，遵循Google技术文档风格。直接返回优化后的完整文本，不要添加解释。""",
    },
    "简洁商务风": {
        "icon": "💼",
        "description": "商务报告改写",
        "system_prompt": """你是一位商务沟通顾问。将内容改写为简洁专业的商务表达，突出关键信息，删除口语化表达，使用数据驱动的方式。直接返回改写后的完整文本，不要添加解释。""",
    },
    "中英翻译": {
        "icon": "🌐",
        "description": "专业学术翻译",
        "system_prompt": """你是一位学术翻译专家。进行中英或英中的专业翻译，保持专业术语一致性，确保翻译后的文本符合目标语言学术表达习惯。直接返回翻译后的完整文本，不要添加解释。""",
    },
    "代码解释": {
        "icon": "🔍",
        "description": "代码注释生成",
        "system_prompt": """你是一位技术导师。为代码添加清晰详尽的中文注释，解释设计思路和实现逻辑，说明函数参数和返回值。直接返回带注释的完整代码，不要添加解释。""",
    },
}

APP_CONFIG = {
    "max_input_length": 5000,
    "default_temperature": 0.3,
    "timeout_seconds": 60,
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

    def _build_headers(self) -> dict:
        return {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}

    def _build_payload(self, text: str, system_prompt: str, stream: bool = False) -> dict:
        return {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": text},
            ],
            "temperature": APP_CONFIG["default_temperature"],
            "stream": stream,
        }

    def polish_text_stream(self, text: str, system_prompt: str) -> Generator[str, None, None]:
        headers = self._build_headers()
        payload = self._build_payload(text, system_prompt, stream=True)
        try:
            response = requests.post(
                self.base_url, headers=headers, json=payload,
                timeout=APP_CONFIG["timeout_seconds"], stream=True,
            )
        except requests.exceptions.Timeout:
            raise TimeoutError(f"请求超时")
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
                delta = chunk.get("choices", [{}])[0].get("delta", {})
                content = delta.get("content", "")
                if content:
                    yield content
            except (json.JSONDecodeError, KeyError, IndexError):
                continue


# ============================================================
# 工具函数
# ============================================================

def count_text_stats(text: str) -> dict:
    clean = text.strip()
    cn = len(re.findall(r"[\u4e00-\u9fff]", clean))
    total = len(clean)
    paras = len([p for p in clean.split("\n\n") if p.strip()]) or 1
    lines = len([l for l in clean.split("\n") if l.strip()])
    return {"总字符数": total, "中文字符": cn, "段落数": paras, "行数": lines}


def read_docx(file_bytes: bytes) -> str:
    from docx import Document
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n\n".join(paragraphs)


def save_docx(text: str) -> bytes:
    from docx import Document
    doc = Document()
    doc.add_heading("AI润色结果", 0)
    doc.add_paragraph("─" * 40)
    for para in text.split("\n\n"):
        if para.strip():
            doc.add_paragraph(para.strip())
    doc.add_paragraph("")
    doc.add_paragraph("─" * 40)
    footer = doc.add_paragraph("由 AI智能文档润色助手 生成")
    footer.alignment = 1
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


# ============================================================
# Streamlit 界面
# ============================================================

st.set_page_config(page_title="AI智能文档润色助手", page_icon="📝", layout="wide")

if "result" not in st.session_state:
    st.session_state.result = ""
if "count" not in st.session_state:
    st.session_state.count = 0
if "role" not in st.session_state:
    st.session_state.role = "严谨学术风"

# CSS
st.markdown("""
<style>
.stat-card { background: #f3f4f6; border-radius: 8px; padding: 10px; text-align: center; }
.stat-value { font-size: 1.5rem; font-weight: 700; color: #2563eb; }
.stat-label { font-size: 0.8rem; color: #6b7280; }
.result-box { background: #f8fafc; border-left: 4px solid #2563eb; padding: 20px; border-radius: 0 8px 8px 0; }
</style>
""", unsafe_allow_html=True)

# 侧边栏
with st.sidebar:
    st.header("⚙️ API配置")
    provider = st.selectbox("选择API提供商", list(API_PROVIDERS.keys()), index=0)

    default_key = ""
    if "DASHSCOPE_API_KEY" in st.secrets:
        default_key = st.secrets["DASHSCOPE_API_KEY"]
    elif "DASHSCOPE_API_KEY" in os.environ:
        default_key = os.environ["DASHSCOPE_API_KEY"]

    if default_key:
        api_key = default_key
        st.success("✅ API Key已加载")
    else:
        api_key = st.text_input("API Key", type="password", placeholder="粘贴API Key")

    model = st.selectbox("选择模型", API_PROVIDERS[provider]["models"], index=0)
    st.markdown("---")
    st.metric("处理次数", st.session_state.count)

# 标题
st.title("📝 AI智能文档润色/翻译助手")
st.caption("基于大模型API的专业文本处理工具")
st.markdown("---")

# 角色选择
st.subheader("🎭 选择处理角色")
cols = st.columns(len(ROLE_TEMPLATES))
for col, (key, info) in zip(cols, ROLE_TEMPLATES.items()):
    with col:
        selected = st.session_state.role == key
        border = "#2563eb" if selected else "#e5e7eb"
        bg = "#eff6ff" if selected else "#fafafa"
        st.markdown(f"""
        <div style="border:2px solid {border};border-radius:10px;padding:15px;background:{bg};text-align:center;">
            <div style="font-size:2rem">{info['icon']}</div>
            <div style="font-weight:600">{key}</div>
            <div style="font-size:0.8rem;color:#6b7280">{info['description']}</div>
        </div>""", unsafe_allow_html=True)
        if st.button("选择", key=f"btn_{key}", use_container_width=True):
            st.session_state.role = key
            st.rerun()

info = ROLE_TEMPLATES[st.session_state.role]
st.info(f"当前角色: {info['icon']} **{st.session_state.role}** - {info['description']}")
st.markdown("---")

# 输入输出
left, right = st.columns(2)

with left:
    st.subheader("📥 输入原文")
    method = st.radio("方式", ["直接输入", "上传文件"], horizontal=True, label_visibility="collapsed")
    input_text = ""

    if method == "直接输入":
        input_text = st.text_area("文本", placeholder="粘贴需要处理的文本...", height=300,
                                   max_chars=APP_CONFIG["max_input_length"], label_visibility="collapsed")
    else:
        file = st.file_uploader("上传", type=["txt", "docx"], label_visibility="collapsed")
        if file:
            try:
                if file.name.endswith(".docx"):
                    input_text = read_docx(file.getvalue())
                    st.success(f"✅ 已读取Word: {file.name}")
                else:
                    input_text = file.getvalue().decode("utf-8")
                    st.success(f"✅ 已读取: {file.name}")
            except Exception as e:
                st.error(f"❌ 读取失败: {e}")

    if input_text.strip():
        stats = count_text_stats(input_text)
        sc = st.columns(4)
        for c, (k, v) in zip(sc, stats.items()):
            c.markdown(f"<div class='stat-card'><div class='stat-value'>{v}</div><div class='stat-label'>{k}</div></div>", unsafe_allow_html=True)

    has_key = bool(api_key and api_key.strip())
    disabled = not (has_key and input_text.strip())
    st.button("🚀 开始处理", type="primary", use_container_width=True, disabled=disabled,
              on_click=lambda: None, key="process_btn")

# 处理逻辑
if st.session_state.get("process_btn") and has_key and input_text.strip():
    with right:
        st.subheader("📤 处理结果")
        placeholder = st.empty()
        try:
            client = LLMClient(provider=provider, api_key=api_key, model=model)
            prompt = ROLE_TEMPLATES[st.session_state.role]["system_prompt"]
            placeholder.info("🤖 AI处理中...")
            full = ""
            t0 = time.time()
            for chunk in client.polish_text_stream(input_text, prompt):
                full += chunk
                safe = full.replace("<", "&lt;").replace(">", "&gt;")
                placeholder.markdown(f"<div class='result-box'>{safe}</div>", unsafe_allow_html=True)
            st.session_state.result = full
            st.session_state.count += 1
            st.success(f"✅ 完成！耗时 {time.time()-t0:.1f}秒")
        except Exception as e:
            placeholder.error(f"❌ {e}")

with right:
    if not st.session_state.get("process_btn") and st.session_state.result:
        st.subheader("📤 处理结果")
        safe = st.session_state.result.replace("<", "&lt;").replace(">", "&gt;")
        st.markdown(f"<div class='result-box'>{safe}</div>", unsafe_allow_html=True)
    elif not st.session_state.get("process_btn"):
        st.subheader("📤 处理结果")
        st.info("👆 输入文本后点击开始处理")

# 导出
if st.session_state.result:
    st.markdown("---")
    st.subheader("💾 导出结果")
    c1, c2 = st.columns(2)
    with c1:
        st.text_area("复制文本", st.session_state.result, height=100, label_visibility="collapsed")
        st.caption("👆 全选复制")
    with c2:
        try:
            docx_data = save_docx(st.session_state.result)
            st.download_button("📄 下载Word", docx_data, "润色结果.docx",
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except:
            txt_data = st.session_state.result.encode("utf-8")
            st.download_button("📝 下载TXT", txt_data, "润色结果.txt", "text/plain")

st.markdown("---")
st.caption("🎓 东北大学秦皇岛分校 · 计算机与通信工程学院 · 启航科创项目")
