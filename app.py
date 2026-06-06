"""
AI智能文档润色/翻译助手 - 主程序
基于Streamlit的Web应用，集成大模型API实现专业文本处理

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
        "name": "通义千问 (阿里)",
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
        "models": ["qwen-turbo", "qwen-plus", "qwen-max"],
        "default_model": "qwen-turbo",
    },
    "DeepSeek": {
        "name": "DeepSeek",
        "base_url": "https://api.deepseek.com/chat/completions",
        "models": ["deepseek-chat", "deepseek-reasoner"],
        "default_model": "deepseek-chat",
    },
    "OpenAI": {
        "name": "OpenAI",
        "base_url": "https://api.openai.com/v1/chat/completions",
        "models": ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"],
        "default_model": "gpt-4o-mini",
    },
}

ROLE_TEMPLATES = {
    "严谨学术风": {
        "icon": "📚",
        "description": "适合学术论文、毕业论文、期刊投稿的语言润色",
        "system_prompt": """你是一位拥有20年经验的资深学术编辑，精通中英文学术写作规范。

【任务要求】
1. 对输入文本进行深度语言润色，提升学术规范性
2. 优化句子结构，增强逻辑连贯性
3. 修正语法错误、拼写错误和标点使用
4. 使用恰当的学术词汇和表达方式
5. 保持原文的学术观点和核心结论不变

【输出规范】
- 直接返回润色后的完整文本，不要添加解释说明
- 保持原文的段落结构和格式
- 对于专业术语，确保翻译准确、用词统一
- 语气应保持客观、严谨、专业""",
    },
    "硅谷极客风": {
        "icon": "💻",
        "description": "适合技术文档、代码注释、README优化",
        "system_prompt": """你是一位在硅谷顶级科技公司工作的技术文档专家，擅长编写清晰、简洁的技术文档。

【任务要求】
1. 优化技术文档的表达，使其简洁、专业、易读
2. 统一技术术语的使用，符合业界惯例
3. 优化代码注释，使其清晰说明"为什么"而不仅是"做什么"
4. 使用主动语态和祈使句，增强文档的指导性
5. 删除冗余表达，保持技术文档的精炼风格

【输出规范】
- 直接返回优化后的完整文本，不要添加解释说明
- 代码注释应保持与代码意图一致
- 保留所有技术术语和API名称的原始形式
- 遵循Google/Apple技术文档风格指南""",
    },
    "简洁商务风": {
        "icon": "💼",
        "description": "适合商务邮件、工作报告、方案策划",
        "system_prompt": """你是一位资深商务沟通顾问，专精于企业级商务写作。

【任务要求】
1. 将内容改写为简洁、专业的商务表达
2. 突出关键信息和行动项（Action Items）
3. 使用商务场景惯用的礼貌用语和格式
4. 删除口语化表达，提升文本的正式程度
5. 优化信息结构，使重点突出、层次清晰

【输出规范】
- 直接返回改写后的完整文本，不要添加解释说明
- 商务邮件应包含恰当的称呼和落款建议
- 使用数据驱动的表达方式
- 语气应专业、礼貌、有说服力""",
    },
    "中英翻译": {
        "icon": "🌐",
        "description": "专业学术中英互译，保持术语一致和语境准确",
        "system_prompt": """你是一位专业的学术翻译专家，精通中英双语学术写作。

【任务要求】
1. 准确理解原文的学术含义和语境
2. 进行中英或英中的专业翻译
3. 保持专业术语的翻译一致性和准确性
4. 确保翻译后的文本符合目标语言的学术表达习惯
5. 保留原文的语气和风格特征

【输出规范】
- 直接返回翻译后的完整文本，不要添加解释说明
- 对于首次出现的专业术语，可在括号内保留原文
- 保持原文的段落结构和格式对应
- 确保翻译后的文本流畅自然，符合母语者表达习惯""",
    },
    "代码解释": {
        "icon": "🔍",
        "description": "为代码添加详细的中文注释和解释说明",
        "system_prompt": """你是一位经验丰富的技术导师，擅长为代码添加清晰、详尽的中文注释。

【任务要求】
1. 为每一行/每一块关键代码添加中文注释
2. 解释代码的设计思路和实现逻辑
3. 说明函数参数的含义和返回值
4. 标注潜在的性能优化点和注意事项
5. 对于复杂算法，解释其时间/空间复杂度

【输出规范】
- 直接返回带有详细注释的完整代码
- 注释应放在对应代码行的上方或行尾
- 保持代码的原始格式和缩进
- 注释使用中文，代码中的变量名和API保持原样""",
    },
}

APP_CONFIG = {
    "title": "AI智能文档润色/翻译助手",
    "subtitle": "基于大模型API的专业文本处理工具",
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
            raise TimeoutError(f"请求超时（{APP_CONFIG['timeout_seconds']}秒）")
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

    def health_check(self) -> bool:
        try:
            headers = self._build_headers()
            payload = {"model": self.model, "messages": [{"role": "user", "content": "Hi"}], "max_tokens": 5}
            response = requests.post(self.base_url, headers=headers, json=payload, timeout=10)
            return response.status_code == 200
        except Exception:
            return False


# ============================================================
# 文档处理
# ============================================================

def count_text_stats(text: str) -> dict:
    clean_text = text.strip()
    chinese_chars = len(re.findall(r"[\u4e00-\u9fff]", clean_text))
    total_chars = len(clean_text)
    paragraphs = len([p for p in clean_text.split("\n\n") if p.strip()])
    if paragraphs == 0 and clean_text:
        paragraphs = 1
    lines = len([l for l in clean_text.split("\n") if l.strip()])
    return {"总字符数": total_chars, "中文字符": chinese_chars, "段落数": paragraphs, "行数": lines}


def detect_language(text: str) -> str:
    chinese_chars = len(re.findall(r"[\u4e00-\u9fff]", text))
    english_chars = len(re.findall(r"[a-zA-Z]", text))
    total_chars = len(text.strip())
    if total_chars == 0:
        return "unknown"
    zh_ratio = chinese_chars / total_chars
    en_ratio = english_chars / total_chars
    if zh_ratio > 0.5:
        return "zh"
    elif en_ratio > 0.5:
        return "en"
    return "mixed"


def save_txt(text: str) -> bytes:
    return text.encode("utf-8")


# ============================================================
# Streamlit 界面
# ============================================================

st.set_page_config(page_title=APP_CONFIG["title"], page_icon="📝", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.main-title { font-size: 2rem !important; font-weight: 700 !important; color: #1f2937 !important; margin-bottom: 0.5rem !important; }
.subtitle { font-size: 1rem !important; color: #6b7280 !important; margin-bottom: 1.5rem !important; }
.stat-card { background-color: #f3f4f6; border-radius: 8px; padding: 10px 15px; text-align: center; }
.stat-value { font-size: 1.5rem; font-weight: 700; color: #2563eb; }
.stat-label { font-size: 0.8rem; color: #6b7280; }
.result-box { background-color: #f8fafc; border-left: 4px solid #2563eb; padding: 20px; border-radius: 0 8px 8px 0; }
.footer { text-align: center; color: #9ca3af; font-size: 0.85rem; margin-top: 3rem; padding-top: 1rem; border-top: 1px solid #e5e7eb; }
</style>
""", unsafe_allow_html=True)

if "processed_result" not in st.session_state:
    st.session_state.processed_result = ""
if "process_count" not in st.session_state:
    st.session_state.process_count = 0

# 侧边栏
with st.sidebar:
    st.header("⚙️ API配置")
    provider = st.selectbox("选择API提供商", options=list(API_PROVIDERS.keys()), index=0)
    provider_info = API_PROVIDERS[provider]

    default_key = ""
    if "DASHSCOPE_API_KEY" in st.secrets:
        default_key = st.secrets["DASHSCOPE_API_KEY"]
    elif "DASHSCOPE_API_KEY" in os.environ:
        default_key = os.environ["DASHSCOPE_API_KEY"]

    if default_key:
        api_key = default_key
        st.success("✅ API Key已加载")
    else:
        api_key = st.text_input("API Key", type="password", placeholder="在此粘贴您的API Key")

    model = st.selectbox("选择模型", options=provider_info["models"], index=0)

    if api_key and not default_key:
        if st.button("🔌 测试连接", use_container_width=True):
            with st.spinner("测试中..."):
                try:
                    client = LLMClient(provider=provider, api_key=api_key, model=model)
                    if client.health_check():
                        st.success("✅ API连接成功！")
                    else:
                        st.error("❌ API Key无效")
                except Exception as e:
                    st.error(f"❌ 连接失败: {str(e)}")

    st.markdown("---")
    st.header("📊 使用统计")
    st.metric("本次处理次数", st.session_state.process_count)

# 主界面
st.markdown('<p class="main-title">📝 AI智能文档润色/翻译助手</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">基于大模型API的专业文本处理工具</p>', unsafe_allow_html=True)
st.markdown("---")

# 角色选择
st.subheader("🎭 选择处理角色")
role_cols = st.columns(len(ROLE_TEMPLATES))
for idx, (col, (role_key, role_info)) in enumerate(zip(role_cols, ROLE_TEMPLATES.items())):
    with col:
        is_selected = st.session_state.get("selected_role_key") == role_key
        border_color = "#2563eb" if is_selected else "#e5e7eb"
        bg_color = "#eff6ff" if is_selected else "#fafafa"
        st.markdown(f"""
        <div style="border: 2px solid {border_color}; border-radius: 10px; padding: 15px; background-color: {bg_color}; text-align: center;">
            <div style="font-size: 2rem; margin-bottom: 5px;">{role_info['icon']}</div>
            <div style="font-weight: 600; font-size: 1rem;">{role_key}</div>
            <div style="font-size: 0.8rem; color: #6b7280; margin-top: 4px;">{role_info['description']}</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button(f"选择", key=f"btn_{role_key}", use_container_width=True):
            st.session_state.selected_role_key = role_key
            st.rerun()

if "selected_role_key" not in st.session_state:
    st.session_state.selected_role_key = list(ROLE_TEMPLATES.keys())[0]

selected_role = ROLE_TEMPLATES[st.session_state.selected_role_key]
st.info(f"当前角色: {selected_role['icon']} **{st.session_state.selected_role_key}** - {selected_role['description']}")
st.markdown("---")

# 输入和输出区域
input_col, result_col = st.columns(2)

with input_col:
    st.subheader("📥 输入原文")
    input_method = st.radio("输入方式", ["直接输入", "上传文件"], horizontal=True, label_visibility="collapsed")
    input_text = ""

    if input_method == "直接输入":
        input_text = st.text_area("输入文本", placeholder="在此粘贴需要处理的文本...\n\n支持：\n• 学术论文段落\n• 技术文档/代码注释\n• 商务邮件/工作报告\n• 待翻译内容", height=350, max_chars=APP_CONFIG["max_input_length"], label_visibility="collapsed")
    else:
        uploaded_file = st.file_uploader("上传文档", type=["txt"], help="支持 .txt 格式")
        if uploaded_file is not None:
            try:
                input_text = uploaded_file.getvalue().decode("utf-8")
                st.success(f"✅ 已读取: {uploaded_file.name}")
            except Exception as e:
                st.error(f"❌ 文件读取失败: {str(e)}")

    if input_text.strip():
        stats = count_text_stats(input_text)
        stat_cols = st.columns(4)
        for col, (label, value) in zip(stat_cols, stats.items()):
            col.markdown(f"""
            <div class="stat-card">
                <div class="stat-value">{value}</div>
                <div class="stat-label">{label}</div>
            </div>
            """, unsafe_allow_html=True)
        lang = detect_language(input_text)
        lang_display = {"zh": "🇨🇳 中文", "en": "🇬🇧 英文", "mixed": "🌐 混合", "unknown": "❓ 未知"}
        st.caption(f"检测到语言: {lang_display.get(lang, lang)}")

    has_key = bool(api_key and api_key.strip())
    has_text = bool(input_text and input_text.strip())
    process_disabled = not (has_key and has_text)
    help_text = "请先输入API Key" if process_disabled and not has_key else "请输入待处理文本" if process_disabled else "点击开始AI处理"

    process_btn = st.button("🚀 开始处理", type="primary", use_container_width=True, disabled=process_disabled, help=help_text)

with result_col:
    st.subheader("📤 处理结果")
    if process_btn and api_key and input_text.strip():
        result_placeholder = st.empty()
        try:
            client = LLMClient(provider=provider, api_key=api_key, model=model)
            system_prompt = selected_role["system_prompt"]
            result_placeholder.info("🤖 AI正在处理中，请稍候...")
            full_result = ""
            start_time = time.time()
            for chunk in client.polish_text_stream(input_text, system_prompt):
                full_result += chunk
                safe_result = full_result.replace("<", "&lt;").replace(">", "&gt;")
                result_placeholder.markdown(f'<div class="result-box">{safe_result}</div>', unsafe_allow_html=True)
            elapsed = time.time() - start_time
            st.session_state.processed_result = full_result
            st.session_state.process_count += 1
            result_stats = count_text_stats(full_result)
            st.success(f"✅ 处理完成！耗时 {elapsed:.1f}秒 | 输出 {result_stats['总字符数']} 字符")
        except TimeoutError as e:
            result_placeholder.error(f"⏱️ {str(e)}")
        except ConnectionError as e:
            result_placeholder.error(f"🔌 {str(e)}")
        except Exception as e:
            result_placeholder.error(f"❌ 处理失败: {str(e)}")
    elif st.session_state.processed_result:
        safe_result = st.session_state.processed_result.replace("<", "&lt;").replace(">", "&gt;")
        st.markdown(f'<div class="result-box">{safe_result}</div>', unsafe_allow_html=True)
    else:
        st.info("👆 请在左侧输入文本后点击【开始处理】按钮")

st.markdown("---")

# 导出
if st.session_state.processed_result:
    st.subheader("💾 导出结果")
    result_text = st.session_state.processed_result
    txt_data = save_txt(result_text)
    st.download_button("📝 下载TXT", data=txt_data, file_name="润色结果.txt", mime="text/plain")

st.markdown('<p class="footer">🎓 东北大学秦皇岛分校 · 计算机与通信工程学院 · 启航科创项目</p>', unsafe_allow_html=True)
