"""
AI智能文档润色/翻译助手 - 主程序
基于Streamlit的Web应用，集成大模型API实现专业文本处理

【项目信息】
- 项目名称: AI智能文档润色/翻译助手
- 技术栈: Python + Streamlit + LLM API
- 版本: 1.0.0
- 作者: 启航科创项目

【运行方式】
1. 安装依赖: pip install -r requirements.txt
2. 配置API Key（在侧边栏输入或使用环境变量）
3. 启动应用: streamlit run app.py
4. 浏览器自动打开 http://localhost:8501

【部署方式】
- Streamlit Cloud: 连接GitHub仓库自动部署
- Docker: docker-compose up -d
- 本地运行: streamlit run app.py
"""

import os
import time

import streamlit as st

# 导入项目配置和工具模块
from config import (
    API_PROVIDERS,
    APP_CONFIG,
    ROLE_TEMPLATES,
    UI_TEXT,
)
from utils.api_client import LLMClient
from utils.document_processor import (
    count_text_stats,
    detect_language,
    read_docx,
    read_txt,
    save_docx,
    save_markdown,
    save_txt,
)

# ============================================================
# 页面配置（必须放在最前面，且只能调用一次）
# ============================================================
st.set_page_config(
    page_title=APP_CONFIG["title"],
    page_icon="📝",
    layout="wide",           # 宽屏布局，充分利用屏幕空间
    initial_sidebar_state="expanded",  # 默认展开侧边栏
)

# ============================================================
# 自定义CSS样式（美化界面）
# ============================================================
st.markdown("""
<style>
    /* 主标题样式 */
    .main-title {
        font-size: 2rem !important;
        font-weight: 700 !important;
        color: #1f2937 !important;
        margin-bottom: 0.5rem !important;
    }
    /* 副标题样式 */
    .subtitle {
        font-size: 1rem !important;
        color: #6b7280 !important;
        margin-bottom: 1.5rem !important;
    }
    /* 统计信息卡片 */
    .stat-card {
        background-color: #f3f4f6;
        border-radius: 8px;
        padding: 10px 15px;
        text-align: center;
    }
    .stat-value {
        font-size: 1.5rem;
        font-weight: 700;
        color: #2563eb;
    }
    .stat-label {
        font-size: 0.8rem;
        color: #6b7280;
    }
    /* 角色卡片 */
    .role-card {
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 10px;
        background-color: #fafafa;
    }
    .role-icon {
        font-size: 2rem;
        margin-right: 10px;
    }
    /* 结果区域 */
    .result-box {
        background-color: #f8fafc;
        border-left: 4px solid #2563eb;
        padding: 20px;
        border-radius: 0 8px 8px 0;
    }
    /* 页脚 */
    .footer {
        text-align: center;
        color: #9ca3af;
        font-size: 0.85rem;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #e5e7eb;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 会话状态初始化
# ============================================================
if "processed_result" not in st.session_state:
    st.session_state.processed_result = ""
if "process_count" not in st.session_state:
    st.session_state.process_count = 0
if "api_valid" not in st.session_state:
    st.session_state.api_valid = None


# ============================================================
# 侧边栏 - API配置
# ============================================================
with st.sidebar:
    st.header("⚙️ API配置")

    # --- API提供商选择 ---
    provider = st.selectbox(
        "选择API提供商",
        options=list(API_PROVIDERS.keys()),
        index=0,
        help="通义千问免费额度充足，推荐新手使用",
    )

    # 显示当前提供商信息
    provider_info = API_PROVIDERS[provider]
    st.caption(f"💡 官网: {provider_info['docs_url']}")

    # --- API Key输入（支持环境变量和Streamlit Secrets） ---
    # 优先级: Streamlit Secrets > 环境变量 > 用户手动输入
    default_key = ""
    if "DASHSCOPE_API_KEY" in st.secrets:
        default_key = st.secrets["DASHSCOPE_API_KEY"]
    elif "DASHSCOPE_API_KEY" in os.environ:
        default_key = os.environ["DASHSCOPE_API_KEY"]

    if default_key:
        api_key = default_key
        st.success("✅ API Key已通过Secrets加载")
    else:
        api_key = st.text_input(
            "API Key",
            type="password",
            placeholder="在此粘贴您的API Key",
            help=UI_TEXT["api_key_help"],
        )

    # --- 模型选择 ---
    model = st.selectbox(
        "选择模型",
        options=provider_info["models"],
        index=0,
    )

    # --- API连接测试 ---
    if api_key:
        if st.button("🔌 测试连接", use_container_width=True):
            with st.spinner("测试中..."):
                try:
                    client = LLMClient(provider=provider, api_key=api_key, model=model)
                    if client.health_check():
                        st.session_state.api_valid = True
                        st.success("✅ API连接成功！")
                    else:
                        st.session_state.api_valid = False
                        st.error("❌ API Key无效")
                except Exception as e:
                    st.session_state.api_valid = False
                    st.error(f"❌ 连接失败: {str(e)}")
    else:
        st.info("👆 请输入API Key以开始")

    st.markdown("---")

    # --- 使用统计 ---
    st.header("📊 使用统计")
    st.metric("本次处理次数", st.session_state.process_count)

    st.markdown("---")

    # --- 使用说明 ---
    with st.expander("📖 使用说明"):
        st.markdown("""
        **快速开始：**
        1. 在上方输入API Key
        2. 选择适合的角色模板
        3. 输入或上传待处理文本
        4. 点击【开始处理】
        5. 查看结果并导出

        **支持的输入方式：**
        - 直接粘贴文本
        - 上传 .docx 文件
        - 上传 .txt 文件

        **API Key获取：**
        - [通义千问](https://dashscope.aliyun.com)（推荐，免费）
        - [DeepSeek](https://platform.deepseek.com)
        - [OpenAI](https://platform.openai.com)
        """)


# ============================================================
# 主界面 - 标题区
# ============================================================
st.markdown(f'<p class="main-title">{APP_CONFIG["title"]}</p>', unsafe_allow_html=True)
st.markdown(f'<p class="subtitle">{APP_CONFIG["subtitle"]}</p>', unsafe_allow_html=True)
st.markdown("---")

# ============================================================
# 角色选择区
# ============================================================
st.subheader("🎭 选择处理角色")

# 使用列布局展示角色卡片
role_cols = st.columns(len(ROLE_TEMPLATES))
selected_role = None

for idx, (col, (role_key, role_info)) in enumerate(zip(role_cols, ROLE_TEMPLATES.items())):
    with col:
        # 判断当前角色是否被选中
        is_selected = st.session_state.get("selected_role_key") == role_key
        border_color = "#2563eb" if is_selected else "#e5e7eb"
        bg_color = "#eff6ff" if is_selected else "#fafafa"

        st.markdown(f"""
        <div style="border: 2px solid {border_color}; border-radius: 10px; padding: 15px; background-color: {bg_color}; cursor: pointer;">
            <div style="font-size: 2rem; text-align: center; margin-bottom: 5px;">{role_info['icon']}</div>
            <div style="font-weight: 600; text-align: center; font-size: 1rem; margin-bottom: 5px;">{role_key}</div>
            <div style="font-size: 0.8rem; color: #6b7280; text-align: center;">{role_info['description']}</div>
        </div>
        """, unsafe_allow_html=True)

        if st.button(f"选择", key=f"btn_{role_key}", use_container_width=True):
            st.session_state.selected_role_key = role_key
            st.rerun()

# 获取当前选中的角色
if "selected_role_key" not in st.session_state:
    st.session_state.selected_role_key = list(ROLE_TEMPLATES.keys())[0]

selected_role = ROLE_TEMPLATES[st.session_state.selected_role_key]
st.info(f"当前角色: {selected_role['icon']} **{st.session_state.selected_role_key}** - {selected_role['description']}")

st.markdown("---")

# ============================================================
# 输入区域
# ============================================================
input_col, result_col = st.columns(2)

with input_col:
    st.subheader("📥 输入原文")

    # 输入方式选择
    input_method = st.radio(
        "输入方式",
        ["直接输入", "上传文件"],
        horizontal=True,
        label_visibility="collapsed",
    )

    input_text = ""

    if input_method == "直接输入":
        # 文本输入框
        input_text = st.text_area(
            "输入文本",
            placeholder=UI_TEXT["input_placeholder"],
            height=350,
            max_chars=APP_CONFIG["max_input_length"],
            label_visibility="collapsed",
        )
    else:
        # 文件上传
        uploaded_file = st.file_uploader(
            "上传文档",
            type=["docx", "txt"],
            help="支持 .docx 和 .txt 格式",
        )

        if uploaded_file is not None:
            try:
                file_bytes = uploaded_file.getvalue()

                if uploaded_file.name.endswith(".docx"):
                    input_text = read_docx(file_bytes)
                    st.success(f"✅ 已读取Word文档: {uploaded_file.name}")
                elif uploaded_file.name.endswith(".txt"):
                    input_text = read_txt(file_bytes)
                    st.success(f"✅ 已读取文本文件: {uploaded_file.name}")

                # 显示读取的内容预览
                st.text_area("文件内容预览", input_text, height=250, disabled=True)

            except Exception as e:
                st.error(f"❌ 文件读取失败: {str(e)}")

    # 文本统计信息
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

        # 语言检测
        lang = detect_language(input_text)
        lang_display = {"zh": "🇨🇳 中文", "en": "🇬🇧 英文", "mixed": "🌐 混合", "unknown": "❓ 未知"}
        st.caption(f"检测到语言: {lang_display.get(lang, lang)}")

    # 处理按钮
    has_key = bool(api_key and api_key.strip())
    has_text = bool(input_text and input_text.strip())
    process_disabled = not (has_key and has_text)
    if process_disabled and not has_key:
        help_text = "请先输入API Key"
    elif process_disabled:
        help_text = "请输入待处理文本"
    else:
        help_text = "点击开始AI处理"

    process_btn = st.button(
        "🚀 开始处理",
        type="primary",
        use_container_width=True,
        disabled=process_disabled,
        help=help_text,
    )


# ============================================================
# 结果区域
# ============================================================
with result_col:
    st.subheader("📤 处理结果")

    # 结果展示区域
    if process_btn and api_key and input_text.strip():
        # 创建占位符用于流式输出
        result_placeholder = st.empty()

        try:
            # 初始化API客户端
            client = LLMClient(
                provider=provider,
                api_key=api_key,
                model=model,
            )

            # 获取角色对应的系统提示词
            system_prompt = selected_role["system_prompt"]

            # 显示处理中状态
            result_placeholder.info(UI_TEXT["processing_hint"])

            # 使用流式输出获取结果
            full_result = ""
            start_time = time.time()

            for chunk in client.polish_text_stream(input_text, system_prompt):
                full_result += chunk
                # 实时更新显示（转义HTML防止渲染问题）
                safe_result = full_result.replace("<", "&lt;").replace(">", "&gt;")
                result_placeholder.markdown(
                    f'<div class="result-box">{safe_result}</div>',
                    unsafe_allow_html=True,
                )

            elapsed_time = time.time() - start_time

            # 保存结果到会话状态
            st.session_state.processed_result = full_result
            st.session_state.process_count += 1

            # 显示处理完成信息
            result_stats = count_text_stats(full_result)
            st.success(
                f"✅ 处理完成！耗时 {elapsed_time:.1f}秒 | "
                f"输出 {result_stats['总字符数']} 字符"
            )

        except TimeoutError as e:
            result_placeholder.error(f"⏱️ 请求超时: {str(e)}")
        except ConnectionError as e:
            result_placeholder.error(f"🔌 连接错误: {str(e)}")
        except ValueError as e:
            result_placeholder.error(f"❌ API错误: {str(e)}")
        except Exception as e:
            result_placeholder.error(f"❌ 处理失败: {str(e)}")

    elif st.session_state.processed_result:
        # 显示之前的处理结果
        safe_result = st.session_state.processed_result.replace("<", "&lt;").replace(">", "&gt;")
        st.markdown(
            f'<div class="result-box">{safe_result}</div>',
            unsafe_allow_html=True,
        )

    else:
        # 空状态提示
        st.info(UI_TEXT["empty_input_hint"])

st.markdown("---")

# ============================================================
# 导出功能区
# ============================================================
if st.session_state.processed_result:
    st.subheader("💾 导出结果")

    export_col1, export_col2, export_col3, export_col4 = st.columns(4)

    result_text = st.session_state.processed_result

    with export_col1:
        # 复制到剪贴板（通过文本框实现）
        st.text_area("复制文本", result_text, height=100, label_visibility="collapsed")
        st.caption("👆 全选复制即可")

    with export_col2:
        # 导出为Word
        try:
            docx_data = save_docx(result_text)
            st.download_button(
                label="📄 下载Word",
                data=docx_data,
                file_name="润色结果.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        except ImportError:
            st.button("📄 Word导出(需安装python-docx)", disabled=True, use_container_width=True)

    with export_col3:
        # 导出为TXT
        txt_data = save_txt(result_text)
        st.download_button(
            label="📝 下载TXT",
            data=txt_data,
            file_name="润色结果.txt",
            mime="text/plain",
            use_container_width=True,
        )

    with export_col4:
        # 导出为Markdown
        md_content = save_markdown(result_text, st.session_state.selected_role_key)
        st.download_button(
            label="🟦 下载Markdown",
            data=md_content,
            file_name="润色结果.md",
            mime="text/markdown",
            use_container_width=True,
        )

st.markdown("---")

# ============================================================
# 页脚
# ============================================================
st.markdown(f'<p class="footer">{UI_TEXT["footer"]} | v{APP_CONFIG["version"]}</p>', unsafe_allow_html=True)
