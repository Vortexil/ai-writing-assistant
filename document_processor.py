"""
文档处理模块
负责文本的导入导出、格式转换、字数统计等功能

核心功能：
- Word文档(.docx)读取与保存
- 文本文件(.txt)读取与保存
- Markdown格式导出
- 字数统计与文本预处理
"""

import io
import re
from typing import Optional

import streamlit as st

# 尝试导入python-docx，如果未安装则提供友好提示
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


def read_docx(file_bytes: bytes) -> str:
    """
    读取Word文档内容

    从上传的.docx文件中提取纯文本内容，保留段落结构。

    Args:
        file_bytes: 上传文件的字节流

    Returns:
        文档的纯文本内容，段落之间用换行分隔

    Raises:
        ImportError: 未安装python-docx库
        ValueError: 文件格式错误或损坏
    """
    if not DOCX_AVAILABLE:
        raise ImportError(
            "未安装python-docx库。请运行: pip install python-docx"
        )

    try:
        doc = Document(io.BytesIO(file_bytes))
        # 提取所有段落的文本，过滤空段落
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        return "\n\n".join(paragraphs)
    except Exception as e:
        raise ValueError(f"无法解析Word文档: {str(e)}")


def save_docx(text: str, filename: str = "润色结果.docx") -> bytes:
    """
    将文本保存为Word文档

    将处理后的文本保存为格式规范的.docx文件，包含标题和正文。

    Args:
        text: 要保存的文本内容
        filename: 输出文件名（用于Streamlit下载按钮）

    Returns:
        文档的字节流数据，可直接用于st.download_button

    Raises:
        ImportError: 未安装python-docx库
    """
    if not DOCX_AVAILABLE:
        raise ImportError(
            "未安装python-docx库。请运行: pip install python-docx"
        )

    doc = Document()

    # 添加标题
    title = doc.add_heading("AI润色/翻译结果", level=0)
    title.alignment = 1  # 居中对齐

    # 添加分隔线（通过段落底部边框模拟）
    doc.add_paragraph("─" * 40)

    # 按段落添加内容
    paragraphs = text.split("\n\n")
    for para_text in paragraphs:
        if not para_text.strip():
            continue
        # 检测是否为标题行（短文本、无标点、以特定词开头）
        if _is_heading(para_text):
            doc.add_heading(para_text.strip(), level=2)
        else:
            doc.add_paragraph(para_text.strip())

    # 添加页脚信息
    doc.add_paragraph("")  # 空行
    doc.add_paragraph("─" * 40)
    footer = doc.add_paragraph("由 AI智能文档润色助手 生成")
    footer.alignment = 1  # 居中

    # 保存到内存
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    return output.getvalue()


def read_txt(file_bytes: bytes, encoding: Optional[str] = None) -> str:
    """
    读取文本文件内容

    自动检测编码格式（UTF-8/GBK），确保中文文本正确读取。

    Args:
        file_bytes: 上传文件的字节流
        encoding: 指定编码格式，None则自动检测

    Returns:
        文件的文本内容
    """
    if encoding:
        return file_bytes.decode(encoding)

    # 尝试常见编码
    for enc in ["utf-8", "gbk", "gb2312", "utf-16"]:
        try:
            return file_bytes.decode(enc)
        except UnicodeDecodeError:
            continue

    # 如果都失败，使用UTF-8并忽略错误
    return file_bytes.decode("utf-8", errors="ignore")


def save_txt(text: str) -> bytes:
    """
    将文本保存为UTF-8编码的.txt文件

    Args:
        text: 要保存的文本内容

    Returns:
        UTF-8编码的字节流数据
    """
    return text.encode("utf-8")


def save_markdown(text: str, role_name: str = "润色") -> str:
    """
    将结果保存为Markdown格式

    添加Markdown元数据和格式，便于后续编辑和发布。

    Args:
        text: 处理后的文本内容
        role_name: 使用的角色名称（用于元数据）

    Returns:
        格式化的Markdown文本
    """
    from datetime import datetime

    md_content = f"""---
title: AI处理结果
processed_by: {role_name}
date: {datetime.now().strftime("%Y-%m-%d %H:%M")}
---

{text}

---
> 由 AI智能文档润色助手 生成
"""
    return md_content


def count_text_stats(text: str) -> dict:
    """
    统计文本基本信息

    计算字符数、词数、段落数等统计信息，在界面上展示。

    Args:
        text: 输入文本

    Returns:
        包含统计信息的字典
    """
    # 去除空白字符后的纯文本长度
    clean_text = text.strip()

    # 中文字符数（不含标点）
    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', clean_text))

    # 总字符数（含空格）
    total_chars = len(clean_text)

    # 段落数
    paragraphs = len([p for p in clean_text.split("\n\n") if p.strip()])
    if paragraphs == 0 and clean_text:
        paragraphs = 1

    # 总行数
    lines = len([l for l in clean_text.split("\n") if l.strip()])

    return {
        "总字符数": total_chars,
        "中文字符": chinese_chars,
        "段落数": paragraphs,
        "行数": lines,
    }


def _is_heading(text: str) -> bool:
    """
    判断文本是否为标题行

    通过启发式规则判断一行文本是否应作为标题：
    - 长度较短（<30字符）
    - 不以常见标点结尾
    - 以特定关键词开头

    Args:
        text: 待判断的文本行

    Returns:
        是标题返回True，否则返回False
    """
    text = text.strip()
    if len(text) > 30:
        return False
    if text.endswith(("。", "，", "；", "！", "？", ".", ",", ";", "!", "?")):
        return False

    # 常见标题关键词
    heading_keywords = [
        "第", "章", "节", "部分", "引言", "结论", "摘要",
        "背景", "方法", "实验", "结果", "讨论", "参考",
        "前言", "总结", "致谢", "附录",
    ]

    for keyword in heading_keywords:
        if text.startswith(keyword):
            return True

    return False


def truncate_text(text: str, max_length: int, add_ellipsis: bool = True) -> str:
    """
    截断文本到指定长度

    用于预览过长文本，保持阅读体验。

    Args:
        text: 原始文本
        max_length: 最大保留字符数
        add_ellipsis: 截断时是否添加省略号

    Returns:
        截断后的文本
    """
    if len(text) <= max_length:
        return text

    truncated = text[:max_length]
    if add_ellipsis:
        truncated += "..."
    return truncated


def detect_language(text: str) -> str:
    """
    简单检测文本主要语言

    基于中文字符比例判断文本是中文还是英文为主。

    Args:
        text: 待检测文本

    Returns:
        "zh"(中文) 或 "en"(英文) 或 "mixed"(混合)
    """
    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
    english_chars = len(re.findall(r'[a-zA-Z]', text))
    total_chars = len(text.strip())

    if total_chars == 0:
        return "unknown"

    zh_ratio = chinese_chars / total_chars
    en_ratio = english_chars / total_chars

    if zh_ratio > 0.5:
        return "zh"
    elif en_ratio > 0.5:
        return "en"
    else:
        return "mixed"
