"""
大模型API客户端模块
负责与各大模型API进行通信，实现文本润色和翻译功能

核心功能：
- 支持多平台API（通义千问/DeepSeek/OpenAI）
- 统一的请求格式和错误处理
- 流式输出支持（实时显示处理进度）
- 完整的日志记录
"""

import json
import logging
import time
from typing import Generator, Optional

import requests
import streamlit as st

from config import API_PROVIDERS, APP_CONFIG

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


class LLMClient:
    """
    大语言模型API客户端

    封装了对各大模型API的调用逻辑，提供统一的接口：
    - 文本润色/翻译（非流式）
    - 文本润色/翻译（流式实时输出）
    - 健康检查（验证API Key有效性）

    使用示例：
        client = LLMClient(provider="通义千问", api_key="sk-xxx")
        result = client.polish_text("需要润色的文本", role="严谨学术风")
    """

    def __init__(self, provider: str, api_key: str, model: Optional[str] = None):
        """
        初始化API客户端

        Args:
            provider: API提供商名称（通义千问/DeepSeek/OpenAI）
            api_key: 从对应平台获取的API密钥
            model: 指定使用的模型名称，None则使用默认模型
        """
        if provider not in API_PROVIDERS:
            raise ValueError(f"不支持的API提供商: {provider}，可选: {list(API_PROVIDERS.keys())}")

        self.provider = provider
        self.api_key = api_key
        self.config = API_PROVIDERS[provider]
        self.model = model or self.config["default_model"]
        self.base_url = self.config["base_url"]

        logger.info(f"初始化LLMClient: provider={provider}, model={self.model}")

    def _build_headers(self) -> dict:
        """
        构建HTTP请求头

        Returns:
            包含认证信息的请求头字典
        """
        return {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }

    def _build_payload(self, text: str, system_prompt: str, stream: bool = False) -> dict:
        """
        构建API请求体（符合OpenAI兼容格式）

        Args:
            text: 用户输入的待处理文本
            system_prompt: 角色预设的系统提示词
            stream: 是否启用流式输出

        Returns:
            符合API规范的请求体字典
        """
        return {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": system_prompt,
                },
                {
                    "role": "user",
                    "content": text,
                },
            ],
            "temperature": APP_CONFIG["default_temperature"],
            "stream": stream,
        }

    def polish_text(self, text: str, system_prompt: str) -> str:
        """
        【核心方法】非流式文本润色/翻译

        将用户文本发送给大模型API，获取完整的润色/翻译结果。
        适合短文本和需要完整结果后处理的场景。

        Args:
            text: 待处理的原始文本
            system_prompt: 角色对应的系统提示词（定义在config.py中）

        Returns:
            处理后的完整文本字符串

        Raises:
            ConnectionError: 网络连接失败
            TimeoutError: 请求超时
            ValueError: API返回错误或余额不足
        """
        headers = self._build_headers()
        payload = self._build_payload(text, system_prompt, stream=False)

        logger.info(f"发送请求: provider={self.provider}, model={self.model}, text_length={len(text)}")
        start_time = time.time()

        try:
            response = requests.post(
                self.base_url,
                headers=headers,
                json=payload,
                timeout=APP_CONFIG["timeout_seconds"],
            )
            # 记录原始响应用于调试
            logger.debug(f"API响应状态: {response.status_code}")

        except requests.exceptions.Timeout:
            raise TimeoutError(f"请求超时（{APP_CONFIG['timeout_seconds']}秒），请检查网络连接或稍后重试")
        except requests.exceptions.ConnectionError:
            raise ConnectionError("网络连接失败，请检查网络设置")

        # 解析响应
        if response.status_code != 200:
            error_msg = self._parse_error(response)
            raise ValueError(f"API请求失败 [{response.status_code}]: {error_msg}")

        try:
            result_data = response.json()
            result = result_data["choices"][0]["message"]["content"]
        except (KeyError, IndexError, json.JSONDecodeError) as e:
            raise ValueError(f"解析API响应失败: {str(e)}")

        elapsed = time.time() - start_time
        logger.info(f"请求完成: 耗时={elapsed:.2f}s, output_length={len(result)}")

        return result

    def polish_text_stream(self, text: str, system_prompt: str) -> Generator[str, None, None]:
        """
        【核心方法】流式文本润色/翻译（实时输出）

        逐字/逐句返回处理结果，实现"打字机"效果，
        大幅提升用户体验，让用户感知到AI正在"思考"。

        Args:
            text: 待处理的原始文本
            system_prompt: 角色对应的系统提示词

        Yields:
            逐块返回的处理结果文本片段

        使用示例（Streamlit中）：
            result_placeholder = st.empty()
            full_result = ""
            for chunk in client.polish_text_stream(text, prompt):
                full_result += chunk
                result_placeholder.markdown(full_result)
        """
        headers = self._build_headers()
        payload = self._build_payload(text, system_prompt, stream=True)

        logger.info(f"发送流式请求: provider={self.provider}, text_length={len(text)}")

        try:
            response = requests.post(
                self.base_url,
                headers=headers,
                json=payload,
                timeout=APP_CONFIG["timeout_seconds"],
                stream=True,  # 启用HTTP流式传输
            )
        except requests.exceptions.Timeout:
            raise TimeoutError(f"请求超时（{APP_CONFIG['timeout_seconds']}秒）")
        except requests.exceptions.ConnectionError:
            raise ConnectionError("网络连接失败")

        if response.status_code != 200:
            error_msg = self._parse_error(response)
            raise ValueError(f"API请求失败 [{response.status_code}]: {error_msg}")

        # 逐行解析SSE（Server-Sent Events）格式的流式响应
        for line in response.iter_lines(decode_unicode=True):
            if not line or not line.startswith("data: "):
                continue

            data = line[6:]  # 去掉 "data: " 前缀

            if data == "[DONE]":  # 流式传输结束标志
                break

            try:
                chunk = json.loads(data)
                # 提取增量内容（delta格式，兼容OpenAI标准）
                delta = chunk.get("choices", [{}])[0].get("delta", {})
                content = delta.get("content", "")
                if content:
                    yield content
            except (json.JSONDecodeError, KeyError, IndexError):
                continue  # 忽略解析失败的块，保证流不中断

    def _parse_error(self, response: requests.Response) -> str:
        """
        解析API错误响应，提取可读的错误信息

        Args:
            response: 失败的HTTP响应对象

        Returns:
            用户友好的错误描述字符串
        """
        try:
            error_data = response.json()
            # 兼容不同厂商的错误格式
            if "error" in error_data:
                error = error_data["error"]
                if isinstance(error, dict):
                    return error.get("message", str(error))
                return str(error)
            return str(error_data)
        except:
            return f"未知错误: {response.text[:200]}"

    def health_check(self) -> bool:
        """
        验证API Key是否有效

        发送一个极短的测试请求，用于在界面上显示连接状态。

        Returns:
            API Key有效返回True，否则返回False
        """
        try:
            # 发送一个极简的测试请求
            headers = self._build_headers()
            payload = {
                "model": self.model,
                "messages": [{"role": "user", "content": "Hi"}],
                "max_tokens": 5,  # 限制返回长度，节省Token
            }
            response = requests.post(
                self.base_url,
                headers=headers,
                json=payload,
                timeout=10,
            )
            return response.status_code == 200
        except Exception as e:
            logger.warning(f"健康检查失败: {str(e)}")
            return False


def get_available_models(provider: str) -> list:
    """
    获取指定提供商支持的所有模型列表

    Args:
        provider: API提供商名称

    Returns:
        模型名称字符串列表
    """
    if provider not in API_PROVIDERS:
        return []
    return API_PROVIDERS[provider]["models"]


def format_stream_output(text: str) -> str:
    """
    格式化流式输出文本，确保Markdown渲染正确

    处理流式输出中可能出现的格式问题：
    - 转义HTML特殊字符防止XSS
    - 保持换行格式

    Args:
        text: 原始文本片段

    Returns:
        格式化后的安全文本
    """
    # 转义HTML标签，防止意外渲染
    text = text.replace("<", "&lt;").replace(">", "&gt;")
    return text
