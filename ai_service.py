import os
import json
import re
from typing import List, Optional, Dict

from dotenv import load_dotenv
import requests

load_dotenv()


def get_ai_config(endpoint: Optional[str] = None, api_key: Optional[str] = None, model: Optional[str] = None) -> Optional[Dict[str, str]]:
    endpoint_value = (endpoint or os.getenv("SILICONFLOW_API_URL", "https://api.siliconflow.cn/v1/chat/completions")).strip()
    api_key_value = (api_key or os.getenv("SILICONFLOW_API_KEY", "")).strip()
    model_value = (model or os.getenv("SILICONFLOW_MODEL", "Qwen/Qwen2.5-72B-Instruct")).strip()
    if not endpoint_value or not api_key_value:
        return None
    return {"endpoint": endpoint_value, "api_key": api_key_value, "model": model_value}


def compact_json(raw_text: str) -> Optional[dict]:
    try:
        return json.loads(raw_text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", raw_text, re.S)
        if not match:
            return None
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            return None


def _extract_message_content(content) -> str:
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for item in content:
            if isinstance(item, dict):
                text = item.get("text") or item.get("content") or ""
                if isinstance(text, str):
                    parts.append(text)
                elif isinstance(text, list):
                    parts.append("".join(str(x) for x in text))
            elif item is not None:
                parts.append(str(item))
        return "".join(parts)
    if isinstance(content, dict):
        return str(content.get("content") or content.get("text") or "")
    return str(content or "")


def parse_ai_sections(raw_text: str, endpoint: Optional[str] = None, api_key: Optional[str] = None, model: Optional[str] = None) -> Optional[List[Dict[str, str]]]:
    config = get_ai_config(endpoint=endpoint, api_key=api_key, model=model)
    if config is None:
        return None

    prompt = (
        "你是学术排版助手。请将以下文章内容按照章节进行切分，保留中英文标题。"
        " 仅返回严格 JSON 格式，格式如下：\n"
        "{\"sections\":[{\"heading\":\"摘要\",\"content\":\"正文内容\"},{\"heading\":\"Introduction\",\"content\":\"正文内容\"}]}\n"
        "如果文章没有明确标题，请基于文本语义生成合理章节标题。"
    )
    messages = [
        {"role": "system", "content": "你是一个用于学术稿件排版的助手，擅长将文章文本切分为章节。"},
        {"role": "user", "content": prompt + "\n\n文章内容：\n" + raw_text},
    ]
    payload = {
        "model": config["model"],
        "messages": messages,
        "temperature": 0.0,
        "max_tokens": 1500,
    }
    headers = {
        "Authorization": f"Bearer {config['api_key']}",
        "Content-Type": "application/json",
    }
    try:
        response = requests.post(config["endpoint"], json=payload, headers=headers, timeout=60)
        response.raise_for_status()
        data = response.json()
        choices = data.get("choices")
        if not choices:
            return None
        message = choices[0].get("message", {}) or {}
        content = _extract_message_content(message.get("content", ""))
        parsed = compact_json(content)
        if parsed and isinstance(parsed.get("sections"), list):
            sections = []
            for item in parsed["sections"]:
                heading = item.get("heading", "").strip()
                content = item.get("content", "").strip()
                if heading and content:
                    sections.append({"heading": heading, "content": content})
            if sections:
                return sections
    except Exception:
        return None
    return None


def parse_sections_with_ai(raw_text: str, endpoint: Optional[str] = None, api_key: Optional[str] = None, model: Optional[str] = None) -> Optional[List[Dict[str, str]]]:
    return parse_ai_sections(raw_text, endpoint=endpoint, api_key=api_key, model=model)
