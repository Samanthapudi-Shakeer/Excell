from __future__ import annotations

import os
import time
from dataclasses import dataclass
from typing import Iterable, List, Protocol

import requests


class Translator(Protocol):
    engine_name: str

    def translate_batch(self, texts: Iterable[str], source_lang: str, target_lang: str) -> List[str]:
        ...


@dataclass
class AzureTranslator:
    endpoint: str
    key: str
    region: str
    timeout_seconds: int = 20
    retries: int = 2
    engine_name: str = "azure"

    def translate_batch(self, texts: Iterable[str], source_lang: str, target_lang: str) -> List[str]:
        text_list = list(texts)
        if not text_list:
            return []

        url = f"{self.endpoint.rstrip('/')}/translate"
        params = {"api-version": "3.0", "from": source_lang, "to": [target_lang]}
        headers = {
            "Ocp-Apim-Subscription-Key": self.key,
            "Ocp-Apim-Subscription-Region": self.region,
            "Content-Type": "application/json",
        }
        body = [{"text": t} for t in text_list]

        last_error = None
        for attempt in range(self.retries + 1):
            try:
                resp = requests.post(url, params=params, headers=headers, json=body, timeout=self.timeout_seconds)
                resp.raise_for_status()
                data = resp.json()
                return [item["translations"][0]["text"] for item in data]
            except Exception as exc:
                last_error = exc
                if attempt < self.retries:
                    time.sleep(1.5 * (attempt + 1))
        raise RuntimeError(f"Azure translation failed after retries: {last_error}")


@dataclass
class OllamaGemmaTranslator:
    model: str = "gemma:2b"
    endpoint: str = "http://localhost:11434/api/generate"
    timeout_seconds: int = 60
    engine_name: str = "ollama_gemma"

    def _prompt(self, text: str, source_lang: str, target_lang: str) -> str:
        return (
            "You are a strict translation engine. Translate exactly from "
            f"{source_lang} to {target_lang}. Preserve punctuation, placeholders, spacing, and formatting. "
            "Do not explain. Do not add notes. Return translation only.\n\n"
            f"Input:\n{text}"
        )

    def translate_batch(self, texts: Iterable[str], source_lang: str, target_lang: str) -> List[str]:
        translated: List[str] = []
        for text in texts:
            payload = {
                "model": self.model,
                "prompt": self._prompt(text, source_lang, target_lang),
                "stream": False,
                "options": {"temperature": 0},
            }
            resp = requests.post(self.endpoint, json=payload, timeout=self.timeout_seconds)
            resp.raise_for_status()
            translated.append(resp.json().get("response", text).strip())
        return translated


class RoutedTranslator:
    """Deterministic routing: azure->fallback local, or local only."""

    def __init__(self, selected_engine: str):
        self.selected_engine = selected_engine

        self.azure = AzureTranslator(
            endpoint=os.getenv("AZURE_TRANSLATOR_ENDPOINT", ""),
            key=os.getenv("AZURE_TRANSLATOR_KEY", ""),
            region=os.getenv("AZURE_TRANSLATOR_REGION", ""),
        )
        self.local = OllamaGemmaTranslator(
            model=os.getenv("OLLAMA_MODEL", "gemma:2b"),
            endpoint=os.getenv("OLLAMA_ENDPOINT", "http://localhost:11434/api/generate"),
        )

    def translate_with_engine(self, text: str, source_lang: str, target_lang: str) -> tuple[str, str]:
        if self.selected_engine == "local":
            return self.local.translate_batch([text], source_lang, target_lang)[0], self.local.engine_name

        try:
            if not (self.azure.endpoint and self.azure.key and self.azure.region):
                raise RuntimeError("Azure credentials are incomplete")
            return self.azure.translate_batch([text], source_lang, target_lang)[0], self.azure.engine_name
        except Exception:
            return self.local.translate_batch([text], source_lang, target_lang)[0], self.local.engine_name
