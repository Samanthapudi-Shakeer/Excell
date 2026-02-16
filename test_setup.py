from __future__ import annotations

"""Basic translator setup check for Azure and Ollama.

This script performs a single-word translation smoke test for both engines
using the input text "hello". It is intended as an operational connectivity
check for environment setup, not as a correctness benchmark.
"""

import os
import sys


def _load_translators():
    try:
        from excel_translator.translators import AzureTranslator, OllamaGemmaTranslator

        return AzureTranslator, OllamaGemmaTranslator, None
    except Exception as exc:  # dependency/import problem for setup diagnostics
        return None, None, exc


def run_azure(AzureTranslator) -> tuple[bool, str]:
    endpoint = os.getenv("AZURE_TRANSLATOR_ENDPOINT", "")
    key = os.getenv("AZURE_TRANSLATOR_KEY", "")
    region = os.getenv("AZURE_TRANSLATOR_REGION", "")
    if not (endpoint and key and region):
        return False, "Azure env vars missing: AZURE_TRANSLATOR_ENDPOINT / AZURE_TRANSLATOR_KEY / AZURE_TRANSLATOR_REGION"

    translator = AzureTranslator(endpoint=endpoint, key=key, region=region)
    translated = translator.translate_batch(["hello"], "en", "fr")[0]
    return True, translated


def run_ollama(OllamaGemmaTranslator) -> tuple[bool, str]:
    translator = OllamaGemmaTranslator(
        model=os.getenv("OLLAMA_MODEL", "gemma:2b"),
        endpoint=os.getenv("OLLAMA_ENDPOINT", "http://localhost:11434/api/generate"),
    )
    translated = translator.translate_batch(["hello"], "en", "fr")[0]
    return True, translated


def main() -> int:
    failures = 0
    AzureTranslator, OllamaGemmaTranslator, import_error = _load_translators()
    if import_error is not None:
        print(f"[FAIL] Could not import translator classes/dependencies: {import_error}")
        return 1

    print("[SETUP] Running Azure smoke test with text: 'hello'")
    try:
        ok, message = run_azure(AzureTranslator)
        if ok:
            print(f"[OK] Azure translated: {message}")
        else:
            failures += 1
            print(f"[FAIL] Azure: {message}")
    except Exception as exc:
        failures += 1
        print(f"[FAIL] Azure request failed: {exc}")

    print("[SETUP] Running Ollama smoke test with text: 'hello'")
    try:
        ok, message = run_ollama(OllamaGemmaTranslator)
        if ok:
            print(f"[OK] Ollama translated: {message}")
        else:
            failures += 1
            print(f"[FAIL] Ollama: {message}")
    except Exception as exc:
        failures += 1
        print(f"[FAIL] Ollama request failed: {exc}")

    if failures:
        print(f"[RESULT] FAILED ({failures} checks)")
        return 1

    print("[RESULT] PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
