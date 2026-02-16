from __future__ import annotations

import io
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from typing import Callable, List

A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
C_NS = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"


@dataclass
class XmlTranslationLog:
    object_id: str
    original_text: str
    translated_text: str
    error: str | None = None


def _translate_in_xml(xml_bytes: bytes, translate_func: Callable[[str, str], tuple[str, str]], object_prefix: str) -> tuple[bytes, List[XmlTranslationLog]]:
    logs: List[XmlTranslationLog] = []
    root = ET.fromstring(xml_bytes)

    nodes = list(root.iter(f"{A_NS}t"))
    nodes.extend(root.iter(f"{C_NS}v"))

    for idx, node in enumerate(nodes):
        original = node.text
        if not original or not original.strip():
            continue
        try:
            translated, _engine = translate_func(original, f"{object_prefix}:{idx}")
            node.text = translated
            logs.append(XmlTranslationLog(object_id=f"{object_prefix}:{idx}", original_text=original, translated_text=translated))
        except Exception as exc:
            logs.append(XmlTranslationLog(object_id=f"{object_prefix}:{idx}", original_text=original, translated_text=original, error=str(exc)))

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), logs


def translate_drawings_and_charts(xlsx_bytes: bytes, translate_func: Callable[[str, str], tuple[str, str]]) -> tuple[bytes, List[XmlTranslationLog]]:
    in_mem = io.BytesIO(xlsx_bytes)
    out_mem = io.BytesIO()
    all_logs: List[XmlTranslationLog] = []

    with zipfile.ZipFile(in_mem, "r") as zin, zipfile.ZipFile(out_mem, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            payload = zin.read(info.filename)
            if info.filename.startswith("xl/drawings/") and info.filename.endswith(".xml"):
                payload, logs = _translate_in_xml(payload, translate_func, info.filename)
                all_logs.extend(logs)
            elif info.filename.startswith("xl/charts/") and info.filename.endswith(".xml"):
                payload, logs = _translate_in_xml(payload, translate_func, info.filename)
                all_logs.extend(logs)
            zout.writestr(info, payload)

    out_mem.seek(0)
    return out_mem.read(), all_logs
