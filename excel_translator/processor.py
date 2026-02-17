from __future__ import annotations

import io
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import List

from .drawing_xml import translate_drawings_and_charts
from .logging_utils import TranslationLogEntry
from .translators import RoutedTranslator

INVALID_SHEET_CHARS = r"[\\/*?:\[\]]"
S_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


@dataclass
class ProcessingResult:
    output_filename: str
    output_bytes: bytes
    logs: List[TranslationLogEntry]


def _safe_sheet_title(name: str, existing: set[str]) -> str:
    cleaned = re.sub(INVALID_SHEET_CHARS, "_", name).strip() or "Sheet"
    cleaned = cleaned[:31]
    base = cleaned
    i = 1
    while cleaned in existing:
        suffix = f"_{i}"
        cleaned = (base[: 31 - len(suffix)] + suffix)[:31]
        i += 1
    return cleaned


def _with_lang_suffix(name: str, target_lang: str) -> str:
    p = Path(name)
    return f"{p.stem}_{target_lang}{p.suffix}"


def _translate_text(translator: RoutedTranslator, text: str, source_lang: str, target_lang: str) -> tuple[str, str]:
    return translator.translate_with_engine(text, source_lang, target_lang)


def _translate_workbook_sheet_names(
    workbook_xml: bytes,
    file_name: str,
    translator: RoutedTranslator,
    source_lang: str,
    target_lang: str,
    logs: List[TranslationLogEntry],
) -> tuple[bytes, dict[str, str]]:
    root = ET.fromstring(workbook_xml)
    existing_titles: set[str] = set()
    rid_to_sheet: dict[str, str] = {}

    for sheet in root.findall(f".//{S_NS}sheet"):
        original_title = sheet.attrib.get("name", "")
        rid = sheet.attrib.get(f"{R_NS}id", "")
        try:
            translated, engine = _translate_text(translator, original_title, source_lang, target_lang)
            safe = _safe_sheet_title(translated, existing_titles)
            sheet.set("name", safe)
            if rid:
                rid_to_sheet[rid] = safe
            existing_titles.add(safe)
            logs.append(
                TranslationLogEntry(
                    file_name=file_name,
                    sheet_name=safe,
                    object_id="sheet_title",
                    original_text=original_title,
                    translated_text=safe,
                    engine=engine,
                    status="ok",
                )
            )
        except Exception as exc:
            safe = original_title
            if rid:
                rid_to_sheet[rid] = safe
            existing_titles.add(safe)
            logs.append(
                TranslationLogEntry(
                    file_name=file_name,
                    sheet_name=safe,
                    object_id="sheet_title",
                    original_text=original_title,
                    translated_text=original_title,
                    engine="none",
                    status="error",
                    error=str(exc),
                )
            )

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), rid_to_sheet


def _workbook_relationships_map(workbook_rels_xml: bytes) -> dict[str, str]:
    root = ET.fromstring(workbook_rels_xml)
    mapping: dict[str, str] = {}
    for rel in root.findall(f".//{PKG_REL_NS}Relationship"):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rid and target:
            mapping[rid] = target.lstrip("/") if target.startswith("/") else f"xl/{target}" if not target.startswith("xl/") else target
    return mapping


def _translate_shared_strings(
    xml_bytes: bytes,
    file_name: str,
    translator: RoutedTranslator,
    source_lang: str,
    target_lang: str,
    logs: List[TranslationLogEntry],
) -> bytes:
    root = ET.fromstring(xml_bytes)
    for idx, node in enumerate(root.iter(f"{S_NS}t")):
        original = node.text
        if not original or not original.strip():
            continue
        try:
            translated, engine = _translate_text(translator, original, source_lang, target_lang)
            node.text = translated
            logs.append(
                TranslationLogEntry(
                    file_name=file_name,
                    sheet_name="<shared-strings>",
                    object_id=f"sharedString:{idx}",
                    original_text=original,
                    translated_text=translated,
                    engine=engine,
                    status="ok",
                )
            )
        except Exception as exc:
            logs.append(
                TranslationLogEntry(
                    file_name=file_name,
                    sheet_name="<shared-strings>",
                    object_id=f"sharedString:{idx}",
                    original_text=original,
                    translated_text=original,
                    engine="none",
                    status="error",
                    error=str(exc),
                )
            )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _translate_sheet_cells(
    xml_bytes: bytes,
    sheet_name: str,
    file_name: str,
    translator: RoutedTranslator,
    source_lang: str,
    target_lang: str,
    logs: List[TranslationLogEntry],
) -> bytes:
    root = ET.fromstring(xml_bytes)

    for cell in root.iter(f"{S_NS}c"):
        coord = cell.attrib.get("r", "?")
        if cell.find(f"{S_NS}f") is not None:
            continue

        inline_t = cell.find(f"{S_NS}is/{S_NS}t")
        if inline_t is not None and inline_t.text and inline_t.text.strip():
            original = inline_t.text
            try:
                translated, engine = _translate_text(translator, original, source_lang, target_lang)
                inline_t.text = translated
                logs.append(TranslationLogEntry(file_name=file_name, sheet_name=sheet_name, object_id=f"cell:{coord}", original_text=original, translated_text=translated, engine=engine, status="ok"))
            except Exception as exc:
                logs.append(TranslationLogEntry(file_name=file_name, sheet_name=sheet_name, object_id=f"cell:{coord}", original_text=original, translated_text=original, engine="none", status="error", error=str(exc)))
            continue

        if cell.attrib.get("t") == "str":
            value_node = cell.find(f"{S_NS}v")
            if value_node is not None and value_node.text and value_node.text.strip():
                original = value_node.text
                try:
                    translated, engine = _translate_text(translator, original, source_lang, target_lang)
                    value_node.text = translated
                    logs.append(TranslationLogEntry(file_name=file_name, sheet_name=sheet_name, object_id=f"cell:{coord}", original_text=original, translated_text=translated, engine=engine, status="ok"))
                except Exception as exc:
                    logs.append(TranslationLogEntry(file_name=file_name, sheet_name=sheet_name, object_id=f"cell:{coord}", original_text=original, translated_text=original, engine="none", status="error", error=str(exc)))

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _translate_comments(
    xml_bytes: bytes,
    sheet_name: str,
    file_name: str,
    translator: RoutedTranslator,
    source_lang: str,
    target_lang: str,
    logs: List[TranslationLogEntry],
) -> bytes:
    root = ET.fromstring(xml_bytes)

    for comment in root.findall(f".//{S_NS}comment"):
        ref = comment.attrib.get("ref", "?")
        text_nodes = list(comment.iter(f"{S_NS}t"))
        for idx, node in enumerate(text_nodes):
            original = node.text
            if not original or not original.strip():
                continue
            try:
                translated, engine = _translate_text(translator, original, source_lang, target_lang)
                node.text = translated
                logs.append(TranslationLogEntry(file_name=file_name, sheet_name=sheet_name, object_id=f"comment:{ref}:{idx}", original_text=original, translated_text=translated, engine=engine, status="ok"))
            except Exception as exc:
                logs.append(TranslationLogEntry(file_name=file_name, sheet_name=sheet_name, object_id=f"comment:{ref}:{idx}", original_text=original, translated_text=original, engine="none", status="error", error=str(exc)))

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def process_excel_file(
    file_name: str,
    file_bytes: bytes,
    source_lang: str,
    target_lang: str,
    selected_engine: str,
) -> ProcessingResult:
    translator = RoutedTranslator(selected_engine=selected_engine)
    logs: List[TranslationLogEntry] = []

    with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as zin:
        parts = {info.filename: zin.read(info.filename) for info in zin.infolist()}

    workbook_path = "xl/workbook.xml"
    workbook_rels_path = "xl/_rels/workbook.xml.rels"

    rid_to_sheet_name: dict[str, str] = {}
    rid_to_target: dict[str, str] = {}

    if workbook_path in parts:
        parts[workbook_path], rid_to_sheet_name = _translate_workbook_sheet_names(
            parts[workbook_path],
            file_name,
            translator,
            source_lang,
            target_lang,
            logs,
        )

    if workbook_rels_path in parts:
        rid_to_target = _workbook_relationships_map(parts[workbook_rels_path])

    for rid, sheet_name in rid_to_sheet_name.items():
        target = rid_to_target.get(rid)
        if target and target in parts and target.startswith("xl/worksheets/"):
            parts[target] = _translate_sheet_cells(parts[target], sheet_name, file_name, translator, source_lang, target_lang, logs)

    for path, payload in list(parts.items()):
        if path == "xl/sharedStrings.xml":
            parts[path] = _translate_shared_strings(payload, file_name, translator, source_lang, target_lang, logs)
        elif path.startswith("xl/comments") and path.endswith(".xml"):
            parts[path] = _translate_comments(payload, "<comments>", file_name, translator, source_lang, target_lang, logs)

    def translate_func(text: str, object_id: str) -> tuple[str, str]:
        translated, engine = translator.translate_with_engine(text, source_lang, target_lang)
        logs.append(
            TranslationLogEntry(
                file_name=file_name,
                sheet_name="<xml-layer>",
                object_id=object_id,
                original_text=text,
                translated_text=translated,
                engine=engine,
                status="ok",
            )
        )
        return translated, engine

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, payload in parts.items():
            zout.writestr(name, payload)

    final_bytes, xml_logs = translate_drawings_and_charts(buf.getvalue(), translate_func)
    for item in xml_logs:
        if item.error:
            logs.append(
                TranslationLogEntry(
                    file_name=file_name,
                    sheet_name="<xml-layer>",
                    object_id=item.object_id,
                    original_text=item.original_text,
                    translated_text=item.translated_text,
                    engine="none",
                    status="error",
                    error=item.error,
                )
            )

    return ProcessingResult(output_filename=_with_lang_suffix(file_name, target_lang), output_bytes=final_bytes, logs=logs)
