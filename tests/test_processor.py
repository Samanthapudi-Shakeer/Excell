from __future__ import annotations

import io

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Font

from excel_translator.processor import process_excel_file


def _sample_workbook_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws["A1"] = "Hello"
    ws["A2"] = 10
    ws["B2"] = 20
    ws["C2"] = "=A2+B2"
    ws["A3"] = "Merged text"
    ws["D4"] = "Far text"
    ws.merge_cells("A3:B3")
    ws["A1"].font = Font(bold=True, color="00FF0000")
    ws["A1"].comment = Comment("Review this", "qa")

    chart = BarChart()
    chart.title = "Quarterly Revenue"
    chart.y_axis.title = "Amount"
    chart.x_axis.title = "Quarter"
    data = Reference(ws, min_col=1, min_row=2, max_col=2, max_row=2)
    chart.add_data(data, titles_from_data=False)
    ws.add_chart(chart, "E2")

    ws2 = wb.create_sheet("Ops")
    ws2["A1"] = "World"

    buf = io.BytesIO()
    wb.save(buf)
    wb.close()
    return buf.getvalue()


def test_translation_preserves_formula_and_formatting(monkeypatch):
    from excel_translator import processor

    monkeypatch.setattr(
        processor.RoutedTranslator,
        "translate_with_engine",
        lambda self, text, source, target: (f"T[{text}]", "fake_engine"),
    )

    result = process_excel_file(
        file_name="input.xlsx",
        file_bytes=_sample_workbook_bytes(),
        source_lang="en",
        target_lang="fr",
        selected_engine="azure",
    )

    wb = load_workbook(io.BytesIO(result.output_bytes))
    ws = wb[wb.sheetnames[0]]

    assert ws["A1"].value == "T[Hello]"
    assert ws["D4"].value == "T[Far text]"
    assert ws["C2"].value == "=A2+B2"
    assert ws["A1"].font.bold is True
    assert ws["A1"].font.color.rgb == "00FF0000"
    assert "A3:B3" in [str(rng) for rng in ws.merged_cells.ranges]
    assert ws["A1"].comment.text == "T[Review this]"

    assert len(wb.sheetnames) == 2
    assert result.output_filename == "input_fr.xlsx"
    assert any(log.object_id == "sheet_title" for log in result.logs)

    wb.close()
