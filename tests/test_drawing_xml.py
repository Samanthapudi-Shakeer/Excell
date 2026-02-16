from __future__ import annotations

from excel_translator.drawing_xml import _translate_in_xml


def test_translate_in_xml_only_translates_drawing_text_nodes():
    xml = b'''<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:p><a:r><a:t>Hello Title</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:ser>
      <c:val><c:numRef><c:f>Sheet1!$A$1:$A$2</c:f><c:numCache><c:pt idx="0"><c:v>123</c:v></c:pt></c:numCache></c:numRef></c:val>
    </c:ser>
  </c:chart>
</c:chartSpace>'''

    def fake_translate(text: str, object_id: str) -> tuple[str, str]:
        return f"T[{text}]", "fake"

    out, logs = _translate_in_xml(xml, fake_translate, "xl/charts/chart1.xml")
    out_text = out.decode("utf-8")

    assert "T[Hello Title]" in out_text
    assert ">123<" in out_text
    assert len(logs) == 1
    assert logs[0].original_text == "Hello Title"
