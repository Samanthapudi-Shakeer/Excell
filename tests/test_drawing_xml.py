from __future__ import annotations

import xml.etree.ElementTree as ET

from excel_translator.drawing_xml import _translate_in_xml


def test_translate_in_xml_only_updates_a_t_nodes():
    xml = b"""<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:p><a:r><a:t>Chart Title</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:valAx>
        <c:title>
          <c:tx><c:rich><a:p><a:r><a:t>Axis Label</a:t></a:r></a:p></c:rich></c:tx>
        </c:title>
      </c:valAx>
      <c:ser>
        <c:val>
          <c:numRef>
            <c:numCache>
              <c:pt idx=\"0\"><c:v>100</c:v></c:pt>
            </c:numCache>
          </c:numRef>
        </c:val>
      </c:ser>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
"""

    translated_xml, logs = _translate_in_xml(xml, lambda text, _id: (f"T[{text}]", "fake"), "xl/charts/chart1.xml")
    root = ET.fromstring(translated_xml)

    translated_text_nodes = [node.text for node in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t")]
    chart_value_nodes = [node.text for node in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/chart}v")]

    assert translated_text_nodes == ["T[Chart Title]", "T[Axis Label]"]
    assert chart_value_nodes == ["100"]
    assert len(logs) == 2


def test_translate_in_xml_skips_empty_or_whitespace_a_t_nodes():
    xml = b"""<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">
  <xdr:twoCellAnchor>
    <xdr:sp>
      <xdr:txBody>
        <a:p><a:r><a:t>  </a:t></a:r></a:p>
        <a:p><a:r><a:t>Flow Step</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
  </xdr:twoCellAnchor>
</xdr:wsDr>
"""

    translated_xml, logs = _translate_in_xml(xml, lambda text, _id: (f"T[{text}]", "fake"), "xl/drawings/drawing1.xml")
    root = ET.fromstring(translated_xml)
    text_nodes = [node.text for node in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t")]

    assert text_nodes == ["  ", "T[Flow Step]"]
    assert len(logs) == 1

