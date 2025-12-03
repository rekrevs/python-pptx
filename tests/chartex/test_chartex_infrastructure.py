"""Tests for ChartEx infrastructure (cx: namespace)."""

import pytest


class DescribeChartExData:
    """Tests for ChartExData class."""

    def it_can_add_a_series(self):
        from pptx.chartex.data import ChartExData

        data = ChartExData()
        series = data.add_series(
            "Sales", categories=["A", "B", "C"], values=[1, 2, 3]
        )

        assert len(data) == 1
        assert series.name == "Sales"
        assert series.categories == ["A", "B", "C"]
        assert series.values == [1, 2, 3]

    def it_can_generate_xlsx_blob(self):
        from pptx.chartex.data import ChartExData

        data = ChartExData()
        data.add_series("Sales", ["A", "B"], [1, 2])

        xlsx = data.xlsx_blob

        # Check it's a valid ZIP file (xlsx is a zip)
        assert len(xlsx) > 0
        assert xlsx[:4] == b"PK\x03\x04"


class DescribeChartExXmlWriter:
    """Tests for ChartEx XML writers."""

    def it_generates_valid_treemap_xml(self):
        from pptx.chartex.data import ChartExData
        from pptx.chartex.xmlwriter import ChartExXmlWriter

        data = ChartExData()
        data.add_series("Sales", ["North", "South", "East", "West"], [100, 150, 200, 120])

        writer = ChartExXmlWriter("treemap", data)
        xml = writer.xml

        assert "<?xml version" in xml
        assert "<cx:chartSpace" in xml
        assert 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"' in xml
        assert "<cx:chartData>" in xml
        assert '<cx:data id="0">' in xml
        assert '<cx:strDim type="cat">' in xml
        assert '<cx:numDim type="size">' in xml  # Treemap uses 'size'
        assert '<cx:series layoutId="treemap"' in xml
        assert '<cx:parentLabelLayout val="overlapping"/>' in xml

    def it_generates_valid_sunburst_xml(self):
        from pptx.chartex.data import ChartExData
        from pptx.chartex.xmlwriter import ChartExXmlWriter

        data = ChartExData()
        data.add_series("Sales", ["A", "B"], [1, 2])

        writer = ChartExXmlWriter("sunburst", data)
        xml = writer.xml

        assert '<cx:series layoutId="sunburst"' in xml
        assert '<cx:numDim type="size">' in xml

    def it_generates_valid_waterfall_xml(self):
        from pptx.chartex.data import ChartExData
        from pptx.chartex.xmlwriter import ChartExXmlWriter

        data = ChartExData()
        series = data.add_series("Revenue", ["Q1", "Q2", "Subtotal"], [100, 50, 150])
        series.set_subtotals([2])

        writer = ChartExXmlWriter("waterfall", data)
        xml = writer.xml

        assert '<cx:series layoutId="waterfall"' in xml
        assert "<cx:subtotals>" in xml
        assert '<cx:idx val="2"/>' in xml

    def it_generates_valid_funnel_xml(self):
        from pptx.chartex.data import ChartExData
        from pptx.chartex.xmlwriter import ChartExXmlWriter

        data = ChartExData()
        data.add_series("Stages", ["Lead", "Qualified", "Won"], [100, 50, 20])

        writer = ChartExXmlWriter("funnel", data)
        xml = writer.xml

        assert '<cx:series layoutId="funnel"' in xml
        assert '<cx:visibility connectorLines="0" seriesLines="0"/>' in xml

    def it_generates_valid_boxwhisker_xml(self):
        from pptx.chartex.data import ChartExData
        from pptx.chartex.xmlwriter import ChartExXmlWriter

        data = ChartExData()
        data.add_series("Data", ["Group1", "Group2"], [10, 20])

        writer = ChartExXmlWriter("boxWhisker", data)
        xml = writer.xml

        assert '<cx:series layoutId="boxWhisker"' in xml
        assert '<cx:visibility meanLine="1"' in xml
        assert '<cx:statistics quartileMethod="inclusive"/>' in xml

    def it_generates_valid_regionmap_xml(self):
        from pptx.chartex.data import ChartExData
        from pptx.chartex.xmlwriter import ChartExXmlWriter

        data = ChartExData()
        data.add_series("Population", ["United States", "Germany", "France"], [331, 83, 67])

        writer = ChartExXmlWriter("regionMap", data)
        xml = writer.xml

        assert '<cx:series layoutId="regionMap"' in xml
        assert '<cx:numDim type="colorVal">' in xml

    def it_raises_for_unsupported_layout(self):
        from pptx.chartex.data import ChartExData
        from pptx.chartex.xmlwriter import ChartExXmlWriter

        data = ChartExData()
        data.add_series("X", ["A"], [1])

        with pytest.raises(NotImplementedError) as exc:
            ChartExXmlWriter("unsupported", data)

        assert "not yet implemented" in str(exc.value)


class DescribeChartExOXML:
    """Tests for ChartEx OXML element classes."""

    def it_registers_cx_namespace(self):
        from pptx.oxml.ns import qn

        # Verify the cx namespace is registered
        clark_name = qn("cx:chartSpace")
        assert clark_name == "{http://schemas.microsoft.com/office/drawing/2014/chartex}chartSpace"

    def it_can_parse_chartex_xml(self):
        from pptx.oxml import parse_xml
        from pptx.oxml.chartex.chartex import CT_ChartExSpace

        xml = (
            '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">'
            "  <cx:chartData/>"
            "  <cx:chart>"
            "    <cx:plotArea>"
            "      <cx:plotAreaRegion/>"
            "    </cx:plotArea>"
            "  </cx:chart>"
            "</cx:chartSpace>"
        )

        chartSpace = parse_xml(xml)

        assert isinstance(chartSpace, CT_ChartExSpace)
        assert chartSpace.chartData is not None
        assert chartSpace.chart is not None


class DescribeChartExPart:
    """Tests for ChartExPart class."""

    def it_is_registered_for_chartex_content_type(self):
        from pptx.opc.constants import CONTENT_TYPE as CT
        from pptx.opc.package import PartFactory
        from pptx.parts.chartex import ChartExPart

        # Verify ChartExPart is registered for the chartex content type
        assert PartFactory.part_type_for.get(CT.OFC_CHART_EX) is ChartExPart
