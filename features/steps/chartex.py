"""Gherkin step implementations for ChartEx chart features."""

from __future__ import annotations

from behave import then, when

from pptx.chartex.data import ChartExData
from pptx.enum.chart import XL_CHARTEX_TYPE
from pptx.oxml.ns import qn
from pptx.spec import GRAPHIC_DATA_URI_CHARTEX
from pptx.util import Inches

# when ====================================================


@when("I add a {chart_type} ChartEx chart with sample data")
def when_I_add_a_chartex_chart_with_sample_data(context, chart_type):
    type_map = {
        "Waterfall": XL_CHARTEX_TYPE.WATERFALL,
        "Funnel": XL_CHARTEX_TYPE.FUNNEL,
        "Box & Whisker": XL_CHARTEX_TYPE.BOX_AND_WHISKER,
        "Region Map": XL_CHARTEX_TYPE.REGION_MAP,
    }
    chart_data = ChartExData()
    chart_data.add_series(
        "Series 1", categories=["A", "B", "C", "D"], values=[10, 20, 15, 25]
    )
    context.graphic_frame = context.slide.shapes.add_chartex(
        type_map[chart_type],
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(4),
        chart_data,
    )


# then ====================================================


@then("the slide has a graphic frame containing a ChartEx chart")
def then_the_slide_has_a_graphic_frame_containing_a_chartex_chart(context):
    graphic_frame = context.graphic_frame
    assert type(graphic_frame).__name__ == "GraphicFrame", (
        "expected GraphicFrame, got %s" % type(graphic_frame).__name__
    )
    uri = graphic_frame._element.graphicData_uri
    assert uri == GRAPHIC_DATA_URI_CHARTEX, (
        "expected ChartEx URI, got %s" % uri
    )


@then("the ChartEx chart layout is {layout_id}")
def then_the_chartex_chart_layout_is_layout_id(context, layout_id):
    graphic_frame = context.graphic_frame
    graphicData = graphic_frame._element.graphic.graphicData
    cx_chart = graphicData.find(qn("cx:chart"))
    assert cx_chart is not None, "no cx:chart element found in graphicData"
    rId = cx_chart.get(qn("r:id"))
    assert rId is not None, "cx:chart element has no r:id attribute"
    chartex_part = graphic_frame.part.related_part(rId)
    actual_layout_id = chartex_part.chartex.layout_id
    assert actual_layout_id == layout_id, (
        "expected layout_id '%s', got '%s'" % (layout_id, actual_layout_id)
    )
