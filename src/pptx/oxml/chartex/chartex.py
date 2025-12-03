"""Custom element classes for ChartEx (cx:) namespace - Office 2016+ chart types.

ChartEx is a completely separate chart vocabulary introduced in Office 2016 for
modern chart types like Treemap, Sunburst, Waterfall, Funnel, and Box & Whisker.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.simpletypes import XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    pass


class CT_ChartExSpace(BaseOxmlElement):
    """`cx:chartSpace` root element of a ChartEx part.

    This is the root element for Office 2016+ chart types (Treemap, Sunburst,
    Waterfall, Funnel, Box & Whisker, Map).
    """

    _tag_seq = (
        "cx:chartData",
        "cx:chart",
        "cx:spPr",
        "cx:txPr",
        "cx:clrMapOvr",
        "cx:printSettings",
        "cx:extLst",
    )
    chartData: CT_ChartExData = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:chartData", successors=_tag_seq[1:]
    )
    chart: CT_ChartExChart = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:chart", successors=_tag_seq[2:]
    )
    del _tag_seq

    @property
    def xlsx_part_rId(self) -> str | None:
        """Return rId of the external data (embedded Excel workbook), or None."""
        chartData = self.chartData
        if chartData is None:
            return None
        externalData = chartData.externalData
        if externalData is None:
            return None
        return externalData.rId


class CT_ChartExData(BaseOxmlElement):
    """`cx:chartData` element containing data definitions and external references."""

    _tag_seq = (
        "cx:externalData",
        "cx:data",
    )
    externalData: CT_ChartExExternalData | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:externalData", successors=_tag_seq[1:]
    )
    data: CT_ChartExDataElement = ZeroOrMore(  # pyright: ignore[reportAssignmentType]
        "cx:data", successors=()
    )
    del _tag_seq


class CT_ChartExExternalData(BaseOxmlElement):
    """`cx:externalData` element linking to embedded Excel workbook."""

    autoUpdate: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "cx:autoUpdate", XsdUnsignedInt, default=0
    )
    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_ChartExDataElement(BaseOxmlElement):
    """`cx:data` element containing dimension data for a series."""

    _tag_seq = (
        "cx:strDim",
        "cx:numDim",
    )
    strDim: CT_ChartExStrDim | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:strDim", successors=_tag_seq[1:]
    )
    numDim: CT_ChartExNumDim | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:numDim", successors=()
    )
    del _tag_seq

    id: int = RequiredAttribute("id", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


class CT_ChartExStrDim(BaseOxmlElement):
    """`cx:strDim` element for string dimension data (categories)."""

    _tag_seq = (
        "cx:f",
        "cx:nf",
        "cx:lvl",
    )
    f: CT_ChartExFormula | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:f", successors=_tag_seq[1:]
    )
    lvl = ZeroOrMore("cx:lvl", successors=())
    del _tag_seq

    type: str = RequiredAttribute("type", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_ChartExNumDim(BaseOxmlElement):
    """`cx:numDim` element for numeric dimension data (values/sizes)."""

    _tag_seq = (
        "cx:f",
        "cx:nf",
        "cx:lvl",
    )
    f: CT_ChartExFormula | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:f", successors=_tag_seq[1:]
    )
    lvl = ZeroOrMore("cx:lvl", successors=())
    del _tag_seq

    type: str = RequiredAttribute("type", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_ChartExFormula(BaseOxmlElement):
    """`cx:f` element containing a formula reference to spreadsheet data."""

    pass


class CT_ChartExLevel(BaseOxmlElement):
    """`cx:lvl` element containing a level of data points."""

    _tag_seq = ("cx:pt",)
    pt = ZeroOrMore("cx:pt", successors=())
    del _tag_seq

    ptCount: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "ptCount", XsdUnsignedInt
    )
    formatCode: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "formatCode", XsdString
    )


class CT_ChartExDataPoint(BaseOxmlElement):
    """`cx:pt` element containing a single data point value."""

    idx: int = RequiredAttribute("idx", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


class CT_ChartExChart(BaseOxmlElement):
    """`cx:chart` element.

    This element serves two purposes:
    1. As a reference element in `a:graphicData` with just an `r:id` attribute pointing to
       a ChartEx part.
    2. As the chart definition element inside `cx:chartSpace` containing title, plotArea,
       legend, etc.

    The same class handles both uses - when used as a reference, only `rId` is accessed.
    When used as the full chart definition, the child elements are accessed.
    """

    _tag_seq = (
        "cx:title",
        "cx:plotArea",
        "cx:legend",
        "cx:extLst",
    )
    title: CT_ChartExTitle | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:title", successors=_tag_seq[1:]
    )
    plotArea: CT_ChartExPlotArea = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "cx:plotArea"
    )
    legend: CT_ChartExLegend | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:legend", successors=_tag_seq[3:]
    )
    del _tag_seq

    # Used when this element is a reference in a:graphicData
    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]

    @staticmethod
    def new_chart(rId: str) -> "CT_ChartExChart":
        """Return a new `cx:chart` element for use in `a:graphicData`.

        This creates a reference element that points to a ChartEx part.
        """
        return cast(
            CT_ChartExChart,
            parse_xml(f'<cx:chart {nsdecls("cx")} {nsdecls("r")} r:id="{rId}"/>'),
        )


class CT_ChartExTitle(BaseOxmlElement):
    """`cx:title` element for chart title."""

    _tag_seq = (
        "cx:tx",
        "cx:txPr",
        "cx:spPr",
        "cx:extLst",
    )
    tx: CT_ChartExText | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:tx", successors=_tag_seq[1:]
    )
    del _tag_seq

    pos: str | None = OptionalAttribute("pos", XsdString)  # pyright: ignore[reportAssignmentType]
    align: str | None = OptionalAttribute("align", XsdString)  # pyright: ignore[reportAssignmentType]
    overlay: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "overlay", XsdUnsignedInt
    )


class CT_ChartExText(BaseOxmlElement):
    """`cx:tx` element for text content."""

    rich = ZeroOrOne("cx:rich", successors=())
    txData = ZeroOrOne("cx:txData", successors=())


class CT_ChartExLegend(BaseOxmlElement):
    """`cx:legend` element for chart legend."""

    pos: str | None = OptionalAttribute("pos", XsdString)  # pyright: ignore[reportAssignmentType]
    align: str | None = OptionalAttribute("align", XsdString)  # pyright: ignore[reportAssignmentType]
    overlay: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "overlay", XsdUnsignedInt
    )


class CT_ChartExPlotArea(BaseOxmlElement):
    """`cx:plotArea` element containing the plot area."""

    _tag_seq = (
        "cx:plotAreaRegion",
        "cx:axis",
        "cx:spPr",
        "cx:extLst",
    )
    plotAreaRegion: CT_ChartExPlotAreaRegion = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "cx:plotAreaRegion"
    )
    axis = ZeroOrMore("cx:axis", successors=_tag_seq[2:])
    del _tag_seq


class CT_ChartExPlotAreaRegion(BaseOxmlElement):
    """`cx:plotAreaRegion` element containing series."""

    _tag_seq = (
        "cx:plotSurface",
        "cx:series",
        "cx:extLst",
    )
    plotSurface = ZeroOrOne("cx:plotSurface", successors=_tag_seq[1:])
    series = ZeroOrMore("cx:series", successors=_tag_seq[2:])
    del _tag_seq

    def iter_series(self):
        """Generate each cx:series element."""
        for ser in self.series_lst:
            yield ser


class CT_ChartExSeries(BaseOxmlElement):
    """`cx:series` element defining a data series.

    The `layoutId` attribute specifies the chart type:
    - treemap: Treemap chart
    - sunburst: Sunburst chart
    - waterfall: Waterfall chart
    - funnel: Funnel chart
    - boxWhisker: Box & Whisker chart
    - clusteredColumn: Clustered Column (new format)
    - paretoLine: Pareto Line
    - regionMap: Map/Geographic chart
    """

    _tag_seq = (
        "cx:tx",
        "cx:spPr",
        "cx:valueColors",
        "cx:valueColorPositions",
        "cx:dataPt",
        "cx:dataLabels",
        "cx:dataId",
        "cx:layoutPr",
        "cx:axisId",
        "cx:extLst",
    )
    tx = ZeroOrOne("cx:tx", successors=_tag_seq[1:])
    dataId: CT_ChartExDataId | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:dataId", successors=_tag_seq[7:]
    )
    layoutPr: CT_ChartExLayoutPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:layoutPr", successors=_tag_seq[8:]
    )
    del _tag_seq

    layoutId: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "layoutId", XsdString
    )
    uniqueId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "uniqueId", XsdString
    )


class CT_ChartExDataId(BaseOxmlElement):
    """`cx:dataId` element referencing data in cx:chartData."""

    val: int = RequiredAttribute("val", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


class CT_ChartExLayoutPr(BaseOxmlElement):
    """`cx:layoutPr` element containing layout-specific properties.

    Different chart types have different children:
    - Treemap: parentLabelLayout
    - Waterfall: subtotals
    - Funnel: visibility
    - Box & Whisker: visibility, statistics
    """

    _tag_seq = (
        "cx:parentLabelLayout",
        "cx:subtotals",
        "cx:visibility",
        "cx:statistics",
        "cx:extLst",
    )
    parentLabelLayout: CT_ChartExParentLabelLayout | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:parentLabelLayout", successors=_tag_seq[1:]
    )
    subtotals: CT_ChartExSubtotals | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:subtotals", successors=_tag_seq[2:]
    )
    visibility: CT_ChartExVisibility | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:visibility", successors=_tag_seq[3:]
    )
    statistics: CT_ChartExStatistics | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "cx:statistics", successors=_tag_seq[4:]
    )
    del _tag_seq


class CT_ChartExParentLabelLayout(BaseOxmlElement):
    """`cx:parentLabelLayout` for Treemap charts."""

    val: str = RequiredAttribute("val", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_ChartExSubtotals(BaseOxmlElement):
    """`cx:subtotals` for Waterfall charts."""

    idx = ZeroOrMore("cx:idx", successors=())


class CT_ChartExIdx(BaseOxmlElement):
    """`cx:idx` element specifying a subtotal index."""

    val: int = RequiredAttribute("val", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


class CT_ChartExVisibility(BaseOxmlElement):
    """`cx:visibility` element for Funnel and Box & Whisker charts."""

    connectorLines: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "connectorLines", XsdUnsignedInt
    )
    seriesLines: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "seriesLines", XsdUnsignedInt
    )
    meanLine: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "meanLine", XsdUnsignedInt
    )
    meanMarker: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "meanMarker", XsdUnsignedInt
    )
    outliers: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "outliers", XsdUnsignedInt
    )
    nonoutliers: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "nonoutliers", XsdUnsignedInt
    )


class CT_ChartExStatistics(BaseOxmlElement):
    """`cx:statistics` element for Box & Whisker charts."""

    quartileMethod: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "quartileMethod", XsdString
    )


class CT_ChartExAxis(BaseOxmlElement):
    """`cx:axis` element for chart axis."""

    _tag_seq = (
        "cx:catScaling",
        "cx:valScaling",
        "cx:title",
        "cx:units",
        "cx:majorGridlines",
        "cx:minorGridlines",
        "cx:majorTickMarks",
        "cx:minorTickMarks",
        "cx:tickLabels",
        "cx:numFmt",
        "cx:spPr",
        "cx:txPr",
        "cx:extLst",
    )
    title = ZeroOrOne("cx:title", successors=_tag_seq[3:])
    del _tag_seq

    id: int = RequiredAttribute("id", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
    hidden: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "hidden", XsdUnsignedInt
    )


