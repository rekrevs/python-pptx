"""ChartEx part objects for Office 2016+ chart types."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.parts.embeddedpackage import EmbeddedXlsxPart
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.chartex.chartex import CT_ChartExSpace
    from pptx.package import Package


class ChartExPart(XmlPart):
    """A ChartEx part for Office 2016+ chart types.

    Corresponds to parts having partnames matching ppt/charts/chartEx[1-9][0-9]*.xml.
    Handles chart types like Treemap, Sunburst, Waterfall, Funnel, and Box & Whisker.
    """

    partname_template = "/ppt/charts/chartEx%d.xml"

    @classmethod
    def new(
        cls,
        layout_id: str,
        chart_data: "ChartExData",
        package: "Package",
    ) -> "ChartExPart":
        """Return new |ChartExPart| instance added to `package`.

        The chart is created with `layout_id` (e.g., 'treemap', 'sunburst')
        and depicts `chart_data`.
        """
        chart_part = cls.load(
            package.next_partname(cls.partname_template),
            CT.OFC_CHART_EX,
            package,
            chart_data.xml_bytes(layout_id),
        )
        chart_part.chart_workbook.update_from_xlsx_blob(chart_data.xlsx_blob)
        return chart_part

    @property
    def chartex(self) -> "ChartEx":
        """Return |ChartEx| object representing the chart in this part."""
        from pptx.chartex.chartex import ChartEx

        return ChartEx(self._element, self)

    @lazyproperty
    def chart_workbook(self) -> "ChartExWorkbook":
        """Return |ChartExWorkbook| for access to external chart data."""
        return ChartExWorkbook(self._element, self)


class ChartExWorkbook:
    """Provides access to external chart data in embedded Excel workbook."""

    def __init__(self, chartSpace: "CT_ChartExSpace", chart_part: ChartExPart):
        super().__init__()
        self._chartSpace = chartSpace
        self._chart_part = chart_part

    def update_from_xlsx_blob(self, xlsx_blob: bytes) -> None:
        """Replace Excel spreadsheet in related EmbeddedXlsxPart.

        If no EmbeddedXlsxPart exists, creates one and establishes the relationship.
        """
        xlsx_part = self.xlsx_part
        if xlsx_part is None:
            self.xlsx_part = EmbeddedXlsxPart.new(xlsx_blob, self._chart_part.package)
            return
        xlsx_part.blob = xlsx_blob

    @property
    def xlsx_part(self) -> EmbeddedXlsxPart | None:
        """Return |EmbeddedXlsxPart| containing data for this chart, or None.

        The related part has its rId at `cx:chartSpace/cx:chartData/cx:externalData/@r:id`.
        Returns None if no rId is specified or if the relationship doesn't exist yet.
        """
        xlsx_part_rId = self._chartSpace.xlsx_part_rId
        if xlsx_part_rId is None:
            return None
        # Check if relationship exists (may not exist for newly created charts)
        try:
            return self._chart_part.related_part(xlsx_part_rId)
        except KeyError:
            return None

    @xlsx_part.setter
    def xlsx_part(self, xlsx_part: EmbeddedXlsxPart) -> None:
        """Set the related |EmbeddedXlsxPart| to `xlsx_part`."""
        rId = self._chart_part.relate_to(xlsx_part, RT.PACKAGE)
        chartData = self._chartSpace.chartData
        if chartData is None:
            raise ValueError("ChartEx has no chartData element")
        externalData = chartData.externalData
        if externalData is None:
            raise ValueError("ChartEx has no externalData element")
        externalData.rId = rId
