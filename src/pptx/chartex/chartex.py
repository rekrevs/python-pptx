"""ChartEx class providing access to Office 2016+ chart types."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.shared import PartElementProxy

if TYPE_CHECKING:
    from pptx.oxml.chartex.chartex import CT_ChartExSpace
    from pptx.parts.chartex import ChartExPart


class ChartEx(PartElementProxy):
    """Provides access to Office 2016+ charts (Treemap, Sunburst, Waterfall, etc.)."""

    def __init__(self, chartSpace: "CT_ChartExSpace", chartex_part: "ChartExPart"):
        super().__init__(chartSpace, chartex_part)
        self._chartSpace = chartSpace
        self._chartex_part = chartex_part

    @property
    def layout_id(self) -> str | None:
        """Return the chart type layout ID (e.g., 'treemap', 'sunburst').

        Returns None if no series is present.
        """
        chart = self._chartSpace.chart
        if chart is None:
            return None
        plotArea = chart.plotArea
        if plotArea is None:
            return None
        plotAreaRegion = plotArea.plotAreaRegion
        if plotAreaRegion is None:
            return None
        # Get first series to determine layout type
        for series in plotAreaRegion.series_lst:
            return series.layoutId
        return None

    @property
    def chart_type(self) -> str:
        """Return a human-readable chart type name."""
        layout_id = self.layout_id
        type_names = {
            "treemap": "Treemap",
            "sunburst": "Sunburst",
            "waterfall": "Waterfall",
            "funnel": "Funnel",
            "boxWhisker": "Box & Whisker",
            "clusteredColumn": "Clustered Column",
            "paretoLine": "Pareto Line",
            "regionMap": "Map",
        }
        return type_names.get(layout_id, "Unknown") if layout_id else "Unknown"
