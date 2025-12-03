"""XML writers for ChartEx (Office 2016+) chart types.

This module generates XML for ChartEx charts (cx: namespace), which is completely
different from traditional charts (c: namespace).
"""

from __future__ import annotations

import uuid
from typing import TYPE_CHECKING, Sequence
from xml.sax.saxutils import escape

if TYPE_CHECKING:
    from pptx.chartex.data import ChartExData, ChartExSeriesData


def ChartExXmlWriter(layout_id: str, chart_data: "ChartExData") -> "_BaseChartExXmlWriter":
    """Factory function returning appropriate XML writer for `layout_id`.

    Args:
        layout_id: Chart type identifier (treemap, sunburst, waterfall, funnel, boxWhisker, regionMap)
        chart_data: Data for the chart

    Returns:
        Appropriate XML writer instance for the chart type.
    """
    try:
        WriterCls = {
            "treemap": _TreemapChartExXmlWriter,
            "sunburst": _SunburstChartExXmlWriter,
            "waterfall": _WaterfallChartExXmlWriter,
            "funnel": _FunnelChartExXmlWriter,
            "boxWhisker": _BoxWhiskerChartExXmlWriter,
            "regionMap": _RegionMapChartExXmlWriter,
        }[layout_id]
    except KeyError:
        raise NotImplementedError(
            f"XML writer for ChartEx layout '{layout_id}' not yet implemented"
        )
    return WriterCls(layout_id, chart_data)


class _BaseChartExXmlWriter:
    """Base class for ChartEx XML writers."""

    def __init__(self, layout_id: str, chart_data: "ChartExData"):
        self._layout_id = layout_id
        self._chart_data = chart_data

    @property
    def xml(self) -> str:
        """Return the complete XML for the ChartEx chart."""
        return (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
            "  <cx:chartData>\n"
            "{external_data_xml}"
            "{data_xml}"
            "  </cx:chartData>\n"
            "  <cx:chart>\n"
            "    <cx:plotArea>\n"
            "      <cx:plotAreaRegion>\n"
            "{series_xml}"
            "      </cx:plotAreaRegion>\n"
            "    </cx:plotArea>\n"
            "  </cx:chart>\n"
            "</cx:chartSpace>\n"
        ).format(
            external_data_xml=self._external_data_xml,
            data_xml=self._data_xml,
            series_xml=self._series_xml,
        )

    @property
    def _external_data_xml(self) -> str:
        """Return XML for the external data reference."""
        return '    <cx:externalData r:id="rId1" cx:autoUpdate="0"/>\n'

    @property
    def _data_xml(self) -> str:
        """Return XML for all data elements."""
        return "".join(
            self._data_element_xml(idx, series)
            for idx, series in enumerate(self._chart_data)
        )

    def _data_element_xml(self, data_id: int, series: "ChartExSeriesData") -> str:
        """Return XML for a single cx:data element."""
        return (
            '    <cx:data id="{data_id}">\n'
            '{str_dim_xml}'
            '{num_dim_xml}'
            "    </cx:data>\n"
        ).format(
            data_id=data_id,
            str_dim_xml=self._str_dim_xml(series),
            num_dim_xml=self._num_dim_xml(series),
        )

    def _str_dim_xml(self, series: "ChartExSeriesData") -> str:
        """Return XML for string dimension (categories)."""
        categories = series.categories
        if not categories:
            return ""

        pts_xml = "".join(
            f'        <cx:pt idx="{idx}">{escape(str(cat))}</cx:pt>\n'
            for idx, cat in enumerate(categories)
        )

        return (
            '      <cx:strDim type="cat">\n'
            "        <cx:f>{formula}</cx:f>\n"
            '        <cx:lvl ptCount="{count}">\n'
            "{pts_xml}"
            "        </cx:lvl>\n"
            "      </cx:strDim>\n"
        ).format(
            formula=escape(series.categories_ref or ""),
            count=len(categories),
            pts_xml=pts_xml,
        )

    def _num_dim_xml(self, series: "ChartExSeriesData") -> str:
        """Return XML for numeric dimension (values)."""
        values = series.values
        if not values:
            return ""

        pts_xml = "".join(
            f'        <cx:pt idx="{idx}">{val}</cx:pt>\n'
            for idx, val in enumerate(values)
            if val is not None
        )

        return (
            '      <cx:numDim type="{dim_type}">\n'
            "        <cx:f>{formula}</cx:f>\n"
            '        <cx:lvl ptCount="{count}" formatCode="General">\n'
            "{pts_xml}"
            "        </cx:lvl>\n"
            "      </cx:numDim>\n"
        ).format(
            dim_type=self._num_dim_type,
            formula=escape(series.values_ref or ""),
            count=len(values),
            pts_xml=pts_xml,
        )

    @property
    def _num_dim_type(self) -> str:
        """Return the numeric dimension type for this chart type."""
        return "val"  # Default, overridden by subclasses

    @property
    def _series_xml(self) -> str:
        """Return XML for all series elements."""
        return "".join(
            self._series_element_xml(idx, series)
            for idx, series in enumerate(self._chart_data)
        )

    def _series_element_xml(self, data_id: int, series: "ChartExSeriesData") -> str:
        """Return XML for a single cx:series element."""
        unique_id = "{" + str(uuid.uuid4()).upper() + "}"
        return (
            '        <cx:series layoutId="{layout_id}" uniqueId="{unique_id}">\n'
            "{tx_xml}"
            '          <cx:dataId val="{data_id}"/>\n'
            "{layout_pr_xml}"
            "        </cx:series>\n"
        ).format(
            layout_id=self._layout_id,
            unique_id=unique_id,
            tx_xml=self._tx_xml(series),
            data_id=data_id,
            layout_pr_xml=self._layout_pr_xml(series),
        )

    def _tx_xml(self, series: "ChartExSeriesData") -> str:
        """Return XML for series title."""
        name = series.name
        if not name:
            return ""

        return (
            "          <cx:tx>\n"
            "            <cx:txData>\n"
            "              <cx:f>{formula}</cx:f>\n"
            "              <cx:v>{name}</cx:v>\n"
            "            </cx:txData>\n"
            "          </cx:tx>\n"
        ).format(
            formula=escape(series.name_ref or ""),
            name=escape(name),
        )

    def _layout_pr_xml(self, series: "ChartExSeriesData") -> str:
        """Return XML for layout properties. Override in subclasses."""
        return "          <cx:layoutPr/>\n"


class _TreemapChartExXmlWriter(_BaseChartExXmlWriter):
    """XML writer for Treemap charts."""

    @property
    def _num_dim_type(self) -> str:
        return "size"

    def _layout_pr_xml(self, series: "ChartExSeriesData") -> str:
        """Return XML with parentLabelLayout for Treemap."""
        return (
            "          <cx:layoutPr>\n"
            '            <cx:parentLabelLayout val="overlapping"/>\n'
            "          </cx:layoutPr>\n"
        )


class _SunburstChartExXmlWriter(_BaseChartExXmlWriter):
    """XML writer for Sunburst charts."""

    @property
    def _num_dim_type(self) -> str:
        return "size"


class _WaterfallChartExXmlWriter(_BaseChartExXmlWriter):
    """XML writer for Waterfall charts."""

    def _layout_pr_xml(self, series: "ChartExSeriesData") -> str:
        """Return XML with subtotals for Waterfall."""
        subtotals = getattr(series, "subtotals", [])
        if not subtotals:
            return "          <cx:layoutPr/>\n"

        idx_xml = "".join(f'            <cx:idx val="{idx}"/>\n' for idx in subtotals)
        return (
            "          <cx:layoutPr>\n"
            "            <cx:subtotals>\n"
            "{idx_xml}"
            "            </cx:subtotals>\n"
            "          </cx:layoutPr>\n"
        ).format(idx_xml=idx_xml)


class _FunnelChartExXmlWriter(_BaseChartExXmlWriter):
    """XML writer for Funnel charts."""

    def _layout_pr_xml(self, series: "ChartExSeriesData") -> str:
        """Return XML with visibility settings for Funnel."""
        return (
            "          <cx:layoutPr>\n"
            '            <cx:visibility connectorLines="0" seriesLines="0"/>\n'
            "          </cx:layoutPr>\n"
        )


class _BoxWhiskerChartExXmlWriter(_BaseChartExXmlWriter):
    """XML writer for Box & Whisker charts."""

    def _layout_pr_xml(self, series: "ChartExSeriesData") -> str:
        """Return XML with visibility and statistics for Box & Whisker."""
        return (
            "          <cx:layoutPr>\n"
            '            <cx:visibility meanLine="1" meanMarker="1" outliers="1" nonoutliers="0"/>\n'
            '            <cx:statistics quartileMethod="inclusive"/>\n'
            "          </cx:layoutPr>\n"
        )


class _RegionMapChartExXmlWriter(_BaseChartExXmlWriter):
    """XML writer for Region Map (geographic) charts.

    Map charts display geographic data where regions are colored based on values.
    The actual map rendering is handled by PowerPoint using Bing Maps.
    """

    @property
    def _num_dim_type(self) -> str:
        """Map charts use colorVal for the numeric dimension (determines color intensity)."""
        return "colorVal"
