.. _XlChartExType:

``XL_CHARTEX_TYPE``
===================

Specifies the type of a modern chart (Office 2016+).

These chart types use the ChartEx format (``cx:`` namespace) which is different
from traditional charts. They are created using ``slide.shapes.add_chartex()``
with ``ChartExData``.

Example::

    from pptx.enum.chart import XL_CHARTEX_TYPE

    slide.shapes.add_chartex(
        XL_CHARTEX_TYPE.TREEMAP, x, y, cx, cy, chart_data
    )

----

TREEMAP
    Treemap chart. Displays hierarchical data as nested rectangles.
    Each category is represented by a rectangle sized proportionally
    to its value.

SUNBURST
    Sunburst chart. Similar to treemap but displayed as concentric rings.
    Great for showing hierarchical relationships.

WATERFALL
    Waterfall chart. Shows how an initial value is affected by a series
    of positive or negative values. Supports subtotals.

FUNNEL
    Funnel chart. Visualizes stages in a process, typically showing
    decreasing quantities at each stage.

BOX_WHISKER
    Box & Whisker chart. Statistical chart showing distribution through
    quartiles, median, and outliers.

REGION_MAP
    Map chart. Geographic visualization with regions colored by value.
    Uses Bing Maps for rendering in PowerPoint.
