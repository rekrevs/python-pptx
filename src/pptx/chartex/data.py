"""ChartExData and related objects for Office 2016+ chart types."""

from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING, Iterator

from pptx.chartex.xmlwriter import ChartExXmlWriter

if TYPE_CHECKING:
    pass


class ChartExData(Sequence):
    """Data object for ChartEx charts (Treemap, Sunburst, Waterfall, Funnel, Box & Whisker).

    Serves as a proxy for the chart data that will be written to an Excel worksheet.
    Operates as a sequence of series.
    """

    def __init__(self, number_format: str = "General"):
        super().__init__()
        self._number_format = number_format
        self._series: list[ChartExSeriesData] = []

    def __getitem__(self, index: int) -> "ChartExSeriesData":
        return self._series[index]

    def __len__(self) -> int:
        return len(self._series)

    def __iter__(self) -> Iterator["ChartExSeriesData"]:
        return iter(self._series)

    def add_series(
        self,
        name: str,
        categories: Sequence[str] | None = None,
        values: Sequence[float | int | None] | None = None,
    ) -> "ChartExSeriesData":
        """Add and return a new series with `name`, `categories`, and `values`."""
        series = ChartExSeriesData(self, name, categories, values)
        self._series.append(series)
        return series

    @property
    def number_format(self) -> str:
        """Return the number format string for this chart's values."""
        return self._number_format

    def xml_bytes(self, layout_id: str) -> bytes:
        """Return XML bytes for a ChartEx chart of type `layout_id`."""
        return ChartExXmlWriter(layout_id, self).xml.encode("utf-8")

    @property
    def xlsx_blob(self) -> bytes:
        """Return bytes of Excel workbook containing this chart's data."""
        from pptx.chartex.xlsx import ChartExWorkbookWriter

        return ChartExWorkbookWriter(self).xlsx_blob


class ChartExSeriesData:
    """Data for a single series in a ChartEx chart."""

    def __init__(
        self,
        chart_data: ChartExData,
        name: str,
        categories: Sequence[str] | None = None,
        values: Sequence[float | int | None] | None = None,
    ):
        self._chart_data = chart_data
        self._name = name
        self._categories = list(categories) if categories else []
        self._values = list(values) if values else []
        self._subtotals: list[int] = []  # For Waterfall charts

    @property
    def name(self) -> str:
        """Return the name of this series."""
        return self._name

    @property
    def name_ref(self) -> str:
        """Return Excel worksheet reference for the series name cell."""
        idx = self._chart_data._series.index(self)
        col = chr(ord("B") + idx)  # B, C, D, ...
        return f"Sheet1!${col}$1"

    @property
    def categories(self) -> list[str]:
        """Return list of category labels for this series."""
        return self._categories

    @property
    def categories_ref(self) -> str:
        """Return Excel worksheet reference for the categories."""
        if not self._categories:
            return ""
        return f"Sheet1!$A$2:$A${len(self._categories) + 1}"

    @property
    def values(self) -> list[float | int | None]:
        """Return list of values for this series."""
        return self._values

    @property
    def values_ref(self) -> str:
        """Return Excel worksheet reference for the values."""
        if not self._values:
            return ""
        idx = self._chart_data._series.index(self)
        col = chr(ord("B") + idx)  # B, C, D, ...
        return f"Sheet1!${col}$2:${col}${len(self._values) + 1}"

    @property
    def subtotals(self) -> list[int]:
        """Return list of subtotal indices for Waterfall charts."""
        return self._subtotals

    def set_subtotals(self, indices: Sequence[int]) -> None:
        """Set which data points are subtotals (for Waterfall charts).

        Args:
            indices: Zero-based indices of values that are subtotals
        """
        self._subtotals = list(indices)
