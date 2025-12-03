"""Excel workbook writer for ChartEx charts."""

from __future__ import annotations

import io
from contextlib import contextmanager
from typing import TYPE_CHECKING

from xlsxwriter import Workbook

if TYPE_CHECKING:
    from pptx.chartex.data import ChartExData


class ChartExWorkbookWriter:
    """Writes Excel workbook containing data for ChartEx charts."""

    def __init__(self, chart_data: "ChartExData"):
        self._chart_data = chart_data

    @property
    def xlsx_blob(self) -> bytes:
        """Return bytes for Excel file containing chart data."""
        xlsx_file = io.BytesIO()
        with self._open_worksheet(xlsx_file) as (workbook, worksheet):
            self._populate_worksheet(workbook, worksheet)
        return xlsx_file.getvalue()

    @contextmanager
    def _open_worksheet(self, xlsx_file):
        """Context manager for Excel workbook creation."""
        workbook = Workbook(xlsx_file, {"in_memory": True})
        worksheet = workbook.add_worksheet()
        yield workbook, worksheet
        workbook.close()

    def _populate_worksheet(self, workbook, worksheet):
        """Write chart data to worksheet.

        Layout:
        - Row 1: Series names (starting from column B)
        - Column A: Categories (starting from row 2)
        - Columns B+: Series values
        """
        chart_data = self._chart_data
        num_format = workbook.add_format({"num_format": chart_data.number_format})

        # Write categories in column A (if any series has categories)
        if chart_data and chart_data[0].categories:
            categories = chart_data[0].categories
            for row_idx, category in enumerate(categories):
                worksheet.write(row_idx + 1, 0, category)

        # Write series data
        for col_idx, series in enumerate(chart_data):
            # Write series name in row 0
            worksheet.write(0, col_idx + 1, series.name)

            # Write values
            for row_idx, value in enumerate(series.values):
                if value is not None:
                    worksheet.write(row_idx + 1, col_idx + 1, value, num_format)
