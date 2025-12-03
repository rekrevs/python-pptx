# Task T-CHART-11

## Header

| Field | Value |
|-------|-------|
| ID | T-CHART-11 |
| Parent | User request |
| State | DONE |
| Created | 2025-12-03 |

## Objective

Add read support for stock charts (`c:stockChart`) so that python-pptx can open presentations containing stock charts without raising `ValueError: unsupported plot type`.

## Acceptance Criteria

- [x] Add `StockPlot` class to `src/pptx/chart/plot.py`
- [x] Register `StockPlot` in `PlotFactory` mapping
- [x] Implement `_differentiate_stock_chart_type()` in `PlotTypeInspector`
- [x] Add `StockSeries` class to `src/pptx/chart/series.py`
- [x] Register `StockSeries` in `_SeriesFactory` mapping
- [x] Stock charts correctly report their `chart_type` (HLC, OHLC)
- [x] `WOTAN/example-docs/test-features/stock-chart-test.pptx` can be opened without error
- [x] Unit tests pass

## Context

### Existing Code
- `src/pptx/chart/plot.py` - Plot classes and PlotFactory (missing StockPlot)
- `src/pptx/oxml/chart/plot.py` - CT_StockChart already defined
- `src/pptx/oxml/__init__.py` - CT_StockChart already registered
- `src/pptx/enum/chart.py` - XL_CHART_TYPE has STOCK_HLC, STOCK_OHLC, STOCK_VHLC, STOCK_VOHLC

### Stock Chart XML Structure
Stock charts use `<c:stockChart>` element with 3-4 series:
- HLC (High-Low-Close): 3 series
- OHLC (Open-High-Low-Close): 4 series
- VHLC (Volume-High-Low-Close): Bar chart + 3 series stock
- VOHLC (Volume-Open-High-Low-Close): Bar chart + 4 series stock

The series names in the test file are: High, Low, Close (for HLC variant).

### Differentiation Logic
- Count series: 3 = HLC, 4 = OHLC
- Volume variants have a separate bar chart overlay (not in stockChart element)

## Implementation Notes

### Approach

1. Add minimal `StockPlot` class inheriting from `_BasePlot`
2. Add mapping in `PlotFactory`: `qn("c:stockChart"): StockPlot`
3. Add `_differentiate_stock_chart_type()` method to `PlotTypeInspector`
4. Add `StockSeries` class inheriting from `_BaseCategorySeries` and `_MarkerMixin`
5. Add mapping in `_SeriesFactory`: `qn("c:stockChart"): StockSeries`
6. Use series count to determine HLC vs OHLC variant

### Files Modified

- `src/pptx/chart/plot.py` - Add StockPlot class, factory mapping, and type differentiation
- `src/pptx/chart/series.py` - Add StockSeries class and factory mapping
- `tests/chart/test_plot.py` - Add test cases for stock chart type detection
- `tests/chart/test_series.py` - Add test cases for StockSeries

## Evidence

### Tests Run

```
$ pytest tests/chart/test_plot.py tests/chart/test_series.py -v --tb=short
... (all tests passed) ...
tests/chart/test_plot.py: 115 passed
tests/chart/test_series.py: 70 passed
```

### Verification with Real File

```python
>>> from pptx import Presentation
>>> prs = Presentation("stock-chart-test.pptx")
>>> for slide in prs.slides:
...     for shape in slide.shapes:
...         if shape.has_chart:
...             print(shape.chart.chart_type)
...             for series in shape.chart.plots[0].series:
...                 print(f"  {series.name}: {series.values[:3]}")
STOCK_HLC
  High: (105.0, 108.0, 107.0)
  Low: (98.0, 100.0, 99.0)
  Close: (102.0, 105.0, 103.0)
STOCK_OHLC
  Open: (100.0, 102.0, 105.0)
  High: (105.0, 108.0, 107.0)
  Low: (98.0, 100.0, 99.0)
  Close: (102.0, 105.0, 103.0)
```

### Test Results

- Unit tests: PASS (185 tests total)
- Verification with real file: PASS

## Outcome

**State**: DONE

Stock chart read support has been successfully implemented. Presentations containing stock charts can now be opened and their chart type, series names, and values can be accessed programmatically.
