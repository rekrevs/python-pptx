# Task T-CHART-09

## Header

| Field | Value |
|-------|-------|
| ID | T-CHART-09 |
| Parent | B-CHART-09 |
| State | DONE |
| Created | 2024-12-03 |
| Updated | 2024-12-03 |

## Objective

Build the foundational ChartEx infrastructure required by all Phase 2 charts (Treemap, Sunburst, Waterfall, Funnel, Box & Whisker).

## Acceptance Criteria

- [x] Register `cx:` namespace in `ns.py`
- [x] Add `RT.CHART_EX` relationship type in `constants.py`
- [x] Content type `OFC_CHART_EX` already exists in `constants.py`
- [x] Create OXML element classes for ChartEx
- [x] Create `ChartExPart` class for `application/vnd.ms-office.chartex+xml`
- [x] Register `ChartExPart` in part factory
- [x] Create base `ChartExXmlWriter` with implementations for all chart types
- [x] Create `ChartExData` class for data handling
- [x] Create `ChartExWorkbookWriter` for Excel data generation
- [x] All tests pass (2711 unit, 973 acceptance)

## Context

ChartEx (`cx:` namespace) is a completely different chart vocabulary from traditional charts (`c:` namespace). It was introduced in Office 2016 for modern chart types.

Key differences:
- Different content type: `application/vnd.ms-office.chartex+xml`
- Different namespace: `http://schemas.microsoft.com/office/drawing/2014/chartex`
- Different relationship type: `http://schemas.microsoft.com/office/2014/relationships/chartEx`
- Uses `layoutId` attribute to specify chart type instead of element names
- Data stored in `cx:chartData` with `cx:strDim` and `cx:numDim` dimensions

## Implementation

### Files Created

- `src/pptx/oxml/chartex/__init__.py` - Package init
- `src/pptx/oxml/chartex/chartex.py` - OXML element classes (24 classes)
- `src/pptx/chartex/__init__.py` - ChartEx package init
- `src/pptx/chartex/chartex.py` - ChartEx API class
- `src/pptx/chartex/data.py` - ChartExData and ChartExSeriesData classes
- `src/pptx/chartex/xmlwriter.py` - XML writers for all chart types
- `src/pptx/chartex/xlsx.py` - Excel workbook writer
- `src/pptx/parts/chartex.py` - ChartExPart and ChartExWorkbook classes
- `tests/chartex/__init__.py` - Test package init
- `tests/chartex/test_chartex_infrastructure.py` - 11 tests

### Files Modified

- `src/pptx/oxml/ns.py` - Added `cx:` namespace
- `src/pptx/opc/constants.py` - Added `RT.CHART_EX` relationship type
- `src/pptx/oxml/__init__.py` - Registered 24 ChartEx element classes
- `src/pptx/__init__.py` - Registered `ChartExPart` for content type

### OXML Element Classes

```python
CT_ChartExSpace          # cx:chartSpace - root element
CT_ChartExData           # cx:chartData
CT_ChartExExternalData   # cx:externalData
CT_ChartExDataElement    # cx:data
CT_ChartExStrDim         # cx:strDim
CT_ChartExNumDim         # cx:numDim
CT_ChartExFormula        # cx:f
CT_ChartExLevel          # cx:lvl
CT_ChartExDataPoint      # cx:pt
CT_ChartExChart          # cx:chart
CT_ChartExTitle          # cx:title
CT_ChartExText           # cx:tx
CT_ChartExLegend         # cx:legend
CT_ChartExPlotArea       # cx:plotArea
CT_ChartExPlotAreaRegion # cx:plotAreaRegion
CT_ChartExSeries         # cx:series
CT_ChartExDataId         # cx:dataId
CT_ChartExLayoutPr       # cx:layoutPr
CT_ChartExParentLabelLayout  # cx:parentLabelLayout (Treemap)
CT_ChartExSubtotals      # cx:subtotals (Waterfall)
CT_ChartExIdx            # cx:idx
CT_ChartExVisibility     # cx:visibility (Funnel, BoxWhisker)
CT_ChartExStatistics     # cx:statistics (BoxWhisker)
CT_ChartExAxis           # cx:axis
```

### XML Writer Classes

```python
_TreemapChartExXmlWriter   # layoutId="treemap"
_SunburstChartExXmlWriter  # layoutId="sunburst"
_WaterfallChartExXmlWriter # layoutId="waterfall"
_FunnelChartExXmlWriter    # layoutId="funnel"
_BoxWhiskerChartExXmlWriter # layoutId="boxWhisker"
```

## Evidence

### Unit tests pass

```
$ pytest tests/ -q
2711 passed in 3.35s
```

### Acceptance tests pass

```
$ behave features/ -q
54 features passed, 0 failed, 0 skipped
973 scenarios passed, 0 failed, 0 skipped
2914 steps passed, 0 failed, 0 skipped
```

### ChartEx infrastructure tests pass

```
$ pytest tests/chartex/ -v
tests/chartex/test_chartex_infrastructure.py::DescribeChartExData::it_can_add_a_series PASSED
tests/chartex/test_chartex_infrastructure.py::DescribeChartExData::it_can_generate_xlsx_blob PASSED
tests/chartex/test_chartex_infrastructure.py::DescribeChartExXmlWriter::it_generates_valid_treemap_xml PASSED
tests/chartex/test_chartex_infrastructure.py::DescribeChartExXmlWriter::it_generates_valid_sunburst_xml PASSED
tests/chartex/test_chartex_infrastructure.py::DescribeChartExXmlWriter::it_generates_valid_waterfall_xml PASSED
tests/chartex/test_chartex_infrastructure.py::DescribeChartExXmlWriter::it_generates_valid_funnel_xml PASSED
tests/chartex/test_chartex_infrastructure.py::DescribeChartExXmlWriter::it_generates_valid_boxwhisker_xml PASSED
tests/chartex/test_chartex_infrastructure.py::DescribeChartExXmlWriter::it_raises_for_unsupported_layout PASSED
tests/chartex/test_chartex_infrastructure.py::DescribeChartExOXML::it_registers_cx_namespace PASSED
tests/chartex/test_chartex_infrastructure.py::DescribeChartExOXML::it_can_parse_chartex_xml PASSED
tests/chartex/test_chartex_infrastructure.py::DescribeChartExPart::it_is_registered_for_chartex_content_type PASSED

11 passed in 0.07s
```

## API Summary

```python
# Create chart data
from pptx.chartex.data import ChartExData

data = ChartExData()
series = data.add_series(
    "Sales",
    categories=["Q1", "Q2", "Q3", "Q4"],
    values=[100, 150, 200, 175]
)

# For Waterfall charts, mark subtotals
series.set_subtotals([3])  # Index 3 is a subtotal

# Generate XML
from pptx.chartex.xmlwriter import ChartExXmlWriter

writer = ChartExXmlWriter("treemap", data)  # or sunburst, waterfall, funnel, boxWhisker
xml = writer.xml

# Generate Excel workbook
xlsx_bytes = data.xlsx_blob
```

## Outcome

**State**: DONE

All acceptance criteria met. ChartEx infrastructure is now complete and unblocks all Phase 2 chart implementations:
- B-CHART-03: Treemap charts
- B-CHART-04: Sunburst charts
- B-CHART-05: Waterfall charts
- B-CHART-06: Funnel charts
- B-CHART-08: Box & Whisker charts

The next step to complete Phase 2 is to add the public API for inserting ChartEx charts via `shapes.add_chartex()` method.
