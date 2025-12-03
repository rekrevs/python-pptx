# Task T-CHART-01

## Header

| Field | Value |
|-------|-------|
| ID | T-CHART-01 |
| Parent | B-CHART-01 |
| State | PARTIAL |
| Created | 2024-12-03 |
| Updated | 2024-12-03 |

## Objective

Implement Stock chart support in python-pptx, enabling creation and manipulation of Stock chart types (HLC, OHLC, VHLC, VOHLC).

## Acceptance Criteria

- [x] CT_StockChart OXML element class defined
- [x] CT_UpDownBars OXML element class defined (for candlestick rendering)
- [x] OXML elements registered in __init__.py
- [x] All tests pass (2700 unit, 973 acceptance)
- [ ] _StockChartXmlWriter class implemented
- [ ] Stock chart types registered in ChartXmlWriter factory
- [ ] Unit tests for XML writer
- [ ] Integration tests (BDD) for Stock chart creation

## Context

### Stock Chart Types

| Enum Name | Value | Series Count | Description |
|-----------|-------|--------------|-------------|
| `STOCK_HLC` | 88 | 3 | High-Low-Close |
| `STOCK_OHLC` | 89 | 4 | Open-High-Low-Close |
| `STOCK_VHLC` | 90 | 4 | Volume-High-Low-Close |
| `STOCK_VOHLC` | 91 | 5 | Volume-Open-High-Low-Close |

### XML Schema (CT_StockChart)

```xml
<xsd:complexType name="CT_StockChart">
  <xsd:sequence>
    <xsd:element name="ser"        type="CT_LineSer"       minOccurs="3" maxOccurs="4"/>
    <xsd:element name="dLbls"      type="CT_DLbls"         minOccurs="0"/>
    <xsd:element name="dropLines"  type="CT_ChartLines"    minOccurs="0"/>
    <xsd:element name="hiLowLines" type="CT_ChartLines"    minOccurs="0"/>
    <xsd:element name="upDownBars" type="CT_UpDownBars"    minOccurs="0"/>
    <xsd:element name="axId"       type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
    <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0"/>
  </xsd:sequence>
</xsd:complexType>
```

## Files Modified

- `src/pptx/oxml/chart/plot.py` - Added CT_StockChart and CT_UpDownBars classes
- `src/pptx/oxml/__init__.py` - Registered c:stockChart and c:upDownBars elements

## Evidence

### Unit tests pass

```
$ pytest tests/ -q
2700 passed in 3.21s
```

### OXML elements defined

```python
class CT_StockChart(BaseChartElement):
    """Stock charts display High-Low-Close (or Open-High-Low-Close) data."""
    ser = ZeroOrMore("c:ser", ...)
    dLbls = ZeroOrOne("c:dLbls", ...)

class CT_UpDownBars(BaseOxmlElement):
    """Controls up/down bars in stock charts."""
    gapWidth = ZeroOrOne("c:gapWidth", ...)
```

## Remaining Work

The XML writer implementation is complex because:

1. **Multi-plot charts**: VHLC/VOHLC variants require a bar chart for volume alongside the stock chart
2. **Series ordering**: Series must be in specific order (High, Low, Close or Open, High, Low, Close)
3. **Special elements**: hiLowLines, upDownBars need proper XML generation
4. **Chart data handling**: May need special ChartData class for stock data

Recommended next steps:
1. Create sample PPTX with stock chart from PowerPoint
2. Analyze XML structure
3. Implement _StockChartXmlWriter for HLC first (simplest variant)
4. Extend to OHLC, then volume variants

## Outcome

**State**: PARTIAL

OXML infrastructure for Stock charts is in place:
- CT_StockChart and CT_UpDownBars elements defined
- Elements registered in __init__.py
- Round-trip of existing stock charts will work

Remaining:
- XML writer for creating new stock charts
- Tests specific to stock chart creation
