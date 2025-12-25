# ChartEx (cx:) Namespace Specification

This document specifies the ChartEx namespace structure for Office 2016+ chart types in python-pptx.

## Overview

Office 2016 introduced a new chart vocabulary using the `cx:` (chartEx) namespace, completely separate from the traditional `c:` (chart) namespace. These charts use:

- **Different content type**: `application/vnd.ms-office.chartex+xml`
- **Different namespace**: `http://schemas.microsoft.com/office/drawing/2014/chartex`
- **Different part location**: `ppt/charts/chartExN.xml`
- **Different structure**: `cx:chartSpace` instead of `c:chartSpace`

## Namespace URIs

| Prefix | URI | Purpose |
|--------|-----|---------|
| `cx` | `http://schemas.microsoft.com/office/drawing/2014/chartex` | ChartEx base namespace |
| `cx1` | `http://schemas.microsoft.com/office/drawing/2015/9/8/chartex` | ChartEx extension |

## Supported Chart Types

ChartEx supports these chart types via the `layoutId` attribute on `cx:series`:

| layoutId | Chart Type | Office Version |
|----------|------------|----------------|
| `treemap` | Treemap | 2016+ |
| `sunburst` | Sunburst | 2016+ |
| `waterfall` | Waterfall | 2016+ |
| `funnel` | Funnel | 2019+ |
| `boxWhisker` | Box & Whisker | 2016+ |
| `clusteredColumn` | Clustered Column (new format) | 2016+ |
| `paretoLine` | Pareto Line | 2016+ |
| `regionMap` | Map/Geographic | 2019+ |

## XML Structure

### Root Element: cx:chartSpace

```xml
<cx:chartSpace
    xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cx:chartData>...</cx:chartData>
  <cx:chart>...</cx:chart>
  <cx:spPr>...</cx:spPr>  <!-- optional shape properties -->
  <cx:txPr>...</cx:txPr>  <!-- optional text properties -->
  <cx:clrMapOvr>...</cx:clrMapOvr>  <!-- optional color mapping -->
  <cx:printSettings>...</cx:printSettings>  <!-- optional -->
  <cx:extLst>...</cx:extLst>  <!-- optional extensions -->
</cx:chartSpace>
```

### cx:chartData

Contains data definitions and external data references:

```xml
<cx:chartData>
  <cx:externalData r:id="rId1" cx:autoUpdate="0"/>
  <cx:data id="0">
    <cx:strDim type="cat">
      <cx:f>Sheet1!$A$2:$A$5</cx:f>
      <cx:lvl ptCount="4">
        <cx:pt idx="0">Category 1</cx:pt>
        <cx:pt idx="1">Category 2</cx:pt>
      </cx:lvl>
    </cx:strDim>
    <cx:numDim type="size">
      <cx:f>Sheet1!$B$2:$B$5</cx:f>
      <cx:lvl ptCount="4" formatCode="General">
        <cx:pt idx="0">4.3</cx:pt>
        <cx:pt idx="1">2.5</cx:pt>
      </cx:lvl>
    </cx:numDim>
  </cx:data>
</cx:chartData>
```

### cx:chart

Contains the visual representation:

```xml
<cx:chart>
  <cx:title pos="t" align="ctr" overlay="0">
    <cx:tx>
      <cx:rich>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:r><a:t>Chart Title</a:t></a:r>
        </a:p>
      </cx:rich>
    </cx:tx>
  </cx:title>
  <cx:plotArea>
    <cx:plotAreaRegion>
      <cx:series layoutId="treemap" uniqueId="{GUID}">
        <cx:tx>
          <cx:txData>
            <cx:f>Sheet1!$B$1</cx:f>
            <cx:v>Series 1</cx:v>
          </cx:txData>
        </cx:tx>
        <cx:dataId val="0"/>
        <cx:layoutPr/>
      </cx:series>
    </cx:plotAreaRegion>
  </cx:plotArea>
  <cx:legend pos="t" align="ctr" overlay="0"/>
</cx:chart>
```

## Chart-Specific Structures

### Treemap

```xml
<cx:series layoutId="treemap" uniqueId="{GUID}">
  <cx:tx>...</cx:tx>
  <cx:dataId val="0"/>
  <cx:layoutPr>
    <cx:parentLabelLayout val="overlapping"/>  <!-- or "banner", "none" -->
  </cx:layoutPr>
</cx:series>
```

Data dimensions:
- `cx:strDim type="cat"` - Category hierarchy (can have multiple levels for nested treemaps)
- `cx:numDim type="size"` - Size values

### Sunburst

```xml
<cx:series layoutId="sunburst" uniqueId="{GUID}">
  <cx:tx>...</cx:tx>
  <cx:dataId val="0"/>
  <cx:layoutPr/>
</cx:series>
```

Data dimensions (same as Treemap):
- `cx:strDim type="cat"` - Category hierarchy (multiple levels for rings)
- `cx:numDim type="size"` - Size values

### Waterfall

```xml
<cx:series layoutId="waterfall" uniqueId="{GUID}">
  <cx:tx>...</cx:tx>
  <cx:dataId val="0"/>
  <cx:layoutPr>
    <cx:subtotals>
      <cx:idx val="3"/>  <!-- Index of subtotal bars -->
      <cx:idx val="6"/>
    </cx:subtotals>
  </cx:layoutPr>
</cx:series>
```

Data dimensions:
- `cx:strDim type="cat"` - Category labels
- `cx:numDim type="val"` - Values (positive = up, negative = down)

### Funnel

```xml
<cx:series layoutId="funnel" uniqueId="{GUID}">
  <cx:tx>...</cx:tx>
  <cx:dataId val="0"/>
  <cx:layoutPr>
    <cx:visibility connectorLines="0" seriesLines="0"/>
  </cx:layoutPr>
</cx:series>
```

### Box & Whisker

```xml
<cx:series layoutId="boxWhisker" uniqueId="{GUID}">
  <cx:tx>...</cx:tx>
  <cx:dataId val="0"/>
  <cx:layoutPr>
    <cx:visibility meanLine="1" meanMarker="1" outliers="1" nonoutliers="0"/>
    <cx:statistics quartileMethod="inclusive"/>
  </cx:layoutPr>
</cx:series>
```

## Relationships

ChartEx charts require different relationship types:

| Purpose | Relationship Type |
|---------|-------------------|
| Chart reference | `http://schemas.microsoft.com/office/2014/relationships/chartEx` |
| Embedded workbook | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/package` |

## Content Types

```xml
<Override PartName="/ppt/charts/chartEx1.xml"
          ContentType="application/vnd.ms-office.chartex+xml"/>
```

## Implementation Approach for python-pptx

### Required Components

1. **New Part Class**: `ChartExPart` for `application/vnd.ms-office.chartex+xml`
2. **New Namespace**: Register `cx:` and `cx1:` prefixes
3. **OXML Elements**: Create CT_ChartExSpace, CT_ChartExData, CT_ChartExPlotArea, etc.
4. **XML Writers**: New ChartExXmlWriter classes for each chart type
5. **API Classes**: ChartEx, TreemapChart, SunburstChart, WaterfallChart, etc.
6. **Relationship Type**: Add CHART_EX to relationship constants

### Key Differences from c: charts

| Aspect | c: (traditional) | cx: (ChartEx) |
|--------|------------------|---------------|
| Part extension | chart1.xml | chartEx1.xml |
| Content type | openxmlformats...chart+xml | ms-office.chartex+xml |
| Data storage | Embedded in chart XML | Referenced via cx:chartData |
| Series type | Per-chart elements (barChart, lineChart) | Generic series with layoutId |
| Hierarchy | Single level categories | Multi-level for treemap/sunburst |

## References

- [MS-PPTX Specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/efd8bb2d-d888-4e2e-af25-cad476730c9f)
- [Open XML SDK ChartSpace](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2016.drawing.chartdrawing.chartspace)
- [python-pptx Issue #371](https://github.com/scanny/python-pptx/issues/371) - Treemap support discussion
- [Open-XML-SDK Issue #675](https://github.com/OfficeDev/Open-XML-SDK/issues/675) - cx:chartSpace implementation
