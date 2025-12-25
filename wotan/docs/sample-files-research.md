# Sample Files Research

This document tracks sample files gathered for research and testing of modern PPTX features.

## Available Sample Files

Located in `WOTAN/example-docs/test-features/`:

| File | Size | Contents | Source |
|------|------|----------|--------|
| `smartart-business-model-canvas.pptx` | 125KB | SmartArt diagrams (Business Model Canvas) | [GitHub: bfritscher](https://github.com/bfritscher/smartart-business-model-canvas) |
| `charts-standard.pptx` | 84KB | 8 standard chart types | Created with python-pptx |
| `freeform-bezier.pptx` | 29KB | Bezier curve freeform shapes | Created with python-pptx |

## SmartArt Analysis

### File Structure
```
ppt/diagrams/
├── data1.xml      # Diagram data model (nodes and connections)
├── drawing1.xml   # Visual representation
├── layout1.xml    # Layout algorithm definition
├── quickStyle1.xml # Quick style settings
├── colors1.xml    # Color scheme
```

### Key XML Structure (dgm:dataModel)

```xml
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
  <dgm:ptLst>
    <!-- Points (nodes) in the diagram -->
    <dgm:pt modelId="{GUID}" type="doc">
      <dgm:prSet loTypeId="..." qsTypeId="..." csTypeId="..."/>
      <dgm:t>  <!-- Text content -->
        <a:p><a:r><a:t>Node Text</a:t></a:r></a:p>
      </dgm:t>
    </dgm:pt>

    <!-- Regular content nodes -->
    <dgm:pt modelId="{GUID}">
      <dgm:t><a:p><a:r><a:t>Content Text</a:t></a:r></a:p></dgm:t>
    </dgm:pt>

    <!-- Transition nodes (parTrans = parent transition, sibTrans = sibling transition) -->
    <dgm:pt modelId="{GUID}" type="parTrans" cxnId="{CONNECTION_GUID}"/>
  </dgm:ptLst>

  <dgm:cxnLst>
    <!-- Connections between nodes -->
    <dgm:cxn srcId="{SRC_GUID}" destId="{DEST_GUID}" type="parOf"/>
  </dgm:cxnLst>
</dgm:dataModel>
```

### Text Extraction Path
- SmartArt text is stored in `dgm:pt/dgm:t/a:p/a:r/a:t`
- Same structure as regular PowerPoint text (DrawingML)
- Node hierarchy defined by `dgm:cxnLst` connections

### Python-pptx Detection
- `shape.shape_type == MSO_SHAPE_TYPE.DIAGRAM` for SmartArt shapes
- `shape.has_smart_art == True` for GraphicFrame containing SmartArt

## Still Needed: PowerPoint-Created Files

The following features require files created in Microsoft PowerPoint (2019+) because they cannot be created programmatically with python-pptx:

### Priority 1: Modern Chart Types (Office 2016+)

| Chart Type | Namespace | Why Needed |
|------------|-----------|------------|
| Stock (HLC, OHLC) | `c:stockChart` | Multi-plot structure with hi-low lines |
| Surface | `c:surfaceChart` | Wireframe and 3D rendering |
| Treemap | `cx:treemap` | New chartEx namespace |
| Sunburst | `cx:sunburst` | Hierarchical data format |
| Waterfall | `cx:waterfall` | Subtotal markers |
| Histogram | `cx:histogram` | Statistical binning |
| Box & Whisker | `cx:boxWhisker` | Quartile calculations |

**Note**: Modern charts (Treemap, Sunburst, Waterfall, etc.) use the `cx:` (chartEx) namespace instead of `c:`, which is fundamentally different from traditional OOXML charts.

### Priority 2: SVG Images

| Feature | Element | Why Needed |
|---------|---------|------------|
| SVG insertion | `a16:svgBlip` | Dual-format storage with PNG fallback |
| MCE wrapper | `mc:AlternateContent` | Compatibility structure |

### Priority 3: Transitions (Office 2019+)

| Feature | Element | Why Needed |
|---------|---------|------------|
| Morph transition | `p14:transition` | Object matching mechanism |

## How to Create Required Files

### Stock Chart
1. PowerPoint > Insert > Chart > Stock > High-Low-Close
2. Add sample data (Date, High, Low, Close values)
3. Save as `stock-chart-test.pptx`

### Surface Chart
1. PowerPoint > Insert > Chart > Surface
2. Select "3-D Surface" or "Wireframe"
3. Add sample 3D data grid
4. Save as `surface-chart-test.pptx`

### Modern Charts (2016+)
1. PowerPoint 2016+ > Insert > Chart > [Treemap|Sunburst|Waterfall|Histogram]
2. Add appropriate data
3. Save as `modern-charts-test.pptx`

### SVG Image
1. Download an SVG file (e.g., from [sample-files.com](https://sample-files.com/images/svg/))
2. PowerPoint 2019+ > Insert > Pictures > Select SVG
3. Save as `svg-image-test.pptx`

### Morph Transition
1. Create slide 1 with a shape
2. Duplicate slide
3. Move/resize the shape on slide 2
4. Apply "Morph" transition to slide 2
5. Save as `morph-transition-test.pptx`

## Chart Namespace Differences

### Traditional Charts (c: namespace)
```xml
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>...</c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
```

### Modern Charts (cx: namespace - Office 2016+)
```xml
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="treemap">...</cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>
```

Key differences:
- Uses `chartEx1.xml` instead of `chart1.xml`
- Tags use `cx:` prefix instead of `c:`
- Different structure for data binding
- Part stored in `ppt/charts/chartExN.xml`

## References

- [python-pptx Issue #583](https://github.com/scanny/python-pptx/issues/583) - New Chart Types in Office 2016
- [python-pptx Issue #83](https://github.com/scanny/python-pptx/issues/83) - SmartArt support feature request
- [ECMA-376](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) - Office Open XML specification
