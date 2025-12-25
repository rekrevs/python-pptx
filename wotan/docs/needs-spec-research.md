# NEEDS-SPEC Research Summary

This document summarizes research findings for backlog items that were marked NEEDS-SPEC.

## B-CHART-07: Map Charts

### Overview
Map charts (Geographic/Region Map charts) are ChartEx charts introduced in Office 2019.

### XML Structure
- **Namespace**: `http://schemas.microsoft.com/office/drawing/2014/chartex` (cx:)
- **Layout ID**: `regionMap`
- **Series Element**: `<cx:series layoutId="regionMap">`

### Key Elements
- Uses the same ChartEx infrastructure as Treemap, Sunburst, etc.
- Geographic data binding (country/region names as categories)
- May require Bing Maps integration for rendering (PowerPoint-side)

### Implementation Priority
**LOW** - Requires sample file to analyze exact XML structure. The ChartEx infrastructure is already in place, so adding regionMap would be straightforward once the exact data format is understood.

### Sample File Needed
Need a PPTX with a Map chart created in PowerPoint to analyze:
- Data dimension types (strDim for regions, numDim for values)
- Any map-specific layout properties
- Color scale handling

---

## B-TRANS-01: Morph Transitions

### Overview
Morph is a slide transition introduced in PowerPoint 2016 that animates smooth movement between slides.

### XML Structure
- **Namespace**: `http://schemas.microsoft.com/office/powerpoint/2015/09/main` (p159:)
- **Element**: `<p159:morph>`
- **Type**: `CT_MorphTransition`

### Key Attributes
- `option` attribute with values like `byObject` (controls morphing behavior)

### Object Matching
- Objects with matching names morph into each other
- Custom naming: prefix with `!!` (e.g., `!!shape1`) to force matching

### Implementation Details
```xml
<mc:AlternateContent>
  <mc:Choice Requires="p159">
    <p:transition>
      <p159:morph option="byObject"/>
    </p:transition>
  </mc:Choice>
  <mc:Fallback>
    <p:transition>
      <p:fade/>
    </p:transition>
  </mc:Fallback>
</mc:AlternateContent>
```

### Implementation Priority
**MEDIUM** - Well-documented, straightforward to implement. Requires:
1. Register p159 namespace
2. Add CT_MorphTransition element class
3. Add transition property to Slide class
4. Handle AlternateContent wrapper for backwards compatibility

---

## B-TRANS-02: Zoom Features

### Overview
Zoom features (Slide Zoom, Section Zoom, Summary Zoom) were introduced in PowerPoint 2016.

### XML Structure
- **Namespace**: `http://schemas.microsoft.com/office/powerpoint/2016/slidezoom` (p166:)
- **Elements**: `sldZm`, `sectionZm`, `summaryZm`

### Schema Details

#### CT_SlideZoomObject
- `zmPr` element (type: `p166:CT_ZoomObjectProperties`)
- `sldId` attribute (required) - target slide ID
- `cId` attribute (optional) - connection ID

#### CT_SlideZoom
- Contains `sldZmObj` element

### Implementation Details
Zoom objects are special shapes that:
1. Display a thumbnail of the target slide
2. Navigate to that slide when clicked
3. Show a zoom transition effect

### Implementation Priority
**MEDIUM** - Well-documented schema. Requires:
1. Register p166 namespace
2. Add zoom element classes
3. Add zoom thumbnail generation
4. Integrate with slide navigation

---

## B-MEDIA-01: 3D Models

### Overview
3D model support was added in Office 2017 for embedding GLB/glTF models.

### XML Structure
- **Namespace**: `http://schemas.microsoft.com/office/drawing/2017/model3d` (am3d:)
- **Element**: `<am3d:model3d>`
- **Content Type**: `model/gltf-binary` (for .glb files)

### Storage
- Model file: `ppt/media/model3d1.glb`
- Fallback PNG: `ppt/media/image1.png`
- Rendering info in: `ppt/slides/slideX.xml`

### Key Child Elements
1. **am3d:ambientLight** - Ambient lighting settings
2. **am3d:ptLight** - Point light settings
3. **am3d:objViewport** - Viewport/camera settings
4. **am3d:meterPerModelUnit** - Scaling factor

### Camera Settings (CT_Model3DCamera)
- Position relative to 3D model
- Field of view
- Output is always square

### Object Mode vs Window Mode
- **Object Mode**: Model not clipped during rotation (default)
- **Window Mode**: Model clipped to rectangular selection box

### Implementation Priority
**LOW** - Complex feature requiring:
1. GLB/glTF parsing
2. 3D rendering metadata
3. Fallback image generation
4. Complex camera/lighting math

---

## B-SHAPE-02: Ink Annotations

### Overview
Ink annotations store pen/stylus input using InkML format.

### XML Structure
- **Content Type**: `application/inkml+xml`
- **Namespace**: `http://www.w3.org/2003/InkML`
- **Relationship**: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml`

### Storage
- Ink files in: `ppt/ink/` folder
- Root element: `<ink>` in InkML namespace

### InkML Elements
- **trace**: Stroke data with `contextRef` and `brushRef`
- **traceGroup**: Container for organizing traces
- **definitions**: Context and inkSource definitions
- **context**: Rendering parameters
- **inkSource**: Pen input source (requires `xml:id`)

### Channel Data
Supported channels: X, Y, Z, S, T, SN, F, TP, BP, OTx, OTy, OA, OE, OR, RP, RR, RY, TW, TH, TC
Units: dev, in, cm, deg, rad, s, lb, g

### Brush Properties
- width, height, color (RGB hex)
- transparency (0-255)
- tip (ellipse/rectangle)
- rasterOp, antiAliased, fitToCurve, ignorePressure

### Implementation Priority
**LOW** - Specialized feature for:
1. Tablet/stylus input preservation
2. Round-trip of existing ink annotations
3. Not commonly needed for programmatic generation

---

## B-MEDIA-02: Cameo (Live Camera)

### Overview
Cameo allows inserting live camera feed as a placeholder, introduced in Microsoft 365 (2022).

### Limited Documentation
The XML structure for Cameo is not well-documented in public specifications. Based on patterns:
- Likely uses a newer namespace (p188 or similar)
- Implemented as a special placeholder type
- Stores as regular shape with specific properties

### Features
- Live camera feed on slide
- Can apply picture effects/styles
- Integrates with Teams PowerPoint Live
- Can be added to Slide Masters as placeholder

### Limitations
- Only direct cameras (not virtual cameras)
- One video feed per slide

### Implementation Priority
**VERY LOW** - Reasons:
1. Minimal public documentation
2. Very new feature
3. Primarily useful in live presentation context
4. Not commonly needed for programmatic generation

---

## Implementation Recommendations

### Priority Order (High to Low)

1. **B-TRANS-01: Morph Transitions** - Well-documented, high user value
2. **B-TRANS-02: Zoom Features** - Well-documented, high user value
3. **B-CHART-07: Map Charts** - Uses existing ChartEx infrastructure
4. **B-SHAPE-02: Ink Annotations** - Standard InkML, useful for round-trip
5. **B-MEDIA-01: 3D Models** - Complex but documented
6. **B-MEDIA-02: Cameo** - Poorly documented, low priority

### Next Steps

1. Create sample PPTX files with each feature in PowerPoint
2. Extract and analyze XML structure
3. Implement in priority order starting with Morph transitions
