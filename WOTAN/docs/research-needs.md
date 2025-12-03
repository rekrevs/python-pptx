# Research Needs

This document tracks what sample files and research are needed to complete feature implementations.

## Status Summary

| Feature | Sample File | XML Analyzed | API Designed | Implemented |
|---------|-------------|--------------|--------------|-------------|
| SmartArt detection | ✅ | ✅ | ✅ | ✅ |
| SmartArt text extraction | ✅ | ✅ | ❌ | ❌ |
| SVG insertion | ❌ Needs PowerPoint | ❌ | ❌ | ❌ |
| Stock chart creation | ❌ Needs PowerPoint | ❌ | ❌ | ❌ |
| Surface chart creation | ❌ Needs PowerPoint | ❌ | ❌ | ❌ |
| Treemap/Sunburst | ❌ Needs PowerPoint 2016+ | ❌ | ❌ | ❌ |
| Morph transitions | ❌ Needs PowerPoint 2019+ | ❌ | ❌ | ❌ |

---

## Completed Research

### SmartArt Diagrams ✅

**File obtained**: `WOTAN/example-docs/test-features/smartart-business-model-canvas.pptx`

**Findings**:
- SmartArt is stored in `ppt/diagrams/` with multiple XML parts
- Data model in `dataX.xml` contains `dgm:ptLst` (points/nodes) and `dgm:cxnLst` (connections)
- Text stored in `dgm:pt/dgm:t/a:p/a:r/a:t` (standard DrawingML format)
- Python-pptx correctly detects as `MSO_SHAPE_TYPE.DIAGRAM` with `has_smart_art=True`

**Next steps**: Implement text extraction API to read SmartArt content

---

## Priority 1: Sample PPTX Files Still Needed

To proceed with Phase 1 completion, we need PPTX files created in PowerPoint containing:

### 1. SVG Image (Critical)
**File needed**: A presentation with an inserted SVG image

**How to create**:
1. Open PowerPoint (2019 or later, or Microsoft 365)
2. Insert > Pictures > This Device
3. Select an SVG file
4. Save as .pptx

**What we'll learn**:
- `a16:svgBlip` element structure
- MCE `AlternateContent` block format
- Relationship between SVG and PNG fallback parts
- How `a16:` namespace is declared

### 2. Stock Chart
**File needed**: A presentation with a stock chart

**How to create**:
1. Open PowerPoint
2. Insert > Chart > Stock
3. Choose "High-Low-Close" initially
4. Also create "Open-High-Low-Close" variant
5. Add sample data
6. Save as .pptx

**What we'll learn**:
- `c:stockChart` element structure
- Series ordering (High, Low, Close positions)
- `c:hiLowLines` and `c:upDownBars` formatting
- Axis configuration

### 3. Surface Chart
**File needed**: A presentation with surface charts

**How to create**:
1. Open PowerPoint
2. Insert > Chart > Surface
3. Create both "3-D Surface" and "Wireframe" variants
4. Save as .pptx

**What we'll learn**:
- `c:surfaceChart` vs `c:surface3DChart` difference
- `c:wireframe` attribute
- Axis requirements (2 vs 3)

---

## Priority 2: 2016+ Features (Phase 2)

These require PowerPoint 2016 or later:

### 4. Modern Chart Types
**File needed**: Presentations with each chart type

| Chart Type | How to Create |
|------------|---------------|
| Treemap | Insert > Chart > Treemap |
| Sunburst | Insert > Chart > Sunburst |
| Waterfall | Insert > Chart > Waterfall |
| Funnel | Insert > Chart > Funnel |
| Map | Insert > Chart > Map |
| Histogram | Insert > Chart > Histogram |
| Box & Whisker | Insert > Chart > Box & Whisker |

**What we'll learn**:
- `cx:` namespace structure (chartEx, not chart)
- Hierarchical data format for Treemap/Sunburst
- Special markers for Waterfall subtotals

### 5. Morph Transition
**File needed**: Two-slide presentation with Morph

**How to create**:
1. Create slide 1 with a shape
2. Duplicate to slide 2
3. Move/resize the shape on slide 2
4. Apply Morph transition to slide 2
5. Save as .pptx

**What we'll learn**:
- `p14:transition` element structure
- Object matching mechanism
- Transition properties

---

## File Naming Convention

Please save files to `WOTAN/example-docs/test-features/` with names:
- `svg-image-test.pptx`
- `stock-chart-test.pptx`
- `surface-chart-test.pptx`
- `modern-charts-test.pptx`
- `morph-transition-test.pptx`

---

## Already Available Files

Located in `WOTAN/example-docs/test-features/`:

| File | Contents |
|------|----------|
| `smartart-business-model-canvas.pptx` | 3 SmartArt diagrams |
| `charts-standard.pptx` | 8 standard chart types (Bar, Line, Pie, etc.) |
| `freeform-bezier.pptx` | Bezier curve freeform shapes |
