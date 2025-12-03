# Task T-TEST-02

## Header

| Field | Value |
|-------|-------|
| ID | T-TEST-02 |
| Parent | B-TEST-01 |
| State | PARTIAL |
| Created | 2024-12-03 |
| Updated | 2024-12-03 |

## Objective

Create test PPTX files containing modern PowerPoint features for research and testing. These files will enable deeper analysis of XML structures needed for implementation.

## Acceptance Criteria

- [x] PPTX with SmartArt diagrams obtained
- [x] SmartArt XML structure analyzed and documented
- [x] Standard chart types sample created
- [x] Bezier freeform sample created
- [ ] PPTX with Stock chart created (requires PowerPoint)
- [ ] PPTX with Surface chart created (requires PowerPoint)
- [ ] PPTX with embedded SVG image created (requires PowerPoint 2019+)
- [ ] PPTX with modern charts (Treemap, Sunburst, etc.) created (requires PowerPoint 2016+)
- [ ] XML structure documented for each feature

## Context

Several Phase 1 and Phase 2 features require sample PPTX files to understand the XML structure before implementation can proceed.

## Files Obtained/Created

### SmartArt Sample
- **File**: `smartart-business-model-canvas.pptx` (125KB)
- **Source**: [GitHub: bfritscher/smartart-business-model-canvas](https://github.com/bfritscher/smartart-business-model-canvas)
- **Contents**: Business Model Canvas with 3 SmartArt diagrams
- **Analysis**: Complete - see `WOTAN/docs/sample-files-research.md`

### Standard Charts Sample
- **File**: `charts-standard.pptx` (84KB)
- **Source**: Created with python-pptx
- **Contents**: 8 slides with Bar, Line, Pie, Scatter, Bubble, Area, Doughnut, Radar charts

### Bezier Freeform Sample
- **File**: `freeform-bezier.pptx` (29KB)
- **Source**: Created with python-pptx (using new Bezier support from T-SHAPE-01)
- **Contents**: Heart shape and wave using cubic Bezier curves

## Research Findings

### SmartArt Structure
SmartArt files contain multiple XML parts in `ppt/diagrams/`:
- `dataX.xml` - Node hierarchy and text content
- `drawingX.xml` - Visual representation
- `layoutX.xml` - Layout algorithm
- `quickStyleX.xml` - Style settings
- `colorsX.xml` - Color scheme

Text is stored in `dgm:pt/dgm:t/a:p/a:r/a:t` using standard DrawingML text format.

### Modern Charts (chartEx)
Office 2016+ charts use a completely different structure:
- Namespace: `cx:` instead of `c:`
- Part location: `ppt/charts/chartExN.xml` instead of `chartN.xml`
- Different element structure and data binding

This requires significant new infrastructure in python-pptx.

### What Still Needs PowerPoint

The following cannot be created programmatically and require Microsoft PowerPoint:

1. **Stock Charts** - Need PowerPoint to create `c:stockChart` with hi-low lines
2. **Surface Charts** - Need PowerPoint for wireframe/3D rendering
3. **Modern Charts** - Treemap, Sunburst, Waterfall use `cx:` namespace
4. **SVG Images** - Need PowerPoint 2019+ to create `a16:svgBlip` with fallback
5. **Morph Transitions** - Need PowerPoint 2019+ for `p14:transition`

## Evidence

### SmartArt Detection Test
```python
>>> from pptx import Presentation
>>> prs = Presentation('smartart-business-model-canvas.pptx')
>>> for slide in prs.slides:
...     for shape in slide.shapes:
...         if hasattr(shape, 'has_smart_art'):
...             print(shape.shape_type, shape.has_smart_art)
MSO_SHAPE_TYPE.DIAGRAM True
MSO_SHAPE_TYPE.DIAGRAM True
MSO_SHAPE_TYPE.DIAGRAM True
```

## Outcome

**State**: PARTIAL

Obtained SmartArt sample and created standard chart/Bezier samples. Remaining files require Microsoft PowerPoint to create:
- Stock and Surface charts
- SVG embedded images
- Modern chart types (Treemap, Sunburst, Waterfall)
- Morph transitions

Detailed documentation created in `WOTAN/docs/sample-files-research.md`.
