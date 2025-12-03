# Backlog

## Categories

| Code | Description |
|------|-------------|
| SVG | SVG image support |
| SMART | SmartArt support |
| CHART | Chart types and features |
| TRANS | Transitions and animations |
| SHAPE | Shape features |
| MEDIA | Media types (3D models, etc.) |
| API | Public API improvements |
| TEST | Test infrastructure |
| DOC | Documentation |
| BUILD | Build and tooling |

---

## Phase 0: Baseline Verification

### B-TEST-00 [DONE]
**Verify current python-pptx functionality and establish baseline**

Thoroughly test that python-pptx does what it is supposed to do. Run all existing tests, verify behavior against test files and example documents, and document the current state including any regressions or issues found.

**Details**:
- Run existing test suites (pytest, behave) and document results
- Examine test files in `tests/test_files/` to understand coverage
- Test against user-provided example documents in `WOTAN/example-docs/`
- Document what works, what fails, and any unexpected behavior
- Create a baseline report of current functionality
- Identify any regressions from documented features
- Establish a regression test baseline before making changes

**Acceptance Criteria**:
- All existing tests run and results documented
- Example documents tested for read/modify/write round-trip
- Current state documented with specific pass/fail for each feature area
- Any regressions or bugs identified and catalogued
- Baseline established that future changes can be measured against

**Next**: T-TEST-00 (DONE)

---

## Phase 1: Critical Modern Features

### B-SVG-01 [READY]
**Add SVG image support**

Enable inserting, reading, and modifying SVG images in presentations.

**Details**:
- Add SVG content type detection in `src/pptx/parts/image.py`
- Remove SVG skip logic in `src/pptx/package.py` (line 164)
- Handle `a:svgBlip` element in picture shapes
- Register SVG MIME type
- Add `a16:` namespace for DrawingML 2016 extensions

**Next**: None

---

### B-SMART-01 [NEEDS-SPEC]
**Add SmartArt read support**

Enable reading SmartArt diagrams from existing presentations without data loss.

**Details**:
- Parse `dgm:` namespace elements (diagrams)
- Parse `dsp:` namespace elements (diagram shapes)
- Expose SmartArt data model (text, hierarchy)
- Preserve SmartArt on round-trip even if not fully editable
- Consider: should SmartArt be convertible to regular shapes?

**Next**: None

---

### B-CHART-01 [READY]
**Implement Stock charts**

Add support for Stock chart types (HLC, OHLC, VHLC, VOHLC).

**Details**:
- Enums already exist in `src/pptx/enum/chart.py`
- Need XML writer in `src/pptx/chart/xmlwriter.py`
- Need CT_StockChart element class
- Need series handling for stock data

**Next**: None

---

### B-CHART-02 [READY]
**Implement Surface charts**

Add support for Surface chart types.

**Details**:
- Referenced in `docs/dev/analysis/cht-series.rst` lines 62-65
- Need CT_SurfaceSer, CT_SurfaceChart element classes
- Need XML writer

**Next**: None

---

## Phase 2: PowerPoint 2016+ Charts

### B-CHART-03 [NEEDS-SPEC]
**Implement Treemap charts**

Add support for Treemap charts (Office 2016+).

**Details**:
- Uses `c16:` namespace (Charts 2016 extensions)
- CT_Treemap element
- Hierarchical data model

**Next**: None

---

### B-CHART-04 [NEEDS-SPEC]
**Implement Sunburst charts**

Add support for Sunburst charts (Office 2016+).

**Details**:
- Uses `c16:` namespace
- CT_Sunburst element
- Hierarchical data model similar to Treemap

**Next**: None

---

### B-CHART-05 [NEEDS-SPEC]
**Implement Waterfall charts**

Add support for Waterfall charts (Office 2016+).

**Details**:
- Uses `c16:` namespace
- CT_Waterfall element
- Special handling for subtotals and totals

**Next**: None

---

### B-CHART-06 [NEEDS-SPEC]
**Implement Funnel charts**

Add support for Funnel charts (Office 2019+).

**Details**:
- CT_Funnel element
- Sequential stage data model

**Next**: None

---

### B-CHART-07 [NEEDS-SPEC]
**Implement Map charts**

Add support for Map/Geographic charts (Office 2019+).

**Details**:
- CT_Map element
- Geographic data binding
- May require Bing Maps integration understanding

**Next**: None

---

### B-CHART-08 [NEEDS-SPEC]
**Implement Histogram and Box & Whisker charts**

Add support for statistical charts (Office 2016+).

**Details**:
- CT_Histogram, CT_BoxWhisker elements
- Statistical data calculations

**Next**: None

---

## Phase 3: Transitions and Animation

### B-TRANS-01 [NEEDS-SPEC]
**Implement Morph transitions**

Add support for Morph transition type (PowerPoint 2019+).

**Details**:
- Uses `p14:` namespace (PowerPoint 2010+ extensions)
- `p14:transition` element with morph settings
- Object matching by name ("!!" prefix convention)
- Add transition property to Slide class

**Next**: None

---

### B-TRANS-02 [NEEDS-SPEC]
**Implement Zoom features**

Add support for Section Zoom, Slide Zoom, Summary Zoom.

**Details**:
- Uses `p14:` namespace
- Interactive navigation elements
- Thumbnail generation

**Next**: None

---

## Phase 4: Advanced Features

### B-MEDIA-01 [NEEDS-SPEC]
**Add 3D model support**

Enable inserting and positioning 3D models.

**Details**:
- `a3d:model3d` elements
- 3D model file formats (GLB, etc.)
- Rotation and positioning

**Next**: None

---

### B-SHAPE-01 [READY]
**Add Bezier curve support to FreeformBuilder**

Enable creating curved freeform shapes with cubic Bezier segments.

**Details**:
- Currently only MoveTo, LineTo, Close supported
- Add `cubicBezTo` path segment support
- Update FreeformBuilder API

**Next**: None

---

### B-SHAPE-02 [NEEDS-SPEC]
**Implement Ink annotations**

Add support for reading/writing ink annotations.

**Details**:
- Ink ML format in `ppt/ink/` folder
- Pen strokes, highlighter
- Tablet/stylus input preservation

**Next**: None

---

### B-MEDIA-02 [NEEDS-SPEC]
**Add Cameo (live camera) support**

Support for camera feed placeholder (Microsoft 365 2022+).

**Details**:
- Camera placeholder shape
- Recording integration
- Very new feature, low priority

**Next**: None

---

## Maintenance and Infrastructure

### B-TEST-01 [READY]
**Create test PPTX files with modern features**

Build a collection of test files containing modern PowerPoint features for testing.

**Details**:
- Create PPTX files in PowerPoint with each modern feature
- Use for round-trip testing
- Document expected behavior

**Next**: None

---

### B-DOC-01 [BLOCKED]
**Document new features as implemented**

Update documentation for each new feature added.

**Details**:
- API documentation
- Usage examples
- Update feature matrix

**Blocked by**: Feature implementation

**Next**: None
