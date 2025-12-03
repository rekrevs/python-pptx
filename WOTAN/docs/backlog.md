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

### B-SVG-01 [DONE]
**Add SVG image support**

Enable inserting, reading, and modifying SVG images in presentations.

**Details**:
- ~~Add SVG content type detection in `src/pptx/parts/image.py`~~ DONE
- ~~Remove SVG skip logic in `src/pptx/package.py` (line 164)~~ DONE
- ~~Register SVG MIME type~~ DONE
- ~~Handle `asvg:svgBlip` element in picture shapes~~ DONE
- ~~Add `asvg:` namespace for SVG extension~~ DONE
- ~~Implement SVG insertion API with PNG fallback~~ DONE (shapes.add_svg_picture())

**Next**: T-SVG-01 (DONE)

---

### B-SMART-01 [DONE]
**Add SmartArt read support**

Enable reading SmartArt diagrams from existing presentations without data loss.

**Details**:
- ~~Register `dgm:` namespace (diagrams)~~ DONE
- ~~Register `dsp:` namespace (diagram shapes)~~ DONE
- ~~Detect SmartArt in GraphicFrame.shape_type~~ DONE (returns MSO_SHAPE_TYPE.DIAGRAM)
- ~~Add has_smart_art property~~ DONE
- ~~Preserve SmartArt on round-trip~~ DONE (was already working)
- ~~Parse diagram data model (text, hierarchy)~~ DONE
- ~~Expose Python API for SmartArt content~~ DONE (shape.smart_art.all_text/text)

**Next**: T-SMART-01 (DONE)

---

### B-CHART-01 [DONE]
**Implement Stock charts**

Add support for Stock chart types (HLC, OHLC, VHLC, VOHLC).

**Details**:
- ~~Enums already exist in `src/pptx/enum/chart.py`~~ DONE
- ~~CT_StockChart element class~~ DONE
- ~~CT_UpDownBars element class~~ DONE
- ~~OXML elements registered~~ DONE
- ~~XML writer in `src/pptx/chart/xmlwriter.py`~~ DONE (_StockChartXmlWriter)
- ~~Series handling for stock data~~ DONE (3 series for HLC, 4 for OHLC)

**Next**: T-CHART-01 (DONE)

---

### B-CHART-02 [DONE]
**Implement Surface charts**

Add support for Surface chart types (SURFACE, SURFACE_WIREFRAME, SURFACE_TOP_VIEW, SURFACE_TOP_VIEW_WIREFRAME).

**Details**:
- ~~CT_SurfaceChart OXML element~~ DONE
- ~~CT_Surface3DChart OXML element~~ DONE
- ~~Elements registered~~ DONE
- ~~XML writer~~ DONE (_SurfaceChartXmlWriter with 3D/2D and wireframe variants)
- PlotTypeInspector detection (future - for reading)
- Plot class with wireframe property (future - for reading)

**Next**: T-CHART-02 (DONE - XML writing complete, reading to be added as needed)

---

## Phase 2: PowerPoint 2016+ Charts (ChartEx)

**Note**: All Phase 2 charts use the ChartEx (`cx:`) namespace, which is completely different from traditional charts (`c:`). See `WOTAN/docs/chartex-spec.md` for full specification.

**Common Infrastructure Required**:
- Register `cx:` namespace (`http://schemas.microsoft.com/office/drawing/2014/chartex`)
- Create `ChartExPart` for content type `application/vnd.ms-office.chartex+xml`
- Add relationship type `http://schemas.microsoft.com/office/2014/relationships/chartEx`
- Create OXML elements: CT_ChartExSpace, CT_ChartExData, CT_ChartExPlotArea, CT_ChartExSeries
- Create base ChartExXmlWriter infrastructure

### B-CHART-03 [DONE]
**Implement Treemap charts**

Add support for Treemap charts (Office 2016+).

**Details**:
- ~~Uses `cx:` namespace (ChartEx) - NOT `c16:`~~ DONE
- ~~`cx:series layoutId="treemap"` element~~ DONE
- ~~Hierarchical data: `cx:strDim type="cat"` (multi-level), `cx:numDim type="size"`~~ DONE
- ~~Parent label layout options: "overlapping", "banner", "none"~~ DONE

**Spec**: See `WOTAN/docs/chartex-spec.md#treemap`

**Next**: T-CHART-10 (DONE)

---

### B-CHART-04 [DONE]
**Implement Sunburst charts**

Add support for Sunburst charts (Office 2016+).

**Details**:
- ~~Uses `cx:` namespace (ChartEx)~~ DONE
- ~~`cx:series layoutId="sunburst"` element~~ DONE
- ~~Same hierarchical data model as Treemap (multi-level categories = rings)~~ DONE

**Spec**: See `WOTAN/docs/chartex-spec.md#sunburst`

**Next**: T-CHART-10 (DONE)

---

### B-CHART-05 [DONE]
**Implement Waterfall charts**

Add support for Waterfall charts (Office 2016+).

**Details**:
- ~~Uses `cx:` namespace (ChartEx)~~ DONE
- ~~`cx:series layoutId="waterfall"` element~~ DONE
- ~~Categories with `cx:strDim type="cat"`, values with `cx:numDim type="val"`~~ DONE
- ~~Subtotals specified via `cx:subtotals/cx:idx` elements~~ DONE

**Spec**: See `WOTAN/docs/chartex-spec.md#waterfall`

**Next**: T-CHART-10 (DONE)

---

### B-CHART-06 [DONE]
**Implement Funnel charts**

Add support for Funnel charts (Office 2019+).

**Details**:
- ~~Uses `cx:` namespace (ChartEx)~~ DONE
- ~~`cx:series layoutId="funnel"` element~~ DONE
- ~~Sequential stage data model~~ DONE
- ~~`cx:visibility` for connector/series lines~~ DONE

**Spec**: See `WOTAN/docs/chartex-spec.md#funnel`

**Next**: T-CHART-10 (DONE)

---

### B-CHART-07 [DONE]
**Implement Map charts**

Add support for Map/Geographic charts (Office 2019+).

**Details**:
- ~~Uses `cx:` namespace (ChartEx) - same infrastructure as other ChartEx charts~~ DONE
- ~~`cx:series layoutId="regionMap"` element~~ DONE
- ~~Geographic data binding (country/region names as categories)~~ DONE
- ~~Color scale handling via `cx:numDim type="colorVal"`~~ DONE

**Next**: T-CHART-12 (DONE)

---

### B-CHART-08 [DONE]
**Implement Box & Whisker charts**

Add support for Box & Whisker charts (Office 2016+).

**Details**:
- ~~Uses `cx:` namespace (ChartEx)~~ DONE
- ~~`cx:series layoutId="boxWhisker"` element~~ DONE
- ~~`cx:visibility` for mean line, markers, outliers~~ DONE
- ~~`cx:statistics quartileMethod="inclusive"` for quartile calculation~~ DONE

**Spec**: See `WOTAN/docs/chartex-spec.md#box--whisker`

**Next**: T-CHART-10 (DONE)

---

### B-CHART-09 [DONE]
**Create ChartEx infrastructure**

Build the foundational ChartEx support required by all Phase 2 charts.

**Details**:
- ~~Register `cx:` namespace in `ns.py`~~ DONE
- ~~Create `ChartExPart` class for `application/vnd.ms-office.chartex+xml`~~ DONE
- ~~Add `RT.CHART_EX` relationship type~~ DONE
- ~~Create OXML element classes~~ DONE:
  - CT_ChartExSpace (root)
  - CT_ChartExData, CT_ChartExChart
  - CT_ChartExPlotArea, CT_ChartExPlotAreaRegion
  - CT_ChartExSeries (with layoutId attribute)
  - CT_ChartExStrDim, CT_ChartExNumDim, CT_ChartExDataId
- ~~Create base ChartExXmlWriter~~ DONE (Treemap, Sunburst, Waterfall, Funnel, BoxWhisker)
- ~~Add ChartExData class for data handling~~ DONE

**Priority**: HIGH (blocks all Phase 2 charts)

**Next**: T-CHART-09 (DONE)

---

## Phase 3: Transitions and Animation

### B-TRANS-01 [DONE]
**Implement Morph transitions**

Add support for Morph transition type (PowerPoint 2016+).

**Details**:
- ~~Uses `p159:` namespace~~ DONE
- ~~`<p159:morph>` element (type CT_MorphTransition)~~ DONE
- ~~`option` attribute: `byObject`, `byWord`, `byChar`~~ DONE
- ~~Add transition property to Slide class~~ DONE
- ~~SlideTransition class with type, duration, morph_option properties~~ DONE
- ~~set_morph() method for easy configuration~~ DONE
- Object matching by name ("!!" prefix convention - documented for users)

**API**:
```python
slide.transition.set_morph(option="byObject", duration_ms=2000)
slide.transition.type  # "morph"
slide.transition.morph_option  # "byObject", "byWord", or "byChar"
slide.transition.duration  # 2000 (ms)
```

**Next**: T-TRANS-01 (DONE)

---

### B-TRANS-02 [DONE]
**Implement Zoom features**

Add support for Section Zoom, Slide Zoom, Summary Zoom (PowerPoint 2016+).

**Details**:
- ~~Register `p166:` namespace~~ DONE
- ~~Register `pslz:` namespace (slide zoom)~~ DONE
- ~~CT_SlideZoom, CT_SlideZoomObject, CT_ZoomObjectProperties OXML elements~~ DONE
- ~~MSO_SHAPE_TYPE.SLIDE_ZOOM enum value~~ DONE
- ~~GraphicFrame.has_slide_zoom property~~ DONE
- ~~GraphicFrame.shape_type returns SLIDE_ZOOM for zoom shapes~~ DONE
- Zoom shapes preserved on round-trip

**API**:
```python
# Detection
shape.has_slide_zoom  # True if shape is a Slide Zoom
shape.shape_type == MSO_SHAPE_TYPE.SLIDE_ZOOM

# Creating new zoom shapes requires thumbnail generation
# which is complex - preserving existing zooms is supported
```

**Next**: T-TRANS-02 (DONE - read/preserve support)

---

## Phase 4: Advanced Features

### B-MEDIA-01 [DONE]
**Add 3D model support**

Enable reading and preserving 3D models (Office 2017+).

**Details**:
- ~~Uses `am3d:` namespace (`http://schemas.microsoft.com/office/drawing/2017/model3d`)~~ DONE
- ~~GLB content type and default mapping~~ DONE
- ~~MSO_SHAPE_TYPE.MODEL_3D enum value~~ DONE
- ~~GraphicFrame.has_model_3d property~~ DONE
- ~~shape_type returns MODEL_3D for 3D model shapes~~ DONE
- 3D models preserved on round-trip

**Next**: T-MEDIA-01 (DONE)

---

### B-SHAPE-01 [DONE]
**Add Bezier curve support to FreeformBuilder**

Enable creating curved freeform shapes with cubic Bezier segments.

**Details**:
- ~~Currently only MoveTo, LineTo, Close supported~~ DONE
- ~~Add `cubicBezTo` path segment support~~ DONE
- ~~Update FreeformBuilder API~~ DONE

**Next**: T-SHAPE-01 (DONE)

---

### B-SHAPE-02 [DONE]
**Implement Ink annotations**

Add support for reading/preserving ink annotations (InkML format).

**Details**:
- ~~Content type: `application/inkml+xml`~~ DONE (already existed as CONTENT_TYPE.INK)
- ~~Namespace: `http://www.w3.org/2003/InkML`~~ DONE (inkml:)
- ~~MSO_SHAPE_TYPE.INK and INK_COMMENT enum values~~ DONE (already existed)
- Ink annotations preserved on round-trip

**Next**: T-SHAPE-02 (DONE)

---

### B-MEDIA-02 [BLOCKED]
**Add Cameo (live camera) support**

Support for camera feed placeholder (Microsoft 365 2022+).

**Details**:
- Camera placeholder shape (special placeholder type)
- Can be added to Slide Masters
- Integrates with Teams PowerPoint Live
- Limited to direct cameras (not virtual cameras)
- One video feed per slide

**Spec**: See `WOTAN/docs/needs-spec-research.md#b-media-02-cameo-live-camera`

**Priority**: VERY LOW - Minimal public documentation, very new feature

**Blocked by**: Lack of public XML schema documentation

**Next**: None

---

## Maintenance and Infrastructure

### B-TEST-01 [DONE]
**Create test PPTX files with modern features**

Build a collection of test files containing modern PowerPoint features for testing.

**Details**:
- ~~Create PPTX files for each modern feature~~ DONE
- ~~Python script to generate test files~~ DONE (create_test_files.py)
- ~~Document expected behavior~~ DONE (README.md)

**Test Files Created**:
- `morph-transition-test.pptx` - Morph transitions demo
- `chartex-test.pptx` - All ChartEx chart types
- `stock-chart-generated.pptx` - Stock charts (HLC, OHLC)
- `surface-chart-generated.pptx` - Surface charts (3D, wireframe)
- `transitions-test.pptx` - Various transition types
- `bezier-generated.pptx` - Bezier curve freeforms

**Location**: `WOTAN/example-docs/test-features/`

**Next**: T-TEST-01 (DONE)

---

### B-DOC-01 [DONE]
**Document new features as implemented**

Update documentation for each new feature added.

**Details**:
- ~~API documentation~~ DONE
- ~~Usage examples~~ DONE
- ~~Update feature matrix~~ DONE

**Documentation Added**:
- Feature Support section updated in index.rst
- Modern Charts section in user/charts.rst (ChartEx examples)
- Slide Transitions section in user/slides.rst (Morph API)
- Shape types updated in user/understanding-shapes.rst
- MSO_SHAPE_TYPE enum docs updated (MODEL_3D, SLIDE_ZOOM)
- XL_CHARTEX_TYPE enum docs added

**Next**: T-DOC-01 (DONE)
