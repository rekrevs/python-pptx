# Gap Analysis: python-pptx vs Modern PPTX

## Executive Summary

python-pptx covers ~70% of core PPTX functionality well, but lacks support for features introduced after ~2016. The library handles PowerPoint 2007-2013 era features comprehensively but has significant gaps for PowerPoint 2016, 2019, 2021, and Microsoft 365 features.

## OOXML Specification Status

**Current Standard**: ISO/IEC 29500:2016
- Part 1: Fundamentals and Markup Language Reference
- Part 2: Open Packaging Conventions (updated 2021)
- Part 3: Markup Compatibility and Extensibility
- Part 4: Transitional Migration Features

**Microsoft Extensions**: MS-PPTX specification adds proprietary extensions.

## Critical Gaps

### 1. SmartArt - HIGH PRIORITY

**Status**: Detection only, no creation/modification

**Impact**: Very commonly used in business presentations

**Current Code**: `src/pptx/shapes/graphfrm.py` lines 32-35, 105
- GraphicFrame can detect SmartArt (returns `None` for `shape_type`)
- No API to create, read structure, or modify SmartArt

**OOXML Elements Needed**:
- `dgm:` namespace (diagrams)
- `dsp:` namespace (diagram shapes)
- CT_DiagramDefinition, CT_DiagramData, CT_DiagramColors, CT_DiagramStyle

**Effort**: HIGH - Complex nested XML structure with data model, layout, colors, and style components

---

### 2. SVG Images - HIGH PRIORITY

**Status**: Explicitly skipped/ignored

**Impact**: Modern vector graphics standard, widely used

**Current Code**: `src/pptx/package.py` line 164 - SVG images are skipped

**OOXML Elements Needed**:
- `a:svgBlip` element for SVG references
- SVG content type registration
- `a16:` namespace for DrawingML 2016 extensions

**Files to Modify**:
- `src/pptx/parts/image.py`: Add SVG content type detection
- `src/pptx/package.py`: Remove SVG skip logic
- `src/pptx/oxml/shapes/picture.py`: Handle `a:svgBlip` element
- `src/pptx/enum/`: Add MSO_PICTURE_TYPE.SVG

**Effort**: MEDIUM

---

### 3. Modern Chart Types - MEDIUM PRIORITY

**Unsupported Chart Types** (enums exist but no implementation):

| Chart Type | Since | OOXML Element |
|------------|-------|---------------|
| Stock (HLC, OHLC, VHLC, VOHLC) | 2007 | CT_StockChart |
| Surface | 2007 | CT_SurfaceChart |
| Treemap | 2016 | CT_Treemap (c16:) |
| Sunburst | 2016 | CT_Sunburst (c16:) |
| Waterfall | 2016 | CT_Waterfall (c16:) |
| Histogram | 2016 | CT_Histogram (c16:) |
| Pareto | 2016 | CT_Pareto (c16:) |
| Box & Whisker | 2016 | CT_BoxWhisker (c16:) |
| Funnel | 2019 | CT_Funnel |
| Map | 2019 | CT_Map |

**Current Code**: `src/pptx/chart/xmlwriter.py` line 52 raises `NotImplementedError`

**Files to Modify**:
- `src/pptx/chart/xmlwriter.py`: Add XML writers for new chart types
- `src/pptx/chart/plot.py`: Add plot classes
- `src/pptx/oxml/chart/`: Add CT_* element classes
- `src/pptx/chart/data.py`: Add data classes if needed

**Effort**: MEDIUM per chart type

---

### 4. Morph Transitions - MEDIUM PRIORITY

**Status**: Not implemented

**Impact**: Key modern feature since PowerPoint 2019

**OOXML Elements Needed**:
- `p14:transition` with morph settings
- `p14:` namespace (PowerPoint 2010+ extensions)

**Files to Add**:
- `src/pptx/oxml/transition.py`: CT_Transition with morph support
- `src/pptx/slide.py`: Add transition property to Slide class

**Effort**: MEDIUM

---

### 5. 3D Models - MEDIUM PRIORITY

**Status**: Not supported

**Impact**: Growing use in modern presentations

**OOXML Elements Needed**:
- `a3d:model3d` elements
- 3D model part relationships

**Effort**: HIGH - Complex 3D model handling and format support

---

### 6. Freeform Bezier Curves - LOW-MEDIUM PRIORITY

**Status**: Only straight lines supported

**Current Code**: `docs/dev/analysis/shp-freeform.rst` lines 59-68
- MoveTo, LineTo, and Close are supported
- `cubicBezTo` NOT supported for creation

**Impact**: Custom shape creation limited

**Effort**: MEDIUM - Add cubicBezTo support to FreeformBuilder

---

### 7. Ink Annotations - LOW PRIORITY

**Status**: Not supported

**OOXML Location**: `ppt/ink/` folder, uses Ink ML

**Effort**: MEDIUM

---

### 8. Zoom Features - LOW PRIORITY

**Status**: Not supported

**Types**: Section zoom, slide zoom, summary zoom

**OOXML Elements**: `p14:` namespace

**Effort**: MEDIUM

---

### 9. Cameo (Live Camera) - LOW PRIORITY

**Status**: Not supported

**Impact**: Very new feature (2022), Microsoft 365 only

**Effort**: MEDIUM

---

### 10. Icons Library Integration - LOW PRIORITY

**Status**: Not supported

**Impact**: Microsoft 365 feature, icons are essentially SVG

**Effort**: LOW (once SVG is supported)

---

## Minor Gaps and Limitations

### Shape Features
- `BaseShape.is_connector` not implemented
- Shadow property on GraphicFrame raises `NotImplementedError`
- Connection points not fully accessible

### Chart Features
- Individual legend entry customization not supported
- Chart titles from Excel cell references not supported
- Some data label configurations limited

### Color/Fill
- Not all fill types support all color operations
- Some color type combinations raise `NotImplementedError`

### Font/Text
- Font parsing limited to OTF/TTF
- Some OS font systems not supported

### Actions
- Start other presentation actions not supported

---

## Coverage Summary

| Category | Supported | Missing | Coverage |
|----------|-----------|---------|----------|
| Basic Shapes | 190+ | 0 | 100% |
| Chart Types | ~15 | ~12 | ~55% |
| Transitions | Basic | Morph, 3D | ~70% |
| Media Types | 8 | SVG, 3D Models | ~80% |
| SmartArt | 0 | All | 0% |
| Modern Features (2016+) | 0 | ~8 | 0% |

---

## Prioritized Implementation Roadmap

### Phase 1: Critical Modern Features
1. **SVG Support** (P0) - Straightforward, high impact
2. **SmartArt Read Support** (P0) - High demand, enables round-trip
3. **Stock/Surface Charts** (P1) - Complete existing enum coverage

### Phase 2: PowerPoint 2016+ Charts
4. **Treemap Charts** (P1)
5. **Sunburst Charts** (P1)
6. **Waterfall Charts** (P1)
7. **Funnel Charts** (P1)
8. **Map Charts** (P2)
9. **Histogram/Box & Whisker** (P2)

### Phase 3: Transitions and Animation
10. **Morph Transitions** (P1)
11. **Zoom Features** (P2)

### Phase 4: Advanced Features
12. **3D Models** (P2)
13. **Bezier Curves in Freeforms** (P2)
14. **Ink Annotations** (P3)
15. **Cameo** (P3)
