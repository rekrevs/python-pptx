# Gap Analysis: python-pptx xtend vs Modern PPTX

## Executive Summary

The xtend branch has comprehensive coverage of modern PPTX features. The library handles PowerPoint 2007-2024 features with near-complete chart support, full theme access, shape effects, and rich text formatting.

## Current State

| Category | Coverage |
|----------|----------|
| Basic Shapes | 100% |
| Chart Types | ~99% (all standard + ChartEx + volume stock) |
| Modern Features (2016+) | ~90% |
| SmartArt | Read: 100%, Write: 0% |
| Text Formatting | ~90% (bullets, numbering done) |
| Theme/Styling | ~85% (color scheme, font scheme done) |
| Shape Effects | ~80% (shadow, glow, reflection, soft edge done) |
| Document Properties | 100% (core + custom) |
| Comments | 100% (read/write) |
| Accessibility | 100% (alt text on all shapes) |

---

## Features Implemented in xtend

### SVG Images ✅ DONE
- Full read/write support
- `shapes.add_svg_picture()` API
- Automatic PNG fallback for older PowerPoint versions
- `asvg:` namespace registered

### SmartArt ✅ DONE (Read Support)
- Detect SmartArt via `shape.has_smart_art`
- `shape.shape_type == MSO_SHAPE_TYPE.DIAGRAM`
- Extract text: `shape.smart_art.text` and `shape.smart_art.all_text`
- `dgm:` and `dsp:` namespaces registered
- Preserved on round-trip

### Stock Charts ✅ DONE (All Variants)
- `XL_CHART_TYPE.STOCK_HLC` (High-Low-Close)
- `XL_CHART_TYPE.STOCK_OHLC` (Open-High-Low-Close)
- `XL_CHART_TYPE.STOCK_VHLC` (Volume-High-Low-Close) — multi-plot with bar + stock
- `XL_CHART_TYPE.STOCK_VOHLC` (Volume-Open-High-Low-Close) — multi-plot with bar + stock
- `StockPlot` and `StockSeries` classes
- Full read/write support including secondary axis for volume

### Surface Charts ✅ DONE
- `XL_CHART_TYPE.SURFACE` (3D Surface)
- `XL_CHART_TYPE.SURFACE_WIREFRAME`
- `XL_CHART_TYPE.SURFACE_TOP_VIEW`
- `XL_CHART_TYPE.SURFACE_TOP_VIEW_WIREFRAME`
- Full XML writer support

### Modern ChartEx Charts ✅ DONE
All using `cx:` namespace (ChartEx infrastructure):

| Chart Type | Enum | Status |
|------------|------|--------|
| Treemap | `XL_CHARTEX_TYPE.TREEMAP` | ✅ |
| Sunburst | `XL_CHARTEX_TYPE.SUNBURST` | ✅ |
| Waterfall | `XL_CHARTEX_TYPE.WATERFALL` | ✅ (with subtotals) |
| Funnel | `XL_CHARTEX_TYPE.FUNNEL` | ✅ |
| Box & Whisker | `XL_CHARTEX_TYPE.BOX_WHISKER` | ✅ |
| Map (Region) | `XL_CHARTEX_TYPE.REGION_MAP` | ✅ |

Infrastructure:
- `ChartExPart` for `application/vnd.ms-office.chartex+xml`
- `ChartExData` class for data handling
- `ChartExXmlWriter` base with type-specific subclasses
- BDD acceptance tests for all 6 types

### Morph Transitions ✅ DONE
- `slide.transition.set_morph(option, duration_ms)`
- Options: `byObject`, `byWord`, `byChar`
- `slide.transition.type`, `.morph_option`, `.duration`
- `p159:` namespace registered

### Theme Access ✅ DONE
- `slide_master.theme` returns `Theme` object
- `theme.color_scheme` with 12 named colors (dark1, light1, accent1-6, hyperlink, etc.)
- Colors readable/writable as `RGBColor`
- Handles both `srgbClr` and `sysClr` elements
- `theme.font_scheme` with `major_font`/`minor_font` (read/write)
- `ThemePart(XmlPart)` properly registered

### Shape Effects ✅ DONE
- `shape.shadow` — expanded with `shadow_type`, `angle`, `blur_radius`, `distance`, `color`
- `shape.glow` — `radius`, `color`
- `shape.reflection` — `blur_radius`, `distance`, `direction`, `start_opacity`, `end_opacity`
- `shape.soft_edge` — `radius`
- OXML classes for all effect list elements
- Existing `.inherit` behavior preserved

### Bullet and Numbering ✅ DONE
- `paragraph.bullet` returns `BulletFormat`
- `bullet.type` — none, char, auto_num (read/write)
- `bullet.char` — character bullet
- `bullet.auto_num_type` — numbering format
- `bullet.font`, `bullet.size`, `bullet.color`
- 10 OXML element classes for bullet elements

### Custom Document Properties ✅ DONE
- `prs.custom_properties` — dict-like interface
- Supports string, int, float, bool, datetime types
- Type inference on write
- CRUD operations (get, set, delete, iterate)

### Slide Comments ✅ DONE
- `slide.comments` — list of Comment objects
- `slide.add_comment(text, author_name)` — create comments
- `comment.text`, `comment.author`, `comment.timestamp`
- Two-part architecture: per-slide comments + presentation-level authors

### Alt Text ✅ DONE
- `shape.alt_text` — read/write on all shape types
- Exposes `cNvPr` `descr` attribute

### Color Transparency ✅ DONE
- `color.alpha` — read/write float 0.0–1.0
- Works with solid fills and font colors
- Alpha = 1.0 removes element (PowerPoint default)

### Slide Zoom ✅ DONE (Read/Preserve)
- `shape.has_slide_zoom` detection
- Preserved on round-trip

### 3D Models ✅ DONE (Read/Preserve)
- `shape.has_model_3d` detection
- Preserved on round-trip

### Ink Annotations ✅ DONE (Preserve)
- Preserved on round-trip

### Bezier Curves ✅ DONE
- `FreeformBuilder.add_bezier()` method
- Full `cubicBezTo` support

---

## Remaining Gaps

### SmartArt Write Support - MEDIUM PRIORITY
**Status**: Read-only, cannot create or modify
**Effort**: HIGH - Complex nested XML with data model, layout, colors, style

### Animations - LOW PRIORITY
**Status**: Preserved on round-trip, no API
**Effort**: HIGH - Complex timing model

### Cameo (Live Camera) - VERY LOW PRIORITY
**Status**: Not supported
**Blocked by**: Lack of public XML schema documentation

### OLE Objects - LOW PRIORITY
**Status**: Analysis doc exists (`docs/dev/analysis/shp-ole-object.rst`), no API
**Effort**: MEDIUM-HIGH

### Equations - LOW PRIORITY
**Status**: Not supported
**Effort**: MEDIUM

### Text Effects - LOW PRIORITY
**Status**: Not supported
**Effort**: MEDIUM

---

## Coverage Summary

| Category | Before xtend | After xtend |
|----------|--------------|-------------|
| Basic Shapes | 100% | 100% |
| Chart Types | ~55% | ~99% |
| Transitions | 0% | ~90% (Morph done) |
| Media Types | ~80% | ~95% (SVG, 3D detect) |
| SmartArt | 0% | ~60% (read) |
| Modern Features (2016+) | 0% | ~90% |
| Theme/Styling | 0% | ~85% |
| Shape Effects | ~5% | ~80% |
| Text Formatting | ~60% | ~90% |
| Document Properties | ~70% | 100% |
| Comments | 0% | 100% |
| Accessibility | 0% | 100% |

---

## Test Coverage

- **3021 unit tests** (pytest)
- **1032 BDD scenarios** (behave)
- **59 feature files**

## Conclusion

The xtend branch has achieved comprehensive coverage:

1. **Complete chart support** — All standard types + ChartEx + volume stock charts
2. **SVG images** — Full read/write with fallback
3. **SmartArt** — Read support with text extraction
4. **Theme access** — Color scheme and font scheme read/write
5. **Shape effects** — Shadow, glow, reflection, soft edge with full property access
6. **Rich text** — Bullet and numbering formatting
7. **Document metadata** — Custom properties with typed values
8. **Collaboration** — Slide comments read/write
9. **Accessibility** — Alt text on all shape types
10. **Color control** — Transparency/alpha support
11. **Round-trip fidelity** — 3D models, ink, zoom preserved

Remaining gaps are either:
- High complexity with limited demand (SmartArt write, animations)
- Blocked by external factors (Cameo lacks docs)
- Niche use cases (OLE objects, equations)
