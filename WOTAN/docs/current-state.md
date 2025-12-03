# Current State: python-pptx xtend branch

## Overview

python-pptx is a mature, production-grade Python library for creating, reading, and updating PowerPoint (.pptx) files. The `xtend` branch extends the upstream v1.0.2 with support for modern PowerPoint features (2016+).

## Architecture

The codebase follows a layered architecture:

```
Public API (pptx.api.Presentation)
    ↓
Presentation/Slide Objects (pptx.presentation, pptx.slide)
    ↓
Shape/Chart/Table Objects (pptx.shapes, pptx.chart, pptx.table)
    ↓
DML (Drawing Markup Language) Objects (pptx.dml)
    ↓
OXML (Office Open XML) Element Classes (pptx.oxml)
    ↓
OPC (Open Packaging Convention) Package System (pptx.opc)
```

## Code Statistics

| Component | Files | Lines of Code |
|-----------|-------|---------------|
| Source Code | ~110 | ~30,000 |
| Unit Tests | 64+ | ~25,000 |
| Behavioral Tests | 58 features | ~2,500 |
| **Total** | ~240 | ~57,500 |

## Dependencies

- **lxml** (3.1.0+) - XML parsing and manipulation
- **Pillow** (3.3.2+) - Image handling
- **XlsxWriter** (0.5.7+) - Chart data Excel writer
- **typing_extensions** (4.9.0+) - Type hint support

## What Works Well

### Presentation Features
- Create/load presentations from files or defaults
- Get/set slide dimensions
- Access slide masters, layouts, and slides
- Document properties (title, author, subject, keywords, etc.)

### Slide Features
- Add/remove slides
- Access slide layouts and placeholders
- Slide background fill (solid, gradient, pattern, image)
- Named slides
- Slide notes
- **Morph transitions (xtend)** - Full API for byObject/byWord/byChar

### Shape Features
- **AutoShapes**: 190+ types (rectangles, circles, arrows, stars, callouts, flowchart, action buttons)
- **Text Shapes**: Text frames with full paragraph/character formatting
- **Pictures**: Add, crop, resize images (JPG, PNG, BMP, GIF, TIFF, EMF, WMF)
- **SVG Images (xtend)**: Vector graphics with automatic PNG fallback
- **Group Shapes**: Group/ungroup multiple shapes
- **Connectors**: Lines, arrows with connection points
- **Freeform Shapes**: Custom shapes with straight lines and **Bezier curves (xtend)**
- **Media**: Embed video/audio files
- **Placeholders**: Title, content, picture, table, chart placeholders with inheritance

### Chart Features
- Full chart data manipulation (XlsxWriter integration)
- **Traditional charts**: column, bar, line, pie, doughnut, area, scatter, bubble, radar
- **Stock charts (xtend)**: HLC, OHLC, VHLC, VOHLC
- **Surface charts (xtend)**: 3D surface, wireframe, top view
- **Modern ChartEx charts (xtend)**: Treemap, Sunburst, Waterfall, Funnel, Box & Whisker, Map
- 3D variants: Area3D, Bar3D, Column3D, Line3D, Pie3D
- Axis customization: scaling, crossing, tick marks, labels, gridlines
- Series formatting: color, dash style, marker style/size
- Data labels with custom format
- Legend positioning and content
- Chart title and axis titles

### Table Features
- Create tables with custom rows/columns
- Cell formatting: fill, borders, text alignment, vertical anchor
- Cell text with full paragraph/character formatting
- Column width and row height management
- First row/column formatting, banding

### Text Features
- Paragraphs with alignment, indentation, spacing, bullets
- Character formatting: font, size, bold, italic, underline, color
- Text anchor options (top, middle, bottom, justified, distributed)
- Auto-fit options (none, shape-to-fit, text-to-fit)
- Word wrap control

### Drawing/Fill Features
- **Fill Types**: Solid, gradient, pattern, picture, group fill, no fill
- **Gradient**: Linear, radial, rectangular; multiple stops
- **Colors**: RGB, scheme colors (theme colors), system colors, HSL, preset colors
- **Color Transformations**: Shade, tint, sat mod, lum mod, alpha modification
- **Lines/Strokes**: Width, dash styles, cap/join styles, arrow styles

### Hyperlinks & Actions
- Click actions (hyperlink, open file, start document, custom action)
- Hover actions
- Action button support

### Media Support
- Image formats: JPEG, PNG, BMP, GIF, TIFF, EMF, WMF, **SVG (xtend)**
- Video formats: ASF, AVI, MOV, MP4, MPEG, SWF, WMV

### Modern Features (xtend)
- **SmartArt**: Read structure and extract text content
- **3D Models**: Detected (MODEL_3D shape type), preserved on round-trip
- **Ink Annotations**: Preserved on round-trip
- **Slide Zoom**: Detected (SLIDE_ZOOM shape type), preserved on round-trip

## OXML Coverage

The library has 200+ custom element classes covering:

| Domain | Elements |
|--------|----------|
| Presentation | CT_Presentation, CT_SlideIdList, CT_SlideMasterIdList |
| Slides | CT_Slide, CT_SlideLayout, CT_SlideMaster, CT_NotesSlide |
| Shapes | CT_Shape, CT_Picture, CT_GraphicalObjectFrame, CT_GroupShape, CT_Connector |
| Tables | CT_Table, CT_TableGrid, CT_TableRow, CT_TableCell |
| Text | CT_TextBody, CT_TextParagraph, CT_TextRun, CT_CharacterProperties |
| Fill | CT_SolidColorFillProperties, CT_GradientFillProperties, CT_PatternFillProperties |
| Color | CT_SRgbColor, CT_SchemeColor, CT_HslColor |
| Charts | Full chart element hierarchy |
| ChartEx (xtend) | CT_ChartExSpace, CT_ChartExSeries, CT_ChartExData |
| Transitions (xtend) | CT_SlideTransition, CT_MorphTransition |

## XML Namespaces Supported

- `p:` - PresentationML (presentations, slides)
- `a:` - DrawingML (shapes, text, colors)
- `r:` - Relationships
- `c:` - Charts
- `cx:` - ChartEx (modern charts) **(xtend)**
- `pic:` - Pictures
- `asvg:` - SVG images **(xtend)**
- `dgm:` - Diagrams/SmartArt **(xtend)**
- `p159:` - PowerPoint 2015 (Morph) **(xtend)**
- `p166:` - PowerPoint 2016 (Zoom) **(xtend)**
- `am3d:` - 3D Models **(xtend)**
- `inkml:` - Ink annotations **(xtend)**

## Test Coverage

- **Unit Tests**: 2700+ tests covering all major modules
- **BDD Tests**: 58 Gherkin feature files with step implementations
- Testing frameworks: pytest, behave
- Test fixtures in `tests/unit/unitdata/`
- All tests passing
