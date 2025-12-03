# python-pptx API Reference

This document provides a comprehensive API reference for python-pptx, including all extensions in the xtend branch.

## Table of Contents

1. [Opening and Saving Presentations](#opening-and-saving-presentations)
2. [Working with Slides](#working-with-slides)
3. [Shapes](#shapes)
4. [Text and Paragraphs](#text-and-paragraphs)
5. [Charts](#charts)
6. [Modern Charts (ChartEx)](#modern-charts-chartex)
7. [Tables](#tables)
8. [Images and Media](#images-and-media)
9. [SmartArt](#smartart)
10. [Transitions](#transitions)
11. [Colors and Fills](#colors-and-fills)
12. [Enumerations](#enumerations)

---

## Opening and Saving Presentations

### Creating a New Presentation

```python
from pptx import Presentation

# Create blank presentation
prs = Presentation()

# Create from template
prs = Presentation('template.pptx')
```

### Saving a Presentation

```python
prs.save('output.pptx')

# Save to file-like object
from io import BytesIO
stream = BytesIO()
prs.save(stream)
```

### Presentation Properties

```python
# Slide dimensions
prs.slide_width   # Emu (914400 = 1 inch)
prs.slide_height  # Emu

# Core properties
prs.core_properties.title
prs.core_properties.author
prs.core_properties.subject
prs.core_properties.keywords
prs.core_properties.comments
prs.core_properties.created
prs.core_properties.modified
```

---

## Working with Slides

### Accessing Slides

```python
# All slides
slides = prs.slides

# By index
slide = prs.slides[0]

# Iterate
for slide in prs.slides:
    print(slide.slide_id)
```

### Adding Slides

```python
# Get layout (0=Title, 1=Title+Content, 5=Blank, 6=Title Only)
slide_layout = prs.slide_layouts[1]

# Add slide
slide = prs.slides.add_slide(slide_layout)
```

### Slide Properties

```python
slide.shapes        # Shape collection
slide.placeholders  # Placeholder collection
slide.slide_layout  # Associated layout
slide.notes_slide   # Notes page
slide.name          # Slide name
slide.slide_id      # Unique ID
```

### Slide Transitions (xtend)

```python
# Add Morph transition
slide.transition.set_morph(option="byObject", duration_ms=1500)

# Check transition
slide.transition.type         # "morph", "fade", etc.
slide.transition.morph_option # "byObject", "byWord", "byChar"
slide.transition.duration     # milliseconds
```

---

## Shapes

### Accessing Shapes

```python
shapes = slide.shapes

# By index
shape = shapes[0]

# By name
for shape in shapes:
    if shape.name == "MyShape":
        # found it
        pass
```

### Adding Shapes

```python
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE

# AutoShape
shape = shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(1), Inches(1),  # left, top
    Inches(2), Inches(1)   # width, height
)

# Text box
textbox = shapes.add_textbox(
    Inches(1), Inches(1),
    Inches(3), Inches(0.5)
)

# Picture
picture = shapes.add_picture(
    'image.png',
    Inches(1), Inches(1),
    width=Inches(2)  # height auto-calculated
)

# SVG with PNG fallback (xtend)
svg_picture = shapes.add_svg_picture(
    'image.svg',
    Inches(1), Inches(1),
    fallback_image='fallback.png'  # optional
)
```

### Shape Properties

```python
# Position and size
shape.left      # Emu from left edge
shape.top       # Emu from top edge
shape.width     # Emu
shape.height    # Emu

# Identification
shape.shape_id  # Unique ID
shape.name      # Shape name
shape.shape_type  # MSO_SHAPE_TYPE enum

# Rotation
shape.rotation  # Degrees (0-360)

# Text
if shape.has_text_frame:
    text_frame = shape.text_frame
```

### Freeform Shapes

```python
# Create custom shape with lines
builder = shapes.build_freeform(Inches(1), Inches(1))
builder.add_line_segments([
    (Inches(2), Inches(1)),
    (Inches(2), Inches(2)),
    (Inches(1), Inches(2)),
], close=True)
shape = builder.convert_to_shape()

# With Bezier curves (xtend)
builder = shapes.build_freeform(Inches(1), Inches(1))
builder.add_line_segments([(Inches(2), Inches(1))])
builder.add_bezier(
    (Inches(3), Inches(1)),   # control point 1
    (Inches(3), Inches(2)),   # control point 2
    (Inches(2), Inches(2))    # end point
)
builder.close()
shape = builder.convert_to_shape()
```

### Shape Type Detection (xtend)

```python
from pptx.enum.shapes import MSO_SHAPE_TYPE

shape.shape_type == MSO_SHAPE_TYPE.PICTURE      # Image
shape.shape_type == MSO_SHAPE_TYPE.CHART        # Chart
shape.shape_type == MSO_SHAPE_TYPE.TABLE        # Table
shape.shape_type == MSO_SHAPE_TYPE.DIAGRAM      # SmartArt
shape.shape_type == MSO_SHAPE_TYPE.MODEL_3D     # 3D Model (xtend)
shape.shape_type == MSO_SHAPE_TYPE.SLIDE_ZOOM   # Slide Zoom (xtend)
shape.shape_type == MSO_SHAPE_TYPE.INK          # Ink annotation
```

---

## Text and Paragraphs

### Text Frames

```python
text_frame = shape.text_frame

# Simple text
text_frame.text = "Hello World"

# Paragraphs
for paragraph in text_frame.paragraphs:
    for run in paragraph.runs:
        print(run.text)
```

### Paragraphs

```python
from pptx.enum.text import PP_ALIGN

paragraph = text_frame.paragraphs[0]

# Alignment
paragraph.alignment = PP_ALIGN.CENTER

# Spacing
paragraph.space_before = Pt(12)
paragraph.space_after = Pt(6)
paragraph.line_spacing = 1.5  # or Pt(18)

# Indentation
paragraph.level = 1  # Bullet level (0-8)
```

### Character Formatting

```python
from pptx.util import Pt
from pptx.dml.color import RGBColor

run = paragraph.add_run()
run.text = "Formatted text"

font = run.font
font.name = "Arial"
font.size = Pt(14)
font.bold = True
font.italic = True
font.underline = True
font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
```

### Hyperlinks

```python
run.hyperlink.address = "https://example.com"
```

---

## Charts

### Adding a Chart

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

chart_data = CategoryChartData()
chart_data.categories = ['East', 'West', 'Midwest']
chart_data.add_series('Q1 Sales', (19.2, 21.4, 16.7))
chart_data.add_series('Q2 Sales', (22.3, 28.6, 15.2))

x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
graphic_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
)
chart = graphic_frame.chart
```

### Chart Types

```python
# Column/Bar
XL_CHART_TYPE.COLUMN_CLUSTERED
XL_CHART_TYPE.COLUMN_STACKED
XL_CHART_TYPE.BAR_CLUSTERED

# Line
XL_CHART_TYPE.LINE
XL_CHART_TYPE.LINE_MARKERS

# Pie/Doughnut
XL_CHART_TYPE.PIE
XL_CHART_TYPE.DOUGHNUT

# Area
XL_CHART_TYPE.AREA
XL_CHART_TYPE.AREA_STACKED

# Scatter/Bubble
XL_CHART_TYPE.XY_SCATTER
XL_CHART_TYPE.BUBBLE

# Stock (xtend)
XL_CHART_TYPE.STOCK_HLC    # High-Low-Close
XL_CHART_TYPE.STOCK_OHLC   # Open-High-Low-Close

# Surface (xtend)
XL_CHART_TYPE.SURFACE
XL_CHART_TYPE.SURFACE_WIREFRAME
```

### Chart Customization

```python
# Title
chart.has_title = True
chart.chart_title.text_frame.text = "Sales Report"

# Legend
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM

# Axes
category_axis = chart.category_axis
value_axis = chart.value_axis

value_axis.maximum_scale = 50.0
value_axis.minimum_scale = 0
value_axis.has_major_gridlines = True

# Data labels
plot = chart.plots[0]
plot.has_data_labels = True
data_labels = plot.data_labels
data_labels.show_value = True
```

---

## Modern Charts (ChartEx)

These chart types use the ChartEx format introduced in Office 2016+.

### Creating ChartEx Charts (xtend)

```python
from pptx.chartex.data import ChartExData
from pptx.enum.chart import XL_CHARTEX_TYPE

chart_data = ChartExData()
chart_data.add_series(
    "Sales",
    ["Q1", "Q2", "Q3", "Q4"],
    [100, 150, 120, 180]
)

x, y, cx, cy = Inches(1), Inches(1), Inches(8), Inches(5)
slide.shapes.add_chartex(
    XL_CHARTEX_TYPE.TREEMAP, x, y, cx, cy, chart_data
)
```

### ChartEx Types

```python
XL_CHARTEX_TYPE.TREEMAP      # Hierarchical rectangles
XL_CHARTEX_TYPE.SUNBURST     # Hierarchical rings
XL_CHARTEX_TYPE.WATERFALL    # Running total
XL_CHARTEX_TYPE.FUNNEL       # Stage progression
XL_CHARTEX_TYPE.BOX_WHISKER  # Statistical distribution
XL_CHARTEX_TYPE.REGION_MAP   # Geographic map
```

### Waterfall with Subtotals

```python
chart_data = ChartExData()
chart_data.add_series(
    "Profit Analysis",
    ["Start", "Revenue", "Costs", "Tax", "End"],
    [100, 50, -30, -10, 110],
    subtotals=[4]  # Index 4 ("End") is a subtotal
)
```

### Map Chart

```python
chart_data = ChartExData()
chart_data.add_series(
    "Population (millions)",
    ["USA", "Germany", "Japan", "Brazil"],
    [331, 83, 125, 213]
)

slide.shapes.add_chartex(
    XL_CHARTEX_TYPE.REGION_MAP, x, y, cx, cy, chart_data
)
```

---

## Tables

### Adding a Table

```python
rows, cols = 3, 4
table = slide.shapes.add_table(
    rows, cols,
    Inches(1), Inches(1),
    Inches(6), Inches(2)
).table
```

### Accessing Cells

```python
cell = table.cell(0, 0)  # row, col
cell.text = "Header"

# Merge cells
cell.merge(table.cell(0, 1))
```

### Cell Formatting

```python
from pptx.enum.text import MSO_ANCHOR

cell.text_frame.paragraphs[0].font.bold = True
cell.vertical_anchor = MSO_ANCHOR.MIDDLE
cell.fill.solid()
cell.fill.fore_color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
```

---

## Images and Media

### Adding Images

```python
# From file
picture = shapes.add_picture('image.jpg', Inches(1), Inches(1))

# From stream
from io import BytesIO
image_stream = BytesIO(image_bytes)
picture = shapes.add_picture(image_stream, Inches(1), Inches(1))

# Specify size
picture = shapes.add_picture(
    'image.jpg',
    Inches(1), Inches(1),
    width=Inches(3),
    height=Inches(2)
)
```

### SVG Images (xtend)

```python
# SVG with automatic PNG fallback generation
svg_picture = shapes.add_svg_picture(
    'image.svg',
    Inches(1), Inches(1),
    width=Inches(3)
)
```

### Supported Image Formats

- JPEG (.jpg, .jpeg)
- PNG (.png)
- GIF (.gif)
- BMP (.bmp)
- TIFF (.tif, .tiff)
- EMF (.emf)
- WMF (.wmf)
- SVG (.svg) - xtend

### Video

```python
movie = shapes.add_movie(
    'video.mp4',
    Inches(1), Inches(1),
    Inches(4), Inches(3),
    poster_frame_image='thumbnail.jpg'
)
```

---

## SmartArt

SmartArt diagrams can be read and their text extracted (xtend).

### Detecting SmartArt

```python
for shape in slide.shapes:
    if shape.has_smart_art:
        smart_art = shape.smart_art

        # Get all text
        all_text = smart_art.all_text  # List of strings

        # Get formatted text
        text = smart_art.text  # Single string with newlines
```

### SmartArt Shape Type

```python
from pptx.enum.shapes import MSO_SHAPE_TYPE

if shape.shape_type == MSO_SHAPE_TYPE.DIAGRAM:
    # This is a SmartArt diagram
    pass
```

---

## Transitions

### Morph Transitions (xtend)

```python
# Set morph transition
slide.transition.set_morph(
    option="byObject",   # "byObject", "byWord", or "byChar"
    duration_ms=1500
)

# Read transition properties
slide.transition.type         # "morph"
slide.transition.morph_option # "byObject"
slide.transition.duration     # 1500
```

### Object Matching for Morph

For Morph to animate objects between slides, they must have the same name:

```python
# Slide 1
shape1 = slide1.shapes.add_shape(MSO_SHAPE.OVAL, ...)
shape1.name = "!!MyCircle"  # "!!" prefix for explicit matching

# Slide 2
shape2 = slide2.shapes.add_shape(MSO_SHAPE.OVAL, ...)
shape2.name = "!!MyCircle"  # Same name = will morph
```

---

## Colors and Fills

### Solid Fill

```python
from pptx.dml.color import RGBColor

shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)
```

### Theme Colors

```python
from pptx.enum.dml import MSO_THEME_COLOR

shape.fill.solid()
shape.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
```

### Gradient Fill

```python
from pptx.enum.dml import MSO_THEME_COLOR

shape.fill.gradient()
shape.fill.gradient_angle = 45.0
shape.fill.gradient_stops[0].color.rgb = RGBColor(0xFF, 0x00, 0x00)
shape.fill.gradient_stops[1].color.rgb = RGBColor(0x00, 0x00, 0xFF)
```

### Pattern Fill

```python
from pptx.enum.dml import MSO_PATTERN

shape.fill.patterned()
shape.fill.pattern = MSO_PATTERN.DARK_HORIZONTAL
shape.fill.fore_color.rgb = RGBColor(0x00, 0x00, 0x00)
shape.fill.back_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
```

### No Fill

```python
shape.fill.background()  # Transparent
```

### Line/Border

```python
from pptx.enum.dml import MSO_LINE_DASH_STYLE

line = shape.line
line.color.rgb = RGBColor(0x00, 0x00, 0x00)
line.width = Pt(2)
line.dash_style = MSO_LINE_DASH_STYLE.DASH_DOT
```

---

## Enumerations

### Shape Types (MSO_SHAPE_TYPE)

```python
from pptx.enum.shapes import MSO_SHAPE_TYPE

MSO_SHAPE_TYPE.AUTO_SHAPE      # AutoShape
MSO_SHAPE_TYPE.PICTURE         # Picture
MSO_SHAPE_TYPE.CHART           # Chart
MSO_SHAPE_TYPE.TABLE           # Table
MSO_SHAPE_TYPE.GROUP           # Group shape
MSO_SHAPE_TYPE.DIAGRAM         # SmartArt
MSO_SHAPE_TYPE.MEDIA           # Video/Audio
MSO_SHAPE_TYPE.FREEFORM        # Freeform shape
MSO_SHAPE_TYPE.TEXT_BOX        # Text box
MSO_SHAPE_TYPE.PLACEHOLDER     # Placeholder
MSO_SHAPE_TYPE.MODEL_3D        # 3D Model (xtend)
MSO_SHAPE_TYPE.SLIDE_ZOOM      # Slide Zoom (xtend)
MSO_SHAPE_TYPE.INK             # Ink annotation
MSO_SHAPE_TYPE.INK_COMMENT     # Ink comment
```

### AutoShape Types (MSO_SHAPE)

```python
from pptx.enum.shapes import MSO_SHAPE

# Basic shapes
MSO_SHAPE.RECTANGLE
MSO_SHAPE.OVAL
MSO_SHAPE.ROUNDED_RECTANGLE
MSO_SHAPE.TRIANGLE

# Arrows
MSO_SHAPE.RIGHT_ARROW
MSO_SHAPE.LEFT_ARROW
MSO_SHAPE.UP_ARROW
MSO_SHAPE.DOWN_ARROW

# Flowchart
MSO_SHAPE.FLOWCHART_PROCESS
MSO_SHAPE.FLOWCHART_DECISION
MSO_SHAPE.FLOWCHART_TERMINATOR

# Stars and banners
MSO_SHAPE.STAR_5_POINT
MSO_SHAPE.EXPLOSION1
MSO_SHAPE.HORIZONTAL_SCROLL

# ... 180+ shapes available
```

### Chart Types (XL_CHART_TYPE)

```python
from pptx.enum.chart import XL_CHART_TYPE

# See Charts section for full list
```

### ChartEx Types (XL_CHARTEX_TYPE) - xtend

```python
from pptx.enum.chart import XL_CHARTEX_TYPE

XL_CHARTEX_TYPE.TREEMAP
XL_CHARTEX_TYPE.SUNBURST
XL_CHARTEX_TYPE.WATERFALL
XL_CHARTEX_TYPE.FUNNEL
XL_CHARTEX_TYPE.BOX_WHISKER
XL_CHARTEX_TYPE.REGION_MAP
```

---

## Units

python-pptx uses English Metric Units (EMU) internally. Helper classes convert common units:

```python
from pptx.util import Inches, Cm, Mm, Pt, Emu

Inches(1)     # 914400 EMU
Cm(2.54)      # 914400 EMU
Mm(25.4)      # 914400 EMU
Pt(72)        # 914400 EMU
Emu(914400)   # 914400 EMU (no conversion)

# Convert back
width_inches = shape.width.inches
height_cm = shape.height.cm
```

---

## Pragmatics

### Round-Trip Preservation

python-pptx preserves elements it doesn't understand:

```python
# Open file with unsupported features
prs = Presentation('complex.pptx')

# Modify what you need
slide = prs.slides[0]
slide.shapes[0].text = "Modified"

# Save - unsupported features preserved
prs.save('modified.pptx')
```

### Preserved Features (xtend)

- 3D Models (GLB format)
- Ink annotations
- Slide zoom shapes
- SmartArt diagrams
- Modern chart types when read

### Best Practices

1. **Use templates**: Start from an existing .pptx with your branding
2. **Preserve placeholders**: Use placeholders for consistent layouts
3. **Check shape types**: Use `shape.shape_type` before accessing specific properties
4. **Handle missing properties**: Some properties return `None` if not set
5. **Test with PowerPoint**: Always verify output in actual PowerPoint

---

## Version Compatibility

| Feature | PowerPoint Version |
|---------|-------------------|
| Basic shapes, charts, tables | 2007+ |
| SVG images | 2016+ |
| ChartEx (Treemap, etc.) | 2016+ |
| Morph transitions | 2016+ |
| Map charts | 2019+ |
| 3D Models | 2019+ |
| Slide Zoom | 2016+ |

Files using modern features will open in older PowerPoint versions but those features may not render or be editable.
