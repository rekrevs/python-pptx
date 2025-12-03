# python-pptx (xtend fork)

This is an **experimental fork** of [python-pptx](https://github.com/python-openxml/python-pptx) with extensions to handle more of the full OOXML (.pptx) format.

The upstream *python-pptx* library provides excellent support for basic presentation operations. This fork extends it with support for modern PowerPoint features (2016+) and improved round-trip fidelity for complex presentations.

## Extensions

This fork adds the following capabilities:

| Feature | Read | Write | Description |
|---------|:----:|:-----:|-------------|
| **SVG Images** | ✓ | ✓ | Vector graphics with automatic PNG fallback |
| **Modern Charts (ChartEx)** | ✓ | ✓ | Treemap, Sunburst, Waterfall, Funnel, Box & Whisker, Map |
| **Stock Charts** | ✓ | ✓ | HLC, OHLC, VHLC, VOHLC stock chart types |
| **Surface Charts** | ✓ | ✓ | 3D surface and wireframe charts |
| **SmartArt** | ✓ | | Read diagram structure and extract text content |
| **Morph Transitions** | ✓ | ✓ | Smooth object/word/character morphing between slides |
| **Slide Zoom** | ✓ | | Interactive zoom navigation shapes |
| **3D Models** | ✓ | | GLB/glTF format models preserved on round-trip |
| **Ink Annotations** | ✓ | | Pen/stylus annotations preserved on round-trip |
| **Bezier Curves** | ✓ | ✓ | Cubic Bezier curves in freeform shapes |

## Installation

```bash
pip install git+https://github.com/sverker/python-pptx.git@xtend
```

## Example

```python
>>> from pptx import Presentation

>>> prs = Presentation()
>>> slide = prs.slides.add_slide(prs.slide_layouts[6])
>>> slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
<pptx.shapes.autoshape.Shape object at 0x...>
>>> prs.save("hello.pptx")

>>> prs = Presentation("hello.pptx")
>>> prs.slides[0].shapes[0].text_frame.text
''
```

### Extension Examples

```python
>>> from pptx import Presentation
>>> from pptx.util import Inches
>>> from pptx.enum.chart import XL_CHARTEX_TYPE
>>> from pptx.chartex.data import ChartExData

# Create a modern Treemap chart
>>> prs = Presentation()
>>> slide = prs.slides.add_slide(prs.slide_layouts[5])
>>> chart_data = ChartExData()
>>> chart_data.add_series("Sales", ["Q1", "Q2", "Q3", "Q4"], [100, 150, 120, 180])
>>> slide.shapes.add_chartex(
...     XL_CHARTEX_TYPE.TREEMAP,
...     Inches(1), Inches(1), Inches(8), Inches(5),
...     chart_data
... )

# Add Morph transition
>>> slide2 = prs.slides.add_slide(prs.slide_layouts[5])
>>> slide2.transition.set_morph(option="byObject", duration_ms=1500)

# Read SmartArt content
>>> prs = Presentation("presentation-with-smartart.pptx")
>>> for shape in prs.slides[0].shapes:
...     if shape.has_smart_art:
...         print(shape.smart_art.text)

# Add SVG image
>>> slide.shapes.add_svg_picture("logo.svg", Inches(1), Inches(1))

# Detect modern shape types
>>> from pptx.enum.shapes import MSO_SHAPE_TYPE
>>> shape.shape_type == MSO_SHAPE_TYPE.MODEL_3D  # 3D model
>>> shape.shape_type == MSO_SHAPE_TYPE.SLIDE_ZOOM  # Slide zoom
```

## Documentation

For core python-pptx functionality, see the [python-pptx documentation](https://python-pptx.readthedocs.org/en/latest/).

For a comprehensive API reference including all extensions, see **[WOTAN/docs/python-pptx-api.md](WOTAN/docs/python-pptx-api.md)**. This includes:
- Complete API reference for all features
- Modern chart examples (Treemap, Waterfall, Map, etc.)
- Transition and animation API
- Tips for working with complex presentations

## Status

This is an experimental fork. All original python-pptx tests pass (2700+ tests). Extensions are additive and should not break existing functionality.

### Feature Coverage

| Category | Upstream | This Fork |
|----------|:--------:|:---------:|
| Basic Shapes | ✓ | ✓ |
| Traditional Charts | ✓ | ✓ |
| Tables | ✓ | ✓ |
| Text/Formatting | ✓ | ✓ |
| Images (raster) | ✓ | ✓ |
| SVG Images | - | ✓ |
| Modern Charts | - | ✓ |
| SmartArt | - | ✓ (read) |
| Morph Transitions | - | ✓ |
| 3D Models | - | ✓ (preserve) |

## License

MIT License - same as upstream python-pptx.
