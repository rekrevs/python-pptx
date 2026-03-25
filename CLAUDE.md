# python-pptx (xtend fork)

> Python library for creating/reading/updating PowerPoint (.pptx) files.

## Project Goals

The `xtend` branch extends python-pptx to support modern PPTX features:
- SVG image support (done)
- Modern chart types: Treemap, Sunburst, Waterfall, Funnel, Box & Whisker, Map (done)
- SmartArt read support (done)
- Stock charts: HLC, OHLC, VHLC, VOHLC (done)
- Surface charts (done)
- Morph transitions (done)
- Theme access: color scheme, font scheme (done)
- Shape effects: shadow, glow, reflection, soft edge (done)
- Bullet and numbering formatting (done)
- Custom document properties (done)
- Slide comments (done)
- Alt text / accessibility (done)
- Color transparency/alpha (done)

## Quick Start

```bash
python -m venv venv && source venv/bin/activate
pip install -e ".[dev]"
pytest tests/ -q
behave features/
```

## Tech Stack

- Python 3.8+, lxml, Pillow, XlsxWriter
- Testing: pytest, behave (BDD)
- Quality: pyright, ruff (advisory only - see below)

## Architecture

```yaml
layers:  # top to bottom
  - api.py: Presentation() entry point
  - presentation.py: Presentation, Slides, SlideMasters
  - shapes/: Shape, Picture, GraphicFrame, Table, Chart
  - chart/: Chart, Series, Plot, Categories
  - chartex/: ChartEx (modern chart types)
  - dml/: DrawingML formatting (color, fill, line, effect)
  - theme.py: Theme, ColorScheme, FontScheme
  - text/: TextFrame, Paragraph, Run, BulletFormat
  - comment.py: Slide comments
  - oxml/: CT_* element classes (XML <-> Python mapping)
  - opc/: ZIP package, Parts, Relationships

src/pptx/:
  api.py: public entry point
  presentation.py: Presentation class
  slide.py: Slide, SlideLayout, SlideMaster
  shapes/: Shape hierarchy (autoshape, picture, group, etc.)
  chart/: Chart types, series, data labels, XML writers
  chartex/: ChartEx data, XML writers (Treemap, Sunburst, etc.)
  table.py: Table, Row, Cell
  dml/: Color, Fill, Line, Effect formatting
  theme.py: Theme, ColorScheme, FontScheme
  text/: TextFrame, Paragraph, Run, BulletFormat
  comment.py: Comment proxy
  oxml/: CT_* element classes
  opc/: Open Packaging Conventions

namespaces:  # in oxml/ns.py
  core: a, r, p, c (DrawingML, Relationships, PresentationML, Charts)
  extended: asvg, dgm, dsp, cx, cust, vt (SVG, Diagrams, ChartEx, Custom Properties)

tests/: unit tests (pytest, it_*/test_* naming)
features/: BDD acceptance tests (behave/gherkin)
wotan/:
  backlog.json: task index
  dev-log/: task files (T-NNNN.md)
  docs/: specifications
```

## Commands

# Testing (required - must pass)
pytest tests/ -q                    # Unit tests
behave features/                    # Acceptance tests
pytest tests/ -q && behave          # Full suite

# Quality (advisory - not enforced in CI)
pyright                             # Type checking (~3900 pre-existing errors)
ruff check                          # Linting (~117 pre-existing warnings)

## Testing

- Unit tests in `tests/` mirror `src/pptx/` structure
- Acceptance tests in `features/` (Gherkin + step definitions)
- Fixtures in `tests/unit/unitdata/` and XML snippets
- Naming: `it_does_something`, `test_feature_behavior`

## Adding OXML Elements

When extending the XML layer:
1. Define element classes in `src/pptx/oxml/` using xmlchemy pattern
2. Register custom element classes with lxml in `src/pptx/oxml/__init__.py`
3. Add namespace prefixes to `src/pptx/oxml/ns.py` if needed
4. Follow `CT_*` naming convention

Key extension namespaces:
- `p14:` - PowerPoint 2010+ (transitions)
- `a16:` - DrawingML 2016
- `dgm:` - Diagrams (SmartArt)
- `cx:` - ChartEx (modern charts)
- `cust:` - Custom properties
- `vt:` - Document property value types

## Extension Features

This fork adds support for:
- SVG images (with PNG fallback)
- SmartArt (read support)
- All chart types: Stock (HLC/OHLC/VHLC/VOHLC), Surface, Treemap, Sunburst, Waterfall, Funnel, Box & Whisker, Region Map
- Morph transitions
- Theme access (color scheme read/write, font scheme read/write)
- Shape effects (shadow properties, glow, reflection, soft edge)
- Bullet and numbering formatting
- Custom document properties
- Slide comments (read/write)
- Alt text on all shape types
- Color transparency/alpha
- Bezier freeform shapes
- Media type detection

See `wotan/docs/` for specifications.

## Verification

Before completing work:
```bash
pytest tests/ -q && behave features/
```

Note: `pyright` and `ruff check` are configured in `pyproject.toml` but have
many pre-existing errors from the upstream codebase. They are not run in CI
(`.github/workflows/ci.yml` only runs pytest + behave). Don't introduce new
errors, but fixing existing ones is out of scope unless specifically requested.
