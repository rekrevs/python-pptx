# python-pptx (xtend fork)

> Python library for creating/reading/updating PowerPoint (.pptx) files.

## Project Goals

The `xtend` branch extends python-pptx to support modern PPTX features:
- SVG image support (done)
- Modern chart types: Treemap, Sunburst, Waterfall, Funnel, Map (partial)
- SmartArt read support (done)
- Stock and Surface charts (partial)
- Morph transitions (planned)
- Other PowerPoint 2016+ features

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
- Quality: pyright (strict), ruff

## Architecture

```yaml
layers:  # top to bottom
  - api.py: Presentation() entry point
  - presentation.py: Presentation, Slides, SlideMasters
  - shapes/: Shape, Picture, GraphicFrame, Table, Chart
  - chart/: Chart, Series, Plot, Categories
  - dml/: DrawingML formatting (color, fill, line)
  - oxml/: CT_* element classes (XML ↔ Python mapping)
  - opc/: ZIP package, Parts, Relationships

src/pptx/:
  api.py: public entry point
  presentation.py: Presentation class
  slide.py: Slide, SlideLayout, SlideMaster
  shapes/: Shape hierarchy (autoshape, picture, group, etc.)
  chart/: Chart types, series, data labels
  table.py: Table, Row, Cell
  dml/: Color, Fill, Line formatting
  oxml/: CT_* element classes
  opc/: Open Packaging Conventions

namespaces:  # in oxml/ns.py
  core: a, r, p, c (DrawingML, Relationships, PresentationML, Charts)
  extended: asvg, dgm, dsp, cx (SVG, Diagrams, ChartEx)

tests/: unit tests (pytest, it_*/test_* naming)
features/: BDD acceptance tests (behave/gherkin)
wotan/:
  backlog.json: task index
  dev-log/: task files (T-NNNN.md)
  docs/: specifications
```

## Commands

```bash
# Testing
pytest tests/ -q                    # Unit tests
behave features/                    # Acceptance tests
pytest tests/ -q && behave          # Full suite

# Quality
pyright                             # Type checking
ruff check                          # Linting

# Task management
/wotan                              # Show active tasks
/wotan add "description"            # Create task
/wotan start                        # Execute next task
```

## Testing

- Unit tests in `tests/unit/` mirror `src/pptx/` structure
- Acceptance tests in `features/` (Gherkin + step definitions)
- Fixtures in `tests/unit/unitdata/` and XML snippets
- Naming: `it_does_something`, `test_feature_behavior`

## Adding OXML Elements

When extending the XML layer:
1. Define element classes in `src/pptx/oxml/` using xmlchemy pattern
2. Register custom element classes with lxml
3. Add namespace prefixes to `src/pptx/oxml/ns.py` if needed
4. Follow `CT_*` naming convention

Key extension namespaces:
- `p14:` - PowerPoint 2010+ (transitions)
- `a16:` - DrawingML 2016
- `dgm:` - Diagrams (SmartArt)
- `cx:` - ChartEx (modern charts)

## Extension Features

This fork adds support for:
- SVG images (with PNG fallback)
- SmartArt (read support)
- Modern charts: Stock, Surface, Treemap, Sunburst (via ChartEx)
- Bezier freeform shapes
- Media type detection

See `wotan/docs/` for specifications.

## Verification

Before completing work:
```bash
pytest tests/ -q && behave features/ && pyright && ruff check
```
