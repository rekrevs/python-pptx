# Baseline Test Report

**Date**: 2024-12-03
**Branch**: xtend
**python-pptx Version**: 1.0.2
**Python Version**: 3.12.9

## Executive Summary

All existing tests pass. Example documents round-trip successfully with no data loss. python-pptx v1.0.2 is stable and functioning as documented.

---

## Test Suite Results

### Unit Tests (pytest)

```
pytest tests/ --tb=no -q
```

| Metric | Result |
|--------|--------|
| Tests Run | 2,700 |
| Passed | 2,700 |
| Failed | 0 |
| Skipped | 0 |
| Duration | 3.01s |

**Status**: ✅ ALL PASS

### Acceptance Tests (behave)

```
behave
```

| Metric | Result |
|--------|--------|
| Features | 54 passed, 0 failed |
| Scenarios | 973 passed, 0 failed |
| Steps | 2,914 passed, 0 failed |
| Duration | 1.67s |

**Status**: ✅ ALL PASS

---

## Test File Inventory

### PPTX Test Files (`tests/test_files/`)

| File | Size | Purpose |
|------|------|---------|
| `minimal.pptx` | 15,802 B | Minimal valid presentation |
| `test.pptx` | 37,859 B | General test fixture |
| `test_slides.pptx` | 38,616 B | Slide-related tests |
| `no-slides.pptx` | 34,036 B | Edge case: no slides |
| `no-core-props.pptx` | 34,436 B | Edge case: missing core properties |
| `missing_rels_item.pptx` | 32,944 B | Edge case: missing relationship items |

### Supporting Test Files

| File | Purpose |
|------|---------|
| `calibriz.ttf` | Font file for text rendering tests |
| `python-icon.jpeg` | JPEG image test |
| `python-powered.png` | PNG image test |
| `python.bmp` | BMP image test |
| `monty-truth.png` | PNG image test |
| `dummy.mp4` | Video media test |
| `cdw-logo.eps` | EPS vector (not directly supported) |

### XML Test Snippets (`tests/test_files/snippets/`)

48 XML snippet files covering:
- Chart types: area, bar, column, line, pie, doughnut, bubble, scatter, radar
- Chart configurations: stacked, clustered, 3D, date axis, float axis
- Multi-category charts
- Various XML structure tests

### Expanded PPTX Structure (`tests/test_files/expanded_pptx/`)

Pre-extracted PPTX structure for XML-level testing:
- `docProps/` - Document properties
- `ppt/slides/` - Slide XML
- `ppt/slideLayouts/` - Layout XML
- `ppt/slideMasters/` - Master XML
- `_rels/` - Relationships

---

## Example Document Analysis

### Documents Tested

| File | Slides | Shapes | Size |
|------|--------|--------|------|
| `2023-03-17-mogren-ai-center-steering-group-deep-learning(1).pptx` | 10 | 108 | 18.8 MB |
| `AI at RISE 2024-05-21.pptx` | 22 | 149 | 37.9 MB |
| `ErdzanHodzic.pptx` | 1 | 4 | 747 KB |

### Round-Trip Test Results

| File | Status | Size Change |
|------|--------|-------------|
| `2023-03-17-mogren...` | ✅ PASS | -1,446 B (-0.0%) |
| `AI at RISE 2024-05-21.pptx` | ✅ PASS | -31,964 B (-0.1%) |
| `ErdzanHodzic.pptx` | ✅ PASS | -5,519 B (-0.7%) |

**Note**: Small size reductions are due to XML formatting/whitespace normalization. No content loss detected.

### Features Detected in Example Documents

#### Shape Types Used

| Shape Type | Doc 1 | Doc 2 | Doc 3 |
|------------|-------|-------|-------|
| PLACEHOLDER | 23 | 52 | 2 |
| PICTURE | 44 | 19 | 1 |
| TEXT_BOX | 30 | 16 | 1 |
| AUTO_SHAPE | 7 | 41 | - |
| LINE | 4 | 1 | - |
| FREEFORM | - | 14 | - |
| GROUP | - | 6 | - |

#### Placeholder Types Used

- TITLE, BODY, SUBTITLE, CENTER_TITLE
- FOOTER, SLIDE_NUMBER
- OBJECT, PICTURE

#### Modern Namespaces Present (preserved through round-trip)

| Namespace | Doc 1 | Doc 2 | Doc 3 | Description |
|-----------|-------|-------|-------|-------------|
| p14 | ✓ | ✓ | ✓ | PowerPoint 2010+ |
| p15 | ✓ | ✓ | ✓ | PowerPoint 2012+ |
| a16 | - | ✓ | ✓ | DrawingML 2014+ |
| dgm | ✓ | - | - | SmartArt/Diagrams |
| Transitions | ✓ | ✓ | - | Slide transitions |

**Status**: ✅ All namespaces preserved through round-trip

---

## Feature Coverage Analysis

### Fully Functional (Tested & Working)

| Feature | Unit Tests | Acceptance Tests | Example Docs |
|---------|------------|------------------|--------------|
| Open/Save PPTX | ✅ | ✅ | ✅ |
| Slides (add/remove/access) | ✅ | ✅ | ✅ |
| Slide layouts | ✅ | ✅ | ✅ |
| Slide masters | ✅ | ✅ | - |
| Placeholders | ✅ | ✅ | ✅ |
| AutoShapes (190+ types) | ✅ | ✅ | ✅ |
| Text frames | ✅ | ✅ | ✅ |
| Paragraphs & runs | ✅ | ✅ | ✅ |
| Font formatting | ✅ | ✅ | ✅ |
| Pictures (PNG, JPEG, BMP, etc.) | ✅ | ✅ | ✅ |
| Tables | ✅ | ✅ | - |
| Charts (basic types) | ✅ | ✅ | - |
| Fill (solid, gradient, pattern) | ✅ | ✅ | - |
| Line formatting | ✅ | ✅ | ✅ |
| Hyperlinks | ✅ | ✅ | - |
| Group shapes | ✅ | - | ✅ |
| Freeform shapes | ✅ | - | ✅ |
| Core properties | ✅ | ✅ | - |
| Slide notes | ✅ | ✅ | - |

### Preserved but Not Editable

These features are preserved through round-trip but cannot be programmatically modified:

| Feature | Preservation | Notes |
|---------|--------------|-------|
| Transitions | ✅ | Present in example docs, preserved |
| SmartArt | ⚠️ | Detected as UNKNOWN shape type |
| p14/p15/a16 extensions | ✅ | Namespaces preserved |

### Not Supported

| Feature | Status |
|---------|--------|
| SVG images | Skipped during read |
| Modern charts (Treemap, Sunburst, etc.) | Not implemented |
| Morph transitions | Not implemented |
| 3D models | Not implemented |
| Ink annotations | Not implemented |

---

## Regressions Identified

**None identified.** All documented functionality works as expected.

---

## Recommendations

1. **Before any changes**: This baseline confirms the library is stable. Any modifications should not break these 2,700+ tests.

2. **SmartArt handling**: Example doc 1 contains SmartArt (dgm namespace). It's preserved through round-trip but detected as "unknown" shape type. This is expected behavior per current documentation.

3. **Test coverage gaps**: No example documents contain:
   - Charts
   - Tables
   - Video/audio media
   - Consider adding test files with these features.

4. **Modern features**: All three example documents use modern namespaces (p14, p15, a16). These are preserved, indicating good forward compatibility even without explicit support.

---

## Test Commands Reference

```bash
# Run all unit tests
pytest tests/ -v

# Run all acceptance tests
behave

# Run specific test module
pytest tests/chart/test_axis.py -v

# Run tests with coverage
pytest tests/ --cov=pptx --cov-report=html
```

---

## Conclusion

python-pptx v1.0.2 on the `xtend` branch is **fully functional** with:
- 100% test pass rate (2,700 unit + 973 acceptance scenarios)
- Successful round-trip of real-world presentations
- Preservation of modern XML namespaces
- No regressions from documented behavior

The library is ready for extension work to add modern PPTX features.
