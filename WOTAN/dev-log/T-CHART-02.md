# Task T-CHART-02

## Header

| Field | Value |
|-------|-------|
| ID | T-CHART-02 |
| Parent | B-CHART-02 |
| State | PARTIAL |
| Created | 2024-12-03 |
| Updated | 2024-12-03 |

## Objective

Implement Surface chart support in python-pptx, enabling reading and creation of Surface chart types (SURFACE, SURFACE_WIREFRAME, SURFACE_TOP_VIEW, SURFACE_TOP_VIEW_WIREFRAME).

## Acceptance Criteria

- [x] CT_SurfaceChart OXML element class defined
- [x] CT_Surface3DChart OXML element class defined
- [x] OXML elements registered in __init__.py
- [x] All tests pass (2700 unit, 973 acceptance)
- [ ] Surface charts detectable via PlotTypeInspector (future)
- [ ] XML writer for Surface charts (future)

## Context

### Surface Chart Types

| Enum | Value | XML Element | Description |
|------|-------|-------------|-------------|
| `SURFACE` | 83 | `c:surfaceChart` | 3D Surface (wireframe=0) |
| `SURFACE_WIREFRAME` | 84 | `c:surfaceChart` | 3D Surface (wireframe=1) |
| `SURFACE_TOP_VIEW` | 85 | `c:surface3DChart` | Surface Top View |
| `SURFACE_TOP_VIEW_WIREFRAME` | 86 | `c:surface3DChart` | Surface Top View wireframe |

### Key Characteristics

- No varyColors element
- No dLbls element (data labels not supported)
- Optional wireframe attribute
- Surface series are simpler (no dPt/markers/labels)
- 2-3 axes for surfaceChart, exactly 3 for surface3DChart

## Files Modified

- `src/pptx/oxml/chart/plot.py` - Added CT_SurfaceChart and CT_Surface3DChart classes
- `src/pptx/oxml/__init__.py` - Registered c:surfaceChart and c:surface3DChart elements

## Evidence

### Unit tests pass

```
$ pytest tests/ -q
2700 passed in 3.22s
```

### Acceptance tests pass

```
$ behave features/ -q
54 features passed, 0 failed, 0 skipped
973 scenarios passed, 0 failed, 0 skipped
2914 steps passed, 0 failed, 0 skipped
```

### OXML elements defined

```python
class CT_SurfaceChart(BaseChartElement):
    """Surface charts with 3D rendering."""
    ser = ZeroOrMore("c:ser", ...)

class CT_Surface3DChart(BaseChartElement):
    """Surface charts with top-view (3D) rendering."""
    ser = ZeroOrMore("c:ser", ...)
```

## Remaining Work

1. **PlotTypeInspector**: Add detection logic to recognize surface chart elements
2. **XML writer**: Implement _SurfaceChartXmlWriter
3. **Plot classes**: Add SurfacePlot with wireframe property

## Outcome

**State**: PARTIAL

OXML infrastructure for Surface charts is in place:
- CT_SurfaceChart and CT_Surface3DChart elements defined
- Elements registered in __init__.py
- Round-trip of existing surface charts will work

Remaining:
- XML writer for creating new surface charts
- Plot type detection
