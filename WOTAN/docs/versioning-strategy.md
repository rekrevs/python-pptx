# OOXML Versioning Strategy

## The Version Problem

PPTX files can come from different versions of PowerPoint (2007, 2010, 2013, 2016, 2019, 365), each potentially using different OOXML features and namespaces. Our goal of "total PPTX mastery" requires handling all of these.

## OOXML Version History

### Standards Evolution

| Version | Standard | Key Changes |
|---------|----------|-------------|
| ECMA-376 1st Ed (2006) | Original | Base spec, Office 2007 |
| ISO/IEC 29500:2008 | ISO adoption | Strict/Transitional split |
| ECMA-376 2nd Ed (2008) | Aligned with ISO | Minor fixes |
| ISO/IEC 29500:2011 | Update | Bug fixes |
| ECMA-376 4th Ed (2012) | Update | Office 2013 features |
| ISO/IEC 29500:2016 | Current | Office 2016 features |

### Strict vs Transitional

Two conformance classes exist:

| Aspect | Strict | Transitional |
|--------|--------|--------------|
| Namespace | `purl.oclc.org/ooxml/...` | `schemas.openxmlformats.org/...` |
| Legacy features | Excluded | Included (VML, etc.) |
| Office default | Never | Always |
| Real-world usage | Rare | ~100% |

**Key insight**: Virtually all real-world PPTX files use Transitional. Office defaults to Transitional for backward compatibility.

### Extension Namespaces by Office Version

| Office Version | Namespaces Added |
|----------------|------------------|
| 2007 | `p:`, `a:`, `r:`, `c:` (base) |
| 2010 | `p14:`, `a14:` |
| 2012 | `p15:`, `a15:` |
| 2014/2016 | `a16:`, `c16:` |
| 2019+ | `p188:`, etc. |

The `p14`, `p15`, `a16` etc. are **extension namespaces** defined via the MCE (Markup Compatibility and Extensibility) mechanism.

## Markup Compatibility and Extensibility (MCE)

MCE (ISO 29500 Part 3) is the key to forward/backward compatibility.

### How MCE Works

```xml
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice Requires="a16">
    <!-- Modern SVG element -->
    <a16:svgBlip r:embed="rId2"/>
  </mc:Choice>
  <mc:Fallback>
    <!-- Fallback PNG for older readers -->
    <a:blip r:embed="rId1"/>
  </mc:Fallback>
</mc:AlternateContent>
```

### MCE Attributes

| Attribute | Purpose |
|-----------|---------|
| `Ignorable` | Namespaces to silently ignore if not understood |
| `MustUnderstand` | Namespaces that must be understood |
| `PreserveElements` | Unknown elements to preserve on save |
| `PreserveAttributes` | Unknown attributes to preserve on save |
| `ProcessContent` | Process children even if parent unknown |

### Current python-pptx MCE Handling

python-pptx currently:
- **Preserves** unknown elements through round-trip (good!)
- **Does not parse** MCE `AlternateContent` blocks
- **Ignores** extension namespace content without API exposure

## Version-Specific Features

### PowerPoint 2010 (p14:)
- Enhanced transitions
- Section support
- Video/audio improvements
- Creation/modification IDs

### PowerPoint 2013 (p15:)
- Guide enhancements
- Chart tracking
- Comments improvements

### PowerPoint 2016/365 (a16:, c16:)
- SVG images (`a16:svgBlip`)
- Modern charts (Treemap, Sunburst, etc.)
- Creation IDs for collaboration

### PowerPoint 2019+
- Morph transitions
- Zoom
- 3D models
- Ink improvements

## Recommended WOTAN Strategy

### 1. Namespace-Aware Parsing

Register all known namespaces:

```python
# In pptx/oxml/ns.py
nsmap = {
    # Base (Office 2007)
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',

    # Extensions
    'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
    'p15': 'http://schemas.microsoft.com/office/powerpoint/2012/main',
    'a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
    'a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
    'c16': 'http://schemas.microsoft.com/office/drawing/2014/chart',

    # MCE
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}
```

### 2. MCE Processing

Implement MCE-aware element access:

```python
def get_effective_element(elem):
    """Get the best available content from AlternateContent."""
    if elem.tag == MC_ALTERNATE_CONTENT:
        for choice in elem.findall('mc:Choice'):
            requires = choice.get('Requires')
            if namespace_supported(requires):
                return choice[0]  # Return choice content
        fallback = elem.find('mc:Fallback')
        if fallback is not None:
            return fallback[0]
    return elem
```

### 3. Version Detection

Add presentation version detection:

```python
class Presentation:
    @property
    def ooxml_version(self):
        """Detect OOXML version based on namespaces used."""
        namespaces = self._collect_namespaces()
        if 'a16' in namespaces:
            return 'Office 2016+'
        if 'p15' in namespaces:
            return 'Office 2013+'
        if 'p14' in namespaces:
            return 'Office 2010+'
        return 'Office 2007'
```

### 4. Feature Availability

Track feature availability by version:

```python
FEATURES = {
    'svg_images': {'min_version': 'Office 2016+', 'namespace': 'a16'},
    'morph_transition': {'min_version': 'Office 2019', 'namespace': 'p14'},
    'modern_charts': {'min_version': 'Office 2016+', 'namespace': 'c16'},
}
```

### 5. Write Strategy

When writing, support target version:

```python
def save(self, path, target_version='Office 2007'):
    """Save with appropriate version features.

    For maximum compatibility, use Office 2007 (Transitional).
    For modern features, use Office 2016+ or Office 2019.
    """
    if target_version == 'Office 2007':
        # Use only base namespaces
        # Generate fallback content for modern features
    else:
        # Use extension namespaces as needed
```

### 6. Graceful Degradation

For features requiring fallback:

```python
def add_svg_image(self, svg_path, png_fallback=None):
    """Add SVG with PNG fallback for older PowerPoint versions."""
    if png_fallback is None:
        png_fallback = convert_svg_to_png(svg_path)

    # Create AlternateContent with both versions
    # Modern readers get SVG, older readers get PNG
```

## Testing Strategy

### Test Matrix

| Test | Office 2007 | 2010 | 2013 | 2016 | 2019 | 365 |
|------|-------------|------|------|------|------|-----|
| Round-trip base features | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Read extension namespaces | - | ✓ | ✓ | ✓ | ✓ | ✓ |
| Write with fallbacks | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Modern feature creation | - | - | - | ✓ | ✓ | ✓ |

### Test Files Needed

1. Minimal files from each Office version
2. Files with AlternateContent blocks
3. Files with each extension namespace
4. Strict conformance files (rare but should handle)

## Known Issues from Community

### From GitHub Issues

1. **SVG Images** ([#394](https://github.com/scanny/python-pptx/issues/394), [#652](https://github.com/scanny/python-pptx/issues/652))
   - SVG stored as dual format (SVG + PNG fallback)
   - Need both for compatibility
   - PIL doesn't support SVG

2. **SmartArt** ([#83](https://github.com/scanny/python-pptx/issues/83))
   - Complex multi-part structure
   - Different placeholder representation
   - Related parts via relationships

3. **Corruption with Office 365** ([#924](https://github.com/scanny/python-pptx/issues/924))
   - Some generated files need repair in Office 365 desktop
   - Works in web version and other apps

### From Stack Overflow

1. **No .ppt support** - Only .pptx (OOXML), not binary .ppt
2. **No PDF export** - Cannot render, only manipulate structure
3. **Auto-fit limitations** - No rendering engine for text fitting
4. **LibreOffice differences** - Some features render differently

## Recommendations Summary

1. **Register all extension namespaces** in `ns.py`
2. **Implement MCE processing** for AlternateContent blocks
3. **Add version detection** to Presentation class
4. **Create fallback content** for modern features
5. **Document version requirements** for each feature
6. **Test against multiple Office versions**
7. **Default to maximum compatibility** (Transitional, with fallbacks)

## Sources

- [OOXML Format Family - Library of Congress](https://www.loc.gov/preservation/digital/formats/fdd/fdd000395.shtml)
- [PPTX Transitional Format](https://www.loc.gov/preservation/digital/formats/fdd/fdd000399.shtml)
- [MCE Specification](https://www.loc.gov/preservation/digital/formats/fdd/fdd000396.shtml)
- [OpenXML SDK Migration Guide](https://learn.microsoft.com/en-us/office/open-xml/migration/migrate-v2-to-v3)
- [MS-ODRAWXML Extensions](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/06cff208-c6e1-4db7-bb68-665135e5f0de)
- [python-pptx GitHub Issues](https://github.com/scanny/python-pptx/issues)
