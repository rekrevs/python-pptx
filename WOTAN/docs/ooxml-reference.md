# OOXML Reference for python-pptx Extensions

## Specification Documents

### ISO/IEC 29500 (OOXML Standard)
- **Part 1** (29500-1:2016): Fundamentals and Markup Language Reference
- **Part 2** (29500-2:2021): Open Packaging Conventions
- **Part 3** (29500-3:2015): Markup Compatibility and Extensibility
- **Part 4** (29500-4:2016): Transitional Migration Features

### Microsoft Extensions
- **[MS-PPTX]**: PowerPoint Extensions to OOXML
- **[MS-ODRAWXML]**: Office Drawing Extensions
- **[MS-OI29500]**: Office Implementation Information

## XML Namespaces

### Currently Supported in python-pptx

| Prefix | Namespace URI | Description |
|--------|--------------|-------------|
| `p` | `http://schemas.openxmlformats.org/presentationml/2006/main` | PresentationML |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | DrawingML |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationships |
| `c` | `http://schemas.openxmlformats.org/drawingml/2006/chart` | Charts |
| `pic` | `http://schemas.openxmlformats.org/drawingml/2006/picture` | Pictures |

### Needed for Extensions

| Prefix | Namespace URI | Description | Features |
|--------|--------------|-------------|----------|
| `p14` | `http://schemas.microsoft.com/office/powerpoint/2010/main` | PowerPoint 2010+ | Morph transitions, Zoom |
| `p15` | `http://schemas.microsoft.com/office/powerpoint/2012/main` | PowerPoint 2012+ | Additional features |
| `a16` | `http://schemas.microsoft.com/office/drawing/2014/main` | DrawingML 2014+ | SVG support |
| `c16` | `http://schemas.microsoft.com/office/drawing/2014/chart` | Charts 2014+ | Modern chart types |
| `dgm` | `http://schemas.openxmlformats.org/drawingml/2006/diagram` | Diagrams | SmartArt |
| `dsp` | `http://schemas.microsoft.com/office/drawing/2008/diagram` | Diagram shapes | SmartArt rendering |

## Key Element Structures

### SVG Images

```xml
<p:pic>
  <p:nvPicPr>...</p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId1">
      <a:extLst>
        <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
          <a16:svgBlip r:embed="rId2"/>
        </a:ext>
      </a:extLst>
    </a:blip>
  </p:blipFill>
  <p:spPr>...</p:spPr>
</p:pic>
```

Note: SVG is stored as a fallback with PNG/EMF as primary for compatibility.

### Morph Transition

```xml
<p:transition spd="slow" p14:dur="2000">
  <p14:morph option="byObject"/>
</p:transition>
```

Options: `byObject`, `byWord`, `byChar`

### SmartArt (Diagram)

SmartArt consists of multiple parts:
- `diagrams/data1.xml` - Data model (text, hierarchy)
- `diagrams/layout1.xml` - Layout definition
- `diagrams/colors1.xml` - Color scheme
- `diagrams/style1.xml` - Style definition
- `diagrams/drawing1.xml` - Rendered shapes

```xml
<!-- In slide -->
<p:graphicFrame>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
      <dgm:relIds r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

### Modern Charts (c16 namespace)

```xml
<c:chartSpace>
  <c:chart>
    <c:plotArea>
      <c16:treemap>
        <!-- Treemap specific elements -->
      </c16:treemap>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
```

### Stock Chart

```xml
<c:stockChart>
  <c:ser>
    <c:idx val="0"/>
    <c:order val="0"/>
    <c:val>
      <c:numRef>...</c:numRef>
    </c:val>
  </c:ser>
  <!-- High, Low, Close (and optionally Open, Volume) series -->
  <c:hiLowLines/>
  <c:upDownBars>
    <c:gapWidth val="150"/>
    <c:upBars/>
    <c:downBars/>
  </c:upDownBars>
</c:stockChart>
```

### Zoom (Section/Slide)

```xml
<p14:sectionZm>
  <p14:nvZmPr>
    <p:cNvPr id="4" name="Section Zoom 3"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p14:nvZmPr>
  <p14:zmPr>
    <p14:returnJump val="1"/>
  </p14:zmPr>
</p14:sectionZm>
```

## Content Types

### New Content Types Needed

| Extension | Content Type |
|-----------|-------------|
| `.svg` | `image/svg+xml` |
| `.glb` | `model/gltf-binary` |
| `.gltf` | `model/gltf+json` |

### Relationship Types

| Type | URI |
|------|-----|
| SVG Image | `http://schemas.microsoft.com/office/2007/relationships/svgImage` |
| 3D Model | `http://schemas.microsoft.com/office/2017/relationships/model3d` |
| Diagram Data | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData` |
| Diagram Layout | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout` |

## Implementation Patterns

### Adding a New Namespace

In `src/pptx/oxml/ns.py`:

```python
nsmap['p14'] = 'http://schemas.microsoft.com/office/powerpoint/2010/main'

def qn(tag):
    # Add p14: prefix handling
```

### Registering Custom Elements

In the appropriate `src/pptx/oxml/` module:

```python
from pptx.oxml import register_element_cls
from pptx.oxml.xmlchemy import BaseOxmlElement

class CT_NewElement(BaseOxmlElement):
    """Custom element class"""
    pass

register_element_cls('p14:morph', CT_NewElement)
```

### Adding New Image Type

In `src/pptx/parts/image.py`:

```python
# Add to MIME type detection
# Add to content type constants
# Update ImagePart class
```

## Testing Approach

1. Create PPTX files in PowerPoint with target feature
2. Unzip and examine XML structure
3. Write unit tests for OXML element parsing
4. Write acceptance tests for API behavior
5. Verify round-trip preservation
