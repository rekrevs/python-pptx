# Vision: python-pptx Complete

## Goal

**Total PPTX Mastery**: Read any PPTX file, parse it completely into every component and detail, understand and manipulate these components in any way desired, and write back a new document with full fidelity.

## Core Principle

Every element in a PPTX file should be:
1. **Parseable** - Read into a Python object model
2. **Inspectable** - All properties accessible and understandable
3. **Manipulable** - Can be created, modified, moved, copied, deleted
4. **Serializable** - Can be written back to valid PPTX

No "black boxes" - if it's in the file, we can work with it.

## Target State

A Python library that can:

1. **Read** any valid PPTX file without errors or data loss
2. **Parse** every element into typed Python objects with full property access
3. **Understand** all relationships, inheritance, and dependencies between elements
4. **Manipulate** any aspect: create, modify, copy, move, delete any element
5. **Write** files that are fully compatible with modern PowerPoint versions

## What "Complete" Means

### Presentation Level
- [x] Core properties (title, author, etc.)
- [x] **Custom properties** ✅ xtend
- [x] Slide dimensions and settings
- [x] **Theme access and modification** ✅ xtend
- [x] **Color schemes** ✅ xtend
- [x] **Font schemes** ✅ xtend

### Slide Level
- [x] All slide types (regular, layout, master, notes, handout)
- [x] Slide relationships and inheritance
- [x] Background (all fill types)
- [x] **Transitions (including Morph)** ✅ xtend
- [ ] Animations and timing
- [x] **Slide comments** ✅ xtend

### Shape Level
- [x] All shape types (auto, text, picture, chart, table, group, connector, media)
- [x] **SmartArt (read support)** ✅ xtend
- [x] **3D models (detection/preserve)** ✅ xtend
- [x] Shape properties (position, size, rotation, flip)
- [x] Shape geometry (preset and custom)
- [x] **Shape effects (shadow, reflection, glow, soft edge)** ✅ xtend
- [x] Shape styles
- [x] **Alt text / accessibility** ✅ xtend

### Text Level
- [x] Text frames and body properties
- [x] Paragraphs (all properties)
- [x] Runs (all character properties)
- [x] **Bullet and numbering (full formatting)** ✅ xtend
- [ ] Tabs and indentation (XML only)
- [ ] Text effects
- [x] Hyperlinks

### Drawing Level
- [x] All fill types (solid, gradient, pattern, picture, group)
- [x] All line types (solid, gradient, pattern)
- [x] Line endings (arrows, etc.)
- [x] Color (RGB, theme, HSL, system, with all transformations)
- [x] **Color transparency/alpha** ✅ xtend

### Rich Content
- [x] **Charts (all types including ChartEx and volume stock)** ✅ xtend
- [x] Tables (all properties, merged cells)
- [x] **SmartArt (read structure and text)** ✅ xtend
- [x] Media (video, audio)
- [x] **Images (all formats including SVG)** ✅ xtend
- [ ] OLE objects
- [ ] Equations

## Scope

### In Scope

- All OOXML PresentationML elements (ISO/IEC 29500)
- PowerPoint 2016, 2019, 2021, and Microsoft 365 features
- Microsoft extensions (MS-PPTX specification)
- Full round-trip fidelity

### Out of Scope

- PowerPoint application automation (COM/AppleScript)
- Real-time collaboration features (cloud-dependent)
- DRM/encryption handling
- Macro execution (VBA)

## Guiding Principles

1. **Total Access** - Every PPTX element accessible via Python
2. **No Data Loss** - Round-trip preserves everything
3. **Backwards Compatible** - Existing python-pptx code continues to work
4. **Layered Architecture** - Maintain separation: API → Parts → OXML
5. **Test-Driven** - All features must have unit and acceptance tests
6. **Spec-Aligned** - Follow OOXML naming conventions and semantics
7. **Pythonic API** - High-level API feels natural to Python developers
