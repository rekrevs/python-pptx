# Task T-TEST-01

## Header

| Field | Value |
|-------|-------|
| ID | T-TEST-01 |
| Parent | User request |
| State | DONE |
| Created | 2024-12-03 |
| Completed | 2024-12-03 |

## Objective

Research developer forums, community discussions, and known issues around PPTX format versioning. Understand how multiple OOXML versions should be handled and document findings.

## Acceptance Criteria

- [x] Search developer forums for python-pptx discussions and known issues
- [x] Research OOXML version differences (2007, 2010, 2013, 2016, 2019, 365)
- [x] Understand how version compatibility is typically handled
- [x] Document findings and recommendations for WOTAN approach
- [x] Update relevant WOTAN docs with versioning strategy

## Context

Before starting implementation, we need to understand:
1. What issues others have encountered
2. How PPTX format versions differ
3. Best practices for handling multiple versions
4. Any existing solutions or patterns

## Research Sources Used

- GitHub issues: #83 (SmartArt), #394 (SVG), #652 (SVG), #924 (Office 365 corruption)
- Stack Overflow: python-pptx tag discussions
- Library of Congress: OOXML format documentation
- Microsoft Learn: OpenXML SDK, MCE specification
- Wikipedia: Office Open XML

## Key Findings

### 1. OOXML Has Two Conformance Classes
- **Strict**: Uses `purl.oclc.org` namespaces, excludes legacy features
- **Transitional**: Uses `schemas.openxmlformats.org`, includes legacy (VML, etc.)
- **Real-world**: ~100% of files are Transitional (Office default)

### 2. Version Extensions via Namespaces
| Office | Namespaces |
|--------|-----------|
| 2007 | p:, a:, r:, c: (base) |
| 2010 | p14:, a14: |
| 2013 | p15:, a15: |
| 2016 | a16:, c16: |
| 2019+ | p188: and others |

### 3. MCE (Markup Compatibility and Extensibility)
- ISO 29500 Part 3 defines forward/backward compatibility
- `AlternateContent` blocks provide modern + fallback versions
- `Ignorable` attribute lists namespaces to skip if unsupported
- **python-pptx preserves unknown elements but doesn't parse MCE**

### 4. Known Community Issues
- **SVG**: Requires dual storage (SVG + PNG fallback), PIL doesn't support SVG
- **SmartArt**: Complex multi-part structure, different placeholder format
- **Office 365**: Some generated files need repair (issue #924)
- **No .ppt support**: Only OOXML (.pptx), not binary format
- **No PDF export**: No rendering engine

### 5. How Other Libraries Handle This
- **.NET OpenXML SDK**: Has versioned APIs (Office2007, Office2010, etc.)
- **Validation**: Separate validation for Strict vs Transitional
- **MCE processing**: SDK has MCE preprocessor for compatibility

## Recommendations for WOTAN

1. **Register all extension namespaces** in ns.py
2. **Implement MCE AlternateContent parsing**
3. **Add version detection** to Presentation class
4. **Generate fallback content** for modern features (SVG→PNG, etc.)
5. **Default to Transitional** conformance for maximum compatibility
6. **Test against multiple Office versions**

## Output

Created: `WOTAN/docs/versioning-strategy.md`

Complete documentation of:
- OOXML version history
- Strict vs Transitional differences
- Extension namespace mapping
- MCE mechanism explanation
- Recommended implementation strategy
- Testing approach
- Known issues from community

## Outcome

**State**: DONE

Research complete. The key insight is that OOXML uses **extension namespaces** (p14, a16, etc.) and **MCE AlternateContent blocks** for version compatibility. python-pptx already preserves these through round-trip but doesn't expose APIs for them. Our approach should:

1. Parse MCE blocks to access modern content
2. Support all extension namespaces
3. Generate fallbacks when creating modern features
4. Default to maximum compatibility
