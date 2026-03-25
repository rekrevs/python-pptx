"""lxml custom element classes for DrawingML effect-related XML elements."""

from __future__ import annotations

from pptx.oxml.simpletypes import (
    ST_PositiveCoordinate,
    ST_PositiveFixedAngle,
    ST_PositiveFixedPercentage,
    XsdBoolean,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
    ZeroOrOneChoice,
)


class CT_EffectList(BaseOxmlElement):
    """`a:effectLst` custom element class.

    Contains optional effect child elements in schema-defined sequence.
    """

    _tag_seq = (
        "a:blur",
        "a:fillOverlay",
        "a:glow",
        "a:innerShdw",
        "a:outerShdw",
        "a:prstShdw",
        "a:reflection",
        "a:softEdge",
    )
    glow = ZeroOrOne("a:glow", successors=_tag_seq[3:])
    innerShdw = ZeroOrOne("a:innerShdw", successors=_tag_seq[4:])
    outerShdw = ZeroOrOne("a:outerShdw", successors=_tag_seq[5:])
    reflection = ZeroOrOne("a:reflection", successors=_tag_seq[6:])
    softEdge = ZeroOrOne("a:softEdge", successors=_tag_seq[7:])
    del _tag_seq


class CT_OuterShadowEffect(BaseOxmlElement):
    """`a:outerShdw` custom element class.

    Contains exactly one color child (EG_ColorChoice group) and attributes
    controlling shadow appearance.
    """

    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
    blurRad = OptionalAttribute("blurRad", ST_PositiveCoordinate)
    dist = OptionalAttribute("dist", ST_PositiveCoordinate)
    dir = OptionalAttribute("dir", ST_PositiveFixedAngle)
    rotWithShape = OptionalAttribute("rotWithShape", XsdBoolean)


class CT_InnerShadowEffect(BaseOxmlElement):
    """`a:innerShdw` custom element class.

    Contains exactly one color child (EG_ColorChoice group) and attributes
    controlling inner shadow appearance.
    """

    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
    blurRad = OptionalAttribute("blurRad", ST_PositiveCoordinate)
    dist = OptionalAttribute("dist", ST_PositiveCoordinate)
    dir = OptionalAttribute("dir", ST_PositiveFixedAngle)


class CT_GlowEffect(BaseOxmlElement):
    """`a:glow` custom element class.

    Contains exactly one color child (EG_ColorChoice group) and a radius
    attribute controlling the glow extent.
    """

    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
    rad = OptionalAttribute("rad", ST_PositiveCoordinate)


class CT_ReflectionEffect(BaseOxmlElement):
    """`a:reflection` custom element class.

    Attributes-only element (no color children). Controls reflection
    appearance via blur, distance, direction, and opacity attributes.
    """

    blurRad = OptionalAttribute("blurRad", ST_PositiveCoordinate)
    stA = OptionalAttribute("stA", ST_PositiveFixedPercentage)
    endA = OptionalAttribute("endA", ST_PositiveFixedPercentage)
    dist = OptionalAttribute("dist", ST_PositiveCoordinate)
    dir = OptionalAttribute("dir", ST_PositiveFixedAngle)


class CT_SoftEdgesEffect(BaseOxmlElement):
    """`a:softEdge` custom element class.

    Attributes-only element with a single required radius attribute.
    """

    rad = RequiredAttribute("rad", ST_PositiveCoordinate)
