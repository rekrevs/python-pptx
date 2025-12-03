"""Custom element classes for slide transitions."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.simpletypes import XsdBoolean, XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrOne,
)

if TYPE_CHECKING:
    pass


class CT_SlideTransition(BaseOxmlElement):
    """`p:transition` element specifying slide transition settings.

    The transition element specifies the kind of slide transition that should be
    used to transition to the current slide from the previous slide. The transition
    information is stored on the slide that appears after the transition.
    """

    # Child elements for transition types (subset - most common)
    _tag_seq = (
        "p:blinds",
        "p:checker",
        "p:circle",
        "p:dissolve",
        "p:comb",
        "p:cover",
        "p:cut",
        "p:diamond",
        "p:fade",
        "p:newsflash",
        "p:plus",
        "p:pull",
        "p:push",
        "p:random",
        "p:randomBar",
        "p:split",
        "p:strips",
        "p:wedge",
        "p:wheel",
        "p:wipe",
        "p:zoom",
        "p14:prism",
        "p14:vortex",
        "p14:ripple",
        "p14:honeycomb",
        "p14:flash",
        "p14:glitter",
        "p14:doors",
        "p14:window",
        "p14:ferris",
        "p14:gallery",
        "p14:conveyor",
        "p14:pan",
        "p14:reveal",
        "p14:wheelReverse",
        "p14:flythrough",
        "p14:warp",
        "p14:shred",
        "p14:switch",
        "p14:flip",
        "p14:cube",
        "p14:box",
        "p14:rotate",
        "p14:orbit",
        "p159:morph",
        "p:sndAc",
        "p:extLst",
    )

    # Standard transition types
    blinds = ZeroOrOne("p:blinds", successors=_tag_seq[1:])
    checker = ZeroOrOne("p:checker", successors=_tag_seq[2:])
    circle = ZeroOrOne("p:circle", successors=_tag_seq[3:])
    dissolve = ZeroOrOne("p:dissolve", successors=_tag_seq[4:])
    comb = ZeroOrOne("p:comb", successors=_tag_seq[5:])
    cover = ZeroOrOne("p:cover", successors=_tag_seq[6:])
    cut = ZeroOrOne("p:cut", successors=_tag_seq[7:])
    diamond = ZeroOrOne("p:diamond", successors=_tag_seq[8:])
    fade = ZeroOrOne("p:fade", successors=_tag_seq[9:])
    newsflash = ZeroOrOne("p:newsflash", successors=_tag_seq[10:])
    plus = ZeroOrOne("p:plus", successors=_tag_seq[11:])
    pull = ZeroOrOne("p:pull", successors=_tag_seq[12:])
    push = ZeroOrOne("p:push", successors=_tag_seq[13:])
    random = ZeroOrOne("p:random", successors=_tag_seq[14:])
    randomBar = ZeroOrOne("p:randomBar", successors=_tag_seq[15:])
    split = ZeroOrOne("p:split", successors=_tag_seq[16:])
    strips = ZeroOrOne("p:strips", successors=_tag_seq[17:])
    wedge = ZeroOrOne("p:wedge", successors=_tag_seq[18:])
    wheel = ZeroOrOne("p:wheel", successors=_tag_seq[19:])
    wipe = ZeroOrOne("p:wipe", successors=_tag_seq[20:])
    zoom = ZeroOrOne("p:zoom", successors=_tag_seq[21:])

    # PowerPoint 2010+ transitions (p14 namespace)
    prism = ZeroOrOne("p14:prism", successors=_tag_seq[22:])
    vortex = ZeroOrOne("p14:vortex", successors=_tag_seq[23:])
    ripple = ZeroOrOne("p14:ripple", successors=_tag_seq[24:])
    honeycomb = ZeroOrOne("p14:honeycomb", successors=_tag_seq[25:])
    flash = ZeroOrOne("p14:flash", successors=_tag_seq[26:])
    glitter = ZeroOrOne("p14:glitter", successors=_tag_seq[27:])
    doors = ZeroOrOne("p14:doors", successors=_tag_seq[28:])
    window = ZeroOrOne("p14:window", successors=_tag_seq[29:])
    ferris = ZeroOrOne("p14:ferris", successors=_tag_seq[30:])
    gallery = ZeroOrOne("p14:gallery", successors=_tag_seq[31:])
    conveyor = ZeroOrOne("p14:conveyor", successors=_tag_seq[32:])
    pan = ZeroOrOne("p14:pan", successors=_tag_seq[33:])
    reveal = ZeroOrOne("p14:reveal", successors=_tag_seq[34:])
    wheelReverse = ZeroOrOne("p14:wheelReverse", successors=_tag_seq[35:])
    flythrough = ZeroOrOne("p14:flythrough", successors=_tag_seq[36:])
    warp = ZeroOrOne("p14:warp", successors=_tag_seq[37:])
    shred = ZeroOrOne("p14:shred", successors=_tag_seq[38:])
    switch = ZeroOrOne("p14:switch", successors=_tag_seq[39:])
    flip = ZeroOrOne("p14:flip", successors=_tag_seq[40:])
    cube = ZeroOrOne("p14:cube", successors=_tag_seq[41:])
    box = ZeroOrOne("p14:box", successors=_tag_seq[42:])
    rotate = ZeroOrOne("p14:rotate", successors=_tag_seq[43:])
    orbit = ZeroOrOne("p14:orbit", successors=_tag_seq[44:])

    # Morph transition (p159 namespace - PowerPoint 2016+)
    morph: CT_MorphTransition | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p159:morph", successors=_tag_seq[45:]
    )

    del _tag_seq

    # Attributes
    spd: str | None = OptionalAttribute("spd", XsdString)  # pyright: ignore[reportAssignmentType]
    advClick: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "advClick", XsdBoolean
    )
    advTm: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "advTm", XsdUnsignedInt
    )
    # p14:dur attribute for duration in milliseconds (PowerPoint 2010+)
    dur: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "p14:dur", XsdUnsignedInt
    )

    @classmethod
    def new(cls) -> "CT_SlideTransition":
        """Return a new `p:transition` element."""
        return cast(CT_SlideTransition, parse_xml(f"<p:transition {nsdecls('p')}/>"))

    @classmethod
    def new_morph(
        cls, option: str = "byObject", duration_ms: int = 2000
    ) -> "CT_SlideTransition":
        """Return a new `p:transition` element with morph transition.

        Args:
            option: Morph option - "byObject", "byWord", or "byChar"
            duration_ms: Duration of the transition in milliseconds
        """
        xml = (
            f'<p:transition {nsdecls("p", "p14", "p159")} '
            f'spd="slow" p14:dur="{duration_ms}">'
            f'<p159:morph option="{option}"/>'
            f"</p:transition>"
        )
        return cast(CT_SlideTransition, parse_xml(xml))

    def _remove_all_transition_types(self) -> None:
        """Remove all transition type child elements."""
        # List of all possible transition type elements
        transition_types = [
            "blinds", "checker", "circle", "dissolve", "comb", "cover", "cut",
            "diamond", "fade", "newsflash", "plus", "pull", "push", "random",
            "randomBar", "split", "strips", "wedge", "wheel", "wipe", "zoom",
            "prism", "vortex", "ripple", "honeycomb", "flash", "glitter",
            "doors", "window", "ferris", "gallery", "conveyor", "pan", "reveal",
            "wheelReverse", "flythrough", "warp", "shred", "switch", "flip",
            "cube", "box", "rotate", "orbit", "morph"
        ]
        for child in list(self):
            local_name = child.tag.split("}")[-1] if "}" in child.tag else child.tag.split(":")[-1]
            if local_name in transition_types:
                self.remove(child)


class CT_MorphTransition(BaseOxmlElement):
    """`p159:morph` element for morph transition settings.

    The morph transition animates smooth movement of objects between slides.
    Objects with matching names morph into each other.
    """

    # option attribute: byObject, byWord, byChar
    option: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "option", XsdString
    )

    @staticmethod
    def new(option: str = "byObject") -> "CT_MorphTransition":
        """Return a new `p159:morph` element.

        Args:
            option: "byObject" (default), "byWord", or "byChar"
        """
        return cast(
            CT_MorphTransition,
            parse_xml(f'<p159:morph {nsdecls("p159")} option="{option}"/>'),
        )


# Standard transition type elements (empty elements with optional attributes)

class CT_EmptyTransition(BaseOxmlElement):
    """Base class for empty transition elements like `p:random`, `p:dissolve`, etc."""

    pass


class CT_OrientationTransition(BaseOxmlElement):
    """Transition with dir attribute (horz/vert) like `p:blinds`, `p:comb`, etc."""

    dir: str | None = OptionalAttribute("dir", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_EightDirectionTransition(BaseOxmlElement):
    """Transition with 8-direction dir attribute like `p:cover`, `p:pull`, etc."""

    dir: str | None = OptionalAttribute("dir", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_SideDirectionTransition(BaseOxmlElement):
    """Transition with dir attribute (l/r/u/d) like `p:push`, `p:wipe`, etc."""

    dir: str | None = OptionalAttribute("dir", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_CornerDirectionTransition(BaseOxmlElement):
    """Transition with corner dir attribute like `p:strips`."""

    dir: str | None = OptionalAttribute("dir", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_WheelTransition(BaseOxmlElement):
    """Transition with spokes attribute like `p:wheel`."""

    spokes: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "spokes", XsdUnsignedInt
    )


class CT_SplitTransition(BaseOxmlElement):
    """Transition with orient and dir attributes like `p:split`."""

    orient: str | None = OptionalAttribute("orient", XsdString)  # pyright: ignore[reportAssignmentType]
    dir: str | None = OptionalAttribute("dir", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_ZoomTransition(BaseOxmlElement):
    """Transition with dir attribute like `p:zoom`."""

    dir: str | None = OptionalAttribute("dir", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_OptionalBlackTransition(BaseOxmlElement):
    """Transition with optional thruBlk attribute like `p:cut`, `p:fade`."""

    thruBlk: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "thruBlk", XsdBoolean
    )
