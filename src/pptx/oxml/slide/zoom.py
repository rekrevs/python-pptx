"""Custom element classes for PowerPoint Zoom features.

Zoom features include Slide Zoom, Section Zoom, and Summary Zoom (PowerPoint 2016+).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import XsdBoolean, XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)

if TYPE_CHECKING:
    pass


class CT_ZoomObjectProperties(BaseOxmlElement):
    """`p166:zmPr` element for zoom object properties.

    Contains the visual properties for a zoom object including the thumbnail
    image and shape properties.

    Attributes:
        id: Unique identifier (GUID) for this zoom object
        returnToParent: Whether to return to the parent slide after zoom (default: true)
        imageType: "preview" or "cover" (default: "preview")
        showBg: Whether to show background (default: true)
    """

    _tag_seq = ("a:blipFill", "a:spPr")
    blipFill = OneAndOnlyOne("a:blipFill")
    spPr = OneAndOnlyOne("a:spPr")
    del _tag_seq

    # Attributes
    id: str = RequiredAttribute("id", XsdString)  # pyright: ignore[reportAssignmentType]
    returnToParent: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "returnToParent", XsdBoolean
    )
    imageType: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "imageType", XsdString
    )
    showBg: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "showBg", XsdBoolean
    )


class CT_SlideZoomObject(BaseOxmlElement):
    """`pslz:sldZmObj` element for a slide zoom object.

    References a specific slide to zoom to.

    Child Elements:
        zmPr: Zoom object properties (thumbnail, shape)

    Attributes:
        sldId: Target slide ID (required)
        cId: Connection ID (optional)
    """

    _tag_seq = ("p166:zmPr", "p166:extLst")
    zmPr: CT_ZoomObjectProperties = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p166:zmPr"
    )
    extLst = ZeroOrOne("p166:extLst", successors=())
    del _tag_seq

    # Attributes
    sldId: int = RequiredAttribute("sldId", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
    cId: int | None = OptionalAttribute("cId", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


class CT_SlideZoom(BaseOxmlElement):
    """`pslz:sldZm` element, root element for slide zoom.

    A slide zoom object displays a thumbnail of another slide and
    navigates to it with a zoom transition when clicked.
    """

    _tag_seq = ("pslz:sldZmObj", "pslz:extLst")
    sldZmObj: CT_SlideZoomObject = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "pslz:sldZmObj"
    )
    extLst = ZeroOrOne("pslz:extLst", successors=())
    del _tag_seq

    @classmethod
    def new(cls, target_slide_id: int, zoom_guid: str) -> "CT_SlideZoom":
        """Return a new `pslz:sldZm` element.

        Args:
            target_slide_id: The slide ID to zoom to
            zoom_guid: Unique GUID identifier for this zoom object
        """
        xml = (
            f'<pslz:sldZm {nsdecls("pslz", "p166", "a")}>'
            f'<pslz:sldZmObj sldId="{target_slide_id}">'
            f'<p166:zmPr id="{zoom_guid}" returnToParent="1" imageType="preview" showBg="1">'
            f"<a:blipFill/>"
            f"<a:spPr/>"
            f"</p166:zmPr>"
            f"</pslz:sldZmObj>"
            f"</pslz:sldZm>"
        )
        return cast(CT_SlideZoom, parse_xml(xml))


class CT_SectionZoomObject(BaseOxmlElement):
    """`psecZm:secZmObj` element for a section zoom object.

    References a specific section to zoom to.
    """

    _tag_seq = ("p166:zmPr", "p166:extLst")
    zmPr: CT_ZoomObjectProperties = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p166:zmPr"
    )
    extLst = ZeroOrOne("p166:extLst", successors=())
    del _tag_seq

    # Attributes
    secId: int = RequiredAttribute("secId", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


class CT_SummaryZoomSection(BaseOxmlElement):
    """`psumZm:sumZmSec` element for a section in summary zoom."""

    _tag_seq = ("p166:zmPr",)
    zmPr: CT_ZoomObjectProperties = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "p166:zmPr"
    )
    del _tag_seq

    # Attributes
    secId: int = RequiredAttribute("secId", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


# Note: Full Section Zoom and Summary Zoom implementation would require
# additional schema research and integration with presentation sections.
# The current implementation provides the core slide zoom functionality.
