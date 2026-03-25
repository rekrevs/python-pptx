"""lxml custom element classes for theme-related XML elements."""

from __future__ import annotations

from pptx.oxml import parse_from_template
from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)


class CT_OfficeStyleSheet(BaseOxmlElement):
    """``<a:theme>`` element, root of a theme part."""

    _tag_seq = (
        "a:themeElements",
        "a:objectDefaults",
        "a:extraClrSchemeLst",
        "a:custClrLst",
        "a:extLst",
    )

    themeElements: CT_BaseThemeElements = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:themeElements"
    )

    @classmethod
    def new_default(cls):
        """Return a new ``<a:theme>`` element containing default settings
        suitable for use with a notes master.
        """
        return parse_from_template("theme")


class CT_BaseThemeElements(BaseOxmlElement):
    """``<a:themeElements>`` element, container for color, font, and format schemes."""

    _tag_seq = (
        "a:clrScheme",
        "a:fontScheme",
        "a:fmtScheme",
        "a:extLst",
    )

    clrScheme: CT_ColorScheme = OneAndOnlyOne("a:clrScheme")  # pyright: ignore[reportAssignmentType]
    fontScheme: CT_FontScheme = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:fontScheme"
    )
    fmtScheme: CT_FormatScheme = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:fmtScheme"
    )


class CT_ColorScheme(BaseOxmlElement):
    """``<a:clrScheme>`` element, defines the 12 theme colors."""

    _tag_seq = (
        "a:dk1",
        "a:lt1",
        "a:dk2",
        "a:lt2",
        "a:accent1",
        "a:accent2",
        "a:accent3",
        "a:accent4",
        "a:accent5",
        "a:accent6",
        "a:hlink",
        "a:folHlink",
        "a:extLst",
    )

    dk1: CT_ThemeColor = OneAndOnlyOne("a:dk1")  # pyright: ignore[reportAssignmentType]
    lt1: CT_ThemeColor = OneAndOnlyOne("a:lt1")  # pyright: ignore[reportAssignmentType]
    dk2: CT_ThemeColor = OneAndOnlyOne("a:dk2")  # pyright: ignore[reportAssignmentType]
    lt2: CT_ThemeColor = OneAndOnlyOne("a:lt2")  # pyright: ignore[reportAssignmentType]
    accent1: CT_ThemeColor = OneAndOnlyOne("a:accent1")  # pyright: ignore[reportAssignmentType]
    accent2: CT_ThemeColor = OneAndOnlyOne("a:accent2")  # pyright: ignore[reportAssignmentType]
    accent3: CT_ThemeColor = OneAndOnlyOne("a:accent3")  # pyright: ignore[reportAssignmentType]
    accent4: CT_ThemeColor = OneAndOnlyOne("a:accent4")  # pyright: ignore[reportAssignmentType]
    accent5: CT_ThemeColor = OneAndOnlyOne("a:accent5")  # pyright: ignore[reportAssignmentType]
    accent6: CT_ThemeColor = OneAndOnlyOne("a:accent6")  # pyright: ignore[reportAssignmentType]
    hlink: CT_ThemeColor = OneAndOnlyOne("a:hlink")  # pyright: ignore[reportAssignmentType]
    folHlink: CT_ThemeColor = OneAndOnlyOne("a:folHlink")  # pyright: ignore[reportAssignmentType]

    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_ThemeColor(BaseOxmlElement):
    """Element class for theme color slot elements (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink).

    Each contains exactly one color child element, either ``<a:srgbClr>`` or ``<a:sysClr>``.
    """

    @property
    def rgb_color(self) -> str | None:
        """Return the RGB hex string for this color slot.

        Reads from ``srgbClr/@val`` if present, otherwise from ``sysClr/@lastClr``.
        Returns None if neither is available.
        """
        srgbClr = self.find(qn("a:srgbClr"))
        if srgbClr is not None:
            return srgbClr.val
        sysClr = self.find(qn("a:sysClr"))
        if sysClr is not None:
            return sysClr.lastClr
        return None

    @rgb_color.setter
    def rgb_color(self, hex_str: str) -> None:
        """Set this color slot to an explicit sRGB color.

        Removes any existing color child (srgbClr or sysClr) and replaces with
        a new ``<a:srgbClr>`` element.
        """
        from pptx.oxml.xmlchemy import OxmlElement

        # Remove existing color children
        for child in list(self):
            self.remove(child)
        # Add new srgbClr element
        srgbClr = OxmlElement("a:srgbClr")
        srgbClr.val = hex_str
        self.append(srgbClr)


class CT_FontScheme(BaseOxmlElement):
    """``<a:fontScheme>`` element, defines major and minor fonts."""

    _tag_seq = (
        "a:majorFont",
        "a:minorFont",
        "a:extLst",
    )

    majorFont: CT_FontCollection = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:majorFont"
    )
    minorFont: CT_FontCollection = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:minorFont"
    )

    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_FontCollection(BaseOxmlElement):
    """``<a:majorFont>`` or ``<a:minorFont>`` element.

    Contains latin, ea, cs font elements and optionally script-specific font elements.
    """

    _tag_seq = (
        "a:latin",
        "a:ea",
        "a:cs",
        "a:font",
        "a:extLst",
    )

    latin = ZeroOrOne("a:latin", successors=("a:ea", "a:cs", "a:font", "a:extLst"))
    ea = ZeroOrOne("a:ea", successors=("a:cs", "a:font", "a:extLst"))
    cs = ZeroOrOne("a:cs", successors=("a:font", "a:extLst"))


class CT_FormatScheme(BaseOxmlElement):
    """``<a:fmtScheme>`` element, defines fill, line, and effect styles.

    Stub class for tag_seq ordering support. Full implementation deferred.
    """

    _tag_seq = (
        "a:fillStyleLst",
        "a:lnStyleLst",
        "a:effectStyleLst",
        "a:bgFillStyleLst",
    )

    name: str = OptionalAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
