"""Theme-related proxy objects for presentation themes."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from pptx.dml.color import RGBColor
from pptx.shared import ElementProxy, PartElementProxy

if TYPE_CHECKING:
    from pptx.oxml.theme import (
        CT_ColorScheme,
        CT_FontScheme,
        CT_OfficeStyleSheet,
    )
    from pptx.parts.slide import ThemePart


class Theme(PartElementProxy):
    """Proxy for a theme part (``<a:theme>`` element).

    Provides access to the color scheme and font scheme defined in this theme.
    """

    _element: CT_OfficeStyleSheet

    def __init__(self, element: CT_OfficeStyleSheet, part: ThemePart):
        super().__init__(element, part)

    @property
    def color_scheme(self) -> ColorScheme:
        """Return the |ColorScheme| object for this theme."""
        return ColorScheme(self._element.themeElements.clrScheme)

    @property
    def font_scheme(self) -> FontScheme:
        """Return the |FontScheme| object for this theme."""
        return FontScheme(self._element.themeElements.fontScheme)

    @property
    def name(self) -> str:
        """Return the name of this theme (from ``<a:theme name="...">``).

        Returns empty string if no name attribute is present.
        """
        return self._element.get("name", "")


class ColorScheme(ElementProxy):
    """Proxy for the ``<a:clrScheme>`` element.

    Provides named access to each of the 12 theme colors. Each color is
    readable and writable as an |RGBColor| value.
    """

    _element: CT_ColorScheme

    # Maps Python-friendly property names to XML element names
    _COLOR_MAP = {
        "dark1": "dk1",
        "light1": "lt1",
        "dark2": "dk2",
        "light2": "lt2",
        "accent1": "accent1",
        "accent2": "accent2",
        "accent3": "accent3",
        "accent4": "accent4",
        "accent5": "accent5",
        "accent6": "accent6",
        "hyperlink": "hlink",
        "followed_hyperlink": "folHlink",
    }

    def _get_color(self, xml_name: str) -> RGBColor | None:
        """Return the |RGBColor| for the color slot named `xml_name`, or None."""
        color_elm = getattr(self._element, xml_name)
        hex_str = color_elm.rgb_color
        if hex_str is None:
            return None
        return RGBColor.from_string(hex_str)

    def _set_color(self, xml_name: str, value: RGBColor) -> None:
        """Set the color slot named `xml_name` to `value`."""
        color_elm = getattr(self._element, xml_name)
        color_elm.rgb_color = str(value)

    @property
    def name(self) -> str:
        """Return the name of this color scheme."""
        return self._element.name

    @property
    def dark1(self) -> RGBColor | None:
        """Return the dark1 theme color as |RGBColor|."""
        return self._get_color("dk1")

    @dark1.setter
    def dark1(self, value: RGBColor) -> None:
        self._set_color("dk1", value)

    @property
    def light1(self) -> RGBColor | None:
        """Return the light1 theme color as |RGBColor|."""
        return self._get_color("lt1")

    @light1.setter
    def light1(self, value: RGBColor) -> None:
        self._set_color("lt1", value)

    @property
    def dark2(self) -> RGBColor | None:
        """Return the dark2 theme color as |RGBColor|."""
        return self._get_color("dk2")

    @dark2.setter
    def dark2(self, value: RGBColor) -> None:
        self._set_color("dk2", value)

    @property
    def light2(self) -> RGBColor | None:
        """Return the light2 theme color as |RGBColor|."""
        return self._get_color("lt2")

    @light2.setter
    def light2(self, value: RGBColor) -> None:
        self._set_color("lt2", value)

    @property
    def accent1(self) -> RGBColor | None:
        """Return the accent1 theme color as |RGBColor|."""
        return self._get_color("accent1")

    @accent1.setter
    def accent1(self, value: RGBColor) -> None:
        self._set_color("accent1", value)

    @property
    def accent2(self) -> RGBColor | None:
        """Return the accent2 theme color as |RGBColor|."""
        return self._get_color("accent2")

    @accent2.setter
    def accent2(self, value: RGBColor) -> None:
        self._set_color("accent2", value)

    @property
    def accent3(self) -> RGBColor | None:
        """Return the accent3 theme color as |RGBColor|."""
        return self._get_color("accent3")

    @accent3.setter
    def accent3(self, value: RGBColor) -> None:
        self._set_color("accent3", value)

    @property
    def accent4(self) -> RGBColor | None:
        """Return the accent4 theme color as |RGBColor|."""
        return self._get_color("accent4")

    @accent4.setter
    def accent4(self, value: RGBColor) -> None:
        self._set_color("accent4", value)

    @property
    def accent5(self) -> RGBColor | None:
        """Return the accent5 theme color as |RGBColor|."""
        return self._get_color("accent5")

    @accent5.setter
    def accent5(self, value: RGBColor) -> None:
        self._set_color("accent5", value)

    @property
    def accent6(self) -> RGBColor | None:
        """Return the accent6 theme color as |RGBColor|."""
        return self._get_color("accent6")

    @accent6.setter
    def accent6(self, value: RGBColor) -> None:
        self._set_color("accent6", value)

    @property
    def hyperlink(self) -> RGBColor | None:
        """Return the hyperlink theme color as |RGBColor|."""
        return self._get_color("hlink")

    @hyperlink.setter
    def hyperlink(self, value: RGBColor) -> None:
        self._set_color("hlink", value)

    @property
    def followed_hyperlink(self) -> RGBColor | None:
        """Return the followed hyperlink theme color as |RGBColor|."""
        return self._get_color("folHlink")

    @followed_hyperlink.setter
    def followed_hyperlink(self, value: RGBColor) -> None:
        self._set_color("folHlink", value)

    def __iter__(self) -> Iterator[tuple[str, RGBColor | None]]:
        """Iterate over (name, color) pairs for all 12 theme colors."""
        for python_name, xml_name in self._COLOR_MAP.items():
            yield python_name, self._get_color(xml_name)

    def items(self) -> list[tuple[str, RGBColor | None]]:
        """Return a list of (name, color) pairs for all 12 theme colors."""
        return list(self)


class FontScheme(ElementProxy):
    """Proxy for the ``<a:fontScheme>`` element.

    Provides access to the major (heading) and minor (body) font typeface names.
    """

    _element: CT_FontScheme

    @property
    def name(self) -> str:
        """Return the name of this font scheme."""
        return self._element.name

    @property
    def major_font(self) -> str:
        """Return the typeface name of the major (heading) font."""
        latin = self._element.majorFont.latin
        if latin is None:
            return ""
        return latin.typeface

    @major_font.setter
    def major_font(self, value: str) -> None:
        """Set the typeface name of the major (heading) font."""
        latin = self._element.majorFont.latin
        if latin is not None:
            latin.typeface = value

    @property
    def minor_font(self) -> str:
        """Return the typeface name of the minor (body) font."""
        latin = self._element.minorFont.latin
        if latin is None:
            return ""
        return latin.typeface

    @minor_font.setter
    def minor_font(self, value: str) -> None:
        """Set the typeface name of the minor (body) font."""
        latin = self._element.minorFont.latin
        if latin is not None:
            latin.typeface = value
