"""Unit-test suite for `pptx.oxml.theme` module."""

from __future__ import annotations

import pytest

from pptx.oxml.theme import (
    CT_BaseThemeElements,
    CT_ColorScheme,
    CT_FontCollection,
    CT_FontScheme,
    CT_OfficeStyleSheet,
    CT_ThemeColor,
)

from ..unitutil.file import snippet_text


class DescribeCT_OfficeStyleSheet(object):
    def it_can_create_a_default_theme_element(self, new_fixture):
        expected_xml = new_fixture
        theme = CT_OfficeStyleSheet.new_default()
        assert theme.xml == expected_xml

    def it_provides_access_to_themeElements(self):
        theme = CT_OfficeStyleSheet.new_default()
        themeElements = theme.themeElements
        assert isinstance(themeElements, CT_BaseThemeElements)

    def it_provides_access_to_clrScheme_via_themeElements(self):
        theme = CT_OfficeStyleSheet.new_default()
        clrScheme = theme.themeElements.clrScheme
        assert isinstance(clrScheme, CT_ColorScheme)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def new_fixture(self):
        expected_xml = snippet_text("default-theme")
        return expected_xml


class DescribeCT_BaseThemeElements(object):
    def it_provides_access_to_clrScheme(self):
        theme = CT_OfficeStyleSheet.new_default()
        themeElements = theme.themeElements
        clrScheme = themeElements.clrScheme
        assert isinstance(clrScheme, CT_ColorScheme)

    def it_provides_access_to_fontScheme(self):
        theme = CT_OfficeStyleSheet.new_default()
        themeElements = theme.themeElements
        fontScheme = themeElements.fontScheme
        assert isinstance(fontScheme, CT_FontScheme)


class DescribeCT_ColorScheme(object):
    """Unit-test suite for `pptx.oxml.theme.CT_ColorScheme` objects."""

    def it_provides_access_to_each_named_color_element(self):
        theme = CT_OfficeStyleSheet.new_default()
        clrScheme = theme.themeElements.clrScheme
        color_names = [
            "dk1", "lt1", "dk2", "lt2",
            "accent1", "accent2", "accent3", "accent4",
            "accent5", "accent6", "hlink", "folHlink",
        ]
        for name in color_names:
            color_elm = getattr(clrScheme, name)
            assert isinstance(color_elm, CT_ThemeColor), (
                f"Expected CT_ThemeColor for {name}, got {type(color_elm)}"
            )

    def it_can_read_srgbClr_val_from_color_elements(self):
        theme = CT_OfficeStyleSheet.new_default()
        clrScheme = theme.themeElements.clrScheme
        # accent1 uses srgbClr val="4F81BD" in the default theme
        assert clrScheme.accent1.rgb_color == "4F81BD"

    def it_can_read_sysClr_lastClr_from_color_elements(self):
        theme = CT_OfficeStyleSheet.new_default()
        clrScheme = theme.themeElements.clrScheme
        # dk1 uses sysClr val="windowText" lastClr="000000"
        assert clrScheme.dk1.rgb_color == "000000"
        # lt1 uses sysClr val="window" lastClr="FFFFFF"
        assert clrScheme.lt1.rgb_color == "FFFFFF"

    def it_can_set_color_to_srgbClr(self):
        theme = CT_OfficeStyleSheet.new_default()
        clrScheme = theme.themeElements.clrScheme

        # Set accent1 to red
        clrScheme.accent1.rgb_color = "FF0000"
        assert clrScheme.accent1.rgb_color == "FF0000"

        # Set dk1 (which starts as sysClr) to a specific color
        clrScheme.dk1.rgb_color = "123456"
        assert clrScheme.dk1.rgb_color == "123456"

    def it_has_a_name_attribute(self):
        theme = CT_OfficeStyleSheet.new_default()
        clrScheme = theme.themeElements.clrScheme
        assert clrScheme.name == "Office"

    @pytest.mark.parametrize(
        ("color_name", "expected_rgb"),
        [
            ("dk1", "000000"),
            ("lt1", "FFFFFF"),
            ("dk2", "1F497D"),
            ("lt2", "EEECE1"),
            ("accent1", "4F81BD"),
            ("accent2", "C0504D"),
            ("accent3", "9BBB59"),
            ("accent4", "8064A2"),
            ("accent5", "4BACC6"),
            ("accent6", "F79646"),
            ("hlink", "0000FF"),
            ("folHlink", "800080"),
        ],
    )
    def it_can_read_all_12_default_theme_colors(self, color_name, expected_rgb):
        theme = CT_OfficeStyleSheet.new_default()
        clrScheme = theme.themeElements.clrScheme
        color_elm = getattr(clrScheme, color_name)
        assert color_elm.rgb_color == expected_rgb


class DescribeCT_FontScheme(object):
    """Unit-test suite for `pptx.oxml.theme.CT_FontScheme` objects."""

    def it_provides_access_to_majorFont(self):
        theme = CT_OfficeStyleSheet.new_default()
        fontScheme = theme.themeElements.fontScheme
        majorFont = fontScheme.majorFont
        assert isinstance(majorFont, CT_FontCollection)

    def it_provides_access_to_minorFont(self):
        theme = CT_OfficeStyleSheet.new_default()
        fontScheme = theme.themeElements.fontScheme
        minorFont = fontScheme.minorFont
        assert isinstance(minorFont, CT_FontCollection)

    def it_has_a_name_attribute(self):
        theme = CT_OfficeStyleSheet.new_default()
        fontScheme = theme.themeElements.fontScheme
        assert fontScheme.name == "Office"


class DescribeCT_FontCollection(object):
    """Unit-test suite for `pptx.oxml.theme.CT_FontCollection` objects."""

    def it_provides_access_to_latin_typeface(self):
        theme = CT_OfficeStyleSheet.new_default()
        majorFont = theme.themeElements.fontScheme.majorFont
        latin = majorFont.latin
        assert latin is not None
        assert latin.typeface == "Calibri"

    def it_provides_access_to_ea_typeface(self):
        theme = CT_OfficeStyleSheet.new_default()
        majorFont = theme.themeElements.fontScheme.majorFont
        ea = majorFont.ea
        assert ea is not None
        assert ea.typeface == ""

    def it_provides_access_to_cs_typeface(self):
        theme = CT_OfficeStyleSheet.new_default()
        majorFont = theme.themeElements.fontScheme.majorFont
        cs = majorFont.cs
        assert cs is not None
        assert cs.typeface == ""
