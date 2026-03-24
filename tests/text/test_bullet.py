# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.text.bullet` module."""

from __future__ import annotations

import pytest

from pptx.dml.color import ColorFormat, RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.text.bullet import BulletFormat
from pptx.text.text import _Paragraph

from ..unitutil.cxml import element, xml


class DescribeBulletFormat:
    """Unit-test suite for `pptx.text.bullet.BulletFormat` object."""

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_value"),
        [
            ("a:pPr", None),
            ("a:pPr/a:buNone", "none"),
            ("a:pPr/a:buChar{char=-}", "char"),
            ("a:pPr/a:buAutoNum{type=arabicPeriod}", "auto_num"),
        ],
    )
    def it_knows_its_type(self, pPr_cxml: str, expected_value: str | None):
        pPr = element(pPr_cxml)
        bullet = BulletFormat(pPr)
        assert bullet.type == expected_value

    @pytest.mark.parametrize(
        ("pPr_cxml", "value", "expected_cxml"),
        [
            ("a:pPr", "none", "a:pPr/a:buNone"),
            ("a:pPr", "auto_num", "a:pPr/a:buAutoNum{type=arabicPeriod}"),
            ("a:pPr/a:buChar{char=-}", "none", "a:pPr/a:buNone"),
            ("a:pPr/a:buNone", "auto_num", "a:pPr/a:buAutoNum{type=arabicPeriod}"),
            ("a:pPr/a:buChar{char=-}", None, "a:pPr"),
            ("a:pPr/a:buAutoNum{type=arabicPeriod}", None, "a:pPr"),
            ("a:pPr/a:buNone", None, "a:pPr"),
        ],
    )
    def it_can_change_its_type(self, pPr_cxml: str, value: str | None, expected_cxml: str):
        pPr = element(pPr_cxml)
        bullet = BulletFormat(pPr)
        bullet.type = value
        assert pPr.xml == xml(expected_cxml)

    def it_sets_default_bullet_char_when_type_set_to_char(self):
        pPr = element("a:pPr")
        bullet = BulletFormat(pPr)
        bullet.type = "char"
        assert bullet.char == "\u2022"

    def it_sets_default_auto_num_when_type_set_to_auto_num_from_char(self):
        pPr = element("a:pPr/a:buChar{char=-}")
        bullet = BulletFormat(pPr)
        bullet.type = "auto_num"
        assert bullet.auto_num_type == "arabicPeriod"

    def it_raises_on_invalid_type(self):
        pPr = element("a:pPr")
        bullet = BulletFormat(pPr)
        with pytest.raises(ValueError, match="bullet type must be"):
            bullet.type = "invalid"

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_value"),
        [
            ("a:pPr", None),
            ("a:pPr/a:buChar{char=-}", "-"),
            ("a:pPr/a:buNone", None),
        ],
    )
    def it_knows_its_char(self, pPr_cxml: str, expected_value: str | None):
        pPr = element(pPr_cxml)
        bullet = BulletFormat(pPr)
        assert bullet.char == expected_value

    @pytest.mark.parametrize(
        ("pPr_cxml", "value", "expected_cxml"),
        [
            ("a:pPr", "-", "a:pPr/a:buChar{char=-}"),
            ("a:pPr/a:buChar{char=-}", ".", "a:pPr/a:buChar{char=.}"),
            ("a:pPr/a:buNone", "-", "a:pPr/a:buChar{char=-}"),
            ("a:pPr/a:buAutoNum{type=arabicPeriod}", "-", "a:pPr/a:buChar{char=-}"),
        ],
    )
    def it_can_set_its_char(self, pPr_cxml: str, value: str, expected_cxml: str):
        pPr = element(pPr_cxml)
        bullet = BulletFormat(pPr)
        bullet.char = value
        assert pPr.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_value"),
        [
            ("a:pPr", None),
            ("a:pPr/a:buAutoNum{type=arabicPeriod}", "arabicPeriod"),
            ("a:pPr/a:buAutoNum{type=romanLcPeriod}", "romanLcPeriod"),
            ("a:pPr/a:buNone", None),
        ],
    )
    def it_knows_its_auto_num_type(self, pPr_cxml: str, expected_value: str | None):
        pPr = element(pPr_cxml)
        bullet = BulletFormat(pPr)
        assert bullet.auto_num_type == expected_value

    @pytest.mark.parametrize(
        ("pPr_cxml", "value", "expected_cxml"),
        [
            ("a:pPr", "arabicPeriod", "a:pPr/a:buAutoNum{type=arabicPeriod}"),
            (
                "a:pPr/a:buAutoNum{type=arabicPeriod}",
                "romanLcPeriod",
                "a:pPr/a:buAutoNum{type=romanLcPeriod}",
            ),
            ("a:pPr/a:buChar{char=-}", "arabicPeriod", "a:pPr/a:buAutoNum{type=arabicPeriod}"),
            ("a:pPr/a:buNone", "arabicPeriod", "a:pPr/a:buAutoNum{type=arabicPeriod}"),
        ],
    )
    def it_can_set_its_auto_num_type(self, pPr_cxml: str, value: str, expected_cxml: str):
        pPr = element(pPr_cxml)
        bullet = BulletFormat(pPr)
        bullet.auto_num_type = value
        assert pPr.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_value"),
        [
            ("a:pPr", None),
            ("a:pPr/a:buFont{typeface=Arial}", "Arial"),
            ("a:pPr/a:buFont{typeface=Wingdings}", "Wingdings"),
        ],
    )
    def it_knows_its_font(self, pPr_cxml: str, expected_value: str | None):
        pPr = element(pPr_cxml)
        bullet = BulletFormat(pPr)
        assert bullet.font == expected_value

    @pytest.mark.parametrize(
        ("pPr_cxml", "value", "expected_cxml"),
        [
            ("a:pPr", "Arial", "a:pPr/a:buFont{typeface=Arial}"),
            (
                "a:pPr/a:buFont{typeface=Arial}",
                "Wingdings",
                "a:pPr/a:buFont{typeface=Wingdings}",
            ),
            ("a:pPr/a:buFont{typeface=Arial}", None, "a:pPr"),
        ],
    )
    def it_can_set_its_font(self, pPr_cxml: str, value: str | None, expected_cxml: str):
        pPr = element(pPr_cxml)
        bullet = BulletFormat(pPr)
        bullet.font = value
        assert pPr.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("pPr_cxml", "expected_value"),
        [
            ("a:pPr", None),
            ("a:pPr/a:buSzPts{val=1200}", 12.0),
            ("a:pPr/a:buSzPts{val=2400}", 24.0),
            ("a:pPr/a:buSzPts{val=100}", 1.0),
        ],
    )
    def it_knows_its_size(self, pPr_cxml: str, expected_value: float | None):
        pPr = element(pPr_cxml)
        bullet = BulletFormat(pPr)
        assert bullet.size == expected_value

    @pytest.mark.parametrize(
        ("pPr_cxml", "value", "expected_cxml"),
        [
            ("a:pPr", 12.0, "a:pPr/a:buSzPts{val=1200}"),
            ("a:pPr/a:buSzPts{val=1200}", 24.0, "a:pPr/a:buSzPts{val=2400}"),
            ("a:pPr/a:buSzPts{val=1200}", None, "a:pPr"),
        ],
    )
    def it_can_set_its_size(self, pPr_cxml: str, value: float | None, expected_cxml: str):
        pPr = element(pPr_cxml)
        bullet = BulletFormat(pPr)
        bullet.size = value
        assert pPr.xml == xml(expected_cxml)

    def it_provides_access_to_color(self):
        pPr = element("a:pPr")
        bullet = BulletFormat(pPr)
        color = bullet.color
        assert isinstance(color, ColorFormat)

    def it_can_set_rgb_color(self):
        pPr = element("a:pPr")
        bullet = BulletFormat(pPr)
        bullet.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        assert bullet.color.rgb == RGBColor(0xFF, 0x00, 0x00)

    def it_can_set_theme_color(self):
        pPr = element("a:pPr")
        bullet = BulletFormat(pPr)
        bullet.color.theme_color = MSO_THEME_COLOR.ACCENT_1
        assert bullet.color.theme_color == MSO_THEME_COLOR.ACCENT_1


class Describe_Paragraph_Bullet:
    """Unit-test suite for the `_Paragraph.bullet` property."""

    def it_provides_access_to_a_BulletFormat(self):
        p = element("a:p")
        paragraph = _Paragraph(p, None)
        bullet = paragraph.bullet
        assert isinstance(bullet, BulletFormat)

    def it_returns_correct_type_for_bulleted_paragraph(self):
        p = element("a:p/a:pPr/a:buChar{char=-}")
        paragraph = _Paragraph(p, None)
        assert paragraph.bullet.type == "char"
        assert paragraph.bullet.char == "-"

    def it_returns_correct_type_for_numbered_paragraph(self):
        p = element("a:p/a:pPr/a:buAutoNum{type=arabicPeriod}")
        paragraph = _Paragraph(p, None)
        assert paragraph.bullet.type == "auto_num"
        assert paragraph.bullet.auto_num_type == "arabicPeriod"

    def it_returns_none_type_for_paragraph_with_no_bullets(self):
        p = element("a:p/a:pPr/a:buNone")
        paragraph = _Paragraph(p, None)
        assert paragraph.bullet.type == "none"

    def it_preserves_other_pPr_children_when_setting_bullet(self):
        p = element("a:p/a:pPr{algn=ctr}")
        paragraph = _Paragraph(p, None)
        paragraph.bullet.type = "char"
        # alignment should still be there
        assert paragraph.alignment is not None
        assert paragraph.bullet.type == "char"
