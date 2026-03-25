"""Unit-test suite for `pptx.oxml.dml.effect` module."""

from __future__ import annotations

import pytest

from pptx.oxml.dml.effect import (
    CT_EffectList,
    CT_GlowEffect,
    CT_InnerShadowEffect,
    CT_OuterShadowEffect,
    CT_ReflectionEffect,
    CT_SoftEdgesEffect,
)
from pptx.util import Emu

from ..unitutil.cxml import element


class DescribeCT_EffectList:
    """Unit-test suite for `pptx.oxml.dml.effect.CT_EffectList` objects."""

    def it_can_access_outerShdw_child(self):
        effectLst = element("a:effectLst/a:outerShdw")
        outerShdw = effectLst.outerShdw
        assert outerShdw is not None
        assert isinstance(outerShdw, CT_OuterShadowEffect)

    def it_returns_None_when_no_outerShdw_child(self):
        effectLst = element("a:effectLst")
        assert effectLst.outerShdw is None

    def it_can_access_innerShdw_child(self):
        effectLst = element("a:effectLst/a:innerShdw")
        innerShdw = effectLst.innerShdw
        assert innerShdw is not None
        assert isinstance(innerShdw, CT_InnerShadowEffect)

    def it_returns_None_when_no_innerShdw_child(self):
        effectLst = element("a:effectLst")
        assert effectLst.innerShdw is None

    def it_can_access_glow_child(self):
        effectLst = element("a:effectLst/a:glow")
        glow = effectLst.glow
        assert glow is not None
        assert isinstance(glow, CT_GlowEffect)

    def it_returns_None_when_no_glow_child(self):
        effectLst = element("a:effectLst")
        assert effectLst.glow is None

    def it_can_access_reflection_child(self):
        effectLst = element("a:effectLst/a:reflection")
        reflection = effectLst.reflection
        assert reflection is not None
        assert isinstance(reflection, CT_ReflectionEffect)

    def it_returns_None_when_no_reflection_child(self):
        effectLst = element("a:effectLst")
        assert effectLst.reflection is None

    def it_can_access_softEdge_child(self):
        effectLst = element("a:effectLst/a:softEdge{rad=40000}")
        softEdge = effectLst.softEdge
        assert softEdge is not None
        assert isinstance(softEdge, CT_SoftEdgesEffect)

    def it_returns_None_when_no_softEdge_child(self):
        effectLst = element("a:effectLst")
        assert effectLst.softEdge is None


class DescribeCT_OuterShadowEffect:
    """Unit-test suite for `pptx.oxml.dml.effect.CT_OuterShadowEffect` objects."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("a:outerShdw{blurRad=40000}", Emu(40000)),
            ("a:outerShdw", None),
        ],
    )
    def it_can_read_blurRad(self, cxml, expected_value):
        outerShdw = element(cxml)
        assert outerShdw.blurRad == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("a:outerShdw{dist=23000}", Emu(23000)),
            ("a:outerShdw", None),
        ],
    )
    def it_can_read_dist(self, cxml, expected_value):
        outerShdw = element(cxml)
        assert outerShdw.dist == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("a:outerShdw{dir=5400000}", 90.0),
            ("a:outerShdw", None),
        ],
    )
    def it_can_read_dir(self, cxml, expected_value):
        outerShdw = element(cxml)
        assert outerShdw.dir == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("a:outerShdw{rotWithShape=true}", True),
            ("a:outerShdw{rotWithShape=false}", False),
            ("a:outerShdw", None),
        ],
    )
    def it_can_read_rotWithShape(self, cxml, expected_value):
        outerShdw = element(cxml)
        assert outerShdw.rotWithShape == expected_value

    def it_can_access_color_choice_child(self):
        outerShdw = element("a:outerShdw/a:srgbClr{val=FF0000}")
        color = outerShdw.eg_colorChoice
        assert color is not None

    def it_returns_None_when_no_color_choice(self):
        outerShdw = element("a:outerShdw")
        assert outerShdw.eg_colorChoice is None


class DescribeCT_InnerShadowEffect:
    """Unit-test suite for `pptx.oxml.dml.effect.CT_InnerShadowEffect` objects."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("a:innerShdw{blurRad=30000}", Emu(30000)),
            ("a:innerShdw", None),
        ],
    )
    def it_can_read_blurRad(self, cxml, expected_value):
        innerShdw = element(cxml)
        assert innerShdw.blurRad == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("a:innerShdw{dist=12000}", Emu(12000)),
            ("a:innerShdw", None),
        ],
    )
    def it_can_read_dist(self, cxml, expected_value):
        innerShdw = element(cxml)
        assert innerShdw.dist == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("a:innerShdw{dir=2700000}", 45.0),
            ("a:innerShdw", None),
        ],
    )
    def it_can_read_dir(self, cxml, expected_value):
        innerShdw = element(cxml)
        assert innerShdw.dir == expected_value

    def it_can_access_color_choice_child(self):
        innerShdw = element("a:innerShdw/a:schemeClr{val=accent1}")
        color = innerShdw.eg_colorChoice
        assert color is not None


class DescribeCT_GlowEffect:
    """Unit-test suite for `pptx.oxml.dml.effect.CT_GlowEffect` objects."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("a:glow{rad=63500}", Emu(63500)),
            ("a:glow", None),
        ],
    )
    def it_can_read_rad(self, cxml, expected_value):
        glow = element(cxml)
        assert glow.rad == expected_value

    def it_can_access_color_choice_child(self):
        glow = element("a:glow/a:srgbClr{val=00FF00}")
        color = glow.eg_colorChoice
        assert color is not None

    def it_returns_None_when_no_color_choice(self):
        glow = element("a:glow")
        assert glow.eg_colorChoice is None


class DescribeCT_ReflectionEffect:
    """Unit-test suite for `pptx.oxml.dml.effect.CT_ReflectionEffect` objects."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("a:reflection{blurRad=6350}", Emu(6350)),
            ("a:reflection", None),
        ],
    )
    def it_can_read_blurRad(self, cxml, expected_value):
        reflection = element(cxml)
        assert reflection.blurRad == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("a:reflection{dist=50800}", Emu(50800)),
            ("a:reflection", None),
        ],
    )
    def it_can_read_dist(self, cxml, expected_value):
        reflection = element(cxml)
        assert reflection.dist == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("a:reflection{dir=5400000}", 90.0),
            ("a:reflection", None),
        ],
    )
    def it_can_read_dir(self, cxml, expected_value):
        reflection = element(cxml)
        assert reflection.dir == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("a:reflection{stA=100000}", 1.0),
            ("a:reflection{stA=50000}", 0.5),
            ("a:reflection", None),
        ],
    )
    def it_can_read_stA(self, cxml, expected_value):
        reflection = element(cxml)
        assert reflection.stA == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("a:reflection{endA=0}", 0.0),
            ("a:reflection{endA=30000}", 0.3),
            ("a:reflection", None),
        ],
    )
    def it_can_read_endA(self, cxml, expected_value):
        reflection = element(cxml)
        assert reflection.endA == expected_value


class DescribeCT_SoftEdgesEffect:
    """Unit-test suite for `pptx.oxml.dml.effect.CT_SoftEdgesEffect` objects."""

    def it_can_read_rad(self):
        softEdge = element("a:softEdge{rad=25400}")
        assert softEdge.rad == Emu(25400)
