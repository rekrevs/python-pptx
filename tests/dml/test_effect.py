"""Unit-test suite for `pptx.dml.effect` module."""

from __future__ import annotations

import pytest

from pptx.dml.color import ColorFormat
from pptx.dml.effect import GlowFormat, ReflectionFormat, ShadowFormat, SoftEdgeFormat
from pptx.util import Emu

from ..unitutil.cxml import element, xml


class DescribeShadowFormat(object):
    """Unit-test suite for `pptx.dml.effect.ShadowFormat` objects."""

    def it_knows_whether_it_inherits(self, inherit_get_fixture):
        shadow, expected_value = inherit_get_fixture
        inherit = shadow.inherit
        assert inherit is expected_value

    def it_can_change_whether_it_inherits(self, inherit_set_fixture):
        shadow, value, expected_xml = inherit_set_fixture
        shadow.inherit = value
        assert shadow._element.xml == expected_xml

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p:spPr/a:effectLst/a:outerShdw", "outer"),
            ("p:spPr/a:effectLst/a:innerShdw", "inner"),
            ("p:spPr/a:effectLst", None),
            ("p:spPr", None),
        ],
    )
    def it_knows_its_shadow_type(self, cxml, expected_value):
        shadow = ShadowFormat(element(cxml))
        assert shadow.shadow_type == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p:spPr/a:effectLst/a:outerShdw{dir=5400000}", 90.0),
            ("p:spPr/a:effectLst/a:outerShdw", 0.0),
            ("p:spPr", None),
        ],
    )
    def it_knows_its_angle(self, cxml, expected_value):
        shadow = ShadowFormat(element(cxml))
        assert shadow.angle == expected_value

    def it_can_change_its_angle(self):
        shadow = ShadowFormat(element("p:spPr/a:effectLst/a:outerShdw"))
        shadow.angle = 45.0
        outerShdw = shadow._element.effectLst.outerShdw
        assert outerShdw.dir == 45.0

    def it_raises_on_setting_angle_with_no_shadow(self):
        shadow = ShadowFormat(element("p:spPr"))
        with pytest.raises(ValueError, match="no shadow element"):
            shadow.angle = 90.0

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p:spPr/a:effectLst/a:outerShdw{blurRad=40000}", Emu(40000)),
            ("p:spPr/a:effectLst/a:outerShdw", Emu(0)),
            ("p:spPr", None),
        ],
    )
    def it_knows_its_blur_radius(self, cxml, expected_value):
        shadow = ShadowFormat(element(cxml))
        assert shadow.blur_radius == expected_value

    def it_can_change_its_blur_radius(self):
        shadow = ShadowFormat(element("p:spPr/a:effectLst/a:outerShdw"))
        shadow.blur_radius = 50000
        outerShdw = shadow._element.effectLst.outerShdw
        assert outerShdw.blurRad == Emu(50000)

    def it_raises_on_setting_blur_radius_with_no_shadow(self):
        shadow = ShadowFormat(element("p:spPr"))
        with pytest.raises(ValueError, match="no shadow element"):
            shadow.blur_radius = 40000

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p:spPr/a:effectLst/a:outerShdw{dist=23000}", Emu(23000)),
            ("p:spPr/a:effectLst/a:outerShdw", Emu(0)),
            ("p:spPr", None),
        ],
    )
    def it_knows_its_distance(self, cxml, expected_value):
        shadow = ShadowFormat(element(cxml))
        assert shadow.distance == expected_value

    def it_can_change_its_distance(self):
        shadow = ShadowFormat(element("p:spPr/a:effectLst/a:outerShdw"))
        shadow.distance = 30000
        outerShdw = shadow._element.effectLst.outerShdw
        assert outerShdw.dist == Emu(30000)

    def it_raises_on_setting_distance_with_no_shadow(self):
        shadow = ShadowFormat(element("p:spPr"))
        with pytest.raises(ValueError, match="no shadow element"):
            shadow.distance = 23000

    def it_provides_access_to_its_color(self):
        shadow = ShadowFormat(
            element("p:spPr/a:effectLst/a:outerShdw/a:srgbClr{val=FF0000}")
        )
        color = shadow.color
        assert isinstance(color, ColorFormat)

    def it_raises_on_color_when_no_shadow(self):
        shadow = ShadowFormat(element("p:spPr"))
        with pytest.raises(ValueError, match="no shadow element"):
            shadow.color

    # Also test inner shadow navigates correctly
    def it_reads_inner_shadow_properties(self):
        shadow = ShadowFormat(
            element("p:spPr/a:effectLst/a:innerShdw{blurRad=30000,dist=12000,dir=2700000}")
        )
        assert shadow.shadow_type == "inner"
        assert shadow.blur_radius == Emu(30000)
        assert shadow.distance == Emu(12000)
        assert shadow.angle == 45.0

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("p:spPr", True),
            ("p:spPr/a:effectLst", False),
            ("p:grpSpPr", True),
            ("p:grpSpPr/a:effectLst", False),
        ]
    )
    def inherit_get_fixture(self, request):
        cxml, expected_value = request.param
        shadow = ShadowFormat(element(cxml))
        return shadow, expected_value

    @pytest.fixture(
        params=[
            ("p:spPr{a:b=c}", False, "p:spPr{a:b=c}/a:effectLst"),
            ("p:grpSpPr{a:b=c}", False, "p:grpSpPr{a:b=c}/a:effectLst"),
            ("p:spPr{a:b=c}/a:effectLst", True, "p:spPr{a:b=c}"),
            ("p:grpSpPr{a:b=c}/a:effectLst", True, "p:grpSpPr{a:b=c}"),
            ("p:spPr", True, "p:spPr"),
            ("p:grpSpPr", True, "p:grpSpPr"),
            ("p:spPr/a:effectLst", False, "p:spPr/a:effectLst"),
            ("p:grpSpPr/a:effectLst", False, "p:grpSpPr/a:effectLst"),
        ]
    )
    def inherit_set_fixture(self, request):
        cxml, value, expected_cxml = request.param
        shadow = ShadowFormat(element(cxml))
        expected_value = xml(expected_cxml)
        return shadow, value, expected_value


class DescribeGlowFormat(object):
    """Unit-test suite for `pptx.dml.effect.GlowFormat` objects."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p:spPr/a:effectLst/a:glow{rad=63500}", Emu(63500)),
            ("p:spPr/a:effectLst/a:glow", Emu(0)),
            ("p:spPr", None),
        ],
    )
    def it_knows_its_radius(self, cxml, expected_value):
        glow = GlowFormat(element(cxml))
        assert glow.radius == expected_value

    def it_can_change_its_radius(self):
        glow = GlowFormat(element("p:spPr/a:effectLst/a:glow{rad=10000}"))
        glow.radius = 50000
        glow_elm = glow._element.effectLst.glow
        assert glow_elm.rad == Emu(50000)

    def it_raises_on_setting_radius_with_no_glow(self):
        glow = GlowFormat(element("p:spPr"))
        with pytest.raises(ValueError, match="no glow element"):
            glow.radius = 50000

    def it_provides_access_to_its_color(self):
        glow = GlowFormat(
            element("p:spPr/a:effectLst/a:glow/a:srgbClr{val=00FF00}")
        )
        color = glow.color
        assert isinstance(color, ColorFormat)

    def it_raises_on_color_when_no_glow(self):
        glow = GlowFormat(element("p:spPr"))
        with pytest.raises(ValueError, match="no glow element"):
            glow.color


class DescribeReflectionFormat(object):
    """Unit-test suite for `pptx.dml.effect.ReflectionFormat` objects."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p:spPr/a:effectLst/a:reflection{blurRad=6350}", Emu(6350)),
            ("p:spPr/a:effectLst/a:reflection", Emu(0)),
            ("p:spPr", None),
        ],
    )
    def it_knows_its_blur_radius(self, cxml, expected_value):
        refl = ReflectionFormat(element(cxml))
        assert refl.blur_radius == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p:spPr/a:effectLst/a:reflection{dist=50800}", Emu(50800)),
            ("p:spPr/a:effectLst/a:reflection", Emu(0)),
            ("p:spPr", None),
        ],
    )
    def it_knows_its_distance(self, cxml, expected_value):
        refl = ReflectionFormat(element(cxml))
        assert refl.distance == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p:spPr/a:effectLst/a:reflection{dir=5400000}", 90.0),
            ("p:spPr/a:effectLst/a:reflection", 0.0),
            ("p:spPr", None),
        ],
    )
    def it_knows_its_direction(self, cxml, expected_value):
        refl = ReflectionFormat(element(cxml))
        assert refl.direction == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p:spPr/a:effectLst/a:reflection{stA=100000}", 1.0),
            ("p:spPr/a:effectLst/a:reflection{stA=50000}", 0.5),
            ("p:spPr/a:effectLst/a:reflection", 0.0),
            ("p:spPr", None),
        ],
    )
    def it_knows_its_start_opacity(self, cxml, expected_value):
        refl = ReflectionFormat(element(cxml))
        assert refl.start_opacity == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p:spPr/a:effectLst/a:reflection{endA=30000}", 0.3),
            ("p:spPr/a:effectLst/a:reflection{endA=0}", 0.0),
            ("p:spPr/a:effectLst/a:reflection", 0.0),
            ("p:spPr", None),
        ],
    )
    def it_knows_its_end_opacity(self, cxml, expected_value):
        refl = ReflectionFormat(element(cxml))
        assert refl.end_opacity == expected_value


class DescribeSoftEdgeFormat(object):
    """Unit-test suite for `pptx.dml.effect.SoftEdgeFormat` objects."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p:spPr/a:effectLst/a:softEdge{rad=25400}", Emu(25400)),
            ("p:spPr", None),
        ],
    )
    def it_knows_its_radius(self, cxml, expected_value):
        soft_edge = SoftEdgeFormat(element(cxml))
        assert soft_edge.radius == expected_value

    def it_can_change_its_radius(self):
        soft_edge = SoftEdgeFormat(element("p:spPr/a:effectLst/a:softEdge{rad=25400}"))
        soft_edge.radius = 50800
        softEdge_elm = soft_edge._element.effectLst.softEdge
        assert softEdge_elm.rad == Emu(50800)

    def it_raises_on_setting_radius_with_no_soft_edge(self):
        soft_edge = SoftEdgeFormat(element("p:spPr"))
        with pytest.raises(ValueError, match="no soft edge element"):
            soft_edge.radius = 50800
