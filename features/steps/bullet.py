"""Gherkin step implementations for bullet formatting features."""

from __future__ import annotations

from behave import given, then, when

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.text.text import _Paragraph


# given ===================================================


@given("a paragraph with no explicit bullet setting")
def given_a_paragraph_with_no_explicit_bullet_setting(context):
    p_xml = "<a:p %s><a:r><a:t>text</a:t></a:r></a:p>" % nsdecls("a")
    p = parse_xml(p_xml)
    context.paragraph = _Paragraph(p, None)


@given("a paragraph with a character bullet")
def given_a_paragraph_with_a_character_bullet(context):
    p_xml = (
        '<a:p %s><a:pPr><a:buChar char="-"/></a:pPr>'
        "<a:r><a:t>text</a:t></a:r></a:p>" % nsdecls("a")
    )
    p = parse_xml(p_xml)
    context.paragraph = _Paragraph(p, None)


@given('a paragraph with bullet font set to "{font_name}"')
def given_a_paragraph_with_bullet_font(context, font_name):
    p_xml = (
        '<a:p %s><a:pPr><a:buFont typeface="%s"/></a:pPr>'
        "<a:r><a:t>text</a:t></a:r></a:p>" % (nsdecls("a"), font_name)
    )
    p = parse_xml(p_xml)
    context.paragraph = _Paragraph(p, None)


@given("a paragraph with bullet size set to {size}")
def given_a_paragraph_with_bullet_size(context, size):
    pts_val = int(float(size) * 100)
    p_xml = (
        '<a:p %s><a:pPr><a:buSzPts val="%d"/></a:pPr>'
        "<a:r><a:t>text</a:t></a:r></a:p>" % (nsdecls("a"), pts_val)
    )
    p = parse_xml(p_xml)
    context.paragraph = _Paragraph(p, None)


# when ====================================================


@when('I assign "{value}" to paragraph.bullet.type')
def when_I_assign_string_to_bullet_type(context, value):
    context.paragraph.bullet.type = value


@when("I assign None to paragraph.bullet.type")
def when_I_assign_None_to_bullet_type(context):
    context.paragraph.bullet.type = None


@when('I assign "{value}" to paragraph.bullet.char')
def when_I_assign_to_bullet_char(context, value):
    context.paragraph.bullet.char = value


@when('I assign "{value}" to paragraph.bullet.auto_num_type')
def when_I_assign_to_bullet_auto_num_type(context, value):
    context.paragraph.bullet.auto_num_type = value


@when('I assign "{value}" to paragraph.bullet.font')
def when_I_assign_string_to_bullet_font(context, value):
    context.paragraph.bullet.font = value


@when("I assign None to paragraph.bullet.font")
def when_I_assign_None_to_bullet_font(context):
    context.paragraph.bullet.font = None


@when("I assign {size:g} to paragraph.bullet.size")
def when_I_assign_to_bullet_size(context, size):
    context.paragraph.bullet.size = size


@when("I assign None to paragraph.bullet.size")
def when_I_assign_None_to_bullet_size(context):
    context.paragraph.bullet.size = None


@when("I assign RGBColor(0xFF, 0x00, 0x00) to paragraph.bullet.color.rgb")
def when_I_assign_rgb_to_bullet_color(context):
    context.paragraph.bullet.color.rgb = RGBColor(0xFF, 0x00, 0x00)


# then ====================================================


@then("paragraph.bullet.type is None")
def then_bullet_type_is_None(context):
    assert context.paragraph.bullet.type is None


@then('paragraph.bullet.type is "{expected_value}"')
def then_bullet_type_is_value(context, expected_value):
    assert context.paragraph.bullet.type == expected_value


@then("paragraph.bullet.char is not None")
def then_bullet_char_is_not_None(context):
    assert context.paragraph.bullet.char is not None


@then('paragraph.bullet.char is "{expected_value}"')
def then_bullet_char_is_value(context, expected_value):
    assert context.paragraph.bullet.char == expected_value


@then('paragraph.bullet.auto_num_type is "{expected_value}"')
def then_bullet_auto_num_type_is_value(context, expected_value):
    assert context.paragraph.bullet.auto_num_type == expected_value


@then('paragraph.bullet.font is "{expected_value}"')
def then_bullet_font_is_value(context, expected_value):
    assert context.paragraph.bullet.font == expected_value


@then("paragraph.bullet.font is None")
def then_bullet_font_is_None(context):
    assert context.paragraph.bullet.font is None


@then("paragraph.bullet.size is {expected_value:g}")
def then_bullet_size_is_value(context, expected_value):
    assert context.paragraph.bullet.size == expected_value


@then("paragraph.bullet.size is None")
def then_bullet_size_is_None(context):
    assert context.paragraph.bullet.size is None


@then("paragraph.bullet.color.rgb is RGBColor(0xFF, 0x00, 0x00)")
def then_bullet_color_rgb_is_red(context):
    assert context.paragraph.bullet.color.rgb == RGBColor(0xFF, 0x00, 0x00)
