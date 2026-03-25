"""Gherkin step implementations for theme feature."""

from __future__ import annotations

from behave import given, then, when
from helpers import test_pptx

from pptx import Presentation
from pptx.dml.color import RGBColor


# given ===================================================


@given("a presentation with a default theme")
def given_a_presentation_with_a_default_theme(context):
    context.prs = Presentation(test_pptx("thm-theme"))
    context.slide_master = context.prs.slide_masters[0]
    context.theme = context.slide_master.theme


# when ====================================================


@when('I set the theme color scheme {color_name} to "{hex_value}"')
def when_I_set_the_theme_color_scheme_color(context, color_name, hex_value):
    color_scheme = context.theme.color_scheme
    setattr(color_scheme, color_name, RGBColor.from_string(hex_value))


@when('I set the theme font scheme major font to "{typeface}"')
def when_I_set_the_theme_font_scheme_major_font(context, typeface):
    context.theme.font_scheme.major_font = typeface


@when('I set the theme font scheme minor font to "{typeface}"')
def when_I_set_the_theme_font_scheme_minor_font(context, typeface):
    context.theme.font_scheme.minor_font = typeface


# then ====================================================


@then('the theme color scheme {color_name} is "{expected_hex}"')
def then_the_theme_color_scheme_color_is(context, color_name, expected_hex):
    # Reload theme from presentation in case it was just reloaded
    theme = context.prs.slide_masters[0].theme
    color_scheme = theme.color_scheme
    actual = getattr(color_scheme, color_name)
    expected = RGBColor.from_string(expected_hex)
    assert actual == expected, (
        "expected %s for %s, got %s" % (expected_hex, color_name, actual)
    )


@then('the theme font scheme major font is "{expected_typeface}"')
def then_the_theme_font_scheme_major_font_is(context, expected_typeface):
    theme = context.prs.slide_masters[0].theme
    actual = theme.font_scheme.major_font
    assert actual == expected_typeface, (
        "expected '%s', got '%s'" % (expected_typeface, actual)
    )


@then('the theme font scheme minor font is "{expected_typeface}"')
def then_the_theme_font_scheme_minor_font_is(context, expected_typeface):
    theme = context.prs.slide_masters[0].theme
    actual = theme.font_scheme.minor_font
    assert actual == expected_typeface, (
        "expected '%s', got '%s'" % (expected_typeface, actual)
    )
