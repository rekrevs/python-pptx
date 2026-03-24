"""Unit-test suite for `pptx.oxml.shapes.shared` module."""

from __future__ import annotations

import pytest

from ...unitutil.cxml import element


class DescribeCT_NonVisualDrawingProps:
    """Unit-test suite for `pptx.oxml.shapes.shared.CT_NonVisualDrawingProps`."""

    @pytest.mark.parametrize(
        ("cNvPr_cxml", "expected_value"),
        [
            ("p:cNvPr{id=1,name=sp}", None),
            ("p:cNvPr{id=1,name=sp,descr=Alt Text}", "Alt Text"),
        ],
    )
    def it_can_get_the_descr_attr(self, cNvPr_cxml: str, expected_value: str | None):
        cNvPr = element(cNvPr_cxml)
        assert cNvPr.descr == expected_value

    @pytest.mark.parametrize(
        ("cNvPr_cxml", "new_value", "expected_cxml"),
        [
            (
                "p:cNvPr{id=1,name=sp}",
                "New alt text",
                "p:cNvPr{id=1,name=sp,descr=New alt text}",
            ),
            (
                "p:cNvPr{id=1,name=sp,descr=Old}",
                "Updated",
                "p:cNvPr{id=1,name=sp,descr=Updated}",
            ),
            (
                "p:cNvPr{id=1,name=sp,descr=Remove me}",
                None,
                "p:cNvPr{id=1,name=sp}",
            ),
        ],
    )
    def it_can_set_the_descr_attr(
        self, cNvPr_cxml: str, new_value: str | None, expected_cxml: str
    ):
        cNvPr = element(cNvPr_cxml)
        cNvPr.descr = new_value
        assert cNvPr.xml == element(expected_cxml).xml
