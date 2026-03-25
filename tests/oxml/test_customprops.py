# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.customprops` module."""

from __future__ import annotations

import pytest

from pptx.oxml.customprops import CT_CustomProperties, _FMTID


class DescribeCT_CustomProperties:
    """Unit-test suite for `pptx.oxml.customprops.CT_CustomProperties` objects."""

    def it_can_construct_a_new_element(self):
        props = CT_CustomProperties.new()
        assert isinstance(props, CT_CustomProperties)
        assert len(props.property_lst) == 0

    def it_can_add_a_property(self):
        props = CT_CustomProperties.new()
        prop = props.add_property("TestProp")
        assert prop.get("name") == "TestProp"
        assert prop.get("fmtid") == _FMTID
        assert prop.get("pid") == "2"
        assert len(props.property_lst) == 1

    def it_increments_pid_for_successive_properties(self):
        props = CT_CustomProperties.new()
        prop1 = props.add_property("First")
        prop2 = props.add_property("Second")
        assert prop1.get("pid") == "2"
        assert prop2.get("pid") == "3"

    def it_can_get_a_property_by_name(self):
        props = CT_CustomProperties.new()
        props.add_property("Target")
        prop = props.get_property("Target")
        assert prop is not None
        assert prop.get("name") == "Target"

    def it_returns_None_for_missing_property(self):
        props = CT_CustomProperties.new()
        assert props.get_property("Missing") is None

    def it_can_remove_a_property(self):
        props = CT_CustomProperties.new()
        props.add_property("ToRemove")
        assert len(props.property_lst) == 1
        props.remove_property("ToRemove")
        assert len(props.property_lst) == 0

    def it_raises_on_remove_missing_property(self):
        props = CT_CustomProperties.new()
        with pytest.raises(KeyError):
            props.remove_property("NoSuch")
