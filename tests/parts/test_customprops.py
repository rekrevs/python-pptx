# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.customprops` module."""

from __future__ import annotations

import datetime as dt

import pytest

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.oxml.customprops import CT_CustomProperties
from pptx.parts.customprops import CustomPropertiesPart


class DescribeCustomPropertiesPart:
    """Unit-test suite for `pptx.parts.customprops.CustomPropertiesPart` objects."""

    def it_can_construct_a_default_custom_props(self):
        custom_props = CustomPropertiesPart.default(None)  # type: ignore

        assert isinstance(custom_props, CustomPropertiesPart)
        assert custom_props.content_type is CT.OFC_CUSTOM_PROPERTIES
        assert custom_props.partname == "/docProps/custom.xml"
        assert isinstance(custom_props._element, CT_CustomProperties)
        assert len(custom_props) == 0

    def it_can_get_a_string_property(self, custom_properties: CustomPropertiesPart):
        assert custom_properties["MyString"] == "Hello World"

    def it_can_get_an_int_property(self, custom_properties: CustomPropertiesPart):
        assert custom_properties["Count"] == 42

    def it_can_get_a_float_property(self, custom_properties: CustomPropertiesPart):
        assert custom_properties["Ratio"] == 3.14

    def it_can_get_a_bool_property(self, custom_properties: CustomPropertiesPart):
        assert custom_properties["Active"] is True

    def it_can_get_a_false_bool_property(self, custom_properties: CustomPropertiesPart):
        assert custom_properties["Inactive"] is False

    def it_can_get_a_datetime_property(self, custom_properties: CustomPropertiesPart):
        assert custom_properties["Created"] == dt.datetime(2024, 1, 15, 10, 30, 0)

    def it_raises_on_missing_property(self, custom_properties: CustomPropertiesPart):
        with pytest.raises(KeyError):
            custom_properties["NoSuchProperty"]

    def it_can_set_a_string_property(self):
        custom_props = CustomPropertiesPart.default(None)  # type: ignore
        custom_props["Title"] = "My Document"
        assert custom_props["Title"] == "My Document"

    def it_can_set_an_int_property(self):
        custom_props = CustomPropertiesPart.default(None)  # type: ignore
        custom_props["Version"] = 7
        assert custom_props["Version"] == 7

    def it_can_set_a_float_property(self):
        custom_props = CustomPropertiesPart.default(None)  # type: ignore
        custom_props["Score"] = 98.6
        assert custom_props["Score"] == 98.6

    def it_can_set_a_bool_property(self):
        custom_props = CustomPropertiesPart.default(None)  # type: ignore
        custom_props["Enabled"] = True
        assert custom_props["Enabled"] is True

    def it_can_set_a_false_bool_property(self):
        custom_props = CustomPropertiesPart.default(None)  # type: ignore
        custom_props["Disabled"] = False
        assert custom_props["Disabled"] is False

    def it_can_set_a_datetime_property(self):
        custom_props = CustomPropertiesPart.default(None)  # type: ignore
        value = dt.datetime(2024, 6, 15, 8, 0, 0)
        custom_props["Due"] = value
        assert custom_props["Due"] == value

    def it_can_overwrite_an_existing_property(self):
        custom_props = CustomPropertiesPart.default(None)  # type: ignore
        custom_props["Color"] = "red"
        assert custom_props["Color"] == "red"
        custom_props["Color"] = "blue"
        assert custom_props["Color"] == "blue"
        assert len(custom_props) == 1

    def it_can_change_property_type_on_overwrite(self):
        custom_props = CustomPropertiesPart.default(None)  # type: ignore
        custom_props["Value"] = "text"
        assert custom_props["Value"] == "text"
        custom_props["Value"] = 42
        assert custom_props["Value"] == 42

    def it_can_delete_a_property(self, custom_properties: CustomPropertiesPart):
        initial_len = len(custom_properties)
        del custom_properties["MyString"]
        assert len(custom_properties) == initial_len - 1
        assert "MyString" not in custom_properties

    def it_raises_on_delete_missing_property(self, custom_properties: CustomPropertiesPart):
        with pytest.raises(KeyError):
            del custom_properties["NoSuchProperty"]

    def it_supports_contains(self, custom_properties: CustomPropertiesPart):
        assert "MyString" in custom_properties
        assert "NoSuchProperty" not in custom_properties

    def it_supports_len(self, custom_properties: CustomPropertiesPart):
        assert len(custom_properties) == 6

    def it_supports_iteration(self, custom_properties: CustomPropertiesPart):
        names = list(custom_properties)
        assert names == ["MyString", "Count", "Ratio", "Active", "Inactive", "Created"]

    def it_supports_get_with_default(self, custom_properties: CustomPropertiesPart):
        assert custom_properties.get("MyString") == "Hello World"
        assert custom_properties.get("Missing") is None
        assert custom_properties.get("Missing", "fallback") == "fallback"

    # -- fixtures ----------------------------------------------------

    @pytest.fixture
    def custom_properties(self) -> CustomPropertiesPart:
        xml = (
            b'<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
            b'<Properties'
            b' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
            b'custom-properties"'
            b' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/'
            b'docPropsVTypes">\n'
            b'  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"'
            b' pid="2" name="MyString">\n'
            b"    <vt:lpwstr>Hello World</vt:lpwstr>\n"
            b"  </property>\n"
            b'  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"'
            b' pid="3" name="Count">\n'
            b"    <vt:i4>42</vt:i4>\n"
            b"  </property>\n"
            b'  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"'
            b' pid="4" name="Ratio">\n'
            b"    <vt:r8>3.14</vt:r8>\n"
            b"  </property>\n"
            b'  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"'
            b' pid="5" name="Active">\n'
            b"    <vt:bool>true</vt:bool>\n"
            b"  </property>\n"
            b'  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"'
            b' pid="6" name="Inactive">\n'
            b"    <vt:bool>false</vt:bool>\n"
            b"  </property>\n"
            b'  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"'
            b' pid="7" name="Created">\n'
            b"    <vt:filetime>2024-01-15T10:30:00Z</vt:filetime>\n"
            b"  </property>\n"
            b"</Properties>\n"
        )
        return CustomPropertiesPart.load(None, None, None, xml)  # type: ignore
