"""Custom properties part, corresponds to ``/docProps/custom.xml`` part in package."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, Iterator, Union

from lxml import etree

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.customprops import CT_CustomProperties
from pptx.oxml.ns import qn

if TYPE_CHECKING:
    from pptx.package import Package

# --- Type alias for values that can be stored in custom properties ---
CustomPropertyValue = Union[str, int, float, bool, dt.datetime]


class CustomPropertiesPart(XmlPart):
    """Corresponds to part named ``/docProps/custom.xml``.

    Provides dict-like access to custom document properties for this package.
    """

    _element: CT_CustomProperties

    @classmethod
    def default(cls, package: Package) -> CustomPropertiesPart:
        """Return a new empty |CustomPropertiesPart| instance."""
        return cls(
            PackURI("/docProps/custom.xml"),
            CT.OFC_CUSTOM_PROPERTIES,
            package,
            CT_CustomProperties.new(),
        )

    def __getitem__(self, name: str) -> CustomPropertyValue | None:
        """Return value of custom property `name`. Raises KeyError if not found."""
        prop = self._element.get_property(name)
        if prop is None:
            raise KeyError(name)
        return self._read_value(prop)

    def __setitem__(self, name: str, value: CustomPropertyValue) -> None:
        """Set custom property `name` to `value`, with automatic type inference."""
        prop = self._element.get_property(name)
        if prop is not None:
            # --- remove old value child elements ---
            for child in list(prop):
                prop.remove(child)
        else:
            prop = self._element.add_property(name)
        self._write_value(prop, value)

    def __delitem__(self, name: str) -> None:
        """Remove custom property `name`. Raises KeyError if not found."""
        self._element.remove_property(name)

    def __contains__(self, name: object) -> bool:
        """Return True if custom property `name` exists."""
        if not isinstance(name, str):
            return False
        return self._element.get_property(name) is not None

    def __len__(self) -> int:
        """Return the number of custom properties."""
        return len(self._element.property_lst)

    def __iter__(self) -> Iterator[str]:
        """Generate the name of each custom property."""
        for prop in self._element.property_lst:
            yield prop.get("name", "")

    def get(self, name: str, default: CustomPropertyValue | None = None) -> CustomPropertyValue | None:
        """Return value of custom property `name`, or `default` if not found."""
        try:
            return self[name]
        except KeyError:
            return default

    @staticmethod
    def _read_value(prop: etree._Element) -> CustomPropertyValue | None:
        """Read typed value from a ``cust:property`` element."""
        for child in prop:
            tag = child.tag
            text = child.text or ""
            if tag == qn("vt:lpwstr"):
                return text
            elif tag == qn("vt:i4"):
                return int(text)
            elif tag == qn("vt:r8"):
                return float(text)
            elif tag == qn("vt:bool"):
                return text.lower() in ("true", "1")
            elif tag == qn("vt:filetime"):
                # --- strip trailing 'Z' if present and parse ISO format ---
                return dt.datetime.fromisoformat(text.rstrip("Z"))
        return None

    @staticmethod
    def _write_value(prop: etree._Element, value: CustomPropertyValue) -> None:
        """Write typed value to a ``cust:property`` element."""
        # --- bool must be checked before int since isinstance(True, int) is True ---
        if isinstance(value, bool):
            child = etree.SubElement(prop, qn("vt:bool"))
            child.text = "true" if value else "false"
        elif isinstance(value, int):
            child = etree.SubElement(prop, qn("vt:i4"))
            child.text = str(value)
        elif isinstance(value, float):
            child = etree.SubElement(prop, qn("vt:r8"))
            child.text = str(value)
        elif isinstance(value, dt.datetime):
            child = etree.SubElement(prop, qn("vt:filetime"))
            child.text = value.strftime("%Y-%m-%dT%H:%M:%SZ")
        else:
            child = etree.SubElement(prop, qn("vt:lpwstr"))
            child.text = str(value)
