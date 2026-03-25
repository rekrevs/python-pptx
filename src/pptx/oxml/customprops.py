"""lxml custom element classes for custom properties XML elements."""

from __future__ import annotations

from typing import cast

from lxml import etree

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.xmlchemy import BaseOxmlElement

_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"


class CT_CustomProperties(BaseOxmlElement):
    """``cust:Properties`` element.

    Root element of the custom properties part stored as ``/docProps/custom.xml``.
    """

    _customProperties_tmpl = "<cust:Properties %s/>\n" % nsdecls("cust", "vt")

    @staticmethod
    def new() -> CT_CustomProperties:
        """Return a new ``cust:Properties`` element."""
        return cast(CT_CustomProperties, parse_xml(CT_CustomProperties._customProperties_tmpl))

    @property
    def property_lst(self) -> list[etree._Element]:
        """Return list of ``cust:property`` child elements."""
        return self.findall(qn("cust:property"))

    @property
    def _next_pid(self) -> int:
        """Return the next available property ID (pid starts at 2)."""
        pids = [int(p.get("pid", "1")) for p in self.property_lst]
        return max(pids, default=1) + 1

    def get_property(self, name: str) -> etree._Element | None:
        """Return the ``cust:property`` element with `name`, or None."""
        for prop in self.property_lst:
            if prop.get("name") == name:
                return prop
        return None

    def add_property(self, name: str) -> etree._Element:
        """Add and return a new ``cust:property`` element with `name`."""
        prop = etree.SubElement(self, qn("cust:property"))
        prop.set("fmtid", _FMTID)
        prop.set("pid", str(self._next_pid))
        prop.set("name", name)
        return prop

    def remove_property(self, name: str) -> None:
        """Remove property with `name`. Raise KeyError if not found."""
        prop = self.get_property(name)
        if prop is None:
            raise KeyError(name)
        self.remove(prop)
