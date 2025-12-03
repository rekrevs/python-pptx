"""Diagram (SmartArt) part objects."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.package import XmlPart
from pptx.smartart.smartart import SmartArt
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.diagram import CT_DiagramDataModel


class DiagramDataPart(XmlPart):
    """A diagram data part.

    Corresponds to parts having partnames matching ppt/diagrams/data[1-9][0-9]*.xml.
    Contains the data model for a SmartArt diagram (nodes, connections, text).
    """

    _element: CT_DiagramDataModel

    @lazyproperty
    def smart_art(self) -> SmartArt:
        """|SmartArt| object representing the diagram in this part."""
        return SmartArt(self._element, self)
