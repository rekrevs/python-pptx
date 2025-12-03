"""OXML element classes for diagram (SmartArt) elements."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrMore

if TYPE_CHECKING:
    from lxml.etree import _Element


class CT_DiagramRelIds(BaseOxmlElement):
    """``dgm:relIds`` element.

    Contains relationship IDs pointing to the diagram parts (data, layout, colors, style).
    """

    dm: str | None = OptionalAttribute("r:dm", XsdString)  # pyright: ignore[reportAssignmentType]
    lo: str | None = OptionalAttribute("r:lo", XsdString)  # pyright: ignore[reportAssignmentType]
    qs: str | None = OptionalAttribute("r:qs", XsdString)  # pyright: ignore[reportAssignmentType]
    cs: str | None = OptionalAttribute("r:cs", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_DiagramDataPoint(BaseOxmlElement):
    """``dgm:pt`` element.

    A point (node) in the diagram data model. Contains text and properties.
    """

    @property
    def model_id(self) -> str | None:
        """Return the modelId attribute value."""
        return self.get("modelId")

    @property
    def node_type(self) -> str | None:
        """Return the type attribute value (doc, parTrans, sibTrans, or None for regular)."""
        return self.get("type")

    @property
    def text(self) -> str:
        """Return the text content of this point, extracting from DrawingML paragraph structure.

        The text is stored in dgm:t/a:p/a:r/a:t elements using standard DrawingML format.
        """
        from pptx.oxml.ns import qn

        text_parts: list[str] = []
        # Find the dgm:t element which contains the text
        t_elements = self.findall(qn("dgm:t"))
        for t_elm in t_elements:
            # Find all a:p (paragraph) elements
            for p_elm in t_elm.findall(qn("a:p")):
                # Find all a:r (run) elements within the paragraph
                for r_elm in p_elm.findall(qn("a:r")):
                    # Find the a:t (text) element
                    at_elm = r_elm.find(qn("a:t"))
                    if at_elm is not None and at_elm.text:
                        text_parts.append(at_elm.text)
        return "".join(text_parts)


class CT_DiagramDataPointList(BaseOxmlElement):
    """``dgm:ptLst`` element.

    Contains the list of points (nodes) in the diagram.
    """

    pt_lst: list[CT_DiagramDataPoint]

    pt = ZeroOrMore("dgm:pt")


class CT_DiagramDataConnection(BaseOxmlElement):
    """``dgm:cxn`` element.

    A connection between two points in the diagram (defines hierarchy).
    """

    @property
    def src_id(self) -> str | None:
        """Return the source point modelId."""
        return self.get("srcId")

    @property
    def dest_id(self) -> str | None:
        """Return the destination point modelId."""
        return self.get("destId")

    @property
    def connection_type(self) -> str | None:
        """Return the type of connection (e.g., 'parOf' for parent-of)."""
        return self.get("type")


class CT_DiagramDataConnectionList(BaseOxmlElement):
    """``dgm:cxnLst`` element.

    Contains the list of connections between points.
    """

    cxn_lst: list[CT_DiagramDataConnection]

    cxn = ZeroOrMore("dgm:cxn")


class CT_DiagramDataModel(BaseOxmlElement):
    """``dgm:dataModel`` element.

    The root element of a diagram data file (ppt/diagrams/dataX.xml).
    Contains the point list and connection list that define the diagram content.
    """

    @property
    def ptLst(self) -> CT_DiagramDataPointList | None:
        """Return the point list element."""
        from pptx.oxml.ns import qn

        return self.find(qn("dgm:ptLst"))  # type: ignore

    @property
    def cxnLst(self) -> CT_DiagramDataConnectionList | None:
        """Return the connection list element."""
        from pptx.oxml.ns import qn

        return self.find(qn("dgm:cxnLst"))  # type: ignore

    @property
    def all_text(self) -> list[str]:
        """Return all text content from the diagram, excluding empty strings.

        Returns a list of text strings from all content nodes (excludes doc,
        parTrans, and sibTrans type nodes which are structural).
        """
        ptLst = self.ptLst
        if ptLst is None:
            return []

        texts: list[str] = []
        for pt in ptLst.pt_lst:
            # Skip structural nodes (doc root, transitions)
            node_type = pt.node_type
            if node_type in ("doc", "parTrans", "sibTrans"):
                continue
            text = pt.text.strip()
            if text:
                texts.append(text)
        return texts
