"""SmartArt object and related classes."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.shared import PartElementProxy

if TYPE_CHECKING:
    from pptx.oxml.diagram import CT_DiagramDataModel
    from pptx.parts.diagram import DiagramDataPart


class SmartArt(PartElementProxy):
    """A SmartArt diagram object.

    Provides access to the content and structure of a SmartArt diagram.
    """

    part: DiagramDataPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, dataModel: CT_DiagramDataModel, diagram_data_part: DiagramDataPart):
        super().__init__(dataModel, diagram_data_part)
        self._dataModel = dataModel

    @property
    def all_text(self) -> list[str]:
        """Return all text content from the SmartArt diagram.

        Returns a list of strings, one for each content node in the diagram.
        Structural nodes (document root, transitions) are excluded.
        The order of text matches the data model order (typically visual order).

        Example::

            >>> smart_art.all_text
            ['Value Propositions', 'Customer Segments', 'Channels', ...]
        """
        return self._dataModel.all_text

    @property
    def text(self) -> str:
        """Return all text content from the SmartArt diagram as a single string.

        Text from different nodes is separated by newlines.

        Example::

            >>> print(smart_art.text)
            Value Propositions
            Customer Segments
            Channels
            ...
        """
        return "\n".join(self.all_text)
