"""Graphic Frame shape and related objects.

A graphic frame is a common container for table, chart, smart art, and media
objects.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.base import BaseShape
from pptx.shared import ParentedElementProxy
from pptx.spec import (
    GRAPHIC_DATA_URI_CHART,
    GRAPHIC_DATA_URI_DIAGRAM,
    GRAPHIC_DATA_URI_MODEL3D,
    GRAPHIC_DATA_URI_OLEOBJ,
    GRAPHIC_DATA_URI_SLIDE_ZOOM,
    GRAPHIC_DATA_URI_TABLE,
)
from pptx.table import Table
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.chart.chart import Chart
    from pptx.dml.effect import GlowFormat, ReflectionFormat, ShadowFormat, SoftEdgeFormat
    from pptx.oxml.shapes.graphfrm import CT_GraphicalObjectData, CT_GraphicalObjectFrame
    from pptx.parts.chart import ChartPart
    from pptx.parts.diagram import DiagramDataPart
    from pptx.parts.slide import BaseSlidePart
    from pptx.smartart import SmartArt
    from pptx.types import ProvidesPart


class GraphicFrame(BaseShape):
    """Container shape for table, chart, smart art, and media objects.

    Corresponds to a `p:graphicFrame` element in the shape tree.
    """

    def __init__(self, graphicFrame: CT_GraphicalObjectFrame, parent: ProvidesPart):
        super().__init__(graphicFrame, parent)
        self._graphicFrame = graphicFrame

    @property
    def chart(self) -> Chart:
        """The |Chart| object containing the chart in this graphic frame.

        Raises |ValueError| if this graphic frame does not contain a chart.
        """
        if not self.has_chart:
            raise ValueError("shape does not contain a chart")
        return self.chart_part.chart

    @property
    def chart_part(self) -> ChartPart:
        """The |ChartPart| object containing the chart in this graphic frame."""
        chart_rId = self._graphicFrame.chart_rId
        if chart_rId is None:
            raise ValueError("this graphic frame does not contain a chart")
        return cast("ChartPart", self.part.related_part(chart_rId))

    @property
    def has_chart(self) -> bool:
        """|True| if this graphic frame contains a chart object. |False| otherwise.

        When |True|, the chart object can be accessed using the `.chart` property.
        """
        return self._graphicFrame.graphicData_uri == GRAPHIC_DATA_URI_CHART

    @property
    def has_table(self) -> bool:
        """|True| if this graphic frame contains a table object, |False| otherwise.

        When |True|, the table object can be accessed using the `.table` property.
        """
        return self._graphicFrame.graphicData_uri == GRAPHIC_DATA_URI_TABLE

    @property
    def ole_format(self) -> _OleFormat:
        """_OleFormat object for this graphic-frame shape.

        Raises `ValueError` on a GraphicFrame instance that does not contain an OLE object.

        An shape that contains an OLE object will have `.shape_type` of either
        `EMBEDDED_OLE_OBJECT` or `LINKED_OLE_OBJECT`.
        """
        if not self._graphicFrame.has_oleobj:
            raise ValueError("not an OLE-object shape")
        return _OleFormat(self._graphicFrame.graphicData, self._parent)

    @lazyproperty
    def glow(self) -> GlowFormat:
        """Unconditionally raises |NotImplementedError|.

        Access to the glow effect for graphic-frame objects is content-specific and has not yet
        been implemented.
        """
        raise NotImplementedError("glow property on GraphicFrame not yet supported")

    @lazyproperty
    def reflection(self) -> ReflectionFormat:
        """Unconditionally raises |NotImplementedError|.

        Access to the reflection effect for graphic-frame objects is content-specific and has not
        yet been implemented.
        """
        raise NotImplementedError("reflection property on GraphicFrame not yet supported")

    @lazyproperty
    def shadow(self) -> ShadowFormat:
        """Unconditionally raises |NotImplementedError|.

        Access to the shadow effect for graphic-frame objects is content-specific (i.e. different
        for charts, tables, etc.) and has not yet been implemented.
        """
        raise NotImplementedError("shadow property on GraphicFrame not yet supported")

    @lazyproperty
    def soft_edge(self) -> SoftEdgeFormat:
        """Unconditionally raises |NotImplementedError|.

        Access to the soft edge effect for graphic-frame objects is content-specific and has not
        yet been implemented.
        """
        raise NotImplementedError("soft_edge property on GraphicFrame not yet supported")

    @property
    def has_model_3d(self) -> bool:
        """|True| if this graphic frame contains a 3D model, |False| otherwise.

        A 3D model is an embedded GLB/glTF file that can be rotated and viewed
        from different angles in PowerPoint.
        """
        return self._graphicFrame.graphicData_uri == GRAPHIC_DATA_URI_MODEL3D

    @property
    def has_slide_zoom(self) -> bool:
        """|True| if this graphic frame contains a Slide Zoom, |False| otherwise.

        A Slide Zoom is a special shape that displays a thumbnail of another slide
        and navigates to that slide with a zoom transition when clicked.
        """
        return self._graphicFrame.graphicData_uri == GRAPHIC_DATA_URI_SLIDE_ZOOM

    @property
    def has_smart_art(self) -> bool:
        """|True| if this graphic frame contains a SmartArt diagram, |False| otherwise."""
        return self._graphicFrame.graphicData_uri == GRAPHIC_DATA_URI_DIAGRAM

    @property
    def smart_art(self) -> SmartArt:
        """The |SmartArt| object containing the SmartArt diagram in this graphic frame.

        Raises |ValueError| if this graphic frame does not contain a SmartArt diagram.

        Example::

            >>> shape = slide.shapes[0]
            >>> shape.has_smart_art
            True
            >>> smart_art = shape.smart_art
            >>> smart_art.all_text
            ['Value Propositions', 'Customer Segments', 'Channels', ...]
            >>> print(smart_art.text)
            Value Propositions
            Customer Segments
            Channels
            ...
        """
        if not self.has_smart_art:
            raise ValueError("shape does not contain SmartArt")
        return self._diagram_data_part.smart_art

    @property
    def _diagram_data_part(self) -> DiagramDataPart:
        """The |DiagramDataPart| object containing the SmartArt data for this graphic frame."""
        diagram_data_rId = self._graphicFrame.diagram_data_rId
        if diagram_data_rId is None:
            raise ValueError("this graphic frame does not contain SmartArt")
        return cast("DiagramDataPart", self.part.related_part(diagram_data_rId))

    @property
    def shape_type(self) -> MSO_SHAPE_TYPE:
        """Optional member of `MSO_SHAPE_TYPE` identifying the type of this shape.

        Possible values are `MSO_SHAPE_TYPE.CHART`, `MSO_SHAPE_TYPE.TABLE`,
        `MSO_SHAPE_TYPE.DIAGRAM` (SmartArt), `MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT`,
        `MSO_SHAPE_TYPE.LINKED_OLE_OBJECT`, `MSO_SHAPE_TYPE.SLIDE_ZOOM`,
        `MSO_SHAPE_TYPE.MODEL_3D`.

        This value is `None` when none of these types apply.
        """
        graphicData_uri = self._graphicFrame.graphicData_uri
        if graphicData_uri == GRAPHIC_DATA_URI_CHART:
            return MSO_SHAPE_TYPE.CHART
        elif graphicData_uri == GRAPHIC_DATA_URI_TABLE:
            return MSO_SHAPE_TYPE.TABLE
        elif graphicData_uri == GRAPHIC_DATA_URI_DIAGRAM:
            return MSO_SHAPE_TYPE.DIAGRAM
        elif graphicData_uri == GRAPHIC_DATA_URI_MODEL3D:
            return MSO_SHAPE_TYPE.MODEL_3D
        elif graphicData_uri == GRAPHIC_DATA_URI_OLEOBJ:
            return (
                MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT
                if self._graphicFrame.is_embedded_ole_obj
                else MSO_SHAPE_TYPE.LINKED_OLE_OBJECT
            )
        elif graphicData_uri == GRAPHIC_DATA_URI_SLIDE_ZOOM:
            return MSO_SHAPE_TYPE.SLIDE_ZOOM
        else:
            return None  # pyright: ignore[reportReturnType]

    @property
    def table(self) -> Table:
        """The |Table| object contained in this graphic frame.

        Raises |ValueError| if this graphic frame does not contain a table.
        """
        if not self.has_table:
            raise ValueError("shape does not contain a table")
        tbl = self._graphicFrame.graphic.graphicData.tbl
        return Table(tbl, self)


class _OleFormat(ParentedElementProxy):
    """Provides attributes on an embedded OLE object."""

    part: BaseSlidePart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, graphicData: CT_GraphicalObjectData, parent: ProvidesPart):
        super().__init__(graphicData, parent)
        self._graphicData = graphicData

    @property
    def blob(self) -> bytes | None:
        """Optional bytes of OLE object, suitable for loading or saving as a file.

        This value is `None` if the embedded object does not represent a "file".
        """
        blob_rId = self._graphicData.blob_rId
        if blob_rId is None:
            return None
        return self.part.related_part(blob_rId).blob

    @property
    def prog_id(self) -> str | None:
        """str "progId" attribute of this embedded OLE object.

        The progId is a str like "Excel.Sheet.12" that identifies the "file-type" of the embedded
        object, or perhaps more precisely, the application (aka. "server" in OLE parlance) to be
        used to open this object.
        """
        return self._graphicData.progId

    @property
    def show_as_icon(self) -> bool | None:
        """True when OLE object should appear as an icon (rather than preview)."""
        return self._graphicData.showAsIcon
