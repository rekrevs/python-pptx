"""Visual effects on a shape such as shadow, glow, and reflection."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.dml.color import ColorFormat

if TYPE_CHECKING:
    from pptx.oxml.dml.effect import (
        CT_GlowEffect,
        CT_InnerShadowEffect,
        CT_OuterShadowEffect,
        CT_ReflectionEffect,
        CT_SoftEdgesEffect,
    )
    from pptx.oxml.shapes.groupshape import CT_GroupShapeProperties
    from pptx.oxml.shapes.shared import CT_ShapeProperties
    from pptx.util import Length


class ShadowFormat(object):
    """Provides access to shadow effect on a shape.

    Wraps the `spPr` or `grpSpPr` element and navigates to the shadow
    child within `a:effectLst`.
    """

    def __init__(self, spPr: CT_ShapeProperties | CT_GroupShapeProperties):
        # ---spPr may also be a grpSpPr; both have a:effectLst child---
        self._element = spPr

    @property
    def angle(self) -> float | None:
        """Shadow direction in degrees (float), or None if no shadow element.

        Read/write. The angle is measured clockwise from the top of the shape.
        For example, 90.0 is a shadow cast to the right.
        """
        shadow = self._shadow_elm
        if shadow is None:
            return None
        dir_val = shadow.dir
        return dir_val if dir_val is not None else 0.0

    @angle.setter
    def angle(self, value: float) -> None:
        shadow = self._shadow_elm
        if shadow is None:
            raise ValueError("no shadow element; set inherit to False first")
        shadow.dir = value

    @property
    def blur_radius(self) -> Length | None:
        """Shadow blur radius as an |Emu| value, or None if no shadow element.

        Read/write. Controls how blurry the shadow edge appears.
        """
        shadow = self._shadow_elm
        if shadow is None:
            return None
        blur = shadow.blurRad
        from pptx.util import Emu

        return blur if blur is not None else Emu(0)

    @blur_radius.setter
    def blur_radius(self, value: int) -> None:
        shadow = self._shadow_elm
        if shadow is None:
            raise ValueError("no shadow element; set inherit to False first")
        shadow.blurRad = value

    @property
    def color(self) -> ColorFormat:
        """Return |ColorFormat| providing access to shadow color.

        Raises |ValueError| if no shadow element is present.
        """
        shadow = self._shadow_elm
        if shadow is None:
            raise ValueError("no shadow element; set inherit to False first")
        return ColorFormat.from_colorchoice_parent(shadow)

    @property
    def distance(self) -> Length | None:
        """Shadow offset distance as an |Emu| value, or None if no shadow element.

        Read/write. Controls how far the shadow is offset from the shape.
        """
        shadow = self._shadow_elm
        if shadow is None:
            return None
        dist = shadow.dist
        from pptx.util import Emu

        return dist if dist is not None else Emu(0)

    @distance.setter
    def distance(self, value: int) -> None:
        shadow = self._shadow_elm
        if shadow is None:
            raise ValueError("no shadow element; set inherit to False first")
        shadow.dist = value

    @property
    def inherit(self):
        """True if shape inherits shadow settings.

        Read/write. An explicitly-defined shadow setting on a shape causes
        this property to return |False|. A shape with no explicitly-defined
        shadow setting inherits its shadow settings from the style hierarchy
        (and so returns |True|).

        Assigning |True| causes any explicitly-defined shadow setting to be
        removed and inheritance is restored. Note this has the side-effect of
        removing **all** explicitly-defined effects, such as glow and
        reflection, and restoring inheritance for all effects on the shape.
        Assigning |False| causes the inheritance link to be broken and **no**
        effects to appear on the shape.
        """
        if self._element.effectLst is None:
            return True
        return False

    @inherit.setter
    def inherit(self, value):
        inherit = bool(value)
        if inherit:
            # ---remove any explicitly-defined effects
            self._element._remove_effectLst()
        else:
            # ---ensure at least the effectLst element is present
            self._element.get_or_add_effectLst()

    @property
    def shadow_type(self) -> str | None:
        """Return the type of shadow as a string: 'outer', 'inner', or None.

        Returns None if no effectLst is present (inheriting) or if the
        effectLst contains no shadow element.
        """
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        if effectLst.outerShdw is not None:
            return "outer"
        if effectLst.innerShdw is not None:
            return "inner"
        return None

    @property
    def _shadow_elm(
        self,
    ) -> CT_OuterShadowEffect | CT_InnerShadowEffect | None:
        """Return the outerShdw or innerShdw element, or None."""
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        outerShdw = effectLst.outerShdw
        if outerShdw is not None:
            return outerShdw
        return effectLst.innerShdw


class GlowFormat(object):
    """Provides access to glow effect on a shape.

    Wraps the `spPr` or `grpSpPr` element and navigates to the glow
    child within `a:effectLst`.
    """

    def __init__(self, spPr: CT_ShapeProperties | CT_GroupShapeProperties):
        self._element = spPr

    @property
    def color(self) -> ColorFormat:
        """Return |ColorFormat| providing access to glow color.

        Raises |ValueError| if no glow element is present.
        """
        glow = self._glow_elm
        if glow is None:
            raise ValueError("no glow element present")
        return ColorFormat.from_colorchoice_parent(glow)

    @property
    def radius(self) -> Length | None:
        """Glow extent radius as an |Emu| value, or None if no glow element.

        Read/write. Controls how far the glow extends from the shape edge.
        """
        glow = self._glow_elm
        if glow is None:
            return None
        rad = glow.rad
        from pptx.util import Emu

        return rad if rad is not None else Emu(0)

    @radius.setter
    def radius(self, value: int) -> None:
        glow = self._glow_elm
        if glow is None:
            raise ValueError("no glow element present")
        glow.rad = value

    @property
    def _glow_elm(self) -> CT_GlowEffect | None:
        """Return the `a:glow` element, or None if not present."""
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        return effectLst.glow


class ReflectionFormat(object):
    """Provides access to reflection effect on a shape.

    Wraps the `spPr` or `grpSpPr` element and navigates to the reflection
    child within `a:effectLst`.
    """

    def __init__(self, spPr: CT_ShapeProperties | CT_GroupShapeProperties):
        self._element = spPr

    @property
    def blur_radius(self) -> Length | None:
        """Reflection blur radius as an |Emu| value, or None if no reflection element.

        Read-only. Controls how blurry the reflection edge appears.
        """
        refl = self._reflection_elm
        if refl is None:
            return None
        blur = refl.blurRad
        from pptx.util import Emu

        return blur if blur is not None else Emu(0)

    @property
    def direction(self) -> float | None:
        """Reflection direction in degrees (float), or None if no reflection element.

        Read-only.
        """
        refl = self._reflection_elm
        if refl is None:
            return None
        dir_val = refl.dir
        return dir_val if dir_val is not None else 0.0

    @property
    def distance(self) -> Length | None:
        """Reflection offset distance as an |Emu| value, or None if no reflection element.

        Read-only.
        """
        refl = self._reflection_elm
        if refl is None:
            return None
        dist = refl.dist
        from pptx.util import Emu

        return dist if dist is not None else Emu(0)

    @property
    def end_opacity(self) -> float | None:
        """Reflection end opacity as a float between 0.0 and 1.0, or None.

        Read-only. Controls the opacity at the far edge of the reflection.
        """
        refl = self._reflection_elm
        if refl is None:
            return None
        endA = refl.endA
        return endA if endA is not None else 0.0

    @property
    def start_opacity(self) -> float | None:
        """Reflection start opacity as a float between 0.0 and 1.0, or None.

        Read-only. Controls the opacity at the near edge of the reflection.
        """
        refl = self._reflection_elm
        if refl is None:
            return None
        stA = refl.stA
        return stA if stA is not None else 0.0

    @property
    def _reflection_elm(self) -> CT_ReflectionEffect | None:
        """Return the `a:reflection` element, or None if not present."""
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        return effectLst.reflection


class SoftEdgeFormat(object):
    """Provides access to soft edge effect on a shape.

    Wraps the `spPr` or `grpSpPr` element and navigates to the soft edge
    child within `a:effectLst`.
    """

    def __init__(self, spPr: CT_ShapeProperties | CT_GroupShapeProperties):
        self._element = spPr

    @property
    def radius(self) -> Length | None:
        """Soft edge radius as an |Emu| value, or None if no soft edge element.

        Read/write. Controls how far the soft edge extends inward from the
        shape boundary.
        """
        softEdge = self._softEdge_elm
        if softEdge is None:
            return None
        return softEdge.rad

    @radius.setter
    def radius(self, value: int) -> None:
        softEdge = self._softEdge_elm
        if softEdge is None:
            raise ValueError("no soft edge element present")
        softEdge.rad = value

    @property
    def _softEdge_elm(self) -> CT_SoftEdgesEffect | None:
        """Return the `a:softEdge` element, or None if not present."""
        effectLst = self._element.effectLst
        if effectLst is None:
            return None
        return effectLst.softEdge
