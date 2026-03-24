"""Bullet formatting object for paragraph bullets."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pptx.dml.color import ColorFormat
    from pptx.oxml.text import CT_TextParagraphProperties


class BulletFormat:
    """Provides access to bullet formatting for a paragraph.

    Accessed via ``paragraph.bullet``. Properties allow reading and writing of
    bullet type, character, auto-numbering scheme, font, size, and color.
    """

    def __init__(self, pPr: CT_TextParagraphProperties):
        self._pPr = pPr

    @property
    def type(self) -> str | None:
        """Bullet type as a string: ``None`` (inherit), ``'none'``, ``'char'``, or ``'auto_num'``.

        ``None`` indicates no explicit bullet setting and the effective value is inherited
        from the paragraph's style hierarchy. Setting to ``None`` removes any explicit
        bullet type, reverting to inheritance.
        """
        if self._pPr.buNone is not None:
            return "none"
        if self._pPr.buChar is not None:
            return "char"
        if self._pPr.buAutoNum is not None:
            return "auto_num"
        return None

    @type.setter
    def type(self, value: str | None) -> None:
        self._pPr._clear_bullet_type()
        if value is None:
            return
        if value == "none":
            self._pPr._add_buNone()
        elif value == "char":
            buChar = self._pPr._add_buChar()
            buChar.char = "\u2022"
        elif value == "auto_num":
            buAutoNum = self._pPr._add_buAutoNum()
            buAutoNum.type = "arabicPeriod"
        else:
            raise ValueError(
                f"bullet type must be None, 'none', 'char', or 'auto_num',"
                f" got {value!r}"
            )

    @property
    def char(self) -> str | None:
        """The bullet character, e.g. ``'\\u2022'`` (bullet).

        Returns ``None`` if this is not a character bullet. Setting this value
        changes the bullet type to ``'char'``.
        """
        buChar = self._pPr.buChar
        return buChar.char if buChar is not None else None

    @char.setter
    def char(self, value: str) -> None:
        self._pPr._clear_bullet_type()
        buChar = self._pPr._add_buChar()
        buChar.char = value

    @property
    def auto_num_type(self) -> str | None:
        """Auto-numbering scheme string, e.g. ``'arabicPeriod'``.

        Returns ``None`` if this is not an auto-numbered bullet. Setting this value
        changes the bullet type to ``'auto_num'``.
        """
        buAutoNum = self._pPr.buAutoNum
        return buAutoNum.type if buAutoNum is not None else None

    @auto_num_type.setter
    def auto_num_type(self, value: str) -> None:
        self._pPr._clear_bullet_type()
        buAutoNum = self._pPr._add_buAutoNum()
        buAutoNum.type = value

    @property
    def font(self) -> str | None:
        """Typeface name for the bullet character.

        Returns ``None`` when the bullet font is inherited from the paragraph's
        default run properties. Setting to ``None`` removes any explicit bullet font.
        """
        buFont = self._pPr.buFont
        return buFont.typeface if buFont is not None else None

    @font.setter
    def font(self, value: str | None) -> None:
        if value is None:
            self._pPr._remove_buFont()
        else:
            buFont = self._pPr.get_or_add_buFont()
            buFont.typeface = value

    @property
    def size(self) -> float | None:
        """Bullet size in points.

        Returns ``None`` when the bullet size is inherited from the paragraph's
        default font size. Setting to ``None`` removes any explicit bullet size.
        """
        buSzPts = self._pPr.buSzPts
        if buSzPts is not None:
            return buSzPts.val / 100.0
        return None

    @size.setter
    def size(self, value: float | None) -> None:
        if value is None:
            self._pPr._remove_buSzPts()
        else:
            buSzPts = self._pPr.get_or_add_buSzPts()
            buSzPts.val = int(value * 100)

    @property
    def color(self) -> ColorFormat:
        """``ColorFormat`` instance providing access to bullet color.

        Accessing this property ensures the ``a:buClr`` element exists. Use
        ``.color.rgb`` or ``.color.theme_color`` to get/set the actual color value.
        """
        from pptx.dml.color import ColorFormat

        buClr = self._pPr.get_or_add_buClr()
        return ColorFormat.from_colorchoice_parent(buClr)
