"""lxml element classes for slide comments XML."""

from __future__ import annotations

import datetime as dt
from typing import cast

from lxml import etree

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.xmlchemy import BaseOxmlElement


class CT_CommentAuthor(BaseOxmlElement):
    """``p:cmAuthor`` element, a comment author entry."""

    @property
    def id(self) -> int:
        """Integer id of this author."""
        return int(self.get("id", "0"))

    @property
    def name(self) -> str:
        """Name of this author."""
        return self.get("name", "")

    @property
    def initials(self) -> str:
        """Initials of this author."""
        return self.get("initials", "")

    @property
    def lastIdx(self) -> int:
        """Last comment index used by this author."""
        return int(self.get("lastIdx", "0"))

    @lastIdx.setter
    def lastIdx(self, value: int) -> None:
        self.set("lastIdx", str(value))


class CT_CommentAuthorList(BaseOxmlElement):
    """``p:cmAuthorLst`` element, the root element for comment authors part."""

    _tmpl = "<p:cmAuthorLst %s/>\n" % nsdecls("p")

    @staticmethod
    def new() -> CT_CommentAuthorList:
        """Return a new ``CT_CommentAuthorList`` element."""
        return cast(CT_CommentAuthorList, parse_xml(CT_CommentAuthorList._tmpl))

    @property
    def cmAuthor_lst(self) -> list[CT_CommentAuthor]:
        """Return list of ``p:cmAuthor`` child elements."""
        return self.findall(qn("p:cmAuthor"))

    def get_author_by_name(self, name: str) -> CT_CommentAuthor | None:
        """Return the ``p:cmAuthor`` element with matching `name`, or None."""
        for author in self.cmAuthor_lst:
            if author.name == name:
                return author
        return None

    def get_author_by_id(self, author_id: int) -> CT_CommentAuthor | None:
        """Return the ``p:cmAuthor`` element with matching `id`, or None."""
        for author in self.cmAuthor_lst:
            if author.id == author_id:
                return author
        return None

    def _next_id(self) -> int:
        """Return the next available author id."""
        ids = [a.id for a in self.cmAuthor_lst]
        return max(ids, default=-1) + 1

    def add_author(self, name: str, initials: str | None = None) -> CT_CommentAuthor:
        """Add and return a new ``p:cmAuthor`` child element.

        `initials` defaults to the uppercase first letters of each word in `name`.
        """
        if initials is None:
            parts = name.split()
            initials = "".join(p[0].upper() for p in parts if p) or "A"
        new_id = self._next_id()
        author = cast(
            CT_CommentAuthor,
            etree.SubElement(self, qn("p:cmAuthor")),
        )
        author.set("id", str(new_id))
        author.set("name", name)
        author.set("initials", initials)
        author.set("lastIdx", "0")
        author.set("clrIdx", str(new_id))
        return author


class CT_Comment(BaseOxmlElement):
    """``p:cm`` element, a single comment."""

    @property
    def text(self) -> str:
        """Text content of this comment."""
        text_elm = self.find(qn("p:text"))
        if text_elm is None:
            return ""
        return text_elm.text or ""

    @property
    def author_id(self) -> int:
        """Integer author id for this comment."""
        return int(self.get("authorId", "0"))

    @property
    def timestamp(self) -> dt.datetime | None:
        """Creation time of this comment as a datetime, or None."""
        dt_str = self.get("dt")
        if dt_str is None:
            return None
        try:
            # Handle ISO format with or without trailing 'Z'
            return dt.datetime.fromisoformat(dt_str.rstrip("Z"))
        except ValueError:
            return None

    @property
    def idx(self) -> int:
        """Integer index of this comment."""
        return int(self.get("idx", "0"))


class CT_CommentList(BaseOxmlElement):
    """``p:cmLst`` element, the root element for a slide comments part."""

    _tmpl = "<p:cmLst %s/>\n" % nsdecls("p", "a")

    @staticmethod
    def new() -> CT_CommentList:
        """Return a new ``CT_CommentList`` element."""
        return cast(CT_CommentList, parse_xml(CT_CommentList._tmpl))

    @property
    def cm_lst(self) -> list[CT_Comment]:
        """Return list of ``p:cm`` child elements."""
        return self.findall(qn("p:cm"))

    def add_comment(
        self, author_id: int, idx: int, text: str, x: int = 0, y: int = 0
    ) -> CT_Comment:
        """Add and return a new ``p:cm`` child element."""
        cm = cast(CT_Comment, etree.SubElement(self, qn("p:cm")))
        cm.set("authorId", str(author_id))
        cm.set("idx", str(idx))
        now = dt.datetime.now(dt.timezone.utc).replace(tzinfo=None)
        cm.set("dt", now.strftime("%Y-%m-%dT%H:%M:%S.000"))
        pos = etree.SubElement(cm, qn("a:pos"))
        pos.set("x", str(x))
        pos.set("y", str(y))
        text_elm = etree.SubElement(cm, qn("p:text"))
        text_elm.text = text
        return cm
