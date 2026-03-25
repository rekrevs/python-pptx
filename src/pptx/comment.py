"""Comment proxy object for slide comments."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pptx.oxml.comment import CT_Comment
    from pptx.parts.comment import CommentAuthorsPart


class Comment:
    """Proxy for a single slide comment.

    Provides access to comment text, author name, and creation timestamp.
    """

    def __init__(self, cm_element: CT_Comment, authors_part: CommentAuthorsPart | None):
        self._cm = cm_element
        self._authors_part = authors_part

    @property
    def text(self) -> str:
        """Text content of this comment."""
        return self._cm.text

    @property
    def author(self) -> str:
        """Name of the comment author."""
        if self._authors_part is None:
            return ""
        author = self._authors_part._element.get_author_by_id(self._cm.author_id)
        return author.name if author is not None else ""

    @property
    def timestamp(self) -> dt.datetime | None:
        """Creation time of this comment, or None if not available."""
        return self._cm.timestamp
