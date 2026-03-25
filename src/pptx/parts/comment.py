"""Comment-related parts for slide comments and comment authors."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.comment import CT_CommentAuthorList, CT_CommentList

if TYPE_CHECKING:
    from pptx.package import Package


class SlideCommentsPart(XmlPart):
    """Part containing comments for a single slide.

    Corresponds to package file ``/ppt/comments/comment[N].xml``.
    """

    _element: CT_CommentList

    @classmethod
    def new(cls, package: Package) -> SlideCommentsPart:
        """Return a new |SlideCommentsPart| instance with a unique partname."""
        partname = package.next_partname("/ppt/comments/comment%d.xml")
        return cls(partname, CT.PML_COMMENTS, package, CT_CommentList.new())


class CommentAuthorsPart(XmlPart):
    """Part containing comment authors for the presentation.

    Corresponds to package file ``/ppt/commentAuthors.xml``.
    """

    _element: CT_CommentAuthorList

    @classmethod
    def default(cls, package: Package) -> CommentAuthorsPart:
        """Return a new default |CommentAuthorsPart| instance."""
        return cls(
            PackURI("/ppt/commentAuthors.xml"),
            CT.PML_COMMENT_AUTHORS,
            package,
            CT_CommentAuthorList.new(),
        )
