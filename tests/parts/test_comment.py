"""Unit-test suite for `pptx.parts.comment` module."""

from __future__ import annotations

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.packuri import PackURI
from pptx.oxml.comment import CT_CommentAuthorList, CT_CommentList
from pptx.parts.comment import CommentAuthorsPart, SlideCommentsPart


class DescribeSlideCommentsPart:
    """Unit-test suite for `pptx.parts.comment.SlideCommentsPart` objects."""

    def it_can_be_constructed_with_a_new_element(self):
        part = SlideCommentsPart(
            PackURI("/ppt/comments/comment1.xml"),
            CT.PML_COMMENTS,
            None,
            CT_CommentList.new(),
        )
        assert part.partname == "/ppt/comments/comment1.xml"
        assert part.content_type == CT.PML_COMMENTS
        assert isinstance(part._element, CT_CommentList)


class DescribeCommentAuthorsPart:
    """Unit-test suite for `pptx.parts.comment.CommentAuthorsPart` objects."""

    def it_can_create_a_default_instance(self):
        part = CommentAuthorsPart.default(None)
        assert part.partname == "/ppt/commentAuthors.xml"
        assert part.content_type == CT.PML_COMMENT_AUTHORS
        assert isinstance(part._element, CT_CommentAuthorList)
