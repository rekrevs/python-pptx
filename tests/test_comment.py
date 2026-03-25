"""Unit-test suite for `pptx.comment` module."""

from __future__ import annotations

import datetime as dt

import pytest

from pptx.comment import Comment
from pptx.oxml.comment import CT_CommentAuthorList, CT_CommentList
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.packuri import PackURI
from pptx.parts.comment import CommentAuthorsPart


class DescribeComment:
    """Unit-test suite for `pptx.comment.Comment` objects."""

    def it_can_report_its_text(self):
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=0, idx=1, text="Hello world")
        comment = Comment(cm, None)
        assert comment.text == "Hello world"

    def it_can_report_its_author(self):
        cmAuthorLst = CT_CommentAuthorList.new()
        cmAuthorLst.add_author("Jane Smith")
        authors_part = CommentAuthorsPart(
            PackURI("/ppt/commentAuthors.xml"),
            CT.PML_COMMENT_AUTHORS,
            None,
            cmAuthorLst,
        )
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=0, idx=1, text="Hi")
        comment = Comment(cm, authors_part)
        assert comment.author == "Jane Smith"

    def it_returns_empty_string_when_no_authors_part(self):
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=0, idx=1, text="Hi")
        comment = Comment(cm, None)
        assert comment.author == ""

    def it_returns_empty_string_for_unknown_author_id(self):
        cmAuthorLst = CT_CommentAuthorList.new()
        cmAuthorLst.add_author("Jane Smith")
        authors_part = CommentAuthorsPart(
            PackURI("/ppt/commentAuthors.xml"),
            CT.PML_COMMENT_AUTHORS,
            None,
            cmAuthorLst,
        )
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=99, idx=1, text="Hi")
        comment = Comment(cm, authors_part)
        assert comment.author == ""

    def it_can_report_its_timestamp(self):
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=0, idx=1, text="Hi")
        comment = Comment(cm, None)
        assert isinstance(comment.timestamp, dt.datetime)

    def it_returns_None_for_missing_timestamp(self):
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        xml = '<p:cm %s authorId="0" idx="1"><p:text>Hi</p:text></p:cm>' % nsdecls("p")
        cm = parse_xml(xml)
        comment = Comment(cm, None)
        assert comment.timestamp is None
