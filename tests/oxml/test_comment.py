"""Unit-test suite for `pptx.oxml.comment` module."""

from __future__ import annotations

import datetime as dt

import pytest

from pptx.oxml.comment import CT_Comment, CT_CommentAuthorList, CT_CommentList


class DescribeCT_CommentAuthorList:
    """Unit-test suite for `pptx.oxml.comment.CT_CommentAuthorList` objects."""

    def it_can_create_a_new_empty_element(self):
        cmAuthorLst = CT_CommentAuthorList.new()
        assert cmAuthorLst.tag.endswith("}cmAuthorLst")
        assert cmAuthorLst.cmAuthor_lst == []

    def it_can_add_an_author(self):
        cmAuthorLst = CT_CommentAuthorList.new()
        author = cmAuthorLst.add_author("Jane Smith")
        assert author.name == "Jane Smith"
        assert author.initials == "JS"
        assert author.id == 0
        assert author.lastIdx == 0

    def it_generates_initials_from_name(self):
        cmAuthorLst = CT_CommentAuthorList.new()
        author = cmAuthorLst.add_author("John Q. Doe")
        assert author.initials == "JQD"

    def it_can_use_custom_initials(self):
        cmAuthorLst = CT_CommentAuthorList.new()
        author = cmAuthorLst.add_author("Jane Smith", initials="JAS")
        assert author.initials == "JAS"

    def it_assigns_sequential_ids(self):
        cmAuthorLst = CT_CommentAuthorList.new()
        a1 = cmAuthorLst.add_author("Author One")
        a2 = cmAuthorLst.add_author("Author Two")
        assert a1.id == 0
        assert a2.id == 1

    def it_can_find_an_author_by_name(self):
        cmAuthorLst = CT_CommentAuthorList.new()
        cmAuthorLst.add_author("Jane Smith")
        cmAuthorLst.add_author("John Doe")
        author = cmAuthorLst.get_author_by_name("John Doe")
        assert author is not None
        assert author.name == "John Doe"

    def it_returns_None_for_missing_author_by_name(self):
        cmAuthorLst = CT_CommentAuthorList.new()
        assert cmAuthorLst.get_author_by_name("Nobody") is None

    def it_can_find_an_author_by_id(self):
        cmAuthorLst = CT_CommentAuthorList.new()
        cmAuthorLst.add_author("Jane Smith")
        cmAuthorLst.add_author("John Doe")
        author = cmAuthorLst.get_author_by_id(1)
        assert author is not None
        assert author.name == "John Doe"

    def it_returns_None_for_missing_author_by_id(self):
        cmAuthorLst = CT_CommentAuthorList.new()
        assert cmAuthorLst.get_author_by_id(99) is None

    def it_can_update_lastIdx(self):
        cmAuthorLst = CT_CommentAuthorList.new()
        author = cmAuthorLst.add_author("Jane Smith")
        assert author.lastIdx == 0
        author.lastIdx = 5
        assert author.lastIdx == 5


class DescribeCT_CommentList:
    """Unit-test suite for `pptx.oxml.comment.CT_CommentList` objects."""

    def it_can_create_a_new_empty_element(self):
        cmLst = CT_CommentList.new()
        assert cmLst.tag.endswith("}cmLst")
        assert cmLst.cm_lst == []

    def it_can_add_a_comment(self):
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=0, idx=1, text="Hello")
        assert cm.text == "Hello"
        assert cm.author_id == 0
        assert cm.idx == 1
        assert cm.timestamp is not None

    def it_can_add_multiple_comments(self):
        cmLst = CT_CommentList.new()
        cmLst.add_comment(author_id=0, idx=1, text="First")
        cmLst.add_comment(author_id=1, idx=1, text="Second")
        assert len(cmLst.cm_lst) == 2
        assert cmLst.cm_lst[0].text == "First"
        assert cmLst.cm_lst[1].text == "Second"


class DescribeCT_Comment:
    """Unit-test suite for `pptx.oxml.comment.CT_Comment` objects."""

    def it_can_report_its_text(self):
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=0, idx=1, text="Test comment")
        assert cm.text == "Test comment"

    def it_returns_empty_string_for_missing_text(self):
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        xml = '<p:cm %s authorId="0" idx="1"/>' % nsdecls("p")
        cm = parse_xml(xml)
        assert cm.text == ""

    def it_can_report_its_author_id(self):
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=42, idx=1, text="Hello")
        assert cm.author_id == 42

    def it_can_report_its_timestamp(self):
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=0, idx=1, text="Hello")
        ts = cm.timestamp
        assert isinstance(ts, dt.datetime)
        # --- timestamp should be close to now ---
        now = dt.datetime.now(dt.timezone.utc).replace(tzinfo=None)
        assert abs((now - ts).total_seconds()) < 5

    def it_returns_None_for_missing_timestamp(self):
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        xml = '<p:cm %s authorId="0" idx="1"><p:text>Hi</p:text></p:cm>' % nsdecls("p")
        cm = parse_xml(xml)
        assert cm.timestamp is None

    def it_can_parse_iso_timestamp_with_Z(self):
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        xml = (
            '<p:cm %s authorId="0" idx="1" dt="2025-01-15T10:30:00.000Z">'
            "<p:text>Hi</p:text></p:cm>" % nsdecls("p")
        )
        cm = parse_xml(xml)
        ts = cm.timestamp
        assert ts == dt.datetime(2025, 1, 15, 10, 30, 0)

    def it_can_report_its_idx(self):
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=0, idx=7, text="Hello")
        assert cm.idx == 7
