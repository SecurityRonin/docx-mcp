"""Comments mixin: get, add, reply."""

from __future__ import annotations

from lxml import etree

from .base import W, W14, W15, _now_iso, _preserve


class CommentsMixin:
    """Comment operations."""

    def get_comments(self) -> list[dict]:
        cm = self._tree("word/comments.xml")
        if cm is None:
            return []
        return [
            {
                "id": int(c.get(f"{W}id", "0")),
                "author": c.get(f"{W}author", ""),
                "date": c.get(f"{W}date", ""),
                "text": self._text(c),
            }
            for c in cm.findall(f"{W}comment")
        ]

    def add_comment(
        self,
        para_id: str,
        text: str,
        *,
        author: str = "Claude",
    ) -> dict:
        """Add a comment anchored to a paragraph."""
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        cm_tree = self._tree("word/comments.xml")
        if cm_tree is None:
            cm_tree = self._create_comments_part()

        comment_id = self._next_comment_id(cm_tree)
        now = _now_iso()
        initials = "".join(w[0].upper() for w in author.split() if w) or "C"

        # Add to comments.xml
        c = etree.SubElement(cm_tree, f"{W}comment")
        c.set(f"{W}id", str(comment_id))
        c.set(f"{W}author", author)
        c.set(f"{W}date", now)
        c.set(f"{W}initials", initials)

        cp = etree.SubElement(c, f"{W}p")
        cp.set(f"{W14}paraId", self._new_para_id())
        cp.set(f"{W14}textId", "77777777")

        # Annotation ref
        ar_run = etree.SubElement(cp, f"{W}r")
        ar_rpr = etree.SubElement(ar_run, f"{W}rPr")
        ar_rs = etree.SubElement(ar_rpr, f"{W}rStyle")
        ar_rs.set(f"{W}val", "CommentReference")
        etree.SubElement(ar_run, f"{W}annotationRef")

        # Comment text
        t_run = etree.SubElement(cp, f"{W}r")
        t_el = etree.SubElement(t_run, f"{W}t")
        _preserve(t_el, text)
        self._mark("word/comments.xml")

        # Add range markers in document.xml
        range_start = etree.Element(f"{W}commentRangeStart")
        range_start.set(f"{W}id", str(comment_id))

        ppr = para.find(f"{W}pPr")
        first_run = para.find(f"{W}r")
        if first_run is not None:
            first_run.addprevious(range_start)
        elif ppr is not None:
            ppr.addnext(range_start)
        else:
            para.insert(0, range_start)

        range_end = etree.SubElement(para, f"{W}commentRangeEnd")
        range_end.set(f"{W}id", str(comment_id))

        ref_run = etree.SubElement(para, f"{W}r")
        ref_rpr = etree.SubElement(ref_run, f"{W}rPr")
        ref_rs = etree.SubElement(ref_rpr, f"{W}rStyle")
        ref_rs.set(f"{W}val", "CommentReference")
        cref = etree.SubElement(ref_run, f"{W}commentReference")
        cref.set(f"{W}id", str(comment_id))
        self._mark("word/document.xml")

        return {"comment_id": comment_id, "para_id": para_id, "author": author, "date": now}

    def reply_to_comment(
        self,
        parent_id: int,
        text: str,
        *,
        author: str = "Claude",
    ) -> dict:
        """Reply to an existing comment."""
        cm_tree = self._require("word/comments.xml")

        # Verify parent exists
        parent_el = None
        for c in cm_tree.findall(f"{W}comment"):
            if c.get(f"{W}id") == str(parent_id):
                parent_el = c
                break
        if parent_el is None:
            raise ValueError(f"Comment {parent_id} not found")

        comment_id = self._next_comment_id(cm_tree)
        now = _now_iso()
        initials = "".join(w[0].upper() for w in author.split() if w) or "C"

        reply = etree.SubElement(cm_tree, f"{W}comment")
        reply.set(f"{W}id", str(comment_id))
        reply.set(f"{W}author", author)
        reply.set(f"{W}date", now)
        reply.set(f"{W}initials", initials)

        rp = etree.SubElement(reply, f"{W}p")
        reply_para_id = self._new_para_id()
        rp.set(f"{W14}paraId", reply_para_id)
        rp.set(f"{W14}textId", "77777777")

        t_run = etree.SubElement(rp, f"{W}r")
        t_el = etree.SubElement(t_run, f"{W}t")
        _preserve(t_el, text)
        self._mark("word/comments.xml")

        # Thread via commentsExtended.xml
        ext = self._tree("word/commentsExtended.xml")
        if ext is not None:
            parent_para = parent_el.find(f"{W}p")
            parent_para_id = parent_para.get(f"{W14}paraId", "") if parent_para is not None else ""
            ce = etree.SubElement(ext, f"{W15}commentEx")
            ce.set(f"{W15}paraId", reply_para_id)
            ce.set(f"{W15}paraIdParent", parent_para_id)
            ce.set(f"{W15}done", "0")
            self._mark("word/commentsExtended.xml")

        return {
            "comment_id": comment_id,
            "parent_id": parent_id,
            "author": author,
            "date": now,
        }
