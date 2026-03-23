"""Track changes mixin: insert, delete."""

from __future__ import annotations

from lxml import etree

from .base import W, W14, _now_iso, _preserve


class TracksMixin:
    """Insert/delete text with tracked changes markup."""

    def insert_text(
        self,
        para_id: str,
        text: str,
        *,
        position: str = "end",
        author: str = "Claude",
    ) -> dict:
        """Insert text with Word track-changes markup (w:ins).

        Args:
            para_id: Target paragraph paraId.
            text: Text to insert.
            position: 'end', 'start', or a substring after which to insert.
            author: Author name for the revision.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        cid = self._next_markup_id(doc)
        now = _now_iso()

        ins = etree.Element(f"{W}ins")
        ins.set(f"{W}id", str(cid))
        ins.set(f"{W}author", author)
        ins.set(f"{W}date", now)

        r = etree.SubElement(ins, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        _preserve(t, text)

        if position == "start":
            ppr = para.find(f"{W}pPr")
            if ppr is not None:
                ppr.addnext(ins)
            else:
                para.insert(0, ins)
        elif position == "end":
            para.append(ins)
        else:
            placed = False
            for run_el in para.findall(f"{W}r"):
                if position in self._text(run_el):
                    run_el.addnext(ins)
                    placed = True
                    break
            if not placed:
                para.append(ins)

        self._mark("word/document.xml")
        return {"change_id": cid, "type": "insertion", "author": author, "date": now}

    def delete_text(
        self,
        para_id: str,
        text: str,
        *,
        author: str = "Claude",
    ) -> dict:
        """Mark text as deleted with Word track-changes markup (w:del).

        Finds the text within runs of the target paragraph and wraps the
        matching portion in <w:del><w:r><w:delText>...</w:delText></w:r></w:del>,
        splitting the run if the match is a substring.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        cid = self._next_markup_id(doc)
        now = _now_iso()

        for run_el in list(para.findall(f"{W}r")):
            t_el = run_el.find(f"{W}t")
            if t_el is None or t_el.text is None:
                continue
            full = t_el.text
            if text not in full:
                continue

            idx = full.index(text)
            rpr = run_el.find(f"{W}rPr")
            rpr_bytes = etree.tostring(rpr) if rpr is not None else None
            parent = run_el.getparent()
            pos = list(parent).index(run_el)
            parent.remove(run_el)

            insert_at = pos

            # Text before
            if idx > 0:
                before = self._make_run(full[:idx], rpr_bytes)
                parent.insert(insert_at, before)
                insert_at += 1

            # Deletion
            del_el = etree.Element(f"{W}del")
            del_el.set(f"{W}id", str(cid))
            del_el.set(f"{W}author", author)
            del_el.set(f"{W}date", now)
            del_run = etree.SubElement(del_el, f"{W}r")
            if rpr_bytes:
                del_run.append(etree.fromstring(rpr_bytes))
            dt = etree.SubElement(del_run, f"{W}delText")
            _preserve(dt, text)
            parent.insert(insert_at, del_el)
            insert_at += 1

            # Text after
            end = idx + len(text)
            if end < len(full):
                after = self._make_run(full[end:], rpr_bytes)
                parent.insert(insert_at, after)

            self._mark("word/document.xml")
            return {"change_id": cid, "type": "deletion", "author": author, "date": now}

        raise ValueError(
            f"Text '{text}' not found in a single run of paragraph '{para_id}'. "
            "If the text spans multiple runs, try searching for a shorter substring."
        )
