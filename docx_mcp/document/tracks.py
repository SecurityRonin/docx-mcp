"""Track changes mixin: insert, delete, replace, accept, reject.

New in this version
-------------------
- Accepted-view flattening: w:ins content is visible for anchor search;
  w:del content is invisible.
- Cascading context anchor: full context → context_before → context_after →
  unique-find (document-global uniqueness check).
- Pre-normalisation + whitespace-collapse for fuzzy matching.
- Multi-run spanning: delete/insert across w:r boundaries with rPr inheritance.
- replace_text with collapseDiff minimisation.
"""

from __future__ import annotations

import difflib
import re

from lxml import etree

from .base import W, _now_iso, _preserve

# ── Normalisation ─────────────────────────────────────────────────────────────

# Values that are empty strings are REMOVED (invisible chars — no rendered glyph).
# Values that are non-empty are SUBSTITUTED (typographic chars → ASCII equivalent).
_PRENORM_MAP: dict[str, str] = {
    # Typographic quotes → ASCII equivalents (1:1 substitution)
    "\u2018": "'",
    "\u2019": "'",  # left / right single quotation mark
    "\u201a": "'",
    "\u201b": "'",  # single low-9 / high-reversed-9
    "\u201c": '"',
    "\u201d": '"',  # left / right double quotation mark
    "\u201e": '"',
    "\u201f": '"',  # double low-9 / high-reversed-9
    # Dashes → hyphen (1:1 substitution)
    "\u2014": "-",  # em dash
    "\u2013": "-",  # en dash
    "\u2011": "-",  # non-breaking hyphen
    # Non-breaking / special spaces → regular space (1:1 substitution)
    "\u00a0": " ",  # non-breaking space (NBSP)
    "\u2007": " ",  # figure space
    "\u2009": " ",  # thin space
    "\u200a": " ",  # hair space
    "\u202f": " ",  # narrow no-break space
    # Invisible characters → removed (empty string)
    # These render as nothing; queries must not include them.
    "\u00ad": "",  # soft hyphen (line-break hint)
    "\u200b": "",  # zero-width space
    "\u200c": "",  # zero-width non-joiner
    "\u200d": "",  # zero-width joiner
    "\u2060": "",  # word joiner
    "\ufeff": "",  # BOM / zero-width no-break space
}


def _pre_normalize(text: str) -> str:
    """Replace/remove typographic characters.  For context strings (no idx needed)."""
    return "".join(_PRENORM_MAP.get(ch, ch) for ch in text)


def _prenorm_with_idx(text: str) -> tuple[str, list[int]]:
    """Pre-normalize and return an index map back to the original string.

    ``prenorm_to_orig[i]`` is the index in *text* of the character that
    produced position *i* in the returned pre-normalized string.  Characters
    that map to the empty string are dropped and do not appear in the map.
    """
    result: list[str] = []
    prenorm_to_orig: list[int] = []
    for i, ch in enumerate(text):
        replacement = _PRENORM_MAP.get(ch, ch)
        for c in replacement:  # empty string → zero iterations (char dropped)
            result.append(c)
            prenorm_to_orig.append(i)
    return "".join(result), prenorm_to_orig


def _normalize_ws(text: str) -> tuple[str, list[int]]:
    """Collapse consecutive whitespace runs to a single space.

    Returns ``(norm_text, orig_idx)`` where ``orig_idx[i]`` is the position in
    *text* that corresponds to position *i* in *norm_text*.
    """
    result: list[str] = []
    orig_idx: list[int] = []
    in_space = False
    for i, ch in enumerate(text):
        if ch in " \t\r\n":
            if not in_space:
                result.append(" ")
                orig_idx.append(i)
            in_space = True
        else:
            result.append(ch)
            orig_idx.append(i)
            in_space = False
    return "".join(result), orig_idx


def _norm(text: str, *, ignore_case: bool = False) -> tuple[str, list[int]]:
    """Pre-normalize, whitespace-collapse, and optionally casefold.

    Returns ``(norm_text, orig_idx)`` where ``orig_idx[i]`` is the index in
    the *original* (un-normalized) *text* that corresponds to position *i* in
    *norm_text*.  The composition is safe even when invisible characters are
    removed by pre-normalization.
    """
    prenorm, prenorm_to_orig = _prenorm_with_idx(text)
    ws_norm, prenorm_idx = _normalize_ws(prenorm)
    # Compose: norm position i → prenorm position prenorm_idx[i] → orig position
    orig_idx = [prenorm_to_orig[j] for j in prenorm_idx]
    if ignore_case:
        ws_norm = ws_norm.casefold()
    return ws_norm, orig_idx


# ── Accepted-view flattening ───────────────────────────────────────────────────


class _Slot:
    """One character in a paragraph's accepted view."""

    __slots__ = ("char", "run_el", "rpr_bytes", "idx", "in_ins")

    def __init__(
        self,
        char: str,
        run_el: etree._Element,
        rpr_bytes: bytes | None,
        idx: int,
        in_ins: etree._Element | None,
    ) -> None:
        self.char = char
        self.run_el = run_el
        self.rpr_bytes = rpr_bytes
        self.idx = idx  # absolute position in the full slots list
        self.in_ins = in_ins  # w:ins parent element, or None


def _flatten_para(para: etree._Element) -> list[_Slot]:
    """Build the accepted-view character list for *para*.

    Includes text from ``w:r``, ``w:ins > w:r``, and ``w:hyperlink > w:r``.
    Excludes text from ``w:del`` (invisible in accepted view).
    """
    slots: list[_Slot] = []
    idx = 0
    for child in para:
        if callable(child.tag):
            continue  # lxml Comment / PI node
        if child.tag == f"{W}r":
            idx = _slots_from_run(child, slots, idx, in_ins=None)
        elif child.tag == f"{W}ins":
            for r in child.findall(f"{W}r"):
                idx = _slots_from_run(r, slots, idx, in_ins=child)
        elif child.tag == f"{W}hyperlink":
            # TODO: w:ins > w:r nested inside w:hyperlink (tracked insertion of hyperlink
            # text) is not yet handled — findall only reaches direct w:r children.
            for r in child.findall(f"{W}r"):
                idx = _slots_from_run(r, slots, idx, in_ins=None)
        # w:del, w:pPr, w:bookmarkStart, etc. → skip
    return slots


def _slots_from_run(
    run_el: etree._Element,
    slots: list[_Slot],
    idx: int,
    in_ins: etree._Element | None,
) -> int:
    rpr = run_el.find(f"{W}rPr")
    rpr_bytes = etree.tostring(rpr) if rpr is not None else None
    for t_el in run_el.findall(f"{W}t"):
        for ch in t_el.text or "":
            slots.append(_Slot(ch, run_el, rpr_bytes, idx, in_ins))
            idx += 1
    return idx


# ── Anchor search ─────────────────────────────────────────────────────────────


def _find_in_norm(
    norm_text: str,
    norm_find: str,
    norm_before: str,
    norm_after: str,
) -> list[tuple[int, int]]:
    """Return all ``(start, end)`` positions of *norm_find* in *norm_text* that
    satisfy context constraints."""
    results: list[tuple[int, int]] = []
    pos = 0
    flen = len(norm_find)
    while True:
        idx = norm_text.find(norm_find, pos)
        if idx == -1:
            break
        end = idx + flen
        before_ok = (not norm_before) or norm_before in norm_text[:idx]
        after_ok = (not norm_after) or norm_after in norm_text[end:]
        if before_ok and after_ok:
            results.append((idx, end))
        pos = idx + 1
    return results


def _doc_norm_count(doc: etree._Element, norm_find: str, *, ignore_case: bool = False) -> int:
    """Count how many paragraphs in *doc* contain *norm_find* (in accepted view)."""
    count = 0
    for p in doc.iter(f"{W}p"):
        slots = _flatten_para(p)
        if not slots:
            continue
        nt, _ = _norm("".join(s.char for s in slots), ignore_case=ignore_case)
        if norm_find in nt:
            count += 1
    return count


def _resolve(
    doc: etree._Element,
    para: etree._Element,
    find: str,
    context_before: str,
    context_after: str,
    *,
    ignore_case: bool = False,
) -> tuple[int, int, list[_Slot]]:
    """Locate *find* in *para* using cascading context anchoring.

    Returns ``(slot_s, slot_e, slots)`` where ``slots[slot_s:slot_e]`` are the
    characters that matched.

    Strategies (tried in order):
      1. find + full context
      2. find + context_before only
      3. find + context_after only
      4. find alone — only if text is doc-globally unique

    Raises :class:`ValueError` if the text is not found or is ambiguous.
    """
    slots = _flatten_para(para)
    accepted = "".join(s.char for s in slots)
    norm_text, orig_idx = _norm(accepted, ignore_case=ignore_case)

    norm_find, _ = _norm(find, ignore_case=ignore_case)
    norm_before = _normalize_ws(_pre_normalize(context_before))[0]
    norm_after = _normalize_ws(_pre_normalize(context_after))[0]
    if ignore_case:
        norm_before = norm_before.casefold()
        norm_after = norm_after.casefold()

    if not norm_find:
        raise ValueError("find text is empty")

    match: tuple[int, int] | None = None

    # Strategy 1: full context
    if norm_before or norm_after:
        m = _find_in_norm(norm_text, norm_find, norm_before, norm_after)
        if len(m) == 1:
            match = m[0]

    # Strategy 2: context_before only
    if match is None and norm_before:
        m = _find_in_norm(norm_text, norm_find, norm_before, "")
        if len(m) == 1:
            match = m[0]

    # Strategy 3: context_after only
    if match is None and norm_after:
        m = _find_in_norm(norm_text, norm_find, "", norm_after)
        if len(m) == 1:
            match = m[0]

    # Strategy 4: unique find
    if match is None:
        m = _find_in_norm(norm_text, norm_find, "", "")
        if not m:
            raise ValueError(f"Text {find!r} not found in paragraph")
        if len(m) > 1:
            raise ValueError(
                f"Ambiguous: {find!r} found {len(m)} times in paragraph; "
                "provide context_before or context_after"
            )
        match = m[0]

    ns, ne = match
    slot_s = orig_idx[ns]
    slot_e = orig_idx[ne - 1] + 1 if ne > ns else slot_s + 1
    return slot_s, slot_e, slots


# ── XML helpers ───────────────────────────────────────────────────────────────


def _build_run(
    text: str,
    rpr_bytes: bytes | None,
    *,
    del_text: bool = False,
) -> etree._Element:
    """Build a ``<w:r>`` with optional ``rPr`` and a ``<w:t>`` or ``<w:delText>``."""
    r = etree.Element(f"{W}r")
    if rpr_bytes:
        r.append(etree.fromstring(rpr_bytes))
    tag = f"{W}delText" if del_text else f"{W}t"
    t = etree.SubElement(r, tag)
    _preserve(t, text)
    return r


def _collapse_diff(find: str, replace: str) -> tuple[str, str, str, str]:
    """Return ``(leading, deleted, inserted, trailing)`` for minimal tracking.

    Uses word-level tokenisation so that shared trailing characters that belong
    to changed words (e.g. the 'd' shared by "bold" and "red") are not
    mistakenly included in the trailing common portion.
    """
    if find == replace:
        return find, "", "", ""
    tokens_f = re.split(r"(\s+)", find)
    tokens_r = re.split(r"(\s+)", replace)
    matcher = difflib.SequenceMatcher(None, tokens_f, tokens_r, autojunk=False)
    changed = [
        (tag, i1, i2, j1, j2) for tag, i1, i2, j1, j2 in matcher.get_opcodes() if tag != "equal"
    ]
    if not changed:
        return find, "", "", ""
    fi1, fi2 = changed[0][1], changed[-1][2]
    ri1, ri2 = changed[0][3], changed[-1][4]
    leading = "".join(tokens_f[:fi1])
    del_part = "".join(tokens_f[fi1:fi2])
    ins_part = "".join(tokens_r[ri1:ri2])
    trailing = "".join(tokens_f[fi2:])
    return leading, del_part, ins_part, trailing


# ── Deletion helper ───────────────────────────────────────────────────────────


def _apply_deletion(
    para: etree._Element,
    slot_s: int,
    slot_e: int,
    slots: list[_Slot],
    cid: int,
    author: str,
    now: str,
) -> None:
    """Wrap ``slots[slot_s:slot_e]`` in a ``<w:del>`` element.

    Handles text spanning multiple ``<w:r>`` elements.  Runs that are only
    partially deleted are split; runs that are entirely deleted are absorbed
    into the ``<w:del>``.

    Raises :class:`ValueError` if the target range touches text inside an
    existing ``<w:ins>`` (accept or reject the insertion first).
    """
    del_slots = slots[slot_s:slot_e]
    if not del_slots:
        return

    # Guard: we don't support deleting text that lives inside an existing w:ins
    for s in del_slots:
        if s.in_ins is not None:
            raise ValueError(
                "Cannot delete text inside an existing w:ins; accept or reject the insertion first"
            )

    # ── Group consecutive del_slots by run_el ─────────────────────────────
    run_groups: list[tuple[etree._Element, bytes | None, str]] = []
    prev: etree._Element | None = None
    acc: list[str] = []
    rpr_b: bytes | None = None
    for s in del_slots:
        if s.run_el is not prev:
            if prev is not None:
                run_groups.append((prev, rpr_b, "".join(acc)))
            prev = s.run_el
            rpr_b = s.rpr_bytes
            acc = [s.char]
        else:
            acc.append(s.char)
    if prev is not None:
        run_groups.append((prev, rpr_b, "".join(acc)))

    first_run = run_groups[0][0]
    last_run = run_groups[-1][0]

    # ── Compute before/after text for split runs ───────────────────────────
    first_all = [s for s in slots if s.run_el is first_run]
    before_text = "".join(s.char for s in first_all if s.idx < slot_s)
    before_rpr = first_all[0].rpr_bytes if first_all else None

    last_all = [s for s in slots if s.run_el is last_run]
    after_text = "".join(s.char for s in last_all if s.idx >= slot_e)
    after_rpr = last_all[0].rpr_bytes if last_all else None

    # ── Find insertion point in parent ────────────────────────────────────
    parent = first_run.getparent()
    children = list(parent)
    insert_pos = children.index(first_run)

    # ── Remove all affected run elements ──────────────────────────────────
    seen: set[int] = set()
    for run_el, _, _ in run_groups:
        if id(run_el) not in seen:
            parent.remove(run_el)
            seen.add(id(run_el))

    pos = insert_pos

    # ── Before split (first run, if partial) ──────────────────────────────
    if before_text:
        parent.insert(pos, _build_run(before_text, before_rpr))
        pos += 1

    # ── w:del element ─────────────────────────────────────────────────────
    del_el = etree.Element(f"{W}del")
    del_el.set(f"{W}id", str(cid))
    del_el.set(f"{W}author", author)
    del_el.set(f"{W}date", now)
    for _run_el, rpr_bytes, del_text in run_groups:
        del_el.append(_build_run(del_text, rpr_bytes, del_text=True))
    parent.insert(pos, del_el)
    pos += 1

    # ── After split (last run, if partial) ────────────────────────────────
    if after_text:
        parent.insert(pos, _build_run(after_text, after_rpr))


def _ins_element(
    text: str,
    rpr_bytes: bytes | None,
    *,
    cid: int,
    author: str,
    now: str,
    tracked: bool,
) -> etree._Element:
    """Return ``w:ins>w:r>w:t`` (tracked) or bare ``w:r>w:t`` (untracked)."""
    r = etree.Element(f"{W}r")
    if rpr_bytes:
        r.append(etree.fromstring(rpr_bytes))
    t = etree.SubElement(r, f"{W}t")
    _preserve(t, text)
    if not tracked:
        return r
    ins = etree.Element(f"{W}ins")
    ins.set(f"{W}id", str(cid))
    ins.set(f"{W}author", author)
    ins.set(f"{W}date", now)
    ins.append(r)
    return ins


def _apply_untracked_deletion(
    para: etree._Element,
    slot_s: int,
    slot_e: int,
    slots: list[_Slot],
) -> tuple[etree._Element, int]:
    """Remove ``slots[slot_s:slot_e]`` directly without ``w:del`` markup.

    Returns ``(parent, insert_pos)`` — the parent element and child index where
    replacement text should be spliced in.
    """
    del_slots = slots[slot_s:slot_e]
    if not del_slots:
        return para, slot_s

    run_groups: list[tuple[etree._Element, bytes | None, str]] = []
    prev: etree._Element | None = None
    acc: list[str] = []
    rpr_b: bytes | None = None
    for s in del_slots:
        if s.run_el is not prev:
            if prev is not None:
                run_groups.append((prev, rpr_b, "".join(acc)))
            prev = s.run_el
            rpr_b = s.rpr_bytes
            acc = [s.char]
        else:
            acc.append(s.char)
    if prev is not None:
        run_groups.append((prev, rpr_b, "".join(acc)))

    first_run = run_groups[0][0]
    last_run = run_groups[-1][0]

    first_all = [s for s in slots if s.run_el is first_run]
    before_text = "".join(s.char for s in first_all if s.idx < slot_s)
    before_rpr = first_all[0].rpr_bytes if first_all else None

    last_all = [s for s in slots if s.run_el is last_run]
    after_text = "".join(s.char for s in last_all if s.idx >= slot_e)
    after_rpr = last_all[0].rpr_bytes if last_all else None

    parent = first_run.getparent()
    children = list(parent)
    insert_pos = children.index(first_run)

    seen: set[int] = set()
    for run_el, _, _ in run_groups:
        if id(run_el) not in seen:
            parent.remove(run_el)
            seen.add(id(run_el))

    pos = insert_pos
    if before_text:
        parent.insert(pos, _build_run(before_text, before_rpr))
        pos += 1

    insert_here = pos  # where replacement text should be inserted

    if after_text:
        parent.insert(pos, _build_run(after_text, after_rpr))

    return parent, insert_here


# ── TracksMixin ───────────────────────────────────────────────────────────────


class TracksMixin:
    """Insert / delete / replace / accept / reject with tracked-changes markup."""

    # ── insert_text ─────────────────────────────────────────────────────────

    def insert_text(
        self,
        para_id: str,
        text: str,
        *,
        position: str = "end",
        author: str = "Claude",
        context_before: str = "",
        context_after: str = "",
        ignore_case: bool = False,
        tracked: bool = True,
    ) -> dict:
        """Insert *text* with ``<w:ins>`` tracked-changes markup.

        When *context_before* is provided the insertion is placed immediately
        after the text matched by *context_before* (optionally disambiguated by
        *context_after*).  The inserted run inherits the ``<w:rPr>`` of the run
        at the insertion point.

        When neither context argument is set the legacy *position* semantics
        apply: ``"start"``, ``"end"``, or a substring to insert after.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        cid = self._next_markup_id(doc)
        now = _now_iso()

        # ── Context-anchored insertion ─────────────────────────────────────
        if context_before:
            slot_s, slot_e, slots = _resolve(
                doc, para, context_before, "", context_after, ignore_case=ignore_case
            )
            last = slots[slot_e - 1]

            el = _ins_element(
                text, last.rpr_bytes, cid=cid, author=author, now=now, tracked=tracked
            )

            # Insert after the element that owns the last matched slot
            if last.in_ins is not None:
                anchor = last.in_ins
            else:
                # Check whether last slot is the last char of its run
                run_slots = [s for s in slots if s.run_el is last.run_el]
                if run_slots[-1].idx == last.idx:
                    anchor = last.run_el
                else:
                    # Split the run at last.idx
                    before_chars = "".join(s.char for s in run_slots if s.idx <= last.idx)
                    after_chars = "".join(s.char for s in run_slots if s.idx > last.idx)
                    t_el = last.run_el.find(f"{W}t")
                    if t_el is not None:
                        _preserve(t_el, before_chars)
                    after_run = _build_run(after_chars, last.rpr_bytes)
                    last.run_el.addnext(el)
                    el.addnext(after_run)
                    self._mark("word/document.xml")
                    return {"change_id": cid, "type": "insertion", "author": author, "date": now}

            anchor.addnext(el)
            self._mark("word/document.xml")
            return {"change_id": cid, "type": "insertion", "author": author, "date": now}

        # ── Legacy position-based insertion ───────────────────────────────
        el = _ins_element(text, None, cid=cid, author=author, now=now, tracked=tracked)

        if position == "start":
            ppr = para.find(f"{W}pPr")
            if ppr is not None:
                ppr.addnext(el)
            else:
                para.insert(0, el)
        elif position == "end":
            para.append(el)
        else:
            placed = False
            for run_el in list(para.findall(f"{W}r")):
                t_el = run_el.find(f"{W}t")
                if t_el is None or t_el.text is None:
                    continue
                if position not in t_el.text:
                    continue
                full = t_el.text
                end = full.index(position) + len(position)
                if end < len(full):
                    rpr = run_el.find(f"{W}rPr")
                    rpr_bytes = etree.tostring(rpr) if rpr is not None else None
                    _preserve(t_el, full[:end])
                    after_run = self._make_run(full[end:], rpr_bytes)
                    run_el.addnext(el)
                    el.addnext(after_run)
                else:
                    run_el.addnext(el)
                placed = True
                break

            if not placed:
                for del_el in para.findall(f"{W}del"):
                    del_text = "".join(t.text for t in del_el.iter(f"{W}delText") if t.text)
                    if position in del_text:
                        del_el.addnext(el)
                        placed = True
                        break

            if not placed:
                para.append(el)

        self._mark("word/document.xml")
        return {"change_id": cid, "type": "insertion", "author": author, "date": now}

    # ── delete_text ─────────────────────────────────────────────────────────

    def delete_text(
        self,
        para_id: str,
        text: str,
        *,
        author: str = "Claude",
        context_before: str = "",
        context_after: str = "",
        ignore_case: bool = False,
        tracked: bool = True,
    ) -> dict:
        """Mark *text* as deleted with ``<w:del>`` markup, or remove directly when tracked=False."""
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        cid = self._next_markup_id(doc)
        now = _now_iso()

        slot_s, slot_e, slots = _resolve(
            doc, para, text, context_before, context_after, ignore_case=ignore_case
        )
        if tracked:
            _apply_deletion(para, slot_s, slot_e, slots, cid, author, now)
        else:
            _apply_untracked_deletion(para, slot_s, slot_e, slots)
        self._mark("word/document.xml")
        return {
            "change_id": cid if tracked else None,
            "type": "deletion",
            "author": author,
            "date": now,
        }

    # ── replace_text ────────────────────────────────────────────────────────

    def replace_text(
        self,
        para_id: str,
        *,
        find: str,
        replace: str,
        author: str = "Claude",
        context_before: str = "",
        context_after: str = "",
        ignore_case: bool = False,
        tracked: bool = True,
    ) -> dict:
        """Replace *find* with *replace* using tracked changes (del + ins).

        Uses :func:`_collapse_diff` so that only the actually-changed portion
        becomes tracked markup; leading and trailing common substrings are left
        as plain runs.

        Returns ``{"type": "replacement", ...}``.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        slot_s, slot_e, slots = _resolve(
            doc, para, find, context_before, context_after, ignore_case=ignore_case
        )
        actual_found = "".join(slots[i].char for i in range(slot_s, slot_e))

        leading, del_text, ins_text, _trailing = _collapse_diff(actual_found, replace)

        # Adjust slot range to the minimal changed portion
        real_slot_s = slot_s + len(leading)
        real_slot_e = real_slot_s + len(del_text)

        if not del_text and not ins_text:
            return {
                "del_id": None,
                "ins_id": None,
                "type": "replacement",
                "author": author,
                "date": _now_iso(),
            }  # noqa: E501

        cid = self._next_markup_id(doc)
        now = _now_iso()
        ins_cid: int | None = None

        if tracked:
            if del_text:
                _apply_deletion(para, real_slot_s, real_slot_e, slots, cid, author, now)
                # Refresh slots after deletion mutated the tree
                slots = _flatten_para(para)
                del_els = list(para.iter(f"{W}del"))
                last_del = del_els[-1] if del_els else None

            if ins_text:
                ins_cid = self._next_markup_id(doc)
                ins_el = etree.Element(f"{W}ins")
                ins_el.set(f"{W}id", str(ins_cid))
                ins_el.set(f"{W}author", author)
                ins_el.set(f"{W}date", now)
                rpr_bytes = slots[real_slot_s].rpr_bytes if real_slot_s < len(slots) else None
                r = etree.SubElement(ins_el, f"{W}r")
                if rpr_bytes:
                    r.append(etree.fromstring(rpr_bytes))
                t = etree.SubElement(r, f"{W}t")
                _preserve(t, ins_text)
                if del_text and last_del is not None:
                    last_del.addnext(ins_el)
                else:
                    para.append(ins_el)
        else:
            rpr_bytes = slots[real_slot_s].rpr_bytes if real_slot_s < len(slots) else None
            parent = para
            insert_pos = real_slot_s
            if del_text:
                parent, insert_pos = _apply_untracked_deletion(
                    para, real_slot_s, real_slot_e, slots
                )
            if ins_text:
                parent.insert(insert_pos, _build_run(ins_text, rpr_bytes))

        self._mark("word/document.xml")
        return {
            "del_id": cid if (tracked and del_text) else None,
            "ins_id": ins_cid if (tracked and ins_text) else None,
            "type": "replacement",
            "author": author,
            "date": now,
        }

    # ── Accept / Reject ──────────────────────────────────────────────────────

    def _matches_author(self, el: etree._Element, author: str | None) -> bool:
        if author is None:
            return True
        return el.get(f"{W}author") == author

    def accept_changes(self, *, author: str | None = None) -> dict:
        """Accept tracked changes: keep insertions, remove deletions."""
        doc = self._require("word/document.xml")
        count = 0

        for ins in list(doc.iter(f"{W}ins")):
            if not self._matches_author(ins, author):
                continue
            parent = ins.getparent()
            idx = list(parent).index(ins)
            for child in list(ins):
                ins.remove(child)
                parent.insert(idx, child)
                idx += 1
            parent.remove(ins)
            count += 1

        for del_el in list(doc.iter(f"{W}del")):
            if not self._matches_author(del_el, author):
                continue
            del_el.getparent().remove(del_el)
            count += 1

        if count > 0:
            self._mark("word/document.xml")
        return {"accepted": count, "scope": "by_author" if author else "all"}

    def reject_changes(self, *, author: str | None = None) -> dict:
        """Reject tracked changes: remove insertions, restore deletions."""
        doc = self._require("word/document.xml")
        count = 0

        for ins in list(doc.iter(f"{W}ins")):
            if not self._matches_author(ins, author):
                continue
            ins.getparent().remove(ins)
            count += 1

        for del_el in list(doc.iter(f"{W}del")):
            if not self._matches_author(del_el, author):
                continue
            parent = del_el.getparent()
            idx = list(parent).index(del_el)
            for child in list(del_el):
                for dt in child.findall(f"{W}delText"):
                    dt.tag = f"{W}t"
                del_el.remove(child)
                parent.insert(idx, child)
                idx += 1
            parent.remove(del_el)
            count += 1

        if count > 0:
            self._mark("word/document.xml")
        return {"rejected": count, "scope": "by_author" if author else "all"}

    def get_body_text(self) -> dict:
        """Return accepted-view full text as {body, footnotes}."""
        doc = self._require("word/document.xml")
        body_lines: list[str] = []
        for para in doc.iter(f"{W}p"):
            slots = _flatten_para(para)
            body_lines.append("".join(s.char for s in slots))

        footnote_lines: list[str] = []
        try:
            fn_doc = self._require("word/footnotes.xml")
            for fn in fn_doc.findall(f"{W}footnote"):
                if fn.get(f"{W}id") in ("-1", "0"):
                    continue
                for para in fn.iter(f"{W}p"):
                    slots = _flatten_para(para)
                    footnote_lines.append("".join(s.char for s in slots))
        except RuntimeError:
            pass

        return {
            "body": "\n".join(line for line in body_lines if line),
            "footnotes": "\n".join(line for line in footnote_lines if line),
        }
