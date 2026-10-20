"""PII scrubbing mixin: detect and permanently redact personal information."""

from __future__ import annotations

import contextlib
import copy
import os
import zipfile
from pathlib import Path

from lxml import etree

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# DrawingML namespaces required for the redaction rectangle
_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

# Fixed redaction rectangle dimensions: 1 inch wide × 0.25 inch tall (EMUs)
_RECT_CX = "914400"
_RECT_CY = "228600"


def _next_drawing_id(doc_tree: etree._Element, _used: set[int] | None = None) -> int:
    """Return an ID not used by any existing wp:docPr in doc_tree."""
    _WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    used = set()
    for el in doc_tree.iter(f"{{{_WP}}}docPr"):
        with contextlib.suppress(ValueError, TypeError):
            used.add(int(el.get("id", 0)))
    if _used:
        used.update(_used)
    n = 1
    while n in used:
        n += 1
    return n


# Lazy singleton — loading AnalyzerEngine instantiates spaCy (expensive once)
_analyzer = None  # lazy-loaded on first scrub_pii call

_MODEL_NAME = "en_core_web_lg"


def _get_analyzer():
    """Lazy-load Presidio AnalyzerEngine with en_core_web_lg.

    On first call, downloads en_core_web_lg (~560MB) if not already installed.
    Download is one-time; cached in the spaCy data directory.
    """
    global _analyzer
    if _analyzer is not None:
        return _analyzer

    import spacy
    from presidio_analyzer import AnalyzerEngine
    from presidio_analyzer.nlp_engine import NlpEngineProvider

    # Auto-download the model on first use if not installed
    try:
        spacy.load(_MODEL_NAME)
    except OSError:
        import sys

        print(
            f"[docx-mcp] Downloading {_MODEL_NAME} NER model (~560MB, one-time)...",
            file=sys.stderr,
        )
        from spacy.cli import download as _spacy_download

        _spacy_download(_MODEL_NAME)

    nlp_config = {
        "nlp_engine_name": "spacy",
        "models": [{"lang_code": "en", "model_name": _MODEL_NAME}],
    }
    provider = NlpEngineProvider(nlp_configuration=nlp_config)
    nlp_engine = provider.create_engine()
    _analyzer = AnalyzerEngine(
        nlp_engine=nlp_engine,
        supported_languages=["en"],
    )
    return _analyzer


def _wt(tag: str) -> str:
    return f"{{{_W}}}{tag}"


# ── Run-level helpers ────────────────────────────────────────────────────────


def _make_run(text: str, rPr: etree._Element | None) -> etree._Element:
    """Create a plain w:r with optional rPr (deep-copied)."""
    run = etree.Element(_wt("r"))
    if rPr is not None:
        run.append(copy.deepcopy(rPr))
    t_el = etree.SubElement(run, _wt("t"))
    t_el.set(_XML_SPACE, "preserve")
    t_el.text = text
    return run


def _make_redaction_drawing(doc_tree: etree._Element, _used_ids: set[int]) -> etree._Element:
    """Create a w:r containing a solid black DrawingML inline rectangle.

    The rectangle is a fixed size (1 inch × 0.25 inch) regardless of the
    length of the redacted text, so no information leaks via character count.
    No w:t element is produced — the original text is deleted entirely.
    """
    drawing_id = _next_drawing_id(doc_tree, _used_ids)
    _used_ids.add(drawing_id)

    run = etree.Element(_wt("r"))
    drawing = etree.SubElement(run, _wt("drawing"))

    inline = etree.SubElement(
        drawing,
        f"{{{_WP_NS}}}inline",
        attrib={"distT": "0", "distB": "0", "distL": "0", "distR": "0"},
    )

    etree.SubElement(
        inline,
        f"{{{_WP_NS}}}extent",
        attrib={"cx": _RECT_CX, "cy": _RECT_CY},
    )
    etree.SubElement(
        inline,
        f"{{{_WP_NS}}}effectExtent",
        attrib={"l": "0", "t": "0", "r": "0", "b": "0"},
    )
    etree.SubElement(
        inline,
        f"{{{_WP_NS}}}docPr",
        attrib={"id": str(drawing_id), "name": "Redacted"},
    )
    etree.SubElement(inline, f"{{{_WP_NS}}}cNvGraphicFramePr")

    graphic = etree.SubElement(inline, f"{{{_A_NS}}}graphic")
    graphic_data = etree.SubElement(
        graphic,
        f"{{{_A_NS}}}graphicData",
        attrib={"uri": _WPS_NS},
    )

    wsp = etree.SubElement(graphic_data, f"{{{_WPS_NS}}}wsp")

    cNvSpPr = etree.SubElement(wsp, f"{{{_WPS_NS}}}cNvSpPr")
    sp_locks = etree.SubElement(cNvSpPr, f"{{{_A_NS}}}spLocks")
    sp_locks.set("noChangeArrowheads", "1")

    spPr = etree.SubElement(wsp, f"{{{_WPS_NS}}}spPr")

    xfrm = etree.SubElement(spPr, f"{{{_A_NS}}}xfrm")
    etree.SubElement(xfrm, f"{{{_A_NS}}}off", attrib={"x": "0", "y": "0"})
    etree.SubElement(xfrm, f"{{{_A_NS}}}ext", attrib={"cx": _RECT_CX, "cy": _RECT_CY})

    prstGeom = etree.SubElement(spPr, f"{{{_A_NS}}}prstGeom", attrib={"prst": "rect"})
    etree.SubElement(prstGeom, f"{{{_A_NS}}}avLst")

    solidFill = etree.SubElement(spPr, f"{{{_A_NS}}}solidFill")
    srgbClr = etree.SubElement(solidFill, f"{{{_A_NS}}}srgbClr")
    srgbClr.set("val", "000000")

    ln = etree.SubElement(spPr, f"{{{_A_NS}}}ln")
    etree.SubElement(ln, f"{{{_A_NS}}}noFill")

    etree.SubElement(wsp, f"{{{_WPS_NS}}}bodyPr")

    return run


# ── Paragraph redaction ──────────────────────────────────────────────────────


def _build_run_char_map(
    para: etree._Element,
) -> tuple[str, list[tuple[int, int, etree._Element]]]:
    """
    Return (full_text, run_ranges) where:
      full_text  — concatenated w:t text (excluding w:del content)
      run_ranges — [(global_start, global_end, run_el), ...]

    Each run's character range in full_text is [global_start, global_end).
    Only w:r elements NOT inside w:del are included.
    """
    full_text = ""
    run_ranges: list[tuple[int, int, etree._Element]] = []

    for run in para.iter(_wt("r")):
        # Skip deleted runs
        parent = run.getparent()
        if parent is not None and parent.tag == _wt("del"):
            continue

        run_text = "".join(t.text or "" for t in run.findall(_wt("t")))
        if run_text:
            start = len(full_text)
            full_text += run_text
            run_ranges.append((start, len(full_text), run))

    return full_text, run_ranges


def _redact_span(
    run_ranges: list[tuple[int, int, etree._Element]],
    span_start: int,
    span_end: int,
    doc_tree: etree._Element,
    _used_ids: set[int],
) -> None:
    """
    Replace the character range [span_start, span_end) in the paragraph
    with a DrawingML black rectangle.  Modifies the XML tree in-place.

    Handles spans that cross multiple runs by splitting each involved run.
    One DrawingML rectangle is inserted for the entire span (replacing the
    first involved run); subsequent involved runs in the same span are simply
    removed with their text.
    """
    involved = [(s, e, r) for s, e, r in run_ranges if e > span_start and s < span_end]
    if not involved:
        return

    drawing_inserted = False

    for run_start, run_end, run in involved:
        run_text = "".join(t.text or "" for t in run.findall(_wt("t")))
        rPr = run.find(_wt("rPr"))

        # Portion of this run that falls within the PII span
        pii_in_run_start = max(span_start - run_start, 0)
        pii_in_run_end = min(span_end - run_start, run_end - run_start)

        before = run_text[:pii_in_run_start]
        after = run_text[pii_in_run_end:]

        # Insert replacement runs immediately before original, then remove it
        if before:
            run.addprevious(_make_run(before, rPr))
        if not drawing_inserted:
            # Insert exactly one black rectangle for the entire redacted span
            run.addprevious(_make_redaction_drawing(doc_tree, _used_ids))
            drawing_inserted = True
        if after:
            run.addprevious(_make_run(after, rPr))

        run.getparent().remove(run)


def _find_all_occurrences(text: str, needle: str) -> list[tuple[int, int]]:
    """Return [(start, end), ...] for all case-insensitive occurrences of needle."""
    needle_lower = needle.casefold()
    text_lower = text.casefold()
    spans = []
    start = 0
    while True:
        pos = text_lower.find(needle_lower, start)
        if pos == -1:
            break
        spans.append((pos, pos + len(needle)))
        start = pos + 1
    return spans


def _merge_overlapping(spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """Merge overlapping or adjacent spans, return sorted non-overlapping list."""
    if not spans:
        return []
    sorted_spans = sorted(spans)
    merged = [sorted_spans[0]]
    for start, end in sorted_spans[1:]:
        if start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


# ── Main mixin ───────────────────────────────────────────────────────────────


class PiiMixin:
    """Detect and permanently redact PII using Presidio + spaCy NER."""

    def scrub_pii(
        self,
        output_path: str,
        *,
        entities: list[str] | None = None,
        confidence_threshold: float = 0.35,
        dry_run: bool = False,
        also_sanitize_metadata: bool = True,
        redact_authors_as: str = "REDACTED",
    ) -> dict:
        """Detect and permanently redact PII from the open document.

        Redaction is true XML redaction: the original text is deleted and
        replaced with a solid black DrawingML rectangle.  No text content
        remains in the output OOXML for the redacted spans.

        Args:
            output_path: Destination path for the scrubbed DOCX.
                         Required when dry_run=False; ignored when dry_run=True.
            entities:    Presidio entity type filter, e.g. ["PERSON","EMAIL_ADDRESS"].
                         None / empty list = all supported types.
            confidence_threshold: Presidio score floor (default 0.35).
            dry_run:     If True, detect only — return entity list, write no file.
            also_sanitize_metadata: When True (default), apply level-3 metadata
                                    sanitization on the output (clears creator,
                                    company, revision, attachedTemplate).
            redact_authors_as: Author label for metadata sanitization (level 2).

        Returns:
            {"path": str|None, "entities": [...]}
        """
        if not dry_run and not output_path:
            raise ValueError("output_path is required when dry_run=False")
        if self.workdir is None:
            raise RuntimeError("No document is open")

        analyzer = _get_analyzer()
        entity_filter = list(entities) if entities else None

        doc_tree = self._trees.get("word/document.xml")
        if doc_tree is None:
            return {"path": None if dry_run else output_path, "entities": []}

        # ── Phase 1: detect PII per paragraph ───────────────────────────────
        body = doc_tree.find(_wt("body"))
        paragraphs = list(body) if body is not None else []

        detected: list[dict] = []  # {type, text, para_index, start, end, score}

        for para_idx, para in enumerate(paragraphs):
            if para.tag != _wt("p"):
                continue
            full_text, _ = _build_run_char_map(para)
            if not full_text.strip():
                continue

            results = analyzer.analyze(
                text=full_text,
                language="en",
                entities=entity_filter,
                score_threshold=confidence_threshold,
            )
            for r in results:
                detected.append(
                    {
                        "type": r.entity_type,
                        "text": full_text[r.start : r.end],
                        "para_index": para_idx,
                        "start": r.start,
                        "end": r.end,
                        "score": round(r.score, 3),
                    }
                )

        if dry_run:
            return {"path": None, "entities": detected}

        # ── Phase 2: deduplication ───────────────────────────────────────────
        # Collect unique PII strings, then find ALL occurrences across the doc
        unique_pii_texts = {d["text"] for d in detected}

        # Build redaction spans per paragraph (index → [(start, end)])
        redaction_map: dict[int, list[tuple[int, int]]] = {}

        for para_idx, para in enumerate(paragraphs):
            if para.tag != _wt("p"):
                continue
            full_text, _ = _build_run_char_map(para)
            if not full_text:
                continue

            spans: list[tuple[int, int]] = []
            for pii_text in unique_pii_texts:
                spans.extend(_find_all_occurrences(full_text, pii_text))

            if spans:
                redaction_map[para_idx] = _merge_overlapping(spans)

        # ── Phase 3: redact on a deep copy of the document tree ─────────────
        doc_copy = copy.deepcopy(doc_tree)
        body_copy = doc_copy.find(_wt("body"))
        copy_paragraphs = list(body_copy) if body_copy is not None else []

        _used_ids: set[int] = set()

        for para_idx, para in enumerate(copy_paragraphs):
            if para.tag != _wt("p"):
                continue
            if para_idx not in redaction_map:
                continue

            full_text, run_ranges = _build_run_char_map(para)
            # Redact spans in reverse order to preserve offsets
            for span_start, span_end in reversed(redaction_map[para_idx]):
                _redact_span(run_ranges, span_start, span_end, doc_copy, _used_ids)
                # Rebuild run_ranges after each modification (indices shift)
                full_text, run_ranges = _build_run_char_map(para)

        out_doc_bytes = etree.tostring(
            doc_copy, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        # ── Phase 4: write output zip ────────────────────────────────────────
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out_zip:
            for root_dir, _dirs, files in os.walk(self.workdir):
                for fname in sorted(files):
                    fpath = Path(root_dir) / fname
                    # "/" separators so the part-name checks below match on Windows too.
                    arcname = fpath.relative_to(self.workdir).as_posix()

                    if arcname == "word/document.xml":
                        out_zip.writestr(arcname, out_doc_bytes)
                    elif also_sanitize_metadata and arcname in (
                        "word/settings.xml",
                        "docProps/core.xml",
                    ):
                        if arcname in self._trees:
                            el = copy.deepcopy(self._trees[arcname])
                            if arcname == "word/settings.xml":
                                self._sanitize_settings_el(el)
                            else:
                                self._sanitize_core_el(el)
                            out_zip.writestr(
                                arcname,
                                etree.tostring(
                                    el, xml_declaration=True, encoding="UTF-8", standalone=True
                                ),
                            )
                        else:
                            out_zip.writestr(arcname, fpath.read_bytes())
                    elif also_sanitize_metadata and arcname == "docProps/app.xml":
                        raw = fpath.read_bytes()
                        app_el = etree.fromstring(raw)
                        self._sanitize_app_el(app_el)
                        out_zip.writestr(
                            arcname,
                            etree.tostring(
                                app_el, xml_declaration=True, encoding="UTF-8", standalone=True
                            ),
                        )
                    elif arcname in self._trees:
                        out_zip.writestr(
                            arcname,
                            etree.tostring(
                                self._trees[arcname],
                                xml_declaration=True,
                                encoding="UTF-8",
                                standalone=True,
                            ),
                        )
                    else:
                        out_zip.writestr(arcname, fpath.read_bytes())

        return {"path": output_path, "entities": detected}
