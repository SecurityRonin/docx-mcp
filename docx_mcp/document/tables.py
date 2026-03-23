"""Tables mixin: read table content."""

from __future__ import annotations

from .base import W


class TablesMixin:
    """Table operations."""

    def get_tables(self) -> list[dict]:
        """Get all tables with their content."""
        doc = self._require("word/document.xml")
        tables = []
        for idx, tbl in enumerate(doc.iter(f"{W}tbl")):
            rows = []
            for tr in tbl.findall(f"{W}tr"):
                cells = []
                for tc in tr.findall(f"{W}tc"):
                    cells.append(self._text(tc))
                rows.append(cells)
            col_count = len(rows[0]) if rows else 0
            tables.append(
                {
                    "index": idx,
                    "row_count": len(rows),
                    "col_count": col_count,
                    "cells": rows,
                }
            )
        return tables
