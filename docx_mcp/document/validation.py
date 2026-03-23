"""Validation mixin: paraids, watermark, audit."""

from __future__ import annotations

from .base import RELS, W14, A, R, V, W


class ValidationMixin:
    """Structural validation and audit."""

    def validate_paraids(self) -> dict:
        """Check paraId uniqueness across all document parts."""
        all_ids: dict[str, list[str]] = {}
        for rel_path, tree in self._trees.items():
            if not rel_path.endswith(".xml"):
                continue
            for elem in tree.iter():
                pid = elem.get(f"{W14}paraId")
                if pid:
                    all_ids.setdefault(pid, []).append(rel_path)

        duplicates = {k: v for k, v in all_ids.items() if len(v) > 1}
        invalid = []
        for pid in all_ids:
            try:
                if int(pid, 16) >= 0x80000000:
                    invalid.append(pid)
            except ValueError:
                invalid.append(pid)

        return {
            "valid": not duplicates and not invalid,
            "total": len(all_ids),
            "duplicates": duplicates,
            "out_of_range": invalid,
        }

    def remove_watermark(self) -> dict:
        """Remove VML watermarks (e.g., DRAFT) from all header XML files."""
        removed = []
        for rel_path, tree in self._trees.items():
            if "header" not in rel_path:
                continue
            for para in list(tree.iter(f"{W}p")):
                for run in list(para.findall(f"{W}r")):
                    pict = run.find(f"{W}pict")
                    if pict is None:
                        continue
                    for shape in pict.iter(f"{V}shape"):
                        tp = shape.find(f"{V}textpath")
                        if tp is not None:
                            wm_text = tp.get("string", "")
                            para.remove(run)
                            removed.append({"header": rel_path, "text": wm_text})
                            self._mark(rel_path)
                            break
        return {"removed": len(removed), "details": removed}

    def audit(self) -> dict:
        """Run comprehensive structural validation."""
        results: dict = {}

        results["footnotes"] = self.validate_footnotes()
        results["paraids"] = self.validate_paraids()

        # Headings
        doc = self._tree("word/document.xml")
        if doc is not None:
            headings = self._find_headings(doc)
            issues = []
            prev = 0
            for h in headings:
                if h["level"] > prev + 1 and prev > 0:
                    issues.append(
                        {
                            "issue": "level_skip",
                            "heading": h["text"][:60],
                            "expected_max": prev + 1,
                            "actual": h["level"],
                        }
                    )
                prev = h["level"]
            results["headings"] = {"count": len(headings), "issues": issues}

            # Bookmarks
            starts = {e.get(f"{W}id") for e in doc.iter(f"{W}bookmarkStart") if e.get(f"{W}id")}
            ends = {e.get(f"{W}id") for e in doc.iter(f"{W}bookmarkEnd") if e.get(f"{W}id")}
            results["bookmarks"] = {
                "total": len(starts),
                "unpaired_starts": len(starts - ends),
                "unpaired_ends": len(ends - starts),
            }

        # Relationships — check targets exist
        rels_tree = self._tree("word/_rels/document.xml.rels")
        rel_issues = []
        if rels_tree is not None:
            for rel in rels_tree.findall(f"{RELS}Relationship"):
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target", "")
                if not (self.workdir / "word" / target).exists():
                    rel_issues.append({"id": rel.get("Id"), "target": target})
        results["relationships"] = {"missing_targets": rel_issues}

        # Images
        img_issues = []
        if doc is not None and rels_tree is not None:
            for blip in doc.iter(f"{A}blip"):
                embed = blip.get(f"{R}embed")
                if not embed:
                    continue
                rel = rels_tree.find(f'{RELS}Relationship[@Id="{embed}"]')
                if rel is not None:
                    target = rel.get("Target", "")
                    if not (self.workdir / "word" / target).exists():
                        img_issues.append({"rId": embed, "target": target})
        results["images"] = {"missing": img_issues}

        # Endnotes
        results["endnotes"] = self.validate_endnotes()

        # Tables — check consistent column counts per table
        table_issues = []
        if doc is not None:
            for idx, tbl in enumerate(doc.iter(f"{W}tbl")):
                row_col_counts = []
                for tr in tbl.findall(f"{W}tr"):
                    row_col_counts.append(len(tr.findall(f"{W}tc")))
                if row_col_counts and len(set(row_col_counts)) > 1:
                    table_issues.append({"table_index": idx, "column_counts": row_col_counts})
        results["tables"] = {"inconsistent_columns": table_issues}

        # Protection status
        settings = self._tree("word/settings.xml")
        if settings is not None:
            prot = settings.find(f"{W}documentProtection")
            if prot is not None:
                results["protection"] = {
                    "edit": prot.get(f"{W}edit", ""),
                    "enforcement": prot.get(f"{W}enforcement", "0"),
                }
            else:
                results["protection"] = {"edit": "none", "enforcement": "0"}
        else:
            results["protection"] = {"edit": "unknown", "enforcement": "0"}

        # Artifacts
        artifacts = []
        for marker in ["DRAFT", "TODO", "FIXME", "XXX"]:
            for hit in self.search_text(marker):
                artifacts.append(
                    {"marker": marker, "source": hit["source"], "context": hit["text"][:100]}
                )
        results["artifacts"] = artifacts

        # Overall
        results["valid"] = (
            results["footnotes"].get("valid", True)
            and results["endnotes"].get("valid", True)
            and results["paraids"].get("valid", True)
            and not results["headings"].get("issues")
            and results["bookmarks"].get("unpaired_starts", 0) == 0
            and results["bookmarks"].get("unpaired_ends", 0) == 0
            and not rel_issues
            and not img_issues
            and not table_issues
        )
        return results
