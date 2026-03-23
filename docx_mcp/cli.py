"""CLI entry point with subcommand dispatch."""

from __future__ import annotations

import shutil
import sys
from pathlib import Path


def _skill_source() -> Path:
    """Return the path to the bundled SKILL.md."""
    return Path(__file__).parent / "skill" / "SKILL.md"


def install_skill(*, target_dir: Path | None = None) -> Path:
    """Copy the bundled SKILL.md into the Claude Code skills directory.

    Returns the path to the installed skill file.
    """
    if target_dir is None:
        target_dir = Path.home() / ".claude" / "skills" / "docx-mcp"
    target_dir.mkdir(parents=True, exist_ok=True)
    dest = target_dir / "SKILL.md"
    shutil.copy2(_skill_source(), dest)
    return dest


def run_server() -> None:
    """Start the MCP server (default behavior)."""
    from docx_mcp.server import main as server_main

    server_main()


def main() -> None:
    """Dispatch: no args → MCP server, subcommand → handle it."""
    args = sys.argv[1:]

    if not args:
        run_server()
        return

    cmd = args[0]

    if cmd in ("install-skill", "update-skill"):
        dest = install_skill()
        print(f"Skill installed to {dest}")
        return

    print(f"Unknown command: {cmd}", file=sys.stderr)
    print("Usage: docx-mcp [install-skill | update-skill]", file=sys.stderr)
    raise SystemExit(1)
