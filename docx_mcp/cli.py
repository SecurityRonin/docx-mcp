"""CLI entry point with subcommand dispatch."""

from __future__ import annotations

import shutil
import sys
from pathlib import Path


def _skill_source() -> Path:
    """Return the path to the bundled SKILL.md."""
    return Path(__file__).parent / "skill" / "SKILL.md"


def _skill_target_dir() -> Path:
    return Path.home() / ".claude" / "skills" / "docx-mcp"


def _needs_update(source: Path, dest: Path) -> bool:
    """Return True if dest is missing or older than source."""
    if not dest.exists():
        return True
    return source.stat().st_size != dest.stat().st_size


def install_skill(*, target_dir: Path | None = None) -> Path:
    """Copy the bundled SKILL.md into the Claude Code skills directory.

    Returns the path to the installed skill file.
    """
    if target_dir is None:
        target_dir = _skill_target_dir()
    target_dir.mkdir(parents=True, exist_ok=True)
    dest = target_dir / "SKILL.md"
    shutil.copy2(_skill_source(), dest)
    return dest


def auto_install_skill() -> None:
    """Silently install or update the skill on server startup.

    Never raises — server startup must not fail because of a skill install issue.
    """
    try:
        source = _skill_source()
        dest = _skill_target_dir() / "SKILL.md"
        if _needs_update(source, dest):
            install_skill()
    except Exception:  # noqa: BLE001
        pass  # Never block server startup


def run_server() -> None:
    """Auto-install skill, then start the MCP server."""
    auto_install_skill()

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
