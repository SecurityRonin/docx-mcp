"""Tests for CLI subcommands (install-skill, update-skill, auto-install)."""

from __future__ import annotations

import sys
from pathlib import Path
from unittest import mock

import pytest

from docx_mcp import cli


class TestInstallSkill:
    def test_copies_skill_to_claude_dir(self, tmp_path: Path):
        target = tmp_path / ".claude" / "skills" / "docx-mcp"
        cli.install_skill(target_dir=target)
        skill_file = target / "SKILL.md"
        assert skill_file.exists()
        content = skill_file.read_text()
        assert "docx-mcp" in content
        assert "open_document" in content

    def test_overwrites_existing(self, tmp_path: Path):
        target = tmp_path / ".claude" / "skills" / "docx-mcp"
        target.mkdir(parents=True)
        (target / "SKILL.md").write_text("old content")
        cli.install_skill(target_dir=target)
        content = (target / "SKILL.md").read_text()
        assert "old content" not in content
        assert "open_document" in content

    def test_default_target_is_home_claude(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
        monkeypatch.setenv("HOME", str(tmp_path))
        monkeypatch.setattr(Path, "home", staticmethod(lambda: tmp_path))
        cli.install_skill()
        expected = tmp_path / ".claude" / "skills" / "docx-mcp" / "SKILL.md"
        assert expected.exists()

    def test_returns_path(self, tmp_path: Path):
        target = tmp_path / ".claude" / "skills" / "docx-mcp"
        result = cli.install_skill(target_dir=target)
        assert result == target / "SKILL.md"


class TestNeedsUpdate:
    def test_missing_file(self, tmp_path: Path):
        source = cli._skill_source()
        dest = tmp_path / "SKILL.md"
        assert cli._needs_update(source, dest) is True

    def test_same_size_no_update(self, tmp_path: Path):
        source = cli._skill_source()
        dest = tmp_path / "SKILL.md"
        # Copy so sizes match
        dest.write_bytes(source.read_bytes())
        assert cli._needs_update(source, dest) is False

    def test_different_size_needs_update(self, tmp_path: Path):
        source = cli._skill_source()
        dest = tmp_path / "SKILL.md"
        dest.write_text("stale")
        assert cli._needs_update(source, dest) is True


class TestAutoInstallSkill:
    def test_installs_when_missing(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
        monkeypatch.setattr(Path, "home", staticmethod(lambda: tmp_path))
        cli.auto_install_skill()
        assert (tmp_path / ".claude" / "skills" / "docx-mcp" / "SKILL.md").exists()

    def test_updates_when_stale(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
        monkeypatch.setattr(Path, "home", staticmethod(lambda: tmp_path))
        target = tmp_path / ".claude" / "skills" / "docx-mcp"
        target.mkdir(parents=True)
        (target / "SKILL.md").write_text("stale")
        cli.auto_install_skill()
        content = (target / "SKILL.md").read_text()
        assert "open_document" in content

    def test_skips_when_current(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
        monkeypatch.setattr(Path, "home", staticmethod(lambda: tmp_path))
        # Install once
        cli.auto_install_skill()
        dest = tmp_path / ".claude" / "skills" / "docx-mcp" / "SKILL.md"
        mtime_before = dest.stat().st_mtime_ns
        # Install again — should skip (no copy)
        cli.auto_install_skill()
        mtime_after = dest.stat().st_mtime_ns
        assert mtime_before == mtime_after

    def test_never_raises(self, monkeypatch: pytest.MonkeyPatch):
        """auto_install_skill swallows all errors — never blocks server startup."""
        monkeypatch.setattr(cli, "_skill_source", lambda: Path("/nonexistent/SKILL.md"))
        cli.auto_install_skill()  # should not raise


class TestRunServer:
    def test_run_server_auto_installs_and_calls_mcp(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ):
        """run_server() auto-installs skill then starts MCP server."""
        monkeypatch.setattr(Path, "home", staticmethod(lambda: tmp_path))
        called = []
        monkeypatch.setattr("docx_mcp.server.mcp.run", lambda: called.append(True))
        cli.run_server()
        assert called == [True]
        assert (tmp_path / ".claude" / "skills" / "docx-mcp" / "SKILL.md").exists()


class TestMainDispatch:
    def test_no_args_runs_mcp(self, monkeypatch: pytest.MonkeyPatch):
        """No arguments → start MCP server."""
        monkeypatch.setattr(sys, "argv", ["docx-mcp"])
        called = []
        with mock.patch("docx_mcp.cli.run_server", side_effect=lambda: called.append(True)):
            cli.main()
        assert called == [True]

    def test_install_skill_subcommand(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
        """'install-skill' arg → install skill, don't start server."""
        monkeypatch.setattr(sys, "argv", ["docx-mcp", "install-skill"])
        monkeypatch.setattr(Path, "home", staticmethod(lambda: tmp_path))
        cli.main()
        assert (tmp_path / ".claude" / "skills" / "docx-mcp" / "SKILL.md").exists()

    def test_update_skill_is_alias(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
        """'update-skill' is an alias for install-skill."""
        monkeypatch.setattr(sys, "argv", ["docx-mcp", "update-skill"])
        monkeypatch.setattr(Path, "home", staticmethod(lambda: tmp_path))
        cli.main()
        assert (tmp_path / ".claude" / "skills" / "docx-mcp" / "SKILL.md").exists()

    def test_unknown_subcommand_prints_help(
        self, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
    ):
        monkeypatch.setattr(sys, "argv", ["docx-mcp", "bogus"])
        with pytest.raises(SystemExit, match="1"):
            cli.main()
        captured = capsys.readouterr()
        assert "install-skill" in captured.err
