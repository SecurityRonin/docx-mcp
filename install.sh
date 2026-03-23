#!/usr/bin/env bash
# Install docx-mcp: MCP server + Claude Code skill
# Usage: curl -sSL https://raw.githubusercontent.com/SecurityRonin/docx-mcp/main/install.sh | bash
set -euo pipefail

echo "Installing docx-mcp..."

# 1. Install the package
if command -v uvx &>/dev/null; then
  uvx --from docx-mcp-server docx-mcp install-skill
  echo "  ✓ Skill installed"
elif command -v pip &>/dev/null; then
  pip install docx-mcp-server
  docx-mcp install-skill
  echo "  ✓ Package + skill installed"
else
  echo "  ⚠ Neither uvx nor pip found — install Python 3.10+ first"
  exit 1
fi

# 2. Add MCP server to Claude Code
if command -v claude &>/dev/null; then
  claude mcp add docx-mcp -- uvx docx-mcp-server
  echo "  ✓ MCP server added to Claude Code"
else
  echo "  ⚠ Claude Code CLI not found — add manually to your MCP config:"
  echo '    {"mcpServers":{"docx-mcp":{"command":"uvx","args":["docx-mcp-server"]}}}'
fi

echo ""
echo "Done! Start a new Claude Code session to use docx-mcp."
echo "Try: \"Open contract.docx and change 'Net 30' to 'Net 60' with track changes\""
