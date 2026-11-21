#!/bin/bash
# Universal LR Database Installer
# Works for: OpenCode, Claude Desktop (MCP), Cursor, CLI, Docker
# Usage: curl -fsSL https://raw.githubusercontent.com/HarshAI01/lr-database-skill/main/install-universal.sh | bash

set -e

REPO="https://github.com/HarshAI01/lr-database-skill"
SKILL_DIR="$HOME/.opencode/skill/lr-database"
CLAUDE_MCP_DIR="$HOME/.config/claude/mcp/lr-database"
CURSOR_DIR="$HOME/.cursor/rules"
CLI_DIR="$HOME/.local/bin"

echo "📦 LR Database Universal Installer"
echo "=================================="
echo ""

# Detect platform
detect_platform() {
    echo "Detecting your setup..."
    
    PLATFORMS=()
    
    if command -v opencode &> /dev/null; then
        PLATFORMS+=("opencode")
        echo "  ✅ OpenCode detected"
    fi
    
    if [ -d "$HOME/Library/Application Support/Claude" ] || [ -d "$HOME/.config/Claude" ]; then
        PLATFORMS+=("claude-mcp")
        echo "  ✅ Claude Desktop detected"
    fi
    
    if [ -d "$HOME/.cursor" ] || [ -d "$HOME/.config/cursor" ]; then
        PLATFORMS+=("cursor")
        echo "  ✅ Cursor detected"
    fi
    
    if command -v code &> /dev/null || command -v cursor &> /dev/null; then
        PLATFORMS+=("vscode")
        echo "  ✅ VS Code detected"
    fi
    
    PLATFORMS+=("cli")
    echo "  ✅ CLI (always available)"
    
    if command -v docker &> /dev/null; then
        PLATFORMS+=("docker")
        echo "  ✅ Docker detected"
    fi
    
    echo ""
}

# Install for OpenCode
install_opencode() {
    echo "📦 Installing for OpenCode..."
    rm -rf "$SKILL_DIR"
    git clone --depth 1 "$REPO" "$SKILL_DIR"
    cd "$SKILL_DIR"
    npm install
    echo "✅ OpenCode skill installed at $SKILL_DIR"
    echo "   Run: opencode run lr-batch"
}

# Install for Claude Desktop (MCP)
install_claude_mcp() {
    echo "📦 Installing for Claude Desktop (MCP)..."
    mkdir -p "$CLAUDE_MCP_DIR"
    cd "$CLAUDE_MCP_DIR"
    git clone --depth 1 "$REPO" .
    cd adapters/claude-mcp
    npm install
    echo "✅ MCP server installed at $CLAUDE_MCP_DIR"
    echo ""
    echo "Add to Claude Desktop config (~/.config/claude/claude_desktop_config.json):"
    echo '{'
    echo '  "mcpServers": {'
    echo '    "lr-database": {'
    echo '      "command": "node",'
    echo '      "args": ["'$CLAUDE_MCP_DIR'/adapters/claude-mcp/server.js"],'
    echo '      "env": { "OPENAI_API_KEY": "sk-..." }'
    echo '    }'
    echo '  }'
    echo '}'
}

# Install for Cursor
install_cursor() {
    echo "📦 Installing for Cursor..."
    mkdir -p "$CURSOR_DIR"
    cat > "$CURSOR_DIR/lr-database.mdc" << 'EOF'
---
description: Literature Review Database Pipeline
globs: ["**/Incoming/*.pdf"]
alwaysApply: false
---

# LR Database Skill for Cursor

When user mentions "literature review", "LR database", "process papers", or drops PDFs in `Incoming/`:
1. Use the `lr-database` commands below
2. Works with any model configured in Cursor (OpenAI, Anthropic, etc.)
3. Output goes to `Outputs/` and `SOP-Log.md`

## Commands

### Run Full Pipeline
```bash
npx tsx ~/.opencode/skill/lr-database/adapters/opencode/tools/lr-batch.ts -r .
```

### Individual Steps
```bash
# 1. Ingest - normalize filenames
npx tsx ~/.opencode/skill/lr-database/adapters/opencode/tools/lr-ingest.ts -r .

# 2. Dedup - find duplicates
npx tsx ~/.opencode/skill/lr-database/adapters/opencode/tools/lr-dedup.ts -r .

# 3. Extract PDFs - get content for LLM
npx tsx ~/.opencode/skill/lr-database/adapters/opencode/tools/lr-extract-pdfs.ts -r .

# 4. Assemble - write to Excel (after LLM extraction)
npx tsx ~/.opencode/skill/lr-database/adapters/opencode/tools/lr-assemble.ts -r . -i extractions.json

# 5. Verify - check quality
npx tsx ~/.opencode/skill/lr-database/adapters/opencode/tools/lr-verify.ts -r .

# 6. Log - write SOP-Log.md
npx tsx ~/.opencode/skill/lr-database/adapters/opencode/tools/lr-log.ts -r . -k 5 -d 0 --serial-start 1 --serial-end 5
```

## Extraction Workflow

1. Run `lr-extract-pdfs` to get PDF content
2. Copy the extraction prompt and PDF content
3. Paste into Cursor Chat with your model
4. Get JSON result, save as `extractions.json`
5. Run `lr-assemble` with the extracted data
EOF
    echo "✅ Cursor rule installed at $CURSOR_DIR/lr-database.mdc"
}

# Install CLI
install_cli() {
    echo "📦 Installing CLI..."
    mkdir -p "$CLI_DIR"
    cd "$CLI_DIR"
    git clone --depth 1 "$REPO" lr-database-cli 2>/dev/null || (cd lr-database-cli && git pull)
    cd lr-database-cli/adapters/cli
    npm install
    npm link
    echo "✅ CLI installed as 'lr-db' command"
    echo "   Run: lr-db batch -r /path/to/project"
}

# Install Docker
install_docker() {
    echo "📦 Installing Docker image..."
    cd /tmp
    git clone --depth 1 "$REPO" lr-database-docker
    cd lr-database-docker/adapters/docker
    docker build -t lr-database .
    echo "✅ Docker image built: lr-database"
    echo "   Run: docker run -v /path/to/project:/project lr-database batch -r /project"
}

# Main
detect_platform

echo "Select installation targets:"
echo "  1) All detected platforms"
echo "  2) OpenCode only"
echo "  3) Claude Desktop (MCP) only"
echo "  4) Cursor only"
echo "  5) CLI only"
echo "  6) Docker only"
echo "  7) Custom selection"
read -p "Choice [1-7]: " choice

case $choice in
    1)
        for p in "${PLATFORMS[@]}"; do install_$p; done
        ;;
    2) install_opencode ;;
    3) install_claude_mcp ;;
    4) install_cursor ;;
    5) install_cli ;;
    6) install_docker ;;
    7)
        for p in "${PLATFORMS[@]}"; do
            read -p "Install for $p? [y/N]: " yn
            [[ $yn =~ ^[Yy]$ ]] && install_$p
        done
        ;;
    *) echo "Invalid choice"; exit 1 ;;
esac

echo ""
echo "🎉 Installation complete!"
echo ""
echo "📖 Next steps:"
echo "  1. Configure your model in OpenCode/Cursor/Claude"
echo "  2. mkdir my-review && cd my-review"
echo "  3. mkdir -p Incoming"
echo "  4. cp ~/papers/*.pdf Incoming/"
echo "  5. Run the pipeline for your platform"
echo ""
echo "📚 Docs: https://github.com/HarshAI01/lr-database-skill"