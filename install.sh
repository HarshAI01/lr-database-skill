#!/bin/bash
# LR Database Skill - One-line installer
# Usage: curl -fsSL https://raw.githubusercontent.com/HarshAI01/lr-database-skill/main/install.sh | bash

set -e

SKILL_NAME="lr-database"
SKILL_DIR="$HOME/.opencode/skill/$SKILL_NAME"
REPO_URL="https://github.com/HarshAI01/lr-database-skill"

echo "📦 Installing LR Database Skill..."

# Remove old version
rm -rf "$SKILL_DIR"

# Clone fresh
git clone --depth 1 "$REPO_URL" "$SKILL_DIR"

# Install dependencies
cd "$SKILL_DIR"
npm install

echo ""
echo "✅ LR Database Skill installed!"
echo ""
echo "🚀 Quick start:"
echo "  mkdir my-review && cd my-review"
echo "  mkdir -p Incoming"
echo "  cp /path/to/papers/*.pdf Incoming/"
echo "  opencode run lr-batch"
echo ""
echo "🤖 Configure your model in OpenCode:"
echo "  opencode config set model openai/gpt-4o"
echo "  # or anthropic/claude-3-5-sonnet, gemini/gemini-1.5-pro, ollama/llama3.1, etc."
echo ""
echo "📖 Full docs: https://github.com/HarshAI01/lr-database-skill"