# LR Database Docker Image
# Multi-platform: linux/amd64, linux/arm64

FROM node:20-alpine AS builder

# Install Python + openpyxl
RUN apk add --no-cache python3 py3-pip && pip3 install --no-cache-dir openpyxl

# Install Ollama (optional - for local LLM)
# RUN curl -fsSL https://ollama.com/install.sh | sh

WORKDIR /app

# Copy package files
COPY core/package*.json ./core/
COPY python/requirements.txt ./python/

# Install core dependencies
RUN cd core && npm ci

# Copy source
COPY core/src ./core/src
COPY core/tsconfig.json ./core/
COPY python ./python

# Build TypeScript
RUN cd core && npm run build

# ============================================================================
# Runtime Stage
# ============================================================================

FROM node:20-alpine

RUN apk add --no-cache python3 py3-pip && pip3 install --no-cache-dir openpyxl

WORKDIR /app

# Copy built core
COPY --from=builder /app/core/dist ./core/dist
COPY --from=builder /app/core/package.json ./core/
COPY --from=builder /app/core/node_modules ./core/node_modules
COPY --from=builder /app/python ./python

# Copy adapters
COPY adapters/cli ./adapters/cli
COPY adapters/claude-mcp ./adapters/claude-mcp

# Create project directories
RUN mkdir -p /project/Incoming /project/Database /project/Outputs /project/Set-Aside

WORKDIR /project

# Entry point
ENTRYPOINT ["node", "/app/adapters/cli/bin/lr-db.js"]

# Default command
CMD ["batch", "-r", "/project"]