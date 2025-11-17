#!/bin/bash

# Master startup script for Outlook MCP services
# This script cleans up old processes and starts both the auth server and MCP server

set -e

PROJECT_DIR="/Users/petermandy/Documents/GitHub/outlook-mcp"
CLEANUP_SCRIPT="$PROJECT_DIR/scripts/cleanup-ports.sh"
LOG_DIR="/tmp"

echo "[$(date +'%Y-%m-%d %H:%M:%S')] Starting Outlook MCP services..."

# Run cleanup script
echo "Running port cleanup..."
if [ -f "$CLEANUP_SCRIPT" ]; then
  bash "$CLEANUP_SCRIPT" 2>&1 || true
else
  echo "Warning: Cleanup script not found at $CLEANUP_SCRIPT"
fi

# Start auth server in background
echo "Starting authentication server..."
cd "$PROJECT_DIR"
npm run auth-server > "$LOG_DIR/outlook-mcp-auth-server.log" 2>&1 &
AUTH_PID=$!
echo "Auth server started (PID: $AUTH_PID)"

# Wait for auth server to be ready
sleep 2

# Start main MCP server (will be managed by parent process)
echo "Starting MCP server..."
exec node "$PROJECT_DIR/index.js"
