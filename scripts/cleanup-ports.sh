#!/bin/bash

# Cleanup script to gracefully shut down old processes before starting new ones
# This prevents port conflicts and zombie processes

set -e

PROJECT_DIR="/Users/petermandy/Documents/GitHub/outlook-mcp"
PORTS=(3333 6274 6277)

echo "[$(date +'%Y-%m-%d %H:%M:%S')] Starting cleanup..."

# Kill any existing outlook-mcp or auth-server processes gracefully
echo "Cleaning up existing outlook-mcp processes..."
pkill -f "outlook-mcp|auth-server" 2>/dev/null || true

# Kill any existing inspector processes
echo "Cleaning up inspector processes..."
pkill -f "@modelcontextprotocol/inspector" 2>/dev/null || true

# Wait a bit for graceful shutdown
sleep 1

# Force kill any remaining processes on our ports
for PORT in "${PORTS[@]}"; do
  echo "Checking port $PORT..."
  PID=$(lsof -i :$PORT 2>/dev/null | grep -v COMMAND | awk '{print $2}' | head -1 || true)
  if [ -n "$PID" ]; then
    echo "  Killing process $PID on port $PORT"
    kill -9 "$PID" 2>/dev/null || true
  fi
done

# Wait for ports to be released
sleep 2

# Verify ports are free
echo "Verifying ports are free..."
for PORT in "${PORTS[@]}"; do
  if lsof -i :$PORT >/dev/null 2>&1; then
    echo "  ⚠️  Port $PORT is still in use (may be OK if it's a system service)"
  else
    echo "  ✓ Port $PORT is free"
  fi
done

echo "[$(date +'%Y-%m-%d %H:%M:%S')] Cleanup completed"
