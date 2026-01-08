#!/bin/bash

# Quick restart script for the app

echo "🔄 Restarting AI One Pager Teaser Maker..."

# Kill any process on port 8501
if lsof -i :8501 > /dev/null 2>&1; then
    echo "⏹️  Stopping existing process..."
    lsof -ti :8501 | xargs kill -9 2>/dev/null || true
    sleep 2
fi

# Kill the monitor if running
if pgrep -f "monitor.sh" > /dev/null; then
    echo "⏹️  Stopping monitor..."
    pkill -f "monitor.sh"
fi

# Start fresh
echo "🚀 Starting fresh..."
bash /workspaces/codespaces-blank/keep-alive.sh
nohup bash /workspaces/codespaces-blank/monitor.sh > /dev/null 2>&1 &

echo "✅ App restarted and monitoring enabled"
echo "🌐 Access at: http://localhost:8501"
