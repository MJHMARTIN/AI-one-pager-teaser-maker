#!/bin/bash

# Quick status check for Streamlit

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "   🔍 Streamlit App Status Check"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

# Check if Streamlit process is running
if pgrep -f "streamlit run" > /dev/null; then
    PID=$(pgrep -f "streamlit run")
    echo "✅ Process Status: RUNNING (PID: $PID)"
else
    echo "❌ Process Status: NOT RUNNING"
fi

echo ""

# Check if port 8501 is listening
if lsof -i :8501 >/dev/null 2>&1; then
    echo "✅ Port 8501: LISTENING"
    echo ""
    lsof -i :8501 | head -5
else
    echo "❌ Port 8501: NOT LISTENING"
fi

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

# Check if monitor is running
if pgrep -f "monitor.sh" > /dev/null; then
    echo "✅ Monitor: ACTIVE (auto-restart enabled)"
else
    echo "⚠️  Monitor: NOT RUNNING (no auto-restart)"
fi

echo ""

# Show recent log lines
if [ -f "/workspaces/codespaces-blank/streamlit.log" ]; then
    echo "📝 Recent logs (last 5 lines):"
    echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    tail -5 /workspaces/codespaces-blank/streamlit.log
fi

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "💡 To restart: bash restart.sh"
echo "📊 To view live logs: tail -f streamlit.log"
echo "🌐 Check Ports tab in VS Code for the URL"
echo ""
