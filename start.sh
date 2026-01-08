#!/bin/bash

# AI One Pager Teaser Maker - Startup Script
# This script ensures Streamlit runs properly on port 8501

echo "🚀 Starting AI One Pager Teaser Maker..."

# Kill any existing process on port 8501
if lsof -i :8501 > /dev/null 2>&1; then
    echo "⚠️  Port 8501 is already in use. Killing existing process..."
    lsof -ti :8501 | xargs kill -9 2>/dev/null || true
    sleep 2
fi

# Ensure dependencies are installed
if ! pip show streamlit > /dev/null 2>&1; then
    echo "📦 Installing dependencies..."
    pip install -r requirements.txt
fi

# Start Streamlit with nohup so it keeps running
echo "✅ Starting Streamlit on port 8501..."
echo "🌐 Access the app at: http://localhost:8501"
echo "📝 Logs are saved to streamlit.log"
echo ""

cd /workspaces/codespaces-blank
nohup streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --server.headless true > streamlit.log 2>&1 &

echo $! > streamlit.pid
echo "✅ Started with PID $(cat streamlit.pid)"
echo ""
echo "💡 To stop: kill \$(cat streamlit.pid)"
echo "📊 To monitor: tail -f streamlit.log"
