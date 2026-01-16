#!/bin/bash

# Auto-start script for Streamlit - Place in .bashrc or run on Codespace start

SCRIPT_DIR="/workspaces/codespaces-blank"
LOGFILE="$SCRIPT_DIR/streamlit.log"

# Function to check if Streamlit is running
is_running() {
    lsof -i :8501 >/dev/null 2>&1
}

# Wait for system to be ready
sleep 2

# Start if not already running
if ! is_running; then
    echo "🚀 Auto-starting Streamlit..."
    cd "$SCRIPT_DIR"
    bash "$SCRIPT_DIR/keep-alive.sh"
    
    # Start monitor in background
    nohup bash "$SCRIPT_DIR/monitor.sh" > /dev/null 2>&1 &
    
    echo "✅ Streamlit auto-started and monitoring enabled"
    echo "🌐 Check Ports tab in VS Code for the forwarded URL"
else
    echo "✅ Streamlit already running on port 8501"
fi
