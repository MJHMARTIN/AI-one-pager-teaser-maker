#!/bin/bash
# Keep Streamlit running persistently

LOGFILE="/workspaces/codespaces-blank/streamlit.log"
PIDFILE="/workspaces/codespaces-blank/streamlit.pid"

# Function to start Streamlit
start_streamlit() {
    echo "$(date): Starting Streamlit..." | tee -a "$LOGFILE"
    cd /workspaces/codespaces-blank
    nohup streamlit run app.py \
        --server.port 8501 \
        --server.address 0.0.0.0 \
        --server.headless true \
        >> "$LOGFILE" 2>&1 &
    
    echo $! > "$PIDFILE"
    echo "$(date): Streamlit started with PID $(cat $PIDFILE)" | tee -a "$LOGFILE"
}

# Check if Streamlit is already running
if [ -f "$PIDFILE" ] && kill -0 $(cat "$PIDFILE") 2>/dev/null; then
    echo "Streamlit is already running with PID $(cat $PIDFILE)"
    exit 0
fi

# Check if port 8501 is in use
if lsof -i :8501 >/dev/null 2>&1; then
    echo "Port 8501 is in use. Killing existing process..."
    lsof -ti :8501 | xargs kill -9 2>/dev/null
    sleep 2
fi

# Start Streamlit
start_streamlit

# Wait a moment and verify it started
sleep 3
if lsof -i :8501 >/dev/null 2>&1; then
    echo "✅ Streamlit is running successfully on port 8501"
    echo "🌐 Access at: http://localhost:8501"
else
    echo "❌ Failed to start Streamlit. Check $LOGFILE for details"
    exit 1
fi
