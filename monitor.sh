#!/bin/bash
# Monitor and restart Streamlit if it stops

LOGFILE="/workspaces/codespaces-blank/streamlit.log"

while true; do
    if ! lsof -i :8501 >/dev/null 2>&1; then
        echo "$(date): Streamlit not running. Restarting..." | tee -a "$LOGFILE"
        bash /workspaces/codespaces-blank/keep-alive.sh
    fi
    sleep 30  # Check every 30 seconds
done
