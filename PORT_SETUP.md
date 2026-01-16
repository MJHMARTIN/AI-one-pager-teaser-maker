# 🚀 Port Management & Auto-Start Setup

## What's Been Configured

Your Streamlit app is now configured to:

1. **Auto-start** when the Codespace opens
2. **Always bind to `0.0.0.0:8501`** (proper for Codespaces port forwarding)
3. **Auto-restart** if the process dies
4. **Use public port visibility** for easy access

## Files Configured

### `.streamlit/config.toml`
- Binds to `0.0.0.0:8501` (all interfaces)
- Disables CORS and XSRF for Codespaces compatibility
- Runs in headless mode (no browser pop-ups)

### `.devcontainer/devcontainer.json`
- Forwards port 8501 with public visibility
- Auto-runs `auto-start.sh` on Codespace startup
- Installs dependencies automatically

### Scripts Available

| Script | Purpose | When to Use |
|--------|---------|-------------|
| `start.sh` | Clean start of Streamlit | First time or clean start |
| `restart.sh` | Kill and restart everything | When you get 502 errors |
| `keep-alive.sh` | Start with process management | For stable running |
| `monitor.sh` | Background auto-restart daemon | Already running automatically |
| `auto-start.sh` | Codespace startup script | Auto-runs on Codespace start |
| `status.sh` | Check if everything is running | Troubleshooting |

## Quick Commands

```bash
# Check status
bash status.sh

# Restart everything (fixes 502 errors)
bash restart.sh

# View live logs
tail -f streamlit.log

# Kill everything manually
pkill -f streamlit
pkill -f monitor.sh

# Check port
lsof -i :8501
```

## How It Works

1. **On Codespace Start**: `.devcontainer/devcontainer.json` → runs `auto-start.sh`
2. **Auto-start runs**: `keep-alive.sh` (starts Streamlit) + `monitor.sh` (background)
3. **Monitor watches**: Checks every 30 seconds if port 8501 is alive
4. **If dead**: Monitor automatically runs `keep-alive.sh` again
5. **Config ensures**: Always binds to `0.0.0.0:8501` for proper forwarding

## Troubleshooting 502 Errors

### Why 502 happens:
- Old forwarded URL from before restart
- Process died but URL still exists
- Wrong host binding (localhost vs 0.0.0.0)

### Solution (every time):
1. Run `bash restart.sh`
2. Open **Ports tab** in VS Code (View → Ports)
3. Use the **fresh URL** shown there (don't reuse old bookmarks!)
4. Ensure visibility is **Public** (not Private)

### The 0.0.0.0 vs localhost issue:
- ❌ `localhost` → Only accessible inside the container
- ✅ `0.0.0.0` → Accessible from port forwarding (required for Codespaces)

This is now configured automatically in `.streamlit/config.toml`.

## What Changed

**Before:**
- CORS/XSRF enabled (can cause issues)
- Port visibility might be private
- No auto-start on Codespace open
- Manual restart required

**After:**
- CORS/XSRF disabled (Codespaces-friendly)
- Port 8501 always public
- Auto-starts on Codespace creation
- Auto-restarts if process dies
- Always binds to 0.0.0.0

## Testing

Run `bash status.sh` to see:
- ✅ Process running
- ✅ Port listening on 8501
- ✅ Monitor active
- 📝 Recent logs

Everything should show ✅ green checkmarks!
