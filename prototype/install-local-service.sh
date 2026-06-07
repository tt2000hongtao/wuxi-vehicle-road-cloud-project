#!/bin/bash
set -euo pipefail

LABEL="com.tianan.wuxi-roadside-prototype"
PROJECT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
SERVICE_DIR="$HOME/Library/Application Support/wuxi-roadside-prototype"
LOG_DIR="$SERVICE_DIR/logs"
RUN_SCRIPT="$SERVICE_DIR/run-local-service.sh"

mkdir -p "$HOME/Library/LaunchAgents" "$LOG_DIR"
cat > "$RUN_SCRIPT" <<SCRIPT
#!/bin/bash
set -euo pipefail
cd "$PROJECT_DIR"
exec /opt/homebrew/bin/node prototype/server.js
SCRIPT
chmod +x "$RUN_SCRIPT"

cat > "$PLIST" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>$LABEL</string>
  <key>ProgramArguments</key>
  <array>
    <string>/bin/bash</string>
    <string>$RUN_SCRIPT</string>
  </array>
  <key>WorkingDirectory</key>
  <string>$SERVICE_DIR</string>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>StandardOutPath</key>
  <string>$LOG_DIR/roadside-service.out.log</string>
  <key>StandardErrorPath</key>
  <string>$LOG_DIR/roadside-service.err.log</string>
  <key>EnvironmentVariables</key>
  <dict>
    <key>PORT</key>
    <string>4173</string>
    <key>PYTHON_BIN</key>
    <string>/Library/Frameworks/Python.framework/Versions/3.11/bin/python3</string>
  </dict>
</dict>
</plist>
PLIST

launchctl bootout "gui/$(id -u)" "$PLIST" >/dev/null 2>&1 || true
launchctl bootstrap "gui/$(id -u)" "$PLIST"
launchctl kickstart -k "gui/$(id -u)/$LABEL"

echo "Installed and started $LABEL"
echo "Open http://127.0.0.1:4173"
echo "State file: $PROJECT_DIR/prototype/storage/roadside-status-state.json"
