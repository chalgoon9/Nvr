#!/usr/bin/env bash
set -euo pipefail

APP_DIR=/opt/rc73xa_watcher
SERVICE_NAME=rc73xa-watcher
SERVICE_FILE=/etc/systemd/system/${SERVICE_NAME}.service
ENV_FILE=/etc/${SERVICE_NAME}.env

echo "[+] Creating app directory: ${APP_DIR}"
sudo mkdir -p "${APP_DIR}"

echo "[+] Copying watcher script"
sudo cp -f "$(dirname "$0")/rc73xa_watcher.py" "${APP_DIR}/rc73xa_watcher.py"

echo "[+] Preparing permissions"
CURRENT_USER=$(whoami)
sudo chown -R "${CURRENT_USER}:${CURRENT_USER}" "${APP_DIR}"

echo "[+] Creating venv and installing deps (as ${CURRENT_USER})"
sudo -u "${CURRENT_USER}" bash -lc "python3 -m venv '${APP_DIR}/.venv' && source '${APP_DIR}/.venv/bin/activate' && pip install --upgrade pip && pip install playwright python-dotenv && playwright install firefox"

echo "[+] (Optional) Installing Playwright system deps via apt"
sudo bash -lc "playwright install-deps || true"

echo "[+] Writing environment file: ${ENV_FILE}"
if [ ! -f "${ENV_FILE}" ]; then
  sudo tee "${ENV_FILE}" >/dev/null <<'EOF'
# Required
TG_BOT_TOKEN=
TG_CHAT_ID=

# Optional
HEARTBEAT_MINUTES=360
TG_COMMANDS_ENABLED=1
TG_ALLOWED_CHAT_ID=
MIN_INTERVAL=3
MAX_INTERVAL=7
MAX_RETURN_PRICE=1000000
QUERY=RC73XA-NH011W
TARGET_MODEL=rc73xa\s*-?\s*nh011w
EXCLUDES=["rc71l\\s*-?\\s*nh001w"]
RETURN_KEYWORDS=["\uBC18\uD488","\uB9AC\uD37C"]
HEADLESS=1
EOF
else
  echo "    Environment already exists; edit ${ENV_FILE} if needed."
fi

echo "[+] Installing systemd service: ${SERVICE_FILE}"
sudo tee "${SERVICE_FILE}" >/dev/null <<EOF
[Unit]
Description=RC73XA Coupang Return Watcher
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
User=${CURRENT_USER}
WorkingDirectory=${APP_DIR}
EnvironmentFile=${ENV_FILE}
ExecStart=${APP_DIR}/.venv/bin/python ${APP_DIR}/rc73xa_watcher.py
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
EOF

echo "[+] Reloading and starting service"
sudo systemctl daemon-reload
sudo systemctl enable --now ${SERVICE_NAME}.service
sudo systemctl status ${SERVICE_NAME}.service --no-pager -l || true

echo "[+] Done. Use: journalctl -u ${SERVICE_NAME} -f"
