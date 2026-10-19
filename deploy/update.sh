#!/bin/bash
# Install dependencies, prepare the database and (re)start the service.
# Called by deploy.sh; can also be run by itself inside the app folder after copying new files.
set -euo pipefail

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
# Support the old zip layout where update.sh sits next to app.py
[ -f "$APP_DIR/app.py" ] || APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ENV_FILE="/etc/lotus-taskmanager.env"
SERVICE="lotus-taskmanager"
cd "$APP_DIR"

if [ "$(id -u)" -ne 0 ]; then echo "Run as root (sudo)."; exit 1; fi

mkdir -p /root/backups instance static/uploads
if [ -f instance/taskmanager.db ]; then
    BK="/root/backups/taskmanager-$(date +%Y%m%d-%H%M%S).db"
    sqlite3 instance/taskmanager.db ".backup '$BK'" 2>/dev/null || cp -a instance/taskmanager.db "$BK"
    echo "==> Database backed up to $BK"
fi
# Keep only the 30 newest DB backups
ls -1t /root/backups/taskmanager-*.db 2>/dev/null | tail -n +31 | xargs -r rm -f

echo "==> Python environment"
[ -d venv ] || python3 -m venv venv
./venv/bin/pip install -q --upgrade pip
./venv/bin/pip install -q -r requirements.txt

if [ ! -f "$ENV_FILE" ]; then
    echo "==> Creating $ENV_FILE (edit it to enable email)"
    SECRET="$( [ -s instance/secret_key ] && cat instance/secret_key || ./venv/bin/python -c 'import secrets; print(secrets.token_hex(32))')"
    cat > "$ENV_FILE" <<ENV
# Lotus Task Manager settings — after editing run: systemctl restart $SERVICE
SECRET_KEY=$SECRET
APP_PORT=5000
MAX_UPLOAD_MB=25
# Office 365 / SMTP email (leave empty to disable email; in-app notifications still work)
MAIL_SERVER=smtp.office365.com
MAIL_PORT=587
MAIL_USE_TLS=1
MAIL_USERNAME=
MAIL_PASSWORD=
ENV
    chmod 600 "$ENV_FILE"
fi

echo "==> Database"
set -a; . "$ENV_FILE"; set +a
LOTUS_DISABLE_SCHEDULER=1 ./venv/bin/python -c "import app; print('Lotus Task Manager v' + app.APP_VERSION + ' - DB OK')"

echo "==> Service"
# Stop any copy started the old way (nohup python app.py) that still holds the port
systemctl stop $SERVICE 2>/dev/null || true
if command -v fuser >/dev/null 2>&1; then fuser -k "${APP_PORT:-5000}/tcp" 2>/dev/null || true; sleep 1; fi
cat > /etc/systemd/system/$SERVICE.service <<UNIT
[Unit]
Description=Lotus Task Manager
After=network.target

[Service]
Type=simple
User=root
WorkingDirectory=$APP_DIR
EnvironmentFile=$ENV_FILE
# One worker only: the background scheduler (auto-cancel, deadlines, recurring tasks) must run once.
# Threads handle concurrent users.
ExecStart=$APP_DIR/venv/bin/gunicorn --workers 1 --threads 8 --timeout 120 --bind 0.0.0.0:\${APP_PORT} --access-logfile - app:app
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
UNIT
systemctl daemon-reload
systemctl enable -q $SERVICE
systemctl restart $SERVICE
sleep 3

# Open the port if the ufw firewall is active
if command -v ufw >/dev/null 2>&1 && ufw status | grep -q "Status: active"; then
    ufw allow "${APP_PORT:-5000}/tcp" >/dev/null || true
fi

if systemctl is-active -q $SERVICE && curl -fsS -o /dev/null "http://127.0.0.1:${APP_PORT:-5000}/login"; then
    echo "==> Running. Open http://$(hostname -I | awk '{print $1}'):${APP_PORT:-5000}/login"
    echo "    Logs:    journalctl -u $SERVICE -f"
    echo "    Restart: systemctl restart $SERVICE"
else
    echo "!! Service failed to start. Logs:"
    journalctl -u $SERVICE -n 40 --no-pager
    exit 1
fi
