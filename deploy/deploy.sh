#!/bin/bash
# Lotus Task Manager — install or update from GitHub on the VPS (Ubuntu/Debian, run as root).
#
#   curl -fsSL https://raw.githubusercontent.com/Refaat1942/Lotus-Task-Manager/main/deploy/deploy.sh | sudo bash
#
# Options (environment variables): BRANCH (default main), APP_DIR (default /root/taskmanager)
# Your data (instance/taskmanager.db, static/uploads, instance/secret_key) is backed up and kept.
set -euo pipefail

REPO_URL="${REPO_URL:-https://github.com/Refaat1942/Lotus-Task-Manager.git}"
BRANCH="${BRANCH:-main}"
APP_DIR="${APP_DIR:-/root/taskmanager}"
STAMP="$(date +%Y%m%d-%H%M%S)"
BACKUP_DIR="/root/backups/lotus-$STAMP"

if [ "$(id -u)" -ne 0 ]; then echo "Run as root (sudo)."; exit 1; fi

echo "==> Installing system packages"
export DEBIAN_FRONTEND=noninteractive
apt-get update -qq
apt-get install -y -qq git curl python3 python3-venv python3-pip sqlite3 psmisc >/dev/null

echo "==> Backing up data to $BACKUP_DIR"
mkdir -p "$BACKUP_DIR"
if [ -d "$APP_DIR" ]; then
    [ -d "$APP_DIR/instance" ] && cp -a "$APP_DIR/instance" "$BACKUP_DIR/instance"
    [ -d "$APP_DIR/static/uploads" ] && cp -a "$APP_DIR/static/uploads" "$BACKUP_DIR/uploads"
    # Old installs kept the DB next to app.py
    [ -f "$APP_DIR/taskmanager.db" ] && cp -a "$APP_DIR/taskmanager.db" "$BACKUP_DIR/taskmanager.db.root"
fi

if [ -d "$APP_DIR/.git" ]; then
    echo "==> Updating code ($BRANCH)"
    git -C "$APP_DIR" fetch -q origin "$BRANCH"
    # Force: older versions tracked instance/taskmanager.db, so the live DB counts as a "local change".
    # It is already backed up above and restored below.
    git -C "$APP_DIR" checkout -q -f -B "$BRANCH" "origin/$BRANCH"
    git -C "$APP_DIR" reset -q --hard "origin/$BRANCH"
else
    if [ -d "$APP_DIR" ]; then
        echo "==> Existing non-git install found; moving it to $APP_DIR.old-$STAMP"
        mv "$APP_DIR" "$APP_DIR.old-$STAMP"
    fi
    echo "==> Cloning $REPO_URL ($BRANCH)"
    git clone -q --branch "$BRANCH" "$REPO_URL" "$APP_DIR"
fi

echo "==> Restoring data"
mkdir -p "$APP_DIR/instance" "$APP_DIR/static/uploads"
if [ -d "$BACKUP_DIR/instance" ]; then
    cp -a "$BACKUP_DIR/instance/." "$APP_DIR/instance/"
elif [ -f "$BACKUP_DIR/taskmanager.db.root" ]; then
    cp -a "$BACKUP_DIR/taskmanager.db.root" "$APP_DIR/instance/taskmanager.db"
fi
[ -d "$BACKUP_DIR/uploads" ] && cp -an "$BACKUP_DIR/uploads/." "$APP_DIR/static/uploads/"

bash "$APP_DIR/deploy/update.sh"
