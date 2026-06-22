#!/bin/bash
# Run on VPS after uploading files to /root/taskmanager
set -e
APP_DIR="/root/taskmanager"
APP_PORT=11000
cd "$APP_DIR"

mkdir -p /root/backups instance static/uploads
cp -a instance/taskmanager.db /root/backups/taskmanager.db.bak 2>/dev/null || true

source venv/bin/activate
pip install -r requirements.txt -q

python -c "
from app import app, db, seed_feature_visibility, APP_VERSION
with app.app_context():
    db.create_all()
    seed_feature_visibility()
    print('Lotus Task Manager v' + APP_VERSION + ' - DB OK')
"

kill $(lsof -t -i:${APP_PORT}) 2>/dev/null || true
sleep 2
nohup python app.py > app.log 2>&1 &
sleep 2
echo "--- Running on port ${APP_PORT} ---"
lsof -i :${APP_PORT} || true
tail -8 app.log
