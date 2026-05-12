#!/bin/sh
cd /config/cellcom

# Create Python venv if it doesn't exist (persists across restarts)
if [ ! -d /config/cellcom/venv ]; then
    python3 -m venv /config/cellcom/venv
fi

# Install packages
/config/cellcom/venv/bin/pip install flask requests openpyxl -q 2>/dev/null
pip3 install flask requests openpyxl -q 2>/dev/null

# הרג ngrok ישן (רץ בטעות על פורט של HA)
pkill -9 -f ngrok 2>/dev/null
kill -9 $(pgrep -f ngrok) 2>/dev/null

# הרג Flask ישן והפעל מחדש
pkill -f chat_server.py 2>/dev/null
sleep 1
# הפעל Flask מתוך ה-container של Home Assistant
docker exec -d homeassistant sh -c 'python3 /config/cellcom/chat_server.py >> /config/cellcom/flask.log 2>&1'
