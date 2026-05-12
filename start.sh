#!/bin/sh
cd /config/cellcom

# התקן pip packages אם חסרים
pip3 install flask requests -q 2>/dev/null

# Flask
if ! pgrep -f chat_server.py > /dev/null 2>&1; then
    nohup python3 chat_server.py > /config/cellcom/flask.log 2>&1 &
    sleep 3
fi

# Cloudflared
if ! pgrep -f cloudflared > /dev/null 2>&1; then
    nohup ./cloudflared-linux-arm64 tunnel --url http://localhost:5000 > /config/cellcom/tunnel.log 2>&1 &
fi
