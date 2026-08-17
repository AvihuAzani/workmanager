# WorkManager — Field Technician Management System

A full-featured web app for Cellcom field technicians — Flask server running on Armbian Linux, accessible remotely via ngrok.

## Features

| Tab | Description |
|-----|-------------|
| 📋 Tasks | Real-time task bank, sorted by location, displayed on a map |
| 🏭 Warehouse | Equipment inventory management, returns with photos, stock report |
| 📊 Reports | Visit history, PPC/LSB/overtime income calculation, Excel export |
| 💰 Pricing | Set price per task type |
| 🛡️ Admin | User management, time-limited permissions |

## Local Setup

```bash
pip install flask requests openpyxl
python chat_server.py
```

## Production Setup (Armbian + ngrok)

```bash
cd /config/cellcom
nohup python3 chat_server.py > /config/cellcom/flask.log 2>&1 &
nohup ngrok http 5050 > /root/ngrok.log 2>&1 &
```

Auto-startup via `/etc/rc.local`. Hourly cron watchdog restarts ngrok if it goes down.

## File Structure

```
chat_server.py # Main Flask server + all frontend (SPA)
cellcom_history.py # Visit history fetcher to SQLite
requirements.txt # Python dependencies
```

## Tech Stack

**Backend**: Python Flask, SQLite, openpyxl
**Frontend**: Vanilla JS, Leaflet.js (maps), Tesseract.js (OCR)
**Deployment**: Armbian Linux / HomeAssistant Docker / Railway
