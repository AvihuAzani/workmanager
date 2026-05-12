# WorkManager — מערכת ניהול עבודה לטכנאי סלקום

אפליקציית Web מלאה לטכנאי סלקום — שרת Flask שרץ ב-HomeAssistant Docker.

## תכונות

| טאב | תיאור |
|-----|-------|
| 📋 פקעות | בנק משימות בזמן אמת, מיון לפי מיקום, הצגה על מפה |
| 🏭 מחסן | ניהול מלאי ציוד, החזרות עם תמונות, דוח ציוד |
| 📊 דוחות | היסטוריית ביקורים, חישוב הכנסות PPC/LSB/חריגה, ייצוא Excel |
| 💰 מחירון | הגדרת מחיר לכל סוג משימה |
| 🛡️ מנהל | ניהול משתמשים, הרשאות עם תוקף |

## הרצה מקומית

```bash
pip install flask requests openpyxl
python chat_server.py
# או עם Cloudflare Tunnel:
python run_local.py
```

## הרצה ב-HomeAssistant Docker

```bash
scp chat_server.py root@<HA_IP>:/usr/share/hassio/homeassistant/cellcom/chat_server.py
ssh root@<HA_IP> "docker exec -d homeassistant bash -c 'pip3 install flask requests openpyxl -q && cd /config/cellcom && python3 chat_server.py > /config/cellcom/flask.log 2>&1'"
```

## מבנה קבצים

```
chat_server.py          # שרת Flask ראשי + כל ה-frontend (SPA)
cellcom_bank_pekaot_3.py # שליפת בנק פקעות עצמאית
cellcom_history.py      # שליפת היסטוריית ביקורים ל-SQLite
cellcom_malai.py        # שליפת מלאי ציוד
run_local.py            # הרצה לוקאלית עם Cloudflare Tunnel
start.sh                # הפעלה בתוך container
Procfile                # Production (Railway/Heroku)
requirements.txt        # תלויות Python
```

## טכנולוגיות

- **Backend**: Python Flask, SQLite, openpyxl
- **Frontend**: Vanilla JS, Leaflet.js (מפות), Tesseract.js (OCR)
- **Deployment**: HomeAssistant Docker / Railway / Cloudflare Tunnel
