"""
chat_server.py — שרת Flask מרכזי לאפליקציית WorkManager לטכנאי סלקום
======================================================================

תיאור כללי:
-----------
אפליקציית Web מלאה (Single Page Application) שרצה בתוך Docker Container
של HomeAssistant (או כ-standalone server) ומאפשרת לטכנאי סלקום:

1. בנק פקעות     — שליפה בזמן אמת מ-API של סלקום, הצגת כרטיסי משימה
                   עם פרטים מלאים (כתובת, לקוח, ציוד, תשתית),
                   מיון לפי מיקום ג'יאוגרפי, הצגה על מפה (Leaflet/OpenStreetMap)

2. מחסן           — ניהול מלאי ציוד: הוספה, הפחתה, היסטוריה,
                   דוח ציוד לפי תאריך, החזרות עם תמונות

3. דוחות          — היסטוריית ביקורים ב-SQLite, סיכום לפי סוג משימה,
                   חישוב הכנסות (PPC/LSB/חריגה), ייצוא Excel עם סימון אדום
                   לביקורים שלא הושלמו, הצגה על מפה

4. מחירון         — הגדרת מחיר לכל סוג משימה (כולל PPC/LSB לתשתית סיבים),
                   שמירה ב-JSON לכל משתמש

5. מנהל           — ניהול משתמשים, הענקת הרשאות עם תפוגה לפי חודשים,
                   לוח ניהול מתקפל

6. צ'אט פנימי     — מערכת הודעות פנימית בין משתמשים ומנהל

מבנה האימות:
------------
- OTP בSMS עם Redis/localStorage
- X-Token / X-Phone headers בכל בקשת API
- Fallback לquery params (לצורך ייצוא Excel ב-window.location.href)
- הרשאות per-feature עם תוקף תאריך

קבצי נתונים:
------------
- cellcom_history.db      — SQLite עם היסטוריית ביקורים, משתמשים, החזרות, הודעות
- saved_tokens.json       — טוקני API של משתמשים
- prices/<phone>.json     — מחירונים per-user
- prices/overrides_<phone>.json — חריגות מחיר per-card
- return_photos/          — תמונות החזרת ציוד
- qr_payment.png          — קוד QR לתשלום

הרצה:
-----
  בתוך HomeAssistant Docker:
    pip3 install flask requests openpyxl && python3 chat_server.py

  לוקאלית עם Cloudflare Tunnel:
    python run_local.py

  Production (Railway/Heroku):
    gunicorn chat_server:app (ראה Procfile)

תלויות:
-------
  flask, requests, urllib3, openpyxl, sqlite3 (builtin)
  Frontend: Leaflet.js (מפות), Tesseract.js (OCR)
"""

import os, json, requests, urllib3, sqlite3
from datetime import datetime, date, timedelta
from collections import defaultdict
from io import BytesIO
from flask import Flask, request, jsonify, session, Response, stream_with_context, send_file

urllib3.disable_warnings()
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "cellcom-secret-2026")

TOKENS_FILE         = os.path.join(os.path.dirname(__file__), "saved_tokens.json")
INVENTORY_HIST_FILE = os.path.join(os.path.dirname(__file__), "inventory_history.json")
DB_PATH             = os.path.join(os.path.dirname(__file__), "cellcom_history.db")
PRICES_DIR          = os.path.join(os.path.dirname(__file__), "prices")
RETURNS_PHOTOS_DIR  = os.path.join(os.path.dirname(__file__), "return_photos")
QR_PAYMENT_FILE     = os.path.join(os.path.dirname(__file__), "qr_payment.png")
ADMIN_PHONE         = "0526845629"   # ← מספר הטלפון של המנהל
TRIAL_DAYS          = 7
SUB_DAYS            = 30

# הרשאות ברירת מחדל למשתמש רגיל
DEFAULT_PERMISSIONS = {
    "premium":           False,   # גישה לכל הלשוניות (מלבד פקעות שתמיד גלויה)
    "full_card":         False,   # כרטיס פקעה מלא
    "inventory":         False,   # לשונית מחסן
    "equipment_report":  False,   # תת-לשונית דוח ציוד
    "reports":           False,   # לשונית דוחות
    "prices":            False,   # לשונית מחירון
}

def get_auth():
    """קורא טוקן ופלאפון מה-header / query-param (client-side storage) — fallback לסשן (OTP flow)."""
    token = (request.headers.get('X-Token', '').strip()
             or request.args.get('token', '').strip()
             or session.get('token', ''))
    phone = (request.headers.get('X-Phone', '').strip()
             or request.args.get('phone', '').strip()
             or session.get('phone', ''))
    return (token or None), phone

# ============================================================
# דוחות — SQLite
# ============================================================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS visits (
            phone TEXT DEFAULT '',
            call_id TEXT, fetch_date TEXT, customer_id TEXT,
            contact_name TEXT, contact_phone TEXT, task_type TEXT,
            street TEXT, city TEXT, appt_start TEXT, appt_finish TEXT,
            status TEXT, infrastructure TEXT, call_type TEXT, is_vip TEXT,
            lsb_flag INTEGER DEFAULT 0,
            PRIMARY KEY (phone, call_id, fetch_date)
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS returns (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            phone TEXT,
            serial TEXT,
            description TEXT,
            return_date TEXT,
            photo_filename TEXT,
            created_at TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            phone TEXT PRIMARY KEY,
            name TEXT,
            employee_id TEXT,
            registered_at TEXT,
            trial_expires TEXT,
            subscription_expires TEXT,
            is_admin INTEGER DEFAULT 0,
            permissions TEXT DEFAULT '{}'
        )
    """)
    try:
        conn.execute("ALTER TABLE returns ADD COLUMN phone TEXT DEFAULT ''")
    except:
        pass
    try:
        conn.execute("ALTER TABLE returns ADD COLUMN quantity INTEGER DEFAULT 1")
    except:
        pass
    # migration — הוסף עמודות אם לא קיימות
    try:
        conn.execute("ALTER TABLE visits ADD COLUMN lsb_flag INTEGER DEFAULT 0")
    except:
        pass
    try:
        conn.execute("ALTER TABLE visits ADD COLUMN phone TEXT DEFAULT ''")
    except:
        pass
    conn.execute("""
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            from_phone TEXT NOT NULL,
            to_phone TEXT NOT NULL,
            body TEXT NOT NULL,
            created_at TEXT NOT NULL,
            is_read INTEGER DEFAULT 0
        )
    """)
    try:
        conn.execute("ALTER TABLE users ADD COLUMN permissions TEXT DEFAULT '{}'")
    except:
        pass
    try:
        conn.execute("ALTER TABLE users ADD COLUMN premium_expires TEXT DEFAULT NULL")
    except:
        pass
    conn.commit()
    return conn

def date_already_fetched(conn, target_date, phone=""):
    cur = conn.execute(
        "SELECT COUNT(*) FROM visits WHERE fetch_date=? AND phone=?",
        (target_date.isoformat(), phone)
    )
    return cur.fetchone()[0] > 0

def save_day_to_db(conn, appointments, fetch_date, phone=""):
    date_str = fetch_date.isoformat()
    for apt in appointments:
        task = apt.get("task", {})
        cd   = apt.get("callDetails", {})
        # בדוק אם LSB מופיע בכל שדות האפוינטמנט
        raw_text = json.dumps(apt, ensure_ascii=False).upper()
        lsb_flag = 1 if "LSB" in raw_text else 0
        conn.execute(
            "INSERT OR REPLACE INTO visits(phone,call_id,fetch_date,customer_id,contact_name,"
            "contact_phone,task_type,street,city,appt_start,appt_finish,status,"
            "infrastructure,call_type,is_vip,lsb_flag) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (
                phone,
                task.get("callId",""), date_str,
                task.get("customer",""), task.get("contactName",""),
                task.get("contactPhoneNumber",""), task.get("taskType",""),
                task.get("street",""), task.get("city",""),
                task.get("formattedAppointmentStart",""), task.get("formattedAppointmentFinish",""),
                task.get("status",{}).get("displayString","") if isinstance(task.get("status"),dict) else task.get("status",""),
                cd.get("infrastructure",""), cd.get("callType",""),
                "כן" if task.get("isVIP")=="1" else "לא",
                lsb_flag,
            )
        )
    conn.commit()

def get_visits_for_range(conn, start_date, end_date, phone=""):
    cur = conn.execute(
        "SELECT phone,call_id,fetch_date,customer_id,contact_name,contact_phone,task_type,"
        "street,city,appt_start,appt_finish,status,infrastructure,call_type,is_vip,lsb_flag "
        "FROM visits WHERE fetch_date>=? AND fetch_date<=? AND phone=? "
        "ORDER BY fetch_date DESC, appt_start ASC",
        (start_date.isoformat(), end_date.isoformat(), phone)
    )
    cols = ["phone","call_id","fetch_date","customer_id","contact_name","contact_phone",
            "task_type","street","city","appt_start","appt_finish",
            "status","infrastructure","call_type","is_vip","lsb_flag"]
    return [dict(zip(cols, row)) for row in cur.fetchall()]

def fetch_schedules_api(token, target_date):
    date_str = target_date.strftime("%Y-%m-%d")
    data = cellcom_post(
        "https://tech-api.cellcom.co.il/api/technician/authorize/technicianscedules/getSchedules",
        {"DeviceModel": DEVICE_MODEL, "IsRefresh": True, "PhoneDeviceId": PHONE_DEVICE_ID,
         "AssignmentStart": f"{date_str}T00:00:00",
         "AssignmentFinish": f"{date_str}T23:59:59"}, token
    )
    if data.get("Header", {}).get("ReturnCode") != "00":
        return []
    return data.get("Body", {}).get("appointments", [])

# ============================================================
# מחירון
# ============================================================
def _prices_file(phone):
    os.makedirs(PRICES_DIR, exist_ok=True)
    safe = "".join(c if c.isalnum() else "_" for c in (phone or "default"))
    return os.path.join(PRICES_DIR, f"{safe}.json")

def load_prices(phone=""):
    path = _prices_file(phone)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_prices(prices, phone=""):
    path = _prices_file(phone)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(prices, f, ensure_ascii=False, indent=2)

def get_all_task_types():
    try:
        conn = sqlite3.connect(DB_PATH)
        cur = conn.execute("SELECT DISTINCT task_type FROM visits WHERE task_type != '' ORDER BY task_type")
        types = [row[0] for row in cur.fetchall()]
        conn.close()
        return types
    except:
        return []

# ============================================================
# מנויים / הרשמה
# ============================================================
def db_get_user(phone):
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT phone,name,employee_id,registered_at,trial_expires,subscription_expires,is_admin,permissions,premium_expires "
        "FROM users WHERE phone=?", (phone,)
    ).fetchone()
    conn.close()
    if not row: return None
    u = dict(zip(['phone','name','employee_id','registered_at','trial_expires','subscription_expires','is_admin','permissions','premium_expires'], row))
    # parse permissions JSON
    try:
        u['permissions'] = json.loads(u['permissions'] or '{}')
    except:
        u['permissions'] = {}
    return u

def get_effective_permissions(phone):
    """מחזיר הרשאות אפקטיביות — מנהל מקבל הכל, משתמש מקבל merge של ברירת מחדל + שמור + פרמיום תאריך"""
    if phone == ADMIN_PHONE:
        return {k: True for k in DEFAULT_PERMISSIONS}
    user = db_get_user(phone)
    if not user: return dict(DEFAULT_PERMISSIONS)
    if user.get('is_admin'):
        return {k: True for k in DEFAULT_PERMISSIONS}
    saved = user.get('permissions', {})
    merged = dict(DEFAULT_PERMISSIONS)
    today = date.today()
    for k, v in saved.items():
        if k not in DEFAULT_PERMISSIONS:
            continue
        if v is True:                      # הרשאה קבועה
            merged[k] = True
        elif v is False or v is None:      # כבוי
            merged[k] = False
        elif isinstance(v, str):           # תאריך ISO
            try:
                merged[k] = date.fromisoformat(v) >= today
            except:
                merged[k] = False
    # פרמיום לפי תאריך תפוגה (grant-premium מהצ'אט)
    prem_exp = user.get('premium_expires', '')
    if prem_exp:
        try:
            if date.fromisoformat(prem_exp) >= today:
                merged['premium'] = True
        except:
            pass
    return merged

def db_register_user(phone, name, employee_id):
    today = date.today()
    trial_exp = (today + timedelta(days=TRIAL_DAYS)).isoformat()
    premium_exp = (today + timedelta(days=14)).isoformat()   # פרמיום אוטומטי 14 יום
    is_adm = 1 if phone == ADMIN_PHONE else 0
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT OR IGNORE INTO users (phone,name,employee_id,registered_at,trial_expires,is_admin,premium_expires) VALUES (?,?,?,?,?,?,?)",
        (phone, name, employee_id, today.isoformat(), trial_exp, is_adm, premium_exp)
    )
    conn.commit(); conn.close()

def db_extend_subscription(phone, days=None):
    if days is None: days = SUB_DAYS
    today = date.today()
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute("SELECT subscription_expires FROM users WHERE phone=?", (phone,)).fetchone()
    if row and row[0]:
        try: base = max(date.fromisoformat(row[0]), today)
        except: base = today
    else:
        base = today
    new_exp = (base + timedelta(days=days)).isoformat()
    conn.execute("UPDATE users SET subscription_expires=? WHERE phone=?", (new_exp, phone))
    conn.commit(); conn.close()
    return new_exp

def get_sub_status(phone):
    """Returns (status_str, expires_date_str)
       status: 'active' | 'not_registered'
    """
    if phone == ADMIN_PHONE:
        return 'active', None
    user = db_get_user(phone)
    if not user:
        return 'not_registered', None
    return 'active', None

# ============================================================
# מלאי — היסטוריה
# ============================================================
def load_inventory_history():
    if os.path.exists(INVENTORY_HIST_FILE):
        with open(INVENTORY_HIST_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"last_updated": "", "serials": {}}

def save_inventory_history(hist):
    with open(INVENTORY_HIST_FILE, "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=2)

def process_inventory_snapshot(categories):
    """
    קורא מלאי חדש, משווה עם ההיסטוריה.
    לכל סריאל: אם כבר קיים → שמור תאריך ישן, אם חדש → רשום היום.
    מחזיר רשימת פריטים מועשרת עם first_seen.
    """
    today = date.today().isoformat()
    hist = load_inventory_history()
    serials_hist = hist.get("serials", {})

    result = []
    for cat in categories:
        cat_title = cat.get("title", "")
        # קביעת סטטוס לפי שם הקטגוריה
        if "תקין" in cat_title:
            status = "תקין"
        elif "חסום" in cat_title or "לא תקין" in cat_title:
            status = "חסום"
        else:
            status = cat_title

        for item in cat.get("inventoryItems", []):
            desc    = item.get("title", "")
            catalog = item.get("catalog", "")
            serials = item.get("inventorySerialItems") or []
            amount  = item.get("amount", 0)

            if serials:
                for sn in serials:
                    if sn in serials_hist:
                        first_seen = serials_hist[sn]["first_seen"]
                    else:
                        first_seen = today
                    # עדכן היסטוריה
                    serials_hist[sn] = {
                        "first_seen": first_seen,
                        "description": desc,
                        "catalog": catalog,
                        "category": cat_title,
                        "status": status,
                    }
                    result.append({
                        "serial": sn,
                        "description": desc,
                        "catalog": catalog,
                        "category": cat_title,
                        "status": status,
                        "first_seen": first_seen,
                        "days_held": (date.today() - date.fromisoformat(first_seen)).days,
                    })
            else:
                # פריט ללא סריאל — מעקב לפי מק"ט+תיאור
                key = f"NO_SN__{catalog}__{desc}"
                if key in serials_hist:
                    first_seen = serials_hist[key]["first_seen"]
                else:
                    first_seen = today
                serials_hist[key] = {
                    "first_seen": first_seen,
                    "description": desc,
                    "catalog": catalog,
                    "category": cat_title,
                    "status": status,
                }
                result.append({
                    "serial": "",
                    "description": desc,
                    "catalog": catalog,
                    "category": cat_title,
                    "status": status,
                    "first_seen": first_seen,
                    "days_held": (date.today() - date.fromisoformat(first_seen)).days,
                    "amount": amount,
                })

    hist["last_updated"] = today
    hist["serials"] = serials_hist
    save_inventory_history(hist)
    return result

# ============================================================
# הגדרות
# ============================================================
PHONE_DEVICE_ID = "FF210C6E-5313-4961-846D-229DF3FAC0FC"
DEVICE_MODEL    = "ios_Apple_iPhone 16_SysVer_26.3.1_appVer_23.03.26.1P"
CLIENT_ID       = "354193a2-8d29-11ea-bc55-0242ac130004"
CLIENT_SECRET   = "354193a2-8d29-11ea-bc55-0242ac130003"

BASE_HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "he-IL,he;q=0.9",
    "Content-Type": "application/json",
    "format": "application/json",
    "phonedeviceid": PHONE_DEVICE_ID,
    "devicemodel": DEVICE_MODEL,
    "User-Agent": "HomeTechAppClient/23.03.26 CFNetwork/3860.400.51 Darwin/25.3.0",
}

INFRA_MAP = {
    "FB": "סיבים IBC", "BF": "סיבים בזק", "BN": "נחושת בזק",
    "HO": "HOT", "NV": "סיבים NV", "IB": "IBC",
}

# ============================================================
# Cellcom API
# ============================================================
def cellcom_post(url, body, token=None):
    headers = {**BASE_HEADERS}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    else:
        headers["Authorization"] = "default"
    r = requests.post(url, headers=headers, json=body, verify=False, timeout=15)
    return r.json()

def login_step1(phone, employee_id):
    data = cellcom_post("https://tech-api.cellcom.co.il/api/technician/loginStep1", {
        "DeviceModel": DEVICE_MODEL, "PhoneNumber": phone, "EmployeeId": employee_id,
        "ClientId": CLIENT_ID, "LoginType": "EMPLOYEEPHONE", "Scope": "USERNAME",
        "SId": "ios_Apple_iPhone 16_SysVer_26.3.1", "PhoneDeviceId": PHONE_DEVICE_ID,
    })
    body = data.get("Body", {})
    if not body.get("isSuccess"):
        raise Exception("שגיאה בשליחת SMS")
    return body["ticketId"]

def login_step2(phone, employee_id, otp, ticket_id):
    data = cellcom_post("https://tech-api.cellcom.co.il/api/technician/loginStep2", {
        "PhoneNumber": phone, "EmployeeId": employee_id,
        "ClientId": CLIENT_ID, "ClientSecret": CLIENT_SECRET,
        "LoginType": "EMPLOYEEPHONE", "Scope": "USERNAME",
        "SId": "ios_Apple_iPhone 16_SysVer_26.3.1",
        "OtpCode": otp, "OtpGuid": ticket_id, "PhoneDeviceId": PHONE_DEVICE_ID,
    })
    rc = data.get("Header", {}).get("ReturnCode")
    if rc != "0":
        raise Exception(data.get("Header", {}).get("ReturnCodeMessage", "קוד שגוי"))
    return data["Body"]["access_token"]

def fetch_inventory(token):
    data = cellcom_post(
        "https://tech-api.cellcom.co.il/api/technician/authorize/Inventory/GetMyInventory",
        {"DeviceModel": DEVICE_MODEL, "IsRefresh": True, "PhoneDeviceId": PHONE_DEVICE_ID},
        token
    )
    if data.get("Header", {}).get("ReturnCode") != "00":
        return []
    return data.get("Body", {}).get("result", {}).get("myInventoryCategories", [])

def fetch_tasks(token):
    data = cellcom_post(
        "https://tech-api.cellcom.co.il/api/technician/authorize/callActivities/getPotentialTasks",
        {"DeviceModel": DEVICE_MODEL, "PhoneDeviceId": PHONE_DEVICE_ID}, token
    )
    if data.get("Header", {}).get("ReturnCode") != "00":
        return []
    rd = data.get("Body", {}).get("ResponseData", "{}")
    if isinstance(rd, str):
        rd = json.loads(rd)
    return rd.get("Tasks", {}).get("Task", [])

def fetch_call_details(token, call_id, customer_id, source):
    data = cellcom_post(
        "https://tech-api.cellcom.co.il/api/technician/authorize/callDetails/getCallDetails",
        {"DeviceModel": DEVICE_MODEL, "CallId": call_id, "CrmCustomerId": customer_id,
         "SourceSystem": source, "IsRefresh": False, "PhoneDeviceId": PHONE_DEVICE_ID}, token
    )
    if data.get("Header", {}).get("ReturnCode") != "00":
        return {}
    return data.get("Body", {})

def parse_equipment(body):
    infra_code = infra_name = technology = ""
    planned, existing, tv = [], [], []
    spm = body.get("supplyProductServicesModel", {})
    technology = spm.get("generalInfo", {}).get("technologyType", "")
    for service in spm.get("services", []):
        if service.get("serviceType") == "INTERNET":
            ic = service.get("infrastructureModel", {}).get("infraComOperator", "") or ""
            infra_code = ic
            infra_name = INFRA_MAP.get(ic, ic)
        elif service.get("serviceType") == "TV":
            for kit in service.get("equipmentList", []):
                for item in kit.get("serialItems", []):
                    if item.get("productStatus") == "Installed":
                        d = item.get("productDescription", "")
                        s = item.get("serialNumber", "")
                        if d: tv.append(f"{d} ({s})" if s else d)
    for item in spm.get("allVisitSerialEquipment", []):
        status = item.get("productStatus", "")
        desc = item.get("productDescription", "") or item.get("originProductDescription", "")
        actions = item.get("availableEquipmentActions", [])
        if "Install" in actions or status in ["SupplyProcess", "Add"]:
            if desc: planned.append(desc)
        elif status == "Installed" and desc:
            existing.append(desc)
    return {
        "infra_name": infra_name, "technology": technology,
        "planned": " | ".join(planned),
        "existing": " | ".join(existing),
        "tv": " | ".join(tv),
    }

def fetch_visit_history(token, ban, user_id, user_type="JET"):
    try:
        data = cellcom_post(
            "https://tech-api.cellcom.co.il/api/technician/authorize/TechnicianVisitsHistory/GetTechnicianVisitsHistory",
            {"DeviceModel": DEVICE_MODEL, "PhoneDeviceId": PHONE_DEVICE_ID,
             "ban": ban, "userId": user_id, "userType": user_type}, token
        )
        if data.get("Header", {}).get("ReturnCode") != "00":
            return []
        return data.get("Body", {}).get("visitsHistory", [])
    except:
        return []

REPEAT_TASK_TYPES = {
    "שימור לקוח",
    "לא צופה ולא גולש",
    "לקוח מושבת ללא אינטרנט",
    "תקלת גלישה תשתית סיבים",
    "תקלה בקו טלפון",
}

def check_recent_visit(visits, current_task_type="", days=30):
    """מחזיר שם טכנאי מהביקור האחרון אם יש חזורת, אחרת None."""
    # בדוק שסוג המשימה הנוכחי רלוונטי לחזורת
    tt = (current_task_type or "").strip()
    if not any(rt in tt for rt in REPEAT_TASK_TYPES):
        return None
    from datetime import timedelta
    cutoff = datetime.now() - timedelta(days=days)
    for v in visits:
        raw = v.get("fullDateTimeTo", "") or v.get("dateEnd", "")
        try:
            dt = datetime.strptime(raw[:16], "%d/%m/%Y %H:%M") if " " in raw else datetime.strptime(raw, "%d/%m/%Y")
            if dt >= cutoff:
                tech = v.get("technicianName", "") or ""
                return tech if tech else "לא ידוע"
        except:
            continue
    return None

def format_visit_history(visits):
    if not visits:
        return ""
    lines = []
    cutoff = datetime.now().replace(hour=0, minute=0, second=0)
    from datetime import timedelta
    cutoff -= timedelta(days=90)
    for v in visits[:5]:
        raw = v.get("fullDateTimeTo", "") or v.get("dateEnd", "")
        try:
            dt = datetime.strptime(raw[:16], "%d/%m/%Y %H:%M") if " " in raw else datetime.strptime(raw, "%d/%m/%Y")
        except:
            dt = None
        date_str = v.get("dateEnd", "")
        tech = v.get("technicianName", "")
        status = v.get("status", "")
        vtype = v.get("visitType", "")
        reasons = " / ".join(
            i.get("visitCloseReason", "") for i in v.get("visitInfo", [])
            if i.get("visitCloseReason")
        )
        parts = [p for p in [date_str, vtype, status, tech, reasons] if p]
        lines.append("  • " + " | ".join(parts))
    return "\n".join(lines)

def build_single_card(task, token, index):
    try:
        start_dt = datetime.fromisoformat(task.get("start_date", "").replace("+03:00", ""))
        end_dt   = datetime.fromisoformat(task.get("end_date", "").replace("+03:00", ""))
        time_str = f"{start_dt.strftime('%H:%M')}–{end_dt.strftime('%H:%M')}"
        date_str = start_dt.strftime("%d/%m/%Y")
    except:
        time_str = date_str = ""
    street  = task.get("street", "")
    home_no = task.get("home_no", "")
    apt     = task.get("apartment_no", "")
    address = f"{street} {home_no}" + (f" דירה {apt}" if apt and apt != "1" else "")
    city    = task.get("city", "")
    district = task.get("district", "")
    cid     = task.get("customer_id", "")
    source  = task.get("system_source", "JET")
    body    = fetch_call_details(token, task.get("call_id", ""), cid, source)
    eq      = parse_equipment(body) if body else {}
    visits  = fetch_visit_history(token, task.get("ban", cid), task.get("user_id", cid), source)
    return {
        "index": index,
        "call_id":    task.get("call_id", ""),
        "date":       date_str, "time": time_str,
        "name":       task.get("contact_name", ""),
        "customer_id": cid,
        "phone":      task.get("contact_phone", ""),
        "address":    f"{address}, {city}" + (f" ({district})" if district else ""),
        "city":       city,
        "task_type":  task.get("task_type", ""),
        "infra":      eq.get("infra_name", ""),
        "technology": eq.get("technology", ""),
        "planned":    eq.get("planned", ""),
        "existing":   eq.get("existing", ""),
        "tv":         eq.get("tv", ""),
        "comment":      task.get("comment_text", "") or "",
        "history":      format_visit_history(visits),
        "recent_visit": check_recent_visit(visits, task.get("task_type", "")),
    }

# ============================================================
# Message Handler
# ============================================================
def handle(text):
    text = text.strip()

    # שלב הגדרה
    if session.get("step") == "setup_phone":
        if not text.replace("-", "").isdigit():
            return "❌ מספר לא תקין. נסה שוב:"
        session["phone"] = text.replace("-", "")
        session["step"] = "setup_employee"
        return "מה מספר העובד שלך?"

    if session.get("step") == "setup_employee":
        if not text.isdigit():
            return "❌ מספר עובד לא תקין. נסה שוב:"
        session["employee_id"] = text
        session["step"] = "idle"
        # בדוק אם יש טוקן שמור תקף
        saved = get_valid_token(session["phone"])
        if saved:
            session["token"] = saved
            return "✅ נרשמת! נמצא טוקן שמור מהיום — מחובר אוטומטית 🔓\nשלח *פקעות* להתחיל."
        return "✅ נרשמת!\nשלח *טוקן* להתחברות."

    # OTP
    if session.get("step") == "awaiting_otp" and text.isdigit() and 4 <= len(text) <= 8:
        try:
            token = login_step2(session["phone"], session["employee_id"], text, session["ticket_id"])
            session["token"] = token
            session["step"] = "idle"
            save_token_for_phone(session["phone"], token)
            return "✅ מחובר! הטוקן נשמר עד 23:59 הלילה.\nשלח *פקעות* להתחיל."
        except Exception as e:
            return f"❌ {e}\nשלח *טוקן* לקוד חדש."

    # בכניסה חדשה — נסה לטעון טוקן שמור
    if session.get("phone") and not session.get("token"):
        saved = get_valid_token(session["phone"])
        if saved:
            session["token"] = saved

    if text in ("טוקן", "התחבר"):
        if not session.get("phone"):
            session["step"] = "setup_phone"
            return "👋 ברוך הבא!\nמה מספר הטלפון שלך? (למשל 0526845629)"
        # בדוק טוקן שמור תקף
        saved = get_valid_token(session["phone"])
        if saved:
            session["token"] = saved
            return "🔓 נמצא טוקן תקף מהיום — מחובר!\nשלח *פקעות* להתחיל."
        try:
            tid = login_step1(session["phone"], session["employee_id"])
            session["ticket_id"] = tid
            session["step"] = "awaiting_otp"
            return "📱 נשלח SMS — שלח את הקוד:"
        except Exception as e:
            return f"❌ {e}"

    if text in ("עזרה", "help"):
        return (
            "• *פקעות* — כל הפקעות\n"
            "• *פקעות [מחוז]* — לפי מחוז\n"
            "• *עזרה* — תפריט זה"
        )

    if text == "/reset":
        session.clear()
        return "🔄 אופס. שלח *טוקן* להתחלה."

    return "שלח *עזרה* לרשימת פקודות."

# ============================================================
# Routes
# ============================================================
@app.route("/logo")
def serve_logo():
    logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
    if os.path.exists(logo_path):
        return send_file(logo_path, mimetype="image/png")
    return "", 404

@app.route("/")
def index():
    return HTML.replace('__ADMIN_PHONE__', ADMIN_PHONE)

@app.route("/api/login", methods=["POST"])
def api_login():
    data = request.json
    phone = data.get("phone", "").replace("-", "").strip()
    employee_id = data.get("employee_id", "").strip()
    # קבל טוקן מה-client אם נשלח (ניסיון התחברות שקטה)
    client_token = data.get("token", "").strip()
    if not phone or not employee_id:
        return jsonify({"status": "error", "message": "מלא את כל השדות"})
    # שמור בסשן רק למהלך OTP
    session["phone"] = phone
    session["employee_id"] = employee_id
    # אם הלקוח שלח טוקן — תמיד תקין (הוא שמר אותו)
    if client_token:
        return jsonify({"status": "ok", "token": client_token, "phone": phone})
    # שלח SMS
    try:
        ticket_id = login_step1(phone, employee_id)
        session["ticket_id"] = ticket_id
        return jsonify({"status": "otp"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

@app.route("/api/verify", methods=["POST"])
def api_verify():
    otp = request.json.get("otp", "").strip()
    phone = session.get("phone")
    employee_id = session.get("employee_id")
    ticket_id = session.get("ticket_id")
    if not all([phone, employee_id, ticket_id, otp]):
        return jsonify({"status": "error", "message": "חסר מידע"})
    try:
        token = login_step2(phone, employee_id, otp, ticket_id)
        # מחזיר טוקן ללקוח — הוא ישמור ב-localStorage
        return jsonify({"status": "ok", "token": token, "phone": phone})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

@app.route("/api/logout", methods=["POST"])
def api_logout():
    session.clear()
    return jsonify({"status": "ok"})

@app.route("/api/status")
def api_status():
    token, phone = get_auth()
    if token and phone:
        sub_status, sub_expires = get_sub_status(phone)
        user = db_get_user(phone)
        is_adm = (phone == ADMIN_PHONE) or bool(user and user.get('is_admin'))
        perms = get_effective_permissions(phone)
        return jsonify({
            "logged_in": True,
            "phone": phone,
            "sub_status": sub_status,
            "sub_expires": sub_expires,
            "is_admin": is_adm,
            "user_name": user.get('name', '') if user else '',
            "permissions": perms
        })
    return jsonify({"logged_in": False})

@app.route("/chat", methods=["POST"])
def chat():
    token, phone = get_auth()
    if not token:
        return jsonify({"type": "text", "text": "⚠️ לא מחובר. רענן את הדף."})
    msg = request.json.get("message", "")
    response = handle(msg)
    if isinstance(response, dict):
        return jsonify(response)
    return jsonify({"type": "text", "text": response})

@app.route("/api/tasks")
def api_tasks():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    district = request.args.get("district", "")
    tasks = fetch_tasks(token)
    if district:
        tasks = [t for t in tasks if district in t.get("district", "")]
    return jsonify({"tasks": tasks, "count": len(tasks)})

@app.route("/api/prices", methods=["GET"])
def api_get_prices():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    prices = load_prices(phone)
    task_types = get_all_task_types()
    return jsonify({"prices": prices, "task_types": task_types})

@app.route("/api/prices", methods=["POST"])
def api_save_prices():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    prices = request.json.get("prices", {})
    save_prices(prices, phone)
    return jsonify({"status": "ok"})

def _overrides_file(phone):
    os.makedirs(PRICES_DIR, exist_ok=True)
    safe = "".join(c if c.isalnum() else "_" for c in (phone or "default"))
    return os.path.join(PRICES_DIR, f"overrides_{safe}.json")

def load_overrides(phone):
    path = _overrides_file(phone)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_overrides(overrides, phone):
    path = _overrides_file(phone)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(overrides, f, ensure_ascii=False, indent=2)

@app.route("/api/reports/set-override", methods=["POST"])
def api_set_override():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    data = request.json
    call_id   = data.get("call_id", "").strip()
    fetch_date = data.get("fetch_date", "").strip()
    price_type = data.get("price_type", "")      # "PPC" / "LSB" / "custom"
    custom_price = data.get("custom_price", None) # float or null
    if not call_id or not fetch_date:
        return jsonify({"error": "חסר מידע"}), 400
    key = f"{call_id}__{fetch_date}"
    overrides = load_overrides(phone)
    overrides[key] = {"price_type": price_type, "custom_price": custom_price}
    save_overrides(overrides, phone)
    return jsonify({"status": "ok"})

@app.route("/api/reports")
def api_reports():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    days = int(request.args.get("days", 30))
    end_str = request.args.get("end", "")
    end_date   = date.fromisoformat(end_str) if end_str else date.today()
    start_date = end_date - timedelta(days=days-1)
    conn = init_db()
    # שלוף מה-API תאריכים שאין במאגר
    fetched_new = 0
    cur = start_date
    today = date.today()
    while cur <= end_date:
        # תמיד רענן את היום הנוכחי (הלוז יכול להשתנות)
        if cur == today:
            conn.execute("DELETE FROM visits WHERE fetch_date=? AND phone=?", (cur.isoformat(), phone))
            conn.commit()
        if cur == today or not date_already_fetched(conn, cur, phone):
            apts = fetch_schedules_api(token, cur)
            if apts:
                save_day_to_db(conn, apts, cur, phone)
                fetched_new += len(apts)
            else:
                # שמור רשומה ריקה כדי לא לשלוף שוב
                conn.execute("INSERT OR IGNORE INTO visits(phone,call_id,fetch_date) VALUES(?,?,?)",
                             (phone, f"EMPTY_{cur.isoformat()}", cur.isoformat()))
                conn.commit()
        cur += timedelta(days=1)
    visits = get_visits_for_range(conn, start_date, end_date, phone)
    # סנן רשומות ריקות
    visits = [v for v in visits if not v["call_id"].startswith("EMPTY_")]
    conn.close()
    prices = load_prices(phone)
    overrides = load_overrides(phone)
    # הוסף מחיר לכל ביקור
    total_earnings = 0
    summary = defaultdict(lambda: {"count": 0, "earnings": 0})
    for v in visits:
        tt = v["task_type"] or ""
        is_incomplete = "לא הושלם" in (v.get("status") or "") or "לא הושלמ" in (v.get("status") or "")
        v["is_incomplete"] = is_incomplete
        ov_key = f"{v['call_id']}__{v['fetch_date']}"
        ov = overrides.get(ov_key, {})
        if is_incomplete:
            price = 0.0
            v["price_type"] = ""
            v["custom_price"] = None
        elif ov.get("price_type") == "custom" and ov.get("custom_price") is not None:
            price = float(ov["custom_price"])
            v["price_type"] = "custom"
            v["custom_price"] = price
        elif ("תשתית סיבים" in tt or "טריפל סיבים" in tt) and "תקלת גלישה" not in tt:
            forced = ov.get("price_type")
            if forced == "LSB" or (not forced and v.get("lsb_flag") == 1):
                price = float(prices.get(tt + "__LSB", 0))
                v["price_type"] = "LSB"
            else:
                price = float(prices.get(tt + "__PPC", 0))
                v["price_type"] = "PPC"
            v["custom_price"] = None
        else:
            price = float(prices.get(tt, 0))
            v["price_type"] = ""
            v["custom_price"] = None
        v["price"] = price
        total_earnings += price
        if tt:
            summary[tt]["count"] += 1
            summary[tt]["earnings"] += price
    summary_sorted = dict(sorted(summary.items(), key=lambda x: -x[1]["count"]))
    return jsonify({
        "visits": visits,
        "total": len(visits),
        "total_earnings": total_earnings,
        "summary": summary_sorted,
        "prices_map": prices,
        "start_date": start_date.isoformat(),
        "end_date": end_date.isoformat(),
        "fetched_new": fetched_new,
    })

@app.route("/api/geocode")
def api_geocode():
    addr = request.args.get("q", "").strip()
    if not addr:
        return jsonify({"error": "no address"}), 400
    conn = init_db()
    conn.execute("CREATE TABLE IF NOT EXISTS geocache (addr TEXT PRIMARY KEY, lat REAL, lng REAL)")
    row = conn.execute("SELECT lat,lng FROM geocache WHERE addr=?", (addr,)).fetchone()
    if row:
        conn.close()
        return jsonify({"lat": row[0], "lng": row[1]})
    try:
        import urllib.parse, urllib.request as ureq
        url = "https://nominatim.openstreetmap.org/search?q=" + urllib.parse.quote(addr + ", ישראל") + "&format=json&limit=1"
        req = ureq.Request(url, headers={"User-Agent": "CellcomApp/1.0"})
        with ureq.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read())
        if data:
            lat, lng = float(data[0]["lat"]), float(data[0]["lon"])
            conn.execute("INSERT OR REPLACE INTO geocache(addr,lat,lng) VALUES(?,?,?)", (addr, lat, lng))
            conn.commit()
            conn.close()
            return jsonify({"lat": lat, "lng": lng})
    except Exception as e:
        pass
    conn.close()
    return jsonify({"error": "not found"}), 404

@app.route("/api/reports/export")
def api_reports_export():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    days    = int(request.args.get("days", 30))
    end_str = request.args.get("end", "")
    end_date   = date.fromisoformat(end_str) if end_str else date.today()
    start_date = end_date - timedelta(days=days - 1)
    conn   = init_db()
    visits = get_visits_for_range(conn, start_date, end_date, phone)
    visits = [v for v in visits if not v["call_id"].startswith("EMPTY_")]
    conn.close()
    prices = load_prices(phone)
    for v in visits:
        tt = v["task_type"] or ""
        is_incomplete = "לא הושלם" in (v.get("status") or "") or "לא הושלמ" in (v.get("status") or "")
        v["is_incomplete"] = is_incomplete
        if is_incomplete:
            v["price"] = 0.0
            v["price_type"] = ""
        elif ("תשתית סיבים" in tt or "טריפל סיבים" in tt) and "תקלת גלישה" not in tt:
            if v.get("lsb_flag") == 1:
                v["price"] = float(prices.get(tt + "__LSB", 0))
                v["price_type"] = "LSB"
            else:
                v["price"] = float(prices.get(tt + "__PPC", 0))
                v["price_type"] = "PPC"
        else:
            v["price"] = float(prices.get(tt, 0))
            v["price_type"] = ""
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        import subprocess
        subprocess.run(
            [os.path.join(os.path.dirname(__file__), "venv", "bin", "pip"),
             "install", "openpyxl", "-q"], check=False)
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment

    # שם מתקין
    installer_phone = phone
    installer_user  = db_get_user(installer_phone)
    installer_name  = (installer_user.get("name") or installer_phone) if installer_user else installer_phone

    BLUE       = "4472C4"
    WHITE      = "FFFFFF"
    HDR_FONT   = Font(color=WHITE, bold=True, size=10)
    HDR_FILL   = PatternFill(start_color=BLUE, end_color=BLUE, fill_type="solid")
    CENTER     = Alignment(horizontal="center", vertical="center", wrap_text=True)
    RIGHT      = Alignment(horizontal="right", vertical="center")
    CENTER_MID = Alignment(horizontal="center", vertical="center")
    HEBREW_MONTHS = {1:"ינואר",2:"פברואר",3:"מרץ",4:"אפריל",5:"מאי",6:"יוני",
                     7:"יולי",8:"אוגוסט",9:"ספטמבר",10:"אוקטובר",11:"נובמבר",12:"דצמבר"}

    headers = [
        "תאריך", "שם טכנאי", "מספר לקוח PSID", "מספר פקע",
        "הערות טכנאי",
        "טריפל", "טי וי בלבד", "התקנת אפליקציה",
        'התקנת שוס/מפא או מפא+שוס', "שינוי שרות",
        "פקע קלה (התקנת סופר WIFI)", "פקע קלה (שדרוג נתב)",
        "קריאת שרות מכל סוג", "הסבות שקטות", "אזור מיוחד",
        "התקנת PPC", "התקנת בניין נמוך LSB",
        "התקנת VILA טריפל", 'התקנת VILA ש"וס',
        "שם מתקין", 'סה"כ תשלום למתקין'
    ]
    col_widths = [
        12, 14, 16, 12,
        22,
        7, 8, 10,
        16, 10,
        20, 16,
        14, 12, 10,
        10, 14,
        12, 12,
        12, 12
    ]

    # מיפוי סוג משימה → דגלים בינאריים (עמודות 14–27)
    def task_type_to_flags(tt, price_type):
        flags = [None] * 14  # cols 14-27
        t = tt or ''
        is_vila = 'VILA' in t or 'וילה' in t
        if is_vila:
            if 'שוס' in t or 'מפא' in t or "ש\"וס" in t or "מפ\"א" in t or "שו\"ס" in t or "מפ\"א" in t:
                flags[13] = 1  # col 27: VILA שוס
            else:
                flags[12] = 1  # col 26: VILA טריפל
        elif 'שקט' in t or 'הסב' in t:
            flags[8] = 1   # col 22: הסבות שקטות
        elif 'מיוחד' in t:
            flags[9] = 1   # col 23: אזור מיוחד
        elif 'שינוי' in t:
            flags[4] = 1   # col 18: שינוי שרות
        elif 'אפליקציה' in t or 'app' in t.lower():
            flags[2] = 1   # col 16: התקנת אפליקציה
        elif ('WIFI' in t.upper() or 'ווייפיי' in t) and ('סופר' in t or 'קלה' in t):
            flags[5] = 1   # col 19: פקע קלה - סופר WIFI
        elif 'נתב' in t or ('שדרוג' in t and 'טריפל' not in t):
            flags[6] = 1   # col 20: פקע קלה - שדרוג נתב
        elif 'טריפל' in t:
            flags[0] = 1   # col 14: טריפל
        elif 'TV' in t or 'טלויזיה' in t or 'טלוויזיה' in t or 'tv' in t.lower():
            flags[1] = 1   # col 15: TV בלבד
        elif 'שוס' in t or 'מפא' in t or "ש\"וס" in t or "מפ\"א" in t:
            flags[3] = 1   # col 17: שוס/מפא
        else:
            flags[7] = 1   # col 21: קריאת שרות (ברירת מחדל)
        # PPC / LSB כסימון נוסף
        if price_type == 'PPC':
            flags[10] = 1  # col 24: PPC
        elif price_type == 'LSB':
            flags[11] = 1  # col 25: LSB
        return flags

    RED_FILL = PatternFill(start_color="FFD0D0", end_color="FFD0D0", fill_type="solid")

    def write_sheet(ws, rows_with_flags):
        ws.sheet_view.rightToLeft = True
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.fill = HDR_FILL
            cell.font = HDR_FONT
            cell.alignment = CENTER
        ws.row_dimensions[1].height = 36
        total_price = 0.0
        for ri, (row_data, incomplete) in enumerate(rows_with_flags, 2):
            if incomplete:
                rf = RED_FILL
            else:
                fc = "EEF4FF" if ri % 2 == 0 else "FFFFFF"
                rf = PatternFill(start_color=fc, end_color=fc, fill_type="solid")
            for col, val in enumerate(row_data, 1):
                cell = ws.cell(row=ri, column=col, value=val)
                cell.fill = rf
                # דגלים בינאריים (cols 6-19) ממורכזים
                if 6 <= col <= 19:
                    cell.alignment = CENTER_MID
                else:
                    cell.alignment = RIGHT
            total_price += float(row_data[20] or 0)
        for col, w in enumerate(col_widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w
        sr = len(rows_with_flags) + 3
        ws.cell(row=sr, column=1,
                value=f'סה"כ: {len(rows_with_flags)} ביקורים').font = Font(bold=True, size=12)
        if total_price > 0:
            ws.cell(row=sr, column=21, value=total_price).font = Font(bold=True, size=12)

    def make_row(v):
        flags = task_type_to_flags(v.get("task_type", ""), v.get("price_type", ""))
        row = [
            v["fetch_date"],        # 1. תאריך
            installer_name,          # 2. שם טכנאי
            v["call_id"],            # 3. מספר לקוח PSID
            v["contact_phone"],      # 4. מספר פקע
            v["contact_name"],       # 5. הערות טכנאי (שם לקוח)
        ] + flags + [               # 6-19: דגלים בינאריים
            installer_name,          # 20: שם מתקין
            v["price"] if v.get("price") else None,  # 21: סה"כ תשלום
        ]
        return row, v.get("is_incomplete", False)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # כל התקופה
    all_rows = [make_row(v) for v in visits]
    ws_all = wb.create_sheet("כל התקופה")
    write_sheet(ws_all, all_rows)

    # גיליון לכל חודש
    monthly = defaultdict(list)
    for v in visits:
        try:
            d_obj = date.fromisoformat(v["fetch_date"])
            monthly[(d_obj.year, d_obj.month)].append(make_row(v))
        except:
            pass
    for (yr, mo), m_rows in sorted(monthly.items(), reverse=True):
        ws_m = wb.create_sheet(f"{HEBREW_MONTHS[mo]} {yr}")
        write_sheet(ws_m, m_rows)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"cellcom_{start_date}_{end_date}.xlsx"
    return send_file(buf,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name=filename)

@app.route("/api/inventory")
def api_inventory():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    categories = fetch_inventory(token)
    if not categories:
        return jsonify({"items": [], "last_updated": "", "total": 0})
    items = process_inventory_snapshot(categories)
    hist  = load_inventory_history()
    return jsonify({
        "items": items,
        "last_updated": hist.get("last_updated", ""),
        "total": len(items),
        "received": hist.get("received", {}),
    })

@app.route("/api/inventory/received", methods=["POST"])
def api_inventory_received():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    data   = request.get_json() or {}
    serial = data.get("serial", "")
    status = data.get("status", "received")
    desc   = data.get("desc", "")
    if not serial:
        return jsonify({"error": "חסר serial"}), 400
    hist = load_inventory_history()
    if "received" not in hist:
        hist["received"] = {}
    hist["received"][serial] = {"status": status, "date": date.today().isoformat(), "desc": desc}
    save_inventory_history(hist)
    return jsonify({"ok": True})

@app.route("/api/save-return", methods=["POST"])
def api_save_return():
    import base64 as b64mod
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    data = request.get_json() or {}
    serial      = data.get("serial", "")
    description = data.get("description", "")
    photo_b64   = data.get("photo", "")
    quantity    = int(data.get("quantity", 1) or 1)
    now = datetime.now()
    photo_filename = ""
    if photo_b64:
        try:
            os.makedirs(RETURNS_PHOTOS_DIR, exist_ok=True)
            safe_serial = "".join(c if c.isalnum() else "_" for c in serial)
            photo_filename = f"{safe_serial}_{now.strftime('%Y%m%d_%H%M%S')}.jpg"
            raw = b64mod.b64decode(photo_b64.split(",")[-1])
            with open(os.path.join(RETURNS_PHOTOS_DIR, photo_filename), "wb") as f:
                f.write(raw)
        except Exception:
            photo_filename = ""
    conn = init_db()
    conn.execute(
        "INSERT INTO returns (phone, serial, description, return_date, photo_filename, created_at, quantity) VALUES (?,?,?,?,?,?,?)",
        (phone, serial, description, now.date().isoformat(), photo_filename, now.isoformat(), quantity)
    )
    conn.commit()
    conn.close()
    return jsonify({"ok": True})

@app.route("/api/returns")
def api_get_returns():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    days = int(request.args.get("days", 30))
    since = (date.today() - timedelta(days=days)).isoformat()
    conn = init_db()
    rows = conn.execute(
        "SELECT id, serial, description, return_date, photo_filename, created_at, quantity FROM returns WHERE return_date >= ? AND phone=? ORDER BY created_at DESC",
        (since, phone)
    ).fetchall()
    conn.close()
    return jsonify({"returns": [
        {"id": r[0], "serial": r[1], "description": r[2],
         "return_date": r[3], "has_photo": bool(r[4]), "photo_filename": r[4],
         "quantity": r[6] or 1}
        for r in rows
    ]})

@app.route("/api/equipment-report")
def api_equipment_report():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    # קבלות
    hist = load_inventory_history()
    received_by_date = {}
    for serial, entry in hist.get("received", {}).items():
        if isinstance(entry, dict) and entry.get("status") == "received":
            d = entry.get("date", "")
            received_by_date.setdefault(d, []).append({
                "serial": serial, "desc": entry.get("desc", "")
            })
    # החזרות
    conn = init_db()
    rows = conn.execute(
        "SELECT serial, description, return_date, quantity FROM returns WHERE phone=? ORDER BY return_date DESC, rowid DESC",
        (phone,)
    ).fetchall()
    conn.close()
    returned_by_date = {}
    for r in rows:
        returned_by_date.setdefault(r[2], []).append({
            "serial": r[0] or "", "desc": r[1] or "", "quantity": r[3] or 1
        })
    return jsonify({"received_by_date": received_by_date, "returned_by_date": returned_by_date})

@app.route("/api/equipment-report/export")
def api_equipment_report_export():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    report_date = request.args.get("date", "")
    report_type = request.args.get("type", "received")
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        import subprocess; subprocess.run(["pip","install","openpyxl","-q"],check=False)
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
    wb = openpyxl.Workbook(); ws = wb.active
    ws.sheet_view.rightToLeft = True
    BLUE = "4472C4"; WHITE = "FFFFFF"
    hdr_font = Font(color=WHITE, bold=True, size=11)
    hdr_fill = PatternFill(start_color=BLUE, end_color=BLUE, fill_type="solid")
    ctr  = Alignment(horizontal="center", vertical="center")
    rght = Alignment(horizontal="right",  vertical="center")
    def hdr(ws, cols):
        for i,(h,w) in enumerate(cols,1):
            c=ws.cell(row=1,column=i,value=h); c.font=hdr_font; c.fill=hdr_fill; c.alignment=ctr
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
        ws.row_dimensions[1].height=22
    if report_type == "received":
        ws.title = f"קבלות {report_date}"
        hdr(ws,[("תאריך",12),("מספר סריאל",20),("תיאור",30)])
        hist = load_inventory_history()
        ri=2
        for serial,entry in hist.get("received",{}).items():
            if isinstance(entry,dict) and entry.get("status")=="received" and entry.get("date")==report_date:
                for i,v in enumerate([report_date,serial,entry.get("desc","")],1):
                    c=ws.cell(row=ri,column=i,value=v); c.alignment=rght
                ri+=1
        ws.cell(row=ri+1,column=1,value=f'סה"כ: {ri-2} פריטים').font=Font(bold=True)
    else:
        ws.title = f"החזרות {report_date}"
        hdr(ws,[("תאריך",12),("מספר סריאל",20),("תיאור",30),("כמות",8)])
        conn=init_db()
        rows=conn.execute(
            "SELECT serial,description,quantity FROM returns WHERE phone=? AND return_date=? ORDER BY rowid DESC",
            (phone,report_date)
        ).fetchall(); conn.close()
        for ri,r in enumerate(rows,2):
            for i,v in enumerate([report_date,r[0],r[1],r[2] or 1],1):
                c=ws.cell(row=ri,column=i,value=v); c.alignment=rght
        ws.cell(row=len(rows)+3,column=1,value=f'סה"כ: {len(rows)} פריטים').font=Font(bold=True)
    buf=BytesIO(); wb.save(buf); buf.seek(0)
    fn=f"equipment_{report_type}_{report_date}.xlsx"
    return send_file(buf,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,download_name=fn)

@app.route("/api/equipment-report/delete-date", methods=["DELETE"])
def api_equipment_report_delete_date():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    data        = request.get_json() or {}
    report_date = data.get("date", "")
    report_type = data.get("type", "")   # "received" | "returned"
    if not report_date or not report_type:
        return jsonify({"error": "חסרים פרמטרים"}), 400
    if report_type == "received":
        hist = load_inventory_history()
        hist["received"] = {
            s: e for s, e in hist.get("received", {}).items()
            if not (isinstance(e, dict) and e.get("date") == report_date)
        }
        save_inventory_history(hist)
    elif report_type == "returned":
        conn = init_db()
        conn.execute("DELETE FROM returns WHERE phone=? AND return_date=?", (phone, report_date))
        conn.commit(); conn.close()
    return jsonify({"ok": True})

@app.route("/api/delete-return/<int:rid>", methods=["DELETE"])
def api_delete_return(rid):
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    conn = init_db()
    row = conn.execute("SELECT photo_filename FROM returns WHERE id=? AND phone=?", (rid, phone)).fetchone()
    if row:
        photo_filename = row[0]
        if photo_filename:
            try:
                os.remove(os.path.join(RETURNS_PHOTOS_DIR, photo_filename))
            except Exception:
                pass
        conn.execute("DELETE FROM returns WHERE id=? AND phone=?", (rid, phone))
        conn.commit()
    conn.close()
    return jsonify({"ok": True})

@app.route("/api/return-photo/<filename>")
def api_return_photo(filename):
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    # אבטחה — רק שם קובץ, בלי path traversal
    safe = os.path.basename(filename)
    path = os.path.join(RETURNS_PHOTOS_DIR, safe)
    if not os.path.exists(path):
        return jsonify({"error": "לא נמצא"}), 404
    return send_file(path, mimetype="image/jpeg")

@app.route("/qr-payment")
def serve_qr_payment():
    if os.path.exists(QR_PAYMENT_FILE):
        return send_file(QR_PAYMENT_FILE, mimetype="image/png")
    return "", 404

@app.route("/api/register", methods=["POST"])
def api_register():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    data = request.json or {}
    name = data.get("name", "").strip()
    if not name:
        return jsonify({"status": "error", "message": "הכנס שם"})
    employee_id = data.get("employee_id", "") or session.get("employee_id", "")
    db_register_user(phone, name, employee_id)
    return jsonify({"status": "ok"})

@app.route("/api/payment-confirm", methods=["POST"])
def api_payment_confirm():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    user = db_get_user(phone)
    if not user:
        return jsonify({"status": "error", "message": "לא רשום"})
    new_exp = db_extend_subscription(phone, SUB_DAYS)
    return jsonify({"status": "ok", "expires": new_exp})

@app.route("/api/admin/users")
def api_admin_users():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    user = db_get_user(phone)
    if phone != ADMIN_PHONE and not (user and user.get('is_admin')):
        return jsonify({"error": "אין הרשאה"}), 403
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT phone,name,employee_id,registered_at,trial_expires,subscription_expires,is_admin,permissions "
        "FROM users ORDER BY registered_at DESC"
    ).fetchall()
    conn.close()
    users_list = []
    for row in rows:
        u = dict(zip(['phone','name','employee_id','registered_at',
                      'trial_expires','subscription_expires','is_admin','permissions'], row))
        try:
            u['permissions'] = json.loads(u['permissions'] or '{}')
        except:
            u['permissions'] = {}
        st, exp = get_sub_status(u['phone'])
        u['sub_status'] = st; u['sub_expires'] = exp
        # merge with defaults — שמור ערכים מקוריים (false/true/date-string)
        merged = dict(DEFAULT_PERMISSIONS)
        merged.update({k: v for k, v in u['permissions'].items() if k in DEFAULT_PERMISSIONS})
        u['permissions'] = merged
        u['premium_expires'] = u.get('premium_expires') or ''
        users_list.append(u)
    return jsonify({"users": users_list, "permission_keys": list(DEFAULT_PERMISSIONS.keys())})

@app.route("/api/admin/permissions/<target_phone>", methods=["POST"])
def api_admin_set_permissions(target_phone):
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    user = db_get_user(phone)
    if phone != ADMIN_PHONE and not (user and user.get('is_admin')):
        return jsonify({"error": "אין הרשאה"}), 403
    data = request.json or {}
    # v יכול להיות: false, true, או מחרוזת ISO date
    new_perms = {}
    for k, v in data.items():
        if k not in DEFAULT_PERMISSIONS: continue
        if v is False or v is None:
            new_perms[k] = False
        elif v is True:
            new_perms[k] = True
        elif isinstance(v, str) and v:
            new_perms[k] = v   # ISO date
        else:
            new_perms[k] = bool(v)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("UPDATE users SET permissions=? WHERE phone=?",
                 (json.dumps(new_perms, ensure_ascii=False), target_phone))
    conn.commit(); conn.close()
    return jsonify({"status": "ok", "permissions": new_perms})

@app.route("/api/admin/extend", methods=["POST"])
def api_admin_extend():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    user = db_get_user(phone)
    if phone != ADMIN_PHONE and not (user and user.get('is_admin')):
        return jsonify({"error": "אין הרשאה"}), 403
    data = request.json or {}
    target = data.get("phone", "").strip()
    days = int(data.get("days", 30))
    if not target:
        return jsonify({"status": "error", "message": "חסר מספר טלפון"})
    new_exp = db_extend_subscription(target, days)
    return jsonify({"status": "ok", "expires": new_exp})

@app.route("/api/task-detail", methods=["POST"])
def api_task_detail():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    data = request.json or {}
    task = data.get("task", {})
    idx  = data.get("index", 1)
    card = build_single_card(task, token, idx)
    return jsonify(card)

# ============================================================
# צ'אט פנימי
# ============================================================
def _is_admin_phone(phone):
    u = db_get_user(phone)
    return phone == ADMIN_PHONE or bool(u and u.get('is_admin'))

@app.route("/api/chat/unread")
def api_chat_unread():
    token, phone = get_auth()
    if not token:
        return jsonify({"count": 0})
    conn = init_db()
    if _is_admin_phone(phone):
        row = conn.execute(
            "SELECT COUNT(*) FROM messages WHERE to_phone=? AND is_read=0", (phone,)
        ).fetchone()
    else:
        row = conn.execute(
            "SELECT COUNT(*) FROM messages WHERE to_phone=? AND is_read=0", (phone,)
        ).fetchone()
    conn.close()
    return jsonify({"count": row[0] if row else 0})

@app.route("/api/chat/conversations")
def api_chat_conversations():
    """רשימת שיחות למנהל — משתמשים שיש איתם הודעות"""
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    if not _is_admin_phone(phone):
        return jsonify({"error": "אין הרשאה"}), 403
    conn = init_db()
    rows = conn.execute(
        "SELECT DISTINCT CASE WHEN from_phone=? THEN to_phone ELSE from_phone END AS other "
        "FROM messages WHERE from_phone=? OR to_phone=? ORDER BY other",
        (phone, phone, phone)
    ).fetchall()
    result = []
    for (other,) in rows:
        if other == phone: continue
        unread = conn.execute(
            "SELECT COUNT(*) FROM messages WHERE from_phone=? AND to_phone=? AND is_read=0",
            (other, phone)
        ).fetchone()[0]
        last = conn.execute(
            "SELECT body,created_at FROM messages WHERE (from_phone=? AND to_phone=?) OR (from_phone=? AND to_phone=?) ORDER BY id DESC LIMIT 1",
            (other, phone, phone, other)
        ).fetchone()
        u = db_get_user(other)
        name = (u.get('name') or other) if u else other
        perms = get_effective_permissions(other)
        prem_exp = (u.get('premium_expires') or '') if u else ''
        result.append({
            "phone": other, "name": name,
            "unread": unread,
            "last_msg": last[0] if last else "",
            "last_at":  last[1][:16] if last else "",
            "premium": perms.get('premium', False),
            "premium_expires": prem_exp
        })
    conn.close()
    result.sort(key=lambda x: x['last_at'], reverse=True)
    return jsonify({"conversations": result})

@app.route("/api/chat/messages")
def api_chat_messages():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    is_adm = _is_admin_phone(phone)
    with_phone = request.args.get("with", ADMIN_PHONE if not is_adm else "")
    if not with_phone:
        return jsonify({"messages": []})
    conn = init_db()
    rows = conn.execute(
        "SELECT id,from_phone,to_phone,body,created_at,is_read FROM messages "
        "WHERE (from_phone=? AND to_phone=?) OR (from_phone=? AND to_phone=?) ORDER BY id ASC",
        (phone, with_phone, with_phone, phone)
    ).fetchall()
    # סמן כנקרא
    conn.execute(
        "UPDATE messages SET is_read=1 WHERE to_phone=? AND from_phone=? AND is_read=0",
        (phone, with_phone)
    )
    conn.commit(); conn.close()
    msgs = [{"id":r[0],"from":r[1],"to":r[2],"body":r[3],"at":r[4][:16],"read":bool(r[5])} for r in rows]
    return jsonify({"messages": msgs})

@app.route("/api/chat/send", methods=["POST"])
def api_chat_send():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    data = request.json or {}
    body = data.get("body", "").strip()
    if not body:
        return jsonify({"error": "הודעה ריקה"}), 400
    is_adm = _is_admin_phone(phone)
    to_phone = data.get("to", ADMIN_PHONE if not is_adm else "")
    if not to_phone:
        return jsonify({"error": "חסר נמען"}), 400
    now = datetime.now().isoformat()
    conn = init_db()
    conn.execute(
        "INSERT INTO messages (from_phone,to_phone,body,created_at) VALUES (?,?,?,?)",
        (phone, to_phone, body, now)
    )
    conn.commit(); conn.close()
    return jsonify({"ok": True})

@app.route("/api/admin/grant-premium", methods=["POST"])
def api_admin_grant_premium():
    token, phone = get_auth()
    if not token:
        return jsonify({"error": "לא מחובר"}), 401
    if not _is_admin_phone(phone):
        return jsonify({"error": "אין הרשאה"}), 403
    data = request.json or {}
    target = data.get("phone", "").strip()
    months = max(1, int(data.get("months", 1)))
    if not target:
        return jsonify({"error": "חסר טלפון"}), 400
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute("SELECT premium_expires FROM users WHERE phone=?", (target,)).fetchone()
    today = date.today()
    base = today
    if row and row[0]:
        try: base = max(date.fromisoformat(row[0]), today)
        except: pass
    new_exp = (base + timedelta(days=months * 30)).isoformat()
    conn.execute("UPDATE users SET premium_expires=? WHERE phone=?", (new_exp, target))
    conn.commit(); conn.close()
    # הודעת מערכת לאותו משתמש
    now = datetime.now().isoformat()
    conn2 = init_db()
    conn2.execute(
        "INSERT INTO messages (from_phone,to_phone,body,created_at) VALUES (?,?,?,?)",
        (phone, target, f"✅ הוענקה גישת פרמיום ל-{months} חודשים (עד {new_exp})", now)
    )
    conn2.commit(); conn2.close()
    return jsonify({"status": "ok", "premium_expires": new_exp})

# ============================================================
# HTML — Mobile-first
# ============================================================
HTML = r"""<!DOCTYPE html>
<html dir="rtl" lang="he">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="WorkManager">
<link rel="apple-touch-icon" href="/logo">
<link rel="icon" type="image/png" href="/logo">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<title>WorkManager</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:#0d1b2a;color:#ddeeff;height:100dvh;display:flex;flex-direction:column;overflow:hidden}

/* ─── Login / OTP screens ─── */
.screen{display:none;flex:1;flex-direction:column;align-items:center;justify-content:center;padding:24px 20px;gap:0}
.screen.active{display:flex}
.logo{margin-bottom:16px}
.logo img{width:90px;height:90px;border-radius:18px}
.screen h1{font-size:21px;font-weight:700;margin-bottom:6px;text-align:center;color:#eaf4ff}
.screen p{font-size:14px;color:#6b94b8;margin-bottom:28px;text-align:center}
.card{background:#1a2d42;border-radius:16px;padding:26px 22px;width:100%;max-width:380px;display:flex;flex-direction:column;gap:14px;box-shadow:0 8px 32px rgba(0,0,0,.5)}
.field{display:flex;flex-direction:column;gap:7px}
.field label{font-size:14px;color:#7baac8;font-weight:500}
.field input{background:#223d58;border:1px solid #284461;border-radius:10px;padding:13px 14px;color:#ddeeff;font-size:16px;outline:none;direction:rtl;width:100%}
.field input:focus{border-color:#3dba6f;box-shadow:0 0 0 2px rgba(61,186,111,.2)}
.field input::placeholder{color:#4a6a88}
.btn-main{background:#3dba6f;border:none;border-radius:10px;padding:14px;color:#fff;font-size:16px;font-weight:700;cursor:pointer;width:100%;margin-top:4px;letter-spacing:.3px}
.btn-main:active{background:#2e9458}
.btn-main:disabled{background:#223d58;color:#4a6a88;cursor:not-allowed}
.err{color:#f07080;font-size:13px;text-align:center;min-height:18px}
.back-link{font-size:13px;color:#6b94b8;text-align:center;cursor:pointer;text-decoration:underline;margin-top:10px}

/* ─── Chat screen ─── */
#chat-screen{display:none;flex:1;flex-direction:column;overflow:hidden}
#chat-screen.active{display:flex}
#header{background:#0d1b2a;padding:12px 16px;display:flex;align-items:center;gap:10px;flex-shrink:0;border-bottom:1px solid #1c3450}
.av{width:42px;height:42px;border-radius:10px;overflow:hidden;flex-shrink:0;display:flex;align-items:center;justify-content:center}
.ht{font-size:15px;font-weight:700;flex:1;color:#eaf4ff}
.hs{font-size:12px;color:#6b94b8}
#logout-btn{background:transparent;border:1px solid #284461;color:#6b94b8;border-radius:8px;padding:5px 12px;font-size:13px;cursor:pointer}
#logout-btn:active{background:#223d58}
#msgs{flex:1;overflow-y:auto;padding:12px 10px;display:flex;flex-direction:column;gap:4px;background:#0d1b2a}
#msgs::-webkit-scrollbar{width:4px}
#msgs::-webkit-scrollbar-thumb{background:#284461;border-radius:4px}
.msg{max-width:85%;padding:8px 12px 5px;border-radius:8px;font-size:14px;line-height:1.6;white-space:pre-wrap;word-break:break-word}
.user{background:#0d4a32;align-self:flex-start;border-radius:8px 8px 0 8px}
.bot{background:#d0e4f7;align-self:flex-end;border-radius:8px 8px 8px 0}
.msg b{font-weight:700}
.mtime{font-size:11px;color:#6b94b8;text-align:left;margin-top:2px}
.qr{display:flex;gap:6px;flex-wrap:wrap;justify-content:flex-end;margin:6px 0}
.qb{background:#1e3a5c;border:1px solid #3d6a9a;color:#ddeeff;border-radius:20px;padding:9px 18px;font-size:14px;font-weight:600;cursor:pointer;white-space:nowrap;box-shadow:0 2px 10px rgba(0,0,0,.5)}
.qb:active{background:#122a44;border-color:#2a5080}
#typing{display:none;align-self:flex-end;background:#d0e4f7;padding:10px 14px;border-radius:8px 8px 8px 0;margin-bottom:2px}
#typing span{display:inline-block;width:7px;height:7px;background:#6b94b8;border-radius:50%;animation:b 1.2s infinite;margin:0 2px}
#typing span:nth-child(2){animation-delay:.2s}
#typing span:nth-child(3){animation-delay:.4s}
@keyframes b{0%,60%,100%{transform:translateY(0)}30%{transform:translateY(-6px)}}
#bar{background:#0d1b2a;padding:8px 10px;display:flex;gap:8px;align-items:center;flex-shrink:0;border-top:1px solid #1c3450}
#inp{flex:1;background:#223d58;border:none;border-radius:20px;padding:10px 16px;color:#ddeeff;font-size:15px;outline:none;direction:rtl}
#inp::placeholder{color:#6b94b8}
#send-btn{width:42px;height:42px;background:#3dba6f;border:none;border-radius:50%;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0}
#send-btn svg{fill:white;width:20px;height:20px}
/* ─── Tabs ─── */
#tabs{background:#0d1b2a;border-bottom:1px solid #1c3450}
.tab-btn{flex:1;padding:12px;background:transparent;border:none;color:#6b94b8;font-size:16px;font-weight:600;cursor:pointer;border-bottom:2px solid transparent}
.tab-btn.active{color:#3dba6f;border-bottom:2px solid #3dba6f;font-weight:700}
/* ─── Inventory cards ─── */
.inv-card{background:#d0e4f7;color:#0d1b2a;border-radius:12px;padding:14px 16px;margin:6px 0;direction:rtl;font-size:15.5px;font-weight:500;line-height:2;box-shadow:0 2px 12px rgba(0,0,0,.5)}
.inv-card.ok{border-right:4px solid #3dba6f}
.inv-card.blocked{border-right:4px solid #f07080}
.inv-badge{display:inline-block;padding:2px 9px;border-radius:10px;font-size:12px;font-weight:700}
.badge-ok{background:#0a3d22;color:#3dba6f}
.badge-blocked{background:#3d0d16;color:#f07080}
.badge-new{background:#0a2a4a;color:#5ab8f5;font-size:11px;margin-right:6px}
/* ─── Accordion ─── */
.acc{border-top:1px solid #1c3450;margin-top:8px;cursor:pointer}
.acc-hdr{color:#6b94b8;font-size:13px;display:flex;justify-content:space-between;padding:6px 0;user-select:none}
.acc-arr{transition:transform .2s;display:inline-block}
.acc-body{display:none;color:#9bbdd8;font-size:13px;padding:4px 0 6px;line-height:1.8}
.acc.open .acc-body{display:block}
.acc.open .acc-arr{transform:rotate(180deg)}
/* ─── Task cards ─── */
.task-card{background:#d0e4f7 !important;border-right-color:#3dba6f !important}
</style>
</head>
<body>

<!-- ══ מסך כניסה ══ -->
<div class="screen" id="login-screen">
  <div class="logo"><img src="/logo" onerror="this.outerHTML='<span style=font-size:52px>🔧</span>'"></div>
  <h1>WorkManager</h1>
  <p>הכנס את הפרטים שלך כדי להמשיך</p>
  <div class="card">
    <div class="field">
      <label>מספר טלפון</label>
      <input id="l-phone" type="tel" placeholder="05XXXXXXXX" inputmode="numeric" />
    </div>
    <div class="field">
      <label>מספר עובד</label>
      <input id="l-emp" type="text" placeholder="מספר עובד" inputmode="numeric" />
    </div>
    <div class="err" id="l-err"></div>
    <button class="btn-main" id="l-btn" onclick="doLogin()">כניסה</button>
  </div>
</div>

<!-- ══ מסך OTP ══ -->
<div class="screen" id="otp-screen">
  <div class="logo">📱</div>
  <h1>אימות SMS</h1>
  <p>נשלח קוד למספר שהזנת.<br>הכנס אותו כאן:</p>
  <div class="card">
    <div class="field">
      <label>קוד אימות</label>
      <input id="otp-inp" type="text" placeholder="123456" inputmode="numeric" maxlength="8"
             onkeypress="if(event.key==='Enter')doVerify()" />
    </div>
    <div class="err" id="otp-err"></div>
    <button class="btn-main" id="otp-btn" onclick="doVerify()">אמת קוד</button>
    <div class="back-link" onclick="showScreen('login')">← חזור לכניסה</div>
  </div>
</div>

<!-- ══ מסך הרשמה ══ -->
<div class="screen" id="reg-screen">
  <div class="logo"><img src="/logo" onerror="this.outerHTML='<span style=font-size:52px>🔧</span>'"></div>
  <h1>ברוך הבא! 👋</h1>
  <p>הרשמה ראשונה — <b style="color:#3dba6f">7 ימי ניסיון חינם</b></p>
  <div class="card">
    <div class="field">
      <label>השם שלך</label>
      <input id="reg-name" type="text" placeholder="שם מלא" autocomplete="name" />
    </div>
    <div class="err" id="reg-err"></div>
    <button class="btn-main" id="reg-btn" onclick="doRegister()">🚀 התחל ניסיון חינם</button>
    <div style="font-size:12px;color:#8fa5b5;text-align:center;margin-top:4px">לאחר 7 ימים — 100 ₪/חודש</div>
  </div>
</div>


<!-- ══ מסך צ'אט ══ -->
<div id="chat-screen">
  <div id="header">
    <div class="av"><img src="/logo" style="width:42px;height:42px;object-fit:cover;border-radius:10px" onerror="this.parentElement.innerHTML='🔧'"></div>
    <div style="flex:1"><div class="ht">סלקום טכנאים</div><div class="hs" id="st">מחובר</div></div>
    <button onclick="location.reload()" style="background:none;border:none;color:#6b94b8;font-size:20px;cursor:pointer;padding:4px 8px" title="רענון">🔄</button>
    <button id="logout-btn" onclick="doLogout()">יציאה</button>
  </div>
  <div id="tabs" style="display:flex;background:#0d1b2a;border-bottom:1px solid #284461;flex-shrink:0">
    <button class="tab-btn active" id="tab-tasks" onclick="switchTab('tasks')">📋 פקעות</button>
    <button class="tab-btn" id="tab-inv" onclick="switchTab('inv')">🏭 מחסן</button>
    <button class="tab-btn" id="tab-rep" onclick="switchTab('rep')">📊 דוחות</button>
    <button class="tab-btn" id="tab-price" onclick="switchTab('price')">💰 מחירון</button>
    <button class="tab-btn" id="tab-admin" onclick="switchTab('admin')" style="display:none">🛡️ מנהל</button>
  </div>
  <div id="msgs">
    <div class="qr">
      <button class="qb" id="sort-btn" onclick="toggleSortByCity()">📍 מיין לפי מיקום</button>
      <button class="qb" onclick="loadTasks()">📋 בנק פקעות</button>
    </div>
  </div>
  <div id="inv-panel" style="display:none;flex:1;overflow-y:auto;direction:rtl">
    <!-- תת-לשוניות מחסן -->
    <div style="display:flex;border-bottom:1px solid #284461;background:#0d1b2a;position:sticky;top:0;z-index:5">
      <button id="sub-tab-inv" onclick="switchSubTab('inv')"
        style="flex:1;padding:10px;background:transparent;border:none;border-bottom:2px solid #3dba6f;color:#3dba6f;font-size:13px;font-weight:700;cursor:pointer">📦 מלאי</button>
      <button id="sub-tab-eqrep" onclick="switchSubTab('eqrep')"
        style="flex:1;padding:10px;background:transparent;border:none;border-bottom:2px solid transparent;color:#6b94b8;font-size:13px;cursor:pointer">🗂️ ציוד</button>
    </div>
    <!-- תוכן מלאי -->
    <div id="inv-content" style="padding:12px 10px"></div>
    <!-- תוכן ציוד -->
    <div id="eqrep-content" style="display:none;padding:12px 10px">
      <div style="display:flex;gap:8px;margin-bottom:12px">
        <button id="eqrep-tab-rcv" onclick="switchEqTab('rcv')"
          style="flex:1;padding:9px;border-radius:8px;border:none;background:#1f7a4a;color:#fff;font-size:14px;font-weight:700;cursor:pointer">
          📥 קיבלתי</button>
        <button id="eqrep-tab-ret" onclick="switchEqTab('ret')"
          style="flex:1;padding:9px;border-radius:8px;border:none;background:#223d58;color:#6b94b8;font-size:14px;cursor:pointer">
          📤 החזרות</button>
      </div>
      <div id="eqrep-rcv-list"></div>
      <div id="eqrep-ret-list" style="display:none"></div>
    </div>
  </div>
  <div id="rep-panel" style="display:none;flex:1;overflow-y:auto;padding:12px 10px;direction:rtl"></div>
  <div id="price-panel" style="display:none;flex:1;overflow-y:auto;padding:12px 10px;direction:rtl"></div>
  <div id="admin-panel" style="display:none;flex:1;overflow-y:auto;padding:12px 10px;direction:rtl"></div>

  <!-- מודל סריקת MAC/סריאל -->
  <div id="scan-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,0.85);z-index:9000;flex-direction:column;align-items:center;justify-content:center;padding:20px">
    <div style="background:#d0e4f7;border-radius:16px;padding:18px;width:100%;max-width:420px;direction:rtl">
      <div style="font-size:17px;font-weight:700;margin-bottom:12px;color:#e9edef">📷 זיהוי ציוד</div>
      <img id="scan-preview" src="" style="width:100%;border-radius:10px;max-height:220px;object-fit:contain;background:#0d1b2a;display:none">
      <div id="scan-status" style="margin-top:10px;font-size:14px;min-height:40px;text-align:center"></div>
      <div id="scan-result" style="margin-top:4px"></div>
    </div>
  </div>
</div>

<script>
/* ── helpers ── */
function showScreen(name){
  ['login-screen','otp-screen','chat-screen','reg-screen','payment-screen'].forEach(id=>{
    const el=document.getElementById(id);
    if(el) el.classList.remove('active');
  });
  document.getElementById('chat-screen').classList.remove('active');
  if(name==='login') document.getElementById('login-screen').classList.add('active');
  else if(name==='otp') document.getElementById('otp-screen').classList.add('active');
  else if(name==='chat') document.getElementById('chat-screen').classList.add('active');
  else if(name==='reg') document.getElementById('reg-screen').classList.add('active');
  else if(name==='payment') document.getElementById('payment-screen').classList.add('active');
}

let _isAdmin=false;
let _userName='';
let _userPerms={premium:false,full_card:false,inventory:false,equipment_report:false,reports:false,prices:false};

/* ── apiFetch — מוסיף X-Token/X-Phone לכל בקשת API ── */
function apiFetch(url, opts={}){
  const token=localStorage.getItem('cc_token')||'';
  const phone=localStorage.getItem('cc_phone')||'';
  opts.headers=Object.assign({'X-Token':token,'X-Phone':phone},opts.headers||{});
  return fetch(url,opts).then(res=>{
    if(res.status===401){
      localStorage.removeItem('cc_token');
      showScreen('login');
      document.getElementById('l-err').textContent='פג תוקף החיבור — התחבר מחדש';
    }
    return res;
  });
}

function saveAuthLocally(token, phone){
  if(token) localStorage.setItem('cc_token', token);
  if(phone) localStorage.setItem('cc_phone', phone);
}

function showAdminTab(){
  const tb=document.getElementById('tab-admin');
  if(tb) tb.style.display='';
}

function applyPermissions(){
  // הלשוניות תמיד גלויות — רק sub-tab ציוד מוסתר לפי הרשאה
  const subEqrep=document.getElementById('sub-tab-eqrep');
  const hasPremium=_isAdmin||_userPerms.premium;
  if(subEqrep) subEqrep.style.display=(hasPremium&&(_userPerms.equipment_report||_isAdmin))?'':'none';
}

const NO_PERM_HTML=`<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;padding:40px 20px;text-align:center;gap:16px">
  <div style="font-size:52px">🔒</div>
  <div style="font-size:17px;font-weight:700;color:#edf2f5">אין הרשאה</div>
  <div style="font-size:14px;color:#8fa5b5;max-width:260px">פנה למנהל לקבלת גישה ללשונית זו</div>
</div>`;

async function handleSubStatus(sub_status){
  if(sub_status==='not_registered'){showScreen('reg');return true;}
  return false;
}

async function attemptLogin(phone, emp, silent){
  const errEl=document.getElementById('l-err');
  const btn=document.getElementById('l-btn');
  errEl.textContent='';
  if(!silent){btn.disabled=true;btn.textContent='מתחבר...';}
  try{
    // שלח טוקן שמור (אם יש) — השרת יאמת אותו
    const savedToken=localStorage.getItem('cc_token')||'';
    const r=await apiFetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({phone,employee_id:emp,token:savedToken})});
    const d=await r.json();
    if(d.status==='ok'){
      // שמור טוקן שהשרת החזיר
      if(d.token) saveAuthLocally(d.token, d.phone||phone);
      const sd=await apiFetch('/api/status').then(r=>r.json());
      _isAdmin=sd.is_admin||false; _userName=sd.user_name||'';
      if(sd.permissions) _userPerms=Object.assign({},_userPerms,sd.permissions);
      if(_isAdmin) showAdminTab();
      applyPermissions();
      const handled=await handleSubStatus(sd.sub_status);
      if(!handled){showScreen('chat');initChat();}
    } else if(d.status==='otp'){
      showScreen('otp');
      setTimeout(()=>document.getElementById('otp-inp').focus(),100);
    } else {
      errEl.textContent=d.message||'שגיאה';
      showScreen('login');
    }
  }catch(e){
    errEl.textContent='שגיאת רשת';
    showScreen('login');
  }
  if(!silent){btn.disabled=false;btn.textContent='כניסה';}
}

/* ── on load ── */
(async function(){
  // נסה טוקן שמור ב-localStorage
  const phone=localStorage.getItem('cc_phone');
  const token=localStorage.getItem('cc_token');
  const emp=localStorage.getItem('cc_emp');
  if(phone && token){
    let d={logged_in:false};
    try{d=await apiFetch('/api/status').then(r=>r.json());}catch(e){}
    if(d.logged_in){
      _isAdmin=d.is_admin||false; _userName=d.user_name||'';
      if(d.permissions) _userPerms=Object.assign({},_userPerms,d.permissions);
      if(_isAdmin) showAdminTab();
      applyPermissions();
      const handled=await handleSubStatus(d.sub_status);
      if(!handled) showScreen('chat');
      return;
    }
  }
  // אין טוקן — נסה כניסה שקטה עם phone+emp
  if(phone && emp){
    const ok=await attemptLoginSilent(phone, emp);
    if(ok) return;
    const masked='05'+'*'.repeat(Math.max(0,phone.length-2));
    document.getElementById('l-phone').value='';
    document.getElementById('l-phone').placeholder=masked;
    document.getElementById('l-emp').value='';
    document.getElementById('l-err').textContent='';
    const hint=document.createElement('div');
    hint.id='quick-login-hint';
    hint.style.cssText='text-align:center;font-size:13px;color:#8fa5b5;margin-top:-6px';
    hint.innerHTML=`מחובר בד"כ כ <b style="color:#dde8ee">${masked}</b> — <span style="color:#3dba6f;cursor:pointer;text-decoration:underline" onclick="quickLogin()">כניסה מהירה</span>`;
    const card=document.querySelector('#login-screen .card');
    card.insertBefore(hint, card.firstChild);
  }
  showScreen('login');
})();

async function attemptLoginSilent(phone, emp){
  try{
    const savedToken=localStorage.getItem('cc_token')||'';
    const r=await apiFetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({phone,employee_id:emp,token:savedToken})});
    const d=await r.json();
    if(d.status==='ok'){
      if(d.token) saveAuthLocally(d.token, d.phone||phone);
      const sd=await apiFetch('/api/status').then(r=>r.json());
      _isAdmin=sd.is_admin||false; _userName=sd.user_name||'';
      if(sd.permissions) _userPerms=Object.assign({},_userPerms,sd.permissions);
      if(_isAdmin) showAdminTab();
      applyPermissions();
      const handled=await handleSubStatus(sd.sub_status);
      if(!handled){showScreen('chat');initChat();}
      return true;
    }
    return false;
  }catch(e){return false;}
}

async function quickLogin(){
  const phone=localStorage.getItem('cc_phone');
  const emp=localStorage.getItem('cc_emp');
  if(phone&&emp) await attemptLogin(phone,emp,false);
}

/* ── login ── */
async function doLogin(){
  const phone=document.getElementById('l-phone').value.replace(/-/g,'').trim();
  const emp=document.getElementById('l-emp').value.trim();
  const errEl=document.getElementById('l-err');
  errEl.textContent='';
  if(!phone||!emp){errEl.textContent='יש למלא את כל השדות';return;}
  // save for future auto-login
  localStorage.setItem('cc_phone', phone);
  localStorage.setItem('cc_emp', emp);
  await attemptLogin(phone, emp, false);
}
document.getElementById('l-phone').addEventListener('keypress',e=>{if(e.key==='Enter')document.getElementById('l-emp').focus()});
document.getElementById('l-emp').addEventListener('keypress',e=>{if(e.key==='Enter')doLogin()});

/* ── OTP ── */
async function doVerify(){
  const otp=document.getElementById('otp-inp').value.trim();
  const errEl=document.getElementById('otp-err');
  const btn=document.getElementById('otp-btn');
  errEl.textContent='';
  if(!otp){errEl.textContent='הכנס קוד';return;}
  btn.disabled=true;btn.textContent='מאמת...';
  try{
    const r=await apiFetch('/api/verify',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({otp})});
    const d=await r.json();
    if(d.status==='ok'){
      // שמור טוקן ב-localStorage
      if(d.token) saveAuthLocally(d.token, d.phone||localStorage.getItem('cc_phone'));
      document.getElementById('otp-inp').value='';
      const sd=await apiFetch('/api/status').then(r=>r.json());
      _isAdmin=sd.is_admin||false; _userName=sd.user_name||'';
      if(sd.permissions) _userPerms=Object.assign({},_userPerms,sd.permissions);
      if(_isAdmin) showAdminTab();
      applyPermissions();
      const handled=await handleSubStatus(sd.sub_status);
      if(!handled){showScreen('chat');initChat();}
    } else {errEl.textContent=d.message||'קוד שגוי';}
  }catch(e){errEl.textContent='שגיאת רשת';}
  btn.disabled=false;btn.textContent='אמת קוד';
}

/* ── logout ── */
async function doLogout(){
  await apiFetch('/api/logout',{method:'POST'});
  localStorage.removeItem('cc_phone');
  localStorage.removeItem('cc_emp');
  localStorage.removeItem('cc_token');
  document.getElementById('msgs').innerHTML=`<div class="qr" id="tasks-qr"><button class="qb" id="sort-btn" onclick="toggleSortByCity()">📍 מיין לפי מיקום</button><button class="qb" onclick="loadTasks()">📋 בנק פקעות</button></div>`;
  showScreen('login');
}

/* ── chat ── */
function ts(){return new Date().toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'})}
function md(t){return t.replace(/\*(.*?)\*/g,'<b>$1</b>')}
function add(t,type){
  const b=document.getElementById('msgs');
  const d=document.createElement('div');
  d.className='msg '+type;
  const safe=t.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  d.innerHTML=md(safe)+'<div class="mtime">'+ts()+'</div>';
  b.appendChild(d);b.scrollTop=9999;
}
function s(t){document.getElementById('inp').value=t;go()}
function toggleAcc(el){el.classList.toggle('open')}
function addCard(c,box){
  if(!box) box=document.getElementById('msgs');
  const card=document.createElement('div');
  card.className='task-card';
  card.style.cssText='background:#d0e4f7;color:#0d1b2a;border-radius:10px;padding:14px 16px;direction:rtl;border-right:4px solid #25d366;font-size:15.5px;font-weight:500;line-height:2;margin:4px 0';

  const repeatBadge=c.recent_visit
    ?`<span style="color:#f15c6e;font-size:12px;font-weight:700">⚠️ חשד לחזורת — <span style="text-decoration:underline">${c.recent_visit}</span></span>`
    :'';

  if(_isAdmin||_userPerms.full_card){
    // מנהל / הרשאת כרטיס מלא
    let html=`<div style="font-weight:700;font-size:15px;color:#25d366;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center">
      <span>${c._num?`<span style="color:#8696a0;font-size:13px;margin-left:6px">${c._num}.</span>`:''}📋 קריאה ${c.call_id}</span>${repeatBadge}</div>`;
    if(c.date)    html+=`<div>📅 ${c.date}&nbsp;&nbsp;⏰ ${c.time}</div>`;
    html+=`<div>👤 ${c.name}&nbsp;&nbsp;<span style="color:#8696a0;font-size:12px">${c.customer_id}</span></div>`;
    if(c.phone)   html+=`<div>📞 <a href="tel:${c.phone}" style="color:#53bdeb">${c.phone}</a></div>`;
    html+=`<div>📍 ${c.address}${c._dist?` <span style="color:#53bdeb;font-size:12px">(${c._dist} ק"מ)</span>`:''}</div>`;
    html+=`<div>🔧 ${c.task_type}</div>`;
    if(c.infra)   html+=`<div>🌐 ${c.infra}${c.technology?' · '+c.technology:''}</div>`;
    if(c.comment) html+=`<div>💬 ${c.comment}</div>`;
    if(c.planned||c.existing||c.tv){
      html+=`<div class="acc" onclick="toggleAcc(this)"><div class="acc-hdr"><span>⚙️ ציוד</span><span class="acc-arr">▾</span></div><div class="acc-body">`;
      if(c.planned)  html+=`<div>📦 להתקנה: ${c.planned}</div>`;
      if(c.existing) html+=`<div>🖥 קיים: ${c.existing}</div>`;
      if(c.tv)       html+=`<div>📺 TV: ${c.tv}</div>`;
      html+=`</div></div>`;
    }
    if(c.history){
      html+=`<div class="acc" onclick="toggleAcc(this)"><div class="acc-hdr"><span>📜 היסטוריה</span><span class="acc-arr">▾</span></div><div class="acc-body">${c.history.replace(/\n/g,'<br>')}</div></div>`;
    }
    card.innerHTML=html;
  } else {
    // משתמש רגיל
    let html=`<div style="font-weight:700;font-size:15px;color:#25d366;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center">
      <span>${c._num?`<span style="color:#8696a0;font-size:13px;margin-left:6px">${c._num}.</span>`:''}📋 קריאה ${c.call_id}</span>${repeatBadge}</div>`;
    if(c.date)  html+=`<div>📅 ${c.date}&nbsp;&nbsp;⏰ ${c.time}</div>`;
    html+=`<div>👤 ${c.name}</div>`;
    if(c.phone) html+=`<div>📞 <a href="tel:${c.phone}" style="color:#53bdeb">${c.phone}</a></div>`;
    html+=`<div>📍 ${c.address}${c._dist?` <span style="color:#53bdeb;font-size:12px">(${c._dist} ק"מ)</span>`:''}</div>`;
    html+=`<div>🔧 ${c.task_type}</div>`;
    if(c.history){
      html+=`<div class="acc" onclick="toggleAcc(this)"><div class="acc-hdr"><span>📜 היסטוריה</span><span class="acc-arr">▾</span></div><div class="acc-body">${c.history.replace(/\n/g,'<br>')}</div></div>`;
    }
    card.innerHTML=html;
  }

  box.appendChild(card);
  box.scrollTop=9999;
}
let _loadTasksAbort=null;
let _allCards=[];
let _sortedByCity=false;
function renderCards(cards,box,hdr,count){
  const existingCards=box.querySelectorAll('.task-card');
  existingCards.forEach(el=>el.remove());
  cards.forEach(c=>addCard(c,box));
  hdr.textContent=`📋 בנק פקעות — ${new Date().toLocaleDateString('he-IL')} | סה"כ: ${count} פקעות`;
}
function haversine(lat1,lon1,lat2,lon2){
  const R=6371, dLat=(lat2-lat1)*Math.PI/180, dLon=(lon2-lon1)*Math.PI/180;
  const a=Math.sin(dLat/2)**2+Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;
  return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}
async function toggleSortByCity(){
  const box=document.getElementById('msgs');
  const btn=document.getElementById('sort-btn');
  if(_sortedByCity){
    _sortedByCity=false;
    btn.textContent='📍 מיין לפי מיקום';
    const hdr=box.querySelector('.tasks-hdr');
    renderCards(_allCards,box,hdr,_allCards.length);
    return;
  }
  btn.textContent='⏳ מאתר מיקום...';
  btn.disabled=true;
  try{
    const pos=await new Promise((res,rej)=>navigator.geolocation.getCurrentPosition(res,rej,{timeout:8000}));
    const uLat=pos.coords.latitude, uLng=pos.coords.longitude;
    btn.textContent='⏳ מחשב מרחקים...';
    const cards=await Promise.all(_allCards.map(async c=>{
      if(c._lat!=null) return c;
      const q=encodeURIComponent((c.city||c.address||''));
      try{const r=await apiFetch('/api/geocode?q='+q);const d=await r.json();if(d.lat){c._lat=d.lat;c._lng=d.lng;}}catch(e){}
      return c;
    }));
    const sorted=[...cards].sort((a,b)=>{
      const da=a._lat!=null?haversine(uLat,uLng,a._lat,a._lng):9999;
      const db=b._lat!=null?haversine(uLat,uLng,b._lat,b._lng):9999;
      return da-db;
    });
    sorted.forEach(c=>{if(c._lat!=null)c._dist=haversine(uLat,uLng,c._lat,c._lng).toFixed(1);});
    _sortedByCity=true;
    const hdr=box.querySelector('.tasks-hdr');
    renderCards(sorted,box,hdr,_allCards.length);
    btn.textContent='🔢 מיון מקורי';
  }catch(e){
    btn.textContent='📍 מיין לפי מיקום';
    alert('לא ניתן לאתר מיקום. אפשר גישה ל-GPS בהגדרות.');
  }
  btn.disabled=false;
}
async function loadTasks(district=''){
  if(_loadTasksAbort){_loadTasksAbort.abort();}
  _loadTasksAbort=new AbortController();
  _allCards=[];_sortedByCity=false;
  const sig=_loadTasksAbort.signal;
  const box=document.getElementById('msgs');
  box.innerHTML=`<div class="qr" id="tasks-qr"><button class="qb" id="sort-btn" onclick="toggleSortByCity()">📍 מיין לפי מיקום</button><button class="qb" onclick="loadTasks()">📋 בנק פקעות</button></div>`;
  document.getElementById('st').textContent='טוען...';
  const hdr=document.createElement('div');
  hdr.className='tasks-hdr';
  hdr.style.cssText='color:#8696a0;font-size:13px;text-align:right;margin:6px 0';
  hdr.textContent='מושך פקעות...';
  box.appendChild(hdr);
  box.scrollTop=9999;
  try{
    const url='/api/tasks'+(district?'?district='+encodeURIComponent(district):'');
    const r=await apiFetch(url,{signal:sig});
    if(r.status===401){hdr.textContent='❌ לא מחובר';document.getElementById('st').textContent='מחובר';return;}
    const d=await r.json();
    if(!d.tasks||d.tasks.length===0){
      hdr.textContent='לא נמצאו פקעות.';
      document.getElementById('st').textContent='מחובר';
      return;
    }
    hdr.textContent=`📋 בנק פקעות — ${new Date().toLocaleDateString('he-IL')} | סה"כ: ${d.count} פקעות`;
    for(let i=0;i<d.tasks.length;i++){
      hdr.textContent=`📋 בנק פקעות — ${new Date().toLocaleDateString('he-IL')} | ${i+1}/${d.count}...`;
      const cr=await apiFetch('/api/task-detail',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({task:d.tasks[i],index:i+1}),signal:sig});
      const card=await cr.json();
      card._num=i+1;
      _allCards.push(card);
      addCard(card,box);
    }
    hdr.textContent=`📋 בנק פקעות — ${new Date().toLocaleDateString('he-IL')} | סה"כ: ${d.count} פקעות`;
    document.getElementById('st').textContent='מחובר';
    // הוסף כפתור מפה אחרי טעינה מוצלחת
    const qr=document.getElementById('tasks-qr');
    if(qr&&!document.getElementById('tasks-map-btn')){
      const mb=document.createElement('button');
      mb.className='qb'; mb.id='tasks-map-btn';
      mb.textContent='🗺️ הצג במפה';
      mb.onclick=showTasksMap;
      qr.appendChild(mb);
    }
  }catch(e){
    if(e.name==='AbortError') return;
    hdr.textContent='❌ שגיאה בטעינה';
    document.getElementById('st').textContent='מחובר';
  }
}
/* ── tabs ── */
let _activeSubTab='inv';
function switchSubTab(sub){
  _activeSubTab=sub;
  const invContent=document.getElementById('inv-content');
  const eqContent=document.getElementById('eqrep-content');
  const btnInv=document.getElementById('sub-tab-inv');
  const btnEq=document.getElementById('sub-tab-eqrep');
  if(sub==='inv'){
    invContent.style.display='block'; eqContent.style.display='none';
    btnInv.style.color='#3dba6f'; btnInv.style.borderBottomColor='#3dba6f'; btnInv.style.fontWeight='700';
    btnEq.style.color='#8696a0';  btnEq.style.borderBottomColor='transparent'; btnEq.style.fontWeight='400';
  } else {
    invContent.style.display='none'; eqContent.style.display='block';
    btnEq.style.color='#3dba6f'; btnEq.style.borderBottomColor='#3dba6f'; btnEq.style.fontWeight='700';
    btnInv.style.color='#8696a0'; btnInv.style.borderBottomColor='transparent'; btnInv.style.fontWeight='400';
    if(!document.getElementById('eqrep-content').dataset.loaded){
      document.getElementById('eqrep-content').dataset.loaded='1';
      loadEqReport();
    }
  }
}

function canAccess(tab){
  if(_isAdmin) return true;
  if(tab==='tasks') return true;
  if(!_userPerms.premium) return false;
  if(tab==='inv')    return _userPerms.inventory;
  if(tab==='rep')    return _userPerms.reports;
  if(tab==='price')  return _userPerms.prices;
  return true;
}

function switchTab(tab){
  ['tasks','inv','rep','price','admin'].forEach(t=>{
    const el=document.getElementById('tab-'+t);
    if(el) el.classList.toggle('active',tab===t);
  });
  const msgs=document.getElementById('msgs');
  const inv=document.getElementById('inv-panel');
  const rep=document.getElementById('rep-panel');
  const price=document.getElementById('price-panel');
  const adm=document.getElementById('admin-panel');
  msgs.style.display='none';
  inv.style.display='none';
  rep.style.display='none'; price.style.display='none'; adm.style.display='none';
  if(tab==='tasks'){
    msgs.style.display='flex';
  } else if(tab==='inv'){
    inv.style.display='flex'; inv.style.flexDirection='column';
    if(!canAccess('inv')){inv.innerHTML=NO_PERM_HTML;return;}
    if(!inv.dataset.loaded) loadInventory();
  } else if(tab==='rep'){
    rep.style.display='block';
    if(!canAccess('rep')){rep.innerHTML=NO_PERM_HTML;return;}
    if(!rep.dataset.loaded) initReports();
  } else if(tab==='price'){
    price.style.display='block';
    if(!canAccess('price')){price.innerHTML=NO_PERM_HTML;return;}
    if(!price.dataset.loaded) loadPriceList();
  } else if(tab==='admin'){
    adm.style.display='block';
    if(!adm.dataset.loaded) loadAdminPanel();
  }
}

/* ── inventory ── */
async function loadInventory(){
  const panel=document.getElementById('inv-content');
  panel.innerHTML='<div style="color:#8696a0;text-align:center;padding:30px">טוען מלאי...</div>';
  try{
    const r=await apiFetch('/api/inventory');
    if(r.status===401){document.getElementById('inv-content').innerHTML='<div style="color:#f15c6e;text-align:center;padding:30px">❌ לא מחובר</div>';return;}
    const d=await r.json();
    if(!d.items||d.items.length===0){panel.innerHTML='<div style="color:#8696a0;text-align:center;padding:30px">לא נמצא מלאי</div>';return;}
    panel.dataset.loaded='1';
    const today=new Date().toISOString().slice(0,10);

    // שמור סריאלים לזיהוי OCR
    _invSerials=[];
    d.items.forEach(item=>{ if(item.serial) _invSerials.push({serial:item.serial,desc:item.description||''}); });

    // קיבוץ לפי קטגוריה → תיאור
    const cats={};
    d.items.forEach(item=>{
      if(!cats[item.category]) cats[item.category]={};
      const key=item.description||'ללא שם';
      if(!cats[item.category][key]) cats[item.category][key]=[];
      cats[item.category][key].push(item);
    });

    let html=`<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;flex-wrap:wrap;gap:6px">
      <div style="color:#8696a0;font-size:13px">עודכן: ${d.last_updated} | סה"כ: ${d.total} פריטים</div>
      <div style="display:flex;gap:6px">
        <button id="scan-mac-btn" class="qb" style="font-size:13px;padding:6px 14px;background:#1a3a5c;border-color:#3d5166;display:none" onclick="scanEquipment()">📷 סרוק ציוד</button>
        <button id="return-mode-btn" class="qb" style="font-size:13px;padding:6px 14px" onclick="toggleReturnMode()">📦 החזרת ציוד</button>
      </div>
    </div>`;
    let uid=0;

    for(const [cat, descMap] of Object.entries(cats)){
      const isBlocked=cat.includes('חסום')||cat.includes('לא תקין');
      const catColor=isBlocked?'#f15c6e':'#25d366';
      const totalInCat=Object.values(descMap).reduce((s,a)=>s+a.length,0);
      html+=`<div style="font-weight:700;color:${catColor};margin:14px 0 6px;font-size:15px">${cat} — ${totalInCat} יח'</div>`;

      for(const [desc, items] of Object.entries(descMap)){
        const count=items.reduce((s,i)=>s+(i.serial?1:(i.amount||1)),0);
        const catalog=items[0].catalog||'';
        const hasNew=items.some(i=>i.first_seen===today);
        const cls=isBlocked?'inv-card blocked':'inv-card ok';
        const badge=isBlocked?'<span class="inv-badge badge-blocked">חסום</span>':'<span class="inv-badge badge-ok">תקין</span>';
        const newBadge=hasNew?'<span class="inv-badge badge-new">חדש היום</span>':'';
        const id='inv'+uid++;

        html+=`<div class="${cls}" style="cursor:pointer" onclick="toggleInvAcc('${id}')">
          <div style="display:flex;justify-content:space-between;align-items:center">
            <span style="font-weight:700">${desc}</span>
            <span style="display:flex;gap:6px;align-items:center">${newBadge}${badge}<span style="background:#223d58;color:#ddeeff;border-radius:12px;padding:2px 10px;font-size:13px;font-weight:700">${count}</span></span>
          </div>`;
        if(catalog) html+=`<div style="color:#8696a0;font-size:12px">מק"ט: ${catalog}</div>`;
        html+=`<div id="${id}" style="display:none;margin-top:8px;border-top:1px solid #284461;padding-top:8px">`;
        items.forEach(item=>{
          const isNew=item.first_seen===today;
          const receivedMap=d.received||{};
          html+=`<div style="padding:5px 0;border-bottom:1px solid #1a252e;font-size:13px">`;
          if(item.serial){
            const safeId='cam_'+item.serial.replace(/[^a-zA-Z0-9]/g,'_');
            const rcvId='rcv_'+item.serial.replace(/[^a-zA-Z0-9]/g,'_');
            const escSerial=item.serial.replace(/'/g,"\\'");
            const escDesc=(desc||'').replace(/'/g,"\\'");
            const rcvEntry=receivedMap[item.serial];
            const rcvStatus=rcvEntry&&rcvEntry.date===today?rcvEntry.status:(rcvEntry===today?'received':null);
            html+=`<div style="display:flex;justify-content:space-between;align-items:center;gap:6px">
              <span>🔢 ${item.serial}</span>
              <div style="display:flex;gap:5px;align-items:center">
                ${isNew?(rcvStatus==='received'?
                  `<button disabled style="background:#1f7a4a;border:1px solid #1f7a4a;color:#e9edef;border-radius:8px;padding:3px 10px;font-size:12px;white-space:nowrap">✅ קיבלתי</button>`:
                  rcvStatus==='not_received'?
                  `<button disabled style="background:#7a1f1f;border:1px solid #7a1f1f;color:#e9edef;border-radius:8px;padding:3px 10px;font-size:12px;white-space:nowrap">❌ לא קיבלתי</button>`:
                  `<button id="${rcvId}_yes" onclick="event.stopPropagation();markRcv('${escSerial}','${escDesc}','${rcvId}','received')"
                    style="background:#1a3a5c;border:1px solid #3d5166;color:#e9edef;border-radius:8px;padding:3px 10px;cursor:pointer;font-size:12px;white-space:nowrap">קיבלתי</button>
                   <button id="${rcvId}_no" onclick="event.stopPropagation();markRcv('${escSerial}','${escDesc}','${rcvId}','not_received')"
                    style="background:#3a1a1a;border:1px solid #7a3a3a;color:#e9edef;border-radius:8px;padding:3px 10px;cursor:pointer;font-size:12px;white-space:nowrap">לא קיבלתי</button>`
                ):''}
                <button class="return-cam-btn" id="${safeId}"
                  onclick="event.stopPropagation();takeReturnPhoto('${escSerial}','${escDesc}')"
                  style="display:none;background:#223d58;border:1px solid #284461;color:#ddeeff;border-radius:8px;padding:3px 10px;cursor:pointer;font-size:15px">📷</button>
              </div>
            </div>`;
          } else {
            const maxQty=item.amount||1;
            const nsId='ns_'+id+'_'+items.indexOf(item);
            const safeNsId=nsId.replace(/[^a-zA-Z0-9]/g,'_');
            const escDescNS=(desc||'').replace(/'/g,"\\'");
            html+=`<div style="display:flex;justify-content:space-between;align-items:center">
              <span style="color:#8696a0">כמות: ${maxQty}</span>
              <div class="return-qty-row" id="qrow_${safeNsId}" style="display:none;gap:6px;align-items:center">
                <span style="color:#8696a0;font-size:11px">להחזיר:</span>
                <input type="number" id="qty_${safeNsId}" min="0" max="${maxQty}" value="0"
                  onclick="event.stopPropagation()"
                  oninput="event.stopPropagation();markNsReturn('${safeNsId}','${escDescNS}',${maxQty})"
                  style="width:50px;background:#223d58;border:1px solid #284461;border-radius:6px;color:#ddeeff;padding:3px 5px;font-size:14px;text-align:center">
                <button id="cam_${safeNsId}"
                  onclick="event.stopPropagation();takeReturnPhoto('${safeNsId}','${escDescNS}')"
                  style="background:#223d58;border:1px solid #284461;color:#ddeeff;border-radius:8px;padding:3px 9px;cursor:pointer;font-size:15px">📷</button>
              </div>
            </div>`;
          }
          html+=`<div style="color:${isNew?'#53bdeb':'#8696a0'}">📅 ${item.first_seen} (${item.days_held} ימים)${isNew?' ← חדש היום':''}</div>`;
          html+=`</div>`;
        });
        html+=`</div></div>`;
      }
    }

    html+=`<div style="margin:14px 0"><button class="qb" onclick="refreshInventory()">🔄 רענן מלאי</button></div>`;
    html+=`<div id="returns-history-box"></div>`;
    document.getElementById('inv-content').innerHTML=html;
    const invPanel=document.getElementById('inv-panel');
    invPanel.dataset.loaded='1';
    loadReturnHistory(d.items||[]);
  }catch(e){
    document.getElementById('inv-content').innerHTML='<div style="color:#f15c6e;text-align:center;padding:30px">❌ שגיאה בטעינה</div>';
  }
}

function toggleInvAcc(id){
  const el=document.getElementById(id);
  el.style.display=el.style.display==='none'?'block':'none';
}

async function markRcv(serial, desc, rcvId, status){
  const btnYes=document.getElementById(rcvId+'_yes');
  const btnNo=document.getElementById(rcvId+'_no');
  if(btnYes) btnYes.disabled=true;
  if(btnNo)  btnNo.disabled=true;
  try{
    await apiFetch('/api/inventory/received',{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({serial, desc, status})
    });
    if(status==='received'){
      if(btnYes){ btnYes.textContent='✅ קיבלתי'; btnYes.style.background='#1f7a4a'; btnYes.style.borderColor='#1f7a4a'; }
      if(btnNo)  btnNo.style.display='none';
    } else {
      if(btnNo){ btnNo.textContent='❌ לא קיבלתי'; btnNo.style.background='#7a1f1f'; btnNo.style.borderColor='#7a1f1f'; }
      if(btnYes) btnYes.style.display='none';
    }
  }catch(e){
    if(btnYes){ btnYes.disabled=false; }
    if(btnNo)  { btnNo.disabled=false; }
  }
}

/* ── החזרת ציוד ── */
let _returnMode=false;
let _pendingReturns={};
let _invSerials=[];  // [{serial, desc}] — מאוכלס בטעינת מלאי

function toggleReturnMode(){
  const btn=document.getElementById('return-mode-btn');
  if(!btn) return;
  if(!_returnMode){
    _returnMode=true;
    _pendingReturns={};
    btn.textContent='✅ סיימתי';
    btn.style.background='#005c4b';
    btn.style.borderColor='#25d366';
    // הוסף כפתור ביטול
    let cancelBtn=document.getElementById('return-cancel-btn');
    if(!cancelBtn){
      cancelBtn=document.createElement('button');
      cancelBtn.id='return-cancel-btn';
      cancelBtn.className='qb';
      cancelBtn.textContent='❌ ביטול';
      cancelBtn.style.cssText='font-size:13px;padding:6px 14px;background:#223d58;border-color:#f15c6e;color:#f15c6e';
      cancelBtn.onclick=cancelReturnMode;
      btn.parentNode.insertBefore(cancelBtn,btn.nextSibling);
    }
    cancelBtn.style.display='inline-block';
    const scanBtn=document.getElementById('scan-mac-btn');
    if(scanBtn) scanBtn.style.display='inline-block';
    document.querySelectorAll('.return-cam-btn').forEach(b=>b.style.display='inline-block');
    document.querySelectorAll('.return-qty-row').forEach(r=>r.style.display='flex');
  } else {
    submitReturns();
  }
}

function cancelReturnMode(){
  _returnMode=false;
  _pendingReturns={};
  const btn=document.getElementById('return-mode-btn');
  if(btn){btn.textContent='📦 החזרת ציוד';btn.style.background='';btn.style.borderColor='';}
  const cancelBtn=document.getElementById('return-cancel-btn');
  if(cancelBtn) cancelBtn.style.display='none';
  document.querySelectorAll('.return-cam-btn').forEach(b=>{
    b.style.display='none';b.textContent='📷';b.style.background='#223d58';b.style.borderColor='#284461';
  });
  document.querySelectorAll('.return-qty-row').forEach(r=>{
    r.style.display='none';
    const inp=r.querySelector('input[type="number"]');
    if(inp) inp.value='0';
    const camBtn=r.querySelector('button');
    if(camBtn){camBtn.textContent='📷';camBtn.style.background='#223d58';camBtn.style.borderColor='#284461';}
  });
}

async function submitReturns(){
  const btn=document.getElementById('return-mode-btn');
  if(btn){btn.disabled=true;btn.textContent='שולח...';}
  // אסוף גם פריטי כמות (NOSN) שסומנו
  document.querySelectorAll('.return-qty-row').forEach(row=>{
    const inp=row.querySelector('input[type="number"]');
    if(!inp) return;
    const qty=parseInt(inp.value)||0;
    if(qty<=0) return;
    const nsId=row.id.replace('qrow_','');
    if(!_pendingReturns[nsId]){
      // מצא שם מהתיאור בכפתור המצלמה
      const camBtn=row.querySelector('button');
      _pendingReturns[nsId]={description:'',photo:'',quantity:qty};
    } else {
      _pendingReturns[nsId].quantity=qty;
    }
  });
  const entries=Object.entries(_pendingReturns);
  let ok=0;
  for(const [serial,data] of entries){
    try{
      await apiFetch('/api/save-return',{method:'POST',headers:{'Content-Type':'application/json'},
        body:JSON.stringify({serial,description:data.description,photo:data.photo||'',quantity:data.quantity||1})});
      ok++;
    }catch(e){}
  }
  _returnMode=false;
  _pendingReturns={};
  if(btn){btn.disabled=false;btn.textContent='📦 החזרת ציוד';btn.style.background='';btn.style.borderColor='';}
  const cancelBtn=document.getElementById('return-cancel-btn');
  if(cancelBtn) cancelBtn.style.display='none';
  const scanBtnS=document.getElementById('scan-mac-btn');
  if(scanBtnS) scanBtnS.style.display='none';
  document.querySelectorAll('.return-cam-btn').forEach(b=>{
    b.style.display='none';b.textContent='📷';b.style.background='#223d58';b.style.borderColor='#284461';
  });
  document.querySelectorAll('.return-qty-row').forEach(r=>{
    r.style.display='none';
    const inp=r.querySelector('input[type="number"]');
    if(inp) inp.value='0';
    const camBtn=r.querySelector('button');
    if(camBtn){camBtn.textContent='📷';camBtn.style.background='#223d58';camBtn.style.borderColor='#284461';}
  });
  if(ok>0){
    loadReturnHistory([]);
    const box=document.getElementById('returns-history-box');
    if(box) box.scrollIntoView({behavior:'smooth'});
  }
}

/* ── סריקת MAC/סריאל עם OCR ── */
function normalizeSerial(s){ return (s||'').replace(/[:\-\.\s]/g,'').toUpperCase(); }

function findSerialInOcr(ocrText){
  const cleanFull = normalizeSerial(ocrText);
  // חפש רצפים hex של 6+ תווים בטקסט
  const hexSeqs = cleanFull.match(/[0-9A-F]{6,}/g) || [];
  // נסה גם עם נקודותיים (MAC בפורמט XX:XX:XX:...)
  const macPatterns = ocrText.toUpperCase().match(/([0-9A-F]{2}[:\-]){5}[0-9A-F]{2}/g) || [];
  const normMacs = macPatterns.map(normalizeSerial);
  const allCandidates = [...hexSeqs, ...normMacs];
  for(const item of _invSerials){
    const ns = normalizeSerial(item.serial);
    if(!ns) continue;
    // התאמה מלאה
    if(allCandidates.some(c => c===ns || c.includes(ns) || ns.includes(c))) return item;
    // התאמה חלקית (8+ תווים מתוך הסריאל)
    if(ns.length>=8 && allCandidates.some(c => c.length>=8 && ns.includes(c.substring(0,8)))) return item;
  }
  return null;
}

async function scanEquipment(){
  // טען Tesseract.js בעת הצורך
  if(!window.Tesseract){
    const scr=document.createElement('script');
    scr.src='https://cdn.jsdelivr.net/npm/tesseract.js@4/dist/tesseract.min.js';
    document.head.appendChild(scr);
    await new Promise((res,rej)=>{ scr.onload=res; scr.onerror=rej; });
  }
  const input=document.createElement('input');
  input.type='file'; input.accept='image/*'; input.capture='environment';
  input.onchange=async(e)=>{
    const file=e.target.files[0]; if(!file) return;
    // הצג מודל טעינה
    const modal=document.getElementById('scan-modal');
    const preview=document.getElementById('scan-preview');
    const status=document.getElementById('scan-status');
    const result=document.getElementById('scan-result');
    modal.style.display='flex';
    preview.style.display='none';
    result.innerHTML='';
    status.innerHTML='<div style="color:#53bdeb;font-size:15px">🔍 מנתח תמונה...</div>';
    try{
      // הצג preview
      const objUrl=URL.createObjectURL(file);
      preview.src=objUrl; preview.style.display='block';
      // הרץ OCR
      const ocr=await Tesseract.recognize(objUrl,'eng',{logger:()=>{}});
      const ocrText=ocr.data.text;
      URL.revokeObjectURL(objUrl);
      // כווץ תמונה לאחסון
      const reader=new FileReader();
      reader.onload=(ev)=>{
        const img=new Image();
        img.onload=()=>{
          const MAX=800; let w=img.width,h=img.height;
          if(w>MAX||h>MAX){ if(w>h){h=Math.round(h*MAX/w);w=MAX;}else{w=Math.round(w*MAX/h);h=MAX;} }
          const c=document.createElement('canvas'); c.width=w; c.height=h;
          c.getContext('2d').drawImage(img,0,0,w,h);
          const photoData=c.toDataURL('image/jpeg',0.75);
          const matched=findSerialInOcr(ocrText);
          if(matched){
            status.innerHTML=`<div style="color:#25d366;font-size:15px">✅ זוהה!</div>
              <div style="font-weight:700;font-size:14px;margin-top:4px">${matched.serial}</div>
              <div style="color:#8696a0;font-size:12px">${matched.desc}</div>`;
            result.innerHTML=`
              <button onclick="confirmScan('${matched.serial.replace(/'/g,"\\'")}','${(matched.desc||'').replace(/'/g,"\\'")}',\`${photoData.replace(/`/g,'\\`')}\`)"
                style="width:100%;padding:12px;background:#25d366;border:none;color:#111b21;border-radius:10px;font-size:15px;font-weight:700;cursor:pointer;margin-top:10px">
                ✅ אשר החזרה</button>
              <button onclick="closeScanModal()"
                style="width:100%;padding:10px;background:#223d58;border:1px solid #284461;color:#ddeeff;border-radius:10px;font-size:14px;cursor:pointer;margin-top:6px">
                ביטול</button>`;
          } else {
            status.innerHTML=`
              <div style="font-size:40px;margin-bottom:8px">⚠️</div>
              <div style="color:#f15c6e;font-size:16px;font-weight:700">ציוד לא מזוהה</div>
              <div style="color:#8696a0;font-size:13px;margin-top:6px">הציוד שצולם אינו שייך למחסן שלך</div>`;
            result.innerHTML=`<button onclick="closeScanModal()"
              style="width:100%;padding:12px;background:#3a1a1a;border:1px solid #7a3a3a;color:#f15c6e;border-radius:10px;font-size:15px;font-weight:700;cursor:pointer;margin-top:10px">
              סגור</button>`;
          }
        };
        img.src=ev.target.result;
      };
      reader.readAsDataURL(file);
    }catch(err){
      status.innerHTML='<div style="color:#f15c6e">❌ שגיאה בזיהוי</div>';
      document.getElementById('scan-result').innerHTML=`<button onclick="closeScanModal()"
        style="width:100%;padding:10px;background:#223d58;border:1px solid #284461;color:#ddeeff;border-radius:10px;font-size:14px;cursor:pointer;margin-top:8px">סגור</button>`;
    }
  };
  input.click();
}

function confirmScan(serial, desc, photoData){
  _pendingReturns[serial]={description:desc, photo:photoData};
  const safeId='cam_'+serial.replace(/[^a-zA-Z0-9]/g,'_');
  const camBtn=document.getElementById(safeId);
  if(camBtn){camBtn.textContent='✅';camBtn.style.background='#005c4b';camBtn.style.borderColor='#25d366';}
  closeScanModal();
  const toast=document.createElement('div');
  toast.style.cssText='position:fixed;bottom:80px;left:50%;transform:translateX(-50%);background:#25d366;color:#111b21;padding:10px 22px;border-radius:20px;font-weight:700;z-index:9999;font-size:14px';
  toast.textContent=`✅ ${serial} נוסף להחזרה`;
  document.body.appendChild(toast);
  setTimeout(()=>toast.remove(),3000);
}

function closeScanModal(){
  document.getElementById('scan-modal').style.display='none';
}

function takeReturnPhoto(serial,desc){
  const input=document.createElement('input');
  input.type='file';input.accept='image/*';input.capture='environment';
  input.onchange=(e)=>{
    const file=e.target.files[0];
    if(!file) return;
    const reader=new FileReader();
    reader.onload=(ev)=>{
      const img=new Image();
      img.onload=()=>{
        const MAX=800;
        let w=img.width,h=img.height;
        if(w>MAX||h>MAX){
          if(w>h){h=Math.round(h*MAX/w);w=MAX;}
          else{w=Math.round(w*MAX/h);h=MAX;}
        }
        const c=document.createElement('canvas');
        c.width=w;c.height=h;
        c.getContext('2d').drawImage(img,0,0,w,h);
        const photoData=c.toDataURL('image/jpeg',0.75);
        _pendingReturns[serial]={description:desc,photo:photoData};
        const safeId='cam_'+serial.replace(/[^a-zA-Z0-9]/g,'_');
        const camBtn=document.getElementById(safeId);
        if(camBtn){camBtn.textContent='✅';camBtn.style.background='#005c4b';camBtn.style.borderColor='#25d366';}
      };
      img.src=ev.target.result;
    };
    reader.readAsDataURL(file);
  };
  input.click();
}

/* ── דוח ציוד ── */
let _eqActiveTab='rcv';

function switchEqTab(t){
  _eqActiveTab=t;
  const btnRcv=document.getElementById('eqrep-tab-rcv');
  const btnRet=document.getElementById('eqrep-tab-ret');
  const listRcv=document.getElementById('eqrep-rcv-list');
  const listRet=document.getElementById('eqrep-ret-list');
  if(t==='rcv'){
    btnRcv.style.background='#1f7a4a'; btnRcv.style.color='#fff'; btnRcv.style.fontWeight='700';
    btnRet.style.background='#223d58'; btnRet.style.color='#6b94b8'; btnRet.style.fontWeight='400';
    listRcv.style.display='block'; listRet.style.display='none';
  } else {
    btnRet.style.background='#c0392b'; btnRet.style.color='#fff'; btnRet.style.fontWeight='700';
    btnRcv.style.background='#223d58'; btnRcv.style.color='#6b94b8'; btnRcv.style.fontWeight='400';
    listRcv.style.display='none'; listRet.style.display='block';
  }
}

async function loadEqReport(){
  const listRcv=document.getElementById('eqrep-rcv-list');
  const listRet=document.getElementById('eqrep-ret-list');
  listRcv.innerHTML=listRet.innerHTML='<div style="text-align:center;padding:20px;color:#8696a0">טוען...</div>';
  try{
    const r=await apiFetch('/api/equipment-report');
    const d=await r.json();
    renderEqList(listRcv, d.received_by_date||{}, 'received');
    renderEqList(listRet, d.returned_by_date||{}, 'returned');
  }catch(e){
    listRcv.innerHTML=listRet.innerHTML='<div style="color:#f15c6e;text-align:center;padding:20px">❌ שגיאה</div>';
  }
}

function renderEqList(container, byDate, type){
  const dates=Object.keys(byDate).sort().reverse();
  if(!dates.length){
    container.innerHTML='<div style="text-align:center;padding:30px;color:#8696a0">אין פריטים</div>';
    return;
  }
  const icon=type==='received'?'📥':'📤';
  let html='';
  for(const d of dates){
    const items=byDate[d];
    const encDate=encodeURIComponent(d);
    const dayId=`eq_${type}_${d.replace(/-/g,'')}`;
    html+=`<div id="card_${dayId}" style="background:#d0e4f7;border-radius:10px;margin-bottom:10px;overflow:hidden">
      <div style="display:flex;justify-content:space-between;align-items:center;padding:10px 12px;border-bottom:1px solid #284461;cursor:pointer"
           onclick="toggleEqDay('${dayId}')">
        <span style="font-weight:700;font-size:15px">${icon} ${d} <span style="color:#8696a0;font-size:12px;font-weight:400">(${items.length} פריטים)</span></span>
        <div style="display:flex;gap:6px;align-items:center" onclick="event.stopPropagation()">
          <a href="/api/equipment-report/export?date=${encDate}&type=${type}&token=${encodeURIComponent(localStorage.getItem('cc_token')||'')}&phone=${encodeURIComponent(localStorage.getItem('cc_phone')||'')}"
             style="background:#1a3a5c;border:1px solid #3d5166;color:#e9edef;border-radius:8px;padding:4px 10px;font-size:12px;text-decoration:none;white-space:nowrap">
            📥 Excel</a>
          <button onclick="deleteEqDate('${d}','${type}','card_${dayId}')"
            style="background:#3a1a1a;border:1px solid #7a3a3a;color:#f15c6e;border-radius:8px;padding:4px 9px;font-size:13px;cursor:pointer"
            title="מחק דוח">🗑️</button>
        </div>
      </div>
      <div id="eq_${type}_${d.replace(/-/g,'')}" style="display:none;padding:8px 12px">`;
    for(const item of items){
      html+=`<div style="display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #1a252e;font-size:13px">
        <span style="color:#e9edef">${item.serial||'—'}</span>
        <span style="color:#8696a0;font-size:12px">${item.desc||item.description||''}</span>
        ${item.quantity&&item.quantity>1?`<span style="color:#53bdeb">×${item.quantity}</span>`:''}
      </div>`;
    }
    html+=`</div></div>`;
  }
  container.innerHTML=html;
}

function toggleEqDay(id){
  const el=document.getElementById(id);
  if(el) el.style.display=el.style.display==='none'?'block':'none';
}

async function deleteEqDate(date, type, cardId){
  if(!confirm(`למחוק את הדוח של ${date}?`)) return;
  try{
    const r=await apiFetch('/api/equipment-report/delete-date',{
      method:'DELETE',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({date, type})
    });
    const d=await r.json();
    if(d.ok){
      const card=document.getElementById(cardId);
      if(card) card.remove();
    }
  }catch(e){ alert('שגיאה במחיקה'); }
}

async function loadReturnHistory(currentItems){
  const box=document.getElementById('returns-history-box');
  if(!box) return;
  try{
    const r=await apiFetch('/api/returns?days=30');
    if(!r.ok) return;
    const d=await r.json();
    if(!d.returns||d.returns.length===0){box.innerHTML='';return;}
    const filtered=d.returns;
    if(filtered.length===0){box.innerHTML='';return;}
    let html=`<div style="margin-top:18px;border-top:2px solid #284461;padding-top:14px">
      <div style="font-weight:700;color:#f15c6e;margin-bottom:10px;font-size:15px">📤 הוחזר (30 ימים אחרונים)</div>`;
    filtered.forEach(ret=>{
      const isNS=ret.serial.startsWith('ns_')||ret.serial.startsWith('NOSN');
      const serialLine=isNS
        ?`<span style="font-weight:700">📦 ${ret.description||'ציוד'} × ${ret.quantity||1}</span>`
        :`<span style="font-weight:700">🔢 ${ret.serial}</span>`;
      const descLine=isNS?'':`<div style="color:#8696a0;margin-top:2px">${ret.description||''}</div>`;
      const qtyLine=(!isNS&&(ret.quantity||1)>1)?`<div style="color:#8696a0;font-size:12px">כמות: ${ret.quantity}</div>`:'';
      html+=`<div id="ret-row-${ret.id}" style="background:#d0e4f7;border-radius:8px;padding:10px 14px;margin-bottom:8px;font-size:13px;direction:rtl;border-right:3px solid #f15c6e">
        <div style="display:flex;justify-content:space-between;align-items:center">
          ${serialLine}
          <div style="display:flex;gap:8px;align-items:center">
            <span style="color:#8696a0;font-size:12px">📅 ${ret.return_date}</span>
            <button onclick="deleteReturn(${ret.id})" style="background:none;border:none;color:#f15c6e;font-size:18px;cursor:pointer;padding:0 4px;line-height:1" title="מחק">🗑</button>
          </div>
        </div>
        ${descLine}${qtyLine}
        ${ret.has_photo?`<div style="margin-top:4px"><a href="/api/return-photo/${ret.photo_filename}" target="_blank" style="color:#53bdeb;font-size:12px">📷 צפה בתמונה</a></div>`:''}
      </div>`;
    });
    html+=`</div>`;
    box.innerHTML=html;
  }catch(e){box.innerHTML='';}
}

function markNsReturn(safeId,desc,maxQty){
  const inp=document.getElementById('qty_'+safeId);
  if(!inp) return;
  let qty=parseInt(inp.value)||0;
  if(qty<0){qty=0;inp.value=0;}
  if(qty>maxQty){qty=maxQty;inp.value=maxQty;}
  if(qty>0){
    if(!_pendingReturns[safeId]) _pendingReturns[safeId]={description:desc,photo:''};
    _pendingReturns[safeId].quantity=qty;
    _pendingReturns[safeId].description=desc;
  } else {
    delete _pendingReturns[safeId];
  }
}

async function deleteReturn(id){
  if(!confirm('למחוק את ההחזרה הזו?')) return;
  try{
    const r=await apiFetch('/api/delete-return/'+id,{method:'DELETE'});
    if(r.ok){
      const row=document.getElementById('ret-row-'+id);
      if(row) row.remove();
    }
  }catch(e){}
}

/* ── reports ── */
function isoDay(offset){
  const d=new Date(); d.setDate(d.getDate()+offset);
  // תאריך מקומי (לא UTC) — חשוב לאזור ישראל UTC+3
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}
function setRepDay(offset){
  const d=isoDay(offset);
  document.getElementById('rep-from').value=d;
  document.getElementById('rep-to').value=d;
  // עדכן סימון כפתור פעיל
  [-1,0,1].forEach(o=>{
    const btn=document.getElementById('rep-day-btn-'+o);
    if(!btn) return;
    if(o===offset){
      btn.style.background='#25d366'; btn.style.color='#fff';
      btn.style.border='none'; btn.style.fontWeight='700';
    } else {
      btn.style.background='#223d58'; btn.style.color='#ddeeff';
      btn.style.border='1px solid #284461'; btn.style.fontWeight='400';
    }
  });
  loadReports();
}

function initReports(){
  const panel=document.getElementById('rep-panel');
  const today=new Date().toISOString().slice(0,10);
  const ago30=new Date(Date.now()-29*86400000).toISOString().slice(0,10);
  panel.dataset.loaded='1';
  panel.innerHTML=`
    <div style="background:#1a2d42;border-radius:10px;padding:14px;margin-bottom:12px">
      <div style="font-weight:700;margin-bottom:10px;color:#25d366">📊 דוחות ביקורים</div>
      <div style="display:flex;gap:6px;margin-bottom:10px">
        <button id="rep-day-btn--1" onclick="setRepDay(-1)" style="flex:1;background:#223d58;border:1px solid #284461;border-radius:8px;padding:7px 4px;color:#ddeeff;font-size:13px;cursor:pointer">אתמול</button>
        <button id="rep-day-btn-0"  onclick="setRepDay(0)"  style="flex:1;background:#223d58;border:1px solid #284461;border-radius:8px;padding:7px 4px;color:#ddeeff;font-size:13px;cursor:pointer">היום</button>
        <button id="rep-day-btn-1"  onclick="setRepDay(1)"  style="flex:1;background:#223d58;border:1px solid #284461;border-radius:8px;padding:7px 4px;color:#ddeeff;font-size:13px;cursor:pointer">מחר</button>
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
        <div style="flex:1;min-width:120px">
          <div style="font-size:12px;color:#6b94b8;margin-bottom:4px">מתאריך</div>
          <input id="rep-from" type="date" value="${ago30}" style="width:100%;background:#223d58;border:1px solid #284461;border-radius:8px;padding:8px;color:#ddeeff;font-size:14px">
        </div>
        <div style="flex:1;min-width:120px">
          <div style="font-size:12px;color:#6b94b8;margin-bottom:4px">עד תאריך</div>
          <input id="rep-to" type="date" value="${today}" style="width:100%;background:#223d58;border:1px solid #284461;border-radius:8px;padding:8px;color:#ddeeff;font-size:14px">
        </div>
        <button class="btn-main" style="margin-top:18px;padding:9px 18px;font-size:14px;width:auto" onclick="loadReports()">הצג</button>
        <button class="btn-main" id="rep-export-btn" style="margin-top:18px;padding:9px 14px;font-size:14px;width:auto;background:#1a3a5c;border:1px solid #3d5166;display:none" onclick="exportReports()">📥 Excel</button>
        <button class="btn-main" id="rep-map-btn" style="margin-top:18px;padding:9px 14px;font-size:14px;width:auto;background:#1a3a5c;border:1px solid #3d5166;display:none" onclick="showReportsMap()">🗺️ מפה</button>
      </div>
    </div>
    <div id="rep-content"></div>`;
}

async function loadReports(){
  const from=document.getElementById('rep-from').value;
  const to=document.getElementById('rep-to').value;
  const box=document.getElementById('rep-content');
  if(!from||!to){box.innerHTML='<div style="color:#f15c6e">בחר תאריכים</div>';return;}
  const fromD=new Date(from), toD=new Date(to);
  const days=Math.round((toD-fromD)/86400000)+1;
  box.innerHTML=`<div style="color:#8696a0;text-align:center;padding:20px">טוען... (${days} ימים)</div>`;
  try{
    const r=await apiFetch('/api/reports?days='+days+'&end='+to);
    if(r.status===401){box.innerHTML='<div style="color:#f15c6e">❌ לא מחובר</div>';return;}
    const d=await r.json();
    if(!d.visits||d.visits.length===0){box.innerHTML='<div style="color:#8696a0;text-align:center;padding:20px">לא נמצאו ביקורים בטווח זה</div>';return;}

    // שמור ביקורים גלובלית למפה
    _lastVisits=d.visits;

    // הצג כפתורי ייצוא ומפה
    const expBtn=document.getElementById('rep-export-btn');
    if(expBtn) expBtn.style.display='inline-block';
    const mapBtn=document.getElementById('rep-map-btn');
    if(mapBtn) mapBtn.style.display='inline-block';

    const hasEarnings=d.total_earnings>0;
    let html=`<div style="background:#1a2d42;border-radius:10px;padding:12px;margin-bottom:10px">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
        <div style="font-weight:700;color:#25d366">סיכום — ${d.total} ביקורים</div>
        ${hasEarnings?`<div style="font-weight:700;font-size:16px;color:#25d366">₪${d.total_earnings.toLocaleString('he-IL',{minimumFractionDigits:0,maximumFractionDigits:2})}</div>`:''}
      </div>`;
    for(const [type,info] of Object.entries(d.summary)){
      const pct=Math.round(info.count/d.total*100);
      html+=`<div style="display:flex;justify-content:space-between;font-size:13px;padding:3px 0;border-bottom:1px solid #284461">
        <span>${type} <span style="color:#6b94b8">${info.count} (${pct}%)</span></span>
        ${info.earnings>0?`<span style="color:#25d366">₪${info.earnings.toLocaleString('he-IL',{minimumFractionDigits:0,maximumFractionDigits:2})}</span>`:'<span style="color:#7a9ab0">ללא מחיר</span>'}
      </div>`;
    }
    if(!hasEarnings) html+=`<div style="color:#6b94b8;font-size:12px;margin-top:8px">💡 הגדר מחירים בלשונית 💰 מחירון</div>`;
    html+=`</div>`;

    // קיבוץ לפי תאריך
    const byDate={};
    d.visits.forEach(v=>{if(!byDate[v.fetch_date])byDate[v.fetch_date]=[];byDate[v.fetch_date].push(v);});
    for(const [dt, rows] of Object.entries(byDate)){
      const dateObj=new Date(dt+'T12:00:00');
      const dayName=['ראשון','שני','שלישי','רביעי','חמישי','שישי','שבת'][dateObj.getDay()];
      const dayTotal=rows.reduce((s,v)=>s+(v.price||0),0);
      html+=`<div style="display:flex;justify-content:space-between;align-items:center;margin:12px 0 6px">
        <div style="font-weight:700;color:#53bdeb;font-size:13px">יום ${dayName} ${dt} (${rows.length})</div>
        <div id="day_total_${dt}" style="font-weight:700;color:#25d366;font-size:14px">${dayTotal>0?`₪${dayTotal.toLocaleString('he-IL',{minimumFractionDigits:0,maximumFractionDigits:2})}`:''}
        </div>
      </div>`;
      rows.forEach((v,vi)=>{
        const isVip=v.is_vip==='כן';
        const ppc=d.prices_map?d.prices_map[(v.task_type||'')+'__PPC']||0:0;
        const lsb=d.prices_map?d.prices_map[(v.task_type||'')+'__LSB']||0:0;
        const cardId=`vc_${dt}_${vi}`;
        const callId=v.call_id||'';
        const isIncomplete=v.is_incomplete===true;
        const isCustom=v.price_type==='custom';
        const isFiber=((v.task_type||'').includes('תשתית סיבים')||(v.task_type||'').includes('טריפל סיבים'))&&!(v.task_type||'').includes('תקלת גלישה');
        html+=`<div class="inv-card ok" data-dt="${dt}" data-price="${v.price||0}" style="margin:4px 0;border-right-color:${isIncomplete?'#f15c6e':isVip?'#f15c6e':'#25d366'}" id="${cardId}">
          <div style="display:flex;justify-content:space-between;align-items:center">
            <span style="font-weight:700;font-size:13px">${v.task_type||'—'}</span>
            <span style="display:flex;gap:6px;align-items:center">
              ${isVip?'<span class="inv-badge badge-blocked">VIP</span>':''}
              ${isIncomplete
                ?'<span style="color:#f15c6e;font-weight:700;font-size:13px">לא הושלם ₪0</span>'
                :isCustom
                  ?`<span id="${cardId}_price" style="color:#f9c846;font-weight:700;font-size:13px">✏️ ₪${v.price}</span>`
                  :(isFiber&&(ppc||lsb))?`
                <button id="${cardId}_ppc" onclick="setPriceType('${cardId}','PPC',${ppc},'${dt}','${callId}')"
                  style="background:${v.price_type==='PPC'?'#005c4b':'#1a3a5c'};border:1px solid ${v.price_type==='PPC'?'#25d366':'#3d5166'};color:${v.price_type==='PPC'?'#25d366':'#53bdeb'};border-radius:6px;padding:3px 8px;font-size:12px;cursor:pointer">
                  ${v.price_type==='PPC'?'✅ ':''}PPC ₪${ppc}</button>
                <button id="${cardId}_lsb" onclick="setPriceType('${cardId}','LSB',${lsb},'${dt}','${callId}')"
                  style="background:${v.price_type==='LSB'?'#005c4b':'#1a3a5c'};border:1px solid ${v.price_type==='LSB'?'#25d366':'#3d5166'};color:${v.price_type==='LSB'?'#25d366':'#53bdeb'};border-radius:6px;padding:3px 8px;font-size:12px;cursor:pointer">
                  ${v.price_type==='LSB'?'✅ ':''}LSB ₪${lsb}</button>`
                  :(v.price>0?`<span id="${cardId}_price" style="color:#25d366;font-size:13px">₪${v.price}</span>`:'')
              }
              ${!isIncomplete?`<button onclick="showOverride('${cardId}','${callId}','${dt}',${v.price||0})"
                style="background:transparent;border:1px solid #6b94b8;color:#6b94b8;border-radius:6px;padding:2px 7px;font-size:11px;cursor:pointer">✏️ חריגה</button>`:''}
            </span>
          </div>
          <div style="color:#0d1b2a;font-size:13px">${v.contact_name||''}</div>
          <div style="color:#4a6a88;font-size:12px">📍 ${v.street||''} ${v.city||''}</div>
          ${v.appt_start?`<div style="color:#4a6a88;font-size:12px">⏰ ${v.appt_start}–${v.appt_finish}</div>`:''}
          ${v.status&&!isIncomplete?`<div style="color:#4a6a88;font-size:12px">סטטוס: ${v.status}</div>`:''}
          ${v.infrastructure?`<div style="color:#4a6a88;font-size:12px">🌐 ${v.infrastructure}</div>`:''}
          <div id="${cardId}_ovr" style="font-size:12px;display:none;margin-top:6px;gap:6px;align-items:center">
            <input type="number" id="${cardId}_ovr_inp" placeholder="סכום חריגה ₪" min="0" step="0.01"
              style="background:#223d58;border:1px solid #284461;border-radius:8px;padding:5px 8px;color:#ddeeff;font-size:13px;width:130px">
            <button onclick="saveOverride('${cardId}','${callId}','${dt}')"
              style="background:#3dba6f;border:none;border-radius:8px;padding:5px 10px;color:#fff;font-size:12px;cursor:pointer">שמור</button>
            <button onclick="clearOverride('${cardId}','${callId}','${dt}')"
              style="background:transparent;border:1px solid #f15c6e;border-radius:8px;padding:5px 10px;color:#f15c6e;font-size:12px;cursor:pointer">בטל</button>
          </div>
          <div id="${cardId}_sel" style="font-size:12px;color:#25d366;margin-top:4px"></div>
        </div>`;
      });
    }
    box.innerHTML=html;
  }catch(e){box.innerHTML='<div style="color:#f15c6e;text-align:center;padding:20px">❌ שגיאה בטעינה</div>';}
}

function exportReports(){
  const from=document.getElementById('rep-from').value;
  const to=document.getElementById('rep-to').value;
  if(!from||!to) return;
  const fromD=new Date(from), toD=new Date(to);
  const days=Math.round((toD-fromD)/86400000)+1;
  const tk=localStorage.getItem('cc_token')||'';
  const ph=localStorage.getItem('cc_phone')||'';
  window.location.href='/api/reports/export?days='+days+'&end='+to+'&token='+encodeURIComponent(tk)+'&phone='+encodeURIComponent(ph);
}

/* ── price list ── */
async function loadPriceList(){
  const panel=document.getElementById('price-panel');
  panel.dataset.loaded='1';
  panel.innerHTML='<div style="color:#8696a0;text-align:center;padding:20px">טוען...</div>';
  try{
    const r=await apiFetch('/api/prices');
    const d=await r.json();
    const prices=d.prices||{};
    const types=d.task_types||[];
    // מיזוג — סוגים שיש מחיר אבל לא בDB
    const allTypes=[...new Set([...types,...Object.keys(prices).filter(k=>!k.includes('__PPC')&&!k.includes('__LSB'))])].sort();
    let html=`<div style="background:#1a2d42;border-radius:10px;padding:14px;margin-bottom:12px">
      <div style="font-weight:700;color:#25d366;margin-bottom:12px;font-size:15px">💰 מחירון לפי סוג משימה</div>
      <div style="color:#6b94b8;font-size:12px;margin-bottom:12px">הזן מחיר לכל סוג — ישמשו לחישוב הדוחות</div>`;
    if(allTypes.length===0){
      html+=`<div style="color:#6b94b8;font-size:13px">טען דוחות קודם כדי לראות את סוגי המשימות</div>`;
    } else {
      allTypes.forEach(t=>{
        const isFiber=(t.includes('תשתית סיבים')||t.includes('טריפל סיבים'))&&!t.includes('תקלת גלישה');
        if(isFiber){
          const valPPC=prices[t+'__PPC']||'';
          const valLSB=prices[t+'__LSB']||'';
          html+=`<div style="padding:8px 0;border-bottom:1px solid #284461">
            <div style="font-size:14px;margin-bottom:6px">${t}</div>
            <div style="display:flex;gap:8px">
              <div style="flex:1">
                <div style="font-size:11px;color:#6b94b8;margin-bottom:3px">PPC</div>
                <div style="display:flex;align-items:center;gap:4px">
                  <input type="number" id="price_${encodeURIComponent(t+'__PPC')}" value="${valPPC}" placeholder="0"
                    style="width:100%;background:#223d58;border:1px solid #284461;border-radius:8px;padding:7px 10px;color:#ddeeff;font-size:14px;text-align:left"
                    min="0" step="0.01">
                  <span style="color:#6b94b8;font-size:13px">₪</span>
                </div>
              </div>
              <div style="flex:1">
                <div style="font-size:11px;color:#6b94b8;margin-bottom:3px">LSB</div>
                <div style="display:flex;align-items:center;gap:4px">
                  <input type="number" id="price_${encodeURIComponent(t+'__LSB')}" value="${valLSB}" placeholder="0"
                    style="width:100%;background:#223d58;border:1px solid #284461;border-radius:8px;padding:7px 10px;color:#ddeeff;font-size:14px;text-align:left"
                    min="0" step="0.01">
                  <span style="color:#6b94b8;font-size:13px">₪</span>
                </div>
              </div>
            </div>
          </div>`;
        } else {
          const val=prices[t]||'';
          html+=`<div style="display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid #284461">
            <span style="font-size:14px;flex:1">${t}</span>
            <div style="display:flex;align-items:center;gap:6px">
              <input type="number" id="price_${encodeURIComponent(t)}" value="${val}" placeholder="0"
                style="width:90px;background:#223d58;border:1px solid #284461;border-radius:8px;padding:7px 10px;color:#ddeeff;font-size:14px;text-align:left"
                min="0" step="0.01">
              <span style="color:#6b94b8;font-size:13px">₪</span>
            </div>
          </div>`;
        }
      });
    }
    html+=`<button class="btn-main" style="margin-top:14px" onclick="savePrices()">💾 שמור מחירון</button>
      <div id="price-msg" style="color:#25d366;font-size:13px;text-align:center;min-height:18px;margin-top:8px"></div>
    </div>
    <div style="background:#d0e4f7;border-radius:10px;padding:12px;font-size:13px;color:#6b94b8">
      <div style="font-weight:700;color:#ddeeff;margin-bottom:6px">💡 טיפ</div>
      לאחר שמירת המחירון, הדוחות יחשבו אוטומטית סכום יומי וסה"כ לכל טווח
    </div>`;
    panel.innerHTML=html;
  }catch(e){
    panel.innerHTML='<div style="color:#f15c6e;text-align:center;padding:20px">❌ שגיאה</div>';
  }
}

async function savePrices(){
  const msg=document.getElementById('price-msg');
  const inputs=document.querySelectorAll('#price-panel input[type=number]');
  const prices={};
  inputs.forEach(inp=>{
    const key=decodeURIComponent(inp.id.replace('price_',''));
    const val=parseFloat(inp.value);
    if(!isNaN(val)&&val>0) prices[key]=val;
  });
  try{
    const r=await apiFetch('/api/prices',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({prices})});
    const d=await r.json();
    msg.textContent=d.status==='ok'?'✅ נשמר!':'❌ שגיאה';
    setTimeout(()=>msg.textContent='',3000);
  }catch(e){msg.textContent='❌ שגיאת רשת';}
}

async function setPriceType(cardId, type, price, dt, callId){
  const ppcBtn=document.getElementById(cardId+'_ppc');
  const lsbBtn=document.getElementById(cardId+'_lsb');
  if(ppcBtn&&lsbBtn){
    const active='background:#005c4b;border:1px solid #25d366;color:#25d366;border-radius:6px;padding:3px 8px;font-size:12px;cursor:pointer';
    const inactive='background:#1a3a5c;border:1px solid #3d5166;color:#53bdeb;border-radius:6px;padding:3px 8px;font-size:12px;cursor:pointer';
    if(type==='PPC'){
      ppcBtn.style.cssText=active; ppcBtn.textContent=`✅ PPC ₪${price}`;
      lsbBtn.style.cssText=inactive; lsbBtn.textContent=lsbBtn.textContent.replace('✅ ','');
    } else {
      lsbBtn.style.cssText=active; lsbBtn.textContent=`✅ LSB ₪${price}`;
      ppcBtn.style.cssText=inactive; ppcBtn.textContent=ppcBtn.textContent.replace('✅ ','');
    }
  }
  // שמור בשרת
  if(callId){
    await apiFetch('/api/reports/set-override',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({call_id:callId,fetch_date:dt,price_type:type,custom_price:null})});
  }
  recalcDayTotal(dt, cardId, price);
}

function recalcDayTotal(dt, changedCard, newPrice){
  // עדכן מחיר הכרטיסייה המשתנה ב-dataset
  if(changedCard) document.getElementById(changedCard).dataset.price=newPrice;
  // חשב מחדש סכום יומי
  const dayEl=document.getElementById('day_total_'+dt);
  if(!dayEl) return;
  let total=0;
  document.querySelectorAll('.inv-card[data-dt="'+dt+'"]').forEach(c=>{
    total+=parseFloat(c.dataset.price||0);
  });
  dayEl.textContent=total>0?`₪${total.toLocaleString('he-IL',{minimumFractionDigits:0,maximumFractionDigits:2})}`:'';
}

function showOverride(cardId, callId, dt, currentPrice){
  const ovr=document.getElementById(cardId+'_ovr');
  if(!ovr) return;
  const isVisible=ovr.style.display==='flex';
  ovr.style.display=isVisible?'none':'flex';
  if(!isVisible){
    const inp=document.getElementById(cardId+'_ovr_inp');
    if(inp){inp.value=currentPrice||'';inp.focus();}
  }
}

async function saveOverride(cardId, callId, dt){
  const inp=document.getElementById(cardId+'_ovr_inp');
  if(!inp) return;
  const val=parseFloat(inp.value);
  if(isNaN(val)||val<0){inp.style.borderColor='#f15c6e';return;}
  await apiFetch('/api/reports/set-override',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({call_id:callId,fetch_date:dt,price_type:'custom',custom_price:val})});
  // עדכן תצוגה
  const priceEl=document.getElementById(cardId+'_price');
  if(priceEl){priceEl.textContent=`✏️ ₪${val}`;priceEl.style.color='#f9c846';}
  document.getElementById(cardId+'_ovr').style.display='none';
  document.getElementById(cardId).dataset.price=val;
  recalcDayTotal(dt);
}

async function clearOverride(cardId, callId, dt){
  await apiFetch('/api/reports/set-override',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({call_id:callId,fetch_date:dt,price_type:'',custom_price:null})});
  document.getElementById(cardId+'_ovr').style.display='none';
  loadReports(); // טען מחדש לחישוב מחיר מקורי
}

async function refreshInventory(){
  document.getElementById('inv-panel').dataset.loaded='';
  loadInventory();
}

/* ══════════════════════════════
   MAP — מפת ביקורים
══════════════════════════════ */
let _lastVisits=[];
let _mapInstance=null;
let _mapLoaded=false;

function _geoCache(){try{return JSON.parse(localStorage.getItem('_geo_cache')||'{}');}catch{return {};}}
function _saveGeoCache(c){try{localStorage.setItem('_geo_cache',JSON.stringify(c));}catch{}}

async function _loadLeaflet(){
  if(_mapLoaded) return;
  await new Promise((res,rej)=>{
    const s=document.createElement('script');
    s.src='https://unpkg.com/leaflet@1.9.4/dist/leaflet.js';
    s.onload=res; s.onerror=rej;
    document.head.appendChild(s);
  });
  _mapLoaded=true;
}

async function showReportsMap(){
  const modal=document.getElementById('map-modal');
  modal.style.display='flex';
  const status=document.getElementById('map-status');
  status.textContent='טוען מפה...';

  try{ await _loadLeaflet(); }catch(e){status.textContent='שגיאה בטעינת מפה';return;}

  // אתחול מפה (פעם ראשונה בלבד)
  if(!_mapInstance){
    _mapInstance=L.map('map-container',{zoomControl:true}).setView([31.8,35.0],8);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{
      attribution:'© <a href="https://openstreetmap.org">OpenStreetMap</a>',
      maxZoom:19
    }).addTo(_mapInstance);
  } else {
    // נקה סמנים קודמים
    _mapInstance.eachLayer(l=>{if(l.options&&(l.options.radius!==undefined||l._icon)) _mapInstance.removeLayer(l);});
  }
  // גרום למפה לחשב גודל נכון אחרי שהמודל נפתח
  setTimeout(()=>_mapInstance.invalidateSize(),100);

  const bounds=[];

  // מיקום המשתמש
  status.textContent='מאתר מיקומך...';
  if(navigator.geolocation){
    try{
      const pos=await new Promise((res,rej)=>navigator.geolocation.getCurrentPosition(res,rej,{timeout:8000}));
      const ul=[pos.coords.latitude,pos.coords.longitude];
      bounds.push(ul);
      L.circleMarker(ul,{radius:13,fillColor:'#3dba6f',color:'#fff',weight:3,fillOpacity:0.95,zIndexOffset:1000})
        .addTo(_mapInstance)
        .bindPopup('<div dir="rtl"><b>📍 המיקום שלי</b></div>');
      _mapInstance.setView(ul,11);
    }catch(e){status.textContent='לא ניתן לאתר מיקום';}
  }

  if(!_lastVisits||!_lastVisits.length){
    status.textContent='אין ביקורים לתצוגה — טען דוח תחילה';
    return;
  }

  // גאוקוד כתובות (עם cache)
  const cache=_geoCache();
  let done=0, placed=0;
  const total=_lastVisits.length;

  for(const v of _lastVisits){
    const addr=((v.street||'')+' '+(v.city||'')+' ישראל').trim();
    if(!addr||addr==='ישראל'){done++;continue;}

    let lat,lng;
    if(cache[addr]){
      lat=cache[addr][0]; lng=cache[addr][1];
    } else {
      try{
        const r=await fetch('https://nominatim.openstreetmap.org/search?q='+encodeURIComponent(addr)+'&format=json&limit=1&countrycodes=il');
        const data=await r.json();
        if(data.length){
          lat=parseFloat(data[0].lat); lng=parseFloat(data[0].lon);
          cache[addr]=[lat,lng];
          _saveGeoCache(cache);
        }
      }catch(e){}
      // המתן כדי לכבד מגבלת Nominatim (1 בקשה/שנייה)
      await new Promise(r=>setTimeout(r,1100));
    }

    if(lat&&lng){
      bounds.push([lat,lng]);
      placed++;
      const isInc=v.is_incomplete===true;
      const isVip=v.is_vip==='כן';
      const col=isInc?'#f15c6e':isVip?'#f9c846':'#53bdeb';
      const priceStr=v.price>0?`<br>💰 ₪${v.price}`:'';
      L.circleMarker([lat,lng],{radius:9,fillColor:col,color:'#0d1b2a',weight:2,fillOpacity:0.9})
        .addTo(_mapInstance)
        .bindPopup(`<div dir="rtl" style="min-width:170px;font-family:Arial,sans-serif">
          <b style="font-size:13px">${v.task_type||''}</b><br>
          👤 ${v.contact_name||''}<br>
          📍 ${v.street||''} ${v.city||''}<br>
          📅 ${v.fetch_date||''} ${v.appt_start?'⏰ '+v.appt_start:''}
          ${priceStr}
          ${isInc?'<br><span style="color:#f15c6e;font-weight:700">⚠️ לא הושלם</span>':''}
          ${isVip?'<br><span style="color:#f9c846;font-weight:700">⭐ VIP</span>':''}
        </div>`);
    }

    done++;
    status.textContent=`ממפה... ${done}/${total}`;
  }

  if(bounds.length>1) _mapInstance.fitBounds(bounds,{padding:[40,40],maxZoom:14});
  status.textContent=`✅ ${placed} ביקורים על המפה`;
}

function closeMap(){
  document.getElementById('map-modal').style.display='none';
}

async function showTasksMap(){
  // מציג פקעות מבנק הפקעות על המפה
  const modal=document.getElementById('map-modal');
  modal.style.display='flex';
  const status=document.getElementById('map-status');
  status.textContent='טוען מפה...';

  try{ await _loadLeaflet(); }catch(e){status.textContent='שגיאה בטעינת מפה';return;}

  if(!_mapInstance){
    _mapInstance=L.map('map-container',{zoomControl:true}).setView([31.8,35.0],8);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{
      attribution:'© <a href="https://openstreetmap.org">OpenStreetMap</a>',maxZoom:19
    }).addTo(_mapInstance);
  } else {
    _mapInstance.eachLayer(l=>{if(l.options&&(l.options.radius!==undefined||l._icon)) _mapInstance.removeLayer(l);});
  }
  setTimeout(()=>_mapInstance.invalidateSize(),100);

  const bounds=[];

  // מיקום המשתמש
  if(navigator.geolocation){
    try{
      const pos=await new Promise((res,rej)=>navigator.geolocation.getCurrentPosition(res,rej,{timeout:8000}));
      const ul=[pos.coords.latitude,pos.coords.longitude];
      bounds.push(ul);
      L.circleMarker(ul,{radius:13,fillColor:'#3dba6f',color:'#fff',weight:3,fillOpacity:0.95,zIndexOffset:1000})
        .addTo(_mapInstance).bindPopup('<div dir="rtl"><b>📍 המיקום שלי</b></div>');
      _mapInstance.setView(ul,11);
    }catch(e){}
  }

  if(!_allCards||!_allCards.length){
    status.textContent='טען פקעות תחילה (בנק פקעות)';return;
  }

  const cache=_geoCache();
  let done=0,placed=0;
  const total=_allCards.length;

  for(const c of _allCards){
    const addr=((c.address||'')+' ישראל').trim();
    if(!addr||addr==='ישראל'){done++;continue;}
    let lat,lng;
    if(cache[addr]){lat=cache[addr][0];lng=cache[addr][1];}
    else{
      try{
        const r=await fetch('https://nominatim.openstreetmap.org/search?q='+encodeURIComponent(addr)+'&format=json&limit=1&countrycodes=il');
        const data=await r.json();
        if(data.length){lat=parseFloat(data[0].lat);lng=parseFloat(data[0].lon);cache[addr]=[lat,lng];_saveGeoCache(cache);}
      }catch(e){}
      await new Promise(r=>setTimeout(r,1100));
    }
    if(lat&&lng){
      bounds.push([lat,lng]);placed++;
      L.circleMarker([lat,lng],{radius:10,fillColor:'#53bdeb',color:'#0d1b2a',weight:2,fillOpacity:0.9})
        .addTo(_mapInstance)
        .bindPopup(`<div dir="rtl" style="min-width:170px;font-family:Arial,sans-serif">
          <b style="font-size:13px">📋 קריאה ${c.call_id||''}</b><br>
          👤 ${c.name||''}<br>
          📍 ${c.address||''}<br>
          🔧 ${c.task_type||''}<br>
          📅 ${c.date||''} ${c.time?'⏰ '+c.time:''}
        </div>`);
    }
    done++;status.textContent=`ממפה... ${done}/${total}`;
  }
  if(bounds.length>1) _mapInstance.fitBounds(bounds,{padding:[40,40],maxZoom:14});
  status.textContent=`✅ ${placed} פקעות על המפה`;
}

/* ── הרשמה ── */
async function doRegister(){
  const name=document.getElementById('reg-name').value.trim();
  const errEl=document.getElementById('reg-err');
  const btn=document.getElementById('reg-btn');
  errEl.textContent='';
  if(!name){errEl.textContent='הכנס שם מלא';return;}
  btn.disabled=true;btn.textContent='שומר...';
  try{
    const r=await apiFetch('/api/register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,employee_id:localStorage.getItem('cc_emp')||''})});
    const d=await r.json();
    if(d.status==='ok'){
      const st=await apiFetch('/api/status');
      const sd=await st.json();
      _isAdmin=sd.is_admin||false; _userName=sd.user_name||'';
      if(sd.permissions) _userPerms=Object.assign({},_userPerms,sd.permissions);
      if(_isAdmin) showAdminTab();
      applyPermissions();
      showScreen('chat');
    } else {
      errEl.textContent=d.message||'שגיאה';
    }
  }catch(e){errEl.textContent='שגיאת רשת';}
  btn.disabled=false;btn.textContent='🚀 התחל ניסיון חינם';
}
document.getElementById('reg-name').addEventListener('keypress',e=>{if(e.key==='Enter')doRegister()});


/* ── אדמין ── */
const PERM_LABELS={
  premium:'⭐ פרמיום (גישה לכל הלשוניות)',
  full_card:'כרטיס פקעה מלא',
  inventory:'לשונית מחסן',
  equipment_report:'דוח ציוד (תת-לשונית)',
  reports:'לשונית דוחות',
  prices:'לשונית מחירון'
};

/* ── permission months modal ── */
let _permModalResolve=null;
function showPermMonthsModal(label){
  return new Promise(resolve=>{
    _permModalResolve=resolve;
    document.getElementById('perm-modal-label').textContent='הרשאה: '+label;
    document.getElementById('perm-modal').style.display='flex';
  });
}
function closePermModal(confirmed){
  document.getElementById('perm-modal').style.display='none';
  if(_permModalResolve){
    const months=confirmed?parseInt(document.getElementById('perm-months-sel').value)||1:null;
    _permModalResolve(months);
    _permModalResolve=null;
  }
}

function toggleUserCard(hdr){
  const details=hdr.nextElementSibling;
  const arrow=hdr.querySelector('.usr-arr');
  const isOpen=details.style.display!=='none';
  details.style.display=isOpen?'none':'block';
  if(arrow) arrow.style.transform=isOpen?'':'rotate(180deg)';
}

function permExpiry(months){
  const d=new Date(); d.setDate(d.getDate()+months*30);
  return d.toISOString().slice(0,10);
}

function daysLeft(isoDate){
  try{
    const exp=new Date(isoDate+'T00:00:00');
    const now=new Date(); now.setHours(0,0,0,0);
    return Math.round((exp-now)/(1000*60*60*24));
  }catch(e){return 0;}
}

function permValueLabel(v){
  if(v===false||v===null||v===undefined) return '';
  if(v===true) return '<span style="font-size:10px;color:#3dba6f;font-weight:600">✓ ללא הגבלה</span>';
  try{
    const days=daysLeft(v);
    if(days<0) return `<span style="font-size:10px;color:#f15c6e">פג תוקף (${v})</span>`;
    const color=days<=7?'#f9c846':days<=30?'#53bdeb':'#3dba6f';
    return `<span style="font-size:10px;color:${color};font-weight:600">נותרו ${days} ימים</span>`;
  }catch(e){return '';}
}

async function revokePerm(phone, permKey, cbId, spanId){
  if(!confirm('לבטל את ההרשאה "' + (PERM_LABELS[permKey]||permKey) + '"?')) return;
  const cb=document.getElementById(cbId);
  if(cb) cb.disabled=true;
  await saveSinglePerm(phone, permKey, false);
  if(cb){ cb.checked=false; cb.disabled=false; }
  const sp=document.getElementById(spanId);
  if(sp) sp.innerHTML='';
  const rb=document.getElementById('rb-'+cbId);
  if(rb) rb.innerHTML='';
}

async function togglePerm(phone, permKey, checkbox, labelSpanId){
  if(!checkbox.checked){
    // כיבוי — שמור מיד ללא חלון
    checkbox.disabled=true;
    await saveSinglePerm(phone, permKey, false);
    document.getElementById(labelSpanId).innerHTML='';
    checkbox.disabled=false;
    return;
  }
  // הפעלה — שאל כמה חודשים
  checkbox.checked=false; // החזר זמנית עד אישור
  const label=PERM_LABELS[permKey]||permKey;
  const months=await showPermMonthsModal(label);
  if(months===null){
    checkbox.checked=false; // ביטול
    return;
  }
  const value=permExpiry(months);
  checkbox.disabled=true;
  await saveSinglePerm(phone, permKey, value);
  checkbox.checked=true;
  document.getElementById(labelSpanId).innerHTML=permValueLabel(value);
  checkbox.disabled=false;
}

async function saveSinglePerm(phone, permKey, value){
  // שלוף הרשאות נוכחיות מה-DOM
  const card=document.querySelector(`[data-usrphone="${phone}"]`);
  const perms={};
  if(card){
    card.querySelectorAll('input[data-perm]').forEach(cb=>{
      perms[cb.dataset.perm]=cb._permValue!==undefined?cb._permValue:cb.checked;
    });
  }
  perms[permKey]=value;
  // שמור _permValue על ה-checkbox
  if(card){
    const cb=card.querySelector(`input[data-perm="${permKey}"]`);
    if(cb) cb._permValue=value;
  }
  try{
    await apiFetch('/api/admin/permissions/'+encodeURIComponent(phone),{
      method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify(perms)
    });
  }catch(e){}
}

async function loadAdminPanel(){
  const panel=document.getElementById('admin-panel');
  panel.dataset.loaded='1';
  panel.innerHTML='<div style="color:#8fa5b5;text-align:center;padding:24px">טוען...</div>';
  try{
    const r=await apiFetch('/api/admin/users');
    if(r.status===403){panel.innerHTML='<div style="color:#f07080;text-align:center;padding:24px">❌ אין הרשאת מנהל</div>';return;}
    const d=await r.json();
    const users=d.users||[];
    const permKeys=d.permission_keys||Object.keys(PERM_LABELS);
    let html=`<div style="font-weight:700;color:#3dba6f;font-size:16px;margin-bottom:14px">🛡️ לוח מנהל — ${users.length} משתמשים</div>`;
    if(!users.length){
      html+='<div style="color:#8fa5b5;text-align:center;padding:20px">אין משתמשים רשומים עדיין</div>';
    }
    users.forEach((u,i)=>{
      const maskedPhone=u.phone?u.phone.slice(0,3)+'****'+u.phone.slice(-2):'-';
      const isAdminUser=!!u.is_admin;
      const perms=u.permissions||{};
      let permsHtml='';
      if(!isAdminUser){
        permsHtml+=`<div style="margin-top:10px;border-top:1px solid #284461;padding-top:10px">
          <div style="font-size:12px;color:#8fa5b5;margin-bottom:8px;font-weight:600">🔐 הרשאות:</div>
          <div style="display:flex;flex-direction:column;gap:9px">`;
        permKeys.forEach(k=>{
          const v=perms.hasOwnProperty(k)?perms[k]:false;
          const isOn=v===true||(typeof v==='string'&&v>=(new Date().toISOString().slice(0,10)));
          const expLbl=permValueLabel(v);
          const spanId=`pv-${u.phone.replace(/\D/g,'')}-${k}`;
          const cbId=`cb-${u.phone.replace(/\D/g,'')}-${k}`;
          const label=PERM_LABELS[k]||k;
          const revokeBtn=isOn
            ?`<button onclick="revokePerm('${u.phone}','${k}','${cbId}','${spanId}')"
                title="בטל הרשאה"
                style="background:none;border:1px solid #f15c6e;color:#f15c6e;border-radius:6px;padding:2px 7px;font-size:12px;cursor:pointer;line-height:1.4">✕ בטל</button>`
            :'';
          permsHtml+=`<div style="display:flex;align-items:center;gap:8px">
            <input id="${cbId}" type="checkbox" data-perm="${k}" ${isOn?'checked':''}
              onchange="togglePerm('${u.phone}','${k}',this,'${spanId}')"
              style="width:17px;height:17px;accent-color:#3dba6f;cursor:pointer">
            <span style="font-size:13px;color:#dde8ee;flex:1">${label}</span>
            <span id="${spanId}" style="font-size:11px">${expLbl}</span>
            <span id="rb-${cbId}">${revokeBtn}</span>
          </div>`;
        });
        permsHtml+=`</div></div>`;
      } else {
        permsHtml+=`<div style="margin-top:8px;font-size:12px;color:#3dba6f">✅ מנהל — גישה מלאה</div>`;
      }
      // תוקף מנוי: ההרשאה הפעילה הקרובה לפקוע
      let subSummary='';
      if(!isAdminUser){
        const activeDates=Object.values(perms).filter(v=>typeof v==='string'&&v);
        if(activeDates.length){
          const minDays=Math.min(...activeDates.map(daysLeft));
          const col=minDays<0?'#f15c6e':minDays<=7?'#f9c846':minDays<=30?'#53bdeb':'#3dba6f';
          subSummary=`<div style="margin-top:4px;font-size:12px;color:${col}">⏱ תוקף מנוי: ${minDays<0?'פג תוקף':`נותרו ${minDays} ימים`}</div>`;
        } else if(Object.values(perms).some(v=>v===true)){
          subSummary=`<div style="margin-top:4px;font-size:12px;color:#3dba6f">⏱ מנוי ללא הגבלה</div>`;
        } else {
          subSummary=`<div style="margin-top:4px;font-size:12px;color:#8fa5b5">⏱ אין מנוי פעיל</div>`;
        }
      }
      html+=`<div data-usrphone="${u.phone}" style="background:#1a2d42;border-radius:10px;margin-bottom:10px;overflow:hidden">
        <div onclick="toggleUserCard(this)" style="padding:12px 14px;cursor:pointer;display:flex;justify-content:space-between;align-items:center;gap:8px">
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:15px;color:#edf2f5">${u.name||'ללא שם'}${isAdminUser?' <span style="font-size:11px;color:#3dba6f">🛡</span>':''}</div>
            <div style="color:#3dba6f;font-size:12px;margin-top:2px">נרשם: ${u.registered_at||'—'}</div>
            ${subSummary}
          </div>
          <span class="usr-arr" style="color:#6b94b8;font-size:20px;display:inline-block;transition:transform .25s;flex-shrink:0">▾</span>
        </div>
        <div style="display:none;padding:0 14px 12px;border-top:1px solid #284461">
          <div style="color:#9eabb5;font-size:13px;margin-top:8px">📱 ${maskedPhone}</div>
          ${permsHtml}
        </div>
      </div>`;
    });
    panel.innerHTML=html;
  }catch(e){
    panel.innerHTML='<div style="color:#f07080;text-align:center;padding:24px">❌ שגיאה בטעינה</div>';
  }
}

/* ════════════════════════════════════════
   צ'אט פנימי
   ════════════════════════════════════════ */
let _chatOpen=false;
let _chatWith='';       // phone of conversation partner (for admin)
let _chatPollTimer=null;

function initChat(){
  // הצג כפתור צ'אט לכל מי שמחובר
  const btn=document.getElementById('chat-float-btn');
  if(btn) btn.style.display='flex';
  startChatPoll();
}

function startChatPoll(){
  if(_chatPollTimer) clearInterval(_chatPollTimer);
  _chatPollTimer=setInterval(async()=>{
    if(_chatOpen && _chatWith){
      await refreshChatMessages();
    } else {
      await refreshUnreadBadge();
    }
  }, 8000);
}

async function refreshUnreadBadge(){
  try{
    const d=await apiFetch('/api/chat/unread').then(r=>r.json());
    const badge=document.getElementById('chat-unread-badge');
    if(!badge) return;
    const n=d.count||0;
    badge.textContent=n;
    badge.style.display=n>0?'flex':'none';
  }catch(e){}
}

function openChatModal(){
  _chatOpen=true;
  document.getElementById('chat-modal').style.display='flex';
  if(_isAdmin){
    showChatConvList();
  } else {
    document.getElementById('chat-input-bar').style.display='flex';
    updateChatHeader('💬 צ\'אט עם מנהל', false);
    openChatConversation('');
  }
}

function closeChatModal(){
  _chatOpen=false;
  _chatWith='';
  document.getElementById('chat-modal').style.display='none';
  refreshUnreadBadge();
}

async function showChatConvList(){
  _chatWith='';
  document.getElementById('chat-input-bar').style.display='none';
  const body=document.getElementById('chat-modal-body');
  body.innerHTML='<div style="color:#8fa5b5;text-align:center;padding:24px">טוען...</div>';
  updateChatHeader('💬 שיחות', false);
  try{
    const d=await apiFetch('/api/chat/conversations').then(r=>r.json());
    const convs=d.conversations||[];
    if(!convs.length){
      body.innerHTML='<div style="color:#8fa5b5;text-align:center;padding:40px;font-size:14px">אין שיחות עדיין</div>';
      return;
    }
    let html='';
    convs.forEach(c=>{
      const unreadBadge=c.unread>0?`<span style="background:#f15c6e;color:#fff;border-radius:50%;min-width:20px;height:20px;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;margin-right:4px">${c.unread}</span>`:'';
      const premBadge=c.premium?`<span style="font-size:11px;color:#f9c846;margin-right:4px">⭐ פרמיום${c.premium_expires?' עד '+c.premium_expires:''}</span>`:`<span style="font-size:11px;color:#8fa5b5">ללא פרמיום</span>`;
      html+=`<div onclick="openChatConversation('${c.phone}')" style="background:#d0e4f7;border-radius:10px;padding:12px 14px;margin-bottom:8px;cursor:pointer;display:flex;justify-content:space-between;align-items:center">
        <div>
          <div style="font-weight:700;color:#edf2f5;font-size:14px">${c.name} ${unreadBadge}</div>
          <div style="font-size:12px;color:#8fa5b5;margin-top:2px">${c.last_msg?c.last_msg.slice(0,40)+(c.last_msg.length>40?'...':''):''}</div>
          <div style="margin-top:4px">${premBadge}</div>
        </div>
        <span style="color:#8fa5b5;font-size:18px">›</span>
      </div>`;
    });
    body.innerHTML=html;
  }catch(e){
    body.innerHTML='<div style="color:#f07080;text-align:center;padding:24px">שגיאה בטעינה</div>';
  }
}

async function openChatConversation(withPhone){
  _chatWith=withPhone||'';
  document.getElementById('chat-input-bar').style.display='flex';
  const body=document.getElementById('chat-modal-body');
  body.innerHTML='<div style="color:#8fa5b5;text-align:center;padding:24px">טוען...</div>';
  const title=_isAdmin?(await getUserName(withPhone)):'צ\'אט עם מנהל';
  updateChatHeader(title, true, withPhone);
  await refreshChatMessages();
}

async function getUserName(phone){
  // try from conv list cache or return masked
  return phone?phone.slice(0,3)+'****'+phone.slice(-2):'—';
}

function updateChatHeader(title, showBack, withPhone=''){
  const hdr=document.getElementById('chat-modal-header');
  let grantBtn='';
  if(_isAdmin && withPhone){
    grantBtn=`<button onclick="showGrantModal('${withPhone}')" style="background:#3dba6f;color:#fff;border:none;border-radius:8px;padding:6px 12px;font-size:12px;font-weight:700;cursor:pointer;margin-right:8px">⭐ הענק פרמיום</button>`;
  }
  hdr.innerHTML=`
    <div style="display:flex;align-items:center;gap:10px">
      ${showBack&&_isAdmin?`<button onclick="showChatConvList()" style="background:none;border:none;color:#3dba6f;font-size:22px;cursor:pointer;padding:0">‹</button>`:''}
      <span style="font-weight:700;font-size:16px;color:#edf2f5">${title}</span>
    </div>
    <div style="display:flex;align-items:center">
      ${grantBtn}
      <button onclick="closeChatModal()" style="background:none;border:none;color:#8fa5b5;font-size:22px;cursor:pointer;padding:0 4px">✕</button>
    </div>`;
}

async function refreshChatMessages(){
  const body=document.getElementById('chat-modal-body');
  if(!body) return;
  const toPhone=_chatWith||(_isAdmin?'':ADMIN_PHONE_JS);
  if(!toPhone && _isAdmin) return;
  try{
    const d=await apiFetch('/api/chat/messages?with='+encodeURIComponent(toPhone)).then(r=>r.json());
    const msgs=d.messages||[];
    const wasAtBottom=body.scrollHeight-body.clientHeight<=body.scrollTop+40;
    const html=msgs.length?msgs.map(m=>{
      const mine=m.from===(_isAdmin?ADMIN_PHONE_JS:localStorage.getItem('cc_phone'));
      return `<div style="display:flex;justify-content:${mine?'flex-start':'flex-end'};margin-bottom:8px">
        <div style="max-width:80%;background:${mine?'#0d4a32':'#d0e4f7'};border-radius:${mine?'4px 14px 14px 14px':'14px 4px 14px 14px'};padding:8px 12px;font-size:14px;color:#ddeeff;line-height:1.5">
          ${m.body.replace(/\n/g,'<br>')}
          <div style="font-size:10px;color:#8fa5b5;margin-top:3px;text-align:${mine?'left':'right'}">${m.at}</div>
        </div>
      </div>`;
    }).join('')
    :'<div style="color:#8fa5b5;text-align:center;padding:30px;font-size:14px">אין הודעות עדיין — שלח הודעה למנהל</div>';
    body.innerHTML=html;
    if(wasAtBottom||!msgs.length) body.scrollTop=body.scrollHeight;
    await refreshUnreadBadge();
  }catch(e){}
}

async function sendChatMessage(){
  const inp=document.getElementById('chat-inp');
  const body=inp.value.trim();
  if(!body) return;
  inp.value='';
  const toPhone=_chatWith||(_isAdmin?'':ADMIN_PHONE_JS);
  if(!toPhone && _isAdmin) return;
  try{
    await apiFetch('/api/chat/send',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({body,to:toPhone})});
    await refreshChatMessages();
  }catch(e){inp.value=body;}
}

let _grantTarget='';
function showGrantModal(phone){
  _grantTarget=phone;
  document.getElementById('grant-modal').style.display='flex';
}
function closeGrantModal(){
  document.getElementById('grant-modal').style.display='none';
}
async function doGrantPremium(){
  const months=parseInt(document.getElementById('grant-months').value)||1;
  if(!_grantTarget) return;
  const btn=document.getElementById('grant-btn');
  btn.disabled=true; btn.textContent='מעניק...';
  try{
    const d=await apiFetch('/api/admin/grant-premium',{method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({phone:_grantTarget,months})}).then(r=>r.json());
    if(d.status==='ok'){
      closeGrantModal();
      await refreshChatMessages();
      // רענן שיחות
      if(!_chatWith) showChatConvList();
    }
  }catch(e){}
  btn.disabled=false; btn.textContent='הענק';
}

async function go(){
  const i=document.getElementById('inp');
  const t=i.value.trim();if(!t)return;
  i.value='';
  if(t.startsWith('פקעות')){
    add(t,'user');
    const parts=t.split(' ');
    loadTasks(parts[1]||'');
    return;
  }
  add(t,'user');
  const ty=document.getElementById('typing');
  document.getElementById('msgs').appendChild(ty);
  ty.style.display='block';
  document.getElementById('st').textContent='מעבד...';
  document.getElementById('msgs').scrollTop=9999;
  try{
    const r=await apiFetch('/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({message:t})});
    const d=await r.json();
    ty.style.display='none';
    add(d.text||d.response||'','bot');
  }catch(e){ty.style.display='none';add('❌ שגיאה','bot');}
  document.getElementById('st').textContent='מחובר';
}
/* ADMIN_PHONE_JS placeholder — filled server-side */
</script>
<script>const ADMIN_PHONE_JS='__ADMIN_PHONE__';</script>

<!-- כפתור צ'אט צף -->
<button id="chat-float-btn" onclick="openChatModal()" style="display:none;position:fixed;bottom:76px;left:16px;width:50px;height:50px;border-radius:50%;background:#25d366;border:none;color:#fff;font-size:22px;cursor:pointer;z-index:900;box-shadow:0 3px 12px rgba(0,0,0,.4);align-items:center;justify-content:center">
  💬
  <span id="chat-unread-badge" style="display:none;position:absolute;top:-3px;right:-3px;background:#f15c6e;color:#fff;border-radius:50%;min-width:18px;height:18px;font-size:10px;font-weight:700;align-items:center;justify-content:center;pointer-events:none"></span>
</button>

<!-- מודל צ'אט -->
<div id="chat-modal" style="display:none;position:fixed;inset:0;z-index:3000;background:#0d1b2a;flex-direction:column;direction:rtl">
  <div id="chat-modal-header" style="display:flex;justify-content:space-between;align-items:center;padding:12px 14px;background:#d0e4f7;border-bottom:1px solid #284461;flex-shrink:0"></div>
  <div id="chat-modal-body" style="flex:1;overflow-y:auto;padding:12px 14px"></div>
  <!-- input bar — נסתר כשמוצגת רשימת שיחות -->
  <div id="chat-input-bar" style="display:flex;gap:8px;padding:10px 12px;background:#d0e4f7;border-top:1px solid #284461;flex-shrink:0">
    <input id="chat-inp" type="text" placeholder="כתוב הודעה..." onkeypress="if(event.key==='Enter')sendChatMessage()"
      style="flex:1;background:#223d58;border:1px solid #284461;border-radius:10px;padding:10px 12px;color:#ddeeff;font-size:15px;outline:none;direction:rtl">
    <button onclick="sendChatMessage()" style="background:#25d366;border:none;border-radius:10px;padding:10px 16px;color:#fff;font-size:20px;cursor:pointer">➤</button>
  </div>
</div>

<!-- מודל חודשים להרשאה -->
<div id="perm-modal" style="display:none;position:fixed;inset:0;z-index:5000;background:rgba(0,0,0,.75);align-items:center;justify-content:center;direction:rtl">
  <div style="background:#d0e4f7;border-radius:16px;padding:22px 20px;width:90%;max-width:310px;display:flex;flex-direction:column;gap:14px">
    <div style="font-weight:700;font-size:15px;color:#ddeeff">⏳ משך ההרשאה</div>
    <div id="perm-modal-label" style="font-size:13px;color:#6b94b8"></div>
    <select id="perm-months-sel" style="background:#223d58;border:1px solid #284461;border-radius:10px;padding:11px 12px;color:#ddeeff;font-size:15px;direction:rtl">
      <option value="1">חודש אחד</option>
      <option value="2">2 חודשים</option>
      <option value="3">3 חודשים</option>
      <option value="4">4 חודשים</option>
      <option value="5">5 חודשים</option>
      <option value="6">6 חודשים</option>
    </select>
    <div style="display:flex;gap:10px">
      <button onclick="closePermModal(true)" style="flex:1;background:#3dba6f;color:#fff;border:none;border-radius:10px;padding:12px;font-size:15px;font-weight:700;cursor:pointer">אשר</button>
      <button onclick="closePermModal(false)" style="flex:1;background:#284461;color:#ddeeff;border:none;border-radius:10px;padding:12px;font-size:15px;cursor:pointer">ביטול</button>
    </div>
  </div>
</div>

<!-- מודל הענקת פרמיום -->
<div id="grant-modal" style="display:none;position:fixed;inset:0;z-index:4000;background:rgba(0,0,0,.7);align-items:center;justify-content:center;direction:rtl">
  <div style="background:#d0e4f7;border-radius:16px;padding:24px 20px;width:90%;max-width:320px;display:flex;flex-direction:column;gap:14px">
    <div style="font-weight:700;font-size:16px;color:#ddeeff">⭐ הענקת גישת פרמיום</div>
    <div style="font-size:14px;color:#6b94b8">בחר לכמה חודשים להעניק גישה:</div>
    <select id="grant-months" style="background:#223d58;border:1px solid #284461;border-radius:10px;padding:11px 12px;color:#ddeeff;font-size:15px;direction:rtl">
      <option value="1">חודש אחד</option>
      <option value="2">2 חודשים</option>
      <option value="3">3 חודשים</option>
      <option value="4">4 חודשים</option>
      <option value="5">5 חודשים</option>
      <option value="6">6 חודשים</option>
    </select>
    <div style="display:flex;gap:10px">
      <button id="grant-btn" onclick="doGrantPremium()" style="flex:1;background:#3dba6f;color:#fff;border:none;border-radius:10px;padding:12px;font-size:15px;font-weight:700;cursor:pointer">הענק</button>
      <button onclick="closeGrantModal()" style="flex:1;background:#284461;color:#ddeeff;border:none;border-radius:10px;padding:12px;font-size:15px;cursor:pointer">ביטול</button>
    </div>
  </div>
</div>
<!-- ══ MAP MODAL ══ -->
<div id="map-modal" style="display:none;position:fixed;inset:0;z-index:9999;background:#0d1b2a;flex-direction:column">
  <div style="display:flex;justify-content:space-between;align-items:center;padding:12px 16px;background:#0d1b2a;border-bottom:1px solid #1c3450;flex-shrink:0">
    <div style="font-weight:700;color:#eaf4ff;font-size:16px">🗺️ מפת ביקורים</div>
    <div style="display:flex;gap:10px;align-items:center">
      <div id="map-status" style="font-size:12px;color:#6b94b8;direction:rtl"></div>
      <button onclick="closeMap()" style="background:transparent;border:1px solid #284461;color:#6b94b8;border-radius:8px;padding:6px 14px;font-size:14px;cursor:pointer">✕ סגור</button>
    </div>
  </div>
  <div style="display:flex;gap:8px;padding:8px 12px;background:#0d1b2a;flex-shrink:0;flex-wrap:wrap">
    <div style="display:flex;align-items:center;gap:4px;font-size:12px;color:#ddeeff"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#3dba6f;border:2px solid #fff"></span>המיקום שלי</div>
    <div style="display:flex;align-items:center;gap:4px;font-size:12px;color:#ddeeff"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#53bdeb;border:2px solid #fff"></span>ביקור רגיל</div>
    <div style="display:flex;align-items:center;gap:4px;font-size:12px;color:#ddeeff"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#f9c846;border:2px solid #fff"></span>VIP</div>
    <div style="display:flex;align-items:center;gap:4px;font-size:12px;color:#ddeeff"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#f15c6e;border:2px solid #fff"></span>לא הושלם</div>
  </div>
  <div id="map-container" style="flex:1;width:100%"></div>
</div>
</body>
</html>"""

# אתחול DB בהפעלה — יוצר את כל הטבלאות כולל users
try:
    init_db().close()
except Exception as _e:
    print(f"DB init warning: {_e}")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    print(f"רץ על http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
