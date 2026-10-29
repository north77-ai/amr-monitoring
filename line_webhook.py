from flask import Flask, request, jsonify
import requests, pandas as pd, os, hashlib, json, subprocess, threading
from datetime import datetime
from zoneinfo import ZoneInfo

app = Flask(__name__)

# ================= CONFIG =================
LINE_TOKEN = "bnwZCoODv2oH4tOKKY9qngoA0Owf3piPTA/dOyZRX621w0FNeg3uPZaW6KjmcXj7Z2cVJarnnaINl24gdFy9hNVR6H3l9P09da/aUWG1IjNGUkHy2F2PhCVJJrfI+WBM5J76vDjbv8c6wpV4/voFgwdB04t89/1O/w1cDnyilFU="
ALERT_USER_ID = "Ca45282f34bacd5d00cb83b10d99ad08c"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "received_files", "data.xlsx")
LAST_ALERT_FILE = os.path.join(BASE_DIR, "last_alert.txt")

AMR_PROLAY_FILE = os.path.join(BASE_DIR, "AMR_prolay.py")
PYTHON_EXE = "python3"

ALLOWED_FILE = "ตามงาน SLA AMR"


# ================= TIME =================
def get_check_time():
    return datetime.now(
        ZoneInfo("Asia/Bangkok")
    ).strftime("%d/%m/%Y %H:%M:%S")


def get_filename(filename=None):
    return filename or os.path.basename(EXCEL_FILE)


# ================= LINE =================
def send_line(text):
    try:
        r = requests.post(
            "https://api.line.me/v2/bot/message/push",
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {LINE_TOKEN}"
            },
            json={
                "to": ALERT_USER_ID,
                "messages": [{"type": "text", "text": text}]
            },
            timeout=30
        )
        return r.status_code == 200
    except Exception:
        return False


def send_long(text):
    if len(text) <= 4500:
        return send_line(text)

    sep = "\n\n━━━━━━━━━━━━━━━━━━━━\n\n"
    blocks, messages, current = text.split(sep), [], ""

    for block in blocks:
        if current and len(current) + len(sep) + len(block) > 4500:
            messages.append(current)
            current = block
        else:
            current += sep if current else ""
            current += block
    if current:
        messages.append(current)

    return all(send_line(x) for x in messages)


# ================= EXCEL =================
def col(df, *names):
    for n in names:
        for c in df.columns:
            if str(c).strip().lower() == n.lower(): return c
    for n in names:
        for c in df.columns:
            if n.lower() in str(c).strip().lower(): return c
    return None

def read_excel():
    if not os.path.exists(EXCEL_FILE): return None
    try:
        x = pd.ExcelFile(EXCEL_FILE)
        sheet = "1.Data" if "1.Data" in x.sheet_names else None
        if not sheet: return None
        return pd.read_excel(EXCEL_FILE, sheet_name=sheet)
    except Exception:
        return None

# ================= GET JOBS =================
def get_jobs():
    df = read_excel()
    if df is None: return []

    problem = col(df, "METER_PROBLEM", "Problem", "PROBLEM")
    jobid = col(df, "JOBID", "JOB ID", "JobID")
    customer = col(df, "PEA No.", "PEA No", "ผู้ใช้ไฟฟ้า/สถานที่")
    group = col(df, "GROUPNAME", "GROUP NAME")
    date = col(df, "วันที่แจ้ง", "DATE")
    
    # ดึงคอลัมน์ "วันที่เหลือก่อนถึงวันสุดท้าย" มาเป็น priority แรก
    duration = col(df, "วันที่เหลือก่อนถึงวันสุดท้าย", "จำนวนวัน", "ระยะเวลางาน", "DURATION")
    
    site = col(df, "USED_NAME")
    status = col(df, "เช็คสถานะ", "สถานะ")
    area = col(df, "Teleport")
    sla_summary_col = col(df, "1.สรุปติดตาม SLA1")

    if not all([jobid, site, area, sla_summary_col]):
        print("Column สำคัญไม่ครบ")
        return []

    w = df.copy()
    w["_AREA_"] = w[area].fillna("").astype(str).str.strip().str.upper()
    w["_SLA_SUMMARY_"] = w[sla_summary_col].fillna("").astype(str).str.strip()

    w = w[(w["_AREA_"] == "BKK") & (w["_SLA_SUMMARY_"] == "1.1.ครบกำหนดนัดหมายวันนี้(เลื่อนได้)")]

    if status:
        closed_keywords = ["ปิด", "close", "completed", "เสร็จสิ้น"]
        for kw in closed_keywords:
            w = w[~w[status].astype(str).str.lower().str.contains(kw, na=False)]

    jobs = []
    for _, r in w.iterrows():
        def val(c):
            if not c: return "-"
            x = r.get(c, "")
            if pd.isna(x): return "-"
            x = str(x).strip()
            return x if x else "-"

        customer_text = val(customer).replace("สถานที่", "").strip()
        date_text = val(date)
        try:
            if date and not pd.isna(r[date]):
                d = pd.to_datetime(r[date], errors="coerce", dayfirst=True)
                if not pd.isna(d): date_text = d.strftime("%d/%m/%Y")
        except: pass

        try:
            if duration:
                d = float(r.get(duration, 0))
                duration_text = str(int(d)) if d.is_integer() else str(d)
            else: duration_text = "-"
        except: duration_text = "-"

        jobs.append({
            "site": val(site),
            "technician": "-",
            "customer": customer_text,
            "jobid": val(jobid),
            "group": val(group),
            "area": "BKK",
            "date": date_text,
            "problem": val(problem),
            "status": val(status),
            "duration": duration_text
        })
    return jobs

# ================= MESSAGE =================
def build_message(jobs, filename=None, check_time=None):
    filename = get_filename(filename)
    check_time = check_time or get_check_time()

    text = (
        "🚨 แจ้งเตือนงาน AMR (ตามวันนี้)\n"
        f"📁 ชื่อไฟล์: {filename}\n"
        f"🕒 เวลาตรวจสอบ: {check_time}\n"
        "พื้นที่: BKK\n"
        f"พบงานที่ต้องตรวจสอบ {len(jobs)} งาน"
    )

    for j in jobs:
        text += (
            "\n\n━━━━━━━━━━━━━━━━━━━━\n\n"
            f"🏢 Site / สถานที่: {j['site']}\n"
            f"👨‍🔧 ช่างผู้รับผิดชอบ: {j['technician']}\n"
            f"👤 ผู้ใช้ไฟฟ้า: {j['customer']}\n"
            f"🆔 JOBID: {j['jobid']}\n"
            f"🏷️ กลุ่ม: {j['group']}\n"
            f"📍 พื้นที่: {j['area']}\n"
            f"📅 วันที่แจ้ง: {j['date']}\n"
            f"⚠️ ปัญหาที่พบ: {j['problem']}\n"
            f"📊 สถานะ: {j['status']}\n"
            f"⏱️ ระยะเวลางาน: {j['duration']} วัน"
        )
    return text

def send_no_jobs_confirmation(filename=None, check_time=None):
    filename = get_filename(filename)
    check_time = check_time or get_check_time()
    message = (
        "✅ ตรวจสอบงาน AMR เรียบร้อย\n"
        f"📁 ชื่อไฟล์: {filename}\n"
        f"🕒 เวลาตรวจสอบ: {check_time}\n"
        "📍 พื้นที่: BKK\n"
        "พบงานที่ต้องตรวจสอบ 0 งาน"
    )
    return send_line(message)

# ================= IMPORT =================
def run_import():
    try:
        p = subprocess.Popen([PYTHON_EXE, AMR_PROLAY_FILE], cwd=os.path.dirname(AMR_PROLAY_FILE))
    except Exception: pass

def process_excel(filename=None):
    filename = get_filename(filename)
    check_time = get_check_time()
    jobs = get_jobs()
    if jobs: send_long(build_message(jobs, filename, check_time))
    else: send_no_jobs_confirmation(filename, check_time)
    threading.Thread(target=run_import, daemon=True).start()

# ================= WEBHOOK =================
@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json(silent=True) or {}
    for event in data.get("events", []):
        msg = event.get("message", {})
        if msg.get("type") != "file": continue

        filename = msg.get("fileName", "")
        message_id = msg.get("id")

        if not (filename.startswith(ALLOWED_FILE) and filename.lower().endswith((".xlsx", ".xls"))):
            continue

        try:
            r = requests.get(
                f"https://api-data.line.me/v2/bot/message/{message_id}/content",
                headers={"Authorization": f"Bearer {LINE_TOKEN}"},
                timeout=60
            )
            if r.status_code != 200: continue

            os.makedirs(os.path.dirname(EXCEL_FILE), exist_ok=True)
            if os.path.exists(EXCEL_FILE):
                try: os.remove(EXCEL_FILE)
                except: pass

            with open(EXCEL_FILE, "wb") as f: f.write(r.content)
            process_excel(filename)
        except Exception: pass
    return jsonify(status="ok"), 200

@app.route("/")
def index():
    return jsonify({"status": "running", "excel": EXCEL_FILE, "allowed_file": ALLOWED_FILE})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)