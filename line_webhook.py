from flask import Flask, request, jsonify
import requests, pandas as pd, os, hashlib, json, subprocess, threading
from datetime import datetime
from zoneinfo import ZoneInfo

app = Flask(__name__)

# ================= CONFIG (Environment Variables / Cloud Ready) =================
# ดึงค่าจาก Environment Variables บน Render หรือใส่ค่าตรงๆ สำหรับรันในเครื่อง
LINE_TOKEN = os.environ.get("LINE_TOKEN", "bnwZCoODv2oH4tOKKY9qngoA0Owf3piPTA/dOyZRX621w0FNeg3uPZaW6KjmcXj7Z2cVJarnnaINl24gdFy9hNVR6H3l9P09da/aUWG1IjNGUkHy2F2PhCVJJrfI+WBM5J76vDjbv8c6wpV4/voFgwdB04t89/1O/w1cDnyilFU=")
ALERT_USER_ID = os.environ.get("ALERT_USER_ID", "Ca45282f34bacd5d00cb83b10d99ad08c")

# กำหนดตำแหน่งไฟล์ Excel และไฟล์ที่ใช้เก็บข้อมูล Alert ล่าสุด
# บน Cloud แนะนำให้ใช้โฟลเดอร์ /tmp สำหรับไฟล์ชั่วคราว
EXCEL_FILE = os.environ.get("EXCEL_FILE", r"C:\project\received_files\data.xlsx")
LAST_ALERT_FILE = os.environ.get("LAST_ALERT_FILE", r"C:\project\last_alert.txt")

# กำหนดไฟล์ Python ที่ใช้ส่งข้อมูลเข้า InfluxDB
AMR_PROLAY_FILE = os.environ.get("AMR_PROLAY_FILE", r"C:\project\AMR_prolay.py")
PYTHON_EXE = os.environ.get("PYTHON_EXE", "python")

# รับเฉพาะไฟล์ที่ชื่อขึ้นต้นด้วยคำนี้
ALLOWED_FILE = "ตามงาน SLA AMR"


# ================= TIME =================
def get_check_time():
    # อ่านเวลาปัจจุบันตามเขตเวลาไทย
    return datetime.now(
        ZoneInfo("Asia/Bangkok")
    ).strftime("%d/%m/%Y %H:%M:%S")


def get_filename(filename=None):
    # ใช้ชื่อไฟล์ที่รับจาก LINE
    # ถ้าไม่มีชื่อ ให้ใช้ชื่อไฟล์ในเครื่องแทน
    return filename or os.path.basename(EXCEL_FILE)


# ================= LINE =================
def send_line(text):
    # ส่งข้อความแจ้งเตือนไปยัง LINE
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

        print("LINE STATUS:", r.status_code)
        print("LINE RESPONSE:", r.text)

        return r.status_code == 200

    except Exception as e:
        print("LINE ERROR:", e)
        return False


def send_long(text):
    # แบ่งข้อความยาวออกเป็นหลายข้อความ
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
    # ค้นหาชื่อคอลัมน์แบบตรงกันก่อน
    for n in names:
        for c in df.columns:
            if str(c).strip().lower() == n.lower():
                return c

    # จากนั้นค้นหาแบบมีคำที่ระบุอยู่ในชื่อ
    for n in names:
        for c in df.columns:
            if n.lower() in str(c).strip().lower():
                return c

    return None


def read_excel():
    # ตรวจสอบว่ามีไฟล์ Excel อยู่หรือไม่
    if not os.path.exists(EXCEL_FILE):
        print("ไม่พบ:", EXCEL_FILE)
        return None

    try:
        x = pd.ExcelFile(EXCEL_FILE)
        sheet = "1.Data" if "1.Data" in x.sheet_names else None

        if not sheet:
            print("ไม่พบ Sheet 1.Data")
            return None

        return pd.read_excel(EXCEL_FILE, sheet_name=sheet)

    except Exception as e:
        print("อ่าน Excel ไม่สำเร็จ:", e)
        return None


# ================= GET JOBS =================
def get_jobs():
    df = read_excel()

    if df is None:
        return []

    problem = col(df, "METER_PROBLEM", "Problem", "PROBLEM")
    jobid = col(df, "JOBID", "JOB ID", "JobID")
    customer = col(df, "PEA No.", "PEA No", "ผู้ใช้ไฟฟ้า/สถานที่")
    group = col(df, "GROUPNAME", "GROUP NAME")
    date = col(df, "วันที่แจ้ง", "DATE")
    duration = col(df, "จำนวนวัน", "ระยะเวลางาน", "DURATION")
    site = col(df, "USED_NAME")
    status = col(df, "เช็คสถานะ", "สถานะ")
    area = col(df, "Teleport")

    print("\n========== COLUMN ==========")
    print("problem :", problem)
    print("jobid   :", jobid)
    print("customer:", customer)
    print("group   :", group)
    print("site    :", site)
    print("status  :", status)
    print("area    :", area)
    print("date    :", date)
    print("duration:", duration)
    print("============================\n")

    # ตรวจสอบว่าคอลัมน์สำคัญครบหรือไม่
    if not all([
        problem, jobid, customer, group, site,
        status, area, date, duration
    ]):
        print("Column สำคัญไม่ครบ")
        return []

    w = df.copy()

    w["_AREA_"] = (
        w[area].fillna("").astype(str).str.strip().str.upper()
    )

    w["_PROBLEM_"] = (
        w[problem].fillna("").astype(str).str.strip()
    )

    w["_DURATION_"] = pd.to_numeric(
        w[duration], errors="coerce"
    )

    # เงื่อนไขเดิม: BKK + มี Problem + ระยะเวลา 0-7 วัน
    w = w[
        (w["_AREA_"] == "BKK") &
        (w["_PROBLEM_"] != "") &
        (w["_DURATION_"].between(0, 7))
    ]

    print("BKK + Problem 0-7 วัน:", len(w))

    jobs = []

    for _, r in w.iterrows():

        def val(c):
            if not c:
                return "-"

            x = r.get(c, "")

            if pd.isna(x):
                return "-"

            x = str(x).strip()
            return x if x else "-"

        customer_text = (
            val(customer).replace("สถานที่", "").strip()
        )

        date_text = val(date)

        try:
            d = pd.to_datetime(
                r[date], errors="coerce", dayfirst=True
            )

            if not pd.isna(d):
                date_text = d.strftime("%d/%m/%Y")
        except:
            pass

        try:
            d = float(r["_DURATION_"])
            duration_text = (
                str(int(d)) if d.is_integer() else str(d)
            )
        except:
            duration_text = "-"

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
        "🚨 แจ้งเตือนงาน AMR\n"
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


# ================= NO JOBS =================
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

    ok = send_line(message)

    if ok:
        print("✅ ส่งข้อความยืนยัน 0 งานสำเร็จ")
    else:
        print("❌ ส่งข้อความยืนยัน 0 งานไม่สำเร็จ")

    return ok


# ================= DUPLICATE =================
def signature(jobs):
    return hashlib.md5(
        json.dumps(
            jobs,
            ensure_ascii=False,
            sort_keys=True
        ).encode()
    ).hexdigest()


def duplicate(sig):
    if not os.path.exists(LAST_ALERT_FILE):
        return False

    try:
        with open(LAST_ALERT_FILE, encoding="utf-8") as f:
            return f.read().strip() == sig
    except:
        return False


def save_sig(sig):
    with open(
        LAST_ALERT_FILE,
        "w",
        encoding="utf-8"
    ) as f:
        f.write(sig)


# ================= IMPORT =================
def run_import():
    try:
        p = subprocess.Popen(
            [PYTHON_EXE, AMR_PROLAY_FILE],
            cwd=os.path.dirname(AMR_PROLAY_FILE) if os.path.dirname(AMR_PROLAY_FILE) else None
        )
        print("เริ่ม AMR_prolay.py PID:", p.pid)

    except Exception as e:
        print("เรียก AMR_prolay.py ไม่สำเร็จ:", e)


def process_excel(filename=None):
    filename = get_filename(filename)
    check_time = get_check_time()

    jobs = get_jobs()

    if jobs:
        message = build_message(
            jobs,
            filename,
            check_time
        )

        if send_long(message):
            print("✅ ส่ง LINE Alert สำเร็จ")
        else:
            print("❌ ส่ง LINE Alert ไม่สำเร็จ")

    else:
        send_no_jobs_confirmation(
            filename,
            check_time
        )

    threading.Thread(
        target=run_import,
        daemon=True
    ).start()


# ================= WEBHOOK =================
@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json(silent=True) or {}

    for event in data.get("events", []):

        msg = event.get("message", {})

        if msg.get("type") != "file":
            continue

        filename = msg.get("fileName", "")
        message_id = msg.get("id")

        print("ได้รับไฟล์:", filename)

        if not (
            filename.startswith(ALLOWED_FILE)
            and filename.lower().endswith((".xlsx", ".xls"))
        ):
            print("⏭️ ข้ามไฟล์:", filename)
            continue

        try:
            r = requests.get(
                f"https://api-data.line.me/v2/bot/message/"
                f"{message_id}/content",
                headers={
                    "Authorization": f"Bearer {LINE_TOKEN}"
                },
                timeout=60
            )

            print("Download status:", r.status_code)

            if r.status_code != 200:
                continue

            os.makedirs(
                os.path.dirname(EXCEL_FILE) if os.path.dirname(EXCEL_FILE) else ".",
                exist_ok=True
            )

            with open(EXCEL_FILE, "wb") as f:
                f.write(r.content)

            print("บันทึก Excel:", EXCEL_FILE)

            process_excel(filename)

        except Exception as e:
            print("ERROR รับไฟล์:", e)

    return jsonify(status="ok"), 200


# ================= GRAFANA =================
@app.route("/grafana-alert", methods=["POST"])
def grafana_alert():
    data = request.get_json(silent=True) or {}
    status = str(data.get("status", "")).lower()

    print("Grafana status:", status)

    if status not in ("firing", "alerting"):
        return jsonify(status="ignored"), 200

    jobs = get_jobs()
    filename = os.path.basename(EXCEL_FILE)
    check_time = get_check_time()

    if not jobs:
        ok = send_no_jobs_confirmation(
            filename,
            check_time
        )

        return jsonify(
            status="no_jobs",
            confirmation_sent=ok
        ), 200

    sig = signature(jobs)

    if duplicate(sig):
        return jsonify(status="duplicate"), 200

    ok = send_long(
        build_message(
            jobs,
            filename,
            check_time
        )
    )

    if ok:
        save_sig(sig)

    return jsonify(
        status="sent" if ok else "failed",
        jobs=len(jobs)
    ), 200


# ================= TEST =================
@app.route("/test-send")
def test_send():
    filename = os.path.basename(EXCEL_FILE)
    check_time = get_check_time()
    jobs = get_jobs()

    if not jobs:
        ok = send_no_jobs_confirmation(
            filename,
            check_time
        )

        return jsonify(
            status="no_jobs",
            confirmation_sent=ok
        ), 200

    ok = send_long(
        build_message(
            jobs,
            filename,
            check_time
        )
    )

    return jsonify(
        status="sent" if ok else "failed",
        jobs=len(jobs)
    ), 200


@app.route("/")
def index():
    return jsonify({
        "status": "running",
        "excel": EXCEL_FILE,
        "alert_to": "ALERT_USER_ID",
        "auto_influx": AMR_PROLAY_FILE,
        "allowed_file": ALLOWED_FILE
    })


# ================= START =================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("================================")
    print("AMR MONITORING SYSTEM (CLOUD READY)")
    print("================================")
    print("Excel:", EXCEL_FILE)
    print("รับไฟล์ขึ้นต้นด้วย:", ALLOWED_FILE)
    print("Alert → ALERT_USER_ID")
    print("พื้นที่: BKK")
    print("Problem: METER_PROBLEM")
    print("Site → USED_NAME")
    print("สถานะ → เช็คสถานะ")
    print("จำนวนวัน: 0-7")
    print("ส่งซ้ำเมื่อรับไฟล์ใหม่: เปิดใช้งาน")
    print("================================")

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )