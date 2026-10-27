import os
import time
import pandas as pd

from influxdb_client import InfluxDBClient, Point, WritePrecision
from influxdb_client.client.write_api import SYNCHRONOUS

# =========================
# CONFIG
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(BASE_DIR, "received_files", "data.xlsx")

INFLUX_URL = "https://us-east-1-1.aws.cloud2.influxdata.com"
INFLUX_TOKEN = "AHRpHuPjeAOov_2Bz2LukBhTScKaEJzbVcUqAdjIg6kNNfViDCWjUVxsvSho-6HA9OycXDjGEg4KfLq0b11JLQ=="
INFLUX_ORG = "AMR-Monitoring"
INFLUX_BUCKET = "AMR_DATA"

# =========================
# FUNCTIONS
# =========================
def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()

def jobid(v):
    v = clean(v)
    return v[:-2] if v.endswith(".0") and v[:-2].isdigit() else v

def col(df, *names):
    for n in names:
        for c in df.columns:
            if str(c).strip().lower() == n.lower():
                return c
    for n in names:
        for c in df.columns:
            if n.lower() in str(c).strip().lower():
                return c
    return None

FIXED_TIMESTAMP = pd.Timestamp.now(tz='UTC').normalize()

def point_base(measurement, row, index, jobid_col, group_col, used_col, teleport_col, detail_col, problem_col, duration_col, date_col):
    jid = jobid(row.get(jobid_col, "")) or f"ROW_{index}"
    actual_date_str = str(row.get(date_col, ""))

    return (
        Point(measurement)
        .tag("JOBID", jid)
        .field("GROUPNAME", clean(row.get(group_col, "")))
        .field("USED_NAME", clean(row.get(used_col, "")))
        .field("Teleport", clean(row.get(teleport_col, "")).upper())
        .field("รายละเอียด", clean(row.get(detail_col, "")))
        .field("METER_PROBLEM", clean(row.get(problem_col, "")))
        .field("จำนวนวัน", float(row.get(duration_col, 0)))
        .field("วันที่แจ้งจริง", actual_date_str)
        .time(FIXED_TIMESTAMP, WritePrecision.NS)
    )

# =========================
# READ EXCEL & FILTER
# =========================
if not os.path.exists(EXCEL):
    print("❌ ไม่พบไฟล์:", EXCEL)
    raise SystemExit

try:
    excel = pd.ExcelFile(EXCEL)
    if "1.Data" not in excel.sheet_names:
        print("❌ ไม่พบ Sheet 1.Data")
        raise SystemExit
    df = pd.read_excel(EXCEL, sheet_name="1.Data")
except Exception as e:
    print("❌ อ่าน Excel ไม่สำเร็จ:", e)
    raise SystemExit

df.columns = [str(c).strip() for c in df.columns]

problem_col = col(df, "METER_PROBLEM", "Problem", "PROBLEM")
jobid_col = col(df, "JOBID", "JOB ID", "JobID")
group_col = col(df, "GROUPNAME", "GROUP NAME")
date_col = col(df, "วันที่แจ้ง", "DATE")
duration_col = col(df, "จำนวนวัน", "ระยะเวลางาน", "DURATION")
used_col = col(df, "USED_NAME")
status_col = col(df, "เช็คสถานะ", "สถานะ")
teleport_col = col(df, "Teleport")
detail_col = col(df, "รายละเอียด")
today_col = col(df, "วันนี้") 

required = [jobid_col, group_col, used_col, teleport_col, detail_col, date_col, problem_col, duration_col, today_col]
if any(c is None for c in required):
    print("❌ Excel ขาด Column สำคัญ (ต้องมีคอลัมน์ 'วันนี้')")
    raise SystemExit

w = df.copy()
w["_AREA_"] = w[teleport_col].fillna("").astype(str).str.strip().str.upper()
w["_PROBLEM_"] = w[problem_col].fillna("").astype(str).str.strip()
w["_DURATION_"] = pd.to_numeric(w[duration_col], errors="coerce").fillna(0)
w["_TODAY_"] = w[today_col].fillna("").astype(str).str.strip()

closed_keywords = ["ปิด", "close", "completed", "เสร็จสิ้น"]
status_val = w[status_col].astype(str).str.lower() if status_col else pd.Series([""]*len(w))
is_closed = status_val.str.contains('|'.join(closed_keywords), na=False, regex=True)

alert_mask = (
    (w["_AREA_"] == "BKK") &
    (w["_TODAY_"] == "ตามวันนี้") &
    (~is_closed)
)
alert_df = w[alert_mask].copy()

overdue_df = w[
    (w["_AREA_"] == "BKK") &
    (w["_DURATION_"] >= 7)
].copy()

print("\n--- สรุปข้อมูล ---")
print("จำนวนข้อมูลทั้งหมดใน Excel:", len(df))
print("BKK + ตามวันนี้ (ตรงกับ LINE):", len(alert_df))
print("BKK >= 7 วัน:", len(overdue_df))

# =========================
# INFLUXDB CONNECTION
# =========================
try:
    client = InfluxDBClient(url=INFLUX_URL, token=INFLUX_TOKEN, org=INFLUX_ORG, timeout=300000)
    write = client.write_api(write_options=SYNCHRONOUS)
except Exception as e:
    print("❌ เชื่อมต่อ InfluxDB ไม่สำเร็จ:", e)
    raise SystemExit

# อัปเกรดระบบ Write ให้ส่งช้าลง และส่งซ้ำเมื่อโดนแบนความเร็ว (429)
def write_in_chunks(points, measurement_name, chunk_size=500):
    success_count = 0
    error_count = 0
    for i in range(0, len(points), chunk_size):
        chunk = points[i:i + chunk_size]
        retries = 3
        while retries > 0:
            try:
                write.write(bucket=INFLUX_BUCKET, org=INFLUX_ORG, record=chunk)
                success_count += len(chunk)
                time.sleep(0.5) # พักครึ่งวินาทีทุกการส่ง 1 ก้อน เพื่อกันการโดนจำกัดความเร็ว
                break
            except Exception as e:
                err_str = str(e)
                if "429" in err_str or "Too Many Requests" in err_str:
                    print(f"⚠️ โดนจำกัดความเร็ว! รอ 15 วินาทีแล้วส่งใหม่... (เหลือโควต้าส่งซ้ำ {retries-1} ครั้ง)")
                    time.sleep(15)
                    retries -= 1
                else:
                    error_count += len(chunk)
                    print(f"❌ {measurement_name} chunk error at index {i}:", err_str)
                    break
        
        if retries == 0:
            error_count += len(chunk)
            print(f"❌ ข้าม Chunk ที่ {i} เนื่องจากส่งซ้ำครบโควต้าแล้ว")
            
    return success_count, error_count

# =========================
# WRITE DATA
# =========================
points_workorder = []
for i, row in df.iterrows():
    try:
        points_workorder.append(point_base("amr_workorder", row, i, jobid_col, group_col, used_col, teleport_col, detail_col, problem_col, duration_col, date_col))
    except Exception as e:
        pass
print("\nกำลังส่ง amr_workorder...")
work_ok, work_err = write_in_chunks(points_workorder, "amr_workorder", chunk_size=400)

points_alert = []
bkk_df = w[w["_AREA_"] == "BKK"]
for i, row in bkk_df.iterrows():
    try:
        is_alert_val = 1 if i in alert_df.index else 0
        p = point_base("amr_alert", row, i, jobid_col, group_col, used_col, teleport_col, detail_col, problem_col, duration_col, date_col).field("is_alert", is_alert_val)
        points_alert.append(p)
    except Exception as e:
        pass
print("กำลังส่ง amr_alert...")
alert_ok, alert_err = write_in_chunks(points_alert, "amr_alert", chunk_size=400)

overdue_ok, overdue_err = 0, 0
try:
    p = Point("amr_overdue_bkk").field("count", int(len(overdue_df))).time(FIXED_TIMESTAMP, WritePrecision.NS)
    # เพิ่ม try เล็กๆ ไว้เผื่อติด Limit ตอนจบ
    try:
        write.write(bucket=INFLUX_BUCKET, org=INFLUX_ORG, record=p)
        overdue_ok = 1
    except Exception as limit_err:
        if "429" in str(limit_err):
            time.sleep(15)
            write.write(bucket=INFLUX_BUCKET, org=INFLUX_ORG, record=p)
            overdue_ok = 1
except Exception as e:
    overdue_err = 1

write.close()
client.close()

print("\n==============================================")
print("IMPORT COMPLETE")
print("==============================================")
print("จำนวนข้อมูลทั้งหมดใน Excel :", len(df))
print("amr_workorder            :", work_ok)
print("amr_alert (ซิงค์ LINE)   :", alert_ok)
print("amr_overdue_bkk          :", len(overdue_df) if overdue_ok else 0)
print("งาน BKK + ตามวันนี้      :", len(alert_df))
print("งาน BKK 7 วันขึ้นไป      :", len(overdue_df))
print("==============================================")
print("✅ ส่งข้อมูลเข้า InfluxDB สำเร็จ!")