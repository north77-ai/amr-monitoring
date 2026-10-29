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

def clean(v):
    if pd.isna(v): return ""
    return str(v).strip()

def jobid(v):
    v = clean(v)
    return v[:-2] if v.endswith(".0") and v[:-2].isdigit() else v

def col(df, *names):
    for n in names:
        for c in df.columns:
            if str(c).strip().lower() == n.lower(): return c
    for n in names:
        for c in df.columns:
            if n.lower() in str(c).strip().lower(): return c
    return None

# =========================
# READ EXCEL
# =========================
if not os.path.exists(EXCEL):
    raise SystemExit("❌ ไม่พบไฟล์:", EXCEL)

try:
    excel = pd.ExcelFile(EXCEL)
    if "1.Data" not in excel.sheet_names:
        raise SystemExit("❌ ไม่พบ Sheet 1.Data")
    df = pd.read_excel(EXCEL, sheet_name="1.Data")
except Exception as e:
    raise SystemExit(f"❌ อ่าน Excel ไม่สำเร็จ: {e}")

df.columns = [str(c).strip() for c in df.columns]

problem_col = col(df, "METER_PROBLEM", "Problem", "PROBLEM")
jobid_col = col(df, "JOBID", "JOB ID", "JobID")
group_col = col(df, "GROUPNAME", "GROUP NAME")
date_col = col(df, "วันที่แจ้ง", "DATE")
duration_col = col(df, "วันที่เหลือก่อนถึงวันสุดท้าย", "จำนวนวัน", "ระยะเวลางาน", "DURATION")
used_col = col(df, "USED_NAME")
status_col = col(df, "เช็คสถานะ", "สถานะ")
teleport_col = col(df, "Teleport")
detail_col = col(df, "รายละเอียด")
sla_summary_col = col(df, "1.สรุปติดตาม SLA1")

required = [jobid_col, group_col, used_col, teleport_col, detail_col, date_col, problem_col, duration_col, sla_summary_col]
if any(c is None for c in required):
    raise SystemExit("❌ Excel ขาด Column สำคัญ")

# =========================
# CLEAN DATA ล่วงหน้า
# =========================
df["_JOBID_"] = df[jobid_col].apply(jobid)
df["_GROUPNAME_"] = df[group_col].apply(clean)
df["_USED_NAME_"] = df[used_col].apply(clean)
df["_TELEPORT_"] = df[teleport_col].fillna("ไม่ระบุ").astype(str).str.upper().str.strip()
df["_DETAIL_"] = df[detail_col].apply(clean)
df["_PROBLEM_"] = df[problem_col].fillna("").astype(str).str.strip()
df["_DURATION_"] = pd.to_numeric(df[duration_col], errors="coerce").fillna(0)
df["_SLA_SUMMARY_"] = df[sla_summary_col].fillna("").astype(str).str.strip()
df["_DATE_"] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
df["_STATUS_"] = df[status_col].fillna("").astype(str).str.lower() if status_col else ""

def point_base(measurement, row, index):
    jid = row["_JOBID_"] or f"ROW_{index}"
    actual_date = row["_DATE_"]
    
    if pd.isna(actual_date):
        t = pd.Timestamp.now(tz='UTC').normalize() + pd.to_timedelta(index, unit="ns")
    else:
        t = actual_date
        if t.tzinfo is None:
            t = t.tz_localize('UTC')
        t = t + pd.to_timedelta(index, unit="ns")

    return (
        Point(measurement)
        .tag("JOBID", jid)
        .field("GROUPNAME", row["_GROUPNAME_"])
        .field("USED_NAME", row["_USED_NAME_"])
        .field("Teleport", row["_TELEPORT_"])
        .field("รายละเอียด", row["_DETAIL_"])
        .field("METER_PROBLEM", row["_PROBLEM_"])
        .field("จำนวนวัน", float(row["_DURATION_"]))
        .time(t, WritePrecision.NS)
    )

# =========================
# FILTER (ปรับเงื่อนไขให้ตรงกับหน้าสรุป Excel 2 งาน)
# =========================
closed_keywords = ["ปิด", "close", "completed", "เสร็จสิ้น"]
is_closed = df["_STATUS_"].str.contains('|'.join(closed_keywords), regex=True)

alert_df = df[
    (df["_TELEPORT_"] == "BKK") &
    (df["_SLA_SUMMARY_"] == "1.1.ครบกำหนดนัดหมายวันนี้(เลื่อนได้)") &
    (~is_closed)
].copy()

overdue_df = df[
    (df["_TELEPORT_"] == "BKK") &
    (df["_DURATION_"] >= 7)
].copy()

# =========================
# INFLUXDB CONNECTION & BATCH WRITE
# =========================
try:
    client = InfluxDBClient(url=INFLUX_URL, token=INFLUX_TOKEN, org=INFLUX_ORG, timeout=300000)
    write = client.write_api(write_options=SYNCHRONOUS)
except Exception as e:
    raise SystemExit(f"❌ เชื่อมต่อ InfluxDB ไม่สำเร็จ: {e}")

def write_in_chunks(points, measurement_name, chunk_size=500):
    success_count, error_count = 0, 0
    for i in range(0, len(points), chunk_size):
        chunk = points[i:i + chunk_size]
        retries = 3
        while retries > 0:
            try:
                write.write(bucket=INFLUX_BUCKET, org=INFLUX_ORG, record=chunk)
                success_count += len(chunk)
                time.sleep(0.5) 
                break
            except Exception as e:
                if "429" in str(e):
                    time.sleep(15)
                    retries -= 1
                else:
                    error_count += len(chunk)
                    break
    return success_count, error_count

# 1. WRITE WORKORDER
points_workorder = []
for i, row in df.iterrows():
    try: points_workorder.append(point_base("amr_workorder", row, i))
    except Exception as e: pass
work_ok, work_err = write_in_chunks(points_workorder, "amr_workorder", chunk_size=300)

# 2. WRITE ALERT
points_alert = []
bkk_df = df[df["_TELEPORT_"] == "BKK"]
for i, row in bkk_df.iterrows():
    try:
        is_alert_val = 1 if i in alert_df.index else 0
        p = point_base("amr_alert", row, i).field("is_alert", is_alert_val)
        points_alert.append(p)
    except Exception as e: pass
alert_ok, alert_err = write_in_chunks(points_alert, "amr_alert", chunk_size=300)

# 3. WRITE OVERDUE
overdue_ok = 0
now_ts = pd.Timestamp.now(tz='UTC').normalize()
try:
    p = Point("amr_overdue_bkk").field("count", int(len(overdue_df))).time(now_ts, WritePrecision.NS)
    write.write(bucket=INFLUX_BUCKET, org=INFLUX_ORG, record=p)
    overdue_ok = 1
except Exception as limit_err:
    if "429" in str(limit_err):
        time.sleep(15)
        write.write(bucket=INFLUX_BUCKET, org=INFLUX_ORG, record=p)
        overdue_ok = 1

# 4. WRITE TOTAL JOBS
total_ok = 0
try:
    p_total = Point("amr_total_jobs").field("count", int(len(df))).time(now_ts, WritePrecision.NS)
    write.write(bucket=INFLUX_BUCKET, org=INFLUX_ORG, record=p_total)
    total_ok = 1
except Exception as limit_err:
    if "429" in str(limit_err):
        time.sleep(15)
        write.write(bucket=INFLUX_BUCKET, org=INFLUX_ORG, record=p_total)
        total_ok = 1

# 5. WRITE AREA SUMMARY 
area_counts = df["_TELEPORT_"].value_counts()
for area_name, count_val in area_counts.items():
    if area_name:
        p_area = Point("amr_area_summary").tag("Teleport", str(area_name)).field("count", int(count_val)).time(now_ts, WritePrecision.NS)
        try:
            write.write(bucket=INFLUX_BUCKET, org=INFLUX_ORG, record=p_area)
        except Exception as e:
            if "429" in str(e):
                time.sleep(15)
                write.write(bucket=INFLUX_BUCKET, org=INFLUX_ORG, record=p_area)

# 6. WRITE AGE SUMMARY
age_groups = {
    "0-1 วัน": len(df[df["_DURATION_"] <= 1]),
    "2-3 วัน": len(df[(df["_DURATION_"] > 1) & (df["_DURATION_"] <= 3)]),
    "4-6 วัน": len(df[(df["_DURATION_"] > 3) & (df["_DURATION_"] <= 6)]),
    "7+ วัน": len(df[df["_DURATION_"] > 6])
}
for age_name, count_val in age_groups.items():
    p_age = Point("amr_age_summary").tag("age_group", age_name).field("count", int(count_val)).time(now_ts, WritePrecision.NS)
    try:
        write.write(bucket=INFLUX_BUCKET, org=INFLUX_ORG, record=p_age)
    except Exception as e:
        if "429" in str(e):
            time.sleep(15)
            write.write(bucket=INFLUX_BUCKET, org=INFLUX_ORG, record=p_age)

write.close()
client.close()
print(f"✅ ส่งข้อมูลเข้า InfluxDB สำเร็จ!")