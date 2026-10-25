import os
import pandas as pd
from datetime import datetime, timezone

from influxdb_client import InfluxDBClient, Point, WritePrecision
from influxdb_client.client.write_api import SYNCHRONOUS


# =========================
# CONFIG (InfluxDB Cloud)
# =========================

EXCEL = r"C:\project\received_files\data.xlsx"

INFLUX_URL = os.environ.get("INFLUX_URL", "https://us-east-1-1.aws.cloud2.influxdata.com")
INFLUX_TOKEN = os.environ.get("INFLUX_TOKEN", "N2ZY7UCs1hz2yDg5TCvEf9oap4thA9fmIjGbA6aEogkpkEnlP4rCVKqm6WRSWfDvWqOT681vh0MCF89DYRD4yQ==")
INFLUX_ORG = os.environ.get("INFLUX_ORG", "AMR-Monitoring")
INFLUX_BUCKET = os.environ.get("INFLUX_BUCKET", "AMR_DATA")


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


# =========================
# CHECK FILE & READ EXCEL
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

required = [
    "JOBID", "GROUPNAME", "USED_NAME", "Teleport",
    "รายละเอียด", "วันที่แจ้ง", "METER_PROBLEM", "จำนวนวัน"
]

missing = [c for c in required if c not in df.columns]
if missing:
    print("❌ Excel ขาด Column:", missing)
    raise SystemExit


# =========================
# CLEAN DATA
# =========================

for c in ["JOBID", "GROUPNAME", "USED_NAME", "Teleport", "รายละเอียด", "METER_PROBLEM"]:
    df[c] = df[c].apply(clean)

df["JOBID"] = df["JOBID"].apply(jobid)
df["Teleport"] = df["Teleport"].str.upper().str.strip()
df["วันที่แจ้ง"] = pd.to_datetime(df["วันที่แจ้ง"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["วันที่แจ้ง"]).copy()
df["จำนวนวัน"] = pd.to_numeric(df["จำนวนวัน"], errors="coerce").fillna(0)


# =========================
# FILTER
# =========================

alert_df = df[
    (df["Teleport"] == "BKK") &
    (df["METER_PROBLEM"] != "") &
    (df["จำนวนวัน"].between(0, 7))
].copy()

overdue_df = df[
    (df["Teleport"] == "BKK") &
    (df["จำนวนวัน"] >= 7)
].copy()

print()
print("จำนวนข้อมูลทั้งหมด:", len(df))
print("BKK + Problem 0-7:", len(alert_df))
print("BKK >= 7 วัน:", len(overdue_df))


# =========================
# INFLUXDB CLOUD (BATCH UPLOAD)
# =========================

try:
    client = InfluxDBClient(
        url=INFLUX_URL,
        token=INFLUX_TOKEN,
        org=INFLUX_ORG,
        timeout=300000
    )
    write_api = client.write_api(write_options=SYNCHRONOUS)
except Exception as x:
    print("❌ เชื่อมต่อ InfluxDB Cloud ไม่สำเร็จ:", x)
    raise SystemExit

now_time = datetime.now(timezone.utc)
points = []

# 1. เตรียมข้อมูล amr_workorder ทั้งหมด
for i, row in df.iterrows():
    jid = jobid(row["JOBID"]) or f"ROW_{i}"
    p = (
        Point("amr_workorder")
        .tag("JOBID", jid)
        .field("GROUPNAME", clean(row["GROUPNAME"]))
        .field("USED_NAME", clean(row["USED_NAME"]))
        .field("Teleport", clean(row["Teleport"]).upper())
        .field("รายละเอียด", clean(row["รายละเอียด"]))
        .field("METER_PROBLEM", clean(row["METER_PROBLEM"]))
        .field("จำนวนวัน", float(row["จำนวนวัน"]))
        .time(now_time, WritePrecision.NS)
    )
    points.append(p)

# 2. เตรียมข้อมูล amr_alert
for i, row in alert_df.iterrows():
    jid = jobid(row["JOBID"]) or f"ROW_{i}"
    p = (
        Point("amr_alert")
        .tag("JOBID", jid)
        .field("GROUPNAME", clean(row["GROUPNAME"]))
        .field("USED_NAME", clean(row["USED_NAME"]))
        .field("Teleport", clean(row["Teleport"]).upper())
        .field("รายละเอียด", clean(row["รายละเอียด"]))
        .field("METER_PROBLEM", clean(row["METER_PROBLEM"]))
        .field("จำนวนวัน", float(row["จำนวนวัน"]))
        .field("alert", 1)
        .time(now_time, WritePrecision.NS)
    )
    points.append(p)

# 3. เตรียมข้อมูล amr_overdue_bkk
p_overdue = (
    Point("amr_overdue_bkk")
    .field("count", int(len(overdue_df)))
    .time(now_time, WritePrecision.NS)
)
points.append(p_overdue)

# ส่งข้อมูลทั้งหมดทีเดียวรวดเดียวจบ (Batch Write)
try:
    print(" đang กำลังส่งข้อมูลทั้งหมดขึ้น InfluxDB Cloud...")
    write_api.write(bucket=INFLUX_BUCKET, org=INFLUX_ORG, record=points)
    print("✅ ส่งข้อมูลสำเร็จเรียบร้อยแล้ว!")
except Exception as e:
    print("❌ เกิดข้อผิดพลาดในการส่งข้อมูล:", e)

write_api.close()
client.close()

print()
print("==============================================")
print("IMPORT COMPLETE (CLOUD)")
print("==============================================")
print("จำนวนข้อมูลทั้งหมด       :", len(df))
print("งาน BKK + Problem 0-7 วัน:", len(alert_df))
print("งาน BKK 7 วันขึ้นไป      :", len(overdue_df))
print("==============================================")