import os
import pandas as pd

from influxdb_client import InfluxDBClient, Point, WritePrecision
from influxdb_client.client.write_api import SYNCHRONOUS


# =========================
# CONFIG (Cloud & Environment Ready)
# =========================

EXCEL = os.environ.get("EXCEL_FILE", r"C:\project\received_files\data.xlsx")
INFLUX_URL = os.environ.get("INFLUX_URL", "https://us-east-1-1.aws.cloud2.influxdata.com")
INFLUX_TOKEN = os.environ.get("INFLUX_TOKEN", "N2ZY7UCs1hz2yDg5TCvEf9oap4thA9fmIjGbA6aEogkpkEnlP4rCVKqm6WRSWfDvWqOT681vh0MCF89DYRD4yQ==")
INFLUX_ORG = os.environ.get("INFLUX_ORG", "AMR")
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
# INFLUXDB CONNECTION (CLOUD)
# =========================

try:
    client = InfluxDBClient(
        url=INFLUX_URL,
        token=INFLUX_TOKEN,
        org=INFLUX_ORG,
        timeout=300000
    )
    write = client.write_api(write_options=SYNCHRONOUS)
except Exception as e:
    print("❌ เชื่อมต่อ InfluxDB ไม่สำเร็จ:", e)
    raise SystemExit


# =========================
# WRITE SUMMARY POINTS FOR GRAFANA
# =========================

try:
    p_summary = (
        Point("amr_summary_bkk")
        .field("total_workorder", int(len(df[df["Teleport"] == "BKK"])))
        .field("alert_count", int(len(alert_df)))
        .field("overdue_count", int(len(overdue_df)))
    )
    
    write.write(
        bucket=INFLUX_BUCKET,
        org=INFLUX_ORG,
        record=p_summary
    )
    print("✅ บันทึกข้อมูลสรุป (amr_summary_bkk) บน Cloud สำเร็จ!")

except Exception as e:
    print("❌ บันทึกข้อมูลสรุปไม่สำเร็จ:", e)

write.close()
client.close()

print("==============================================")
print("IMPORT COMPLETE & READY FOR GRAFANA")
print("==============================================")