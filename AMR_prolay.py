import os
import pandas as pd

from influxdb_client import InfluxDBClient, Point, WritePrecision
from influxdb_client.client.write_api import SYNCHRONOUS


# =========================
# CONFIG
# =========================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(BASE_DIR, "received_files", "data.xlsx")

# ⚠️ หมายเหตุ: หากเก็บบน Render แนะนำให้ใช้ InfluxDB Cloud URL และ Token จริง
INFLUX_URL = "https://your-influxdb-cloud-url"
INFLUX_TOKEN = "C35rZpoQAlyjXsgVT9vLbDQMW2GlQ8LPK956yzfWbqwWpfil98oNk2OWfRvQ-Y6CuV_9Z4yW7ZiW1g-5DfLrKQ=="
INFLUX_ORG = "AMR"
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


def point_base(measurement, row, index):
    jid = jobid(row["JOBID"]) or f"ROW_{index}"

    # ทำให้งานที่วันที่เหมือนกันไม่ทับกัน
    t = row["วันที่แจ้ง"] + pd.to_timedelta(index, unit="ns")

    return (
        Point(measurement)
        .tag("JOBID", jid)
        .field("GROUPNAME", clean(row["GROUPNAME"]))
        .field("USED_NAME", clean(row["USED_NAME"]))
        .field("Teleport", clean(row["Teleport"]).upper())
        .field("รายละเอียด", clean(row["รายละเอียด"]))
        .field("METER_PROBLEM", clean(row["METER_PROBLEM"]))
        .field("จำนวนวัน", float(row["จำนวนวัน"]))
        .time(t, WritePrecision.NS)
    )


# =========================
# CHECK FILE
# =========================

if not os.path.exists(EXCEL):
    print("❌ ไม่พบไฟล์:", EXCEL)
    raise SystemExit


# =========================
# READ EXCEL
# =========================

try:
    excel = pd.ExcelFile(EXCEL)

    if "1.Data" not in excel.sheet_names:
        print("❌ ไม่พบ Sheet 1.Data")
        raise SystemExit

    df = pd.read_excel(
        EXCEL,
        sheet_name="1.Data"
    )

except Exception as e:
    print("❌ อ่าน Excel ไม่สำเร็จ:", e)
    raise SystemExit


df.columns = [str(c).strip() for c in df.columns]

required = [
    "JOBID",
    "GROUPNAME",
    "USED_NAME",
    "Teleport",
    "รายละเอียด",
    "วันที่แจ้ง",
    "METER_PROBLEM",
    "จำนวนวัน"
]

missing = [c for c in required if c not in df.columns]

if missing:
    print("❌ Excel ขาด Column:", missing)
    raise SystemExit


# =========================
# CLEAN DATA
# =========================

for c in [
    "JOBID",
    "GROUPNAME",
    "USED_NAME",
    "Teleport",
    "รายละเอียด",
    "METER_PROBLEM"
]:
    df[c] = df[c].apply(clean)

df["JOBID"] = df["JOBID"].apply(jobid)

df["Teleport"] = (
    df["Teleport"]
    .str.upper()
    .str.strip()
)

df["วันที่แจ้ง"] = pd.to_datetime(
    df["วันที่แจ้ง"],
    dayfirst=True,
    errors="coerce"
)

df = df.dropna(
    subset=["วันที่แจ้ง"]
).copy()

df["จำนวนวัน"] = pd.to_numeric(
    df["จำนวนวัน"],
    errors="coerce"
).fillna(0)


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
# INFLUXDB
# =========================

try:
    client = InfluxDBClient(
        url=INFLUX_URL,
        token=INFLUX_TOKEN,
        org=INFLUX_ORG,
        timeout=300000
    )

    write = client.write_api(
        write_options=SYNCHRONOUS
    )

except Exception as e:
    print("❌ เชื่อมต่อ InfluxDB ไม่สำเร็จ:", e)
    raise SystemExit


# =========================
# DELETE OLD DATA
# =========================

for measurement in [
    "amr_workorder",
    "amr_alert",
    "amr_overdue_bkk"
]:
    try:
        client.delete_api().delete(
            start="1970-01-01T00:00:00Z",
            stop="2100-01-01T00:00:00Z",
            predicate=f'_measurement="{measurement}"',
            bucket=INFLUX_BUCKET,
            org=INFLUX_ORG
        )
        print("ลบ", measurement, "สำเร็จ")

    except Exception as e:
        print("⚠️ ลบ", measurement, "ไม่สำเร็จ:", e)


# =========================
# WORKORDER
# =========================

work_ok = 0
work_err = 0

for i, row in df.iterrows():
    try:
        write.write(
            bucket=INFLUX_BUCKET,
            org=INFLUX_ORG,
            record=point_base(
                "amr_workorder",
                row,
                i
            )
        )
        work_ok += 1

    except Exception as e:
        work_err += 1
        print("workorder error:", e)


# =========================
# ALERT
# =========================

alert_ok = 0
alert_err = 0

for i, row in alert_df.iterrows():
    try:
        p = point_base(
            "amr_alert",
            row,
            i
        ).field(
            "alert",
            1
        )

        write.write(
            bucket=INFLUX_BUCKET,
            org=INFLUX_ORG,
            record=p
        )

        alert_ok += 1

    except Exception as e:
        alert_err += 1
        print("alert error:", e)


# =========================
# OVERDUE
# =========================

overdue_ok = 0
overdue_err = 0

try:
    p = (
        Point("amr_overdue_bkk")
        .field("count", int(len(overdue_df)))
    )

    write.write(
        bucket=INFLUX_BUCKET,
        org=INFLUX_ORG,
        record=p
    )

    overdue_ok = 1

except Exception as e:
    overdue_err = 1
    print("overdue error:", e)


# =========================
# CLOSE
# =========================

write.close()
client.close()


# =========================
# SUMMARY
# =========================

print()
print("==============================================")
print("IMPORT COMPLETE")
print("==============================================")

print("จำนวนข้อมูลทั้งหมด       :", len(df))
print("amr_workorder            :", work_ok)
print("amr_alert                :", alert_ok)
print("amr_overdue_bkk          :", len(overdue_df))

print("งาน BKK + Problem 0-7 วัน:", len(alert_df))
print("งาน BKK 7 วันขึ้นไป      :", len(overdue_df))

print("Error workorder          :", work_err)
print("Error alert              :", alert_err)
print("Error overdue            :", overdue_err)

print("==============================================")
print("ส่งข้อมูลเข้า InfluxDB สำเร็จ!")
print("==============================================")