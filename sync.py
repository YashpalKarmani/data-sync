"""
FusionSolar -> Google Sheet sync (writes straight to Sheets, no Apps Script needed).

Usage:
    python sync.py live    # every 15 minutes: performance, devices, alarms, Latest_Plants
    python sync.py daily   # once a day (~23:55 PKT): plant/device lists, Daily_History, cleanup of old rows
    python sync.py all     # daily + live (use for the very first run)

Settings come from environment variables (use these on hosting):
    FS_BASE, FS_USER, FS_SYSCODE, SHEET_ID,
    GOOGLE_SA_JSON (full contents of the service-account json)  OR  GOOGLE_SA_FILE (path to the json file)
Optional: RETENTION_DAYS (default 7), ALARM_LOOKBACK_HOURS (default 24)
"""
import json
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import gspread
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.getenv("FS_BASE", "https://intl.fusionsolar.huawei.com/thirdData")
USER = os.getenv("FS_USER", "google_sheets_sync")
SYSCODE = os.getenv("FS_SYSCODE", "YAHAN_SYSTEM_CODE")
SHEET_ID = os.getenv("SHEET_ID", "YAHAN_GOOGLE_SHEET_ID")
SA_FILE = os.getenv("GOOGLE_SA_FILE", os.path.join(HERE, "service_account.json"))
RETENTION_DAYS = int(os.getenv("RETENTION_DAYS", "7"))
ALARM_LOOKBACK_H = int(os.getenv("ALARM_LOOKBACK_HOURS", "24"))

PKT = timezone(timedelta(hours=5))  # Pakistan, no DST
INVERTER_TYPES = {1, 38}           # device types whose active_power counts as generation
SKIP_TYPES = {62}                  # dongle: no real-time KPI from Huawei
DEFAULT_CITIES = ["Karachi", "Hyderabad", "Lahore", "Islamabad", "Rawalpindi", "Faisalabad",
                  "Multan", "Peshawar", "Quetta", "Sialkot", "Gujranwala", "Sukkur"]
HEALTH = {1: "Offline", 2: "Faulty", 3: "Healthy"}


# ------------------------------------------------------------------ helpers
def serial(dt):
    """Google Sheets date serial for a datetime (sheet time zone = Asia/Karachi)."""
    return round((dt.replace(tzinfo=None) - datetime(1899, 12, 30)).total_seconds() / 86400, 6)


def fmt_ms(ms):
    try:
        return datetime.fromtimestamp(int(ms) / 1000, PKT).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return ""


def norm(v):
    """Make a value safe to write with RAW input (long ids stay text)."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int) and abs(v) >= 10 ** 12:
        return str(v)
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    return v


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# ------------------------------------------------------------------ Google Sheet
class Book:
    def __init__(self):
        raw = os.getenv("GOOGLE_SA_JSON")
        if raw:
            gc = gspread.service_account_from_dict(json.loads(raw))
        else:
            gc = gspread.service_account(filename=SA_FILE)
        self.ss = gc.open_by_key(SHEET_ID)
        self.ws = {w.title: w for w in self.ss.worksheets()}
        self.new = set()

    def get(self, title, rows=1000, cols=26):
        if title not in self.ws:
            self.ws[title] = self.ss.add_worksheet(title=title, rows=rows, cols=cols)
            self.new.add(title)
        return self.ws[title]

    def _format_first_col(self, w, header):
        patterns = {"time": "yyyy-mm-dd hh:mm", "updated": "yyyy-mm-dd hh:mm", "date": "yyyy-mm-dd"}
        if header and header[0] in patterns:
            kind = "DATE" if header[0] == "date" else "DATE_TIME"
            w.format("A:A", {"numberFormat": {"type": kind, "pattern": patterns[header[0]]}})

    def append(self, title, rows, key=None):
        """Append rows. With key=<column>, rows whose key already exists are skipped."""
        if not rows:
            print(f"{title}: nothing to add")
            return
        w = self.get(title)
        header = [] if title in self.new else w.row_values(1)
        keys = list(dict.fromkeys(k for r in rows for k in r))
        added = [k for k in keys if k not in header]
        if added:
            header = header + added
            if w.col_count < len(header):
                w.add_cols(len(header) - w.col_count)
            w.update(range_name="A1", values=[header])
            if title in self.new:
                self._format_first_col(w, header)
        if key:
            existing = set(w.col_values(header.index(key) + 1)[1:]) if key in header else set()
            fresh = []
            for r in rows:
                k = str(r.get(key, ""))
                if k not in existing:
                    existing.add(k)
                    fresh.append(r)
            rows = fresh
            if not rows:
                print(f"{title}: no new rows")
                return
        data = [[norm(r.get(h)) for h in header] for r in rows]
        w.append_rows(data, value_input_option="RAW", insert_data_option="INSERT_ROWS", table_range="A1")
        print(f"{title}: +{len(data)} rows")

    def replace(self, title, rows):
        """Overwrite a whole tab with the given rows."""
        w = self.get(title)
        if not rows:
            w.clear()
            print(f"{title}: cleared (no rows)")
            return
        header = list(dict.fromkeys(k for r in rows for k in r))
        data = [header] + [[norm(r.get(h)) for h in header] for r in rows]
        w.clear()
        w.resize(rows=max(len(data), 2), cols=max(len(header), 1))
        w.update(range_name="A1", values=data)
        self._format_first_col(w, header)
        print(f"{title}: {len(rows)} rows")

    def prune(self, title, cutoff_serial):
        """Delete rows from the top whose first column (a date serial) is older than the cutoff."""
        w = self.ws.get(title)
        if not w:
            return
        vals = w.col_values(1, value_render_option="UNFORMATTED_VALUE")[1:]
        n = 0
        for v in vals:
            if isinstance(v, (int, float)) and v < cutoff_serial:
                n += 1
            else:
                break
        if n:
            w.delete_rows(2, n + 1)
        print(f"{title}: removed {n} old rows")


def ensure_settings(book):
    """Create the Settings tab if missing and return the list of cities."""
    if "Settings" not in book.ws:
        w = book.get("Settings", rows=20, cols=3)
        w.update(range_name="A1", values=[
            ["Setting", "Value", "Note"],
            ["Critical if performance below (%)", 30, "Red alert"],
            ["Low if performance below (%)", 50, "Orange / low alert"],
            ["Daytime starts (hour, 0-23)", 9, "Performance alerts only run between these hours"],
            ["Daytime ends (hour, 0-23)", 16, "Outside this window plants show NIGHT, not an alert"],
            ["Alarm window (hours)", 24, "Show alarms raised in the last N hours"],
            ["Cities (comma separated)", ", ".join(DEFAULT_CITIES), "City is detected from plant address / name"],
        ])
        w.format("A1:C1", {"textFormat": {"bold": True}})
    for row in book.ws["Settings"].get_all_values():
        if row and str(row[0]).strip().lower().startswith("cities") and len(row) > 1:
            cities = [c.strip() for c in row[1].split(",") if c.strip()]
            if cities:
                return cities
    return DEFAULT_CITIES


# ------------------------------------------------------------------ Huawei API
s = requests.Session()


def login():
    r = s.post(BASE + "/login", json={"userName": USER, "systemCode": SYSCODE}, timeout=60)
    j = r.json()
    if not j.get("success"):
        raise RuntimeError("Login failed: " + r.text)
    s.headers["XSRF-TOKEN"] = s.cookies.get("XSRF-TOKEN")


def call(path, body):
    last = None
    for _ in range(4):
        r = s.post(BASE + path, json=body, timeout=60)
        j = r.json()
        if j.get("success"):
            return j
        last = r.text
        code = j.get("failCode")
        if code == 305:
            login()
        elif code == 407:
            time.sleep(65)
        else:
            break
    raise RuntimeError(f"{path} failed: {last}")


def code_of(p):
    return p.get("stationCode") or p.get("plantCode")


def name_of(p):
    return p.get("plantName") or p.get("stationName") or code_of(p)


def batches(seq, n=100):
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


def fetch_plants():
    plants, page = [], 1
    while True:
        d = call("/getStationList", {"pageNo": page, "pageSize": 100})["data"]
        plants += d["list"]
        if page >= d.get("pageCount", 1):
            break
        page += 1
    return plants


def fetch_devices(codes):
    devs = []
    for b in batches(codes):
        devs += call("/getDevList", {"stationCodes": ",".join(b)})["data"]
    return devs


def fetch_kpi(codes):
    out = {}
    for b in batches(codes):
        for d in call("/getStationRealKpi", {"stationCodes": ",".join(b)})["data"]:
            out[d["stationCode"]] = d["dataItemMap"]
    return out


def fetch_alarms(codes, hours):
    end = int(time.time() * 1000)
    out = []
    for b in batches(codes):
        out += call("/getAlarmList", {"stationCodes": ",".join(b), "beginTime": end - hours * 3600000,
                                      "endTime": end, "language": "en"})["data"]
    return out


def city_of(p, cities):
    text = (str(p.get("plantAddress") or p.get("stationAddr") or p.get("address") or "") + " " + str(name_of(p))).lower()
    for c in cities:
        if c.lower() in text:
            return c
    return "Unknown"


def device_kpis(devs, meta, ts):
    """Real-time KPI per device type. Returns (rows_by_type, kw_by_plant, ok)."""
    by_type = {}
    for d in devs:
        by_type.setdefault(d["devTypeId"], []).append(d)
    rows_by_type, kw, ok = {}, {}, True
    for t, lst in by_type.items():
        if t in SKIP_TYPES:
            continue
        try:
            info = {str(d["id"]): d for d in lst}
            rows = []
            for b in batches(lst):
                res = call("/getDevRealKpi", {"devIds": ",".join(str(d["id"]) for d in b), "devTypeId": t})
                for x in res["data"]:
                    d = info.get(str(x["devId"]), {})
                    pc = d.get("stationCode", "")
                    rows.append({"time": ts, "devId": str(x["devId"]), "devName": d.get("devName", ""),
                                 "plantCode": pc, "plantName": meta.get(pc, {}).get("name", ""),
                                 **x["dataItemMap"]})
                    if t in INVERTER_TYPES:
                        kw[pc] = kw.get(pc, 0.0) + max(0.0, num(x["dataItemMap"].get("active_power")))
            rows_by_type[t] = rows
        except Exception as e:
            if "20013" in str(e):
                print(f"device type {t}: not supported by API, skipped")
            else:
                print(f"ERROR device type {t}: {e}")
                ok = False if t in INVERTER_TYPES else ok
    return rows_by_type, kw, ok


# ------------------------------------------------------------------ modes
def build_meta(plants, cities):
    return {code_of(p): {"name": str(name_of(p)), "city": city_of(p, cities),
                         "address": str(p.get("plantAddress") or p.get("stationAddr") or ""),
                         "cap": num(p.get("capacity"))} for p in plants}


def live(book, now, errors):
    cities = ensure_settings(book)
    ts = serial(now)
    plants = fetch_plants()
    codes = [code_of(p) for p in plants]
    meta = build_meta(plants, cities)
    kpi = fetch_kpi(codes)

    try:
        book.append("Plant_Performance", [
            {"time": ts, "plantCode": c, "plantName": meta[c]["name"], "city": meta[c]["city"], **kpi.get(c, {})}
            for c in codes])
    except Exception as e:
        errors.append(f"performance: {e}")
        print("ERROR performance:", e)

    kw, dev_ok, n_inv = {}, False, {}
    try:
        devs = fetch_devices(codes)
        for d in devs:
            if d["devTypeId"] in INVERTER_TYPES:
                n_inv[d.get("stationCode")] = n_inv.get(d.get("stationCode"), 0) + 1
        rows_by_type, kw, dev_ok = device_kpis(devs, meta, ts)
        for t, rows in rows_by_type.items():
            book.append(f"Dev_Type_{t}", rows)
    except Exception as e:
        errors.append(f"devices: {e}")
        print("ERROR devices:", e)

    try:
        latest = []
        for c in codes:
            d = kpi.get(c, {})
            st = int(num(d.get("real_health_state")))
            latest.append({
                "updated": ts, "plantCode": c, "plantName": meta[c]["name"], "city": meta[c]["city"],
                "address": meta[c]["address"], "capacity_kw": meta[c]["cap"],
                "health_state": st, "health": HEALTH.get(st, "No data"),
                "day_kwh": num(d.get("day_power")), "month_kwh": num(d.get("month_power")),
                "total_kwh": num(d.get("total_power")),
                "current_kw": round(kw.get(c, 0.0), 2) if dev_ok else "",
                "inverters": n_inv.get(c, 0),
            })
        book.replace("Latest_Plants", latest)
    except Exception as e:
        errors.append(f"latest: {e}")
        print("ERROR latest:", e)

    try:
        alarms = fetch_alarms(codes, ALARM_LOOKBACK_H)
        for a in alarms:
            a["_key"] = f'{a.get("stationCode")}|{a.get("alarmId")}|{a.get("raiseTime")}|{a.get("esnCode")}'
            a["raiseTimeReadable"] = fmt_ms(a.get("raiseTime"))
            a["plantName"] = a.get("stationName") or meta.get(a.get("stationCode"), {}).get("name", "")
        book.append("Plant_Alarms", alarms, key="_key")
    except Exception as e:
        errors.append(f"alarms: {e}")
        print("ERROR alarms:", e)


def daily(book, now, errors):
    try:
        book.ss.batch_update({"requests": [{"updateSpreadsheetProperties": {
            "properties": {"timeZone": "Asia/Karachi"}, "fields": "timeZone"}}]})
    except Exception as e:
        print("time zone not changed:", e)

    cities = ensure_settings(book)
    plants = fetch_plants()
    codes = [code_of(p) for p in plants]
    meta = build_meta(plants, cities)
    book.replace("Plant_Details", plants)

    try:
        book.replace("Devices", fetch_devices(codes))
    except Exception as e:
        errors.append(f"devices list: {e}")
        print("ERROR devices list:", e)

    try:
        kpi = fetch_kpi(codes)
        d0 = (now.replace(tzinfo=None) - datetime(1899, 12, 30)).days
        rows = []
        for c in codes:
            d = kpi.get(c, {})
            rows.append({"date": d0, "plantCode": c, "plantName": meta[c]["name"], "city": meta[c]["city"],
                         "capacity_kw": meta[c]["cap"], "day_kwh": num(d.get("day_power")),
                         "month_kwh": num(d.get("month_power")), "total_kwh": num(d.get("total_power")),
                         "_key": f"{d0}|{c}"})
        book.append("Daily_History", rows, key="_key")
    except Exception as e:
        errors.append(f"daily history: {e}")
        print("ERROR daily history:", e)

    cutoff = serial(now - timedelta(days=RETENTION_DAYS))
    for title in [t for t in book.ws if t == "Plant_Performance" or t.startswith("Dev_Type_")]:
        try:
            book.prune(title, cutoff)
        except Exception as e:
            errors.append(f"prune {title}: {e}")
            print(f"ERROR prune {title}:", e)


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "live"
    if mode not in ("live", "daily", "all"):
        raise SystemExit("usage: python sync.py [live|daily|all]")
    now = datetime.now(PKT)
    errors = []
    book = Book()
    login()
    if mode in ("daily", "all"):
        daily(book, now, errors)
    if mode in ("live", "all"):
        live(book, now, errors)
    if errors:
        print(f"Finished with {len(errors)} error(s)")
        sys.exit(1)
    print("Done")


if __name__ == "__main__":
    main()
