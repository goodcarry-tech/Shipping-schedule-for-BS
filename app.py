"""
船期整理系統 - Shipping Schedule Organizer
支援: CNC (上傳), TSL (上傳), IAL (網頁爬取), KMTC (網頁爬取), YML (網頁爬取)
POL: HAIPHONG | POD: HONG KONG / SHEKOU / KAOHSIUNG / TAICHUNG
"""

import streamlit as st
import pandas as pd
import io
import re
import os
import time
import warnings
from datetime import datetime, timedelta
from copy import copy

warnings.filterwarnings("ignore")

try:
    import pdfplumber
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────
POD_LIST    = ["HONG KONG", "SHEKOU", "KAOHSIUNG", "TAICHUNG"]
OUTPUT_COLS = ["POL", "POD", "Vessel", "Voyage", "ETD", "ETA", "T/T Time", "CY Cut-off", "SI Cut-off"]
CARRIER_LIST = ["CNC", "TSL", "IAL", "KMTC", "YML"]

DAY_TO_NUM = {"MON": 0, "TUE": 1, "WED": 2, "THU": 3, "FRI": 4, "SAT": 5, "SUN": 6}

# ─────────────────────────────────────────────
# Helper utilities
# ─────────────────────────────────────────────

def normalize_pod(pod_str: str) -> str:
    """Map raw POD string to standard form."""
    s = pod_str.strip().upper()
    for std in POD_LIST:
        if std in s:
            return std
    return s

def safe_date_str(val) -> str:
    """Convert various date types to YYYY/MM/DD string."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, (datetime,)):
        return val.strftime("%Y/%m/%d")
    if isinstance(val, str):
        val = val.strip()
        # Try common formats
        for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(val, fmt).strftime("%Y/%m/%d")
            except ValueError:
                pass
        return val
    try:
        import pandas as pd
        if pd.notna(val):
            return pd.Timestamp(val).strftime("%Y/%m/%d")
    except Exception:
        pass
    return str(val)

def days_back_from_etd(etd_date: datetime, day_name: str) -> datetime:
    """
    Given ETD date and a day name (e.g. 'SAT'), return the
    most recent occurrence of that day BEFORE or ON etd_date.
    """
    day_name = day_name.strip().upper()[:3]
    target_wd = DAY_TO_NUM.get(day_name)
    if target_wd is None:
        return etd_date
    etd_wd = etd_date.weekday()
    delta = (etd_wd - target_wd) % 7
    return etd_date - timedelta(days=delta)

def empty_df() -> pd.DataFrame:
    return pd.DataFrame(columns=OUTPUT_COLS)

# ─────────────────────────────────────────────
# CNC Parser
# ─────────────────────────────────────────────

def parse_cnc(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Parse CNC schedule file (CSV or PDF)."""
    rows = []
    fname_lower = filename.lower()

    if fname_lower.endswith(".csv"):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes))
            df.columns = [str(c).strip() for c in df.columns]

            col_map = {
                "Origin":       "POL",
                "Destination":  "POD",
                "Vessel name":  "Vessel",
                "Vessel Name":  "Vessel",
                "Voyage ref.":  "Voyage",
                "Voyage Ref.":  "Voyage",
                "Voyage ref":   "Voyage",
                "Departure Date": "ETD",
                "Arrival Date":   "ETA",
                "Transit Time":   "T/T Time",
                "Port cut-off":   "CY Cut-off",
                "Port Cut-off":   "CY Cut-off",
                "SI cut-off":     "SI Cut-off",
                "SI Cut-off":     "SI Cut-off",
            }
            df.rename(columns={k: v for k, v in col_map.items() if k in df.columns}, inplace=True)

            for _, row in df.iterrows():
                pol = str(row.get("POL", "")).strip().upper()
                pod = normalize_pod(str(row.get("POD", "")))
                if pod not in POD_LIST:
                    continue
                if "HAIPHONG" not in pol:
                    continue

                tt = str(row.get("T/T Time", "")).strip()
                if tt and not tt.lower().endswith("day") and not tt.lower().endswith("days"):
                    tt = f"{tt} days"

                rows.append({
                    "POL":        "HAIPHONG",
                    "POD":        pod,
                    "Vessel":     str(row.get("Vessel", "")).strip(),
                    "Voyage":     str(row.get("Voyage", "")).strip(),
                    "ETD":        safe_date_str(row.get("ETD", "")),
                    "ETA":        safe_date_str(row.get("ETA", "")),
                    "T/T Time":   tt,
                    "CY Cut-off": str(row.get("CY Cut-off", "")).strip(),
                    "SI Cut-off": str(row.get("SI Cut-off", "")).strip(),
                })
        except Exception as e:
            st.error(f"CNC CSV 解析錯誤: {e}")

    elif fname_lower.endswith(".pdf"):
        if not PDF_AVAILABLE:
            st.error("需要安裝 pdfplumber 才能解析 PDF")
            return empty_df()
        try:
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table:
                            continue
                        # Find header row
                        header_idx = None
                        header = []
                        for i, row in enumerate(table):
                            row_text = " ".join([str(c or "") for c in row]).lower()
                            if "vessel" in row_text or "voyage" in row_text or "departure" in row_text:
                                header_idx = i
                                header = [str(c or "").strip() for c in row]
                                break
                        if header_idx is None:
                            continue

                        col_map_pdf = {}
                        for j, h in enumerate(header):
                            hl = h.lower()
                            if "origin" in hl:        col_map_pdf["POL"]        = j
                            elif "destination" in hl: col_map_pdf["POD"]        = j
                            elif "vessel name" in hl: col_map_pdf["Vessel"]     = j
                            elif "voyage" in hl:      col_map_pdf["Voyage"]     = j
                            elif "departure" in hl:   col_map_pdf["ETD"]        = j
                            elif "arrival date" in hl: col_map_pdf["ETA"]       = j
                            elif "transit" in hl:     col_map_pdf["T/T Time"]   = j
                            elif "port cut" in hl:    col_map_pdf["CY Cut-off"] = j
                            elif "si cut" in hl:      col_map_pdf["SI Cut-off"] = j

                        for row in table[header_idx + 1:]:
                            def g(key):
                                idx = col_map_pdf.get(key)
                                return str(row[idx] or "").strip() if idx is not None and idx < len(row) else ""

                            pol = g("POL").upper()
                            pod = normalize_pod(g("POD"))
                            if pod not in POD_LIST:
                                continue

                            tt = g("T/T Time")
                            if tt and not tt.lower().endswith("day") and not tt.lower().endswith("days"):
                                tt = f"{tt} days"

                            rows.append({
                                "POL":        "HAIPHONG",
                                "POD":        pod,
                                "Vessel":     g("Vessel"),
                                "Voyage":     g("Voyage"),
                                "ETD":        safe_date_str(g("ETD")),
                                "ETA":        safe_date_str(g("ETA")),
                                "T/T Time":   tt,
                                "CY Cut-off": g("CY Cut-off"),
                                "SI Cut-off": g("SI Cut-off"),
                            })
        except Exception as e:
            st.error(f"CNC PDF 解析錯誤: {e}")

    else:
        # Try Excel
        try:
            df = pd.read_excel(io.BytesIO(file_bytes))
            df.columns = [str(c).strip() for c in df.columns]
            col_map = {
                "Origin": "POL", "Destination": "POD",
                "Vessel name": "Vessel", "Vessel Name": "Vessel",
                "Voyage ref.": "Voyage", "Voyage ref": "Voyage",
                "Departure Date": "ETD", "Arrival Date": "ETA",
                "Transit Time": "T/T Time",
                "Port cut-off": "CY Cut-off",
                "SI cut-off": "SI Cut-off",
            }
            df.rename(columns={k: v for k, v in col_map.items() if k in df.columns}, inplace=True)
            for _, row in df.iterrows():
                pod = normalize_pod(str(row.get("POD", "")))
                if pod not in POD_LIST:
                    continue
                tt = str(row.get("T/T Time", "")).strip()
                if tt and not tt.lower().endswith("day") and not tt.lower().endswith("days"):
                    tt = f"{tt} days"
                rows.append({
                    "POL": "HAIPHONG", "POD": pod,
                    "Vessel": str(row.get("Vessel", "")).strip(),
                    "Voyage": str(row.get("Voyage", "")).strip(),
                    "ETD": safe_date_str(row.get("ETD", "")),
                    "ETA": safe_date_str(row.get("ETA", "")),
                    "T/T Time": tt,
                    "CY Cut-off": str(row.get("CY Cut-off", "")).strip(),
                    "SI Cut-off": str(row.get("SI Cut-off", "")).strip(),
                })
        except Exception as e:
            st.error(f"CNC Excel 解析錯誤: {e}")

    return pd.DataFrame(rows, columns=OUTPUT_COLS) if rows else empty_df()


# ─────────────────────────────────────────────
# TSL Parser
# ─────────────────────────────────────────────

def parse_tsl(file_bytes: bytes, filename: str,
              start_date: datetime, end_date: datetime) -> pd.DataFrame:
    """
    Parse TSL schedule file (Excel or PDF).
    Handles day-based CY cut-off and SI cut-off calculation.
    """
    rows = []
    fname_lower = filename.lower()

    def _parse_tsl_excel(wb):
        for ws in wb.worksheets:
            _parse_tsl_sheet(ws, rows, start_date, end_date)

    def _parse_tsl_sheet(ws, out_rows, start_dt, end_dt):
        # Read all cells into a 2D list for easy traversal
        data = []
        for row in ws.iter_rows(values_only=True):
            data.append([str(c).strip() if c is not None else "" for c in row])

        if not data:
            return

        # ── Find the main header row ──────────────────────────────────────
        header_row_idx = None
        header = []
        for i, row in enumerate(data):
            row_upper = [c.upper() for c in row]
            if ("VESSEL" in row_upper or "VESSEL NAME" in " ".join(row_upper)) and \
               ("VOYAGE" in row_upper or any("VOYAGE" in c for c in row_upper)):
                header_row_idx = i
                header = row_upper
                break

        if header_row_idx is None:
            # Try to find any row with ETD
            for i, row in enumerate(data):
                row_upper = [c.upper() for c in row]
                if any("ETD" in c for c in row_upper) and any("ETA" in c for c in row_upper):
                    header_row_idx = i
                    header = row_upper
                    break

        if header_row_idx is None:
            return

        # ── Identify column indices ───────────────────────────────────────
        def find_col(keywords):
            for j, h in enumerate(header):
                if any(kw.upper() in h for kw in keywords):
                    return j
            return None

        c_service = find_col(["SERVICE"])
        c_vessel  = find_col(["VESSEL"])
        c_voyage  = find_col(["VOYAGE"])
        c_etd     = find_col(["ETD"])
        c_eta     = find_col(["ETA"])

        # POD columns: columns with POD port names in header
        pod_cols = {}
        for j, h in enumerate(header):
            for pod in POD_LIST:
                if pod in h:
                    pod_cols[pod] = j
                    break

        if c_vessel is None or c_voyage is None or c_etd is None:
            return

        # ── Look for SERVICE reference table (cut-off rules) ─────────────
        # Format: SERVICE NAME | ETD day | CY time+day | SI time+day
        service_rules = {}  # {service_name: {"cy_day": str, "cy_time": str, "si_day": str, "si_time": str}}

        # Search for a "SERVICE NAME" section anywhere before/after data
        for i, row in enumerate(data):
            row_upper = [c.upper() for c in row]
            if "SERVICE NAME" in " ".join(row_upper) or any("SERVICE NAME" in c for c in row_upper):
                # This row or next rows contain the reference table
                svc_header_idx = i
                # Column positions in this section
                svc_col = next((j for j, c in enumerate(row_upper) if "SERVICE" in c and "NAME" in c), 0)
                etd_hph_col = next((j for j, c in enumerate(row_upper) if "ETD" in c and ("HPH" in c or "HAI" in c or "DEP" in c)), None)
                cy_col  = next((j for j, c in enumerate(row_upper) if "CY" in c and "CUT" in c.replace(" ", "") or c == "CY"), None)
                si_col  = next((j for j, c in enumerate(row_upper) if "SUBMIT" in c or ("SI" in c and "VGM" in c) or c == "SI/VGM"), None)

                if etd_hph_col is None:
                    etd_hph_col = next((j for j, c in enumerate(row_upper) if "ETD HPH" in c or "ETD" in c), None)
                if cy_col is None:
                    cy_col = next((j for j, c in enumerate(row_upper) if "CY" in c), None)
                if si_col is None:
                    si_col = next((j for j, c in enumerate(row_upper) if "SI" in c or "VGM" in c), None)

                # Read subsequent rows as service rules
                for k in range(svc_header_idx + 1, min(svc_header_idx + 30, len(data))):
                    rule_row = data[k]
                    if not any(rule_row):
                        break
                    svc_name = rule_row[svc_col].strip().upper() if svc_col is not None and svc_col < len(rule_row) else ""
                    if not svc_name:
                        continue

                    etd_day_raw = rule_row[etd_hph_col].strip().upper() if etd_hph_col is not None and etd_hph_col < len(rule_row) else ""
                    cy_raw  = rule_row[cy_col].strip().upper()  if cy_col  is not None and cy_col  < len(rule_row) else ""
                    si_raw  = rule_row[si_col].strip().upper()  if si_col  is not None and si_col  < len(rule_row) else ""

                    # Parse cy_raw like "24:00 SAT" or "SAT 24:00"
                    cy_time, cy_day = _parse_time_day(cy_raw)
                    si_time, si_day = _parse_time_day(si_raw)
                    etd_day = etd_day_raw[:3] if etd_day_raw else ""

                    service_rules[svc_name] = {
                        "etd_day": etd_day,
                        "cy_day": cy_day, "cy_time": cy_time,
                        "si_day": si_day, "si_time": si_time,
                    }
                break

        # ── Parse data rows ───────────────────────────────────────────────
        for i in range(header_row_idx + 1, len(data)):
            row = data[i]
            if not any(row):
                continue

            vessel = row[c_vessel].strip() if c_vessel < len(row) else ""
            voyage = row[c_voyage].strip() if c_voyage < len(row) else ""
            etd_raw = row[c_etd].strip() if c_etd < len(row) else ""
            eta_raw = row[c_eta].strip() if c_eta < len(row) else ""

            if not vessel or not voyage or not etd_raw:
                continue
            if vessel.upper() in ("VESSEL", "VESSEL NAME", ""):
                continue

            # Parse ETD date
            etd_str = safe_date_str(etd_raw)
            if not etd_str:
                continue

            try:
                etd_dt = datetime.strptime(etd_str, "%Y/%m/%d")
            except ValueError:
                continue

            # Filter by date range
            if etd_dt < start_dt or etd_dt > end_dt:
                continue

            # Parse ETA (may have transit time in brackets)
            eta_str = ""
            tt_str = ""
            eta_raw_clean = re.sub(r"\(.*?\)", "", eta_raw).strip()
            eta_str = safe_date_str(eta_raw_clean)

            # Extract T/T from brackets e.g. "(2 days)"
            tt_match = re.search(r"\((\d+)\s*days?\)", eta_raw, re.IGNORECASE)
            if tt_match:
                tt_str = f"{tt_match.group(1)} days"
            elif eta_str and etd_str:
                try:
                    eta_dt = datetime.strptime(eta_str, "%Y/%m/%d")
                    tt_days = (eta_dt - etd_dt).days
                    if tt_days >= 0:
                        tt_str = f"{tt_days} days"
                except Exception:
                    pass

            # Get service name
            svc_name = ""
            if c_service is not None and c_service < len(row):
                svc_name = row[c_service].strip().upper()

            # Calculate CY and SI cut-off
            cy_cutoff = ""
            si_cutoff = ""
            rule = service_rules.get(svc_name, {})
            if rule:
                cy_day  = rule.get("cy_day", "")
                cy_time = rule.get("cy_time", "")
                si_day  = rule.get("si_day", "")
                si_time = rule.get("si_time", "")
                if cy_day:
                    cy_dt = days_back_from_etd(etd_dt, cy_day)
                    cy_cutoff = cy_dt.strftime("%Y/%m/%d") + (f" {cy_time}" if cy_time else "")
                if si_day:
                    si_dt = days_back_from_etd(etd_dt, si_day)
                    si_cutoff = si_dt.strftime("%Y/%m/%d") + (f" {si_time}" if si_time else "")

            # Determine PODs for this row
            pods_in_row = []
            if pod_cols:
                for pod, j in pod_cols.items():
                    if j < len(row) and row[j].strip():
                        pods_in_row.append(pod)
            else:
                # Try to detect POD from row itself
                row_upper = " ".join(row).upper()
                for pod in POD_LIST:
                    if pod in row_upper:
                        pods_in_row.append(pod)

            if not pods_in_row:
                # Default: check if there's explicit POD column
                pods_in_row = [POD_LIST[0]]  # fallback

            for pod in pods_in_row:
                out_rows.append({
                    "POL": "HAIPHONG", "POD": pod,
                    "Vessel": vessel, "Voyage": voyage,
                    "ETD": etd_str, "ETA": eta_str,
                    "T/T Time": tt_str,
                    "CY Cut-off": cy_cutoff,
                    "SI Cut-off": si_cutoff,
                })

    def _parse_time_day(raw: str):
        """Parse strings like '24:00 SAT' or 'SAT 09:00' or '09:00 FRI'"""
        raw = raw.strip().upper()
        time_match = re.search(r"\d{1,2}:\d{2}", raw)
        day_match  = re.search(r"\b(MON|TUE|WED|THU|FRI|SAT|SUN)\b", raw)
        t = time_match.group(0) if time_match else ""
        d = day_match.group(0)[:3]  if day_match  else ""
        return t, d

    if fname_lower.endswith((".xlsx", ".xls")):
        try:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            _parse_tsl_excel(wb)
        except Exception as e:
            st.error(f"TSL Excel 解析錯誤: {e}")
    elif fname_lower.endswith(".pdf"):
        if not PDF_AVAILABLE:
            st.error("需要安裝 pdfplumber 才能解析 PDF")
            return empty_df()
        try:
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table:
                            continue
                        # Build a minimal worksheet-like structure
                        # Find VESSEL in first few rows
                        header_idx = None
                        for i, row in enumerate(table):
                            if any("VESSEL" in str(c or "").upper() for c in row):
                                header_idx = i
                                break
                        if header_idx is None:
                            continue

                        header = [str(c or "").strip().upper() for c in table[header_idx]]

                        def find_col_pdf(kws):
                            for j, h in enumerate(header):
                                if any(k.upper() in h for k in kws):
                                    return j
                            return None

                        c_v   = find_col_pdf(["VESSEL"])
                        c_voy = find_col_pdf(["VOYAGE"])
                        c_svc = find_col_pdf(["SERVICE"])
                        c_etd_p = find_col_pdf(["ETD"])
                        c_eta_p = find_col_pdf(["ETA"])

                        if c_v is None or c_voy is None or c_etd_p is None:
                            continue

                        for row in table[header_idx + 1:]:
                            def gp(idx):
                                return str(row[idx] or "").strip() if idx is not None and idx < len(row) else ""

                            vessel  = gp(c_v)
                            voyage  = gp(c_voy)
                            etd_raw = gp(c_etd_p)
                            eta_raw = gp(c_eta_p)

                            if not vessel or not etd_raw:
                                continue

                            etd_str = safe_date_str(etd_raw)
                            if not etd_str:
                                continue
                            try:
                                etd_dt = datetime.strptime(etd_str, "%Y/%m/%d")
                            except ValueError:
                                continue

                            if etd_dt < start_date or etd_dt > end_date:
                                continue

                            eta_str = safe_date_str(re.sub(r"\(.*?\)", "", eta_raw).strip())
                            tt_match = re.search(r"\((\d+)\s*days?\)", eta_raw, re.IGNORECASE)
                            tt_str = f"{tt_match.group(1)} days" if tt_match else ""

                            # No service rules for PDF; CY/SI left blank
                            rows.append({
                                "POL": "HAIPHONG", "POD": "HONG KONG",
                                "Vessel": vessel, "Voyage": voyage,
                                "ETD": etd_str, "ETA": eta_str,
                                "T/T Time": tt_str,
                                "CY Cut-off": "",
                                "SI Cut-off": "",
                            })
        except Exception as e:
            st.error(f"TSL PDF 解析錯誤: {e}")
    else:
        st.warning(f"TSL 不支援的檔案格式: {filename}")

    return pd.DataFrame(rows, columns=OUTPUT_COLS) if rows else empty_df()


# ─────────────────────────────────────────────
# Web Scraping (Selenium)
# ─────────────────────────────────────────────

def get_driver():
    """Create a Selenium headless Chrome driver."""
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service

        options = Options()
        options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--lang=zh-TW")
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36")

        # Try system chromium first (Streamlit Cloud)
        for binary in ["/usr/bin/chromium", "/usr/bin/chromium-browser",
                       "/usr/bin/google-chrome", "/usr/bin/google-chrome-stable"]:
            if os.path.exists(binary):
                options.binary_location = binary
                break

        for driver_path in ["/usr/bin/chromedriver", "/usr/lib/chromium/chromedriver",
                            "/usr/lib/chromium-browser/chromedriver"]:
            if os.path.exists(driver_path):
                service = Service(driver_path)
                return webdriver.Chrome(service=service, options=options)

        # Fallback: webdriver_manager
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            service = Service(ChromeDriverManager().install())
            return webdriver.Chrome(service=service, options=options)
        except Exception:
            pass

        # Last resort: no service arg
        return webdriver.Chrome(options=options)

    except Exception as e:
        st.error(f"無法啟動瀏覽器驅動: {e}\n請確認已安裝 chromium / chromedriver")
        return None


def scrape_ial(year: int, month: int, pods: list) -> pd.DataFrame:
    """
    Scrape IAL schedule from https://www.interasia.cc/Service/Form?servicetype=3
    POL: HAIPHONG -> PODs: HONG KONG, SHEKOU, KAOHSIUNG, TAICHUNG
    """
    try:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import Select, WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
    except ImportError:
        st.error("請安裝 selenium: pip install selenium")
        return empty_df()

    # POD -> (Country, Port) mapping for IAL form
    pod_country_map = {
        "HONG KONG": ("HONG KONG", "HONG KONG"),
        "SHEKOU":    ("CHINA",     "SHEKOU"),
        "KAOHSIUNG": ("TAIWAN",    "KAOHSIUNG"),
        "TAICHUNG":  ("TAIWAN",    "TAICHUNG"),
    }

    rows = []
    driver = get_driver()
    if driver is None:
        return empty_df()

    try:
        url = "https://www.interasia.cc/Service/Form?servicetype=3"
        for pod in pods:
            if pod not in pod_country_map:
                continue
            country, port = pod_country_map[pod]
            try:
                driver.get(url)
                wait = WebDriverWait(driver, 20)

                # Select origin: VIETNAM
                try:
                    orig_sel = Select(wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//select[contains(@id,'pol_country') or contains(@name,'pol_country') or contains(@id,'origin_country')]"))))
                    orig_sel.select_by_visible_text("VIETNAM")
                    time.sleep(1)
                except Exception:
                    # Try selecting any "origin" dropdown
                    selects = driver.find_elements(By.TAG_NAME, "select")
                    if len(selects) >= 1:
                        Select(selects[0]).select_by_visible_text("VIETNAM")
                        time.sleep(1)

                # Select POL: HAIPHONG
                try:
                    pol_sel = Select(wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//select[contains(@id,'pol') and not(contains(@id,'country'))]"))))
                    pol_sel.select_by_visible_text("HAIPHONG")
                    time.sleep(1)
                except Exception:
                    selects = driver.find_elements(By.TAG_NAME, "select")
                    if len(selects) >= 2:
                        try:
                            Select(selects[1]).select_by_visible_text("HAIPHONG")
                            time.sleep(1)
                        except Exception:
                            pass

                # Select destination country
                try:
                    dst_sel = Select(wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//select[contains(@id,'pod_country') or contains(@name,'pod_country') or contains(@id,'dest_country')]"))))
                    dst_sel.select_by_visible_text(country)
                    time.sleep(1)
                except Exception:
                    selects = driver.find_elements(By.TAG_NAME, "select")
                    if len(selects) >= 3:
                        try:
                            Select(selects[2]).select_by_visible_text(country)
                            time.sleep(1)
                        except Exception:
                            pass

                # Select POD port
                try:
                    pod_sel = Select(wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//select[contains(@id,'pod') and not(contains(@id,'country'))]"))))
                    pod_sel.select_by_visible_text(port)
                    time.sleep(1)
                except Exception:
                    selects = driver.find_elements(By.TAG_NAME, "select")
                    if len(selects) >= 4:
                        try:
                            Select(selects[3]).select_by_visible_text(port)
                            time.sleep(1)
                        except Exception:
                            pass

                # Click search button
                try:
                    search_btn = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(@class,'search') or contains(text(),'查詢') or contains(text(),'Search')]")))
                    search_btn.click()
                except Exception:
                    buttons = driver.find_elements(By.TAG_NAME, "button")
                    for btn in buttons:
                        txt = btn.text.strip()
                        if any(k in txt for k in ["查詢", "Search", "SEARCH", "搜尋"]):
                            btn.click()
                            break
                time.sleep(3)

                # Extract results table
                tables = driver.find_elements(By.TAG_NAME, "table")
                for table in tables:
                    header_cells = table.find_elements(By.XPATH, ".//th")
                    if not header_cells:
                        continue
                    headers = [h.text.strip() for h in header_cells]

                    def find_col_ial(kws):
                        for j, h in enumerate(headers):
                            if any(k.lower() in h.lower() for k in kws):
                                return j
                        return None

                    c_vessel  = find_col_ial(["出發船名", "vessel", "船名"])
                    c_voyage  = find_col_ial(["出發航次", "voyage", "航次"])
                    c_etd_i   = find_col_ial(["出發日期", "etd", "departure"])
                    c_eta_i   = find_col_ial(["抵達日期", "eta", "arrival"])
                    c_tt      = find_col_ial(["預估運輸", "transit", "t/t"])

                    if c_vessel is None or c_etd_i is None:
                        continue

                    tbody_rows = table.find_elements(By.XPATH, ".//tbody/tr")
                    for tr in tbody_rows:
                        cells = tr.find_elements(By.TAG_NAME, "td")
                        if not cells:
                            continue
                        def gc(idx):
                            return cells[idx].text.strip() if idx is not None and idx < len(cells) else ""

                        vessel  = gc(c_vessel)
                        voyage  = gc(c_voyage)
                        etd_raw = gc(c_etd_i)
                        eta_raw = gc(c_eta_i)
                        tt_raw  = gc(c_tt)

                        if not vessel or not etd_raw:
                            continue

                        etd_str = safe_date_str(etd_raw)
                        eta_str = safe_date_str(eta_raw)

                        # Filter by month/year
                        try:
                            etd_dt = datetime.strptime(etd_str, "%Y/%m/%d")
                            if etd_dt.year != year or etd_dt.month != month:
                                continue
                        except Exception:
                            pass

                        tt_str = tt_raw if tt_raw else ""
                        if tt_str and re.match(r"^\d+$", tt_str):
                            tt_str = f"{tt_str} Days"

                        rows.append({
                            "POL": "HAIPHONG", "POD": pod,
                            "Vessel": vessel, "Voyage": voyage,
                            "ETD": etd_str, "ETA": eta_str,
                            "T/T Time": tt_str,
                            "CY Cut-off": "",
                            "SI Cut-off": "",
                        })

            except Exception as e:
                st.warning(f"IAL 爬取 {pod} 時發生錯誤: {e}")
                continue

    finally:
        driver.quit()

    return pd.DataFrame(rows, columns=OUTPUT_COLS) if rows else empty_df()


def scrape_kmtc(year: int, month: int, pods: list) -> pd.DataFrame:
    """
    Scrape KMTC schedule from https://www.ekmtc.com
    POL: HAIPHONG -> PODs: HONG KONG, SHEKOU, KAOHSIUNG, TAICHUNG
    """
    try:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import Select, WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
    except ImportError:
        st.error("請安裝 selenium: pip install selenium")
        return empty_df()

    # KMTC URL parameters for each POD
    pod_url_map = {
        "HONG KONG": f"https://www.ekmtc.com/index.html#/schedule/leg?porCtrCd=VN&porPlcCd=HPH&dlyCtrCd=HK&dlyPlcCd=HKG&yyyymm={year:04d}{month:02d}&loginChk=",
        "SHEKOU":    f"https://www.ekmtc.com/index.html#/schedule/leg?porCtrCd=VN&porPlcCd=HPH&dlyCtrCd=CN&dlyPlcCd=SKU&yyyymm={year:04d}{month:02d}&loginChk=",
        "KAOHSIUNG": f"https://www.ekmtc.com/index.html#/schedule/leg?porCtrCd=VN&porPlcCd=HPH&dlyCtrCd=TW&dlyPlcCd=KHH&yyyymm={year:04d}{month:02d}&loginChk=",
        "TAICHUNG":  f"https://www.ekmtc.com/index.html#/schedule/leg?porCtrCd=VN&porPlcCd=HPH&dlyCtrCd=TW&dlyPlcCd=TXG&yyyymm={year:04d}{month:02d}&loginChk=",
    }

    rows = []
    driver = get_driver()
    if driver is None:
        return empty_df()

    try:
        for pod in pods:
            if pod not in pod_url_map:
                continue
            url = pod_url_map[pod]
            try:
                driver.get(url)
                wait = WebDriverWait(driver, 25)
                time.sleep(4)  # Wait for SPA to load

                # Click Search button if present
                try:
                    search_btn = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(text(),'Search') or contains(@class,'search-btn') or contains(@id,'searchBtn')]")))
                    search_btn.click()
                    time.sleep(3)
                except Exception:
                    pass

                # KMTC shows a calendar; find all date links
                date_links = driver.find_elements(By.XPATH,
                    "//td[contains(@class,'calendar') or contains(@class,'schedule')]//a | "
                    "//div[contains(@class,'cal')]//a | //td//a[contains(@href,'schedule')]")

                if not date_links:
                    # Try to find any clickable schedule items
                    date_links = driver.find_elements(By.XPATH,
                        "//a[contains(@class,'sched') or contains(@onclick,'schedule')]")

                schedule_urls = set()
                for link in date_links:
                    href = link.get_attribute("href") or ""
                    if href and "schedule" in href.lower():
                        schedule_urls.add(href)

                # If no links found, try to parse the current page table
                if not schedule_urls:
                    _extract_kmtc_table(driver, pod, year, month, rows)
                else:
                    for sched_url in list(schedule_urls)[:50]:
                        try:
                            driver.get(sched_url)
                            time.sleep(2)
                            _extract_kmtc_detail(driver, pod, year, month, rows)
                        except Exception:
                            continue

                # Also try to parse any visible schedule tables on current page
                _extract_kmtc_table(driver, pod, year, month, rows)

            except Exception as e:
                st.warning(f"KMTC 爬取 {pod} 時發生錯誤: {e}")
                continue

    finally:
        driver.quit()

    return pd.DataFrame(rows, columns=OUTPUT_COLS) if rows else empty_df()


def _extract_kmtc_table(driver, pod: str, year: int, month: int, rows: list):
    """Extract schedule info from KMTC table on current page."""
    try:
        from selenium.webdriver.by import By
        tables = driver.find_elements(By.TAG_NAME, "table")
        for table in tables:
            ths = table.find_elements(By.TAG_NAME, "th")
            headers = [th.text.strip().lower() for th in ths]
            if not any("vessel" in h or "departure" in h or "arrival" in h for h in headers):
                continue

            def fci(kws):
                for j, h in enumerate(headers):
                    if any(k.lower() in h for k in kws):
                        return j
                return None

            c_vessel  = fci(["vessel", "ship"])
            c_voyage  = fci(["voyage"])
            c_etd     = fci(["departure", "etd"])
            c_eta     = fci(["arrival", "eta"])
            c_tt      = fci(["t/t", "transit", "total t"])
            c_cy      = fci(["cy cut", "cy"])
            c_si      = fci(["vgm", "si cut", "si closing"])

            tbody_rows = table.find_elements(By.XPATH, ".//tbody/tr")
            for tr in tbody_rows:
                tds = tr.find_elements(By.TAG_NAME, "td")
                def gc(i): return tds[i].text.strip() if i is not None and i < len(tds) else ""

                vessel  = gc(c_vessel)
                voyage  = gc(c_voyage)
                etd_str = safe_date_str(gc(c_etd))
                eta_str = safe_date_str(gc(c_eta))
                tt_str  = gc(c_tt)
                cy_str  = gc(c_cy)
                si_str  = gc(c_si)

                if not vessel or not etd_str:
                    continue
                try:
                    etd_dt = datetime.strptime(etd_str, "%Y/%m/%d")
                    if etd_dt.year != year or etd_dt.month != month:
                        continue
                except Exception:
                    pass

                if tt_str and re.match(r"^\d+$", tt_str):
                    tt_str = f"{tt_str} Days"

                # Avoid duplicates
                is_dup = any(
                    r["Vessel"] == vessel and r["Voyage"] == voyage
                    and r["POD"] == pod for r in rows
                )
                if not is_dup:
                    rows.append({
                        "POL": "HAIPHONG", "POD": pod,
                        "Vessel": vessel, "Voyage": voyage,
                        "ETD": etd_str, "ETA": eta_str,
                        "T/T Time": tt_str,
                        "CY Cut-off": cy_str,
                        "SI Cut-off": si_str,
                    })
    except Exception:
        pass


def _extract_kmtc_detail(driver, pod: str, year: int, month: int, rows: list):
    """Extract from KMTC detail page."""
    try:
        from selenium.webdriver.by import By
        page_text = driver.find_element(By.TAG_NAME, "body").text

        vessel_m  = re.search(r"Vessel[/\s]+Voyage[:\s]+([A-Z][A-Z\s]+?)\s+(\w+)", page_text)
        etd_m     = re.search(r"Date of Departure[:\s]+([\d/\-]+)", page_text)
        eta_m     = re.search(r"Date of Arrival[:\s]+([\d/\-]+)", page_text)
        tt_m      = re.search(r"Total T/T[:\s]+(\d+\s*Days?)", page_text, re.IGNORECASE)
        cy_m      = re.search(r"CY CUT[:\s]+([\d/\-\s:]+)", page_text, re.IGNORECASE)
        vgm_m     = re.search(r"VGM closing[:\s]+([\d/\-\s:]+)", page_text, re.IGNORECASE)

        if vessel_m and etd_m:
            vessel  = vessel_m.group(1).strip()
            voyage  = vessel_m.group(2).strip()
            etd_str = safe_date_str(etd_m.group(1).strip())
            eta_str = safe_date_str(eta_m.group(1).strip()) if eta_m else ""
            tt_str  = tt_m.group(1).strip() if tt_m else ""
            cy_str  = cy_m.group(1).strip()  if cy_m  else ""
            si_str  = vgm_m.group(1).strip() if vgm_m else ""

            try:
                etd_dt = datetime.strptime(etd_str, "%Y/%m/%d")
                if etd_dt.year != year or etd_dt.month != month:
                    return
            except Exception:
                return

            is_dup = any(
                r["Vessel"] == vessel and r["Voyage"] == voyage
                and r["POD"] == pod for r in rows
            )
            if not is_dup:
                rows.append({
                    "POL": "HAIPHONG", "POD": pod,
                    "Vessel": vessel, "Voyage": voyage,
                    "ETD": etd_str, "ETA": eta_str,
                    "T/T Time": tt_str,
                    "CY Cut-off": cy_str,
                    "SI Cut-off": si_str,
                })
    except Exception:
        pass


def scrape_yml(year: int, month: int, pods: list) -> pd.DataFrame:
    """
    Scrape YML schedule from https://www.yangming.com/en/esolution/schedule/point_to_point_search
    POL: HAIPHONG -> PODs: HONG KONG, SHEKOU, KAOHSIUNG, TAICHUNG
    """
    try:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.common.keys import Keys
    except ImportError:
        st.error("請安裝 selenium: pip install selenium")
        return empty_df()

    pod_keyword_map = {
        "HONG KONG": "Hong Kong",
        "SHEKOU":    "Shekou",
        "KAOHSIUNG": "Kaohsiung",
        "TAICHUNG":  "Taichung",
    }

    rows = []
    driver = get_driver()
    if driver is None:
        return empty_df()

    base_url = "https://www.yangming.com/en/esolution/schedule/point_to_point_search"

    try:
        for pod in pods:
            if pod not in pod_keyword_map:
                continue
            pod_kw = pod_keyword_map[pod]
            try:
                driver.get(base_url)
                wait = WebDriverWait(driver, 20)
                time.sleep(3)

                # Click Point-To-Point tab if needed
                try:
                    ptp_tab = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//a[contains(text(),'Point') or contains(text(),'point')]")))
                    ptp_tab.click()
                    time.sleep(2)
                except Exception:
                    pass

                # Fill "From" field: HaiPhong
                try:
                    from_input = wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//input[contains(@placeholder,'From') or contains(@id,'from') or contains(@name,'from')]")))
                    from_input.clear()
                    from_input.send_keys("HaiPhong")
                    time.sleep(1)
                    # Select autocomplete suggestion
                    suggestions = driver.find_elements(By.XPATH,
                        "//li[contains(text(),'Haiphong') or contains(text(),'HAIPHONG') or contains(text(),'HaiPhong')]")
                    if suggestions:
                        suggestions[0].click()
                        time.sleep(1)
                    else:
                        from_input.send_keys(Keys.RETURN)
                        time.sleep(1)
                except Exception as e:
                    st.warning(f"YML From 欄位填寫失敗: {e}")

                # Fill "To" field: destination
                try:
                    to_input = wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//input[contains(@placeholder,'To') or contains(@id,'to') or contains(@name,'to')]")))
                    to_input.clear()
                    to_input.send_keys(pod_kw)
                    time.sleep(1)
                    suggestions = driver.find_elements(By.XPATH,
                        f"//li[contains(text(),'{pod_kw}')]")
                    if suggestions:
                        suggestions[0].click()
                        time.sleep(1)
                    else:
                        to_input.send_keys(Keys.RETURN)
                        time.sleep(1)
                except Exception as e:
                    st.warning(f"YML To 欄位填寫失敗: {e}")

                # Set period (select the month)
                yyyymm = f"{year:04d}/{month:02d}"
                try:
                    period_input = driver.find_elements(By.XPATH,
                        "//input[contains(@id,'period') or contains(@name,'period') or contains(@placeholder,'Period')]")
                    if period_input:
                        period_input[0].clear()
                        period_input[0].send_keys(yyyymm)
                        time.sleep(1)
                except Exception:
                    pass

                # Click Search
                try:
                    search_btn = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(text(),'Search') or @type='submit']")))
                    search_btn.click()
                    time.sleep(4)
                except Exception as e:
                    st.warning(f"YML Search 按鈕點擊失敗: {e}")

                # Extract results
                tables = driver.find_elements(By.TAG_NAME, "table")
                for table in tables:
                    ths = table.find_elements(By.TAG_NAME, "th")
                    headers = [th.text.strip().lower() for th in ths]
                    if not any("vessel" in h or "etd" in h or "departure" in h for h in headers):
                        continue

                    def fci(kws):
                        for j, h in enumerate(headers):
                            if any(k.lower() in h for k in kws):
                                return j
                        return None

                    c_vessel  = fci(["vessel"])
                    c_voyage  = fci(["voyage"])
                    c_etd_y   = fci(["etd", "departure"])
                    c_eta_y   = fci(["eta", "arrival"])
                    c_tt      = fci(["t/t", "transit"])
                    c_cy      = fci(["cy cut", "cy"])
                    c_si      = fci(["si cut", "si "])

                    tbody_rows = table.find_elements(By.XPATH, ".//tbody/tr")
                    for tr in tbody_rows:
                        tds = tr.find_elements(By.TAG_NAME, "td")
                        def gc(i): return tds[i].text.strip() if i is not None and i < len(tds) else ""

                        vessel  = gc(c_vessel)
                        voyage  = gc(c_voyage)
                        etd_str = safe_date_str(gc(c_etd_y))
                        eta_str = safe_date_str(gc(c_eta_y))
                        tt_str  = gc(c_tt)
                        cy_str  = gc(c_cy)
                        si_str  = gc(c_si)

                        if not vessel or not etd_str:
                            continue

                        try:
                            etd_dt = datetime.strptime(etd_str, "%Y/%m/%d")
                            if etd_dt.year != year or etd_dt.month != month:
                                continue
                        except Exception:
                            pass

                        if tt_str and re.match(r"^\d+$", tt_str):
                            tt_str = f"{tt_str} Days"

                        rows.append({
                            "POL": "HAIPHONG", "POD": pod,
                            "Vessel": vessel, "Voyage": voyage,
                            "ETD": etd_str, "ETA": eta_str,
                            "T/T Time": tt_str,
                            "CY Cut-off": cy_str,
                            "SI Cut-off": si_str,
                        })

                # Also try to parse text-based results (YML may use divs)
                result_divs = driver.find_elements(By.XPATH,
                    "//div[contains(@class,'result') or contains(@class,'schedule-item') or contains(@class,'row-item')]")
                for div in result_divs:
                    text = div.text
                    vessel_m = re.search(r"(?:Vessel|Ship)[:\s]+([A-Z][A-Z\s]+?)(?:\n|Voyage|/)", text, re.IGNORECASE)
                    voyage_m = re.search(r"(?:Voyage)[:\s]*(\w+)", text, re.IGNORECASE)
                    etd_m    = re.search(r"(?:ETD|Departure)[:\s]*([\d]{4}[/\-][\d]{2}[/\-][\d]{2})", text, re.IGNORECASE)
                    eta_m    = re.search(r"(?:ETA|Arrival)[:\s]*([\d]{4}[/\-][\d]{2}[/\-][\d]{2})", text, re.IGNORECASE)
                    tt_m     = re.search(r"(?:T/T|Transit)[:\s]*(\d+\s*Days?)", text, re.IGNORECASE)
                    cy_m     = re.search(r"(?:CY)[:\s]*([\d/\-\s:]+)", text, re.IGNORECASE)
                    si_m     = re.search(r"(?:SI Cut)[:\s]*([\w\s]+)", text, re.IGNORECASE)

                    if vessel_m and etd_m:
                        etd_str = safe_date_str(etd_m.group(1))
                        try:
                            etd_dt = datetime.strptime(etd_str, "%Y/%m/%d")
                            if etd_dt.year != year or etd_dt.month != month:
                                continue
                        except Exception:
                            continue

                        vessel = vessel_m.group(1).strip()
                        voyage = voyage_m.group(1).strip() if voyage_m else ""
                        is_dup = any(
                            r["Vessel"] == vessel and r["Voyage"] == voyage
                            and r["POD"] == pod for r in rows
                        )
                        if not is_dup:
                            rows.append({
                                "POL": "HAIPHONG", "POD": pod,
                                "Vessel": vessel,
                                "Voyage": voyage,
                                "ETD": etd_str,
                                "ETA": safe_date_str(eta_m.group(1)) if eta_m else "",
                                "T/T Time": tt_m.group(1).strip() if tt_m else "",
                                "CY Cut-off": cy_m.group(1).strip() if cy_m else "",
                                "SI Cut-off": si_m.group(1).strip() if si_m else "",
                            })

            except Exception as e:
                st.warning(f"YML 爬取 {pod} 時發生錯誤: {e}")
                continue

    finally:
        driver.quit()

    return pd.DataFrame(rows, columns=OUTPUT_COLS) if rows else empty_df()


# ─────────────────────────────────────────────
# Excel Export
# ─────────────────────────────────────────────

def export_to_excel(all_data: pd.DataFrame) -> bytes:
    """Create formatted Excel with one sheet per POD, sorted by ETD."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Styles
    header_fill  = PatternFill("solid", fgColor="1F4788")
    header_font  = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    alt_fill     = PatternFill("solid", fgColor="DCE6F1")
    normal_font  = Font(name="Calibri", size=10)
    center_align = Alignment(horizontal="center", vertical="center")

    thin = Side(border_style="thin", color="B8CCE4")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    col_widths = [12, 12, 20, 10, 12, 12, 10, 20, 22]

    for pod in POD_LIST:
        pod_df = all_data[all_data["POD"] == pod].copy()
        if pod_df.empty:
            continue

        # Sort by ETD
        pod_df["_etd_sort"] = pd.to_datetime(pod_df["ETD"], format="%Y/%m/%d", errors="coerce")
        pod_df = pod_df.sort_values("_etd_sort").drop(columns=["_etd_sort"])
        pod_df = pod_df.reset_index(drop=True)

        # Add carrier column at front
        if "Carrier" in pod_df.columns:
            cols = ["Carrier"] + OUTPUT_COLS
        else:
            cols = OUTPUT_COLS

        ws = wb.create_sheet(title=pod.title())

        # Header row
        ws.row_dimensions[1].height = 30
        for j, col_name in enumerate(OUTPUT_COLS, 1):
            cell = ws.cell(row=1, column=j, value=col_name)
            cell.fill   = header_fill
            cell.font   = header_font
            cell.alignment = header_align
            cell.border = border

        # Data rows
        for i, (_, row) in enumerate(pod_df.iterrows(), 2):
            fill = alt_fill if i % 2 == 0 else None
            for j, col_name in enumerate(OUTPUT_COLS, 1):
                val = row.get(col_name, "")
                cell = ws.cell(row=i, column=j, value=str(val) if val else "")
                cell.font   = normal_font
                cell.alignment = center_align
                cell.border = border
                if fill:
                    cell.fill = fill

        # Column widths
        for j, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w

        # Freeze top row
        ws.freeze_panes = "A2"

    # Summary sheet
    ws_sum = wb.create_sheet(title="Summary", index=0)
    ws_sum.cell(1, 1, "POD").font = Font(bold=True)
    ws_sum.cell(1, 2, "Count").font = Font(bold=True)
    ws_sum.cell(1, 3, "Carriers").font = Font(bold=True)
    for i, pod in enumerate(POD_LIST, 2):
        subset = all_data[all_data["POD"] == pod]
        ws_sum.cell(i, 1, pod)
        ws_sum.cell(i, 2, len(subset))
        carriers = ", ".join(sorted(subset["Carrier"].unique())) if "Carrier" in subset.columns else "—"
        ws_sum.cell(i, 3, carriers)
    ws_sum.column_dimensions["A"].width = 14
    ws_sum.column_dimensions["B"].width = 8
    ws_sum.column_dimensions["C"].width = 30

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="船期整理系統",
        page_icon="🚢",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # Custom CSS
    st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(90deg, #1F4788 0%, #2E6BC9 100%);
        color: white; padding: 16px 24px; border-radius: 8px;
        margin-bottom: 20px;
    }
    .main-header h1 { margin:0; font-size:1.8rem; }
    .main-header p  { margin:4px 0 0; font-size:0.9rem; opacity:0.85; }
    .carrier-badge {
        display:inline-block; background:#1F4788; color:white;
        padding:2px 10px; border-radius:12px; font-size:0.8rem;
        margin:2px;
    }
    .stat-box {
        background:#F0F4FC; border:1px solid #B8CCE4; border-radius:6px;
        padding:12px; text-align:center;
    }
    .stat-num { font-size:1.8rem; font-weight:700; color:#1F4788; }
    .stat-lbl { font-size:0.8rem; color:#666; }
    div[data-testid="stTabs"] button { font-size:0.95rem; padding:8px 16px; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="main-header">
      <h1>🚢 船期整理系統</h1>
      <p>Shipping Schedule Organizer｜POL: HAIPHONG → HKG / SKU / KHH / TXG</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Session state ────────────────────────────────────────────────────
    if "data_store" not in st.session_state:
        st.session_state.data_store = {c: empty_df() for c in CARRIER_LIST}

    def store_data(carrier: str, df: pd.DataFrame):
        if not df.empty:
            df["Carrier"] = carrier
        st.session_state.data_store[carrier] = df

    def get_all_data() -> pd.DataFrame:
        frames = []
        for carrier in CARRIER_LIST:
            df = st.session_state.data_store.get(carrier, empty_df())
            if not df.empty:
                if "Carrier" not in df.columns:
                    df["Carrier"] = carrier
                frames.append(df)
        if frames:
            return pd.concat(frames, ignore_index=True)
        return empty_df()

    # ── Sidebar ──────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ⚙️ 設定")

        # Month & Year selector
        today = datetime.today()
        col_y, col_m = st.columns(2)
        with col_y:
            sel_year  = st.selectbox("年份", list(range(today.year - 1, today.year + 3)),
                                     index=1)
        with col_m:
            sel_month = st.selectbox("月份", list(range(1, 13)),
                                     index=today.month - 1,
                                     format_func=lambda m: f"{m:02d}月")

        # Date range for file upload filtering
        import calendar
        last_day = calendar.monthrange(sel_year, sel_month)[1]
        range_start = datetime(sel_year, sel_month, 1)
        range_end   = datetime(sel_year, sel_month, last_day, 23, 59, 59)

        st.divider()
        st.markdown("### 🎯 目的港 (POD)")
        sel_pods = []
        for pod in POD_LIST:
            if st.checkbox(pod, value=True, key=f"pod_{pod}"):
                sel_pods.append(pod)

        st.divider()
        st.markdown("### 📊 已收集資料")
        for carrier in CARRIER_LIST:
            df = st.session_state.data_store.get(carrier, empty_df())
            cnt = len(df)
            color = "#1F4788" if cnt > 0 else "#aaa"
            st.markdown(f"<span class='carrier-badge' style='background:{color}'>{carrier}: {cnt} 筆</span>",
                        unsafe_allow_html=True)

        st.divider()
        if st.button("🗑️ 清除所有資料", type="secondary", use_container_width=True):
            st.session_state.data_store = {c: empty_df() for c in CARRIER_LIST}
            st.rerun()

    # ── Main tabs ────────────────────────────────────────────────────────
    tab_upload, tab_web, tab_preview, tab_export = st.tabs([
        "📂 上傳檔案 (CNC / TSL)",
        "🌐 網頁爬取 (IAL / KMTC / YML)",
        "📋 資料預覽",
        "📥 匯出 Excel",
    ])

    # ── Tab 1: File Upload ───────────────────────────────────────────────
    with tab_upload:
        st.subheader("📂 上傳船期檔案")
        st.info(f"篩選日期範圍：**{range_start.strftime('%Y/%m/%d')}** 至 **{range_end.strftime('%Y/%m/%d')}**")

        col1, col2 = st.columns(2)

        # CNC Upload
        with col1:
            st.markdown("#### 🚢 CNC")
            st.caption("支援格式：CSV / PDF / Excel｜檔名需含 **CNC**")
            cnc_files = st.file_uploader("選擇 CNC 檔案", type=["csv", "pdf", "xlsx", "xls"],
                                          accept_multiple_files=True, key="cnc_upload",
                                          label_visibility="collapsed")
            if cnc_files:
                if st.button("解析 CNC 檔案", type="primary", key="parse_cnc"):
                    all_cnc = []
                    with st.spinner("解析中..."):
                        for f in cnc_files:
                            df = parse_cnc(f.read(), f.name)
                            if not df.empty:
                                all_cnc.append(df)
                    if all_cnc:
                        merged = pd.concat(all_cnc, ignore_index=True)
                        # Filter to selected PODs
                        merged = merged[merged["POD"].isin(sel_pods)]
                        store_data("CNC", merged)
                        st.success(f"✅ CNC 解析完成：共 {len(merged)} 筆")
                        st.dataframe(merged, use_container_width=True)
                    else:
                        st.warning("⚠️ 未找到符合條件的 CNC 資料")

        # TSL Upload
        with col2:
            st.markdown("#### 🚢 TSL")
            st.caption("支援格式：Excel / PDF｜檔名需含 **TSL**")
            tsl_files = st.file_uploader("選擇 TSL 檔案", type=["xlsx", "xls", "pdf"],
                                          accept_multiple_files=True, key="tsl_upload",
                                          label_visibility="collapsed")
            if tsl_files:
                if st.button("解析 TSL 檔案", type="primary", key="parse_tsl"):
                    all_tsl = []
                    with st.spinner("解析中..."):
                        for f in tsl_files:
                            df = parse_tsl(f.read(), f.name, range_start, range_end)
                            if not df.empty:
                                all_tsl.append(df)
                    if all_tsl:
                        merged = pd.concat(all_tsl, ignore_index=True)
                        merged = merged[merged["POD"].isin(sel_pods)]
                        store_data("TSL", merged)
                        st.success(f"✅ TSL 解析完成：共 {len(merged)} 筆")
                        st.dataframe(merged, use_container_width=True)
                    else:
                        st.warning("⚠️ 未找到符合條件的 TSL 資料")

    # ── Tab 2: Web Scraping ──────────────────────────────────────────────
    with tab_web:
        st.subheader("🌐 網頁爬取船期")
        st.info(f"爬取月份：**{sel_year} 年 {sel_month:02d} 月** ｜ 目的港：{', '.join(sel_pods) if sel_pods else '（未選擇）'}")

        if not sel_pods:
            st.warning("請在側邊欄選擇至少一個目的港")
        else:
            col_ial, col_kmtc, col_yml = st.columns(3)

            # IAL
            with col_ial:
                st.markdown("#### 🌐 IAL")
                st.caption("https://www.interasia.cc")
                ial_cnt = len(st.session_state.data_store.get("IAL", empty_df()))
                if ial_cnt > 0:
                    st.success(f"已有 {ial_cnt} 筆資料")
                if st.button("爬取 IAL", type="primary", key="scrape_ial", use_container_width=True):
                    with st.spinner(f"爬取 IAL {sel_year}/{sel_month:02d}..."):
                        df = scrape_ial(sel_year, sel_month, sel_pods)
                    if not df.empty:
                        store_data("IAL", df)
                        st.success(f"✅ IAL 完成：{len(df)} 筆")
                        st.dataframe(df, use_container_width=True)
                    else:
                        st.warning("⚠️ 未取得 IAL 資料（請確認瀏覽器驅動已安裝）")

            # KMTC
            with col_kmtc:
                st.markdown("#### 🌐 KMTC")
                st.caption("https://www.ekmtc.com")
                kmtc_cnt = len(st.session_state.data_store.get("KMTC", empty_df()))
                if kmtc_cnt > 0:
                    st.success(f"已有 {kmtc_cnt} 筆資料")
                if st.button("爬取 KMTC", type="primary", key="scrape_kmtc", use_container_width=True):
                    with st.spinner(f"爬取 KMTC {sel_year}/{sel_month:02d}..."):
                        df = scrape_kmtc(sel_year, sel_month, sel_pods)
                    if not df.empty:
                        store_data("KMTC", df)
                        st.success(f"✅ KMTC 完成：{len(df)} 筆")
                        st.dataframe(df, use_container_width=True)
                    else:
                        st.warning("⚠️ 未取得 KMTC 資料（請確認瀏覽器驅動已安裝）")

            # YML
            with col_yml:
                st.markdown("#### 🌐 YML")
                st.caption("https://www.yangming.com")
                yml_cnt = len(st.session_state.data_store.get("YML", empty_df()))
                if yml_cnt > 0:
                    st.success(f"已有 {yml_cnt} 筆資料")
                if st.button("爬取 YML", type="primary", key="scrape_yml", use_container_width=True):
                    with st.spinner(f"爬取 YML {sel_year}/{sel_month:02d}..."):
                        df = scrape_yml(sel_year, sel_month, sel_pods)
                    if not df.empty:
                        store_data("YML", df)
                        st.success(f"✅ YML 完成：{len(df)} 筆")
                        st.dataframe(df, use_container_width=True)
                    else:
                        st.warning("⚠️ 未取得 YML 資料（請確認瀏覽器驅動已安裝）")

            st.divider()
            st.markdown("#### 一鍵爬取全部")
            if st.button("🚀 爬取 IAL + KMTC + YML", type="primary", use_container_width=True):
                carriers_web = [("IAL", scrape_ial), ("KMTC", scrape_kmtc), ("YML", scrape_yml)]
                for carrier_name, scrape_fn in carriers_web:
                    with st.spinner(f"爬取 {carrier_name}..."):
                        df = scrape_fn(sel_year, sel_month, sel_pods)
                    if not df.empty:
                        store_data(carrier_name, df)
                        st.success(f"✅ {carrier_name}: {len(df)} 筆")
                    else:
                        st.warning(f"⚠️ {carrier_name}: 無資料")

    # ── Tab 3: Data Preview ──────────────────────────────────────────────
    with tab_preview:
        st.subheader("📋 資料預覽")
        all_data = get_all_data()

        if all_data.empty:
            st.info("尚無資料。請先上傳檔案或爬取網頁船期。")
        else:
            # Summary stats
            st.markdown("##### 📊 彙整統計")
            stat_cols = st.columns(len(POD_LIST) + 1)
            with stat_cols[0]:
                st.markdown(f"""
                <div class="stat-box">
                  <div class="stat-num">{len(all_data)}</div>
                  <div class="stat-lbl">總筆數</div>
                </div>""", unsafe_allow_html=True)
            for i, pod in enumerate(POD_LIST, 1):
                cnt = len(all_data[all_data["POD"] == pod])
                with stat_cols[i]:
                    st.markdown(f"""
                    <div class="stat-box">
                      <div class="stat-num">{cnt}</div>
                      <div class="stat-lbl">{pod}</div>
                    </div>""", unsafe_allow_html=True)

            st.divider()

            # Preview by POD
            view_mode = st.radio("檢視方式", ["依 POD 分頁", "全部合併"], horizontal=True)

            if view_mode == "依 POD 分頁":
                pod_tabs = st.tabs([p.title() for p in POD_LIST])
                for i, pod in enumerate(POD_LIST):
                    with pod_tabs[i]:
                        pod_df = all_data[all_data["POD"] == pod].copy()
                        if pod_df.empty:
                            st.info(f"無 {pod} 資料")
                        else:
                            pod_df["_sort"] = pd.to_datetime(pod_df["ETD"], format="%Y/%m/%d", errors="coerce")
                            pod_df = pod_df.sort_values("_sort").drop(columns=["_sort"])
                            display_cols = ["Carrier"] + OUTPUT_COLS if "Carrier" in pod_df.columns else OUTPUT_COLS
                            st.dataframe(pod_df[display_cols], use_container_width=True, height=400)
                            st.caption(f"共 {len(pod_df)} 筆")
            else:
                all_data["_sort"] = pd.to_datetime(all_data["ETD"], format="%Y/%m/%d", errors="coerce")
                all_data = all_data.sort_values(["POD", "_sort"]).drop(columns=["_sort"])
                display_cols = ["Carrier"] + OUTPUT_COLS if "Carrier" in all_data.columns else OUTPUT_COLS
                st.dataframe(all_data[display_cols], use_container_width=True, height=500)

    # ── Tab 4: Export ────────────────────────────────────────────────────
    with tab_export:
        st.subheader("📥 匯出 Excel")
        all_data = get_all_data()

        if all_data.empty:
            st.warning("⚠️ 尚無資料可匯出，請先收集船期資料。")
        else:
            st.markdown(f"""
            **匯出摘要：**
            - 總筆數：**{len(all_data)}** 筆
            - 船公司：{', '.join(sorted(all_data['Carrier'].unique()) if 'Carrier' in all_data.columns else [])}
            - 目的港 Sheets：{', '.join([p for p in POD_LIST if p in all_data['POD'].values])}
            """)

            filename = f"船期整理_{sel_year}{sel_month:02d}.xlsx"
            if st.button("🔄 生成 Excel 報表", type="primary"):
                with st.spinner("生成中..."):
                    excel_bytes = export_to_excel(all_data)
                st.success("✅ Excel 報表生成完畢！")
                st.download_button(
                    label=f"⬇️ 下載 {filename}",
                    data=excel_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )

            st.divider()
            st.markdown("#### 📑 預覽各 Sheet")
            for pod in POD_LIST:
                pod_df = all_data[all_data["POD"] == pod]
                if not pod_df.empty:
                    with st.expander(f"📄 {pod.title()} ({len(pod_df)} 筆)"):
                        pod_df2 = pod_df.copy()
                        pod_df2["_s"] = pd.to_datetime(pod_df2["ETD"], format="%Y/%m/%d", errors="coerce")
                        pod_df2 = pod_df2.sort_values("_s").drop(columns=["_s"])
                        display_cols = ["Carrier"] + OUTPUT_COLS if "Carrier" in pod_df2.columns else OUTPUT_COLS
                        st.dataframe(pod_df2[display_cols], use_container_width=True)


if __name__ == "__main__":
    main()
