"""
Shipping Schedule Organizer v4.0
Supports: PDF / Excel / PNG / JPG upload
Exports: Excel with sheets per Carrier-POD-Month (e.g. CNC - KHH - MARCH)
"""

import streamlit as st
import pandas as pd
import pdfplumber
import json, re, base64, io, os
from datetime import datetime
from anthropic import Anthropic
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Shipping Schedule Organizer",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.upload-hint { color:#aaa; font-size:0.85rem; margin-top:4px; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# API KEY
# ══════════════════════════════════════════════════════════════
try:
    API_KEY = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

# ══════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════
COLUMNS = ["CARRIER","POL","POD","Vessel","Voyage","ETD","ETA","T/T Time","CY Cut-off","SI Cut-off"]

CARRIERS = ["CNC","IAL","KMTC","SITC","TSL","YML","COSCO","EVERGREEN","ONE","PIL","RCL","WHL","OTHER"]

COMMON_PORTS = [
    "HAIPHONG","HO CHI MINH CITY","DA NANG",
    "HONG KONG","SHEKOU","NANSHA","GUANGZHOU",
    "KAOHSIUNG","TAICHUNG","KEELUNG",
    "SHANGHAI","NINGBO","QINGDAO","TIANJIN",
    "YANGON","PORT KLANG","SINGAPORE","BANGKOK","LAEM CHABANG",
    "JAKARTA","SURABAYA","COLOMBO","CHATTOGRAM","BUSAN","TOKYO",
]

PORT_CODE = {
    "HAIPHONG":"HPH","HO CHI MINH CITY":"SGN","DA NANG":"DAD",
    "HONG KONG":"HKG","SHEKOU":"SKU","NANSHA":"NSA","GUANGZHOU":"CAN",
    "KAOHSIUNG":"KHH","TAICHUNG":"TXG","KEELUNG":"KEL",
    "SHANGHAI":"SHA","NINGBO":"NGB","QINGDAO":"TAO","TIANJIN":"TSN",
    "YANGON":"RGN","PORT KLANG":"PKG","SINGAPORE":"SIN",
    "BANGKOK":"BKK","LAEM CHABANG":"LCB",
    "JAKARTA":"JKT","SURABAYA":"SUB","COLOMBO":"CMB",
    "CHATTOGRAM":"CGP","BUSAN":"PUS","TOKYO":"TYO",
}

MONTHS = {1:"JAN",2:"FEB",3:"MAR",4:"APR",5:"MAY",6:"JUN",
          7:"JUL",8:"AUG",9:"SEP",10:"OCT",11:"NOV",12:"DEC"}

# ══════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════════
def get_port_code(name: str) -> str:
    """Convert port name to 3-letter code."""
    n = name.upper().strip()
    if n in PORT_CODE: 
        return PORT_CODE[n]
    for k, v in PORT_CODE.items():
        if k in n or n in k: 
            return v
    return n[:3].upper()

def get_month_from_etd(etd: str) -> int:
    """Extract month number from ETD like '02-06' or '2026-02-06'."""
    if not etd: 
        return datetime.now().month
    m = re.search(r'(\d{2})-\d{2}$', etd.strip())
    if m: 
        return int(m.group(1))
    m = re.search(r'\d{4}-(\d{2})-\d{2}', etd.strip())
    if m: 
        return int(m.group(1))
    return datetime.now().month

def make_sheet_name(carrier: str, pod: str, month_num: int) -> str:
    """Generate sheet name like 'CNC - KHH - MAR' (max 31 chars)."""
    pod_code = get_port_code(pod)
    mon = MONTHS.get(month_num, "???")
    return f"{carrier[:6]} - {pod_code} - {mon}"[:31]

def norm_row(r: dict) -> dict:
    """Normalize a row dict to standard COLUMNS."""
    def g(*keys):
        for k in keys:
            v = r.get(k, "")
            if v: 
                return str(v).strip()
        return ""
    
    return {
        "CARRIER":    g("CARRIER","carrier"),
        "POL":        g("POL","pol","origin"),
        "POD":        g("POD","pod","destination"),
        "Vessel":     g("Vessel","vessel","VESSEL","ship"),
        "Voyage":     g("Voyage","voyage","VOYAGE","voy"),
        "ETD":        g("ETD","etd","departure"),
        "ETA":        g("ETA","eta","arrival"),
        "T/T Time":   g("T/T Time","transit_time","T/T","transit"),
        "CY Cut-off": g("CY Cut-off","cy_cutoff","CY","cy"),
        "SI Cut-off": g("SI Cut-off","si_cutoff","SI","si","doc_cutoff"),
    }

# ══════════════════════════════════════════════════════════════
# CLAUDE PARSING - TEXT
# ══════════════════════════════════════════════════════════════
def parse_text_claude(text: str, carrier: str, pol: str, pod: str, api_key: str) -> list:
    """Parse text schedule using Claude AI."""
    client = Anthropic(api_key=api_key)
    
    prompt = f"""Extract ALL shipping schedule entries from the text below.

Carrier: {carrier}
POL (Port of Loading): {pol}
POD (Port of Discharge): {pod}

RULES:
1. ETD = departure date from {pol} → format MM-DD (e.g. 02-15)
2. ETA = arrival date at {pod} → format MM-DD
3. T/T = transit days as integer (round UP: 1d 15h → 2)
4. CY Cut-off = CY CUT / Cargo Closing / CY Closing
5. SI Cut-off = Document Closing / Doc Cut / S/I Cut / Booking Doc Closing
   (KMTC: "Document Closing" = SI Cut-off)
6. All dates format: MM-DD
7. Extract EVERY sailing visible. Do not skip.

Return ONLY a valid JSON array, no explanation:
[
  {{
    "carrier": "{carrier}",
    "pol": "{pol}",
    "pod": "{pod}",
    "vessel": "VESSEL NAME",
    "voyage": "VOYAGE NUMBER",
    "etd": "MM-DD",
    "eta": "MM-DD",
    "transit_time": "2",
    "cy_cutoff": "MM-DD",
    "si_cutoff": "MM-DD"
  }}
]

TEXT TO PARSE:
{text[:6000]}"""

    try:
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=2000,
            messages=[{"role":"user","content":prompt}]
        )
        raw = resp.content[0].text.strip()
        # Remove markdown code blocks
        raw = re.sub(r'^```json\s*|^```\s*|```$','', raw, flags=re.MULTILINE).strip()
        data = json.loads(raw)
        return [norm_row(r) for r in data if r.get("etd")]
    except Exception as e:
        st.error(f"Claude parsing error: {e}")
        return []

# ══════════════════════════════════════════════════════════════
# CLAUDE PARSING - IMAGE
# ══════════════════════════════════════════════════════════════
def parse_image_claude(img_bytes: bytes, ext: str, carrier: str, pol: str, pod: str, api_key: str) -> list:
    """Parse schedule image using Claude Vision."""
    client = Anthropic(api_key=api_key)
    b64 = base64.standard_b64encode(img_bytes).decode()
    media_type = "image/png" if ext == "png" else "image/jpeg"

    prompt = f"""This image shows a shipping schedule from {carrier}.

POL: {pol}  →  POD: {pod}

Extract ALL sailing rows visible in the image.

RULES:
1. ETD from {pol} → format MM-DD
2. ETA at {pod} → format MM-DD
3. T/T = integer days (round UP)
4. CY Cut-off = CY CUT / Cargo Closing
5. SI Cut-off = Document Closing / Doc Cut / Booking Cut
6. Extract EVERY row visible.

Return ONLY a JSON array:
[
  {{
    "carrier": "{carrier}",
    "pol": "{pol}",
    "pod": "{pod}",
    "vessel": "VESSEL NAME",
    "voyage": "VOYAGE",
    "etd": "MM-DD",
    "eta": "MM-DD",
    "transit_time": "2",
    "cy_cutoff": "MM-DD",
    "si_cutoff": "MM-DD"
  }}
]"""

    try:
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=2000,
            messages=[{
                "role": "user",
                "content": [
                    {"type":"image","source":{"type":"base64","media_type":media_type,"data":b64}},
                    {"type":"text","text":prompt}
                ]
            }]
        )
        raw = resp.content[0].text.strip()
        raw = re.sub(r'^```json\s*|^```\s*|```$','', raw, flags=re.MULTILINE).strip()
        data = json.loads(raw)
        return [norm_row(r) for r in data if r.get("etd")]
    except Exception as e:
        st.error(f"Claude Vision error: {e}")
        return []

# ══════════════════════════════════════════════════════════════
# EXCEL FILE PARSER
# ══════════════════════════════════════════════════════════════
def parse_excel_upload(fbytes: bytes, carrier: str, pol: str, pod: str) -> list:
    """Parse uploaded Excel/CSV file."""
    try:
        df = pd.read_excel(io.BytesIO(fbytes))
    except Exception:
        try:
            df = pd.read_csv(io.BytesIO(fbytes))
        except Exception as e:
            st.error(f"Excel/CSV read error: {e}")
            return []

    # Auto-detect columns
    col_map = {}
    for col in df.columns:
        c = str(col).upper().strip()
        if any(x in c for x in ["VESSEL","SHIP","VSL"]):    
            col_map["Vessel"] = col
        elif any(x in c for x in ["VOY","VOYAGE"]):          
            col_map["Voyage"] = col
        elif "ETD" in c or "DEPARTURE" in c:                 
            col_map["ETD"] = col
        elif "ETA" in c or "ARRIVAL" in c:                   
            col_map["ETA"] = col
        elif "T/T" in c or "TRANSIT" in c:                   
            col_map["T/T Time"] = col
        elif "CY" in c and "CUT" in c:                       
            col_map["CY Cut-off"] = col
        elif any(x in c for x in ["SI","DOC","S/I"]):        
            col_map["SI Cut-off"] = col
        elif "POL" in c or "ORIGIN" in c:                    
            col_map["POL"] = col
        elif "POD" in c or "DEST" in c:                      
            col_map["POD"] = col

    rows = []
    for _, row in df.iterrows():
        vessel = str(row.get(col_map.get("Vessel",""), "") or "").strip()
        if not vessel or vessel == "nan": 
            continue
        
        rows.append({
            "CARRIER":    carrier,
            "POL":        str(row.get(col_map.get("POL",""), pol) or pol).strip(),
            "POD":        str(row.get(col_map.get("POD",""), pod) or pod).strip(),
            "Vessel":     vessel,
            "Voyage":     str(row.get(col_map.get("Voyage",""), "") or "").strip(),
            "ETD":        str(row.get(col_map.get("ETD",""), "") or "").strip(),
            "ETA":        str(row.get(col_map.get("ETA",""), "") or "").strip(),
            "T/T Time":   str(row.get(col_map.get("T/T Time",""), "") or "").strip(),
            "CY Cut-off": str(row.get(col_map.get("CY Cut-off",""), "") or "").strip(),
            "SI Cut-off": str(row.get(col_map.get("SI Cut-off",""), "") or "").strip(),
        })
    return rows

# ══════════════════════════════════════════════════════════════
# PDF FILE PARSER
# ══════════════════════════════════════════════════════════════
def parse_pdf_upload(fbytes: bytes, carrier: str, pol: str, pod: str, api_key: str) -> list:
    """Parse uploaded PDF file."""
    text = ""
    try:
        with pdfplumber.open(io.BytesIO(fbytes)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t: 
                    text += t + "\n"
    except Exception as e:
        st.error(f"PDF read error: {e}")
        return []
    
    if not text.strip(): 
        return []
    
    return parse_text_claude(text, carrier, pol, pod, api_key)

# ══════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ══════════════════════════════════════════════════════════════
def _write_worksheet(ws, df: pd.DataFrame):
    """Write header + data rows to worksheet with styling."""
    WIDTHS = {
        "CARRIER":10, "POL":14, "POD":14, "Vessel":22, "Voyage":9,
        "ETD":8, "ETA":8, "T/T Time":7, "CY Cut-off":11, "SI Cut-off":11
    }
    
    # Header styling
    h_fill = PatternFill("solid", fgColor="1F4E79")
    h_font = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
    h_alig = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin   = Side(style="thin", color="AAAAAA")
    bord   = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    # Data row styling
    e_fill = PatternFill("solid", fgColor="EBF3FB")
    
    # Write headers
    for ci, col in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.fill = h_fill
        cell.font = h_font
        cell.alignment = h_alig
        cell.border = bord
        ws.column_dimensions[get_column_letter(ci)].width = WIDTHS.get(col, 12)
    
    ws.row_dimensions[1].height = 28
    
    # Write data rows
    for ri, (_, row) in enumerate(df.iterrows(), 2):
        for ci, col in enumerate(COLUMNS, 1):
            val = str(row.get(col, "") or "").strip()
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = bord
            if ri % 2 == 0:  # Even rows get blue tint
                cell.fill = e_fill
    
    ws.freeze_panes = "A2"

def create_excel(df: pd.DataFrame) -> bytes:
    """Create Excel workbook with multiple sheets."""
    wb = Workbook()
    
    # Sheet 1: All Schedules
    ws_all = wb.active
    ws_all.title = "All Schedules"
    df_sorted = df.sort_values("ETD", na_position="last").reset_index(drop=True)
    _write_worksheet(ws_all, df_sorted)
    
    # Create sheets per Carrier + POD + Month
    bucket = {}
    for _, row in df_sorted.iterrows():
        c = str(row.get("CARRIER", "")).strip().upper() or "UNK"
        p = str(row.get("POD", "")).strip().upper() or "UNK"
        m = get_month_from_etd(str(row.get("ETD", "")))
        key = make_sheet_name(c, p, m)
        bucket.setdefault(key, []).append(row)
    
    # Create each sheet
    for sname in sorted(bucket.keys()):
        ws = wb.create_sheet(title=sname)
        sub_df = pd.DataFrame(bucket[sname], columns=COLUMNS).fillna("")
        sub_df = sub_df.sort_values("ETD", na_position="last").reset_index(drop=True)
        _write_worksheet(ws, sub_df)
    
    # Save to bytes
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame(columns=COLUMNS)

def add_rows(rows: list):
    """Add new rows to session state dataframe."""
    if not rows: 
        return
    
    new = pd.DataFrame(rows, columns=COLUMNS).fillna("")
    merged = pd.concat([st.session_state.df, new], ignore_index=True)
    
    # Remove duplicates
    merged = merged.drop_duplicates(
        subset=["CARRIER","POL","POD","Vessel","Voyage","ETD"], 
        keep="last"
    )
    
    # Sort by ETD
    merged = merged.sort_values("ETD", na_position="last").reset_index(drop=True)
    st.session_state.df = merged

# ══════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.title("🚢 Schedule Organizer")
    st.markdown("---")
    
    st.markdown("#### ⚙️ File Settings")
    sel_carrier = st.selectbox("Carrier", CARRIERS)
    sel_pol = st.selectbox("POL (Origin)", COMMON_PORTS,
                            index=COMMON_PORTS.index("HAIPHONG"))
    sel_pod = st.selectbox("POD (Destination)", COMMON_PORTS,
                            index=COMMON_PORTS.index("HONG KONG"))
    
    st.markdown("---")
    
    # API Key status
    if API_KEY:
        st.success("✅ API Key configured")
    else:
        st.error("❌ No API Key")
        st.caption("PDF/Image requires API Key")
        st.caption("Add to Streamlit Secrets:")
        st.code('ANTHROPIC_API_KEY = "sk-ant-..."', language="toml")
    
    st.markdown("---")
    
    # Stats
    df_side = st.session_state.df
    st.metric("Total Schedules", len(df_side))
    st.metric("Destinations", df_side["POD"].nunique() if len(df_side) else 0)
    
    # Clear button
    if len(df_side) > 0:
        if st.button("🗑️ Clear All Data", use_container_width=True):
            st.session_state.df = pd.DataFrame(columns=COLUMNS)
            st.rerun()

# ══════════════════════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════════════════════
st.title("🚢 Shipping Schedule Organizer")
st.caption("Upload PDF / Excel / Image → Preview → Export Excel (sheets: CNC - KHH - MAR)")

tab1, tab2, tab3 = st.tabs(["📁 Upload Files", "✏️ Preview & Edit", "📊 Export Excel"])

# ──────────────────────────────────────────────────────────────
# TAB 1: UPLOAD FILES
# ──────────────────────────────────────────────────────────────
with tab1:
    st.markdown("### 📁 Upload Schedule Files")
    
    st.markdown("""
**Supported formats:**

| Type | How it works | API needed? | Cost |
|------|--------------|-------------|------|
| 📊 **Excel / CSV** | Auto-detect columns | ❌ No | Free |
| 📄 **PDF** | Claude AI reads text | ✅ Yes | ~$0.01/file |
| 🖼️ **PNG / JPG** | Claude Vision reads screenshot | ✅ Yes | ~$0.03/image |

💡 **Tip:** For best results with images, take clear screenshots of the schedule table.
""")
    
    if not API_KEY:
        st.warning("⚠️ PDF and Image files require an Anthropic API Key. Excel files work without API Key.")
    
    # File uploader
    uploaded = st.file_uploader(
        "📎 Drop files here or click to browse",
        type=["pdf","xlsx","xls","csv","png","jpg","jpeg"],
        accept_multiple_files=True,
        help="You can upload multiple files at once"
    )
    
    # Parse button
    _, col_btn, _ = st.columns([2,1,2])
    with col_btn:
        parse_btn = st.button(
            "🚀 Parse All Files",
            type="primary",
            use_container_width=True,
            disabled=not uploaded
        )
    
    # Process uploaded files
    if parse_btn and uploaded:
        total_added = 0
        
        for f in uploaded:
            fname = f.name
            ext = fname.rsplit(".", 1)[-1].lower()
            fbytes = f.read()
            
            with st.spinner(f"Processing **{fname}**..."):
                try:
                    # Route to appropriate parser
                    if ext in ("xlsx", "xls", "csv"):
                        rows = parse_excel_upload(fbytes, sel_carrier, sel_pol, sel_pod)
                        source = "Excel"
                        
                    elif ext == "pdf":
                        if not API_KEY:
                            st.warning(f"⚠️ {fname}: API Key required for PDF. Skipped.")
                            continue
                        rows = parse_pdf_upload(fbytes, sel_carrier, sel_pol, sel_pod, API_KEY)
                        source = "PDF→Claude"
                        
                    elif ext in ("png", "jpg", "jpeg"):
                        if not API_KEY:
                            st.warning(f"⚠️ {fname}: API Key required for images. Skipped.")
                            continue
                        rows = parse_image_claude(fbytes, ext, sel_carrier, sel_pol, sel_pod, API_KEY)
                        source = "Image→Claude"
                        
                    else:
                        st.warning(f"⚠️ {fname}: Unsupported format")
                        continue
                    
                    # Add rows to dataframe
                    if rows:
                        add_rows(rows)
                        total_added += len(rows)
                        st.success(f"✅ **{fname}** ({source}) → **{len(rows)}** schedules extracted")
                        
                        # Show preview
                        with st.expander(f"Preview extracted data from {fname}"):
                            preview_df = pd.DataFrame(rows, columns=COLUMNS)
                            st.dataframe(preview_df, use_container_width=True, hide_index=True)
                    else:
                        st.warning(f"⚠️ {fname}: No schedules found")
                        
                except Exception as e:
                    st.error(f"❌ {fname}: {str(e)}")
        
        # Summary
        if total_added > 0:
            st.balloons()
            st.success(f"🎉 Successfully added **{total_added}** schedules! Go to **Preview & Edit** tab to review.")

# ──────────────────────────────────────────────────────────────
# TAB 2: PREVIEW & EDIT
# ──────────────────────────────────────────────────────────────
with tab2:
    st.markdown("### ✏️ Preview & Edit Schedules")
    df = st.session_state.df
    
    if df.empty:
        st.info("📭 No data yet. Upload files in the **Upload Files** tab.")
    else:
        # Summary metrics
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Rows", len(df))
        c2.metric("Carriers", df["CARRIER"].nunique())
        c3.metric("PODs", df["POD"].nunique())
        
        # Count expected sheets
        sheet_names = set()
        for _, row in df.iterrows():
            c = str(row.get("CARRIER", "")).upper()
            p = str(row.get("POD", "")).upper()
            m = get_month_from_etd(str(row.get("ETD", "")))
            sheet_names.add(make_sheet_name(c, p, m))
        c4.metric("Excel Sheets", len(sheet_names) + 1)  # +1 for "All Schedules"
        
        st.markdown("---")
        
        # Filters
        fc1, fc2 = st.columns(2)
        with fc1:
            carriers_list = ["All"] + sorted(df["CARRIER"].unique().tolist())
            flt_carrier = st.selectbox("Filter by Carrier", carriers_list)
        with fc2:
            pods_list = ["All"] + sorted(df["POD"].unique().tolist())
            flt_pod = st.selectbox("Filter by POD", pods_list)
        
        # Apply filters
        view = df.copy()
        if flt_carrier != "All":
            view = view[view["CARRIER"] == flt_carrier]
        if flt_pod != "All":
            view = view[view["POD"] == flt_pod]
        
        st.caption(f"Showing {len(view)} of {len(df)} rows")
        
        # Editable table
        edited = st.data_editor(
            view[COLUMNS],
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "CARRIER": st.column_config.SelectboxColumn("CARRIER", options=CARRIERS),
            }
        )
        
        # Save button
        if st.button("💾 Save Changes", type="primary"):
            if flt_carrier == "All" and flt_pod == "All":
                st.session_state.df = edited.reset_index(drop=True)
            else:
                st.session_state.df.loc[view.index] = edited.values
            st.success("✅ Changes saved!")
            st.rerun()

# ──────────────────────────────────────────────────────────────
# TAB 3: EXPORT EXCEL
# ──────────────────────────────────────────────────────────────
with tab3:
    st.markdown("### 📊 Export to Excel")
    df = st.session_state.df
    
    if df.empty:
        st.info("📭 No data to export. Upload files first.")
    else:
        # Preview sheet names
        sheet_info = {}
        for _, row in df.sort_values("ETD", na_position="last").iterrows():
            c = str(row.get("CARRIER", "")).upper()
            p = str(row.get("POD", "")).upper()
            m = get_month_from_etd(str(row.get("ETD", "")))
            sn = make_sheet_name(c, p, m)
            sheet_info[sn] = sheet_info.get(sn, 0) + 1
        
        st.markdown("**📋 Sheets that will be created:**")
        all_sheets = ["📄 All Schedules"] + [f"🗂️ {k} ({v} rows)" for k, v in sorted(sheet_info.items())]
        
        # Display in columns
        for i in range(0, len(all_sheets), 3):
            cols = st.columns(3)
            for j, sheet in enumerate(all_sheets[i:i+3]):
                cols[j].markdown(sheet)
        
        st.markdown("---")
        
        # Filename input
        today = datetime.now().strftime("%Y%m%d")
        default_fname = f"Shipping_Schedule_{today}.xlsx"
        fname = st.text_input("📝 Export filename", value=default_fname)
        if not fname.endswith(".xlsx"):
            fname += ".xlsx"
        
        # Build & Download buttons
        b1, b2, _ = st.columns([1, 1, 3])
        
        with b1:
            if st.button("⚙️ Build Excel", type="primary", use_container_width=True):
                with st.spinner("Building Excel file..."):
                    excel_bytes = create_excel(df)
                st.session_state["excel_bytes"] = excel_bytes
                st.session_state["excel_fname"] = fname
                st.success("✅ Excel file ready!")
        
        with b2:
            if "excel_bytes" in st.session_state:
                st.download_button(
                    label="⬇️ Download",
                    data=st.session_state["excel_bytes"],
                    file_name=st.session_state.get("excel_fname", fname),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
