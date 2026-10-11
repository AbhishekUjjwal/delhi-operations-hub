import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import fitz  # PyMuPDF
import re
import io
import json
import os
import zipfile
import barcode
from barcode.writer import ImageWriter
from datetime import datetime, timezone, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

IST = timezone(timedelta(hours=5, minutes=30))
def get_current_ist_str():
    return datetime.now(IST).strftime("%d-%b %I:%M %p")

STORAGE_DIR = ".pan_history"
METADATA_FILE = os.path.join(STORAGE_DIR, "history_db.json")
RETENTION_DAYS = 30

os.makedirs(STORAGE_DIR, exist_ok=True)

def load_disk_history():
    if not os.path.exists(METADATA_FILE):
        return []
    try:
        with open(METADATA_FILE, "r", encoding="utf-8") as f:
            meta_list = json.load(f)

        now_ts = datetime.now(IST).timestamp()
        cutoff_ts = now_ts - (RETENTION_DAYS * 24 * 60 * 60)

        valid_items = []
        updated_meta = []
        cleaned_up = False

        for m in meta_list:
            created_ts = m.get("created_ts", now_ts)
            if created_ts < cutoff_ts:
                cleaned_up = True
                for f_key in ["disk_filename", "disk_excel_filename", "disk_json_filename"]:
                    fname = m.get(f_key)
                    if fname:
                        fpath = os.path.join(STORAGE_DIR, fname)
                        if os.path.exists(fpath):
                            try:
                                os.remove(fpath)
                            except Exception:
                                pass
                continue

            pdf_path = os.path.join(STORAGE_DIR, m.get("disk_filename", ""))
            if os.path.exists(pdf_path):
                m_copy = dict(m)
                with open(pdf_path, "rb") as f_bin:
                    m_copy["data"] = f_bin.read()

                excel_fname = m.get("disk_excel_filename")
                if excel_fname:
                    epath = os.path.join(STORAGE_DIR, excel_fname)
                    if os.path.exists(epath):
                        with open(epath, "rb") as fe:
                            m_copy["excel_data"] = fe.read()

                json_fname = m.get("disk_json_filename")
                if json_fname:
                    jpath = os.path.join(STORAGE_DIR, json_fname)
                    if os.path.exists(jpath):
                        with open(jpath, "rb") as fj:
                            m_copy["json_data"] = fj.read()

                valid_items.append(m_copy)
                updated_meta.append(m)
            else:
                cleaned_up = True

        if cleaned_up:
            with open(METADATA_FILE, "w", encoding="utf-8") as f:
                json.dump(updated_meta, f, indent=2)

        return valid_items
    except Exception:
        return []

def save_item_to_disk_history(item):
    base_name = item.get("file_name", "invoice.pdf")
    wh = item.get("location", "DEL")
    mod = item.get("module", "GEN")

    existing_items = load_disk_history()
    existing_fnames = [
        x.get("file_name") for x in existing_items
        if x.get("location") == wh and x.get("module") == mod
    ]

    new_file_name = base_name
    count = 1
    name_root = base_name[:-4] if base_name.endswith(".pdf") else base_name
    ext = ".pdf" if base_name.endswith(".pdf") else ""

    while new_file_name in existing_fnames:
        new_file_name = f"{name_root}_edit_{count}{ext}"
        count += 1

    current_dt = datetime.now(IST)
    current_ts = current_dt.timestamp()
    prefix = int(current_ts * 1000)

    disk_fname = f"{prefix}_{new_file_name}"
    disk_path = os.path.join(STORAGE_DIR, disk_fname)
    with open(disk_path, "wb") as f_out:
        f_out.write(item["data"])

    disk_excel_fname = None
    if "excel_data" in item and item["excel_data"]:
        disk_excel_fname = f"{prefix}_{name_root}.xlsx"
        with open(os.path.join(STORAGE_DIR, disk_excel_fname), "wb") as fe:
            fe.write(item["excel_data"])

    disk_json_fname = None
    if "json_data" in item and item["json_data"]:
        disk_json_fname = f"{prefix}_{name_root}.json"
        with open(os.path.join(STORAGE_DIR, disk_json_fname), "wb") as fj:
            fj.write(item["json_data"])

    meta_record = {
        "original_name": item.get("original_name", base_name),
        "file_name": new_file_name,
        "disk_filename": disk_fname,
        "disk_excel_filename": disk_excel_fname,
        "disk_json_filename": disk_json_fname,
        "pages": item.get("pages", 1),
        "stamped": item.get("stamped", 0),
        "location": wh,
        "module": mod,
        "created_ts": current_ts,
        "time": get_current_ist_str()
    }

    current_meta = []
    if os.path.exists(METADATA_FILE):
        try:
            with open(METADATA_FILE, "r", encoding="utf-8") as f:
                current_meta = json.load(f)
        except Exception:
            current_meta = []

    current_meta.append(meta_record)
    with open(METADATA_FILE, "w", encoding="utf-8") as f:
        json.dump(current_meta, f, indent=2)

def save_batch_to_disk_history(batch_items):
    for item in batch_items:
        save_item_to_disk_history(item)

def clear_disk_history(wh_filter=None, mod_filter=None):
    if not os.path.exists(METADATA_FILE):
        return
    try:
        with open(METADATA_FILE, "r", encoding="utf-8") as f:
            meta_list = json.load(f)

        remaining_meta = []
        for m in meta_list:
            should_delete = True
            if mod_filter and m.get("module") != mod_filter:
                should_delete = False
            if wh_filter and wh_filter != "Headoffice" and m.get("location") != wh_filter:
                should_delete = False

            if should_delete:
                for f_key in ["disk_filename", "disk_excel_filename", "disk_json_filename"]:
                    fname = m.get(f_key)
                    if fname:
                        fpath = os.path.join(STORAGE_DIR, fname)
                        if os.path.exists(fpath):
                            try:
                                os.remove(fpath)
                            except Exception:
                                pass
            else:
                remaining_meta.append(m)

        with open(METADATA_FILE, "w", encoding="utf-8") as f:
            json.dump(remaining_meta, f, indent=2)
    except Exception:
        pass

st.set_page_config(
    page_title="PAN INDIA OPERATIONS HUB | Central Portal",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    #MainMenu {visibility: hidden; display: none !important;}
    header {visibility: hidden; display: none !important;}
    footer {visibility: hidden; display: none !important;}
    .stDeployButton {display: none !important;}
    div[data-testid="stToolbar"] {visibility: hidden; display: none !important;}
    div[data-testid="stDecoration"] {visibility: hidden; display: none !important;}
    div[data-testid="stStatusWidget"] {visibility: hidden; display: none !important;}

    [data-testid="stViewerBadge"] {display: none !important; visibility: hidden !important;}
    div[class*="viewerBadge"] {display: none !important; visibility: hidden !important;}
    iframe[title*="streamlit"] {display: none !important;}
    div[data-testid="manage-app-button"] {display: none !important;}
    div[class*="manage-app"] {display: none !important;}
    div[class*="floating"] {display: none !important;}

    .block-container {
        padding-top: 0.4rem !important;
        padding-bottom: 0.8rem !important;
        padding-left: 1.2rem !important;
        padding-right: 1.2rem !important;
    }

    hr {
        margin-top: 6px !important;
        margin-bottom: 8px !important;
        border-color: rgba(58, 80, 107, 0.3) !important;
    }

    /* Small action buttons */
    .right-action-cluster div[data-testid="stButton"] button {
        border-radius: 16px !important;
        font-size: 12px !important;
        font-weight: 600 !important;
        min-height: 28px !important;
        height: 28px !important;
        line-height: 1 !important;
        width: 100% !important;
    }

    div[data-testid="stSelectbox"] label { display: none !important; }
    div[data-testid="stSelectbox"] { width: 100% !important; }
    div[data-testid="stSelectbox"] > div > div {
        min-height: 28px !important;
        height: 28px !important;
        padding: 0px 8px !important;
        font-size: 12px !important;
        border-radius: 16px !important;
        border: 1px solid rgba(148, 163, 184, 0.3) !important;
    }

    .pan-brand-title {
        font-size: 22px;
        font-weight: 900;
        letter-spacing: -0.3px;
        background: linear-gradient(90deg, #38bdf8 0%, #818cf8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin: 0px 0px 1px 0px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .pan-brand-subtitle {
        font-size: 12.5px;
        font-weight: 500;
        color: #94a3b8;
        margin-bottom: 8px;
    }

    .pro-success-card {
        background: linear-gradient(135deg, rgba(16, 185, 129, 0.15) 0%, rgba(5, 150, 105, 0.25) 100%);
        border: 1px solid rgba(52, 211, 153, 0.5);
        border-radius: 10px;
        padding: 12px 16px;
        color: #ecfdf5;
        margin: 8px 0px;
    }

    /* Pinned Tabs Styling */
    .top-tab-strip div[data-testid="stButton"] button {
        border-radius: 18px !important;
        font-size: 12px !important;
        font-weight: 700 !important;
        padding: 4px 14px !important;
        min-height: 32px !important;
        height: 32px !important;
    }

    /* OVAL CARD BOX STYLING */
    .oval-card-wrapper {
        position: relative;
        width: 100%;
        margin-bottom: 16px;
    }
    .oval-card {
        background: #ffffff;
        border: 2px solid #e2e8f0;
        border-radius: 24px;
        padding: 24px 26px 18px 26px;
        box-shadow: 0 4px 18px rgba(15, 23, 42, 0.05);
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        pointer-events: none;
    }
    .card-top-flex {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 18px;
    }
    .icon-bubble {
        width: 48px;
        height: 48px;
        border-radius: 14px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 24px;
    }
    .badge-pill {
        font-size: 11px;
        font-weight: 800;
        padding: 5px 14px;
        border-radius: 30px;
        letter-spacing: 0.6px;
    }
    .card-heading {
        font-size: 21px;
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 22px;
        letter-spacing: -0.2px;
    }
    .card-foot-flex {
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-top: 1px solid #f1f5f9;
        padding-top: 14px;
    }
    .foot-label {
        font-size: 14.5px;
        font-weight: 800;
    }
    .foot-circle {
        width: 32px;
        height: 32px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 14px;
        font-weight: 900;
    }

    /* FULL INVISIBLE BUTTON OVERLAY TO MAKE ENTIRE OVAL CLICKABLE */
    .oval-card-wrapper div[data-testid="stButton"] {
        position: absolute !important;
        top: 0 !important;
        left: 0 !important;
        width: 100% !important;
        height: 100% !important;
        z-index: 5 !important;
        margin: 0 !important;
        padding: 0 !important;
    }
    .oval-card-wrapper div[data-testid="stButton"] button {
        width: 100% !important;
        height: 100% !important;
        opacity: 0 !important;
        cursor: pointer !important;
        border-radius: 24px !important;
        border: none !important;
        padding: 0 !important;
        margin: 0 !important;
    }

    /* DYNAMIC HOVER COLOR EFFECTS */
    .oval-card-wrapper:hover .oval-card {
        transform: translateY(-5px);
        box-shadow: 0 12px 28px rgba(15, 23, 42, 0.12);
    }
    .hover-amz:hover .oval-card { border-color: #2563eb !important; }
    .hover-blnk-po:hover .oval-card { border-color: #059669 !important; }
    .hover-blnk:hover .oval-card { border-color: #d97706 !important; }
    .hover-lbl:hover .oval-card { border-color: #7c3aed !important; }
    .hover-conf:hover .oval-card { border-color: #475569 !important; }
    </style>
""", unsafe_allow_html=True)

LOCATION_DISPLAY_MAP = {
    "DEL": "📍 DEL",
    "KOL": "📍 KOL",
    "HYD": "📍 HYD",
    "PUN": "📍 PUN",
    "BLR": "📍 BLR",
    "CHE": "📍 CHE",
    "Headoffice": "🏢 Headoffice"
}

ORDERED_LOCATIONS = list(LOCATION_DISPLAY_MAP.keys())

AVAILABLE_MODULES = {
    "AMZ": {
        "mod_no": "MODULE 01",
        "name": "1. Amazon Invoice Editor",
        "code": "AMZ-ED",
        "short_open": "Open AMZ-ED",
        "logo": "📄",
        "color": "#1d4ed8",
        "badge_bg": "#eff6ff",
        "icon_bg": "#dbeafe",
        "arrow_bg": "#eff6ff",
        "hover_cls": "hover-amz"
    },
    "BLNK_PO": {
        "mod_no": "MODULE 02",
        "name": "2. Blinkit PO Pouch",
        "code": "BLNK-PO",
        "short_open": "Open BLNK-PO",
        "logo": "🛍️",
        "color": "#059669",
        "badge_bg": "#ecfdf5",
        "icon_bg": "#d1fae5",
        "arrow_bg": "#ecfdf5",
        "hover_cls": "hover-blnk-po"
    },
    "BLNK": {
        "mod_no": "MODULE 03",
        "name": "3. Blinkit e-Invoice Tool",
        "code": "BLNK-INV",
        "short_open": "Open BLNK-INV",
        "logo": "🧾",
        "color": "#d97706",
        "badge_bg": "#fffbeb",
        "icon_bg": "#fef3c7",
        "arrow_bg": "#fffbeb",
        "hover_cls": "hover-blnk"
    },
    "LBL": {
        "mod_no": "MODULE 04",
        "name": "4. Label Editor Tool",
        "code": "LBL-ED",
        "short_open": "Open LBL-ED",
        "logo": "🏷️",
        "color": "#7c3aed",
        "badge_bg": "#f5f3ff",
        "icon_bg": "#ede9fe",
        "arrow_bg": "#f5f3ff",
        "hover_cls": "hover-lbl"
    },
    "CONF": {
        "mod_no": "MODULE 05",
        "name": "5. Settings & Config",
        "code": "SETTINGS",
        "short_open": "Open SETTINGS",
        "logo": "⚙️",
        "color": "#475569",
        "badge_bg": "#f1f5f9",
        "icon_bg": "#e2e8f0",
        "arrow_bg": "#f1f5f9",
        "hover_cls": "hover-conf"
    }
}

if "selected_wh" not in st.session_state:
    st.session_state.selected_wh = "DEL"
if "open_tabs" not in st.session_state:
    st.session_state.open_tabs = ["HOME", "AMZ"]
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "HOME"
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "show_history" not in st.session_state:
    st.session_state.show_history = False
if "current_batch_results" not in st.session_state:
    st.session_state.current_batch_results = []
if "blnk_batch_results" not in st.session_state:
    st.session_state.blnk_batch_results = []

def get_current_location_history():
    current_tab = st.session_state.active_tab
    current_wh = st.session_state.selected_wh

    all_disk_history = load_disk_history()
    module_items = [item for item in all_disk_history if item.get("module") == current_tab]

    if current_wh == "Headoffice":
        return module_items

    return [item for item in module_items if item.get("location") == current_wh]

def reset_current_workspace():
    st.session_state.uploader_key += 1
    st.session_state.current_batch_results = []
    st.session_state.blnk_batch_results = []
    st.session_state.show_history = False
    st.rerun()

def clear_all_pinned_tabs():
    st.session_state.open_tabs = ["HOME"]
    st.session_state.active_tab = "HOME"
    st.session_state.show_history = False
    st.rerun()

# ----------------- SIDEBAR: LIVE INDIA ANALOG CLOCK -----------------
with st.sidebar:
    clock_html = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
    .clock-wrapper {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 10px 0px;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    .analog-clock {
        width: 140px;
        height: 140px;
        border-radius: 50%;
        background: #0f172a;
        border: 4px solid #38bdf8;
        box-shadow: 0 0 18px rgba(56, 189, 248, 0.4);
        position: relative;
        margin-bottom: 12px;
    }
    .center-dot {
        width: 8px;
        height: 8px;
        background: #38bdf8;
        border-radius: 50%;
        position: absolute;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        z-index: 10;
    }
    .hand {
        position: absolute;
        bottom: 50%;
        left: 50%;
        transform-origin: bottom center;
        border-radius: 4px;
    }
    .hour-hand {
        width: 4px;
        height: 38px;
        background: #f8fafc;
        margin-left: -2px;
    }
    .minute-hand {
        width: 2.5px;
        height: 52px;
        background: #94a3b8;
        margin-left: -1.25px;
    }
    .second-hand {
        width: 1.5px;
        height: 58px;
        background: #ef4444;
        margin-left: -0.75px;
    }
    .time-ist-tag {
        font-size: 11px;
        font-weight: 800;
        color: #38bdf8;
        letter-spacing: 1px;
    }
    .digital-time {
        font-size: 15px;
        font-weight: 700;
        color: #f1f5f9;
        margin-top: 2px;
    }
    .digital-date {
        font-size: 11px;
        color: #94a3b8;
        margin-top: 1px;
    }
    </style>
    </head>
    <body>
    <div class="clock-wrapper">
        <div class="analog-clock">
            <div class="center-dot"></div>
            <div class="hand hour-hand" id="hour"></div>
            <div class="hand minute-hand" id="min"></div>
            <div class="hand second-hand" id="sec"></div>
        </div>
        <div class="time-ist-tag">🇮🇳 INDIA TIME (IST)</div>
        <div class="digital-time" id="digi-time">--:--:-- --</div>
        <div class="digital-date" id="digi-date">--</div>
    </div>
    <script>
    function updateClock() {
        const now = new Date();
        const utc = now.getTime() + (now.getTimezoneOffset() * 60000);
        const ist = new Date(utc + (3600000 * 5.5));
        
        const h = ist.getHours();
        const m = ist.getMinutes();
        const s = ist.getSeconds();

        const hDeg = (h % 12) * 30 + (m / 2);
        const mDeg = m * 6;
        const sDeg = s * 6;

        document.getElementById('hour').style.transform = `rotate(${hDeg}deg)`;
        document.getElementById('min').style.transform = `rotate(${mDeg}deg)`;
        document.getElementById('sec').style.transform = `rotate(${sDeg}deg)`;

        const ampm = h >= 12 ? 'PM' : 'AM';
        const h12 = h % 12 || 12;
        const pad = (n) => n < 10 ? '0' + n : n;
        document.getElementById('digi-time').innerText = `${pad(h12)}:${pad(m)}:${pad(s)} ${ampm}`;
        
        const options = { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' };
        document.getElementById('digi-date').innerText = ist.toLocaleDateString('en-IN', options);
    }
    setInterval(updateClock, 1000);
    updateClock();
    </script>
    </body>
    </html>
    """
    components.html(clock_html, height=230)
    st.divider()
    st.markdown(f"**WH Location:** `{LOCATION_DISPLAY_MAP.get(st.session_state.selected_wh, st.session_state.selected_wh)}`")
    st.caption("Operations Hub System Engine v3.2")

# ----------------- TOP BAR: HEADER BRANDING & CONTROLS -----------------
h_left, h_right = st.columns([7.5, 2.5])

with h_left:
    st.markdown("""
        <div class="pan-brand-title">🏛️ PAN INDIA OPERATIONS HUB</div>
        <div class="pan-brand-subtitle">Romsons prime pvt ltd operation portal</div>
    """, unsafe_allow_html=True)

with h_right:
    st.markdown('<div class="right-action-cluster">', unsafe_allow_html=True)
    loc_sub1, loc_sub2 = st.columns([0.38, 1.62])
    with loc_sub1:
        if st.button("🔄", key="btn_refresh_pinned", help="Clear all extra tabs", use_container_width=True):
            clear_all_pinned_tabs()
    with loc_sub2:
        selected_display = st.selectbox(
            "WH | Location",
            options=ORDERED_LOCATIONS,
            format_func=lambda x: LOCATION_DISPLAY_MAP.get(x, x),
            index=ORDERED_LOCATIONS.index(st.session_state.selected_wh) if st.session_state.selected_wh in ORDERED_LOCATIONS else 0
        )
        st.session_state.selected_wh = selected_display

    loc_history = get_current_location_history()
    hist_col, reset_col = st.columns([1.2, 0.8])
    with hist_col:
        hist_title = f"🕒 History ({len(loc_history)})"
        if st.button(hist_title, key="top_hist_btn", type="primary" if st.session_state.show_history else "secondary", use_container_width=True):
            st.session_state.show_history = not st.session_state.show_history
            st.rerun()
    with reset_col:
        if st.button("🔄 Reset", key="top_reset_btn", type="secondary", help="Reset workspace", use_container_width=True):
            reset_current_workspace()

    st.markdown('</div>', unsafe_allow_html=True)

# ----------------- PINNED TABS: 'HOME' REPLACES 'DASHBOARD' -----------------
st.markdown('<div class="top-tab-strip">', unsafe_allow_html=True)
tab_cols = st.columns([1.0] * len(st.session_state.open_tabs) + [3.0])
for idx, t_key in enumerate(st.session_state.open_tabs):
    if t_key == "HOME":
        tab_label = "🏠 HOME"
    else:
        info = AVAILABLE_MODULES[t_key]
        tab_label = f"{info['logo']} {info['code']} ✕" if len(st.session_state.open_tabs) > 1 else f"{info['logo']} {info['code']}"

    is_active = (t_key == st.session_state.active_tab and not st.session_state.show_history)
    with tab_cols[idx]:
        if st.button(tab_label, key=f"top_chip_{t_key}", type="primary" if is_active else "secondary"):
            if is_active and t_key != "HOME" and len(st.session_state.open_tabs) > 1:
                st.session_state.open_tabs.remove(t_key)
                st.session_state.active_tab = st.session_state.open_tabs[0]
            else:
                st.session_state.active_tab = t_key
                st.session_state.show_history = False
            st.rerun()
st.markdown('</div>', unsafe_allow_html=True)

st.divider()

# ----------------- DASHBOARD VIEW: OVAL CARDS (CLICK ANYWHERE TO OPEN) -----------------
if st.session_state.active_tab == "HOME" and not st.session_state.show_history:
    st.markdown('<div style="font-size: 13px; font-weight: 800; color: #64748b; margin-bottom: 12px; letter-spacing: 0.8px;">SELECT OPERATION MODULE</div>', unsafe_allow_html=True)

    keys = list(AVAILABLE_MODULES.keys())
    for i in range(0, len(keys), 2):
        c1, c2 = st.columns(2)
        row_keys = keys[i:i+2]
        
        for col_obj, m_key in zip([c1, c2], row_keys):
            mod_data = AVAILABLE_MODULES[m_key]
            with col_obj:
                # OVAL CARD HTML + INVISIBLE FULL BUTTON (NO SEPARATE BOTTOM BUTTONS)
                card_markup = f"""
                <div class="oval-card-wrapper {mod_data['hover_cls']}">
                    <div class="oval-card">
                        <div class="card-top-flex">
                            <div class="icon-bubble" style="background-color: {mod_data['icon_bg']}; color: {mod_data['color']};">
                                {mod_data['logo']}
                            </div>
                            <span class="badge-pill" style="background-color: {mod_data['badge_bg']}; color: {mod_data['color']};">
                                {mod_data['mod_no']}
                            </span>
                        </div>
                        <div class="card-heading">{mod_data['name']}</div>
                        <div class="card-foot-flex">
                            <span class="foot-label" style="color: {mod_data['color']};">{mod_data['short_open']}</span>
                            <div class="foot-circle" style="background-color: {mod_data['arrow_bg']}; color: {mod_data['color']};">➔</div>
                        </div>
                    </div>
                """
                st.markdown(card_markup, unsafe_allow_html=True)
                
                # Hidden full overlay button on entire box
                if st.button(f"open_mod_{m_key}", key=f"ovl_btn_{m_key}"):
                    if m_key not in st.session_state.open_tabs:
                        st.session_state.open_tabs.append(m_key)
                    st.session_state.active_tab = m_key
                    st.rerun()

                st.markdown("</div>", unsafe_allow_html=True)

# ----------------- HISTORY VIEW -----------------
elif st.session_state.show_history:
    cur_mod_info = AVAILABLE_MODULES.get(st.session_state.active_tab, {"name": "Module", "logo": "📁"})
    h_top_c1, h_top_c2 = st.columns([4, 1])
    with h_top_c1:
        if st.session_state.selected_wh == "Headoffice":
            st.subheader(f"📜 {cur_mod_info['logo']} {cur_mod_info['name']} — Central Master History (All Locations)")
        else:
            st.subheader(f"📜 {cur_mod_info['logo']} {cur_mod_info['name']} — [{LOCATION_DISPLAY_MAP.get(st.session_state.selected_wh, st.session_state.selected_wh)}]")

    with h_top_c2:
        if loc_history:
            clear_btn_label = "🗑️ Clear Master History" if st.session_state.selected_wh == "Headoffice" else "🗑️ Clear Location History"
            if st.button(clear_btn_label, key="btn_del_hist", use_container_width=True):
                clear_disk_history(wh_filter=st.session_state.selected_wh, mod_filter=st.session_state.active_tab)
                st.rerun()

    if not loc_history:
        st.info("No files in history.")
    else:
        st.success(f"💾 Total Saved Files: **{len(loc_history)}** (Auto retained for 30 days)")

        if len(loc_history) > 1:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for idx, item in enumerate(loc_history):
                    loc_tag = item.get("location", "GEN")
                    name_root = item['file_name'][:-4] if item['file_name'].endswith(".pdf") else item['file_name']

                    safe_pdf_name = f"{idx+1}_{loc_tag}_{item['file_name']}"
                    zf.writestr(safe_pdf_name, item["data"])

                    if "excel_data" in item and item["excel_data"]:
                        zf.writestr(f"{idx+1}_{loc_tag}_{name_root}.xlsx", item["excel_data"])

                    if "json_data" in item and item["json_data"]:
                        zf.writestr(f"{idx+1}_{loc_tag}_{name_root}.json", item["json_data"])

            zip_bytes = zip_buf.getvalue()

            zip_title = f"📦 Download All Complete Invoices Bundle as ZIP ({len(loc_history)} Records)"
            st.download_button(
                label=zip_title,
                data=zip_bytes,
                file_name=f"{st.session_state.selected_wh}_{st.session_state.active_tab}_Master_Bundle.zip",
                mime="application/zip",
                key="btn_zip_history_download",
                use_container_width=True
            )

        st.write("---")
        for idx, item in enumerate(loc_history):
            name_root = item['file_name'][:-4] if item['file_name'].endswith(".pdf") else item['file_name']
            
            d_col1, d_col2, d_col3, d_col4 = st.columns([3.2, 0.9, 0.9, 0.9])
            with d_col1:
                st.write(
                    f"📁 **{item['file_name']}** — `[{LOCATION_DISPLAY_MAP.get(item.get('location'), item.get('location'))}]` | "
                    f"⏰ *{item.get('time', '')}*"
                )
            with d_col2:
                st.download_button(
                    label="📥 PDF",
                    data=item["data"],
                    file_name=item["file_name"],
                    mime="application/pdf",
                    key=f"dl_hist_pdf_{st.session_state.selected_wh}_{idx}",
                    use_container_width=True
                )
            with d_col3:
                if "excel_data" in item and item["excel_data"]:
                    st.download_button(
                        label="📊 Excel",
                        data=item["excel_data"],
                        file_name=f"{name_root}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_hist_excel_{st.session_state.selected_wh}_{idx}",
                        use_container_width=True
                    )
                else:
                    st.button("📊 Excel", key=f"dis_xl_{idx}", disabled=True, use_container_width=True)
            with d_col4:
                if "json_data" in item and item["json_data"]:
                    st.download_button(
                        label="🧾 JSON",
                        data=item["json_data"],
                        file_name=f"{name_root}.json",
                        mime="application/json",
                        key=f"dl_hist_json_{st.session_state.selected_wh}_{idx}",
                        use_container_width=True
                    )
                else:
                    st.button("🧾 JSON", key=f"dis_js_{idx}", disabled=True, use_container_width=True)

# ----------------- MODULE 1: AMAZON INVOICE EDITOR -----------------
elif st.session_state.active_tab == "AMZ":
    st.subheader("📑 Amazon Invoice Editor & Barcode Stamper")

    TARGET_PAN = "aalcr5906l"

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        uploaded_csv = st.file_uploader(
            "1. Upload Shipment Report (CSV / Excel)", 
            type=["csv", "xlsx", "xls"], 
            key=f"amz_csv_{st.session_state.selected_wh}_{st.session_state.uploader_key}"
        )
    with col_u2:
        uploaded_pdfs = st.file_uploader(
            "2. Upload Invoice PDF(s) [Multiple Files Allowed]", 
            type=["pdf"], 
            accept_multiple_files=True, 
            key=f"amz_pdf_{st.session_state.selected_wh}_{st.session_state.uploader_key}"
        )

    REGEX_ORDER_HYPHEN = re.compile(r'(\d{3})\s*[-–—\s]\s*(\d{7})\s*[-–—\s]\s*(\d{7})')
    REGEX_ORDER_KEYWORD = re.compile(r'order\s*number\s*[:\s]*([0-9\-–—\s]{17,25})', re.IGNORECASE)
    REGEX_WORDS = re.compile(r'[a-zA-Z0-9]+')
    STOP_WORDS = {
        'the', 'and', 'for', 'with', 'pcs', 'piece', 'pieces', 'only', 'total', 
        'hsn', 'gst', 'rs', 'inr', 'tax', 'invoice', 'date', 'order', 'qty',
        'rate', 'amount', 'shipping', 'charges', 'state', 'code', 'pan', 'bill', 'to'
    }

    def clean_val(v):
        if pd.isna(v):
            return ""
        s = str(v).strip()
        return s.strip("=\"'")

    def clean_alphanumeric(text):
        return re.sub(r'[^a-zA-Z0-9]', '', str(text)).lower()

    def get_token_words(text):
        if not text:
            return set()
        words = REGEX_WORDS.findall(str(text).lower())
        return {w for w in words if len(w) >= 2 and w not in STOP_WORDS}

    def has_valid_description(raw_text):
        text_lower = raw_text.lower()
        if not any(k in text_lower for k in ("description", "item description", "product description", "title")):
            return False
        tokens = get_token_words(raw_text)
        return len(tokens) >= 5

    BARCODE_CACHE = {}
    code128_cls = barcode.get_barcode_class('code128')
    barcode_writer = ImageWriter()
    barcode_writer.font_path = None

    def get_cached_barcode(code_text):
        clean_code = str(code_text).strip()
        if clean_code in BARCODE_CACHE:
            return BARCODE_CACHE[clean_code]
        try:
            barcode_instance = code128_cls(clean_code, writer=barcode_writer)
            buffer = io.BytesIO()
            barcode_instance.write(
                buffer,
                options={
                    'write_text': False,
                    'module_width': 0.45,
                    'module_height': 15.0,
                    'quiet_zone': 1.5,
                    'dpi': 200
                }
            )
            data = buffer.getvalue()
            BARCODE_CACHE[clean_code] = data
            return data
        except Exception:
            return None

    if uploaded_csv and uploaded_pdfs:
        try:
            if uploaded_csv.name.endswith(".csv"):
                try:
                    df = pd.read_csv(uploaded_csv, dtype=str)
                except UnicodeDecodeError:
                    uploaded_csv.seek(0)
                    df = pd.read_csv(uploaded_csv, encoding="latin1", dtype=str)
            else:
                df = pd.read_excel(uploaded_csv, dtype=str)
        except Exception as e:
            st.error(f"CSV read error: {e}")
            st.stop()

        col_mapping = {str(col).strip().lower(): col for col in df.columns}
        order_col = next((col_mapping[c] for c in col_mapping if "order" in c and "id" in c), None) or \
                    next((col_mapping[c] for c in col_mapping if "order" in c), None)
        tracking_col = next((col_mapping[c] for c in col_mapping if "track" in c or "tracing" in c or "awb" in c), None)
        title_col = next((col_mapping[c] for c in col_mapping if any(k in c for k in ["title", "item-name", "item_name", "product", "desc"])), None)
        sku_col = next((col_mapping[c] for c in col_mapping if "sku" in c or "msku" in c), None)

        if not (order_col and tracking_col):
            st.error("CSV me Order ID aur Tracking ID column hona zaroori hai!")
            st.stop()

        order_records_map = {}
        for _, row in df.iterrows():
            raw_oid = clean_val(row.get(order_col, ""))
            clean_oid = clean_alphanumeric(raw_oid)
            track_val = clean_val(row.get(tracking_col, ""))

            desc_parts = []
            if title_col and not pd.isna(row.get(title_col, "")):
                desc_parts.append(str(row[title_col]))
            if sku_col and not pd.isna(row.get(sku_col, "")):
                desc_parts.append(str(row[sku_col]))

            full_desc = " ".join(desc_parts).strip()
            desc_words = get_token_words(full_desc)

            if clean_oid and track_val:
                if clean_oid not in order_records_map:
                    order_records_map[clean_oid] = []
                order_records_map[clean_oid].append({
                    "full_desc": full_desc,
                    "desc_words": desc_words,
                    "track": track_val,
                    "used": False
                })

        st.info(f"📊 Unique Orders in CSV: **{len(order_records_map)}** | Selected PDFs: **{len(uploaded_pdfs)} file(s)**")

        if st.button("🚀 Process & Generate Stamped Invoices", type="primary", use_container_width=True):
            progress_bar = st.progress(0)
            status_text = st.empty()

            processed_items = []
            total_files = len(uploaded_pdfs)

            barcode_rect = fitz.Rect(40, 58, 235, 82)
            white_color = (1.0, 1.0, 1.0)
            black_color = (0, 0, 0)

            for file_idx, pdf_file in enumerate(uploaded_pdfs):
                status_text.text(f"Processing File {file_idx+1}/{total_files}: {pdf_file.name}...")
                pdf_file.seek(0)
                pdf_bytes = pdf_file.read()
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                new_doc = fitz.open()

                file_stamped_count = 0
                file_removed_pan = 0
                file_removed_desc = 0

                for page_num in range(len(doc)):
                    page = doc[page_num]
                    raw_text = page.get_text()
                    text_lower = raw_text.lower()

                    if TARGET_PAN not in text_lower:
                        file_removed_pan += 1
                        continue

                    if not has_valid_description(raw_text):
                        file_removed_desc += 1
                        continue

                    order_clean = None
                    order_match = REGEX_ORDER_HYPHEN.search(raw_text)
                    if order_match:
                        order_clean = f"{order_match.group(1)}{order_match.group(2)}{order_match.group(3)}"
                    else:
                        num_match = REGEX_ORDER_KEYWORD.search(raw_text)
                        if num_match:
                            extracted = clean_alphanumeric(num_match.group(1))
                            if len(extracted) == 17:
                                order_clean = extracted

                    target_tracking_id = None

                    if order_clean and order_clean in order_records_map:
                        matching_rows = order_records_map[order_clean]

                        if len(matching_rows) == 1:
                            target_tracking_id = matching_rows[0]["track"]
                        else:
                            page_words = get_token_words(raw_text)
                            best_match_row = None
                            highest_overlap = -1

                            for r in matching_rows:
                                if not r["desc_words"]:
                                    continue
                                common = page_words.intersection(r["desc_words"])
                                overlap_score = len(common)

                                if overlap_score > highest_overlap:
                                    highest_overlap = overlap_score
                                    best_match_row = r

                            if best_match_row and highest_overlap > 0:
                                target_tracking_id = best_match_row["track"]
                                best_match_row["used"] = True
                            else:
                                unused_rows = [r for r in matching_rows if not r["used"]]
                                if unused_rows:
                                    target_tracking_id = unused_rows[0]["track"]
                                    unused_rows[0]["used"] = True
                                else:
                                    target_tracking_id = matching_rows[0]["track"]

                    if target_tracking_id:
                        page.draw_rect(barcode_rect, color=white_color, fill=white_color, width=0)

                        barcode_img_bytes = get_cached_barcode(target_tracking_id)
                        if barcode_img_bytes:
                            page.insert_image(barcode_rect, stream=barcode_img_bytes, keep_proportion=False)

                        page.insert_text(
                            (barcode_rect.x0 + 10, 95),
                            f"TRACKING: {target_tracking_id}",
                            fontsize=10.5,
                            fontname="hebo",
                            color=black_color
                        )

                        date_instances = page.search_for("Order Date:") or page.search_for("Order Date")
                        if date_instances:
                            first_date_rect = date_instances[0]
                            page.insert_text(
                                (first_date_rect.x0, first_date_rect.y1 + 13),
                                f"Tracking ID: {target_tracking_id}",
                                fontsize=9.5,
                                fontname="hebo",
                                color=black_color
                            )

                        file_stamped_count += 1

                    new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)

                final_page_count = len(new_doc)
                out_bytes = new_doc.tobytes(deflate=True, clean=False)
                new_doc.close()
                doc.close()

                processed_items.append({
                    "original_name": pdf_file.name,
                    "file_name": f"Stamped_{pdf_file.name}",
                    "data": out_bytes,
                    "pages": final_page_count,
                    "stamped": file_stamped_count,
                    "removed_pan": file_removed_pan,
                    "removed_desc": file_removed_desc,
                    "location": st.session_state.selected_wh,
                    "module": "AMZ"
                })

                progress_bar.progress((file_idx + 1) / total_files)

            status_text.empty()
            progress_bar.empty()

            st.session_state.current_batch_results = processed_items

            st.markdown(f"""
                <div class="pro-success-card">
                    <b style="font-size: 15px; color: #34d399;">✔ Operation Completed Successfully!</b><br>
                    <span style="font-size: 13px; color: #cbd5e1;">All {len(processed_items)} invoice file(s) for <b>{st.session_state.selected_wh}</b> have been accurately verified, filtered, and stamped with dispatch barcodes. Download below to record in History.</span>
                </div>
            """, unsafe_allow_html=True)

    batch_list = st.session_state.current_batch_results
    if batch_list:
        st.write("---")
        st.subheader("📥 Download Invoices")

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
            for idx, item in enumerate(batch_list):
                unique_name = f"{idx+1}_{item['file_name']}"
                zip_file.writestr(unique_name, item["data"])
        bulk_zip_bytes = zip_buffer.getvalue()

        st.download_button(
            label=f"📦 Download All Invoices as ZIP ({len(batch_list)} Files)",
            data=bulk_zip_bytes,
            file_name=f"{st.session_state.selected_wh}_All_Invoices.zip",
            mime="application/zip",
            key="btn_live_zip_download",
            on_click=save_batch_to_disk_history,
            args=(batch_list,),
            use_container_width=True
        )
        st.write("")

        for idx, item in enumerate(batch_list):
            d_col1, d_col2 = st.columns([3, 1])
            with d_col1:
                st.write(
                    f"📁 **{item['original_name']}** — `[{LOCATION_DISPLAY_MAP.get(item.get('location'), item.get('location'))}]` | "
                    f"`{item['pages']} Pages Kept` | `{item['stamped']} Stamped`"
                )
            with d_col2:
                st.download_button(
                    label="📥 Download PDF",
                    data=item["data"],
                    file_name=item["file_name"],
                    mime="application/pdf",
                    key=f"dl_single_live_{st.session_state.selected_wh}_{idx}",
                    on_click=save_item_to_disk_history,
                    args=(item,),
                    use_container_width=True
                )

# ----------------- MODULE 2: BLINKIT PO UPLOADER -----------------
elif st.session_state.active_tab == "BLNK_PO":
    st.subheader("🛍️ Blinkit PO Pouch & Manifest Processor")
    st.caption("Upload Blinkit Purchase Order (PO) PDFs for extraction and dispatch processing.")

    po_col1, po_col2 = st.columns(2)
    with po_col1:
        st.file_uploader(
            "Upload Blinkit PO PDF(s)", 
            type=["pdf"], 
            accept_multiple_files=True, 
            key=f"blnk_po_pdf_{st.session_state.selected_wh}_{st.session_state.uploader_key}"
        )
    with po_col2:
        st.file_uploader(
            "Upload Reference / SKU Mapping Sheet (Optional Excel/CSV)", 
            type=["xlsx", "xls", "csv"], 
            key=f"blnk_po_map_{st.session_state.selected_wh}_{st.session_state.uploader_key}"
        )

# ----------------- MODULE 3: BLINKIT E-INVOICE TOOL -----------------
elif st.session_state.active_tab == "BLNK":
    st.subheader("🧾 Blinkit e-Invoice Management Tool & Reconciler")
    st.caption("Clean Party Name | Dedicated Address Only in Addr1/Addr2 | Exact Taxable Base")

    STATE_CODE_MAP = {
        "01": "JAMMU AND KASHMIR", "02": "HIMACHAL PRADESH", "03": "PUNJAB", "04": "CHANDIGARH",
        "05": "UTTARAKHAND", "06": "HARYANA", "07": "DELHI", "08": "RAJASTHAN",
        "09": "UTTAR PRADESH", "10": "BIHAR", "11": "SIKKIM", "12": "ARUNACHAL PRADESH",
        "13": "NAGALAND", "14": "MANIPUR", "15": "MIZORAM", "16": "TRIPURA",
        "17": "MEGHALAYA", "18": "ASSAM", "19": "WEST BENGAL", "20": "JHARKHAND",
        "21": "ODISHA", "22": "CHATTISGARH", "23": "MADHYA PRADESH", "24": "GUJARAT",
        "26": "DADRA AND NAGAR HAVELI AND DAMAN AND DIU", "27": "MAHARASHTRA", "29": "KARNATAKA",
        "30": "GOA", "31": "LAKSHADWEEP", "32": "KERALA", "33": "TAMIL NADU",
        "34": "PUDUCHERRY", "35": "ANDAMAN AND NICOBAR ISLANDS", "36": "TELANGANA", "37": "ANDHRA PRADESH",
        "38": "LADAKH"
    }

    STATE_NAME_TO_CODE = {v.upper(): k for k, v in STATE_CODE_MAP.items()}

    MAJOR_CITIES = [
        "New Delhi", "Delhi", "Varanasi", "Lucknow", "Jaipur", "Gurgaon", "Gurugram", 
        "Noida", "Ghaziabad", "Kanpur", "Bengaluru", "Bangalore", "Mumbai", "Pune", 
        "Kolkata", "Ahmedabad", "Patna", "Ranchi", "Chandigarh", "Faridabad", "Agra", "Meerut", "Hapur", "Dehradun"
    ]

    def clean_extracted_city(text, state_name):
        m = re.search(r"([A-Za-z\s]+),\s*(?:[A-Za-z\s]+)[\-\s]*[0-9]{6}", text)
        if m:
            c_cand = m.group(1).strip()
            parts = [p.strip() for p in c_cand.split(",") if p.strip()]
            last_p = parts[-1]
            if len(last_p) > 2 and not re.search(r"(road|street|nagar|colony|floor|block|house|marg)", last_p, re.I):
                return last_p.title()

        for city in MAJOR_CITIES:
            if re.search(rf"\b{city}\b", text, re.I):
                return city

        return state_name.title() if state_name else "Delhi"

    def format_best_address(raw_lines, party_name, pincode, city, state_name):
        clean_parts = []
        for l in raw_lines:
            s = l.strip()
            if not s:
                continue
            
            if party_name and party_name.lower() in s.lower():
                continue
            if re.search(r"(@|contact|phone|mob|gstin|pan|india|cgst|sgst|ugst|igst|s\.no|item code|tax invoice|billing address|shipping address)", s, re.I):
                continue
            
            s = re.sub(rf"\b{re.escape(str(pincode))}\b", "", s).strip()
            s = re.sub(rf"\b{re.escape(str(city))}\b", "", s, flags=re.I).strip()
            s = re.sub(rf"\b{re.escape(str(state_name))}\b", "", s, flags=re.I).strip()
            s = re.sub(r'[\,\-\s]+$', '', s).strip()
            s = re.sub(r'^[\,\-\s]+', '', s).strip()
            
            if len(s) > 2:
                clean_parts.append(s)

        full_addr = ", ".join(clean_parts) if clean_parts else "Commercial Facility"
        full_addr = re.sub(r"\s+", " ", full_addr).strip()
        
        if party_name:
            full_addr = re.sub(rf"^{re.escape(party_name)}[\,\s\-]*", "", full_addr, flags=re.I).strip()

        if len(full_addr) <= 100:
            addr1 = full_addr
            addr2 = ""
        else:
            split_idx = full_addr[:98].rfind(",")
            if split_idx > 25:
                addr1 = full_addr[:split_idx].strip()
                addr2 = full_addr[split_idx+1:].strip()[:100]
            else:
                addr1 = full_addr[:98].strip()
                addr2 = full_addr[98:].strip()[:100]

        return addr1, addr2

    def fmt_dec(val):
        try:
            f = float(str(val).replace(",", "").strip())
            return f"{f:.2f}"
        except (ValueError, TypeError):
            return "0.00"

    def clean_description_completely(desc, hsn_code=None):
        if not desc:
            return ""
        if hsn_code:
            desc = re.sub(rf"\b{re.escape(str(hsn_code))}\b", "", desc)
        desc = re.sub(r"\b\d{6,8}\b", "", desc)
        desc = re.sub(r'(\b[A-Za-z]+)\s+([a-z]{1,4}\b)', lambda m: m.group(1) + m.group(2) if (m.group(1)+m.group(2)).lower() in ["three", "cock", "valve", "dispo", "cannula", "piece", "stopcock"] else m.group(0), desc)
        desc = re.sub(r"\s+", " ", desc).strip()
        return desc

    def overwrite_exact_area(page, rect, new_text, font_size=7.2):
        safe_rect = fitz.Rect(rect.x0, rect.y0 + 0.3, rect.x1, rect.y1 - 0.3)
        page.draw_rect(safe_rect, color=None, fill=(1, 1, 1))
        page.insert_text((safe_rect.x0 + 0.5, safe_rect.y1 - 1.0), str(new_text), fontsize=font_size, fontname="helv", color=(0, 0, 0))

    def extract_metadata_from_doc(doc):
        full_text = ""
        for p in doc:
            full_text += p.get_text() + "\n"

        page0 = doc[0]
        words = page0.get_text("words")

        ext_order_id = None
        ext_matches = [w for w in words if "extern" in w[4].lower()]
        if ext_matches:
            anchor = ext_matches[0]
            num_words = [
                w for w in words 
                if (anchor[1] - 4 <= w[1] <= anchor[3] + 11.5) 
                and (anchor[0] + 45 <= w[0] <= 445)
                and w[4].strip().rstrip('.').isdigit()
            ]
            num_words.sort(key=lambda w: (round(w[1] / 6) * 6, w[0]))
            if num_words:
                ext_order_id = "".join([w[4].strip().rstrip('.') for w in num_words])

        if not ext_order_id:
            ext_match = re.search(r"Extern(?:al)?\s*Order\s*(?:No\.?|ID)?\s*[:\s]*([0-9\s]+?)(?=\n\s*[A-Za-z]|Invoice|Date|$)", full_text, re.IGNORECASE)
            if ext_match:
                ext_order_id = "".join(ext_match.group(1).split()).rstrip('.')
        if not ext_order_id:
            digits_match = re.findall(r"\b(49\d{8,14}|5\d{8,14}|\d{12,18})\b", full_text)
            if digits_match:
                ext_order_id = digits_match[0]
        if not ext_order_id:
            ext_order_id = "ExtOrder"

        ext_order_id = re.sub(r'[\.\s]+$', '', str(ext_order_id))

        inv_match = re.search(r"Invoice\s*No\s*[:\s]*([A-Za-z0-9\-_/]+)", full_text, re.IGNORECASE)
        invoice_no = inv_match.group(1).strip() if inv_match else "INV-001"

        date_match = re.search(r"Invoice\s*Date\s*[:\s]*([A-Za-z0-9,\s\.\-\/]+?)(?=\n|Ship\s*Date|Order|$)", full_text, re.IGNORECASE)
        raw_date = date_match.group(1).strip() if date_match else "Date"

        clean_date = raw_date
        std_date_for_json = datetime.now(IST).strftime("%d/%m/%Y")
        for fmt in ("%b %d, %Y", "%B %d, %Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                d_obj = datetime.strptime(raw_date, fmt)
                clean_date = d_obj.strftime("%d-%m-%Y")
                std_date_for_json = d_obj.strftime("%d/%m/%Y")
                break
            except ValueError:
                pass

        clean_ext = re.sub(r'[^A-Za-z0-9\-_]', '', ext_order_id)
        clean_inv = re.sub(r'[^A-Za-z0-9\-_]', '', invoice_no)
        clean_dt = re.sub(r'[^A-Za-z0-9\-_]', '', clean_date)
        pdf_filename = f"{clean_ext}_{clean_inv}_{clean_dt}.pdf"

        # --- SELLER EXTRACTION ---
        sold_header = [w for w in words if "sold" in w[4].lower()]
        bill_header = [w for w in words if "billing" in w[4].lower()]
        y_sold = sold_header[0][1] if sold_header else 80
        y_bill = bill_header[0][1] if bill_header else 200

        seller_rect = fitz.Rect(20, y_sold, 295, y_bill)
        seller_text = page0.get_text("text", clip=seller_rect).strip()
        s_lines = [l.strip() for l in seller_text.split("\n") if l.strip() and not re.search(r"sold\s*by", l, re.I)]
        seller_name = s_lines[0] if s_lines else "ROMSONS PRIME PRIVATE LIMITED"

        s_gst = re.search(r"GSTIN\s*[:\s]*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", seller_text, re.I)
        seller_gstin = s_gst.group(1).strip() if s_gst else "07AALCR5906L1ZW"

        s_pin = re.findall(r"\b[1-9][0-9]{5}\b", seller_text)
        seller_pin = int(s_pin[0]) if s_pin else 110028
        seller_state_code = seller_gstin[:2]
        seller_state_name = STATE_CODE_MAP.get(seller_state_code, "DELHI")
        seller_loc = clean_extracted_city(seller_text, seller_state_name)
        s_addr1, s_addr2 = format_best_address(s_lines[1:], seller_name, seller_pin, seller_loc, seller_state_name)

        # ----------------- PARSING BILLING & SHIPPING ADDRESSES -----------------
        b_idx = full_text.find("Billing Address")
        s_idx = full_text.find("Shipping Address")
        t_idx = full_text.find("S.No")
        if t_idx == -1:
            t_idx = full_text.find("Item Code")
        if t_idx == -1:
            t_idx = len(full_text)

        start_addr_pos = min(pos for pos in [b_idx, s_idx] if pos != -1) if (b_idx != -1 or s_idx != -1) else 0
        raw_combined_addr = full_text[start_addr_pos:t_idx]

        cleaned_block = re.sub(r"Billing\s*Address", "", raw_combined_addr, flags=re.I)
        cleaned_block = re.sub(r"Shipping\s*Address", "", cleaned_block, flags=re.I).strip()

        pan_matches = list(re.finditer(r"PAN/IT\s*No\s*[:\s]*[A-Z0-9]+", cleaned_block, re.I))
        if len(pan_matches) >= 2:
            split_pos = pan_matches[0].end()
            billing_text_block = cleaned_block[:split_pos].strip()
            shipping_text_block = cleaned_block[split_pos:].strip()
        else:
            billing_text_block = cleaned_block
            shipping_text_block = cleaned_block

        b_lines = [l.strip() for l in billing_text_block.split("\n") if l.strip()]
        shp_lines = [l.strip() for l in shipping_text_block.split("\n") if l.strip()]

        full_lower = full_text.lower()
        if "blink" in full_lower or "blinkit" in full_lower:
            buyer_name = "BLINK COMMERCE PRIVATE LIMITED"
        elif "g n healthcare" in full_lower or "g. n healthcare" in full_lower or "g.n healthcare" in full_lower:
            buyer_name = "G N Healthcare"
        else:
            cand = b_lines[0] if b_lines else "COMMERCIAL BUYER"
            buyer_name = cand if len(cand) > 2 else "COMMERCIAL BUYER"

        em_m = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", billing_text_block)
        buyer_email = em_m.group(0).strip() if em_m else ""
        ph_m = re.search(r"(?:Contact|Phone|Mob)?\s*[:\s]*([6-9]\d{9})", billing_text_block, re.IGNORECASE)
        buyer_phone = ph_m.group(1).strip() if ph_m else ""

        b_gst_m = re.search(r"GSTIN\s*[:\s]*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", billing_text_block, re.IGNORECASE)
        buyer_gstin = b_gst_m.group(1).strip() if b_gst_m else ""
        if not buyer_gstin:
            all_gstins = re.findall(r"\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b", full_text)
            buyer_gstin = all_gstins[1] if len(all_gstins) > 1 else (all_gstins[0] if all_gstins else "06AAFCG9846E1ZF")

        buyer_state_code = buyer_gstin[:2]
        buyer_state_name = STATE_CODE_MAP.get(buyer_state_code, "DELHI")
        b_pin_m = re.findall(r"\b[1-9][0-9]{5}\b", billing_text_block)
        buyer_pin = int(b_pin_m[0]) if b_pin_m else int(buyer_state_code + "0001")
        buyer_loc = clean_extracted_city(billing_text_block, buyer_state_name)
        buyer_addr1, buyer_addr2 = format_best_address(b_lines, buyer_name, buyer_pin, buyer_loc, buyer_state_name)

        # Shipping Details
        shp_gst_m = re.search(r"GSTIN\s*[:\s]*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", shipping_text_block, re.IGNORECASE)
        ship_gstin = shp_gst_m.group(1).strip() if shp_gst_m else buyer_gstin
        ship_state_code = ship_gstin[:2] if len(ship_gstin) >= 2 else buyer_state_code
        ship_state_name = STATE_CODE_MAP.get(ship_state_code, buyer_state_name)
        shp_pin_m = re.findall(r"\b[1-9][0-9]{5}\b", shipping_text_block)
        ship_pin = int(shp_pin_m[0]) if shp_pin_m else buyer_pin
        ship_loc = clean_extracted_city(shipping_text_block, ship_state_name)

        ship_name = buyer_name
        for line in shp_lines:
            clean_l = re.sub(r'[^a-zA-Z0-9\s]', '', line).strip()
            if len(clean_l) > 3 and not re.search(r"(shipping|billing|gstin|contact|phone|mob|email|pan|india)", clean_l, re.I):
                ship_name = line.strip()
                break

        ship_addr1, ship_addr2 = format_best_address(shp_lines, ship_name, ship_pin, ship_loc, ship_state_name)

        is_blinkit = bool(re.search(r"blink\s*commerce|blinkit", f"{buyer_name} {full_text}", re.IGNORECASE))

        # ----------------- SHIPPING RULE -----------------
        if is_blinkit:
            shipping_is_different = False
        else:
            norm_b_addr = re.sub(r'[^a-zA-Z0-9]', '', f"{buyer_addr1}{buyer_addr2}{buyer_pin}").lower()
            norm_s_addr = re.sub(r'[^a-zA-Z0-9]', '', f"{ship_addr1}{ship_addr2}{ship_pin}").lower()
            shipping_is_different = bool(norm_s_addr and (norm_s_addr != norm_b_addr))

        pos_match = re.search(r"Place\s*of\s*Supply\s*[:\s]*([A-Za-z0-9\s\-]+?)(?=\n|GST|State|$)", full_text, re.IGNORECASE)
        pos_raw = pos_match.group(1).strip() if pos_match else ""
        pos_digits = re.findall(r"\b\d{2}\b", pos_raw)

        if pos_digits:
            buyer_pos_code = pos_digits[0]
            pos_state_name = STATE_CODE_MAP.get(buyer_pos_code, buyer_state_name)
        else:
            matched_code = None
            for s_name, s_code in STATE_NAME_TO_CODE.items():
                if s_name in pos_raw.upper():
                    matched_code = s_code
                    pos_state_name = s_name.title()
                    break
            buyer_pos_code = matched_code if matched_code else buyer_state_code
            pos_state_name = STATE_CODE_MAP.get(buyer_pos_code, buyer_state_name)

        is_interstate = (seller_state_code != buyer_pos_code)

        inv_total_m = re.search(r"(?:Total\s*Invoice\s*Value|Invoice\s*Total|Grand\s*Total|Total\s*Amount)[:\s]*[₹\s]*([0-9,]+\.\d{2})", full_text, re.I)
        printed_grand_total = float(inv_total_m.group(1).replace(",", "")) if inv_total_m else None

        summary_m = re.search(r"Total\s*Amount\(\+\)\s*[:\s]*([0-9,]+\.\d{2})", full_text, re.I)
        printed_taxable_total = float(summary_m.group(1).replace(",", "")) if summary_m else None

        tax_sum_m = re.search(r"Total\s*Tax\(\+\)\s*[:\s]*([0-9,]+\.\d{2})", full_text, re.I)
        printed_tax_total = float(tax_sum_m.group(1).replace(",", "")) if tax_sum_m else None

        return {
            "invoice_no": invoice_no,
            "ext_order_id": ext_order_id,
            "doc_date": std_date_for_json,
            "clean_date": clean_date,
            "pdf_filename": pdf_filename,
            "seller_name": seller_name,
            "seller_gstin": seller_gstin,
            "seller_pin": seller_pin,
            "seller_addr1": s_addr1,
            "seller_addr2": s_addr2,
            "seller_loc": seller_loc,
            "seller_state_code": seller_state_code,
            "seller_state_name": seller_state_name,
            "buyer_name": buyer_name,
            "buyer_trade_name": buyer_name,
            "buyer_addr1": buyer_addr1,
            "buyer_addr2": buyer_addr2,
            "buyer_loc": buyer_loc,
            "buyer_pin": buyer_pin,
            "buyer_phone": buyer_phone,
            "buyer_email": buyer_email,
            "buyer_gstin": buyer_gstin,
            "buyer_state_code": buyer_state_code,
            "buyer_state_name": buyer_state_name,
            "buyer_pos": buyer_pos_code,
            "pos_state_name": pos_state_name,
            "is_interstate": is_interstate,
            "shipping_is_different": shipping_is_different,
            "ship_name": ship_name,
            "ship_addr1": ship_addr1,
            "ship_addr2": ship_addr2,
            "ship_loc": ship_loc,
            "ship_pin": ship_pin,
            "ship_gstin": ship_gstin,
            "ship_state_name": ship_state_name,
            "ship_state_code": ship_state_code,
            "is_blinkit": is_blinkit,
            "printed_grand_total": printed_grand_total,
            "printed_taxable_total": printed_taxable_total,
            "printed_tax_total": printed_tax_total
        }

    def process_and_reconcile_pdf(pdf_bytes):
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        meta = extract_metadata_from_doc(doc)
        is_blinkit = meta["is_blinkit"]
        line_items = []

        for p_idx, page in enumerate(doc):
            words = page.get_text("words")

            table_header_bottom_y = 200
            y_groups = {}
            for w in words:
                y_key = round(w[1] / 7) * 7
                y_groups.setdefault(y_key, []).append(w)

            for y_key, g_words in sorted(y_groups.items()):
                texts = [w[4].lower() for w in g_words]
                full_line = " ".join(texts)
                matches = sum(1 for k in ["hsn", "qty", "quantity", "unit price", "price", "rate", "taxable", "s.no", "item code"] if k in full_line)
                if matches >= 2:
                    table_header_bottom_y = max(w[3] for w in g_words)
                    break

            table_bottom_y = page.rect.height - 30
            for w in words:
                wt = w[4].lower().strip()
                if w[1] > (table_header_bottom_y + 30) and any(k in wt for k in ["total", "subtotal", "gst summary", "summary", "words"]):
                    table_bottom_y = min(table_bottom_y, w[1])

            sno_candidates = []
            for w in words:
                val = w[4].strip()
                if val.isdigit() and 1 <= int(val) <= 99:
                    if w[0] < 85 and w[1] > (table_header_bottom_y + 2):
                        if not any(abs(c["center_y"] - ((w[1]+w[3])/2)) < 8 for c in sno_candidates):
                            sno_candidates.append({
                                "sno": int(val),
                                "y0": w[1],
                                "y1": w[3],
                                "center_y": (w[1] + w[3]) / 2
                            })

            sno_candidates.sort(key=lambda x: x["center_y"])

            if sno_candidates:
                for idx, item_anchor in enumerate(sno_candidates):
                    y_top = item_anchor["y0"] - 3
                    if idx + 1 < len(sno_candidates):
                        y_bottom = sno_candidates[idx + 1]["y0"] - 2
                    else:
                        y_bottom = min(item_anchor["y1"] + 90, table_bottom_y)

                    row_words = [w for w in words if y_top <= ((w[1] + w[3]) / 2) <= y_bottom]
                    row_words.sort(key=lambda w: w[0])

                    if not row_words:
                        continue

                    hsn_w = None
                    hsn_code = "63079091"
                    for w in row_words:
                        clean_w = w[4].replace(",", "").strip()
                        if clean_w.isdigit() and len(clean_w) in [6, 7, 8]:
                            hsn_w = w
                            hsn_code = clean_w
                            break

                    hsn_x0 = hsn_w[0] if hsn_w else 200

                    desc_words = [w for w in row_words if 30 <= w[0] < (hsn_x0 - 4) and not w[4].isdigit()]
                    desc_words.sort(key=lambda w: (round(w[1] / 6) * 6, w[0]))
                    raw_desc = " ".join([w[4] for w in desc_words]).strip()
                    pure_desc = clean_description_completely(raw_desc, hsn_code=hsn_code)

                    if any(k in pure_desc.lower() for k in ["private limited", "tehsil", "district", "kumaspur", "cod charge", "shipping charge", "item code"]):
                        continue

                    if not pure_desc:
                        pure_desc = f"Item {item_anchor['sno']}"

                    factor = 1
                    if is_blinkit:
                        upper_desc = pure_desc.upper()
                        if "DISPO" in upper_desc:
                            factor = 50
                        elif "COMFIT" in upper_desc:
                            factor = 25

                    search_words = [w for w in row_words if w[0] > (hsn_w[2] if hsn_w else (hsn_x0 + 5))]
                    search_words.sort(key=lambda w: w[0])

                    after_hsn_nums = []
                    for w in search_words:
                        val = w[4].replace(",", "").strip()
                        if re.match(r"^\d+(\.\d{1,2})?$", val):
                            after_hsn_nums.append({
                                "val": float(val),
                                "str": val,
                                "rect": fitz.Rect(w[0], w[1], w[2], w[3]),
                                "x0": w[0]
                            })

                    final_q = None
                    final_p = None
                    final_disc = 0.0
                    final_taxable = None
                    final_item_tot = None

                    if len(after_hsn_nums) >= 2:
                        final_q = int(after_hsn_nums[0]["val"])
                        if factor > 1 and final_q >= factor:
                            final_q = final_q // factor
                            overwrite_exact_area(page, after_hsn_nums[0]["rect"], f"{final_q}", font_size=7.2)

                        final_p = after_hsn_nums[1]["val"]
                        if factor > 1:
                            final_p = round(final_p * factor, 2)
                            overwrite_exact_area(page, after_hsn_nums[1]["rect"], f"{final_p:.2f}", font_size=7.2)

                    disc_cand = [n for n in after_hsn_nums if 245 <= n["x0"] <= 295]
                    if disc_cand:
                        final_disc = disc_cand[0]["val"]
                    elif len(after_hsn_nums) >= 3 and after_hsn_nums[2]["x0"] < 298:
                        final_disc = after_hsn_nums[2]["val"]

                    taxable_cand = [n for n in after_hsn_nums if 298 <= n["x0"] <= 345]
                    if taxable_cand:
                        final_taxable = taxable_cand[0]["val"]
                    elif len(after_hsn_nums) >= 4:
                        final_taxable = after_hsn_nums[3]["val"]

                    tot_cand = [n for n in after_hsn_nums if n["x0"] >= 520]
                    if tot_cand:
                        final_item_tot = tot_cand[0]["val"]

                    if final_q is not None and final_p is not None and final_q > 0:
                        if final_taxable is not None and final_taxable > 0:
                            taxable_val = round(final_taxable, 2)
                        else:
                            taxable_val = round((final_q * final_p) - final_disc, 2)

                        gross_amt = round(taxable_val + final_disc, 2)
                        unit_label = "BOX" if (is_blinkit and factor > 1) else "PACKS"

                        line_items.append({
                            "sno": len(line_items) + 1,
                            "desc": pure_desc,
                            "hsn": hsn_code,
                            "qty": final_q,
                            "unit": unit_label,
                            "unit_price": final_p,
                            "gross_amt": gross_amt,
                            "discount": final_disc,
                            "taxable_val": taxable_val,
                            "gst_rate": 5.0,
                            "item_tot": final_item_tot
                        })

                    if is_blinkit and factor > 1:
                        uom_in_row = [
                            w for w in row_words 
                            if "UOM" in w[4].upper() and 30 <= w[0] <= 180
                        ]
                        for uw in uom_in_row:
                            hanging_s = [
                                w for w in row_words 
                                if w[4].strip() == "S" and abs(w[0] - uw[0]) < 15 and 0 < (w[1] - uw[1]) < 12
                            ]
                            uom_r = fitz.Rect(uw[0], uw[1], uw[2] + 4, uw[3])
                            if hanging_s:
                                uom_r |= fitz.Rect(hanging_s[0][0], hanging_s[0][1], hanging_s[0][2] + 4, hanging_s[0][3])
                            overwrite_exact_area(page, uom_r, "UOM-BOX", font_size=6.8)

        out_pdf_buf = io.BytesIO()
        doc.save(out_pdf_buf)
        doc.close()
        out_pdf_buf.seek(0)

        meta["line_items"] = line_items
        return out_pdf_buf, meta

    def build_einvoice_json_v101(data):
        is_interstate = data["is_interstate"]
        seller_state_code = data["seller_state_code"]
        buyer_pos = data["buyer_pos"]

        item_list = []
        tot_taxable = 0.0
        tot_cgst = 0.0
        tot_sgst = 0.0
        tot_igst = 0.0

        for idx, item in enumerate(data["line_items"], 1):
            qty = item["qty"]
            unit_price = item["unit_price"]
            unit_label = item.get("unit", "BOX")
            taxable_amt = item.get("taxable_val", round(qty * unit_price, 2))
            tot_taxable += taxable_amt

            gst_rate = item.get("gst_rate", 5.0)

            if is_interstate:
                igst_amt = round((taxable_amt * gst_rate) / 100, 2)
                cgst_amt = 0.0
                sgst_amt = 0.0
                tot_igst += igst_amt
            else:
                cgst_amt = round((taxable_amt * (gst_rate / 2)) / 100, 2)
                sgst_amt = round((taxable_amt * (gst_rate / 2)) / 100, 2)
                igst_amt = 0.0
                tot_cgst += cgst_amt
                tot_sgst += sgst_amt

            tot_item_val = round(taxable_amt + cgst_amt + sgst_amt + igst_amt, 2)

            item_list.append({
                "SlNo": str(idx),
                "PrdDesc": item["desc"][:250],
                "IsServc": "N",
                "HsnCd": item["hsn"],
                "Qty": qty,
                "Unit": unit_label,
                "UnitPrice": round(unit_price, 2),
                "TotAmt": round(item.get("gross_amt", taxable_amt), 2),
                "Discount": round(item.get("discount", 0.0), 2),
                "AssAmt": round(taxable_amt, 2),
                "GstRt": gst_rate,
                "IgstAmt": igst_amt,
                "CgstAmt": cgst_amt,
                "SgstAmt": sgst_amt,
                "CesRt": 0.0,
                "CesAmt": 0.0,
                "TotItemVal": tot_item_val
            })

        exact_calc_total = data.get("printed_grand_total") or round(tot_taxable + tot_cgst + tot_sgst + tot_igst, 2)

        ship_dtls = None
        if data.get("shipping_is_different", False):
            ship_dtls = {
                "Gstin": data["ship_gstin"],
                "LglNm": data["ship_name"][:100],
                "TrdNm": data["ship_name"][:100],
                "Addr1": data["ship_addr1"][:100],
                "Addr2": data.get("ship_addr2", "")[:100] if data.get("ship_addr2") else None,
                "Loc": data["ship_loc"],
                "Pin": int(data["ship_pin"]),
                "Stcd": str(data["ship_state_code"])
            }

        return {
            "Version": "1.01",
            "TranDtls": {
                "TaxSch": "GST",
                "SupTyp": "B2B",
                "RegRev": "N",
                "EcmGstin": None,
                "IgstOnIntra": "N"
            },
            "DocDtls": {
                "Typ": "INV",
                "No": data["invoice_no"],
                "Dt": data["doc_date"]
            },
            "SellerDtls": {
                "Gstin": data["seller_gstin"],
                "LglNm": data["seller_name"],
                "TrdNm": data["seller_name"],
                "Addr1": data["seller_addr1"][:100],
                "Loc": data["seller_loc"],
                "Pin": data["seller_pin"],
                "Stcd": seller_state_code
            },
            "BuyerDtls": {
                "Gstin": data["buyer_gstin"],
                "LglNm": data["buyer_name"],
                "TrdNm": data["buyer_trade_name"],
                "Pos": buyer_pos,
                "Addr1": data["buyer_addr1"][:100],
                "Loc": data["buyer_loc"],
                "Pin": data["buyer_pin"],
                "Stcd": buyer_pos
            },
            "DispDtls": None,
            "ShipDtls": ship_dtls,
            "ItemList": item_list,
            "ValDtls": {
                "AssVal": round(tot_taxable, 2),
                "CgstVal": round(tot_cgst, 2),
                "SgstVal": round(tot_sgst, 2),
                "IgstVal": round(tot_igst, 2),
                "CesVal": 0.0,
                "StCesVal": 0.0,
                "RndOffAmt": 0.0,
                "TotInvVal": round(exact_calc_total, 2)
            }
        }

    def generate_official_nic_v101_excel_bytes(data_list):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.views.sheetView[0].showGridLines = True

        c_supply = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        c_doc = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
        c_buyer = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
        c_disp = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
        c_ship = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
        c_item = PatternFill(start_color="D8E4BC", end_color="D8E4BC", fill_type="solid")
        c_bch = PatternFill(start_color="D8E4BC", end_color="D8E4BC", fill_type="solid")
        c_val = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
        c_export = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
        c_eway = PatternFill(start_color="E6B8B7", end_color="E6B8B7", fill_type="solid")

        font_title = Font(name="Arial", size=18, bold=True)
        font_sec = Font(name="Arial", size=10, bold=True)
        font_col = Font(name="Arial", size=9, bold=True)
        font_data = Font(name="Arial", size=9)
        font_code = Font(name="Arial", size=8, color="555555")

        thin_border = Border(
            left=Side(style='thin', color='B0B0B0'),
            right=Side(style='thin', color='B0B0B0'),
            top=Side(style='thin', color='B0B0B0'),
            bottom=Side(style='thin', color='B0B0B0')
        )

        ws.merge_cells("B1:Q1")
        ws["B1"] = "E-Invoice System "
        ws["B1"].font = font_title
        ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
        ws["R1"] = "*  Indicates Mandatory Fields"
        ws["R1"].font = Font(name="Arial", size=9, italic=True, color="FF0000")

        sections = [
            ("Document Details", 5, 7, c_doc),
            ("Buyer Details", 8, 18, c_buyer),
            ("Dispatch Details", 19, 24, c_disp),
            ("Shipping Details", 25, 32, c_ship),
            ("Product Details", 33, 57, c_item),
            ("Batch Details", 58, 60, c_bch),
            ("Value Details", 61, 70, c_val),
            ("Export Details", 71, 77, c_export),
            ("E-way-bill Details", 78, 85, c_eway)
        ]

        for name, sc, ec, fill in sections:
            ws.merge_cells(start_row=2, start_column=sc, end_row=2, end_column=ec)
            c = ws.cell(row=2, column=sc, value=name)
            c.fill = fill
            c.font = font_sec
            c.alignment = Alignment(horizontal="center", vertical="center")
            for i in range(sc, ec + 1):
                ws.cell(row=2, column=i).border = thin_border

        col_names = [
            "Supply Type Code *", "Reverse Charge", "e-Comm GSTIN", "Igst On Intra",
            "Document Type *", "Document Number *", "Document Date (DD/MM/YYYY) *",
            "Buyer GSTIN *", "Buyer Legal Name *", "Buyer Trade Name ", "Buyer POS *",
            "Buyer Addr1 *", "Buyer Addr2", "Buyer Location *", "Buyer Pin Code",
            "Buyer State *", "Buyer Phone Number", "Buyer Email Id",
            "Dispatch Name ", "Dispatch Addr1 ", "Dispatch Addr2", "Dispatch Location ", "Dispatch Pin Code ", "Dispatch State ",
            "Shipping GSTIN ", "Shipping Legal Name ", "Shipping Trade Name ", "Shipping Addr1 ", "Shipping Addr2", "Shipping Location ", "Shipping Pin Code ", "Shipping State ",
            "Sl.No. *", "Product Description", "Is_Service *", "HSN code *", "Bar code ", "Quantity *", "Free Quantity", "Unit *", "Unit Price *",
            "Gross Amount *", "Discount", "Pre Tax Value", "Taxable value *", "GST Rate (%) *", "Sgst Amt(Rs)", "Cgst Amt (Rs)",
            "Igst Amt (Rs)", "Cess Rate (%)", "Cess Amt Adval (Rs)", "Cess Non Adval Amt (Rs)", "State Cess Rate (%)",
            "State Cess Adval Amt (Rs)", "State Cess Non-Adval Amt (Rs)", "Other Charges  ", "Item Total *",
            "Batch Name", "Batch Expiry Dt", "Warranty Dt",
            "Total Taxable value *", "Sgst Amt", "Cgst Amt", "Igst Amt", "Cess Amt", "State Cess Amt", "Discount", "Other charges", "Round off", "Total Invoice value *",
            "Shipping Bill No", "Shipping Bill Dt", "Port", "Refund claim", "Foreign Currency ", "Country Code ", "Export Duty Amount",
            "Trans ID", "Trans Name", "Trans Mode ", "Distance ", "Trans Doc No.", "Trans Doc Date", "Vehicle No.", "Vehicle Type", "Error List"
        ]

        col_keys = [
            "colSupType", "colRevCharge", "colEcomGstin", "colIgstIntra",
            "colDoctype", "colDocno", "colDocdate",
            "colBgstin", "colBLegalname", "colBTradname", "colPos",
            "colBaddr1", "colBaddr2", "colBLoc", "colBPin",
            "colBState", "colBPhno", "colBEmail",
            "colDName", "colDaddr1", "colDaddr2", "colDLoc", "colDPin", "colDState",
            "colSgstin", "colSLegalname", "colSTradname", "colSaddr1", "colSaddr2", "colSLoc", "colSPin", "colSState",
            "colProdSlno", "colProddesc", "colProdservice", "colHsn", "colBar", "colQuantity", "colFreeQuanty", "colUnit", "colUnitPrice",
            "colTotal", "colDiscount", "colPreTaxValue", "colAssValue", "colGstrate", "colSgst", "colCgst",
            "colIgst", "colCessrate", "colCessadval", "colCessnonad", "colStCessrate",
            "colStCessadval", "colStCessnonad", "colOthChrgs", "colTolitemval",
            "colBchname", "colBchExpDt", "colBchWarDt",
            "colTotTaxval", "colTsgstval", "colTcgstval", "colTigstval", "colTcessval", "colTstcessval", "colTDiscount", "colTOthChrgs", "colRoundOff", "colTinvoiceval",
            "colShipBilNo", "colShipBilDt", "colPort", "colSupRefund", "colForCur", "colCntryCode", "colExpDty",
            "colTid", "colTName", "colTMode", "colTDistance", "colTDocNo", "colTDocDt", "colTVehno", "colTVehTyp", ""
        ]

        for c_idx, col_name in enumerate(col_names, 1):
            cell_h = ws.cell(row=3, column=c_idx, value=col_name)
            cell_k = ws.cell(row=4, column=c_idx, value=col_keys[c_idx-1])

            if c_idx <= 4:
                fill_target = c_supply
            elif c_idx <= 7:
                fill_target = c_doc
            elif c_idx <= 18:
                fill_target = c_buyer
            elif c_idx <= 24:
                fill_target = c_disp
            elif c_idx <= 32:
                fill_target = c_ship
            elif c_idx <= 57:
                fill_target = c_item
            elif c_idx <= 60:
                fill_target = c_bch
            elif 61 <= c_idx <= 70:
                fill_target = c_val
            elif 71 <= c_idx <= 77:
                fill_target = c_export
            elif 78 <= c_idx <= 85:
                fill_target = c_eway
            else:
                fill_target = PatternFill(fill_type=None)

            cell_h.fill = fill_target
            cell_h.font = font_col
            cell_h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell_h.border = thin_border

            cell_k.font = font_code
            cell_k.alignment = Alignment(horizontal="center", vertical="center")
            cell_k.border = thin_border

        ws.row_dimensions[3].height = 36
        ws.row_dimensions[4].height = 18

        curr_row = 5
        for inv in data_list:
            is_interstate = inv["is_interstate"]
            pos_str = inv["pos_state_name"].title()

            item_rows_data = []
            tot_taxable = 0.0
            tot_cgst = 0.0
            tot_sgst = 0.0
            tot_igst = 0.0
            tot_discount = 0.0

            for it in inv["line_items"]:
                qty = it["qty"]
                price = it["unit_price"]
                disc = it.get("discount", 0.0)

                taxable = it.get("taxable_val", round(qty * price, 2))
                gross_amt = round(taxable + disc, 2)
                gst_rate = it.get("gst_rate", 5.0)

                tot_taxable += taxable
                tot_discount += disc

                if is_interstate:
                    igst = round(taxable * (gst_rate / 100), 2)
                    cgst = 0.0
                    sgst = 0.0
                    tot_igst += igst
                else:
                    cgst = round(taxable * (gst_rate / 200), 2)
                    sgst = round(taxable * (gst_rate / 200), 2)
                    igst = 0.0
                    tot_cgst += cgst
                    tot_sgst += sgst

                item_tot = it.get("item_tot") or round(taxable + cgst + sgst + igst, 2)
                item_rows_data.append({
                    "it": it, "qty": qty, "price": price, "gross_amt": gross_amt,
                    "disc": disc, "taxable": taxable, "gst_rate": gst_rate,
                    "cgst": cgst, "sgst": sgst, "igst": igst, "item_tot": item_tot
                })

            final_taxable_sum = inv.get("printed_taxable_total") or round(tot_taxable, 2)
            final_tax_sum = inv.get("printed_tax_total") or round(tot_cgst + tot_sgst + tot_igst, 2)
            final_invoice_val = inv.get("printed_grand_total") or round(final_taxable_sum + final_tax_sum, 2)

            if inv.get("shipping_is_different", False):
                ship_vals = [
                    str(inv.get("ship_gstin", "")),
                    str(inv.get("ship_name", "")),
                    str(inv.get("ship_name", "")),
                    str(inv.get("ship_addr1", "")),
                    str(inv.get("ship_addr2", "")),
                    str(inv.get("ship_loc", "")),
                    str(inv.get("ship_pin", "")),
                    str(inv.get("ship_state_name", ""))
                ]
            else:
                ship_vals = ["", "", "", "", "", "", "", ""]

            for s_no, r_data in enumerate(item_rows_data, 1):
                it = r_data["it"]
                row_data = [
                    "B2B", "No", "", "",
                    "Tax Invoice", str(inv["invoice_no"]), str(inv["doc_date"]),
                    str(inv["buyer_gstin"]), str(inv["buyer_name"]), "", pos_str,
                    str(inv.get("buyer_addr1", "")), str(inv.get("buyer_addr2", "")), str(inv["buyer_loc"]), str(inv["buyer_pin"]),
                    str(inv["buyer_state_name"]), str(inv["buyer_phone"]), str(inv["buyer_email"]),
                    "", "", "", "", "", "",
                    *ship_vals,
                    str(s_no), str(it["desc"]), "No", str(it["hsn"]), "", str(r_data["qty"]), "",
                    "BOX" if str(it.get("unit", "")).upper() == "BOX" else "PACKS",
                    fmt_dec(r_data["price"]), fmt_dec(r_data["gross_amt"]),
                    fmt_dec(r_data["disc"]) if r_data["disc"] > 0 else "",
                    "",
                    fmt_dec(r_data["taxable"]), str(int(r_data["gst_rate"])),
                    fmt_dec(r_data["sgst"]) if (r_data["sgst"] > 0 and not is_interstate) else "",
                    fmt_dec(r_data["cgst"]) if (r_data["cgst"] > 0 and not is_interstate) else "",
                    fmt_dec(r_data["igst"]) if (r_data["igst"] > 0 and is_interstate) else "",
                    "", "", "", "", "", "", "",
                    fmt_dec(r_data["item_tot"]),
                    "", "", "",
                    fmt_dec(final_taxable_sum),
                    fmt_dec(tot_sgst) if (tot_sgst > 0 and not is_interstate) else "",
                    fmt_dec(tot_cgst) if (tot_cgst > 0 and not is_interstate) else "",
                    fmt_dec(final_tax_sum) if is_interstate else "",
                    "", "", "", "",
                    "",
                    fmt_dec(final_invoice_val),
                    "", "", "", "", "", "", "",
                    "", "", "", "", "", "", "", "",
                    ""
                ]

                for c_idx, val in enumerate(row_data, 1):
                    cell = ws.cell(row=curr_row, column=c_idx, value=str(val) if val != "" else None)
                    cell.font = font_data
                    cell.border = thin_border
                    if cell.value is not None:
                        cell.number_format = '@'

                    if c_idx in [1, 2, 5, 7, 8, 11, 15, 16, 25, 31, 32, 33, 35, 36, 38, 40, 46]:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif c_idx in [41, 42, 43, 45, 47, 48, 49, 57, 61, 62, 63, 64, 70]:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

                curr_row += 1

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row < 3:
                    continue
                v_str = str(cell.value or '')
                if len(v_str) > max_len:
                    max_len = len(v_str)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

        out_io = io.BytesIO()
        wb.save(out_io)
        return out_io.getvalue()

    def render_exact_government_einvoice_preview(meta):
        is_interstate = meta["is_interstate"]
        tot_taxable = meta.get("printed_taxable_total") or sum([it['taxable_val'] for it in meta['line_items']])
        tot_discount = sum([it.get('discount', 0.0) for it in meta['line_items']])

        if is_interstate:
            tot_igst = meta.get("printed_tax_total") or round(tot_taxable * 0.05, 2)
            tot_cgst = 0.0
            tot_sgst = 0.0
        else:
            tot_igst = 0.0
            tot_cgst = round(tot_taxable * 0.025, 2)
            tot_sgst = round(tot_taxable * 0.025, 2)

        exact_total = meta.get("printed_grand_total") or round(tot_taxable + tot_cgst + tot_sgst + tot_igst, 2)
        ack_date_str = f"{meta['clean_date']} 17:00:00"

        table_rows = ""
        for idx, it in enumerate(meta['line_items'], 1):
            q = it['qty']
            p = it['unit_price']
            taxable = it['taxable_val']
            disc = it.get('discount', 0.0)
            gst_r = it.get('gst_rate', 5.0)
            tot_val = it.get("item_tot") or round(taxable + round(taxable * (gst_r / 100), 2), 2)

            table_rows += f"""
            <tr style="border-bottom: 1px solid #000; font-size: 11px;">
                <td style="border-right: 1px solid #000; padding: 5px 3px; text-align: center;">{idx}</td>
                <td style="border-right: 1px solid #000; padding: 5px 4px; text-align: left; font-weight: 500;">{it['desc']}</td>
                <td style="border-right: 1px solid #000; padding: 5px 3px; text-align: center;">{it['hsn']}</td>
                <td style="border-right: 1px solid #000; padding: 5px 3px; text-align: right;">{q}</td>
                <td style="border-right: 1px solid #000; padding: 5px 3px; text-align: center;">{it.get('unit', 'PAC')}</td>
                <td style="border-right: 1px solid #000; padding: 5px 3px; text-align: right;">{p:.2f}</td>
                <td style="border-right: 1px solid #000; padding: 5px 3px; text-align: right;">{disc:.2f}</td>
                <td style="border-right: 1px solid #000; padding: 5px 3px; text-align: right;">{taxable:.2f}</td>
                <td style="border-right: 1px solid #000; padding: 5px 3px; text-align: center; line-height: 1.2;">
                    {gst_r:.2f}+0.00<br><span style="font-size: 9px; color: #555;">0.00+0</span>
                </td>
                <td style="border-right: 1px solid #000; padding: 5px 3px; text-align: right;">0</td>
                <td style="padding: 5px 4px; text-align: right; font-weight: bold;">{tot_val:.2f}</td>
            </tr>
            """

        html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
body {{ background-color: #f1f5f9; font-family: Arial, Helvetica, sans-serif; margin: 0; padding: 10px; }}
.invoice-box {{ background-color: #ffffff; color: #000000; padding: 16px; border: 2px solid #000; max-width: 930px; margin: 0 auto; box-shadow: 0 4px 10px rgba(0,0,0,0.15); }}
.sec-title {{ background-color: #d1d5db; color: #000; font-weight: bold; font-size: 11.5px; padding: 3px 6px; border: 1px solid #000; }}
.sec-content {{ border: 1px solid #000; border-top: none; padding: 6px; font-size: 11px; margin-bottom: 8px; line-height: 1.5; }}
table.items {{ width: 100%; border-collapse: collapse; border: 1px solid #000; border-top: none; }}
table.items th {{ background-color: #f3f4f6; border-right: 1px solid #000; border-bottom: 1px solid #000; padding: 5px 3px; font-size: 10.5px; }}
</style>
</head>
<body>
<div class="invoice-box">
    <table style="width: 100%; border-collapse: collapse; margin-bottom: 8px;">
        <tr>
            <td style="vertical-align: top; width: 75%;">
                <div style="font-size: 19px; font-weight: bold; letter-spacing: 0.5px;">{meta['seller_gstin']}</div>
                <div style="font-size: 17px; font-weight: bold; margin-top: 4px; text-transform: uppercase;">{meta['seller_name']}</div>
            </td>
            <td style="vertical-align: top; width: 25%; text-align: right;">
                <div style="display: inline-block; width: 95px; height: 95px; border: 1.5px solid #000; text-align: center; background: #fff; padding: 4px;">
                    <svg viewBox="0 0 100 100" style="width: 100%; height: 100%;">
                        <rect width="100" height="100" fill="#fff" />
                        <path d="M10 10 h30 v30 h-30 z M15 15 v20 h20 v-20 z M20 20 h10 v10 h-10 z" fill="#000" />
                        <path d="M60 10 h30 v30 h-30 z M65 15 v20 h20 v-20 z M70 20 h10 v10 h-10 z" fill="#000" />
                        <path d="M10 60 h30 v30 h-30 z M15 65 v20 h20 v-20 z M20 70 h10 v10 h-10 z" fill="#000" />
                        <rect x="50" y="50" width="10" height="10" fill="#000" />
                        <rect x="65" y="65" width="15" height="15" fill="#000" />
                        <rect x="75" y="50" width="10" height="10" fill="#000" />
                        <rect x="50" y="75" width="10" height="10" fill="#000" />
                    </svg>
                </div>
            </td>
        </tr>
    </table>
    <div class="sec-title">1. e-Invoice Details</div>
    <div class="sec-content">
        <table style="width: 100%;">
            <tr>
                <td style="width: 50%;"><b>IRN :</b> <span style="font-size: 9.5px; word-break: break-all;">3521a723ac0d702f87a9ee33b47e4f25eb9b8ac3e0d2160434014ee0c0102180</span></td>
                <td style="width: 25%;"><b>Ack. No :</b> 172621231914553</td>
                <td style="width: 25%; text-align: right;"><b>Ack. Date :</b> {ack_date_str}</td>
            </tr>
        </table>
    </div>
    <div class="sec-title">2. Transaction Details</div>
    <div class="sec-content" style="line-height: 1.6;">
        <table style="width: 100%;">
            <tr>
                <td style="width: 32%;"><b>Supply Type Code :</b> B2B</td>
                <td style="width: 35%;"><b>Document No :</b> {meta['invoice_no']}</td>
                <td style="width: 33%;" rowspan="2"><b>IGST applicable despite Supplier and Recipient located in same State :</b> No</td>
            </tr>
            <tr>
                <td><b>Place of Supply :</b> {meta['pos_state_name'].upper()}</td>
                <td></td>
            </tr>
            <tr>
                <td><b>Document Type :</b> Tax Invoice</td>
                <td><b>Document Date :</b> {meta['clean_date']}</td>
                <td></td>
            </tr>
        </table>
    </div>
    <div class="sec-title">3. Party Details</div>
    <div style="border: 1px solid #000; border-top: none; margin-bottom: 8px;">
        <table style="width: 100%; border-collapse: collapse; font-size: 11px;">
            <tr>
                <td style="width: 50%; border-right: 1px solid #000; padding: 6px; vertical-align: top;">
                    <div style="font-weight: bold; font-size: 13px; text-decoration: underline; margin-bottom: 3px;">Supplier</div>
                    <div><b>GSTIN :</b> {meta['seller_gstin']}</div>
                    <div style="font-weight: bold; margin-top: 2px;">{meta['seller_name']}</div>
                    <div>{meta['seller_addr1']}</div>
                    <div>{meta['seller_loc'].upper()}</div>
                    <div>{meta['seller_pin']} &nbsp; {meta['seller_state_name'].upper()}</div>
                    <div style="margin-top: 3px;">+91-7070701513 &nbsp; info@romsons.in</div>
                </td>
                <td style="width: 50%; padding: 6px; vertical-align: top;">
                    <div style="font-weight: bold; font-size: 13px; text-decoration: underline; margin-bottom: 3px;">Recipient</div>
                    <div><b>GSTIN :</b> {meta['buyer_gstin']}</div>
                    <div style="font-weight: bold; margin-top: 2px;">{meta['buyer_name']}</div>
                    <div>{meta['buyer_addr1']} {meta.get('buyer_addr2', '')}</div>
                    <div>{meta['buyer_loc']} &nbsp; Place of Supply: {meta['pos_state_name'].upper()}</div>
                    <div>{meta['buyer_pin']} &nbsp; {meta['buyer_state_name'].upper()}</div>
                    <div style="margin-top: 3px;">{meta['buyer_phone']} &nbsp; {meta['buyer_email']}</div>
                </td>
            </tr>
        </table>
    </div>
    <div class="sec-title">4. Details of Goods / Services</div>
    <table class="items">
        <thead>
            <tr>
                <th style="width: 4%;">SlNo</th>
                <th style="width: 28%; text-align: left;">Item Description</th>
                <th style="width: 8%;">HSN Code</th>
                <th style="width: 6%;">Quantity</th>
                <th style="width: 5%;">Unit</th>
                <th style="width: 8%;">Unit Price(Rs)</th>
                <th style="width: 7%;">Discount(Rs)</th>
                <th style="width: 9%;">Taxable Amount(Rs)</th>
                <th style="width: 14%; line-height: 1.2;">Tax Rate<br><span style="font-size: 9px; font-weight: normal;">(GST+Cess)</span></th>
                <th style="width: 6%;">Other charges(Rs)</th>
                <th style="width: 9%; border-right: none;">Total</th>
            </tr>
        </thead>
        <tbody>
            {table_rows}
        </tbody>
    </table>
    <table style="width: 100%; border-collapse: collapse; border: 1px solid #000; border-top: none; font-size: 11px; margin-top: -1px;">
        <tr>
            <td style="width: 60%; padding: 6px; vertical-align: top; border-right: 1px solid #000; font-size: 10px; color: #444;">
                * Official verification format. e-Invoice generated via PAN Operations Portal.
            </td>
            <td style="width: 40%; padding: 6px; vertical-align: top;">
                <table style="width: 100%; line-height: 1.6;">
                    <tr><td>Total Taxable Value :</td><td style="text-align: right; font-weight: bold;">₹{tot_taxable:.2f}</td></tr>
                    <tr><td>Total Discount :</td><td style="text-align: right; color: #dc2626;">₹{tot_discount:.2f}</td></tr>
                    <tr><td>Total CGST :</td><td style="text-align: right;">₹{tot_cgst:.2f}</td></tr>
                    <tr><td>Total SGST :</td><td style="text-align: right;">₹{tot_sgst:.2f}</td></tr>
                    <tr><td>Total IGST :</td><td style="text-align: right;">₹{tot_igst:.2f}</td></tr>
                    <tr style="border-top: 1.5px solid #000; font-size: 13px; font-weight: bold;"><td>Total Invoice Value :</td><td style="text-align: right;">₹{exact_total:.2f}</td></tr>
                </table>
            </td>
        </tr>
    </table>
</div>
</body>
</html>
"""
        return html

    uploaded_invoices = st.file_uploader(
        "Upload Invoices (PDF) - Single ya Bulk", 
        type=["pdf"], 
        accept_multiple_files=True, 
        key=f"blnk_inv_{st.session_state.selected_wh}_{st.session_state.uploader_key}"
    )

    if uploaded_invoices:
        processed_docs = []
        for file in uploaded_invoices:
            try:
                file.seek(0)
                pdf_bytes = file.read()
                reconciled_pdf_buf, meta = process_and_reconcile_pdf(pdf_bytes)
                processed_docs.append({
                    "original_name": file.name,
                    "pdf_buf": reconciled_pdf_buf,
                    "meta": meta
                })
            except Exception as e:
                st.error(f"Error processing {file.name}: {str(e)}")

        if processed_docs:
            st.session_state.blnk_batch_results = processed_docs

    if st.session_state.blnk_batch_results:
        active_docs = st.session_state.blnk_batch_results
        st.success(f"✅ Successfully processed {len(active_docs)} Invoices!")
        parsed_invoices = [doc["meta"] for doc in active_docs]
        excel_raw_bytes = generate_official_nic_v101_excel_bytes(parsed_invoices)

        if len(active_docs) == 1:
            doc = active_docs[0]
            meta = doc["meta"]
            pdf_buf = doc["pdf_buf"]
            json_payload = build_einvoice_json_v101(meta)
            json_raw_bytes = json.dumps(json_payload, indent=4).encode('utf-8')
            inv_no = meta["invoice_no"]
            pdf_name = meta["pdf_filename"]
            is_blinkit = meta["is_blinkit"]

            tax_badge = "🔴 Inter-State (IGST Only)" if meta["is_interstate"] else "🟢 Intra-State (CGST + SGST)"
            ship_status = "⚠️ Different (Shipping Details Filled in Excel)" if meta["shipping_is_different"] else "✅ Same (Shipping Blank in Excel)"

            st.markdown(f"""
                <div style="background-color: #1e293b; padding: 14px 18px; border-radius: 8px; margin-bottom: 15px; border-left: 5px solid {'#10b981' if is_blinkit else '#3b82f6'};">
                    <b>Buyer Name:</b> <code>{meta['buyer_name']}</code> &nbsp;|&nbsp; <b>Tax Rule:</b> <code>{tax_badge}</code><br>
                    <b>Address 1:</b> <code>{meta['buyer_addr1']}</code> &nbsp;|&nbsp; <b>Address 2:</b> <code>{meta.get('buyer_addr2', '')}</code><br>
                    <b>City:</b> <code>{meta['buyer_loc']}</code> ({meta['buyer_pin']}) &nbsp;|&nbsp; <b>Shipping Status:</b> <code>{ship_status}</code><br>
                    <b>Seller:</b> {meta['seller_name']} ({meta['seller_gstin']})
                </div>
            """, unsafe_allow_html=True)

            hist_item = {
                "original_name": doc["original_name"],
                "file_name": pdf_name,
                "data": pdf_buf.getvalue(),
                "excel_data": excel_raw_bytes,
                "json_data": json_raw_bytes,
                "pages": 1,
                "stamped": len(meta["line_items"]),
                "location": st.session_state.selected_wh,
                "module": "BLNK"
            }

            c1, c2, c3 = st.columns(3)
            with c1:
                st.download_button(
                    label="📥 Download Processed PDF",
                    data=pdf_buf.getvalue(),
                    file_name=pdf_name,
                    mime="application/pdf",
                    on_click=save_item_to_disk_history,
                    args=(hist_item,),
                    key="dl_blnk_single_pdf"
                )
            with c2:
                st.download_button(
                    label="📊 Download eInvoice.xlsx (Official Format)",
                    data=excel_raw_bytes,
                    file_name=f"{inv_no}_eInvoice.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    on_click=save_item_to_disk_history,
                    args=(hist_item,),
                    key="dl_blnk_single_xlsx"
                )
            with c3:
                st.download_button(
                    label="🧾 Download e-Invoice JSON (v1.01)",
                    data=json_raw_bytes,
                    file_name=f"{inv_no}_eInvoice_v1.01.json",
                    mime="application/json",
                    on_click=save_item_to_disk_history,
                    args=(hist_item,),
                    key="dl_blnk_single_json"
                )

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("### 📄 Official Government e-Invoice Live Preview")
            calc_height = 560 + (len(meta['line_items']) * 36)
            preview_html = render_exact_government_einvoice_preview(meta)
            components.html(preview_html, height=calc_height, scrolling=True)

        else:
            zip_buffer = io.BytesIO()
            batch_hist = []
            bulk_json = [build_einvoice_json_v101(d["meta"]) for d in active_docs]
            date_str = datetime.now(IST).strftime('%d-%m-%Y')

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, doc in enumerate(active_docs, 1):
                    pdf_name = doc["meta"]["pdf_filename"]
                    name_root = pdf_name[:-4] if pdf_name.endswith(".pdf") else pdf_name
                    raw_data = doc["pdf_buf"].getvalue()
                    
                    zip_file.writestr(pdf_name, raw_data)

                    single_json = bulk_json[idx-1]
                    single_json_bytes = json.dumps(single_json, indent=4).encode('utf-8')
                    single_excel_bytes = generate_official_nic_v101_excel_bytes([doc["meta"]])

                    zip_file.writestr(f"{name_root}.xlsx", single_excel_bytes)
                    zip_file.writestr(f"{name_root}.json", single_json_bytes)

                    batch_hist.append({
                        "original_name": doc["original_name"],
                        "file_name": pdf_name,
                        "data": raw_data,
                        "excel_data": single_excel_bytes,
                        "json_data": single_json_bytes,
                        "pages": 1,
                        "stamped": len(doc["meta"]["line_items"]),
                        "location": st.session_state.selected_wh,
                        "module": "BLNK"
                    })
            zip_buffer.seek(0)

            c1, c2, c3 = st.columns(3)
            with c1:
                st.download_button(
                    label=f"📦 Download All Processed Files ({len(active_docs)} Invoices Bundle - ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Invoices_Complete_Bundle_{date_str}.zip",
                    mime="application/zip",
                    on_click=save_batch_to_disk_history,
                    args=(batch_hist,),
                    key="dl_blnk_bulk_zip"
                )
            with c2:
                st.download_button(
                    label=f"📊 Download Bulk eInvoice.xlsx ({len(active_docs)} Invoices)",
                    data=excel_raw_bytes,
                    file_name=f"Bulk_eInvoice_{date_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    on_click=save_batch_to_disk_history,
                    args=(batch_hist,),
                    key="dl_blnk_bulk_xlsx"
                )
            with c3:
                st.download_button(
                    label=f"🧾 Download Bulk e-Invoice JSON ({len(bulk_json)} Invoices)",
                    data=json.dumps(bulk_json, indent=4).encode('utf-8'),
                    file_name=f"Bulk_eInvoice_NIC_v1.01_{date_str}.json",
                    mime="application/json",
                    on_click=save_batch_to_disk_history,
                    args=(batch_hist,),
                    key="dl_blnk_bulk_json"
                )

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("### 📄 Invoices Official e-Invoice Preview")
            preview_tabs = st.tabs([f"INV: {d['meta']['invoice_no']}" for d in processed_docs])
            for i, tab in enumerate(preview_tabs):
                with tab:
                    curr_meta = active_docs[i]['meta']
                    calc_height = 560 + (len(curr_meta['line_items']) * 36)
                    p_html = render_exact_government_einvoice_preview(curr_meta)
                    components.html(p_html, height=calc_height, scrolling=True)

# ----------------- MODULE 4: LABEL EDITOR -----------------
elif st.session_state.active_tab == "LBL":
    st.subheader("🏷️ Shipping & Product Label Editor")
    st.info("No files")

# ----------------- MODULE 5: SETTINGS & REPORTS -----------------
elif st.session_state.active_tab == "CONF":
    st.subheader("⚙️ System Operations & Units Configuration")
    st.markdown(f"""
    - **Active Location:** `{LOCATION_DISPLAY_MAP.get(st.session_state.selected_wh, st.session_state.selected_wh)}`
    - **Default Barcode Type:** Code128 (High DPI Crisp Vector)
    - **Engine Status:** Online & Healthy
    """)
