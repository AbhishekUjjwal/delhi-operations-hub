import streamlit as st
import pandas as pd
import pymupdf as fitz
import re
import io
import zipfile
import barcode
from barcode.writer import ImageWriter
from datetime import datetime

# Page Configuration
st.set_page_config(
    page_title="PAN OPERATIONS HUB | Central Portal",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling: Strict Letter-Fit Tabs & Clean Icons
st.markdown("""
    <style>
    /* 1. HIDE TOP-RIGHT FORK, GITHUB LOGO & STREAMLIT HEADER */
    #MainMenu {visibility: hidden; display: none !important;}
    header {visibility: hidden; display: none !important;}
    footer {visibility: hidden; display: none !important;}
    .stDeployButton {display: none !important;}
    div[data-testid="stToolbar"] {visibility: hidden; display: none !important;}
    div[data-testid="stDecoration"] {visibility: hidden; display: none !important;}
    div[data-testid="stStatusWidget"] {visibility: hidden; display: none !important;}

    /* 2. TOP ACTION ROW - COMPACT HORIZONTAL BAR */
    div[data-testid="stHorizontalBlock"] {
        align-items: center !important;
        gap: 6px !important;
    }

    /* 3. TIGHT LETTER-FIT BUTTONS */
    div[data-testid="stButton"] {
        width: fit-content !important;
    }
    div[data-testid="stButton"] button {
        border-radius: 16px !important;
        font-size: 12px !important;
        font-weight: 600 !important;
        padding: 4px 10px !important;
        min-height: 28px !important;
        height: 28px !important;
        line-height: 1 !important;
        border: 1px solid rgba(148, 163, 184, 0.3) !important;
        white-space: nowrap !important;
        width: fit-content !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-shadow: none !important;
    }
    div[data-testid="stButton"] button:hover {
        border-color: #38bdf8 !important;
        color: #38bdf8 !important;
    }

    /* 4. COMPACT SELECTBOX - FULL TEXT VISIBLE */
    div[data-testid="stSelectbox"] label {
        display: none !important;
    }
    div[data-testid="stSelectbox"] {
        min-width: 125px !important;
        width: fit-content !important;
    }
    div[data-testid="stSelectbox"] > div > div {
        min-height: 28px !important;
        height: 28px !important;
        padding: 0px 8px !important;
        font-size: 12px !important;
        border-radius: 16px !important;
        border: 1px solid rgba(148, 163, 184, 0.3) !important;
    }

    /* 5. SIDEBAR TITLE */
    .hub-title {
        font-size: 18px;
        font-weight: 800;
        white-space: nowrap;
        background: linear-gradient(90deg, #38bdf8 0%, #818cf8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 2px;
        display: flex;
        align-items: center;
        gap: 6px;
    }

    .pro-success-card {
        background: linear-gradient(135deg, rgba(16, 185, 129, 0.15) 0%, rgba(5, 150, 105, 0.25) 100%);
        border: 1px solid rgba(52, 211, 153, 0.5);
        border-radius: 10px;
        padding: 14px 18px;
        color: #ecfdf5;
        margin: 12px 0px;
    }
    </style>
""", unsafe_allow_html=True)

# Locations list without RPPL
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
    "AMZ": {"name": "Amazon Invoice Editor", "code": "AMZ-ED", "logo": "📦"},
    "BLNK": {"name": "Blinkit e-Invoice Tool", "code": "BLNK-INV", "logo": "⚡"},
    "LBL": {"name": "Label Editor", "code": "LBL-ED", "logo": "🏷️"},
    "CONF": {"name": "Settings & Reports", "code": "SETTINGS", "logo": "⚙️"}
}

# ----------------- SESSION STATE MANAGEMENT -----------------
if "selected_wh" not in st.session_state:
    st.session_state.selected_wh = "DEL"
if "open_tabs" not in st.session_state:
    st.session_state.open_tabs = ["AMZ"]
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "AMZ"
if "permanent_history" not in st.session_state:
    st.session_state.permanent_history = []
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "show_history" not in st.session_state:
    st.session_state.show_history = False

def get_current_location_history():
    if st.session_state.selected_wh == "Headoffice":
        return st.session_state.permanent_history
    return [
        item for item in st.session_state.permanent_history 
        if item.get("location") == st.session_state.selected_wh
    ]

def reset_current_workspace():
    st.session_state.uploader_key += 1
    st.session_state.show_history = False
    st.rerun()

# ----------------- SIDEBAR NAVIGATION -----------------
with st.sidebar:
    st.markdown('<div class="hub-title">🏛️ PAN OPERATIONS HUB</div>', unsafe_allow_html=True)
    st.caption("Romsons prime pvt ltd operation portal")
    st.divider()

    for key, info in AVAILABLE_MODULES.items():
        is_pinned = key in st.session_state.open_tabs
        btn_label = f"📌 {info['logo']} {info['code']}" if is_pinned else f"➕ {info['logo']} {info['code']}"
        if st.button(btn_label, key=f"side_nav_{key}", use_container_width=True):
            if key not in st.session_state.open_tabs:
                st.session_state.open_tabs.append(key)
            st.session_state.active_tab = key
            st.session_state.show_history = False
            st.rerun()

    st.divider()
    st.caption("System Version: v3.2 Enterprise")

# ----------------- TOP ACTION BAR: CLEAN COMPACT ROW -----------------
top_cols = st.columns([0.8] * len(st.session_state.open_tabs) + [1.1, 0.9, 0.65, 3.5])

col_idx = 0

# 1. Pinned Tabs
for t_key in st.session_state.open_tabs:
    info = AVAILABLE_MODULES[t_key]
    is_active = (t_key == st.session_state.active_tab and not st.session_state.show_history)
    
    if len(st.session_state.open_tabs) > 1:
        tab_label = f"{info['logo']} {info['code']} ✕"
    else:
        tab_label = f"{info['logo']} {info['code']}"

    with top_cols[col_idx]:
        if st.button(tab_label, key=f"tab_chip_{t_key}", type="primary" if is_active else "secondary"):
            if is_active and len(st.session_state.open_tabs) > 1:
                st.session_state.open_tabs.remove(t_key)
                st.session_state.active_tab = st.session_state.open_tabs[0]
            else:
                st.session_state.active_tab = t_key
                st.session_state.show_history = False
            st.rerun()
    col_idx += 1

# 2. Location Dropdown (RPPL Removed & Full Visible)
with top_cols[col_idx]:
    selected_display = st.selectbox(
        "WH | Location",
        options=ORDERED_LOCATIONS,
        format_func=lambda x: LOCATION_DISPLAY_MAP.get(x, x),
        index=ORDERED_LOCATIONS.index(st.session_state.selected_wh) if st.session_state.selected_wh in ORDERED_LOCATIONS else 0
    )
    st.session_state.selected_wh = selected_display
col_idx += 1

# 3. History Button
loc_history = get_current_location_history()
with top_cols[col_idx]:
    if st.button(f"🕒 History ({len(loc_history)})", key="top_hist_btn", type="primary" if st.session_state.show_history else "secondary"):
        st.session_state.show_history = not st.session_state.show_history
        st.rerun()
col_idx += 1

# 4. Reset Button
with top_cols[col_idx]:
    if st.button("🔄 Reset", key="top_reset_btn", type="secondary", help="Reset upload area"):
        reset_current_workspace()

st.divider()

# =========================================================================
# HISTORY VIEW (PERMANENT STORAGE)
# =========================================================================
if st.session_state.show_history:
    h_top_c1, h_top_c2 = st.columns([4, 1])
    with h_top_c1:
        st.subheader("📜 Processed History & Downloads (Safe Storage)")
    with h_top_c2:
        if loc_history:
            if st.button("🗑️ Clear History", key="btn_del_hist", use_container_width=True):
                if st.session_state.selected_wh == "Headoffice":
                    st.session_state.permanent_history = []
                else:
                    st.session_state.permanent_history = [
                        item for item in st.session_state.permanent_history 
                        if item.get("location") != st.session_state.selected_wh
                    ]
                st.rerun()

    if not loc_history:
        st.info(f"Abhi tak **{st.session_state.selected_wh}** ke liye koi history saved nahi hai.")
    else:
        st.success(f"💾 Total Saved Files in Storage: **{len(loc_history)}**")
        
        if len(loc_history) > 1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, item in enumerate(loc_history):
                    zip_file.writestr(f"{idx+1}_{item['file_name']}", item["data"])
            zip_buffer.seek(0)

            st.download_button(
                label=f"📦 Download All {st.session_state.selected_wh} Files as ZIP ({len(loc_history)} Files)",
                data=zip_buffer,
                file_name=f"{st.session_state.selected_wh}_History.zip",
                mime="application/zip",
                key="btn_zip_history_download",
                use_container_width=True
            )

        st.write("---")
        for idx, item in enumerate(loc_history):
            d_col1, d_col2 = st.columns([3, 1])
            with d_col1:
                st.write(
                    f"📁 **{item['original_name']}** — `[{LOCATION_DISPLAY_MAP.get(item.get('location'), item.get('location'))}]` | "
                    f"`{item['pages']} Pages Kept` | `{item['stamped']} Stamped` | "
                    f"`{item['removed_pan']} Non-PAN Filtered` | `{item.get('removed_desc', 0)} No-Desc Filtered` | "
                    f"⏰ *{item.get('time', '')}*"
                )
            with d_col2:
                st.download_button(
                    label="📥 Download PDF",
                    data=item["data"],
                    file_name=item["file_name"],
                    mime="application/pdf",
                    key=f"dl_hist_perm_{st.session_state.selected_wh}_{idx}",
                    use_container_width=True
                )

# =========================================================================
# MODULE 1: AMAZON INVOICE EDITOR
# =========================================================================
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

    def clean_val(v):
        if pd.isna(v):
            return ""
        s = str(v).strip()
        s = re.sub(r'^[="\']+|["\']+$', '', s)
        return s.strip()

    def clean_alphanumeric(text):
        return re.sub(r'[^a-zA-Z0-9]', '', str(text)).lower()

    def get_token_words(text):
        if not text:
            return set()
        words = re.findall(r'[a-zA-Z0-9]+', str(text).lower())
        stop_words = {
            'the', 'and', 'for', 'with', 'pcs', 'piece', 'pieces', 'only', 'total', 
            'hsn', 'gst', 'rs', 'inr', 'tax', 'invoice', 'date', 'order', 'qty',
            'rate', 'amount', 'shipping', 'charges', 'state', 'code', 'pan', 'bill', 'to'
        }
        return {w for w in words if len(w) >= 2 and w not in stop_words}

    def has_valid_description(raw_text):
        text_lower = raw_text.lower()
        desc_keywords = ["description", "item description", "product description", "title"]
        has_header = any(k in text_lower for k in desc_keywords)
