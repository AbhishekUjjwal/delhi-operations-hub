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

# Custom Styling: Zero Top Padding, Hide Bottom Watermarks & Streamlit Badges
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

    /* 2. STRICTLY HIDE STREAMLIT BOTTOM LOGO, CROWN & BADGES */
    [data-testid="stViewerBadge"] {display: none !important; visibility: hidden !important;}
    div[class*="viewerBadge"] {display: none !important; visibility: hidden !important;}
    div[data-testid="manage-app-button"] {display: none !important; visibility: hidden !important;}
    div[class*="manage-app"] {display: none !important; visibility: hidden !important;}
    div[class*="floating"] {display: none !important; visibility: hidden !important;}
    div[data-testid="stBottom"] {display: none !important;}
    iframe[title*="streamlit"] {display: none !important;}
    footer:after {display: none !important;}
    a[href*="streamlit.io"] {display: none !important;}

    /* 3. ZERO EXTRA TOP GAP */
    .block-container {
        padding-top: 0.2rem !important;
        padding-bottom: 0.8rem !important;
        padding-left: 1.2rem !important;
        padding-right: 1.2rem !important;
    }

    hr {
        margin-top: 4px !important;
        margin-bottom: 6px !important;
        border-color: rgba(58, 80, 107, 0.4) !important;
    }

    /* 4. TIGHT LETTER-FIT BUTTONS */
    div[data-testid="stButton"] {
        width: fit-content !important;
    }
    div[data-testid="stButton"] button {
        border-radius: 16px !important;
        font-size: 12px !important;
        font-weight: 600 !important;
        padding: 3px 10px !important;
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

    /* 5. RIGHT STACK ACTION ROW BUTTONS */
    .right-action-cluster div[data-testid="stButton"] {
        width: 100% !important;
    }
    .right-action-cluster div[data-testid="stButton"] button {
        width: 100% !important;
    }

    /* 6. COMPACT SELECTBOX (LOCATION) */
    div[data-testid="stSelectbox"] label {
        display: none !important;
    }
    div[data-testid="stSelectbox"] {
        width: 100% !important;
    }
    div[data-testid="stSelectbox"] > div > div {
        min-height: 28px !important;
        height: 28px !important;
        padding: 0px 8px !important;
        font-size: 12px !important;
        border-radius: 16px !important;
        border: 1px solid rgba(148, 163, 184, 0.3) !important;
    }

    h2, h3, h4 {
        margin-top: 0px !important;
        margin-bottom: 4px !important;
        padding-top: 0px !important;
    }

    /* 7. SIDEBAR TITLE */
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
        padding: 12px 16px;
        color: #ecfdf5;
        margin: 8px 0px;
    }
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

# Blinkit Invoice se pehle Blinkit PO Uploader add kiya gaya hai
AVAILABLE_MODULES = {
    "AMZ": {"name": "Amazon Invoice Editor", "code": "AMZ-ED", "logo": "📦"},
    "BLNK_PO": {"name": "Blinkit PO Uploader", "code": "BLNK-PO", "logo": "📑"},
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
if "current_batch_results" not in st.session_state:
    st.session_state.current_batch_results = []

def get_current_location_history():
    current_tab = st.session_state.active_tab
    current_wh = st.session_state.selected_wh
    
    module_items = [
        item for item in st.session_state.permanent_history 
        if item.get("module") == current_tab
    ]
    if current_wh == "Headoffice":
        return module_items
        
    return [
        item for item in module_items 
        if item.get("location") == current_wh
    ]

def record_to_history(item):
    existing_ids = {f"{x.get('file_name')}_{x.get('location')}_{x.get('module')}" for x in st.session_state.permanent_history}
    item_id = f"{item.get('file_name')}_{item.get('location')}_{item.get('module')}"
    if item_id not in existing_ids:
        st.session_state.permanent_history.append(item)

def record_batch_to_history(batch_items):
    for item in batch_items:
        record_to_history(item)

def reset_current_workspace():
    st.session_state.uploader_key += 1
    st.session_state.current_batch_results = []
    st.session_state.show_history = False
    st.rerun()

def clear_all_pinned_tabs():
    st.session_state.open_tabs = ["AMZ"]
    st.session_state.active_tab = "AMZ"
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

    st.divider
