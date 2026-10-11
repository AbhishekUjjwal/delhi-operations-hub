import streamlit as st
import pandas as pd
import pymupdf as fitz
import re
import io
import zipfile
import barcode
from barcode.writer import ImageWriter

# Page Configuration - Enterprise Wide Layout
st.set_page_config(
    page_title="PAN OPERATIONS HUB | Central Portal",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling: Sleek, Compact & Theme-Adaptive
st.markdown("""
    <style>
    /* Metric Cards - Modern Gradient */
    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.7) 0%, rgba(15, 23, 42, 0.9) 100%);
        border: 1px solid rgba(56, 189, 248, 0.3);
        border-radius: 10px;
        padding: 10px 14px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    }
    
    /* Top action selectbox label hide & compact size */
    div[data-testid="stSelectbox"] label {
        display: none;
    }
    div[data-testid="stSelectbox"] > div > div {
        padding-top: 2px !important;
        padding-bottom: 2px !important;
        min-height: 34px !important;
        font-size: 13px !important;
        border-radius: 18px !important;
    }
    
    /* Compact Pill buttons */
    div[data-testid="stButton"] button {
        border-radius: 18px !important;
        font-size: 12.5px !important;
        font-weight: 600 !important;
        padding: 3px 12px !important;
        min-height: 34px !important;
        line-height: 1.2 !important;
        transition: all 0.2s ease-in-out !important;
        border: 1px solid rgba(148, 163, 184, 0.25) !important;
    }
    
    div[data-testid="stButton"] button:hover {
        transform: translateY(-1px);
        box-shadow: 0 3px 10px rgba(56, 189, 248, 0.25);
    }

    /* Single line gradient title */
    .hub-title {
        font-size: 19px;
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
    </style>
""", unsafe_allow_html=True)

# Locations list with icons
LOCATION_DISPLAY_MAP = {
    "RPPL - DEL": "📍 RPPL - DEL",
    "RPPL - KOL": "📍 RPPL - KOL",
    "RPPL - HYD": "📍 RPPL - HYD",
    "RPPL - PUN": "📍 RPPL - PUN",
    "RPPL - BLR": "📍 RPPL - BLR",
    "RPPL - CHE": "📍 RPPL - CHE",
    "Headoffice": "🏢 Headoffice"
}

ORDERED_LOCATIONS = list(LOCATION_DISPLAY_MAP.keys())

# Short Name + Unique Logo for each tab
AVAILABLE_MODULES = {
    "AMZ": {"name": "Amazon Invoice Editor", "code": "AMZ-ED", "logo": "📦"},
    "BLNK": {"name": "Blinkit e-Invoice Tool", "code": "BLNK-INV", "logo": "⚡"},
    "LBL": {"name": "Label Editor", "code": "LBL-ED", "logo": "🏷️"},
    "CONF": {"name": "Settings & Reports", "code": "SETTINGS", "logo": "⚙️"}
}

# ----------------- SESSION STATE MANAGEMENT -----------------
if "selected_wh" not in st.session_state:
    st.session_state.selected_wh = "RPPL - DEL"
if "open_tabs" not in st.session_state:
    st.session_state.open_tabs = ["AMZ"]
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "AMZ"
if "processed_history" not in st.session_state:
    st.session_state.processed_history = []
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "show_history" not in st.session_state:
    st.session_state.show_history = False

def get_current_location_history():
    if st.session_state.selected_wh == "Headoffice":
        return st.session_state.processed_history
    return [
        item for item in st.session_state.processed_history 
        if item.get("location") == st.session_state.selected_wh
    ]

def reset_current_workspace():
    current_wh = st.session_state.selected_wh
    if current_wh == "Headoffice":
        st.session_state.processed_history = []
    else:
        st.session_state.processed_history = [
            item for item in st.session_state.processed_history 
            if item.get("location") != current_wh
        ]
    st.session_state.uploader_key += 1
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

    st.divider()
    st.caption("System Version: v3.2 Enterprise")

# ----------------- TOP ACTION BAR: SLEEK & COMPACT -----------------
top_tabs_col, top_wh_col, top_hist_col, top_reset_col = st.columns([4.8, 2.2, 1.6, 1.1])

with top_tabs_col:
    num_tabs = len(st.session_state.open_tabs)
    tab_cols = st.columns(num_tabs + 1)
    
    for idx, t_key in enumerate(st.session_state.open_tabs):
        info = AVAILABLE_MODULES[t_key]
        is_active = (t_key == st.session_state.active_tab and not st.session_state.show_history)
        tab_label = f"{info['logo']} {info['code']}"
        
        with tab_cols[idx]:
            sub_col1, sub_col2 = st.columns([3, 1])
            with sub_col1:
                if st.button(tab_label, key=f"pin_tab_{t_key}", type="primary" if is_active else "secondary", use_container_width=True):
                    st.session_state.active_tab = t_key
                    st.session_state.show_history = False
                    st.rerun()
            with sub_col2:
                if len(st.session_state.open_tabs) > 1:
                    if st.button("\u2715", key=f"close_tab_{t_key}", help="Close tab"):
                        st.session_state.open_tabs.remove(t_key)
                        if st.session_state.active_tab == t_key:
                            st.session_state.active_tab = st.session_state.open_tabs[0]
                        st.rerun()
                        
    with tab_cols[num_tabs]:
        if st.button("🔄 Unpin", key="btn_refresh_tabs", help="Clear pinned tabs"):
            clear_all_pinned_tabs()

with top_wh_col:
    selected_display = st.selectbox(
        "WH | Location",
        options=ORDERED_LOCATIONS,
        format_func=lambda x: LOCATION_DISPLAY_MAP.get(x, x),
        index=ORDERED_LOCATIONS.index(st.session_state.selected_wh) if st.session_state.selected_wh in ORDERED_LOCATIONS else 0,
        help="Select Warehouse Location"
    )
    st.session_state.selected_wh = selected_display

loc_history = get_current_location_history()

with top_hist_col:
    hist_btn_label = f"🕒 History ({len(loc_history)})"
    if st.button(hist_btn_label, use_container_width=True, type="primary" if st.session_state.show_history else "secondary"):
        st.session_state.show_history = not st.session_state.show_history
        st.rerun()

with top_reset_col:
    if st.button("🔄 Reset", use_container_width=True, type="secondary"):
        reset_current_workspace()

st.divider()

# =========================================================================
# HISTORY VIEW (IF TOGGLED)
# =========================================================================
if st.session_state.show_history:
    st.subheader("📜 Processed History & Downloads")
    if not loc_history:
        st.info(f"Abhi tak **{st.session_state.selected_wh}** ke liye koi file process nahi hui hai.")
    else:
        st.success(f"💾 Total Saved Files: **{len(loc_history)}**")
        if len(loc_history) > 1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, item in enumerate(loc_history):
                    zip_file.writestr(f"{idx+1}_{item['file_name']}", item["data"])
            zip_buffer.seek(0)

            st.download_button(
                label=f"📦 Download All {st.session_state.selected_wh} Files as ZIP",
                data=zip_buffer,
                file_name=f"{st.session_state.selected_wh.replace(' ', '_')}_Invoices.zip",
                mime="application/zip",
                use_container_width=True
            )

        st.write("---")
        for idx, item in enumerate(loc_history):
            d_col1, d_col2 = st.columns([3, 1])
            with d_col1:
                st.write(
                    f"📁 **{item['original_name']}** — `[{LOCATION_DISPLAY_MAP.get(item.get('location'), item.get('location'))}]` | "
                    f"`{item['pages']} Pages Kept` | `{item['stamped']} Stamped` | "
                    f"`{item['removed_pan']} Non-PAN Filtered` | `{item.get('removed_desc', 0)} No-Desc Filtered`"
                )
            with d_col2:
                st.download_button(
                    label="📥 Download PDF",
                    data=item["data"],
                    file_name=item["file_name"],
                    mime="application/pdf",
                    key=f"dl_hist_{st.session_state.selected_wh}_{idx}",
                    use_container_width=True
                )

# =========================================================================
# MODULE 1: AMAZON INVOICE EDITOR
# =========================================================================
elif st.session_state.active_tab == "AMZ":
    # Clean subheader without location brackets
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
        tokens = get_token_words(raw_text)
        return has_header and len(tokens) >= 5

    def generate_barcode_image(code_text):
        try:
            code128 = barcode.get_barcode_class('code128')
            writer = ImageWriter()
            writer.font_path = None
            barcode_instance = code128(str(code_text).strip(), writer=writer)
            
            buffer = io.BytesIO()
            barcode_instance.write(
                buffer,
                options={
                    'write_text': False,
                    'module_width': 0.45,
                    'module_height': 15.0,
                    'quiet_zone': 1.5,
                    'dpi': 300
                }
            )
            buffer.seek(0)
            return buffer.getvalue()
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
                df
