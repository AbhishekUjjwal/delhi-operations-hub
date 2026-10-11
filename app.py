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
    page_title="Operations Hub | Romsons Prime Pvt Ltd",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling: Sleek ERP / Enterprise Multi-Tab Dashboard
st.markdown("""
    <style>
    .main { background-color: #0b132b; color: #f8fafc; }
    
    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        background-color: #1c2541;
        border-right: 1px solid #3a506b;
    }
    
    /* Top Action Bar Styling */
    div[data-testid="stSelectbox"] label {
        display: none;
    }
    </style>
""", unsafe_allow_html=True)

# ----------------- WAREHOUSE USER CREDENTIALS SETUP -----------------
# Aap yahan har warehouse ke liye custom Username aur Password define kar sakte hain
WAREHOUSE_USERS = {
    "delhi_user": {"password": "delhi@123", "role": "warehouse", "location": "RPPL - DEL"},
    "kolkata_user": {"password": "kolkata@123", "role": "warehouse", "location": "RPPL - KOL"},
    "hyd_user": {"password": "hyd@123", "role": "warehouse", "location": "RPPL - HYD"},
    "pune_user": {"password": "pune@123", "role": "warehouse", "location": "RPPL - PUN"},
    "blr_user": {"password": "blr@123", "role": "warehouse", "location": "RPPL - BLR"},
    "chennai_user": {"password": "chennai@123", "role": "warehouse", "location": "RPPL - CHE"},
    "headoffice_admin": {"password": "admin@rppl", "role": "admin", "location": "Headoffice"}
}

ORDERED_LOCATIONS = [
    "RPPL - DEL",
    "RPPL - KOL",
    "RPPL - HYD",
    "RPPL - PUN",
    "RPPL - BLR",
    "RPPL - CHE",
    "Headoffice"
]

# ----------------- SESSION STATE MANAGEMENT -----------------
AVAILABLE_MODULES = {
    "AMZ": {"name": "Amazon Invoice Editor", "code": "AMZ-ED", "icon": "📑"},
    "BLNK": {"name": "Blinkit e-Invoice Tool", "code": "BLNK-INV", "icon": "⚡"},
    "LBL": {"name": "Label Editor", "code": "LBL-ED", "icon": "🏷️"},
    "HIST": {"name": "History / Downloads", "code": "HISTORY", "icon": "📜"},
    "CONF": {"name": "Settings & Reports", "code": "SETTINGS", "icon": "⚙️"}
}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_info" not in st.session_state:
    st.session_state.user_info = None
if "selected_wh" not in st.session_state:
    st.session_state.selected_wh = "RPPL - DEL"
if "open_tabs" not in st.session_state:
    st.session_state.open_tabs = ["AMZ", "HIST"]
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "AMZ"
if "processed_history" not in st.session_state:
    st.session_state.processed_history = []
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

def reset_all_workspace():
    st.session_state.processed_history = []
    st.session_state.open_tabs = ["AMZ", "HIST"]
    st.session_state.active_tab = "AMZ"
    st.session_state.uploader_key += 1
    st.rerun()

# ----------------- LOGIN SCREEN -----------------
if not st.session_state.logged_in:
    st.write("")
    st.write("")
    l_col1, l_col2, l_col3 = st.columns([1, 2, 1])
    with l_col2:
        st.markdown("""
            <div style="background-color: #1c2541; padding: 25px; border-radius: 12px; border: 1px solid #3a506b;">
                <h3 style="margin-top:0; color:#48cae4; text-align:center;">🏛️ Delhi Operations Hub</h3>
                <p style="text-align:center; color:#cbd5e1; font-size:14px;">Romsons prime pvt ltd operation portal</p>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        with st.form("login_form"):
            username = st.text_input("Username").strip().lower()
            password = st.text_input("Password", type="password")
            submit = st.form_submit_button("🔐 Login to Workspace", use_container_width=True)

            if submit:
                if username in WAREHOUSE_USERS and WAREHOUSE_USERS[username]["password"] == password:
                    st.session_state.logged_in = True
                    st.session_state.user_info = WAREHOUSE_USERS[username]
                    st.session_state.selected_wh = WAREHOUSE_USERS[username]["location"]
                    st.success("Login Successful!")
                    st.rerun()
                else:
                    st.error("Invalid Username ya Password!")
    st.stop()

# ----------------- SIDEBAR NAVIGATION -----------------
with st.sidebar:
    st.title("🏛️ Delhi Operations Hub")
    st.caption("Romsons prime pvt ltd operation portal")
    st.caption(f"👤 Logged in: **{st.session_state.selected_wh}**")
    
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.user_info = None
        st.rerun()

    st.divider()
    st.markdown("### 🧭 Workspace Modules")
    st.caption("Click to pin and open workspace:")

    for key, info in AVAILABLE_MODULES.items():
        is_pinned = key in st.session_state.open_tabs
        pin_label = f"📌 {info['code']} | {info['name']}" if is_pinned else f"➕ {info['code']} | {info['name']}"
        if st.button(pin_label, key=f"side_nav_{key}", use_container_width=True):
            if key not in st.session_state.open_tabs:
                st.session_state.open_tabs.append(key)
            st.session_state.active_tab = key
            st.rerun()

    st.divider()
    st.caption("System Version: v3.2 Enterprise")

# ----------------- TOP BAR: TABS, WH LOCATION & RESET -----------------
top_tabs_col, top_wh_col, top_reset_col = st.columns([5, 2, 1])

with top_tabs_col:
    tab_cols = st.columns(len(st.session_state.open_tabs) + 1)
    for idx, t_key in enumerate(st.session_state.open_tabs):
        info = AVAILABLE_MODULES[t_key]
        is_active = (t_key == st.session_state.active_tab)
        label = f"📍 {info['code']}" if is_active else f"{info['code']}"
        
        with tab_cols[idx]:
            sub_col1, sub_col2 = st.columns([3, 1])
            with sub_col1:
                if st.button(label, key=f"pin_tab_{t_key}", type="primary" if is_active else "secondary", use_container_width=True):
                    st.session_state.active_tab = t_key
                    st.rerun()
            with sub_col2:
                if len(st.session_state.open_tabs) > 1:
                    if st.button("✖", key=f"close_tab_{t_key}", help="Close Tab"):
                        st.session_state.open_tabs.remove(t_key)
                        if st.session_state.active_tab == t_key:
                            st.session_state.active_tab = st.session_state.open_tabs[0]
                        st.rerun()

with top_wh_col:
    # Headoffice/Admin sabhi locations choose kar sakte hain, regular users ka locked rehta hai
    if st.session_state.user_info.get("role") == "admin":
        st.session_state.selected_wh = st.selectbox(
            "WH | Location",
            options=ORDERED_LOCATIONS,
            index=ORDERED_LOCATIONS.index(st.session_state.selected_wh),
            help="Select Warehouse / Location"
        )
    else:
        st.markdown(f"""
            <div style="background-color: #1c2541; padding: 7px 12px; border-radius: 6px; border: 1px solid #3a506b; text-align: center; font-weight: bold; color: #48cae4;">
                🏢 {st.session_state.selected_wh}
            </div>
        """, unsafe_allow_html=True)

with top_reset_col:
    if st.button("🔄 Reset All", use_container_width=True, type="secondary"):
        reset_all_workspace()

st.divider()

# =========================================================================
# MODULE 1: AMAZON INVOICE EDITOR
# =========================================================================
if st.session_state.active_tab == "AMZ":
    st.subheader(f"📑 Amazon Invoice Editor & Barcode Stamper [{st.session_state.selected_wh}]")

    TARGET_PAN = "aalcr5906l"

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        uploaded_csv = st.file_uploader(
            "1. Upload Shipment Report (CSV / Excel)", 
            type=["csv", "xlsx", "xls"], 
            key=f"amz_csv_{st.session_state.uploader_key}"
        )
    with col_u2:
        uploaded_pdfs = st.file_uploader(
            "2. Upload Invoice PDF(s) [Multiple Files Allowed]", 
            type=["pdf"], 
            accept_multiple_files=True, 
            key=f"amz_pdf_{st.session_state.uploader_key}"
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
                    'module_width
