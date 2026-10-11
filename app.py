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

# Custom Styling: Vibrant, Modern & Theme-Adaptive (Works great in both Dark & Light themes)
st.markdown("""
    <style>
    /* Metric Cards - Modern Gradient */
    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.7) 0%, rgba(15, 23, 42, 0.9) 100%);
        border: 1px solid rgba(56, 189, 248, 0.3);
        border-radius: 12px;
        padding: 14px 18px;
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.15);
    }
    
    /* Top action selectbox label hide */
    div[data-testid="stSelectbox"] label {
        display: none;
    }
    
    /* Pill buttons - Matches custom SaaS look */
    div[data-testid="stButton"] button {
        border-radius: 24px !important;
        font-weight: 600 !important;
        transition: all 0.25s ease-in-out !important;
        border: 1px solid rgba(148, 163, 184, 0.25) !important;
    }
    
    div[data-testid="stButton"] button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(56, 189, 248, 0.25);
    }
    
    /* File Uploader Cards */
    div[data-testid="stFileUploader"] {
        border-radius: 12px;
        padding: 8px;
    }
    
    /* Status Badge highlight */
    .status-badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 14px;
        font-size: 12px;
        font-weight: 700;
        letter-spacing: 0.3px;
    }
    .badge-del { background: rgba(59, 130, 246, 0.2); color: #60a5fa; border: 1px solid #3b82f6; }
    .badge-kol { background: rgba(245, 158, 11, 0.2); color: #fbbf24; border: 1px solid #f59e0b; }
    .badge-hyd { background: rgba(16, 185, 129, 0.2); color: #34d399; border: 1px solid #10b981; }
    .badge-pun { background: rgba(236, 72, 153, 0.2); color: #f472b6; border: 1px solid #ec4899; }
    .badge-blr { background: rgba(139, 92, 246, 0.2); color: #a78bfa; border: 1px solid #8b5cf6; }
    .badge-che { background: rgba(6, 182, 212, 0.2); color: #22d3ee; border: 1px solid #06b6d4; }
    .badge-ho  { background: rgba(239, 68, 68, 0.2); color: #f87171; border: 1px solid #ef4444; }
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

AVAILABLE_MODULES = {
    "AMZ": {"name": "Amazon Invoice Editor", "code": "AMZ-ED", "icon": "📑"},
    "BLNK": {"name": "Blinkit e-Invoice Tool", "code": "BLNK-INV", "icon": "⚡"},
    "LBL": {"name": "Label Editor", "code": "LBL-ED", "icon": "🏷️"},
    "CONF": {"name": "Settings & Reports", "code": "SETTINGS", "icon": "⚙️"}
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

# ----------------- SIDEBAR NAVIGATION -----------------
with st.sidebar:
    st.title("🏛️ PAN OPERATIONS HUB")
    st.caption("Romsons prime pvt ltd operation portal")
    st.divider()

    st.markdown("### 🧭 Module Navigation")
    st.caption("Click to pin and open workspace:")

    for key, info in AVAILABLE_MODULES.items():
        is_pinned = key in st.session_state.open_tabs
        pin_label = f"📌 {info['code']} | {info['name']}" if is_pinned else f"➕ {info['code']} | {info['name']}"
        if st.button(pin_label, key=f"side_nav_{key}", use_container_width=True):
            if key not in st.session_state.open_tabs:
                st.session_state.open_tabs.append(key)
            st.session_state.active_tab = key
            st.session_state.show_history = False
            st.rerun()

    st.divider()
    st.caption("System Version: v3.2 Enterprise")

# ----------------- TOP ACTION BAR: PINNED TABS, WH LOCATION, HISTORY & RESET -----------------
top_tabs_col, top_wh_col, top_hist_col, top_reset_col = st.columns([5, 2.4, 1.8, 1.2])

with top_tabs_col:
    tab_cols = st.columns(len(st.session_state.open_tabs) + 1)
    for idx, t_key in enumerate(st.session_state.open_tabs):
        info = AVAILABLE_MODULES[t_key]
        is_active = (t_key == st.session_state.active_tab and not st.session_state.show_history)
        label = f"📍 {info['code']}" if is_active else f"{info['code']}"
        
        with tab_cols[idx]:
            sub_col1, sub_col2 = st.columns([3, 1])
            with sub_col1:
                if st.button(label, key=f"pin_tab_{t_key}", type="primary" if is_active else "secondary", use_container_width=True):
                    st.session_state.active_tab = t_key
                    st.session_state.show_history = False
                    st.rerun()
            with sub_col2:
                if len(st.session_state.open_tabs) > 1:
                    if st.button("✖", key=f"close_tab_{t_key}", help="Close Tab"):
                        st.session_state.open_tabs.remove(t_key)
                        if st.session_state.active_tab == t_key:
                            st.session_state.active_tab = st.session_state.open_tabs[0]
                        st.rerun()

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
    st.subheader(f"📜 Processed History & Downloads — [{LOCATION_DISPLAY_MAP.get(st.session_state.selected_wh, st.session_state.selected_wh)}]")
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
    st.subheader(f"📑 Amazon Invoice Editor & Barcode Stamper — [{LOCATION_DISPLAY_MAP.get(st.session_state.selected_wh, st.session_state.selected_wh)}]")

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

        m1, m2, m3 = st.columns(3)
        m1.metric("📦 Order Column", str(order_col))
        m2.metric("🏷️ Product Column", str(title_col if title_col else sku_col))
        m3.metric("🚚 Tracking Column", str(tracking_col))

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
            progress_bar = st.progress
