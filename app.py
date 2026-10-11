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

# Custom Styling: Sleek, Compact & Theme-Adaptive
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

    /* 2. COMPACT TAB & PILL BUTTONS */
    div[data-testid="stButton"] button {
        border-radius: 16px !important;
        font-size: 12px !important;
        font-weight: 600 !important;
        padding: 4px 10px !important;
        min-height: 30px !important;
        height: 30px !important;
        line-height: 1 !important;
        border: 1px solid rgba(148, 163, 184, 0.3) !important;
        white-space: nowrap !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-shadow: none !important;
    }

    div[data-testid="stButton"] button:hover {
        border-color: #38bdf8 !important;
        color: #38bdf8 !important;
    }

    /* 3. COMPACT SELECTBOX */
    div[data-testid="stSelectbox"] label {
        display: none !important;
    }
    div[data-testid="stSelectbox"] > div > div {
        min-height: 30px !important;
        height: 30px !important;
        padding: 0px 8px !important;
        font-size: 12px !important;
        border-radius: 16px !important;
        border: 1px solid rgba(148, 163, 184, 0.3) !important;
    }

    /* 4. SINGLE LINE GRADIENT TITLE */
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

    /* Professional Success Badge */
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

# RESET SIRF WORKSPACE UPLOADERS KO CLEAR KAREGA, HISTORY KO NAHI CHHEDEGA
def reset_current_workspace():
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

# ----------------- TOP ACTION BAR -----------------
top_tabs_col, top_wh_col, top_hist_col, top_reset_col = st.columns([5.2, 2.0, 1.4, 1.0])

with top_tabs_col:
    cols = st.columns(len(st.session_state.open_tabs) + 1)
    
    for idx, t_key in enumerate(st.session_state.open_tabs):
        info = AVAILABLE_MODULES[t_key]
        is_active = (t_key == st.session_state.active_tab and not st.session_state.show_history)
        
        if len(st.session_state.open_tabs) > 1:
            tab_label = f"{info['logo']} {info['code']}  ✕"
        else:
            tab_label = f"{info['logo']} {info['code']}"

        with cols[idx]:
            if st.button(tab_label, key=f"tab_chip_{t_key}", type="primary" if is_active else "secondary", use_container_width=True):
                if is_active and len(st.session_state.open_tabs) > 1:
                    st.session_state.open_tabs.remove(t_key)
                    st.session_state.active_tab = st.session_state.open_tabs[0]
                else:
                    st.session_state.active_tab = t_key
                    st.session_state.show_history = False
                st.rerun()

    with cols[len(st.session_state.open_tabs)]:
        if st.button("🔄 Unpin", key="btn_unpin_all", use_container_width=True, help="Clear pinned tabs"):
            clear_all_pinned_tabs()

with top_wh_col:
    selected_display = st.selectbox(
        "WH | Location",
        options=ORDERED_LOCATIONS,
        format_func=lambda x: LOCATION_DISPLAY_MAP.get(x, x),
        index=ORDERED_LOCATIONS.index(st.session_state.selected_wh) if st.session_state.selected_wh in ORDERED_LOCATIONS else 0
    )
    st.session_state.selected_wh = selected_display

loc_history = get_current_location_history()

with top_hist_col:
    if st.button(f"🕒 History ({len(loc_history)})", use_container_width=True, type="primary" if st.session_state.show_history else "secondary"):
        st.session_state.show_history = not st.session_state.show_history
        st.rerun()

with top_reset_col:
    if st.button("🔄 Reset", use_container_width=True, type="secondary", help="Clears upload inputs only; history remains safe"):
        reset_current_workspace()

st.divider()

# =========================================================================
# HISTORY VIEW (PERMANENT & SAVED ACROSS RESETS)
# =========================================================================
if st.session_state.show_history:
    h_top_c1, h_top_c2 = st.columns([4, 1])
    with h_top_c1:
        st.subheader("📜 Processed History & Downloads (Safe Storage)")
    with h_top_c2:
        if loc_history:
            if st.button("🗑️ Clear This Location History", key="btn_del_hist", use_container_width=True):
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
        
        # BULK ZIP DOWNLOAD IN HISTORY
        if len(loc_history) > 1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, item in enumerate(loc_history):
                    zip_file.writestr(f"{idx+1}_{item['file_name']}", item["data"])
            zip_buffer.seek(0)

            st.download_button(
                label=f"📦 Download All {st.session_state.selected_wh} Files as ZIP ({len(loc_history)} Files)",
                data=zip_buffer,
                file_name=f"{st.session_state.selected_wh.replace(' ', '_')}_History.zip",
                mime="application/zip",
                key="btn_zip_history_download",
                use_container_width=True
            )

        st.write("---")
        # SINGLE-SINGLE FILE DOWNLOAD IN HISTORY
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

        # Silent Column Mapping
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

            current_processed = []
            total_files = len(uploaded_pdfs)

            for file_idx, pdf_file in enumerate(uploaded_pdfs):
                status_text.text(f"Processing File {file_idx+1}/{total_files}: {pdf_file.name}...")

                pdf_bytes = pdf_file.read()
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                new_doc = fitz.open()

                total_pages = len(doc)
                file_stamped_count = 0
                file_removed_pan = 0
                file_removed_desc = 0

                for page_num in range(total_pages):
                    page = doc[page_num]
                    raw_text = page.get_text()
                    text_lower = raw_text.lower()

                    # Rule 1: Strict PAN Filter
                    if TARGET_PAN not in text_lower:
                        file_removed_pan += 1
                        continue

                    # Rule 2: Description Check Filter
                    if not has_valid_description(raw_text):
                        file_removed_desc += 1
                        continue

                    # Extract 17-digit Order Number
                    order_clean = None
                    order_match = re.search(r'(\d{3})\s*
