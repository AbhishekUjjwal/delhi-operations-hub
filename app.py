import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import pymupdf as fitz
import re
import io
import json
import zipfile
import barcode
from barcode.writer import ImageWriter
from datetime import datetime, timezone, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# IST TIME HELPER (Exact Indian Standard Time)
IST = timezone(timedelta(hours=5, minutes=30))
def get_current_ist_str():
    return datetime.now(IST).strftime("%d-%b %I:%M %p")

# Page Configuration
st.set_page_config(
    page_title="PAN OPERATIONS HUB | Central Portal",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling: Zero Top Padding, Compact Actions & Snappy Layout
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

    /* 2. HIDE BOTTOM RIGHT BADGES & WATERMARKS */
    [data-testid="stViewerBadge"] {display: none !important; visibility: hidden !important;}
    div[class*="viewerBadge"] {display: none !important; visibility: hidden !important;}
    iframe[title*="streamlit"] {display: none !important;}
    div[data-testid="manage-app-button"] {display: none !important;}
    div[class*="manage-app"] {display: none !important;}
    div[class*="floating"] {display: none !important;}

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

    /* 7. HEADINGS SPACING COMPACT */
    h2, h3, h4 {
        margin-top: 0px !important;
        margin-bottom: 4px !important;
        padding-top: 0px !important;
    }

    /* 8. SIDEBAR TITLE */
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

AVAILABLE_MODULES = {
    "AMZ": {"name": "Amazon Invoice Editor", "code": "AMZ-ED", "logo": "📦"},
    "BLNK_PO": {"name": "Blinkit PO Uploader", "code": "BLNK-PO", "logo": "⚡"},
    "BLNK": {"name": "Blinkit e-Invoice Tool", "code": "BLNK-INV", "logo": "🧾"},
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

    st.divider()
    st.caption("System Version: v3.2 Enterprise")

# ----------------- TOP LAYOUT -----------------
header_left, header_right = st.columns([7.8, 2.2])

with header_left:
    tab_cols = st.columns([0.8] * len(st.session_state.open_tabs) + [3.8])
    for idx, t_key in enumerate(st.session_state.open_tabs):
        info = AVAILABLE_MODULES[t_key]
        is_active = (t_key == st.session_state.active_tab and not st.session_state.show_history)
        tab_label = f"{info['logo']} {info['code']} ✕" if len(st.session_state.open_tabs) > 1 else f"{info['logo']} {info['code']}"
        with tab_cols[idx]:
            if st.button(tab_label, key=f"tab_chip_{t_key}", type="primary" if is_active else "secondary"):
                if is_active and len(st.session_state.open_tabs) > 1:
                    st.session_state.open_tabs.remove(t_key)
                    st.session_state.active_tab = st.session_state.open_tabs[0]
                else:
                    st.session_state.active_tab = t_key
                    st.session_state.show_history = False
                st.rerun()

with header_right:
    st.markdown('<div class="right-action-cluster">', unsafe_allow_html=True)
    
    loc_sub1, loc_sub2 = st.columns([0.38, 1.62])
    with loc_sub1:
        if st.button("🔄", key="btn_refresh_pinned", help="Click to unpin all extra tabs", use_container_width=True):
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
        if st.button("🔄 Reset", key="top_reset_btn", type="secondary", help="Reset upload area", use_container_width=True):
            reset_current_workspace()
        
    st.markdown('</div>', unsafe_allow_html=True)

st.divider()

# =========================================================================
# HISTORY VIEW (WITH EXACT IST TIMEZONE)
# =========================================================================
if st.session_state.show_history:
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
                if st.session_state.selected_wh == "Headoffice":
                    st.session_state.permanent_history = [
                        item for item in st.session_state.permanent_history 
                        if item.get("module") != st.session_state.active_tab
                    ]
                else:
                    st.session_state.permanent_history = [
                        item for item in st.session_state.permanent_history 
                        if not (item.get("module") == st.session_state.active_tab and item.get("location") == st.session_state.selected_wh)
                    ]
                st.rerun()

    if not loc_history:
        st.info("No files")
    else:
        st.success(f"💾 Total Saved Files: **{len(loc_history)}**")
        
        if len(loc_history) > 1:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for idx, item in enumerate(loc_history):
                    loc_tag = item.get("location", "GEN")
                    safe_fname = f"{idx+1}_{loc_tag}_{item['file_name']}"
                    zf.writestr(safe_fname, item["data"])
            zip_bytes = zip_buf.getvalue()

            zip_title = f"📦 Download All Central Master Invoices as ZIP ({len(loc_history)} Files)" if st.session_state.selected_wh == "Headoffice" else f"📦 Re-Download All History as ZIP ({len(loc_history)} Files)"
            st.download_button(
                label=zip_title,
                data=zip_bytes,
                file_name=f"{st.session_state.selected_wh}_{st.session_state.active_tab}_Master_History.zip",
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
                    "module": "AMZ",
                    "time": get_current_ist_str()
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

    # BULK & SINGLE DOWNLOAD SECTION
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
            on_click=record_batch_to_history,
            args=(batch_list,),
            use_container_width=True
        )
        st.write("")

        for idx, item in enumerate(batch_list):
            d_col1, d_col2 = st.columns([3, 1])
            with d_col1:
                st.write(
                    f"📁 **{item['original_name']}** — `[{LOCATION_DISPLAY_MAP.get(item.get('location'), item.get('location'))}]` | "
                    f"`{item['pages']} Pages Kept` | `{item['stamped']} Stamped` | "
                    f"⏰ *{item.get('time', '')}*"
                )
            with d_col2:
                st.download_button(
                    label="📥 Download PDF",
                    data=item["data"],
                    file_name=item["file_name"],
                    mime="application/pdf",
                    key=f"dl_single_live_{st.session_state.selected_wh}_{idx}",
                    on_click=record_to_history,
                    args=(item,),
                    use_container_width=True
                )

# =========================================================================
# MODULE 2: BLINKIT PO UPLOADER
# =========================================================================
elif st.session_state.active_tab == "BLNK_PO":
    st.subheader("⚡ Blinkit PO Uploader & Converter")
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

# =========================================================================
# MODULE 3: BLINKIT E-INVOICE TOOL (STRICT ORDER ID & IST CLOCK)
# =========================================================================
elif st.session_state.active_tab == "BLNK":
    st.subheader("🧾 Blinkit e-Invoice Management Tool & Reconciler")
    st.caption("Reconciles Blinkit 50x/25x Box factor, generates 85-Column NIC standard Excel, JSON v1.01, and preview.")

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
        "Kolkata", "Ahmedabad", "Patna", "Ranchi", "Chandigarh", "Faridabad", "Agra", "Meerut"
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

        # 1. CLEAN EXTERNAL ORDER EXTRACTION (STRICTLY NUMBERS - NO INVOICE NO LEAKAGE)
        ext_order_id = None
        ext_matches = [w for w in words if "extern" in w[4].lower()]
        if ext_matches:
            anchor = ext_matches[0]
            # Capture words on the right OR immediately below (<= 11 points) that are strictly digits
            num_words = [
                w for w in words 
                if (anchor[1] - 4 <= w[1] <= anchor[3] + 11.5) 
                and (anchor[0] + 45 <= w[0] <= 445)
                and w[4].strip().isdigit()  # Strictly digit only (excludes M90-...)
            ]
            num_words.sort(key=lambda w: (round(w[1] / 6) * 6, w[0]))
            if num_words:
                ext_order_id = "".join([w[4].strip() for w in num_words])

        if not ext_order_id:
            ext_match = re.search(r"Extern(?:al)?\s*Order\s*(?:No\.?|ID)?\s*[:\s]*([0-9\s]+?)(?=\n\s*[A-Za-z]|Invoice|Date|$)", full_text, re.IGNORECASE)
            if ext_match:
                ext_order_id = "".join(ext_match.group(1).split())
        if not ext_order_id:
            digits_match = re.findall(r"\b(49\d{8,14}|5\d{8,14}|\d{12,18})\b", full_text)
            if digits_match:
                ext_order_id = digits_match[0]
        if not ext_order_id:
            ext_order_id = "ExtOrder"

        # 2. INVOICE NO EXTRACTION
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

        sold_header = [w for w in words if "sold" in w[4].lower()]
        bill_header = [w for w in words if "billing" in w[4].lower()]
        table_header = [w for w in words if w[4].lower() in ["s.no", "sl.no", "item", "hsn"]]

        y_sold = sold_header[0][1] if sold_header else 80
        y_bill = bill_header[0][1] if bill_header else 215
        y_table = table_header[0][1] if table_header else 340
        split_x = 300

        seller_rect = fitz.Rect(20, y_sold, split_x, y_bill)
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

        seller_addr_candidates = [l for l in s_lines[1:] if not re.search(r"(@|contact|phone|mob|gstin|pan|india)", l, re.I)]
        seller_addr = ", ".join(seller_addr_candidates[:2]) if seller_addr_candidates else "Ground Floor, B-248 B Block, Naraina Industrial Area, Phase 1"

        billing_rect = fitz.Rect(20, y_bill, split_x, y_table)
        billing_text = page0.get_text("text", clip=billing_rect).strip()
        b_lines = [l.strip() for l in billing_text.split("\n") if l.strip() and not re.search(r"billing\s*addr", l, re.I)]

        buyer_name = b_lines[0] if b_lines else "BLINK COMMERCE PRIVATE LIMITED"
        b_addr_candidates = [l for l in b_lines[1:] if not re.search(r"(@|contact|phone|mob|gstin|pan|india)", l, re.I)]
        buyer_addr1 = ", ".join(b_addr_candidates[:2]) if b_addr_candidates else (b_lines[1] if len(b_lines) > 1 else "Commercial Facility")

        em_m = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", billing_text)
        buyer_email = em_m.group(0).strip() if em_m else ""

        ph_m = re.search(r"(?:Contact|Phone|Mob)?\s*[:\s]*([6-9]\d{9})", billing_text, re.IGNORECASE)
        buyer_phone = ph_m.group(1).strip() if ph_m else ""

        b_gst_m = re.search(r"GSTIN\s*[:\s]*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", billing_text, re.IGNORECASE)
        buyer_gstin = b_gst_m.group(1).strip() if b_gst_m else ""

        if not buyer_gstin:
            all_gstins = re.findall(r"\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b", full_text)
            buyer_gstin = all_gstins[1] if len(all_gstins) > 1 else (all_gstins[0] if all_gstins else "06AAFCG9846E1ZF")

        buyer_state_code = buyer_gstin[:2]
        buyer_state_name = STATE_CODE_MAP.get(buyer_state_code, "DELHI")

        b_pin_m = re.findall(r"\b[1-9][0-9]{5}\b", billing_text)
        buyer_pin = int(b_pin_m[0]) if b_pin_m else int(buyer_state_code + "0001")
        buyer_loc = clean_extracted_city(billing_text, buyer_state_name)

        shipping_rect = fitz.Rect(split_x + 5, y_bill, 585, y_table)
        shipping_text = page0.get_text("text", clip=shipping_rect).strip()
        shp_lines = [l.strip() for l in shipping_text.split("\n") if l.strip() and not re.search(r"shipping\s*addr", l, re.I)]

        ship_name = shp_lines[0] if shp_lines else buyer_name
        s_addr_candidates = [l for l in shp_lines[1:] if not re.search(r"(@|contact|phone|mob|gstin|pan|india)", l, re.I)]
        ship_addr1 = ", ".join(s_addr_candidates[:2]) if s_addr_candidates else (shp_lines[1] if len(shp_lines) > 1 else buyer_addr1)

        shp_gst_m = re.search(r"GSTIN\s*[:\s]*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", shipping_text, re.IGNORECASE)
        ship_gstin = shp_gst_m.group(1).strip() if shp_gst_m else buyer_gstin
        ship_state_code = ship_gstin[:2]
        ship_state_name = STATE_CODE_MAP.get(ship_state_code, buyer_state_name)

        shp_pin_m = re.findall(r"\b[1-9][0-9]{5}\b", shipping_text)
        ship_pin = int(shp_pin_m[0]) if shp_pin_m else buyer_pin
        ship_loc = clean_extracted_city(shipping_text, ship_state_name)

        clean_b_str = re.sub(r'[^A-Za-z0-9]', '', f"{buyer_addr1}{buyer_pin}").lower()
        clean_s_str = re.sub(r'[^A-Za-z0-9]', '', f"{ship_addr1}{ship_pin}").lower()
        shipping_is_different = (clean_b_str != clean_s_str) and (buyer_pin != ship_pin or buyer_addr1 != ship_addr1)

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

        is_blinkit = bool(re.search(r"blink\s*commerce|blinkit", f"{buyer_name} {full_text}", re.IGNORECASE))

        has_cgst_header = bool(re.search(r"\bCGST\b", full_text, re.I))
        cgst_vals = re.findall(r"CGST[\s\S]*?(\d+\.\d{2})", full_text)
        has_positive_cgst = any(float(v) > 0 for v in cgst_vals[:4]) if cgst_vals else False

        if has_cgst_header and has_positive_cgst:
            is_interstate = False
        else:
            is_interstate = (seller_state_code != buyer_pos_code)

        inv_total_m = re.search(r"(?:Total\s*Invoice\s*Value|Invoice\s*Total|Grand\s*Total|Total\s*Amount)[:\s]*[₹\s]*([0-9,]+\.\d{2})", full_text, re.I)
        printed_grand_total = float(inv_total_m.group(1).replace(",", "")) if inv_total_m else None

        return {
            "invoice_no": invoice_no,
            "ext_order_id": ext_order_id,
            "doc_date": std_date_for_json,
            "clean_date": clean_date,
            "pdf_filename": pdf_filename,
            "seller_name": seller_name,
            "seller_gstin": seller_gstin,
            "seller_pin": seller_pin,
            "seller_addr": seller_addr,
            "seller_loc": seller_loc,
            "seller_state_code": seller_state_code,
            "seller_state_name": seller_state_name,
            "buyer_name": buyer_name,
            "buyer_trade_name": buyer_name,
            "buyer_addr1": buyer_addr1,
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
            "ship_loc": ship_loc,
            "ship_pin": ship_pin,
            "ship_gstin": ship_gstin,
            "ship_state_name": ship_state_name,
            "ship_state_code": ship_state_code,
            "is_blinkit": is_blinkit,
            "printed_grand_total": printed_grand_total
        }

    def process_and_reconcile_pdf(pdf_bytes):
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        meta = extract_metadata_from_doc(doc)
        is_blinkit = meta["is_blinkit"]
        ext_order_str = meta["ext_order_id"]
        line_items = []

        for p_idx, page in enumerate(doc):
            words = page.get_text("words")

            # ACCURATE SINGLE-LINE OVERWRITE FOR EXTERNAL ORDER ID
            if p_idx == 0 and ext_order_str and ext_order_str != "ExtOrder":
                ext_matches = [w for w in words if "extern" in w[4].lower()]
                if ext_matches:
                    anchor = ext_matches[0]
                    # Mask exactly the order number area (without touching Invoice No below)
                    ext_num_box = fitz.Rect(anchor[2] + 2, anchor[1] - 0.5, anchor[2] + 75, anchor[3] + 10)
                    page.draw_rect(ext_num_box, color=None, fill=(1, 1, 1))
                    page.insert_text(
                        (anchor[2] + 3, anchor[3] - 1.0), 
                        str(ext_order_str), 
                        fontsize=7.2, 
                        fontname="helv", 
                        color=(0, 0, 0)
                    )

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

                    hsn_x0 = hsn_w[0] if hsn_w else 220

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
                            if len(val) >= 9:
                                continue
                            after_hsn_nums.append({
                                "val": float(val),
                                "str": val,
                                "rect": fitz.Rect(w[0], w[1], w[2], w[3])
                            })

                    final_q = None
                    final_p = None
                    final_disc = 0.0
                    final_taxable = None

                    if len(after_hsn_nums) >= 2:
                        final_q = int(after_hsn_nums[0]["val"])
                        if factor > 1 and final_q >= factor:
                            final_q = final_q // factor
                            overwrite_exact_area(page, after_hsn_nums[0]["rect"], f"{final_q}", font_size=7.2)

                        final_p = after_hsn_nums[1]["val"]
                        if factor > 1:
                            final_p = round(final_p * factor, 2)
                            overwrite_exact_area(page, after_hsn_nums[1]["rect"], f"{final_p:.2f}", font_size=7.2)

                        if len(after_hsn_nums) >= 3:
                            final_disc = after_hsn_nums[2]["val"]

                        if len(after_hsn_nums) >= 4:
                            final_taxable = after_hsn_nums[3]["val"]

                    if final_q is not None and final_p is not None and final_q > 0:
                        if final_taxable is not None and final_taxable > 0:
                            taxable_val = final_taxable
                        else:
                            taxable_val = round((final_q * final_p) - final_disc, 2)

                        gross_amt = round(taxable_val + final_disc, 2)
                        unit_label = "BOX" if (is_blinkit and factor > 1) else "PAC"
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
                            "gst_rate": 5.0
                        })

                    # SAFE IN-LINE UOM UPDATE
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

        exact_calc_total = round(tot_taxable + tot_cgst + tot_sgst + tot_igst, 2)
        final_rounded_total = round(exact_calc_total)
        round_off_amt = round(final_rounded_total - exact_calc_total, 2)

        ship_dtls = None
        if data.get("shipping_is_different", False):
            ship_dtls = {
                "Gstin": data["ship_gstin"],
                "LglNm": data["ship_name"],
                "TrdNm": data["ship_name"],
                "Addr1": data["ship_addr1"][:100],
                "Loc": data["ship_loc"],
                "Pin": data["ship_pin"],
                "Stcd": data["ship_state_code"]
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
                "Addr1": f"Warehouse Facility, {data['seller_loc']}",
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
                "RndOffAmt": round_off_amt,
                "TotInvVal": round(final_rounded_total, 2)
            }
        }

    def generate_official_nic_v101_excel_bytes(data_list):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "eInvoice"
        ws.views.sheetView[0].showGridLines = True

        c_supply = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        c_doc = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
        c_buyer = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
        c_disp = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
        c_ship = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
        c_item = PatternFill(start_color="D8E4BC", end_color="D8E4BC", fill_type="solid")
        c_val = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
        c_export = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
        c_eway = PatternFill(start_color="E6B8B7", end_color="E6B8B7", fill_type="solid")

        font_title = Font(name="Arial", size=18, bold=True)
        font_btn = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        font_sec = Font(name="Arial", size=10, bold=True)
        font_col = Font(name="Arial", size=9, bold=True)
        font_data = Font(name="Arial", size=9)

        thin_border = Border(
            left=Side(style='thin', color='B0B0B0'),
            right=Side(style='thin', color='B0B0B0'),
            top=Side(style='thin', color='B0B0B0'),
            bottom=Side(style='thin', color='B0B0B0')
        )

        ws.merge_cells("H1:L2")
        ws["H1"] = "E-Invoice System"
        ws["H1"].font = font_title
        ws["H1"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("M1:N2")
        ws["M1"] = "Validate"
        ws["M1"].fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        ws["M1"].font = font_btn
        ws["M1"].alignment = Alignment(horizontal="center", vertical="center")

        sections = [
            ("Supply Details", 1, 4, c_supply),
            ("Document Details", 5, 7, c_doc),
            ("Buyer Details", 8, 18, c_buyer),
            ("Dispatch Details", 19, 24, c_disp),
            ("Shipping Details", 25, 32, c_ship),
            ("Product Details", 33, 57, c_item),
            ("Value Details", 61, 70, c_val),
            ("Export Details", 71, 77, c_export),
            ("E-way-bill Details", 78, 85, c_eway)
        ]

        for name, sc, ec, fill in sections:
            ws.merge_cells(start_row=3, start_column=sc, end_row=3, end_column=ec)
            c = ws.cell(row=3, column=sc, value=name)
            c.fill = fill
            c.font = font_sec
            c.alignment = Alignment(horizontal="center", vertical="center")
            for i in range(sc, ec + 1):
                ws.cell(row=3, column=i).border = thin_border

        col_names = [
            "Supply Type Code *", "Reverse Charge", "e-Comm GSTIN", "Igst On Intra",
            "Document Type *", "Document Number *", "Document Date (DD/MM/YYYY) *",
            "Buyer GSTIN *", "Buyer Legal Name *", "Buyer Trade Name", "Buyer POS *",
            "Buyer Addr1 *", "Buyer Addr2", "Buyer Location *", "Buyer Pin Code *",
            "Buyer State *", "Buyer Phone Number", "Buyer Email Id",
            "Dispatch Name", "Dispatch Addr1", "Dispatch Addr2", "Dispatch Location", "Dispatch Pin Code", "Dispatch State",
            "Shipping GSTIN", "Shipping Legal Name", "Shipping Trade Name", "Shipping Addr1", "Shipping Addr2", "Shipping Location", "Shipping Pin Code", "Shipping State",
            "Sl.No. *", "Product Description", "Is Service *", "HSN Code *", "Bar Code", "Quantity *", "Free Quantity", "Unit *", "Unit Price *",
            "Gross Amount", "Discount", "Pre Tax Value", "Taxable value *", "GST Rate (%) *", "Sgst Amt(Rs)", "Cgst Amt(Rs)",
            "Igst Amt(Rs)", "Cess Rate (%)", "Cess Amt Adval (Rs)", "Cess Non Adval Amt (Rs)", "State Cess Rate (%)",
            "State Cess Adval Amt (Rs)", "State Cess Non-Adval Amt (Rs)", "Other Charges", "Item Total *",
            "", "", "",
            "Total Taxable value *", "Sgst Amt", "Cgst Amt", "Igst Amt", "Cess Amt", "State Cess Amt", "Discount", "Other charges", "Round off", "Total Invoice value *",
            "Shipping Bill No", "Shipping Bill Dt", "Port", "Refund claim", "Foreign Currency", "Country Code", "Export Duty Amount",
            "Trans ID", "Trans Name", "Trans Mode", "Distance", "Trans Doc No", "Trans Doc Date", "Vehicle No", "Vehicle Type"
        ]

        for c_idx, col_name in enumerate(col_names, 1):
            cell = ws.cell(row=4, column=c_idx, value=col_name)
            if c_idx <= 4:
                cell.fill = c_supply
            elif c_idx <= 7:
                cell.fill = c_doc
            elif c_idx <= 18:
                cell.fill = c_buyer
            elif c_idx <= 24:
                cell.fill = c_disp
            elif c_idx <= 32:
                cell.fill = c_ship
            elif c_idx <= 57:
                cell.fill = c_item
            elif 61 <= c_idx <= 70:
                cell.fill = c_val
            elif 71 <= c_idx <= 77:
                cell.fill = c_export
            elif 78 <= c_idx <= 85:
                cell.fill = c_eway
            cell.font = font_col
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        ws.row_dimensions[4].height = 42

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

                item_tot = round(taxable + cgst + sgst + igst, 2)
                item_rows_data.append({
                    "it": it, "qty": qty, "price": price, "gross_amt": gross_amt,
                    "disc": disc, "taxable": taxable, "gst_rate": gst_rate,
                    "cgst": cgst, "sgst": sgst, "igst": igst, "item_tot": item_tot
                })

            exact_total = round(tot_taxable + tot_cgst + tot_sgst + tot_igst, 2)
            final_inv_val = round(exact_total)
            exact_round_off = round(final_inv_val - exact_total, 2)

            if inv.get("shipping_is_different", False):
                ship_vals = [
                    str(inv["ship_gstin"]), str(inv["ship_name"]), str(inv["ship_name"]),
                    str(inv["ship_addr1"]), "", str(inv["ship_loc"]), str(inv["ship_pin"]), str(inv["ship_state_name"])
                ]
            else:
                ship_vals = ["", "", "", "", "", "", "", ""]

            for s_no, r_data in enumerate(item_rows_data, 1):
                it = r_data["it"]
                row_data = [
                    "B2B", "N", "", "N",
                    "Tax Invoice", str(inv["invoice_no"]), str(inv["doc_date"]),
                    str(inv["buyer_gstin"]), str(inv["buyer_name"]), str(inv["buyer_trade_name"]), pos_str,
                    str(inv["buyer_addr1"]), "", str(inv["buyer_loc"]), str(inv["buyer_pin"]),
                    str(inv["buyer_state_name"]), str(inv["buyer_phone"]), str(inv["buyer_email"]),
                    "", "", "", "", "", "",
                    *ship_vals,
                    str(s_no), str(it["desc"]), "N", str(it["hsn"]), "", str(r_data["qty"]), "0",
                    str(it.get("unit", "PAC")), fmt_dec(r_data["price"]), fmt_dec(r_data["gross_amt"]),
                    fmt_dec(r_data["disc"]), "", fmt_dec(r_data["taxable"]), str(int(r_data["gst_rate"])),
                    fmt_dec(r_data["sgst"]), fmt_dec(r_data["cgst"]), fmt_dec(r_data["igst"]),
                    "0", "0.00", "0.00", "0", "0.00", "0.00", "0.00",
                    fmt_dec(r_data["item_tot"]),
                    "", "", "",
                    fmt_dec(tot_taxable), fmt_dec(tot_sgst), fmt_dec(tot_cgst), fmt_dec(tot_igst),
                    "0.00", "0.00", fmt_dec(tot_discount), "0.00", fmt_dec(exact_round_off), fmt_dec(final_inv_val),
                    "", "", "", "", "", "", "",
                    "", "", "", "", "", "", "", ""
                ]

                for c_idx, val in enumerate(row_data, 1):
                    cell = ws.cell(row=curr_row, column=c_idx, value=str(val))
                    cell.font = font_data
                    cell.border = thin_border
                    cell.number_format = '@'
                    if c_idx in [1, 2, 4, 5, 7, 8, 11, 15, 16, 17, 25, 31, 32, 33, 35, 36, 38, 40, 46]:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif c_idx in [41, 42, 43, 45, 47, 48, 49, 57, 61, 62, 63, 64, 67, 69, 70]:
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

        tabs = ["Welcome", "Profile", "Master Codes", "Sample Invoice", "Format A,B,C,D", "Schema", "Validation", "Calculations", "FAQs"]
        for t in tabs:
            d_ws = wb.create_sheet(title=t)
            d_ws.sheet_view.showGridLines = True
            d_ws["A1"] = f"{t} - Official e-Invoice System Utility"
            d_ws["A1"].font = Font(size=14, bold=True, color="1F497D")

        wb.active = ws
        out_io = io.BytesIO()
        wb.save(out_io)
        return out_io.getvalue()

    def render_exact_government_einvoice_preview(meta):
        is_interstate = meta["is_interstate"]
        tot_taxable = sum([it['taxable_val'] for it in meta['line_items']])
        tot_discount = sum([it.get('discount', 0.0) for it in meta['line_items']])

        if is_interstate:
            tot_igst = round(tot_taxable * 0.05, 2)
            tot_cgst = 0.0
            tot_sgst = 0.0
        else:
            tot_igst = 0.0
            tot_cgst = round(tot_taxable * 0.025, 2)
            tot_sgst = round(tot_taxable * 0.025, 2)

        exact_total = round(tot_taxable + tot_cgst + tot_sgst + tot_igst, 2)
        grand_total = round(exact_total)
        round_off = round(grand_total - exact_total, 2)
        ack_date_str = f"{meta['clean_date']} 17:00:00"

        table_rows = ""
        for idx, it in enumerate(meta['line_items'], 1):
            q = it['qty']
            p = it['unit_price']
            taxable = it['taxable_val']
            disc = it.get('discount', 0.0)
            gst_r = it.get('gst_rate', 5.0)
            tax_amt = round(taxable * (gst_r / 100), 2)
            tot_val = round(taxable + tax_amt, 2)

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
                    <div>{meta['seller_addr']}</div>
                    <div>{meta['seller_loc'].upper()}</div>
                    <div>{meta['seller_pin']} &nbsp; {meta['seller_state_name'].upper()}</div>
                    <div style="margin-top: 3px;">+91-7070701513 &nbsp; info@romsons.in</div>
                </td>
                <td style="width: 50%; padding: 6px; vertical-align: top;">
                    <div style="font-weight: bold; font-size: 13px; text-decoration: underline; margin-bottom: 3px;">Recipient</div>
                    <div><b>GSTIN :</b> {meta['buyer_gstin']}</div>
                    <div style="font-weight: bold; margin-top: 2px;">{meta['buyer_name']}</div>
                    <div>{meta['buyer_addr1']}</div>
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
                    <tr><td>Round Off :</td><td style="text-align: right;">₹{round_off:.2f}</td></tr>
                    <tr style="border-top: 1.5px solid #000; font-size: 13px; font-weight: bold;"><td>Total Invoice Value :</td><td style="text-align: right;">₹{grand_total:.2f}</td></tr>
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
        "Upload Blinkit Invoices (PDF) - Single ya Bulk", 
        type=["pdf"], 
        accept_multiple_files=True,
        key=f"blnk_inv_{st.session_state.selected_wh}_{st.session_state.uploader_key}"
    )

    if uploaded_invoices:
        processed_docs = []
        for file in uploaded_invoices:
            try:
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
            st.success(f"✅ Successfully processed {len(processed_docs)} Invoices!")
            parsed_invoices = [doc["meta"] for doc in processed_docs]
            excel_raw_bytes = generate_official_nic_v101_excel_bytes(parsed_invoices)

            if len(processed_docs) == 1:
                doc = processed_docs[0]
                meta = doc["meta"]
                pdf_buf = doc["pdf_buf"]
                json_payload = build_einvoice_json_v101(meta)
                inv_no = meta["invoice_no"]
                pdf_name = meta["pdf_filename"]
                is_blinkit = meta["is_blinkit"]

                channel_badge = "🟢 Blinkit Conversion Active (50x/25x Box Fix)" if is_blinkit else "🔵 Standard Buyer"
                ship_status = "⚠️ Different (Shipping Details Filled)" if meta["shipping_is_different"] else "✅ Same (Shipping Blank in Excel)"

                st.markdown(f"""
                    <div style="background-color: #1e293b; padding: 14px 18px; border-radius: 8px; margin-bottom: 15px; border-left: 5px solid {'#10b981' if is_blinkit else '#3b82f6'};">
                        <b>Buyer Name:</b> <code>{meta['buyer_name']}</code> &nbsp;|&nbsp; <b>Rule:</b> <code>{channel_badge}</code><br>
                        <b>City:</b> <code>{meta['buyer_loc']}</code> ({meta['buyer_pin']}) &nbsp;|&nbsp; <b>Shipping Status:</b> <code>{ship_status}</code><br>
                        <b>Seller:</b> {meta['seller_name']} ({meta['seller_gstin']})
                    </div>
                """, unsafe_allow_html=True)

                hist_item = {
                    "original_name": doc["original_name"],
                    "file_name": pdf_name,
                    "data": pdf_buf.getvalue(),
                    "pages": 1,
                    "stamped": len(meta["line_items"]),
                    "location": st.session_state.selected_wh,
                    "module": "BLNK",
                    "time": get_current_ist_str()
                }

                c1, c2, c3 = st.columns(3)
                with c1:
                    st.download_button(
                        label="📥 Download Processed PDF",
                        data=pdf_buf.getvalue(),
                        file_name=pdf_name,
                        mime="application/pdf",
                        on_click=record_to_history,
                        args=(hist_item,),
                        key="dl_blnk_single_pdf"
                    )
                with c2:
                    st.download_button(
                        label="📊 Download eInvoice.xlsx (Exact Template)",
                        data=excel_raw_bytes,
                        file_name=f"{inv_no}_eInvoice.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_blnk_single_xlsx"
                    )
                with c3:
                    st.download_button(
                        label="🧾 Download e-Invoice JSON (v1.01)",
                        data=json.dumps(json_payload, indent=4),
                        file_name=f"{inv_no}_eInvoice_v1.01.json",
                        mime="application/json",
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
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for idx, doc in enumerate(processed_docs, 1):
                        pdf_name = doc["meta"]["pdf_filename"]
                        raw_data = doc["pdf_buf"].getvalue()
                        zip_file.writestr(pdf_name, raw_data)
                        batch_hist.append({
                            "original_name": doc["original_name"],
                            "file_name": pdf_name,
                            "data": raw_data,
                            "pages": 1,
                            "stamped": len(doc["meta"]["line_items"]),
                            "location": st.session_state.selected_wh,
                            "module": "BLNK",
                            "time": get_current_ist_str()
                        })
                zip_buffer.seek(0)

                bulk_json = [build_einvoice_json_v101(d["meta"]) for d in processed_docs]
                date_str = datetime.now(IST).strftime('%d-%m-%Y')

                c1, c2, c3 = st.columns(3)
                with c1:
                    st.download_button(
                        label=f"📦 Download All Processed PDFs ({len(processed_docs)} Files - ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"Invoices_Processed_{date_str}.zip",
                        mime="application/zip",
                        on_click=record_batch_to_history,
                        args=(batch_hist,),
                        key="dl_blnk_bulk_zip"
                    )
                with c2:
                    st.download_button(
                        label=f"📊 Download Bulk eInvoice.xlsx ({len(processed_docs)} Invoices)",
                        data=excel_raw_bytes,
                        file_name=f"Bulk_eInvoice_{date_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_blnk_bulk_xlsx"
                    )
                with c3:
                    st.download_button(
                        label=f"🧾 Download Bulk e-Invoice JSON ({len(bulk_json)} Invoices)",
                        data=json.dumps(bulk_json, indent=4),
                        file_name=f"Bulk_eInvoice_NIC_v1.01_{date_str}.json",
                        mime="application/json",
                        key="dl_blnk_bulk_json"
                    )

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("### 📄 Invoices Official e-Invoice Preview")
                preview_tabs = st.tabs([f"INV: {d['meta']['invoice_no']}" for d in processed_docs])
                for i, tab in enumerate(preview_tabs):
                    with tab:
                        curr_meta = processed_docs[i]['meta']
                        calc_height = 560 + (len(curr_meta['line_items']) * 36)
                        p_html = render_exact_government_einvoice_preview(curr_meta)
                        components.html(p_html, height=calc_height, scrolling=True)

# =========================================================================
# MODULE 4: LABEL EDITOR
# =========================================================================
elif st.session_state.active_tab == "LBL":
    st.subheader("🏷️ Shipping & Product Label Editor")
    st.info("No files")

# =========================================================================
# MODULE 5: SETTINGS & REPORTS
# =========================================================================
elif st.session_state.active_tab == "CONF":
    st.subheader("⚙️ System Operations & Units Configuration")
    st.markdown(f"""
    - **Active Location:** `{LOCATION_DISPLAY_MAP.get(st.session_state.selected_wh, st.session_state.selected_wh)}`
    - **Default Barcode Type:** Code128 (High DPI Crisp Vector)
    - **Engine Status:** Online & Healthy
    """)
