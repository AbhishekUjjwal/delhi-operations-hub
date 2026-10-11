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

# Custom Styling: Zero Top Padding, Compact Actions & Snappy Layout
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

    .right-action-cluster div[data-testid="stButton"] {
        width: 100% !important;
    }
    .right-action-cluster div[data-testid="stButton"] button {
        width: 100% !important;
    }

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
# HISTORY VIEW
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
# MODULE 1: AMAZON INVOICE EDITOR (ULTRA-FAST OPTIMIZED ENGINE)
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

    # Pre-compiled high performance regex
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

    # Memory Barcode Cache: Same barcode image renders once only
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
                    'dpi': 200  # High crisp scan without overhead
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

                # Lightning-Fast PDF stream save
                out_bytes = new_doc.tobytes(deflate=True, clean=False)
                new_doc.close()
                doc.close()

                processed_items.append({
                    "original_name": pdf_file.name,
                    "file_name": f"Stamped_{pdf_file.name}",
                    "data": out_bytes,
                    "pages": len(new_doc) if new_doc else file_stamped_count,
                    "stamped": file_stamped_count,
                    "removed_pan": file_removed_pan,
                    "removed_desc": file_removed_desc,
                    "location": st.session_state.selected_wh,
                    "module": "AMZ",
                    "time": datetime.now().strftime("%d-%b %H:%M")
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

    # =====================================================================
    # BULK & SINGLE DOWNLOAD SECTION
    # =====================================================================
    batch_list = st.session_state.current_batch_results
    if batch_list:
        st.write("---")
        st.subheader("📥 Download Invoices")

        # FAST ZIP CREATION
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
# MODULE 2: BLINKIT E-INVOICE TOOL
# =========================================================================
elif st.session_state.active_tab == "BLNK":
    st.subheader("⚡ Blinkit e-Invoice Management Tool")
    b_col1, b_col2 = st.columns(2)
    with b_col1:
        st.file_uploader("Upload Blinkit PO / Data Sheet (Excel/CSV)", type=["xlsx", "xls", "csv"], key="blinkit_data")
    with b_col2:
        st.file_uploader("Upload Master Template / Supporting PDF (Optional)", type=["pdf", "json"], key="blinkit_template")

# =========================================================================
# MODULE 3: LABEL EDITOR
# =========================================================================
elif st.session_state.active_tab == "LBL":
    st.subheader("🏷️ Shipping & Product Label Editor")
    st.info("No files")

# =========================================================================
# MODULE 4: SETTINGS & REPORTS
# =========================================================================
elif st.session_state.active_tab == "CONF":
    st.subheader("⚙️ System Operations & Units Configuration")
    st.markdown(f"""
    - **Active Location:** `{LOCATION_DISPLAY_MAP.get(st.session_state.selected_wh, st.session_state.selected_wh)}`
    - **Default Barcode Type:** Code128 (High DPI Crisp Vector)
    - **Engine Status:** Online & Healthy
    """)
