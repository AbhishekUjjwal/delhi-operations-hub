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

# Custom Styling: Vibrant, Clean & Theme-Adaptive
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
    
    /* Pill buttons */
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

    /* Single line gradient title */
    .hub-title {
        font-size: 20px;
        font-weight: 800;
        white-space: nowrap;
        background: linear-gradient(90deg, #38bdf8 0%, #818cf8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 2px;
        display: flex;
        align-items: center;
        gap: 8px;
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

# ----------------- SIDEBAR NAVIGATION -----------------
with st.sidebar:
    # Single-line stylish title
    st.markdown('<div class="hub-title">🏛️ PAN OPERATIONS HUB</div>', unsafe_allow_html=True)
    st.caption("Romsons prime pvt ltd operation portal")
    st.divider()

    # Short Name + Logo only (Clean ERP Look)
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

# ----------------- TOP ACTION BAR: PINNED TABS, WH LOCATION, HISTORY & RESET -----------------
top_tabs_col, top_wh_col, top_hist_col, top_reset_col = st.columns([5, 2.4, 1.8, 1.2])

with top_tabs_col:
    tab_cols = st.columns(len(st.session_state.open_tabs) + 1)
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
                    # Clean Cancel Mark (✕)
                    if st.button("✕", key=f"close_tab_{t_key}", help="Close Tab"):
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
# MODULE 1: AMAZON INVOICE EDITOR (100% UNTOUCHED CORE LOGIC)
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

                    # Rule 2: Description Check Filter (Remove if no description)
                    if not has_valid_description(raw_text):
                        file_removed_desc += 1
                        continue

                    # Extract 17-digit Order Number
                    order_clean = None
                    order_match = re.search(r'(\d{3})\s*[-–—\s]\s*(\d{7})\s*[-–—\s]\s*(\d{7})', raw_text)
                    if order_match:
                        order_clean = f"{order_match.group(1)}{order_match.group(2)}{order_match.group(3)}"
                    else:
                        num_match = re.search(r'order\s*number\s*[:\s]*([0-9\-–—\s]{17,25})', raw_text, re.IGNORECASE)
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

                    # Stamping Barcode and Tracking ID
                    if target_tracking_id:
                        barcode_rect = fitz.Rect(40, 58, 235, 82)
                        page.draw_rect(barcode_rect, color=(1.0, 1.0, 1.0), fill=(1.0, 1.0, 1.0), width=0)

                        barcode_img_bytes = generate_barcode_image(target_tracking_id)
                        if barcode_img_bytes:
                            page.insert_image(barcode_rect, stream=barcode_img_bytes, keep_proportion=False)

                        page.insert_text(
                            (barcode_rect.x0 + 10, 95),
                            f"TRACKING: {target_tracking_id}",
                            fontsize=10.5,
                            fontname="hebo",
                            color=(0, 0, 0)
                        )

                        date_instances = page.search_for("Order Date:") or page.search_for("Order Date")
                        if date_instances:
                            first_date_rect = date_instances[0]
                            page.insert_text(
                                (first_date_rect.x0, first_date_rect.y1 + 13),
                                f"Tracking ID: {target_tracking_id}",
                                fontsize=9.5,
                                fontname="hebo",
                                color=(0, 0, 0)
                            )

                        file_stamped_count += 1

                    new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)

                out_buf = io.BytesIO()
                new_doc.save(out_buf)
                out_buf.seek(0)

                current_processed.append({
                    "original_name": pdf_file.name,
                    "file_name": f"Stamped_{pdf_file.name}",
                    "data": out_buf.getvalue(),
                    "pages": len(new_doc),
                    "stamped": file_stamped_count,
                    "removed_pan": file_removed_pan,
                    "removed_desc": file_removed_desc,
                    "location": st.session_state.selected_wh
                })

                progress_bar.progress((file_idx + 1) / total_files)

            status_text.empty()
            progress_bar.empty()

            st.session_state.processed_history.extend(current_processed)
            st.balloons()
            st.success(f"🎉 **Total {len(current_processed)} File(s) Processed for {st.session_state.selected_wh}!**")

    # Quick download list for current location
    if loc_history:
        st.write("---")
        st.subheader("📄 Saved Files:")
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
                    key=f"dl_quick_{st.session_state.selected_wh}_{idx}",
                    use_container_width=True
                )

# =========================================================================
# MODULE 2: BLINKIT E-INVOICE TOOL
# =========================================================================
elif st.session_state.active_tab == "BLNK":
    st.subheader(f"⚡ Blinkit e-Invoice Management Tool [{LOCATION_DISPLAY_MAP.get(st.session_state.selected_wh, st.session_state.selected_wh)}]")
    b_col1, b_col2 = st.columns(2)
    with b_col1:
        st.file_uploader("Upload Blinkit PO / Data Sheet (Excel/CSV)", type=["xlsx", "xls", "csv"], key="blinkit_data")
    with b_col2:
        st.file_uploader("Upload Master Template / Supporting PDF (Optional)", type=["pdf", "json"], key="blinkit_template")

# =========================================================================
# MODULE 3: LABEL EDITOR
# =========================================================================
elif st.session_state.active_tab == "LBL":
    st.subheader(f"🏷️ Shipping & Product Label Editor [{LOCATION_DISPLAY_MAP.get(st.session_state.selected_wh, st.session_state.selected_wh)}]")
    st.info("Label Editor module ready hai.")

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
