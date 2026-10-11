import streamlit as st
import pandas as pd
import pymupdf as fitz
import re
import io
import zipfile
import barcode
from barcode.writer import ImageWriter

# Page Configuration - Original Enterprise Dark Dashboard
st.set_page_config(
    page_title="Delhi Operations Hub | Central Portal",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling: Sleek ERP / Enterprise Dashboard
st.markdown("""
    <style>
    .main { background-color: #0b132b; color: #f8fafc; }
    
    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        background-color: #1c2541;
        border-right: 1px solid #3a506b;
    }
    
    /* Metrics card */
    .stMetric {
        background-color: #1c2541;
        padding: 12px;
        border-radius: 8px;
        border: 1px solid #3a506b;
    }
    
    /* Top action selectbox label hide */
    div[data-testid="stSelectbox"] label {
        display: none;
    }
    
    /* Custom button styling */
    div[data-testid="stButton"] button {
        border-radius: 20px;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

ORDERED_LOCATIONS = [
    "RPPL - DEL",
    "RPPL - KOL",
    "RPPL - HYD",
    "RPPL - PUN",
    "RPPL - BLR",
    "RPPL - CHE",
    "Headoffice"
]

AVAILABLE_MODULES = {
    "AMZ": {"name": "Amazon Invoice Editor", "code": "AMZ-ED", "icon": "📑"},
    "BLNK": {"name": "Blinkit e-Invoice Tool", "code": "BLNK-INV", "icon": "⚡"},
    "LBL": {"name": "Label Editor", "code": "LBL-ED", "icon": "🏷️"},
    "CONF": {"name": "Settings & Reports", "code": "SETTINGS", "icon": "⚙️"}
}

# ----------------- SESSION STATE -----------------
if "selected_wh" not in st.session_state:
    st.session_state.selected_wh = "RPPL - DEL"
if "active_module" not in st.session_state:
    st.session_state.active_module = "AMZ"
if "show_history_view" not in st.session_state:
    st.session_state.show_history_view = False
if "processed_history" not in st.session_state:
    st.session_state.processed_history = []
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

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
    st.session_state.show_history_view = False
    st.rerun()

# ----------------- SIDEBAR NAVIGATION -----------------
with st.sidebar:
    st.title("🏛️ Delhi Operations Hub")
    st.caption("Romsons prime pvt ltd operation portal")
    st.divider()

    st.markdown("### 🧭 Module Navigation")
    
    for key, info in AVAILABLE_MODULES.items():
        is_active = (st.session_state.active_module == key and not st.session_state.show_history_view)
        label = f"{info['icon']} {info['name']}"
        if st.button(label, key=f"side_btn_{key}", use_container_width=True, type="primary" if is_active else "secondary"):
            st.session_state.active_module = key
            st.session_state.show_history_view = False
            st.rerun()

    st.divider()
    st.caption("System Version: v3.2 Enterprise")

# ----------------- TOP ACTION BAR (HISTORY & RESET & LOCATION) -----------------
top_title_col, top_loc_col, top_hist_col, top_reset_col = st.columns([5, 2.5, 1.8, 1.2])

with top_loc_col:
    st.session_state.selected_wh = st.selectbox(
        "WH | Location",
        options=ORDERED_LOCATIONS,
        index=ORDERED_LOCATIONS.index(st.session_state.selected_wh) if st.session_state.selected_wh in ORDERED_LOCATIONS else 0,
        help="Select Warehouse Location"
    )

current_history = get_current_location_history()
hist_count = len(current_history)

with top_hist_col:
    hist_btn_label = f"🕒 History ({hist_count})"
    if st.button(hist_btn_label, use_container_width=True, type="primary" if st.session_state.show_history_view else "secondary"):
        st.session_state.show_history_view = not st.session_state.show_history_view
        st.rerun()

with top_reset_col:
    if st.button("🔄 Reset", use_container_width=True, type="secondary"):
        reset_current_workspace()

st.divider()

# =========================================================================
# VIEW 1: HISTORY DRAWER / VIEW (JAB HISTORY CLICKED HO)
# =========================================================================
if st.session_state.show_history_view:
    st.subheader(f"📜 History & Processed Invoices — [{st.session_state.selected_wh}]")
    
    if not current_history:
        st.info(f"Abhi tak **{st.session_state.selected_wh}** ke liye koi file process nahi hui hai.")
    else:
        st.success(f"💾 Total Saved Files: **{len(current_history)}**")

        if len(current_history) > 1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, item in enumerate(current_history):
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
        for idx, item in enumerate(current_history):
            d_col1, d_col2 = st.columns([3, 1])
            with d_col1:
                st.write(
                    f"📁 **{item['original_name']}** — `[Location: {item.get('location')}]` | "
                    f"`{item['pages']} Pages Kept` | `{item['stamped']} Stamped` | "
                    f"`{item['removed_pan']} Non-PAN Filtered` | `{item.get('removed_desc', 0)} No-Desc Filtered`"
                )
            with d_col2:
                st.download_button(
                    label="📥 Download PDF",
                    data=item["data"],
                    file_name=item["file_name"],
                    mime="application/pdf",
                    key=f"dl_history_btn_{idx}",
                    use_container_width=True
                )

# =========================================================================
# VIEW 2: ORIGINAL MAIN DESKTOP WORKSPACE (AMAZON INVOICE EDITOR)
# =========================================================================
elif st.session_state.active_module == "AMZ":
    st.subheader(f"📑 Amazon Invoice Editor & Barcode Stamper")

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
        m1.metric("Order Column", str(order_col))
        m2.metric("Product Description Column", str(title_col if title_col else sku_col))
        m3.metric("Tracking Column", str(tracking_col))

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
            st.success(f"🎉 **Total {len(current_processed)} File(s) Processed Successfully!**")

    # Quick download list for recent operation
    if current_history:
        st.write("---")
        st.subheader("📄 Recent Stamped Files:")
        for idx, item in enumerate(current_history):
            d_col1, d_col2 = st.columns([3, 1])
            with d_col1:
                st.write(
                    f"📁 **{item['original_name']}** — `[Location: {item.get('location')}]` | "
                    f"`{item['pages']} Pages Kept` | `{item['stamped']} Stamped` | "
                    f"`{item['removed_pan']} Non-PAN Filtered` | `{item.get('removed_desc', 0)} No-Desc Filtered`"
                )
            with d_col2:
                st.download_button(
                    label="📥 Download PDF",
                    data=item["data"],
                    file_name=item["file_name"],
                    mime="application/pdf",
                    key=f"dl_quick_btn_{idx}",
                    use_container_width=True
                )

# =========================================================================
# VIEW 3: OTHER MODULES
# =========================================================================
elif st.session_state.active_module == "BLNK":
    st.subheader(f"⚡ Blinkit e-Invoice Management Tool [{st.session_state.selected_wh}]")
    b_col1, b_col2 = st.columns(2)
    with b_col1:
        st.file_uploader("Upload Blinkit PO / Data Sheet (Excel/CSV)", type=["xlsx", "xls", "csv"], key="blinkit_data")
    with b_col2:
        st.file_uploader("Upload Master Template / Supporting PDF (Optional)", type=["pdf", "json"], key="blinkit_template")

elif st.session_state.active_module == "LBL":
    st.subheader(f"🏷️ Shipping & Product Label Editor [{st.session_state.selected_wh}]")
    st.info("Label Editor module ready hai.")

elif st.session_state.active_module == "CONF":
    st.subheader("⚙️ System Operations & Units Configuration")
    st.markdown(f"""
    - **Active Location:** `{st.session_state.selected_wh}`
    - **Default Barcode Type:** Code128 (High DPI Crisp Vector)
    - **Engine Status:** Online & Healthy
    """)
