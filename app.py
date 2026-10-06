import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import pymupdf as fitz
import re
import io
import json
import zipfile
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import barcode
from barcode.writer import ImageWriter

# Page Configuration - Enterprise Wide Layout
st.set_page_config(
    page_title="Delhi Operations Hub",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Universal CSS: Light & Dark Mode Support + Romsons Branding Watermark
st.markdown("""
    <style>
    /* Metric Cards */
    .stMetric {
        background-color: rgba(28, 37, 65, 0.08);
        padding: 10px;
        border-radius: 8px;
        border: 1px solid rgba(148, 163, 184, 0.2);
    }
    
    /* Background Watermark */
    .stApp::before {
        content: "Romsons";
        position: fixed;
        top: 50%;
        left: 55%;
        transform: translate(-50%, -50%) rotate(-12deg);
        font-family: 'Brush Script MT', 'Lucida Handwriting', cursive, sans-serif;
        font-size: 14vw;
        font-weight: 900;
        color: rgba(11, 79, 59, 0.035);
        pointer-events: none;
        z-index: 0;
        white-space: nowrap;
        user-select: none;
    }

    /* Logo Container in Sidebar */
    .brand-logo-card {
        background: #ffffff;
        border-radius: 10px;
        padding: 12px 10px;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
        margin-bottom: 12px;
    }
    .brand-logo-card .logo-title {
        font-family: 'Brush Script MT', 'Lucida Handwriting', cursive, sans-serif;
        font-size: 32px;
        font-weight: 900;
        color: #0b4f3b;
        letter-spacing: -0.5px;
        margin: 0;
        line-height: 1;
    }
    .brand-logo-card .logo-tagline {
        font-family: Arial, Helvetica, sans-serif;
        font-size: 10px;
        font-weight: 700;
        color: #222222;
        letter-spacing: 0.5px;
        margin-top: 4px;
        text-transform: none;
    }

    /* High Visibility Radio Nav Buttons */
    div[data-testid="stRadio"] > div {
        gap: 6px;
    }
    div[data-testid="stRadio"] label {
        padding: 10px 14px;
        border-radius: 8px;
        border: 1px solid rgba(148, 163, 184, 0.4);
        cursor: pointer;
        font-weight: 600;
        transition: all 0.2s ease-in-out;
    }
    div[data-testid="stRadio"] label p {
        font-size: 14px !important;
        font-weight: 600 !important;
    }
    </style>
""", unsafe_allow_html=True)

# ----------------- SIDEBAR NAVIGATION -----------------
with st.sidebar:
    st.title("🏛️ Delhi Operations Hub")
    st.divider()

    # Romsons Brand Logo Header
    st.markdown("""
        <div class="brand-logo-card">
            <div class="logo-title">Romsons</div>
            <div class="logo-tagline">Sustaining the life force</div>
        </div>
    """, unsafe_allow_html=True)

    st.divider()
    
    selected_module = st.radio(
        label="Select Workspace",
        options=[
            "📑 Amazon Invoice Editor",
            "⚡ Blinkit e-Invoice Tool",
            "🏷️ Label Editor",
            "⚙️ Unit Settings"
        ],
        index=0,
        label_visibility="collapsed"
    )

# =========================================================================
# MODULE 1: AMAZON INVOICE EDITOR
# =========================================================================
if selected_module == "📑 Amazon Invoice Editor":
    st.subheader("📑 Amazon Invoice Editor")

    TARGET_PAN = "aalcr5906l"

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        uploaded_csv = st.file_uploader("Upload Shipment Report (CSV / Excel)", type=["csv", "xlsx", "xls"], key="amz_csv")
    with col_u2:
        uploaded_pdfs = st.file_uploader("Upload Invoice PDF(s)", type=["pdf"], accept_multiple_files=True, key="amz_pdf")

    def clean_val(v):
        if pd.isna(v):
            return ""
        s = str(v).strip()
        s = re.sub(r'^[="\']+|["\']+$', '', s)
        return s.strip()

    def clean_alphanumeric(text):
        return re.sub(r'[^a-zA-Z0-9]', '', str(text)).lower()

    def get_token_words(text):
        words = re.findall(r'[a-zA-Z0-9]+', str(text).lower())
        stop_words = {'the', 'and', 'for', 'with', 'pcs', 'piece', 'pieces', 'only', 'total', 'hsn', 'gst', 'rs', 'inr'}
        return set([w for w in words if len(w) >= 2 and w not in stop_words])

    def generate_barcode_image(code_text):
        try:
            code128 = barcode.get_barcode_class('code128')
            writer = ImageWriter()
            writer.font_path = None
            barcode_instance = code128(code_text, writer=writer)
            
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
        order_col = next((col_mapping[c] for c in col_mapping if "order" in c), None)
        tracking_col = next((col_mapping[c] for c in col_mapping if "track" in c or "tracing" in c), None)
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

        if st.button("🚀 Process Invoices", type="primary", use_container_width=True):
            progress_bar = st.progress(0)
            status_text = st.empty()

            processed_files = []
            total_files = len(uploaded_pdfs)

            for file_idx, pdf_file in enumerate(uploaded_pdfs):
                status_text.text(f"Processing File {file_idx+1}/{total_files}...")

                pdf_bytes = pdf_file.read()
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                new_doc = fitz.open()

                total_pages = len(doc)
                file_stamped_count = 0
                file_removed_pan = 0

                for page_num in range(total_pages):
                    page = doc[page_num]
                    raw_text = page.get_text()
                    text_lower = raw_text.lower()

                    if TARGET_PAN not in text_lower:
                        file_removed_pan += 1
                        continue

                    order_clean = None
                    order_match = re.search(r'(\d{3})\s*[-–—]\s*(\d{7})\s*[-–—]\s*(\d{7})', raw_text)
                    if order_match:
                        order_clean = f"{order_match.group(1)}{order_match.group(2)}{order_match.group(3)}"
                    else:
                        num_match = re.search(r'order\s*number\s*[:\s]*(\d{3}[-–—\d]{14,16}\d)', raw_text, re.IGNORECASE)
                        if num_match:
                            order_clean = clean_alphanumeric(num_match.group(1))

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

                processed_files.append({
                    "original_name": pdf_file.name,
                    "file_name": f"Stamped_{pdf_file.name}",
                    "data": out_buf.getvalue(),
                    "pages": len(new_doc),
                    "stamped": file_stamped_count,
                    "removed": file_removed_pan
                })

                progress_bar.progress((file_idx + 1) / total_files)

            status_text.empty()
            progress_bar.empty()

            st.balloons()
            st.success(f"🎉 **Total {len(processed_files)} File(s) Processed Successfully!**")

            if len(processed_files) > 1:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for item in processed_files:
                        zip_file.writestr(item["file_name"], item["data"])
                zip_buffer.seek(0)

                st.download_button(
                    label="📦 Download All Invoices as ZIP",
                    data=zip_buffer,
                    file_name="All_Stamped_Invoices.zip",
                    mime="application/zip",
                    use_container_width=True
                )

            st.write("---")
            for idx, item in enumerate(processed_files):
                d_col1, d_col2 = st.columns([3, 1])
                with d_col1:
                    st.write(f"📁 **{item['original_name']}** — `{item['pages']} Pages Kept` | `{item['stamped']} Stamped` | `{item['removed']} Filtered`")
                with d_col2:
                    st.download_button(
                        label="📥 Download PDF",
                        data=item["data"],
                        file_name=item["file_name"],
                        mime="application/pdf",
                        key=f"dl_btn_{idx}",
                        use_container_width=True
                    )

# =========================================================================
# MODULE 2: BLINKIT E-INVOICE TOOL (UNIVERSAL OPERATIONS & E-INVOICE SUITE)
# =========================================================================
elif selected_module == "⚡ Blinkit e-Invoice Tool":
    st.subheader("⚡ Blinkit Bulk Invoice Gateway & e-Invoice Engine")

    uploaded_invoices = st.file_uploader("Upload Blinkit Invoices (PDF) - Single ya Bulk", type=["pdf"], accept_multiple_files=True, key="blinkit_uploader")

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
            if parts:
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
        desc = re.sub(r'(\b[A-Za-z]+)\s+([a-z]{1,4}\b)', lambda m: m.group
