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

# Page Configuration
st.set_page_config(
    page_title="Delhi Operations Hub",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Universal Styling & Watermark
st.markdown("""
    <style>
    .stMetric {
        background-color: rgba(28, 37, 65, 0.08);
        padding: 10px;
        border-radius: 8px;
        border: 1px solid rgba(148, 163, 184, 0.2);
    }
    
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

# ----------------- SIDEBAR -----------------
with st.sidebar:
    st.title("🏛️ Delhi Operations Hub")
    st.divider()

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
# MODULE 1: AMAZON INVOICE EDITOR (DESCRIPTION MATCHING)
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

            for file_idx, pdf_file in enumerate(
