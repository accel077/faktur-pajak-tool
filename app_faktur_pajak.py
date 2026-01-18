import streamlit as st
import pdfplumber
from PyPDF2 import PdfMerger
import re
import io
import zipfile
from openpyxl import load_workbook

# --- FUNGSI LOGIKA ---

def identify_color_name(color_hex):
    if not color_hex or color_hex == "00000000":
        return "Tanpa_Warna"
    hex_val = color_hex[-6:].upper() if len(color_hex) > 6 else color_hex.upper()
    colors = {
        "FFFF00": "Kuning", "FFC000": "Orange", "ED7D31": "Orange",
        "92D050": "Light Green", "00B050": "Green", "0070C0": "Blue",
        "00B0F0": "Light Blue", "5B9BD5": "Light Blue"
    }
    return colors.get(hex_val, f"Warna_{hex_val}")

def get_color_mapping(excel_file):
    color_map = {}
    try:
        wb = load_workbook(excel_file, data_only=True)
        ws = wb.active
        ref_col_idx = None
        for cell in ws[1]:
            if cell.value == "Referensi":
                ref_col_idx = cell.column
                break
        if not ref_col_idx: return None
        for row in ws.iter_rows(min_row=2):
            cell = row[ref_col_idx-1]
            ref_val = str(cell.value).strip() if cell.value else None
            if ref_val and ref_val.startswith("PJ"):
                fill = cell.fill
                color_hex = fill.start_color.index if fill and hasattr(fill.start_color, 'index') else None
                color_name = identify_color_name(str(color_hex)) if color_hex else "Tanpa_Warna"
                color_map[ref_val] = color_name
        return color_map
    except: return None

def extract_referensi(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    match = re.search(r'PJ\d+', text)
                    if match: return match.group(0)
        return None
    except: return None

# --- TAMPILAN UTAMA ---

st.set_page_config(page_title="Faktur Pajak Tool", layout="centered")
st.title("📑 Faktur Pajak Tool")

# Sidebar untuk Database
with st.sidebar:
    st.header("Database")
    excel_db = st.file_uploader("Upload Excel Warna", type=["xlsx"])
    color_map = get_color_mapping(excel_db) if excel_db else None
    if color_map: st.success("Database Dimuat!")

tab1, tab2, tab3 = st.tabs(["Rename", "Klasifikasi", "Merge"])

# --- TAB 1: RENAME ---
with tab1:
    st.subheader("Rename File (PJ + 3 Kata Terakhir)")
    files = st.file_uploader("Pilih PDF", type="pdf", accept_multiple_files=True, key="ren")
    
    if st.button("Proses Rename") and files:
        rename_buffer = io.BytesIO()
        with zipfile.ZipFile(rename_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
            for f in files:
                ref = extract_referensi(f)
                if ref:
                    # Logika mengambil 3 bagian terakhir dari nama file asli
                    name_only = f.name.replace(".pdf", "").replace(".PDF", "")
                    parts = name_only.split('-')
                    suffix = "-".join(parts[-3:]) if len(parts) >= 3 else name_only
                    
                    new_name = f"{ref} {suffix}.pdf"
                    zip_f.writestr(new_name, f.getvalue())
                    st.success(f"✅ {f.name} -> **{new_name}**")
        
        st.download_button("⬇️ Download ZIP Hasil Rename", rename_buffer.getvalue(), "hasil_rename.zip", "application/zip", use_container_width=True)

# --- TAB 2: KLASIFIKASI ---
with tab2:
    st.subheader("Klasifikasi ke Folder ZIP")
    if not color_map: st.info("Upload Excel di sidebar dahulu.")
    else:
        c_files = st.file_uploader("Pilih PDF", type="pdf", accept_multiple_files=True, key="cls")
        if st.button("Proses Klasifikasi") and c_files:
            cls_buffer = io.BytesIO()
            with zipfile.ZipFile(cls_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_c:
                for f in c_files:
                    ref = extract_referensi(f)
                    folder = color_map.get(ref, "Tidak_Ada_Di_Excel")
                    zip_c.writestr(f"{folder}/{f.name}", f.getvalue())
            st.download_button("⬇️ Download ZIP Terklasifikasi", cls_buffer.getvalue(), "klasifikasi.zip", "application/zip", use_container_width=True)

# --- TAB 3: MERGE ---
with tab3:
    st.subheader("Gabung PDF")
    m_files = st.file_uploader("Pilih minimal 2 PDF", type="pdf", accept_multiple_files=True, key="mrg")
    if st.button("Gabungkan") and len(m_files) >= 2:
        # Urutkan berdasarkan nama file secara ascending
        sorted_files = sorted(m_files, key=lambda x: x.name)
        merger = PdfMerger()
        for f in sorted_files: merger.append(f)
        out = io.BytesIO()
        merger.write(out)
        st.download_button("⬇️ Download PDF Gabungan", out.getvalue(), "hasil_merge.pdf", "application/pdf", use_container_width=True)
