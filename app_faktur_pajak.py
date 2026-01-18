import streamlit as st
import pdfplumber
from PyPDF2 import PdfMerger
import re
import zipfile
import io
from openpyxl import load_workbook

# --- FUNGSI LOGIKA (Diambil dari kode asli Anda) ---

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

# --- TAMPILAN STREAMLIT ---

st.set_page_config(page_title="Faktur Pajak Tool", layout="centered")

st.title("📑 Faktur Pajak Tool")
st.write("Versi Web (Mobile Friendly)")

# Sidebar untuk Database Excel
with st.sidebar:
    st.header("Settings")
    excel_db = st.file_uploader("Upload Excel Warna (Opsional)", type=["xlsx"])
    color_map = None
    if excel_db:
        color_map = get_color_mapping(excel_db)
        if color_map:
            st.success("Database Dimuat!")

# Tab Menu
tab1, tab2, tab3 = st.tabs(["Rename", "Klasifikasi", "Merge"])

with tab1:
    st.subheader("Rename File PDF")
    files = st.file_uploader("Pilih PDF untuk di-rename", type="pdf", accept_multiple_files=True, key="rename_upload")
    if st.button("Proses Rename") and files:
        for f in files:
            ref = extract_referensi(f)
            if ref:
                st.success(f"File: {f.name} -> **{ref}_{f.name}**")
            else:
                st.warning(f"Referensi tidak ditemukan pada: {f.name}")

with tab2:
    st.subheader("Klasifikasi & Download ZIP")
    if not color_map:
        st.info("Silakan upload file Excel di sidebar terlebih dahulu.")
    else:
        files = st.file_uploader("Pilih PDF untuk diklasifikasi", type="pdf", accept_multiple_files=True, key="class_upload")
        
        if files:
            # Siapkan Buffer untuk ZIP
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for f in files:
                    ref = extract_referensi(f)
                    folder_name = color_map.get(ref, "Tidak_Ada_Di_Excel")
                    
                    # Simpan file ke dalam folder di dalam ZIP
                    # Format: Nama_Folder/Nama_File.pdf
                    zip_path = f"{folder_name}/{f.name}"
                    zip_file.writestr(zip_path, f.getvalue())
            
            st.success(f"{len(files)} file siap dikelompokkan!")
            
            # Tombol Download ZIP
            st.download_button(
                label="Download Hasil Klasifikasi (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="faktur_terklasifikasi.zip",
                mime="application/zip",
                use_container_width=True
            )

with tab3:
    st.subheader("Gabung PDF (Merge)")
    files = st.file_uploader("Pilih minimal 2 PDF", type="pdf", accept_multiple_files=True, key="merge_upload")
    if st.button("Gabungkan PDF") and len(files) >= 2:
        merger = PdfMerger()
        for f in files:
            merger.append(f)
        
        output = io.BytesIO()
        merger.write(output)
        st.download_button(
            label="Download PDF Gabungan",
            data=output.getvalue(),
            file_name="hasil_gabungan.pdf",
            mime="application/pdf"
        )
    
