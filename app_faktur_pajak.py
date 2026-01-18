import streamlit as st
import pdfplumber
from PyPDF2 import PdfMerger
import re
import io
import zipfile
import pandas as pd
import time
from openpyxl import load_workbook

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Faktur Pajak Tool Pro", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# --- FUNGSI LOGIKA ---

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
                color_map[ref_val] = str(color_hex) # Simpan hex sementara
        return color_map
    except: return None

# --- SIDEBAR PANDUAN ---
with st.sidebar:
    st.header("📘 Panduan Penggunaan")
    st.markdown("""
    1. **Upload Excel**: Masukkan database warna jika ingin menggunakan fitur **Klasifikasi**.
    2. **Pilih Mode**:
        * **Rename**: Mengubah nama file otomatis.
        * **Klasifikasi**: Mengelompokkan file ke folder ZIP.
        * **Merge**: Menggabungkan banyak PDF.
    3. **Live Preview**: Cek tabel hasil ekstraksi sebelum klik download.
    4. **Download**: Klik tombol download ZIP untuk mengambil hasil.
    """)
    st.divider()
    st.info("Aplikasi ini memproses file di memori sementara dan tidak menyimpannya secara permanen.")
    
    st.header("⚙️ Database")
    excel_db = st.file_uploader("Upload Excel Warna", type=["xlsx"])
    color_map = get_color_mapping(excel_db) if excel_db else None
    if color_map: st.success("✅ Database Dimuat!")

# --- TAMPILAN UTAMA ---
st.title("📑 Faktur Pajak Tool Pro")

tab1, tab2, tab3 = st.tabs(["🔄 Rename & Preview", "📁 Klasifikasi ZIP", "🔗 Merge PDF"])

# --- TAB 1: RENAME & LIVE PREVIEW ---
with tab1:
    st.subheader("Rename dengan Live Preview")
    files = st.file_uploader("Upload PDF", type="pdf", accept_multiple_files=True, key="ren")
    
    if files:
        results = []
        rename_buffer = io.BytesIO()
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        with zipfile.ZipFile(rename_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
            for i, f in enumerate(files):
                # Update Progress
                percent = (i + 1) / len(files)
                progress_bar.progress(percent)
                status_text.text(f"Memproses {i+1}/{len(files)}: {f.name}")
                
                ref = extract_referensi(f)
                status = "✅ Berhasil" if ref else "❌ Gagal"
                
                # Logika Penamaan
                name_only = f.name.replace(".pdf", "").replace(".PDF", "")
                parts = name_only.split('-')
                suffix = "-".join(parts[-3:]) if len(parts) >= 3 else name_only
                new_name = f"{ref} {suffix}.pdf" if ref else f.name
                
                if ref:
                    zip_f.writestr(new_name, f.getvalue())
                
                results.append({
                    "Status": status,
                    "Nama Asli": f.name,
                    "Referensi": ref if ref else "Tidak Terbaca",
                    "Nama Baru": new_name
                })
        
        # Tampilkan Tabel Preview
        df = pd.DataFrame(results)
        st.table(df) # Menampilkan tabel hasil ekstraksi secara live
        
        if any(r['Status'] == "✅ Berhasil" for r in results):
            st.download_button(
                "⬇️ Download Hasil Rename (ZIP)", 
                rename_buffer.getvalue(), 
                "rename_faktur.zip", 
                "application/zip", 
                use_container_width=True
            )

# --- TAB 2: KLASIFIKASI ---
with tab2:
    st.subheader("Klasifikasi Otomatis ke Folder")
    if not color_map:
        st.warning("Silakan unggah database Excel di Sidebar terlebih dahulu.")
    else:
        c_files = st.file_uploader("Upload PDF untuk dikelompokkan", type="pdf", accept_multiple_files=True, key="cls")
        if c_files:
            cls_buffer = io.BytesIO()
            cls_results = []
            
            p_bar_cls = st.progress(0)
            with zipfile.ZipFile(cls_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_c:
                for i, f in enumerate(c_files):
                    p_bar_cls.progress((i + 1) / len(c_files))
                    ref = extract_referensi(f)
                    folder = color_map.get(ref, "Tidak_Ada_Di_Excel")
                    status = "✅ Berhasil" if ref else "❌ Gagal (Ref Tidak Ada)"
                    
                    zip_c.writestr(f"{folder}/{f.name}", f.getvalue())
                    cls_results.append({"File": f.name, "Referensi": ref, "Target Folder": folder, "Status": status})
            
            st.dataframe(pd.DataFrame(cls_results), use_container_width=True)
            st.download_button("⬇️ Download ZIP Terklasifikasi", cls_buffer.getvalue(), "klasifikasi.zip", use_container_width=True)

# --- TAB 3: MERGE ---
with tab3:
    st.subheader("Gabungkan PDF (A-Z)")
    m_files = st.file_uploader("Upload PDF untuk digabung", type="pdf", accept_multiple_files=True, key="mrg")
    if m_files:
        sorted_files = sorted(m_files, key=lambda x: x.name)
        st.write("Urutan penggabungan:")
        st.text(" ➔ ".join([f.name for f in sorted_files]))
        
        if st.button("Proses Gabung"):
            merger = PdfMerger()
            m_progress = st.progress(0)
            for i, f in enumerate(sorted_files):
                m_progress.progress((i + 1) / len(sorted_files))
                merger.append(f)
            
            out = io.BytesIO()
            merger.write(out)
            st.success("✅ Penggabungan Selesai!")
            st.download_button("⬇️ Download PDF Gabungan", out.getvalue(), "hasil_merge.pdf", use_container_width=True)
