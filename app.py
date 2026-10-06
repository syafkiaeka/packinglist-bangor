# app.py - BANGOR APS PACKING LIST GENERATOR (PRO UI)
import streamlit as st
import os
import tempfile
from engine import PackingListEngine
# ==========================================
# 1. KONFIGURASI HALAMAN & TEMA
# ==========================================
st.set_page_config(
    page_title="Bangor - Packing List Generator",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 2. CUSTOM CSS (TEMA PT BANGOR)
# ==========================================
st.markdown("""
<style>
    /* Sembunyikan Footer & Menu Default Streamlit agar terlihat seperti App Asli */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Warna Tema Korporat (Navy Blue & Clean White) */
    .stButton>button {
        background-color: #003366; /* Navy Blue */
        color: white;
        border-radius: 6px;
        padding: 10px 24px;
        font-weight: bold;
        border: none;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #004080;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    
    /* Styling Header */
    .bangor-header {
        background: linear-gradient(90deg, #003366 0%, #00509E 100%);
        color: white;
        padding: 25px;
        border-radius: 10px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .bangor-header h1 {
        margin: 0;
        font-size: 28px;
        font-weight: 800;
        letter-spacing: 1px;
    }
    .bangor-header p {
        margin: 5px 0 0 0;
        font-size: 14px;
        opacity: 0.9;
    }

    /* Styling Container/Box */
    .upload-box {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 20px;
    }

    /* Styling File Uploader */
    .stFileUploader > div > button {
        background-color: #003366;
        color: white;
    }

    /* Footer */
    .app-footer {
        margin-top: 50px;
        padding-top: 20px;
        border-top: 1px solid #e9ecef;
        text-align: center;
        color: #6c757d;
        font-size: 12px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. HEADER & SIDEBAR
# ==========================================
# Header Utama
st.markdown("""
<div class="bangor-header">
    <h1>📦 Bangor APP Packinglist Generator </h1>
    <p>Sistem Generator Packing List Otomatis (Delivery Order)</p>
</div>
""", unsafe_allow_html=True)

# Sidebar Informasi
with st.sidebar:
    st.header("ℹ️ Informasi Sistem")
    st.markdown("""
    **Versi:** 2.0 (Enterprise)  
    **Status:** 🟢 Online  
    **Format Input:** `.xlsx`, `.xlsm`  
    **Format Output:** `.xlsx` (Standard)
    """)
    st.divider()
    st.subheader("📋 Panduan Penggunaan")
    st.markdown("""
    1. Pastikan file Excel memiliki sheet **Master Data Packinglist**.
    2. Sheet DO harus berupa **Rincian Pemindahan Barang** atau **Delivery Order Detail**.
    3. Upload file, klik Generate, dan download hasilnya.
    """)

# ==========================================
# 4. MAIN CONTENT (LAYOUT 2 KOLOM)
# ==========================================
col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("📤 Upload File Delivery Order")
    st.markdown("Silakan upload file Excel Master Data & DO Anda di bawah ini.")
    
    # File Uploader
    uploaded = st.file_uploader(
        "Pilih file Excel (.xlsx / .xlsm)", 
        type=["xlsx", "xlsm"],
        help="File harus mengandung sheet 'Master Data Packinglist' dan sheet DO yang valid."
    )

with col2:
    st.subheader("📊 Status Proses")
    if uploaded is not None:
        st.success(f"✅ **{uploaded.name}** siap diproses.")
        st.info(f"Ukuran: {round(uploaded.size / 1024, 2)} KB")
    else:
        st.warning(" Menunggu file diupload...")

st.divider()

# ==========================================
# 5. PROSES & DOWNLOAD
# ==========================================
if uploaded is not None:
    # Tombol Generate
    generate_clicked = st.button(
        "🚀 GENERATE PACKING LIST", 
        type="primary", 
        use_container_width=True
    )

    if generate_clicked:
        with st.spinner("️ Sedang memproses data... Mohon tunggu."):
            log_box = st.empty()
            logs = []
            
            def cb(msg):
                logs.append(msg)
                log_box.code("\n".join(logs), language="text")

            # Simpan file input sementara
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
                tmp_in.write(uploaded.getvalue())
                tmp_in_path = tmp_in.name

            tmp_out_path = tmp_in_path.replace(".xlsx", "_PACKINGLIST.xlsx")
            base_name = os.path.splitext(uploaded.name)[0]
            download_filename = f"PACKINGLIST_{base_name}.xlsx"

            try:
                # Jalankan Engine
                engine = PackingListEngine(tmp_in_path, tmp_out_path, progress_callback=cb)
                engine.run()

                # Baca file hasil
                with open(tmp_out_path, "rb") as f:
                    data = f.read()

                st.balloons()
                st.success(f"✅ **BERHASIL!** Total **{len(engine.FinalResult)}** Delivery Order berhasil diproses.")
                
                # Tombol Download
                st.download_button(
                    label="⬇️ DOWNLOAD HASIL PACKING LIST (.xlsx)",
                    data=data,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"❌ **GAGAL:** {e}")
            finally:
                # Hapus file sementara
                for p in [tmp_in_path, tmp_out_path]:
                    if os.path.exists(p): 
                        os.remove(p)

# ==========================================
# 6. FOOTER
# ==========================================
st.markdown("""
<div class="app-footer">
    &copy; 2026 Eka Logistik. All rights reserved.
</div>
""", unsafe_allow_html=True)
