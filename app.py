# app.py - PT BANGOR PACKING LIST GENERATOR (FINAL FIXED)
import streamlit as st
import os
import tempfile
from collections import OrderedDict
from engine import PackingListEngine

st.set_page_config(page_title="PT Bangor - Packing List Generator", page_icon="📦", layout="wide")

st.markdown("""
<style>
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;} .stDeployButton {display: none;}
    .stButton>button { background-color: #003366; color: white; border-radius: 6px; padding: 10px 24px; font-weight: bold; border: none; transition: all 0.3s ease; }
    .stButton>button:hover { background-color: #004080; transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0,0,0,0.2); }
    .bangor-header { background: linear-gradient(90deg, #003366 0%, #00509E 100%); color: white; padding: 25px; border-radius: 10px; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .bangor-header h1 { margin: 0; font-size: 28px; font-weight: 800; letter-spacing: 1px; }
    .bangor-header p { margin: 5px 0 0 0; font-size: 14px; opacity: 0.9; }
    .app-footer { margin-top: 50px; padding-top: 20px; border-top: 1px solid #e9ecef; text-align: center; color: #6c757d; font-size: 12px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="bangor-header">
    <h1>📦 Packinglist Generator - Logistic </h1>
    <p>Sistem Generator Packing List Otomatis (Multi-File)</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("️ Informasi Sistem")
    st.markdown("**Versi:** 3.4 (No Double Export)<br>**Status:** 🟢 Online", unsafe_allow_html=True)
    st.divider()
    st.subheader("📋 Panduan")
    st.markdown("1. Upload 1 atau **banyak file Excel** sekaligus.<br>2. Klik Generate.<br>3. Download 1 file gabungan.", unsafe_allow_html=True)

st.subheader("📤 Upload File Delivery Order")
st.markdown("Kamu bisa upload **banyak file Excel sekaligus** dari gudang yang berbeda!")

uploaded = st.file_uploader(
    "Pilih file Excel (.xlsx / .xlsm)", 
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    help="Bisa pilih banyak file sekaligus dengan menahan Ctrl/Cmd saat memilih."
)

# ANTI-DUPLIKAT FILE
if uploaded:
    unique_uploaded = []
    seen_files = set()
    for f in uploaded:
        file_key = (f.name, f.size)
        if file_key not in seen_files:
            unique_uploaded.append(f)
            seen_files.add(file_key)
    uploaded = unique_uploaded

    st.success(f"✅ **{len(uploaded)} file unik** siap diproses.")
    for f in uploaded:
        st.text(f"📄 {f.name} ({round(f.size / 1024, 2)} KB)")
else:
    st.warning("⏳ Menunggu file diupload...")

st.divider()

if uploaded:
    generate_clicked = st.button("🚀 GENERATE PACKING LIST (GABUNG SEMUA)", type="primary", use_container_width=True)

    if generate_clicked:
        master_wb = None
        total_do = 0
        processed_do_set = set()
        
        with st.spinner("️ Sedang memproses semua file... Mohon tunggu."):
            log_box = st.empty()
            logs = []
            def cb(msg):
                logs.append(msg)
                log_box.code("\n".join(logs), language="text")

            try:
                for idx, up_file in enumerate(uploaded):
                    st.write(f"🔄 Memproses file {idx+1}/{len(uploaded)}: **{up_file.name}**...")
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
                        tmp_in.write(up_file.getvalue())
                        tmp_in_path = tmp_in.name
                    tmp_out_path = tmp_in_path.replace(".xlsx", "_temp.xlsx")

                    # 1. Jalankan Engine (Hanya baca & proses data, TIDAK export)
                    engine = PackingListEngine(tmp_in_path, tmp_out_path, progress_callback=cb, master_wb=master_wb)
                    engine.run()
                    
                    # 2. Filter DO yang sudah pernah diproses (Anti-Duplikat)
                    unique_final_result = OrderedDict()
                    for do_no, data in engine.FinalResult.items():
                        if do_no not in processed_do_set:
                            unique_final_result[do_no] = data
                            processed_do_set.add(do_no)
                    
                    engine.FinalResult = unique_final_result
                    
                    # 3. Export ke Excel HANYA SEKALI untuk data yang unik
                    master_wb = engine.export()
                    total_do += len(unique_final_result)
                    
                    if os.path.exists(tmp_in_path): os.remove(tmp_in_path)
                    if os.path.exists(tmp_out_path): os.remove(tmp_out_path)

                final_out_path = tempfile.mktemp(suffix=".xlsx")
                master_wb.save(final_out_path)

                with open(final_out_path, "rb") as f:
                    data = f.read()

                                # ==========================================
                # ANIMASI HUJAN BURGER (IFRAME METHOD)
                # ==========================================
                burger_html = """
                <!DOCTYPE html>
                <html>
                <head>
                <style>
                    body { margin: 0; padding: 0; overflow: hidden; background: transparent; }
                    .falling-item {
                        position: absolute;
                        top: -50px;
                        font-size: 2.5rem;
                        animation: fall linear forwards;
                    }
                    @keyframes fall {
                        0% { transform: translateY(0) rotate(0deg); opacity: 1; }
                        100% { transform: translateY(110vh) rotate(720deg); opacity: 0; }
                    }
                </style>
                </head>
                <body>
                <script>
                    const emojis = ['🍔', '🍟', '🥤', '🍗'];
                    for(let i = 0; i < 50; i++) {
                        let el = document.createElement('div');
                        el.className = 'falling-item';
                        el.innerText = emojis[Math.floor(Math.random() * emojis.length)];
                        el.style.left = Math.random() * 100 + 'vw';
                        el.style.animationDuration = (Math.random() * 3 + 2) + 's';
                        el.style.animationDelay = Math.random() * 2 + 's';
                        document.body.appendChild(el);
                    }
                    // Hapus animasi setelah 6 detik agar tidak berat
                    setTimeout(() => { document.body.innerHTML = ''; }, 6000);
                </script>
                </body>
                </html>
                """

                st.markdown(f"""
                    <iframe srcdoc="{burger_html}" 
                            style="position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; pointer-events: none; z-index: 99999; border: none;" 
                            sandbox="allow-scripts">
                    </iframe>
                """, unsafe_allow_html=True)

                st.success(f"✅ **BERHASIL!** Total **{total_do} Delivery Order Unik** dari {len(uploaded)} file berhasil digabung.")
                
                st.download_button(
                    label="️ DOWNLOAD HASIL GABUNGAN (.xlsx)",
                    data=data,
                    file_name=f"PACKINGLIST_GABUNGAN_{len(uploaded)}FILE.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"❌ **GAGAL:** {e}")
            finally:
                if os.path.exists(final_out_path): os.remove(final_out_path)

st.markdown("""
<div class="app-footer">
    &copy; 2026 Bangor - Internal Logistics System. All rights reserved.
</div>
""", unsafe_allow_html=True)
