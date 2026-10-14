# app.py - PT BANGOR PACKING LIST GENERATOR
import streamlit as st
import os
import tempfile
from collections import OrderedDict
from engine import PackingListEngine

st.set_page_config(
    page_title="Bangor - Packing List Generator", 
    page_icon="📦", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    #MainMenu {visibility: hidden;} 
    footer {visibility: hidden;} 
    header {visibility: hidden;} 
    .stDeployButton {display: none;}
    
    body { background-color: #f0f2f6; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    
    .bangor-header { 
        background: linear-gradient(135deg, #003366 0%, #00509E 100%); 
        color: white; padding: 30px; border-radius: 12px; margin-bottom: 25px; 
        box-shadow: 0 8px 16px rgba(0,51,102,0.15); 
    }
    .bangor-header h1 { margin: 0; font-size: 32px; font-weight: 700; letter-spacing: 0.5px; }
    .bangor-header p { margin: 8px 0 0 0; font-size: 15px; opacity: 0.9; font-weight: 300; }

    .dashboard-card {
        background-color: white; border-radius: 10px; padding: 25px; 
        box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #e9ecef;
        margin-bottom: 20px; height: 100%;
    }
    
    .stButton>button {
        background-color: #003366; color: white; border-radius: 8px; padding: 12px 24px;
        font-weight: 600; border: none; transition: all 0.2s ease; width: 100%;
    }
    .stButton>button:hover { background-color: #004080; transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0,51,102,0.2); }
    .stButton>button:disabled { background-color: #cbd5e1; cursor: not-allowed; }

    .app-footer { margin-top: 40px; padding: 20px; text-align: center; color: #94a3b8; font-size: 13px; border-top: 1px solid #e2e8f0; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="bangor-header">
    <h1>📦 Generator Packinglist By-Logistik</h1>
    <p>Sistem Generator Packing List Otomatis • Multi-File Processing • Anti-Duplikat</p>
</div>
""", unsafe_allow_html=True)

col_input, col_status = st.columns([1.8, 1])

with col_input:
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.subheader("📤 Upload Delivery Order")
    st.caption("Upload satu atau banyak file Excel (.xlsx / .xlsm) sekaligus. Sistem akan otomatis menggabungkan dan menyaring duplikat.")
    
    uploaded = st.file_uploader(
        "", 
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )
    
    if uploaded:
        unique_uploaded = []
        seen_files = set()
        for f in uploaded:
            file_key = (f.name, f.size)
            if file_key not in seen_files:
                unique_uploaded.append(f)
                seen_files.add(file_key)
        uploaded = unique_uploaded
        
        st.success(f"✅ **{len(uploaded)} file unik** siap diproses")
        with st.expander("Lihat detail file yang dipilih"):
            for f in uploaded:
                st.text(f"📄 {f.name} ({round(f.size / 1024, 2)} KB)")
    else:
        st.info("⏳ Belum ada file yang dipilih")
        
    st.markdown('</div>', unsafe_allow_html=True)

with col_status:
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.subheader("📊 Ringkasan Proses")
    
    if uploaded:
        st.metric(label="Total File Diupload", value=len(uploaded))
        st.divider()
        st.caption("Klik tombol di bawah untuk mulai memproses semua file secara bersamaan.")
        
        generate_clicked = st.button(
            "🚀 MULAI GENERATE PACKING LIST", 
            type="primary", 
            use_container_width=True
        )
    else:
        st.metric(label="Total File Diupload", value="-")
        st.divider()
        st.caption("Silakan upload file Excel terlebih dahulu untuk mengaktifkan tombol Generate.")
        st.button("🚀 GENERATE PACKING LIST", disabled=True, use_container_width=True)
        
    st.markdown('</div>', unsafe_allow_html=True)

if uploaded and generate_clicked:
    master_wb = None
    total_do = 0
    processed_do_set = set()
    
    st.divider()
    st.subheader("⚙️ Sedang Memproses...")
    progress_bar = st.progress(0.0)
    status_text = st.empty()
    log_box = st.code("", language="text", height=150)
    
    logs = []
    def cb(msg):
        logs.append(msg)
        log_box.text("\n".join(logs[-5:]))

    try:
        total_files = len(uploaded)
        
        for idx, up_file in enumerate(uploaded):
            current_progress = (idx + 1) / total_files
            progress_bar.progress(current_progress)
            status_text.text(f"Memproses file {idx+1}/{total_files}: **{up_file.name}**")
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
                tmp_in.write(up_file.getvalue())
                tmp_in_path = tmp_in.name
            tmp_out_path = tmp_in_path.replace(".xlsx", "_temp.xlsx")

            engine = PackingListEngine(tmp_in_path, tmp_out_path, progress_callback=cb, master_wb=master_wb)
            engine.run()
            
            unique_final_result = OrderedDict()
            for do_no, data in engine.FinalResult.items():
                if do_no not in processed_do_set:
                    unique_final_result[do_no] = data
                    processed_do_set.add(do_no)
            
            engine.FinalResult = unique_final_result
            
            master_wb = engine.export()
            total_do += len(unique_final_result)
            
            if os.path.exists(tmp_in_path): os.remove(tmp_in_path)
            if os.path.exists(tmp_out_path): os.remove(tmp_out_path)

        final_out_path = tempfile.mktemp(suffix=".xlsx")
        master_wb.save(final_out_path)

        with open(final_out_path, "rb") as f:
            data = f.read()

        progress_bar.progress(1.0)
        status_text.text("✅ Semua file berhasil diproses!")
        
        # Animasi Burger Hujan
        burger_html = """
        <!DOCTYPE html><html><head><style>body{margin:0;padding:0;overflow:hidden;background:transparent}.falling-item{position:absolute;top:-50px;font-size:2.5rem;animation:fall linear forwards}@keyframes fall{0%{transform:translateY(0) rotate(0deg);opacity:1}100%{transform:translateY(110vh) rotate(720deg);opacity:0}}</style></head><body><script>const emojis=['🍔','🍟','','🍗'];for(let i=0;i<50;i++){let el=document.createElement('div');el.className='falling-item';el.innerText=emojis[Math.floor(Math.random()*emojis.length)];el.style.left=Math.random()*100+'vw';el.style.animationDuration=(Math.random()*3+2)+'s';el.style.animationDelay=Math.random()*2+'s';document.body.appendChild(el);}setTimeout(()=>{document.body.innerHTML='';},6000);</script></body></html>"""
        st.markdown(f'<iframe srcdoc="{burger_html}" style="position:fixed;top:0;left:0;width:100vw;height:100vh;pointer-events:none;z-index:99999;border:none;" sandbox="allow-scripts"></iframe>', unsafe_allow_html=True)

        # TOMBOL DOWNLOAD LANGSUNG DI SINI (TANPA CONTAINER)
        st.success(f"**BERHASIL!** Total **{total_do} Delivery Order Unik** dari {len(uploaded)} file telah digabung.")
        
        st.download_button(
            label="⬇️ DOWNLOAD PACKINGLIST (.xlsx)",
            data=data,
            file_name=f"PACKINGLIST_GABUNGAN_{len(uploaded)}FILE.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
            
    except Exception as e:
        progress_bar.progress(0.0)
        status_text.text("❌ Terjadi kesalahan!")
        st.error(f"**GAGAL MEMPROSES:** {e}")
    finally:
        if 'final_out_path' in locals() and os.path.exists(final_out_path): 
            os.remove(final_out_path)

st.markdown("""
<div class="app-footer">
    &copy; 2026 Bangor - Internal Logistics System v3.5<br>
    Developed for Warehouse Operations Team
</div>
""", unsafe_allow_html=True)
