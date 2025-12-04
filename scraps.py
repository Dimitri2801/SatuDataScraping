import streamlit as st
import pandas as pd
import requests
import io
import xlsxwriter 
import zipfile 

# ==========================================
# KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(
    page_title="BPS DKI Satu Data Mining",
    page_icon="📊",
    layout="wide"
)

# ==========================================
# FUNGSI UTILITAS & LOGIC
# ==========================================
@st.cache_data(show_spinner=False)
def fetch_data(url):
    """Mengambil data dari API dan mengembalikan DataFrame"""
    try:
        response = requests.get(url, timeout=30)
        # Jika status code 404/500, ini akan raise error dan masuk ke except
        response.raise_for_status()
        try:
            data_json = response.json()
            content = data_json.get('data', data_json)
            return pd.DataFrame(content)
        except ValueError:
            return pd.read_excel(io.BytesIO(response.content))
    except Exception as e:
        # Return None jika gagal, agar bisa dideteksi oleh logic pemanggil
        return None

def convert_df_to_excel(df):
    """Mengubah DataFrame menjadi binary Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

def generate_safe_filename(row, max_length=120):
    """Membuat nama file yang aman dan tidak kepanjangan"""
    raw_name = f"{str(row['No.'])}-{str(row['Dinas/lnstansi Pemerintah Daerah'])}-{str(row['Judul Tabel'])}"
    clean_name = raw_name.replace(" ", "_")
    clean_name = "".join([c for c in clean_name if c.isalnum() or c in ('.', '_', '-')]).strip()
    
    if len(clean_name) > max_length:
        clean_name = clean_name[:max_length]
        
    return f"{clean_name}.xlsx"

def create_zip_archive(selected_indices, df_source):
    """
    Membuat file ZIP dan Laporan Status (Sukses/Gagal).
    """
    zip_buffer = io.BytesIO()
    url_cache = {} 
    used_filenames = set()
    
    # Variabel untuk menampung laporan
    report = {
        'success': [], # List nama file yang berhasil
        'failed': []   # List nama file yang gagal
    }

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        progress_text = "Sedang memproses file..."
        my_bar = st.progress(0, text=progress_text)
        total = len(selected_indices)
        
        for i, index in enumerate(selected_indices):
            row = df_source.loc[index]
            url_target = row['link_download']
            
            # 1. Generate Nama File
            base_filename = generate_safe_filename(row)
            
            # Handle Duplikasi Nama
            final_filename = base_filename
            counter = 1
            while final_filename in used_filenames:
                name_without_ext = base_filename.replace(".xlsx", "")
                final_filename = f"{name_without_ext}_({counter}).xlsx"
                counter += 1
            
            used_filenames.add(final_filename)
            
            # Update Progress
            my_bar.progress((i + 1) / total, text=f"Memproses: {final_filename}")

            # 2. Fetching Data
            df_data = None
            if f"data_{index}" in st.session_state:
                df_data = st.session_state[f"data_{index}"]
            elif url_target in url_cache:
                df_data = url_cache[url_target]
            else:
                df_data = fetch_data(url_target)
                if df_data is not None:
                    url_cache[url_target] = df_data
                    st.session_state[f"data_{index}"] = df_data

            # 3. Validasi & Writing
            if df_data is not None and not df_data.empty:
                # SUKSES
                excel_bytes = convert_df_to_excel(df_data)
                zip_file.writestr(final_filename, excel_bytes)
                report['success'].append(final_filename)
            else:
                # GAGAL (Link mati / Timeout / Data Kosong)
                # Catat nama file dan linknya untuk laporan
                report['failed'].append({
                    'file': final_filename,
                    'url': url_target
                })
            
        my_bar.empty()
        
    return zip_buffer.getvalue(), report

def toggle_all_checkboxes(df_len, target_state):
    for i in range(df_len):
        st.session_state[f"check_{i}"] = target_state

# ==========================================
# SIDEBAR
# ==========================================
with st.sidebar:
    st.header("🎛️ Panel Kontrol")
    uploaded_file = st.file_uploader("Upload 'list.xlsx'", type=['xlsx'])

    st.divider()
    if st.button("🔴 RESET / MATIKAN SISTEM", type="primary"):
        st.session_state.clear()
        st.rerun()

# ==========================================
# MAIN DASHBOARD
# ==========================================
st.title("📊 Pengunduh Satu Data Untuk BPS DKI")
st.markdown(f"Selamat datang pada sistem pengunduhan satu data.")
st.divider()

if uploaded_file is not None:
    try:
        df_input = pd.read_excel(uploaded_file)
        
        required_cols = ['link_download', 'No.', 'Dinas/lnstansi Pemerintah Daerah', 'Judul Tabel']
        if not all(col in df_input.columns for col in required_cols):
            st.error(f"❌ Format salah! Kolom wajib: {required_cols}")
            st.stop() 
        
        else:
            cols_to_check = ['link_download', 'No.', 'Dinas/lnstansi Pemerintah Daerah', 'Judul Tabel']
            df_clean = df_input.dropna(subset=cols_to_check)
            df_dirty = df_input[df_input[cols_to_check].isnull().any(axis=1)]

            if not df_dirty.empty:
                st.warning(f"⚠️ Perhatian: Ditemukan **{len(df_dirty)} baris data tidak lengkap** (NaN).")
                with st.expander("Lihat Data yang Bermasalah"):
                    st.dataframe(df_dirty)
            
            if df_clean.empty:
                st.error("⛔ Semua data kosong.")
                st.stop()

            df_input = df_clean.reset_index(drop=True)
            st.info(f"✅ Siap Memproses **{len(df_input)}** file yang valid.")

            for col in cols_to_check:
                df_input[col] = df_input[col].astype(str)

            # ==========================================
            # AREA BULK ACTION
            # ==========================================
            st.markdown("### 📦 Bulk Action")
            
            action_col1, action_col2 = st.columns([2, 3])
            
            with action_col1:
                st.write("**Seleksi Cepat:**")
                sub_c1, sub_c2 = st.columns(2)
                with sub_c1:
                    st.button("✅ Select All", on_click=toggle_all_checkboxes, args=(len(df_input), True))
                with sub_c2:
                    st.button("❌ Unselect All", on_click=toggle_all_checkboxes, args=(len(df_input), False))

            with action_col2:
                selected_indices = []
                for i in range(len(df_input)):
                    if st.session_state.get(f"check_{i}", False):
                        selected_indices.append(i)
                
                st.write(f"**Terpilih: {len(selected_indices)} file**")
                
                if selected_indices:
                    if st.button("📦 ZIP Selected Files", type="primary"):
                        with st.spinner("Sedang memproses & memverifikasi link..."):
                            # Logic baru: terima 2 return value (zip & report)
                            zip_data, zip_report = create_zip_archive(selected_indices, df_input)
                            
                            st.session_state['zip_ready'] = zip_data
                            st.session_state['zip_report'] = zip_report # Simpan laporan ke session
                            st.rerun()

            # --- TAMPILAN REPORT & DOWNLOAD ---
            if 'zip_ready' in st.session_state:
                report = st.session_state.get('zip_report', {'success': [], 'failed': []})
                count_success = len(report['success'])
                count_failed = len(report['failed'])
                
                # Container Report
                with st.container(border=True):
                    st.markdown("#### 📑 Laporan Pembuatan ZIP")
                    
                    # Kolom Metric
                    m1, m2 = st.columns(2)
                    m1.metric("Berhasil Masuk ZIP", f"{count_success} File")
                    m2.metric("Gagal / Link Mati", f"{count_failed} File", delta_color="inverse")
                    
                    # Jika ada yang gagal, beri peringatan detail
                    if count_failed > 0:
                        st.error(f"⚠️ Ada {count_failed} file yang gagal didownload karena link tidak valid atau file terlalu besar.")
                        with st.expander("Lihat Detail File Gagal"):
                            # Tampilkan tabel file yg gagal
                            st.table(pd.DataFrame(report['failed']))
                    else:
                        st.success("✅ Sempurna! Semua link valid dan berhasil dikompres.")

                    # Tombol Download Final
                    st.download_button(
                        label="⬇️ KLIK UNTUK UNDUH HASIL ZIP",
                        data=st.session_state['zip_ready'],
                        file_name="BPS_Data_Archive.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )

            st.markdown("---")

            # ==========================================
            # LOOP LIST FILE
            # ==========================================
            h_col1, h_col2 = st.columns([0.5, 9.5])
            h_col1.markdown("**#**")
            h_col2.markdown("**Daftar File**")

            for index, row in df_input.iterrows():
                file_name_full = generate_safe_filename(row)
                url_target = row['link_download']
                
                col_check, col_exp = st.columns([0.5, 9.5])
                
                with col_check:
                    st.checkbox("", key=f"check_{index}")

                with col_exp:
                    with st.expander(f"📄 {file_name_full}"):
                        c1, c2 = st.columns([1, 3])
                        
                        with c1:
                            st.text(f"Dinas: \n{row['Dinas/lnstansi Pemerintah Daerah']}")
                            st.text(f"Rilis: \n{row['Judul Tabel']}")
                            if st.button("🔍 Cek Data", key=f"btn_fetch_{index}"):
                                with st.spinner('Loading...'):
                                    res = fetch_data(url_target)
                                    if res is not None:
                                        st.session_state[f"data_{index}"] = res
                                        st.success("OK")
                                    else:
                                        st.error("Gagal, karena file terlalu besar atau link tidak valid")

                        with c2:
                            if f"data_{index}" in st.session_state:
                                df_s = st.session_state[f"data_{index}"]
                                st.dataframe(df_s, use_container_width=True)
                                st.download_button(
                                    "⬇️ Unduh ini saja",
                                    data=convert_df_to_excel(df_s),
                                    file_name=file_name_full,
                                    key=f"dl_{index}"
                                )

    except Exception as e:
        st.error(f"Error: {e}")

else:
    st.markdown("<div style='text-align:center;color:grey;padding:50px;'>Waiting for file upload...</div>", unsafe_allow_html=True)