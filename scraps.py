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
        response.raise_for_status()
        try:
            data_json = response.json()
            content = data_json.get('data', data_json)
            return pd.DataFrame(content)
        except ValueError:
            return pd.read_excel(io.BytesIO(response.content))
    except Exception as e:
        return None

def convert_df_to_excel(df):
    """Mengubah DataFrame menjadi binary Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

def generate_safe_filename(row, name_columns, max_length=120):
    """
    Membuat nama file dari gabungan kolom yang dipilih user.
    """
    # Gabungkan isi kolom-kolom yang dipilih dengan tanda strip (-)
    parts = [str(row[col]) for col in name_columns if pd.notna(row[col])]
    raw_name = "-".join(parts)
    
    # Bersihkan nama file
    clean_name = raw_name.replace(" ", "_")
    clean_name = "".join([c for c in clean_name if c.isalnum() or c in ('.', '_', '-')]).strip()
    
    # Jika hasil gabungan kosong, beri nama default
    if not clean_name:
        clean_name = "downloaded_file"

    if len(clean_name) > max_length:
        clean_name = clean_name[:max_length]
        
    return f"{clean_name}.xlsx"

def create_zip_archive(selected_indices, df_source, url_col_name, name_cols_list):
    """
    Membuat file ZIP dengan parameter kolom dinamis.
    """
    zip_buffer = io.BytesIO()
    url_cache = {} 
    used_filenames = set()
    
    report = {
        'success': [],
        'failed': []
    }

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        progress_text = "Sedang memproses file..."
        my_bar = st.progress(0, text=progress_text)
        total = len(selected_indices)
        
        for i, index in enumerate(selected_indices):
            row = df_source.loc[index]
            url_target = row[url_col_name] # Pakai kolom URL pilihan user
            
            # 1. Generate Nama File (Pakai kolom Nama pilihan user)
            base_filename = generate_safe_filename(row, name_cols_list)
            
            # Handle Duplikasi
            final_filename = base_filename
            counter = 1
            while final_filename in used_filenames:
                name_without_ext = base_filename.replace(".xlsx", "")
                final_filename = f"{name_without_ext}_({counter}).xlsx"
                counter += 1
            
            used_filenames.add(final_filename)
            
            my_bar.progress((i + 1) / total, text=f"Memproses: {final_filename}")

            # 2. Fetching
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

            # 3. Writing
            if df_data is not None and not df_data.empty:
                excel_bytes = convert_df_to_excel(df_data)
                zip_file.writestr(final_filename, excel_bytes)
                report['success'].append(final_filename)
            else:
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
    uploaded_file = st.file_uploader("Upload File Excel List", type=['xlsx'])

    st.divider()
    if st.button("🔴 RESET / MATIKAN SISTEM", type="primary"):
        st.session_state.clear()
        st.rerun()

# ==========================================
# MAIN DASHBOARD
# ==========================================
st.title("📊 Pengunduh Satu Data Untuk BPS DKI")
st.markdown(f"Sistem downloader fleksibel dengan mapping kolom dinamis.")
st.divider()

if uploaded_file is not None:
    try:
        # Baca Excel mentah
        df_raw = pd.read_excel(uploaded_file)
        all_columns = df_raw.columns.tolist()

        # ==========================================
        # 1. KONFIGURASI MAPPING KOLOM (INPUT USER)
        # ==========================================
        st.subheader("⚙️ Konfigurasi Kolom")
        col_conf1, col_conf2 = st.columns(2)
        
        with col_conf1:
            st.info("Pilih kolom yang berisi **Link Download (URL)**:")
            # Coba cari otomatis kolom yg namanya mirip 'link' atau 'url'
            default_url_ix = 0
            for ix, col in enumerate(all_columns):
                if 'link' in col.lower() or 'url' in col.lower():
                    default_url_ix = ix
                    break
            
            selected_url_col = st.selectbox("Kolom URL Target", all_columns, index=default_url_ix)

        with col_conf2:
            st.info("Pilih kolom untuk menyusun **Nama File** (Urutan berpengaruh):")
            # Default select semua kolom kecuali link download
            default_name_cols = [c for c in all_columns if c != selected_url_col]
            selected_name_cols = st.multiselect("Kolom Penamaan File", all_columns, default=default_name_cols[:3]) # Default ambil 3 kolom pertama

        # Tombol Konfirmasi Konfigurasi
        if st.button("🚀 Terapkan Konfigurasi & Validasi Data"):
            st.session_state['config_confirmed'] = True
            st.session_state['url_col'] = selected_url_col
            st.session_state['name_cols'] = selected_name_cols
        
        # ==========================================
        # 2. LOGIC PROSES (SETELAH KONFIRMASI)
        # ==========================================
        if st.session_state.get('config_confirmed', False):
            st.divider()
            
            # Ambil config dari session state
            url_col = st.session_state['url_col']
            name_cols = st.session_state['name_cols']
            
            # Validasi NaN pada kolom URL
            df_clean = df_raw.dropna(subset=[url_col])
            df_dirty = df_raw[df_raw[url_col].isnull()]

            if not df_dirty.empty:
                st.warning(f"⚠️ Ada **{len(df_dirty)} baris** yang kolom URL-nya kosong (NaN) dan akan dilewati.")
            
            if df_clean.empty:
                st.error("⛔ Tidak ada data valid (semua URL kosong).")
                st.stop()

            # Siapkan Data Bersih
            df_input = df_clean.reset_index(drop=True)
            st.success(f"✅ Konfigurasi tersimpan! Menggunakan kolom URL: **'{url_col}'**. Siap memproses **{len(df_input)}** file.")

            # ==========================================
            # 3. AREA BULK ACTION
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
                        with st.spinner("Sedang memproses..."):
                            # PANGGIL FUNGSI ZIP DENGAN PARAMETER DINAMIS
                            zip_data, zip_report = create_zip_archive(
                                selected_indices, 
                                df_input, 
                                url_col,    # Pass nama kolom URL
                                name_cols   # Pass list kolom Nama
                            )
                            
                            st.session_state['zip_ready'] = zip_data
                            st.session_state['zip_report'] = zip_report
                            st.rerun()

            # --- REPORT AREA ---
            if 'zip_ready' in st.session_state:
                report = st.session_state.get('zip_report', {'success': [], 'failed': []})
                count_success = len(report['success'])
                count_failed = len(report['failed'])
                
                with st.container(border=True):
                    st.markdown("#### 📑 Laporan Pembuatan ZIP")
                    m1, m2 = st.columns(2)
                    m1.metric("Berhasil", f"{count_success} File")
                    m2.metric("Gagal", f"{count_failed} File", delta_color="inverse")
                    
                    if count_failed > 0:
                        st.error(f"⚠️ {count_failed} file gagal.")
                        with st.expander("Detail Gagal"):
                            st.table(pd.DataFrame(report['failed']))
                    
                    st.download_button(
                        label="⬇️ KLIK UNTUK UNDUH ZIP",
                        data=st.session_state['zip_ready'],
                        file_name="BPS_Data_Archive.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )

            st.markdown("---")

            # ==========================================
            # 4. LIST FILE (LOOPING)
            # ==========================================
            h_col1, h_col2 = st.columns([0.5, 9.5])
            h_col1.markdown("**#**")
            h_col2.markdown("**Daftar File & URL**")

            for index, row in df_input.iterrows():
                # Generate Nama File pakai kolom dinamis
                file_name_full = generate_safe_filename(row, name_cols)
                # Ambil URL pakai kolom dinamis
                url_target = row[url_col]
                
                col_check, col_exp = st.columns([0.5, 9.5])
                
                with col_check:
                    st.checkbox("", key=f"check_{index}")

                with col_exp:
                    with st.expander(f"📄 {file_name_full}"):
                        c1, c2 = st.columns([1, 3])
                        
                        with c1:
                            st.markdown("**Info Kolom:**")
                            # Tampilkan preview nilai dari kolom penamaan yg dipilih user
                            for col_name in name_cols:
                                st.caption(f"**{col_name}**: {row[col_name]}")
                            
                            st.caption(f"**URL**: {url_target}")

                            if st.button("🔍 Cek Data", key=f"btn_fetch_{index}"):
                                with st.spinner('Loading...'):
                                    res = fetch_data(url_target)
                                    if res is not None:
                                        st.session_state[f"data_{index}"] = res
                                        st.success("OK")
                                    else:
                                        st.error("Gagal Fetch")

                        with c2:
                            if f"data_{index}" in st.session_state:
                                df_s = st.session_state[f"data_{index}"]
                                st.dataframe(df_s, use_container_width=True)
                                st.download_button(
                                    "⬇️ Unduh File Ini",
                                    data=convert_df_to_excel(df_s),
                                    file_name=file_name_full,
                                    key=f"dl_{index}"
                                )

    except Exception as e:
        st.error(f"Error membaca file: {e}")

else:
    st.markdown("<div style='text-align:center;color:grey;padding:50px;'>Silakan Upload File Excel yang berisi Link Download...</div>", unsafe_allow_html=True)