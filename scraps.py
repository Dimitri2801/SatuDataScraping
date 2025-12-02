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
    page_title="BPS DKI Scraper",
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

def create_zip_archive(selected_indices, df_source):
    """
    Membuat file ZIP berisi file Excel terpilih dengan penanganan duplikasi cerdas.
    """
    zip_buffer = io.BytesIO()
    
    # Dictionary untuk menyimpan cache data URL yang sudah didownload dalam sesi ini
    # Format: { 'url_download': dataframe_pandas }
    url_cache = {} 
    
    # Set untuk melacak nama file yang sudah ada di dalam ZIP agar tidak bentrok
    used_filenames = set()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        progress_text = "Sedang memproses file..."
        my_bar = st.progress(0, text=progress_text)
        total = len(selected_indices)
        
        for i, index in enumerate(selected_indices):
            row = df_source.loc[index]
            url_target = row['link_download']
            
            # 1. Tentukan Nama File Dasar
            clean_name = f"{row['No.']}-{row['Dinas/lnstansi Pemerintah Daerah']}-{row['Judul Tabel']}".replace(" ", "_")
            # Bersihkan karakter ilegal untuk nama file windows/linux
            clean_name = "".join([c for c in clean_name if c.isalnum() or c in (' ', '.', '_', '-')]).strip()
            base_filename = f"{clean_name}.xlsx"
            
            # 2. Handle Nama File Duplikat (Agar tidak menimpa/folder-in-folder)
            final_filename = base_filename
            counter = 1
            while final_filename in used_filenames:
                # Jika nama sudah ada, tambahkan suffix (1), (2), dst
                name_without_ext = base_filename.replace(".xlsx", "")
                final_filename = f"{name_without_ext}_({counter}).xlsx"
                counter += 1
            
            used_filenames.add(final_filename)
            
            # Update Progress
            my_bar.progress((i + 1) / total, text=f"Memproses: {final_filename}")

            # 3. Smart Fetching (Cek Cache Dulu)
            df_data = None
            
            # Cek apakah data untuk URL ini sudah ada di session state (preview)
            if f"data_{index}" in st.session_state:
                df_data = st.session_state[f"data_{index}"]
            
            # Cek apakah URL ini sudah didownload di putaran loop sebelumnya (cache lokal)
            elif url_target in url_cache:
                df_data = url_cache[url_target]
            
            # Jika belum ada di mana-mana, baru fetch dari internet
            else:
                df_data = fetch_data(url_target)
                if df_data is not None:
                    # Simpan ke cache lokal agar url sama tidak perlu download ulang
                    url_cache[url_target] = df_data
                    # Opsional: Simpan juga ke session state agar nanti kalau user klik preview, datanya ada
                    st.session_state[f"data_{index}"] = df_data

            # 4. Tulis ke ZIP
            if df_data is not None and not df_data.empty:
                excel_bytes = convert_df_to_excel(df_data)
                zip_file.writestr(final_filename, excel_bytes)
            
        my_bar.empty()
        
    return zip_buffer.getvalue()

# --- FUNGSI BARU: SELECT/UNSELECT ALL ---
def toggle_all_checkboxes(df_len, target_state):
    """
    Mengubah semua session_state checkbox menjadi True/False.
    Callback ini dijalankan saat tombol diklik.
    """
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
st.title("📊 BPS DKI Scraper")
st.markdown(f"Selamat datang pada sistem listing downloading.")
st.divider()

if uploaded_file is not None:
    try:
        df_input = pd.read_excel(uploaded_file)
        
        # Validasi Keberadaan Kolom (Header)
        required_cols = ['link_download', 'No.', 'Dinas/lnstansi Pemerintah Daerah', 'Judul Tabel']
        if not all(col in df_input.columns for col in required_cols):
            st.error(f"❌ Format salah! Kolom wajib: {required_cols}")
            st.stop() # Kalau header salah, stop total karena sistem pasti error
        
        else:
            # --- LOGIKA BARU: FILTER DATA ---
            # Kolom yang tidak boleh kosong datanya
            cols_to_check = ['link_download', 'No.', 'Dinas/lnstansi Pemerintah Daerah', 'Judul Tabel']
            
            # Pisahkan data bersih dan data kotor
            df_clean = df_input.dropna(subset=cols_to_check)
            df_dirty = df_input[df_input[cols_to_check].isnull().any(axis=1)]

            # Jika ada data kotor, beritahu user tapi JANGAN STOP (kecuali semua data kotor)
            if not df_dirty.empty:
                st.warning(f"⚠️ Perhatian: Ditemukan **{len(df_dirty)} baris data tidak lengkap** (NaN) yang akan dilewati.")
                with st.expander("Lihat Data yang Bermasalah"):
                    st.dataframe(df_dirty)
            
            # Jika setelah dibersihkan datanya habis (kosong semua), baru kita stop
            if df_clean.empty:
                st.error("⛔ Semua data dalam file ini tidak lengkap/kosong. Tidak ada yang bisa diproses.")
                st.stop()

            # Ganti df_input menjadi df_clean untuk proses selanjutnya
            df_input = df_clean.reset_index(drop=True)

            st.info(f"✅ Siap Memproses **{len(df_input)}** file yang valid.")

            for col in cols_to_check:
                df_input[col] = df_input[col].astype(str)

            # ==========================================
            # AREA BULK ACTION (Sama seperti sebelumnya)
            # ==========================================
            st.markdown("### 📦 Bulk Action")
            
            
            # Layout Kontrol: Tombol Select di Kiri, Tombol ZIP di Kanan
            action_col1, action_col2 = st.columns([2, 3])
            
            with action_col1:
                st.write("**Seleksi Cepat:**")
                sub_c1, sub_c2 = st.columns(2)
                
                # TOMBOL SELECT ALL
                with sub_c1:
                    st.button(
                        "✅ Select All", 
                        on_click=toggle_all_checkboxes, 
                        args=(len(df_input), True)
                    )
                
                # TOMBOL UNSELECT ALL
                with sub_c2:
                    st.button(
                        "❌ Unselect All", 
                        on_click=toggle_all_checkboxes, 
                        args=(len(df_input), False)
                    )

            with action_col2:
                # Logic menghitung yang dipilih
                selected_indices = []
                for i in range(len(df_input)):
                    if st.session_state.get(f"check_{i}", False):
                        selected_indices.append(i)
                
                st.write(f"**Terpilih: {len(selected_indices)} file**")
                
                if selected_indices:
                    if st.button("📦 ZIP Selected Files", type="primary"):
                        with st.spinner("Sedang mengompres file..."):
                            zip_data = create_zip_archive(selected_indices, df_input)
                            st.session_state['zip_ready'] = zip_data
                            st.rerun()

            # Tampilkan tombol download ZIP jika sudah siap
            if 'zip_ready' in st.session_state:
                st.success("Arsip ZIP Siap!")
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
            
            # Header List
            h_col1, h_col2 = st.columns([0.5, 9.5])
            h_col1.markdown("**#**")
            h_col2.markdown("**Daftar File**")

            for index, row in df_input.iterrows():
                clean_name = f"{row['No.']}-{row['Dinas/lnstansi Pemerintah Daerah']}-{row['Judul Tabel']}".replace(" ", "_")
                file_name_full = f"{clean_name}.xlsx"
                url_target = row['link_download']
                
                col_check, col_exp = st.columns([0.5, 9.5])
                
                with col_check:
                    # Checkbox terhubung dengan session_state 'check_{index}'
                    # Ini kuncinya agar tombol Select All bisa mengontrol checkbox ini
                    st.checkbox("", key=f"check_{index}")

                with col_exp:
                    with st.expander(f"📄 {file_name_full}"):
                        c1, c2 = st.columns([1, 3])
                        
                        with c1:
                            st.text(f"Dinas: \n{row['Dinas/lnstansi Pemerintah Daerah']}")
                            st.text(f"Rilis: \n{row['Bulan Rilis']}")
                            if st.button("🔍 Cek Data", key=f"btn_fetch_{index}"):
                                with st.spinner('Loading...'):
                                    res = fetch_data(url_target)
                                    if res is not None:
                                        st.session_state[f"data_{index}"] = res
                                        st.success("OK")
                                    else:
                                        st.error("Gagal")

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