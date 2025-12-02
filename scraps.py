
import streamlit as st
import pandas as pd
import requests
import io
import xlsxwriter 
import zipfile # Library tambahan untuk membuat ZIP

# ==========================================
# KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(
    page_title="BPS DKI Scraper",
    page_icon="📊",
    layout="wide"
)

# ==========================================
# FUNGSI UTILITAS
# ==========================================
@st.cache_data(show_spinner=False)
def fetch_data(url):
    """Mengambil data dari API dan mengembalikan DataFrame"""
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        try:
            # Coba parsing JSON
            data_json = response.json()
            content = data_json.get('data', data_json)
            return pd.DataFrame(content)
        except ValueError:
            # Jika Excel/CSV binary
            return pd.read_excel(io.BytesIO(response.content))
    except Exception as e:
        return None

def convert_df_to_excel(df):
    """Mengubah DataFrame menjadi binary Excel untuk didownload"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

def create_zip_archive(selected_indices, df_source):
    """
    Membuat file ZIP berisi file-file Excel dari index yang dipilih.
    Melakukan fetching otomatis jika data belum ada di session state.
    """
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        progress_text = "Sedang memproses file..."
        my_bar = st.progress(0, text=progress_text)
        
        total = len(selected_indices)
        
        for i, index in enumerate(selected_indices):
            row = df_source.loc[index]
            
            # Konstruksi Nama File
            clean_name = f"{row['Penamaan_Data']}_{row['PIC']}_{row['Bulan_rilis']}".replace(" ", "_")
            file_name_full = f"{clean_name}.xlsx"
            url_target = row['link_download']
            
            # Update Progress Bar
            my_bar.progress((i + 1) / total, text=f"Memproses: {file_name_full}")
            
            # Cek apakah data sudah ada di session state (sudah di-preview)
            # Jika belum, kita fetch sekarang
            if f"data_{index}" in st.session_state:
                df_data = st.session_state[f"data_{index}"]
            else:
                df_data = fetch_data(url_target)
                # Simpan ke session biar gak fetch ulang kalau mau preview nanti
                if df_data is not None:
                    st.session_state[f"data_{index}"] = df_data
            
            # Tulis ke dalam ZIP jika data valid
            if df_data is not None and not df_data.empty:
                excel_bytes = convert_df_to_excel(df_data)
                zip_file.writestr(file_name_full, excel_bytes)
            
        my_bar.empty() # Hapus progress bar setelah selesai
        
    return zip_buffer.getvalue()

# ==========================================
# SIDEBAR (UPLOAD & RESET)
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

        # 1. VALIDASI NAMA KOLOM
        required_cols = ['link_download', 'Penamaan_Data', 'PIC', 'Bulan_rilis']
        if not all(col in df_input.columns for col in required_cols):
            st.error(f"❌ Format salah! File wajib memiliki kolom: {required_cols}")
        
        else:
            # 2. VALIDASI ISI DATA (TIDAK BOLEH NaN)
            cols_to_check = ['Penamaan_Data', 'PIC', 'Bulan_rilis']
            
            if df_input[cols_to_check].isnull().any().any():
                st.error("⛔ **VALIDASI GAGAL: Data Tidak Lengkap!**")
                st.warning("Kolom 'Penamaan_Data', 'PIC', atau 'Bulan_rilis' tidak boleh kosong (NaN).")
                error_rows = df_input[df_input[cols_to_check].isnull().any(axis=1)]
                st.dataframe(error_rows, use_container_width=True)
                st.stop() 

            # --- JIKA LOLOS VALIDASI ---
            st.info(f"✅ Validasi Sukses. Ditemukan {len(df_input)} target file.")

            # Pastikan tipe data string
            for col in cols_to_check:
                df_input[col] = df_input[col].astype(str)

            # ==========================================
            # AREA BULK DOWNLOAD (TOMBOL DOWNLOAD SEMUA)
            # ==========================================
            st.markdown("### 📦 Bulk Action")
            
            # Wadah untuk tombol download zip
            bulk_col1, bulk_col2 = st.columns([1, 4])
            
            # Logic mencari mana saja yang dicentang
            selected_indices = []
            for i in range(len(df_input)):
                if st.session_state.get(f"check_{i}", False):
                    selected_indices.append(i)
            
            with bulk_col1:
                if selected_indices:
                    st.write(f"Terpilih: **{len(selected_indices)} file**")
                    
                    # Tombol Generate ZIP
                    # Kita gunakan callback sederhana logic button -> generate -> download_button
                    if st.button("📦 ZIP Selected Files"):
                        with st.spinner("Sedang mengompres file..."):
                            zip_data = create_zip_archive(selected_indices, df_input)
                            st.session_state['zip_ready'] = zip_data
                            st.rerun() # Rerun untuk memunculkan tombol download di bawah
                else:
                    st.write("Belum ada file dipilih.")

            with bulk_col2:
                # Jika ZIP sudah siap, tampilkan tombol download final
                if 'zip_ready' in st.session_state:
                    st.download_button(
                        label="⬇️ KLIK UNTUK UNDUH HASIL ZIP",
                        data=st.session_state['zip_ready'],
                        file_name="BPS_Data_Archive.zip",
                        mime="application/zip",
                        type="primary"
                    )

            st.markdown("---")

            # ==========================================
            # LOOP UTAMA (LIST FILE DENGAN CHECKBOX)
            # ==========================================
            
            # Header Tabel Sederhana
            h_col1, h_col2 = st.columns([0.5, 9.5])
            h_col1.markdown("**Pilih**")
            h_col2.markdown("**Daftar File**")

            for index, row in df_input.iterrows():
                clean_name = f"{row['Penamaan_Data']}_{row['PIC']}_{row['Bulan_rilis']}".replace(" ", "_")
                file_name_full = f"{clean_name}.xlsx"
                url_target = row['link_download']
                
                # Layout: Checkbox di kiri, Expander di kanan
                col_check, col_exp = st.columns([0.5, 9.5])
                
                with col_check:
                    # Checkbox seleksi
                    st.checkbox("", key=f"check_{index}")

                with col_exp:
                    with st.expander(f"📄 {file_name_full}"):
                        col1, col2 = st.columns([1, 3])
                        
                        with col1:
                            st.markdown("**Detail File:**")
                            st.text(f"PIC: {row['PIC']}")
                            st.text(f"Rilis: {row['Bulan_rilis']}")
                            
                            fetch_key = f"btn_fetch_{index}"
                            
                            # Tombol Preview/Cek (Tetap ada untuk cek manual)
                            if st.button("🔍 Cek & Siapkan Data", key=fetch_key):
                                with st.spinner('Sedang menghubungi server...'):
                                    df_result = fetch_data(url_target)
                                    if df_result is not None and not df_result.empty:
                                        st.session_state[f"data_{index}"] = df_result
                                        st.success("Data siap!")
                                    else:
                                        st.error("Gagal mengambil data.")

                        with col2:
                            # Jika data sudah di-fetch (baik lewat tombol Cek atau lewat proses ZIP sebelumnya)
                            if f"data_{index}" in st.session_state:
                                df_show = st.session_state[f"data_{index}"]
                                st.subheader("Preview Data")
                                st.dataframe(df_show.tail(5), use_container_width=True)
                                
                                excel_data = convert_df_to_excel(df_show)
                                
                                # Tombol Download Satuan
                                st.download_button(
                                    label="⬇️ DOWNLOAD FILE INI SAJA (XLSX)",
                                    data=excel_data,
                                    file_name=file_name_full,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=f"btn_down_{index}",
                                )

    except Exception as e:
        st.error(f"Terjadi kesalahan teknis: {e}")

else:
    st.markdown("""
    <div style='text-align: center; color: gray; padding: 50px;'>
        Waiting for file upload...
    </div>
    """, unsafe_allow_html=True)