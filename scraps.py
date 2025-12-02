
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
    """Membuat file ZIP berisi file Excel terpilih"""
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        progress_text = "Sedang memproses file..."
        my_bar = st.progress(0, text=progress_text)
        total = len(selected_indices)
        
        for i, index in enumerate(selected_indices):
            row = df_source.loc[index]
            clean_name = f"{row['Penamaan_Data']}_{row['PIC']}_{row['Bulan_rilis']}".replace(" ", "_")
            file_name_full = f"{clean_name}.xlsx"
            url_target = row['link_download']
            
            my_bar.progress((i + 1) / total, text=f"Memproses: {file_name_full}")
            
            # Cek session state atau fetch baru
            if f"data_{index}" in st.session_state:
                df_data = st.session_state[f"data_{index}"]
            else:
                df_data = fetch_data(url_target)
                if df_data is not None:
                    st.session_state[f"data_{index}"] = df_data
            
            if df_data is not None and not df_data.empty:
                excel_bytes = convert_df_to_excel(df_data)
                zip_file.writestr(file_name_full, excel_bytes)
            
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
        
        # Validasi Kolom
        required_cols = ['link_download', 'Penamaan_Data', 'PIC', 'Bulan_rilis']
        if not all(col in df_input.columns for col in required_cols):
            st.error(f"❌ Format salah! Kolom wajib: {required_cols}")
        
        else:
            cols_to_check = ['Penamaan_Data', 'PIC', 'Bulan_rilis']
            if df_input[cols_to_check].isnull().any().any():
                st.error("⛔ Validasi Gagal: Ada data kosong (NaN).")
                st.stop()

            st.info(f"✅ Validasi Sukses. Total: {len(df_input)} file.")

            for col in cols_to_check:
                df_input[col] = df_input[col].astype(str)

            # ==========================================
            # AREA BULK ACTION
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
                clean_name = f"{row['Penamaan_Data']}_{row['PIC']}_{row['Bulan_rilis']}".replace(" ", "_")
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
                            st.text(f"PIC: {row['PIC']}")
                            st.text(f"Rilis: {row['Bulan_rilis']}")
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