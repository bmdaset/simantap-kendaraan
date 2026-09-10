import glob
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="SIMANTAP - Kendaraan Dinas",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
    <style>
    .stButton>button{
        width: 100%;
        border-radius: 8px;
        font-weight: bold;
        padding: 12px;
        font-size: 14px;
    }
    </style>
""", unsafe_allow_html=True)

if 'keyword' not in st.session_state:
    st.session_state.keyword = ""
if 'title' not in st.session_state:
    st.session_state.title = "Semua Kendaraan Dinas"

@st.cache_data
def load_data():
    file_path = "REKAP KENDARAAN TA. 2026.YP.xlsx"
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        return df, file_path
    else:
        files = glob.glob("*.xlsx")
        if files:
            df = pd.read_excel(files[0])
            return df, files[0]
        return None, None

df, file_path = load_data()

if df is not None:
    kolom_no = None
    for col in df.columns:
        if 'no' in str(col).lower():
            kolom_no = col
            break
    if not kolom_no:
        kolom_no = df.columns[0]

    df = df.dropna(subset=[kolom_no])
    df = df[df[kolom_no].astype(str).str.strip() != '']
    df = df[~df[kolom_no].astype(str).str.lower().str.contains('jumlah|total|no', na=False)]
    df = df.reset_index(drop=True)

    def deteksi_kategori(row):
        text = " ".join(row.fillna("").astype(str)).lower()
        if "sepeda motor" in text or "motor" in text or "trail" in text:
            return "Sepeda Motor"
        elif "pick up" in text or "pickup" in text:
            return "Pick Up"
        elif "mobil" in text or "minibus" in text or "jeep" in text or "station wagon" in text or "sedan" in text:
            return "Mobil"
        else:
            return "Lainnya"

    df['Kategori_Jenis'] = df.apply(deteksi_kategori, axis=1)

    def hitung_total(kategori=""):
        if kategori == "":
            return len(df)
        else:
            return len(df[df['Kategori_Jenis'].str.lower() == kategori.lower()])

    t_semua = hitung_total("")
    t_motor = hitung_total("sepeda motor")
    t_mobil = hitung_total("mobil")
    t_pickup = hitung_total("pick up")

    st.markdown("## 🚗 SIMANTAP - Kendaraan Dinas")

    col1, col2 = st.columns(2)
    with col1:
        if st.button(f"🔴 SEMUA KENDARAAN\n\nTotal: {t_semua} Data", key="btn_semua"):
            st.session_state.keyword = ""
            st.session_state.title = "Semua Kendaraan Dinas"
            st.rerun()
        
        if st.button(f"🟡 MOBIL\n\nTotal: {t_mobil} Data", key="btn_mobil"):
            st.session_state.keyword = "mobil"
            st.session_state.title = "Data Kendaraan Mobil"
            st.rerun()

    with col2:
        if st.button(f"🟢 SEPEDA MOTOR\n\nTotal: {t_motor} Data", key="btn_motor"):
            st.session_state.keyword = "sepeda motor"
            st.session_state.title = "Data Sepeda Motor"
            st.rerun()

        if st.button(f"🔵 PICK UP\n\nTotal: {t_pickup} Data", key="btn_pickup"):
            st.session_state.keyword = "pick up"
            st.session_state.title = "Data Pick Up"
            st.rerun()

    st.markdown("---")
    st.markdown(f"### 📋 {st.session_state.title}")

    search_query = st.text_input("🔍 Cari Kendaraan (Nomor Polisi, Jenis, SKPD, dll):", value=st.session_state.keyword)

    filtered_df = df.copy()
    if search_query:
        mask_search = filtered_df.astype(str).apply(
            lambda col: col.str.lower().str.contains(search_query.lower(), na=False)
        ).any(axis=1)
        filtered_df = filtered_df[mask_search]

    display_df = filtered_df.drop(columns=["Harga_Clean", "Kategori_Jenis"], errors="ignore").reset_index(drop=True)

    st.table(display_df)

    st.markdown("---")
    st.markdown("### 📊 Rekap Rinci Jumlah Kendaraan per SKPD")
    
    # Hanya ambil kolom asli dari Excel (buang kolom helper seperti Kategori_Jenis dan kolom Unnamed)
    all_columns = [col for col in df.columns if col not in ['Kategori_Jenis', 'Harga_Clean'] and not str(col).startswith('Unnamed')]
    
    default_idx = 0
    for i, col in enumerate(all_columns):
        col_lower = str(col).lower()
        if any(k in col_lower for k in ['skpd', 'unit', 'opd', 'nama', 'satuan', 'dinas', 'keterangan']):
            default_idx = i
            break

    selected_skpd_col = st.selectbox("Pilih Kolom untuk Nama SKPD:", all_columns, index=default_idx)

    if selected_skpd_col:
        try:
            df_rekap = df.dropna(subset=[selected_skpd_col]).copy()
            df_rekap[selected_skpd_col] = df_rekap[selected_skpd_col].astype(str).str.strip()
            df_rekap = df_rekap[df_rekap[selected_skpd_col] != '']

            rekap_skpd = pd.pivot_table(
                df_rekap,
                index=selected_skpd_col,
                columns='Kategori_Jenis',
                values=kolom_no,
                aggfunc='count',
                fill_value=0
            ).reset_index()

            for kat in ['Mobil', 'Sepeda Motor', 'Pick Up', 'Lainnya']:
                if kat not in rekap_skpd.columns:
                    rekap_skpd[kat] = 0

            kolom_tersedia = [c for c in ['Mobil', 'Sepeda Motor', 'Pick Up', 'Lainnya'] if c in rekap_skpd.columns]
            rekap_skpd['Total'] = rekap_skpd[kolom_tersedia].sum(axis=1)
            rekap_skpd = rekap_skpd.sort_values(by='Total', ascending=False).reset_index(drop=True)

            st.table(rekap_skpd)
        except Exception as e:
            st.warning("Silakan pilih kolom lain pada pilihan di atas yang berisi teks nama SKPD/Unit Kerja yang valid.")
    else:
        st.warning("Silakan pilih kolom SKPD yang sesuai.")
else:
    st.error("File Excel 'REKAP KENDARAAN TA. 2026.YP.xlsx' tidak ditemukan di repositori.")
