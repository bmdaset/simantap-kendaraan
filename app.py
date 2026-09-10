import glob
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="SIMANTAP - Kendaraan Dinas",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Styling tombol agar rapi di HP
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

# Inisialisasi session state untuk navigasi tombol
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
    # --- PEMBERSIHAN DATA BERDASARKAN NOMOR URUT ---
    kolom_no = None
    for col in df.columns:
        if 'no' in str(col).lower():
            kolom_no = col
            break
    if not kolom_no:
        kolom_no = df.columns[0]

    # Membuang baris kosong atau baris total agar jumlah akurat seperti Excel
    df = df.dropna(subset=[kolom_no])
    df = df[df[kolom_no].astype(str).str.strip() != '']
    df = df[~df[kolom_no].astype(str).str.lower().str.contains('jumlah|total|no', na=False)]
    df = df.reset_index(drop=True)

    # Fungsi untuk menghitung total per kategori tombol
    def hitung_total(kategori=""):
        if kategori == "":
            return len(df)
        else:
            filtered = df[df.astype(str).apply(lambda col: col.str.lower().str.contains(kategori.lower(), na=False)).any(axis=1)]
            return len(filtered)

    t_semua = hitung_total("")
    t_motor = hitung_total("sepeda motor")
    t_mobil = hitung_total("mobil")
    t_pickup = hitung_total("pick up")

    st.markdown("## 🚗 SIMANTAP - Kendaraan Dinas")

    # --- MENU TOMBOL KATEGORI ---
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

    # Kolom Pencarian / Filter
    search_query = st.text_input("🔍 Cari Kendaraan (Nomor Polisi, Jenis, SKPD, dll):", value=st.session_state.keyword)

    filtered_df = df.copy()
    if search_query:
        mask_search = filtered_df.astype(str).apply(
            lambda col: col.str.lower().str.contains(search_query.lower(), na=False)
        ).any(axis=1)
        filtered_df = filtered_df[mask_search]

    display_df = filtered_df.drop(columns=["Harga_Clean"], errors="ignore").reset_index(drop=True)

    # --- TABEL UTAMA DENGAN FORMAT GRID EXCEL ---
    st.table(display_df)

    st.markdown("---")

    # --- REKAP JUMLAH KENDARAAN PER SKPD ---
    st.markdown("### 📊 Rekap Jumlah Kendaraan per SKPD")
    
    kolom_skpd_candidates = ['SKPD', 'Unit Kerja', 'OPD', 'Nama SKPD', 'Satuan Kerja']
    kolom_skpd = next((col for col in kolom_skpd_candidates if col in df.columns), None)
    
    if not kolom_skpd and len(df.columns) > 1:
        kolom_skpd = df.columns[1]

    if kolom_skpd:
        rekap_skpd = df.groupby(kolom_skpd).size().reset_index(name='Jumlah Kendaraan')
        rekap_skpd = rekap_skpd.sort_values(by='Jumlah Kendaraan', ascending=False).reset_index(drop=True)
        st.table(rekap_skpd)
    else:
        st.warning("Kolom SKPD tidak terdeteksi otomatis di file Excel.")

else:
    st.error("File Excel 'REKAP KENDARAAN TA. 2026.YP.xlsx' tidak ditemukan di repositori.")