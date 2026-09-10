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
    html, body, [class*="st-"] {
        color: #1f1f1f;
    }
    .stTextInput input, .stSelectbox select {
        color: #1f1f1f !important;
        background-color: #ffffff !important;
    }
    </style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data():
    file_path = "REKAP KENDARAAN TA. 2026.YP.xlsx"
    if not os.path.exists(file_path):
        files = glob.glob("*.xlsx")
        if files:
            file_path = files[0]
        else:
            return None, None
    
    df_raw = pd.read_excel(file_path, header=None)
    header_row = 0
    for idx, row in df_raw.head(15).iterrows():
        row_str = " ".join(row.fillna("").astype(str)).lower()
        if "no" in row_str and any(k in row_str for k in ["jenis", "merek", "skpd", "kode", "nopol", "polisi", "barang"]):
            header_row = idx
            break
    
    df = pd.read_excel(file_path, header=header_row)
    return df, file_path

df, file_path = load_data()

if df is not None:
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.contains('^Unnamed', na=False)]

    kolom_no = None
    for col in df.columns:
        if 'no' in str(col).lower() and len(str(col)) < 10:
            kolom_no = col
            break
    if not kolom_no:
        kolom_no = df.columns[0]

    df = df.dropna(subset=[kolom_no])
    df = df[df[kolom_no].astype(str).str.strip() != '']
    df = df[~df[kolom_no].astype(str).str.lower().str.contains('jumlah|total|no', na=False)]
    df = df.reset_index(drop=True)

    # Deteksi Kolom SKPD secara presisi
    skpd_col = None
    for col in df.columns:
        c_lower = str(col).lower()
        if any(k in c_lower for k in ['skpd', 'unit kerja', 'opd', 'instansi', 'dinas']):
            skpd_col = col
            break
    if not skpd_col:
        skpd_col = df.columns[min(2, len(df.columns)-1)]

    # Fungsi Kategori yang Presisi agar sama persis dengan Excel
    def deteksi_kategori(row):
        text = " ".join(row.fillna("").astype(str)).lower()
        if any(k in text for k in ["pick up", "pickup", "bak terbuka"]):
            return "Pick Up"
        elif any(k in text for k in ["sepeda motor", "roda dua", "trail", "matic", "klx", "crf", "bebek", "scoopy", "beat", "vario", "mio", "suzuki skuter", "honda", "yamaha"]):
            # Pastikan jika itu motor spesifik atau mengandung kata sepeda motor
            if "mobil" in text or "minibus" in text:
                return "Mobil"
            return "Sepeda Motor"
        elif any(k in text for k in ["mobil", "minibus", "station", "jeep", "sedan", "bus", "truk", "truck", "doka", "double", "suv", "mpv", "pemadam", "ambulance", "dump truck"]):
            return "Mobil"
        else:
            return "Lainnya"

    df['Kategori_Jenis'] = df.apply(deteksi_kategori, axis=1)

    # Hitung Nilai Aset jika ada kolom nilai/harga
    kolom_nilai = None
    for col in df.columns:
        if any(k in str(col).lower() for k in ['nilai', 'harga', 'perolehan', 'rp']):
            kolom_nilai = col
            break

    total_nilai = 0
    if kolom_nilai:
        df['Nilai_Clean'] = pd.to_numeric(df[kolom_nilai].astype(str).str.replace(r'[^0-9]', '', regex=True), errors='fillna').fillna(0)
        total_nilai = df['Nilai_Clean'].sum()

    t_semua = len(df)
    t_motor = len(df[df['Kategori_Jenis'] == "Sepeda Motor"])
    t_mobil = len(df[df['Kategori_Jenis'] == "Mobil"])
    t_pickup = len(df[df['Kategori_Jenis'] == "Pick Up"])

    st.markdown("<h3 style='text-align: center; background-color: #2c3e50; color: white; padding: 10px; border-radius: 5px;'>SISTEM INFORMASI MANAJEMEN ASET - KENDARAAN DINAS (SIMANTAP)</h3>", unsafe_allow_html=True)
    st.write("Silakan Pilih Menu / Lihat Rekapitulasi per SKPD:")

    # Tampilan Menu / Kartu Ringkasan seperti semula
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label="🔴 SEMUA KENDARAAN", value=f"{t_semua} Data", delta=f"Rp {total_nilai:,.0f}" if kolom_nilai else None)
        st.metric(label="🟡 MOBIL", value=f"{t_mobil} Data")
    with col2:
        st.metric(label="🟢 SEPEDA MOTOR", value=f"{t_motor} Data")
        st.metric(label="🔵 PICK UP", value=f"{t_pickup} Data")

    st.markdown("---")

    # Tabel Utama & Rekap per SKPD
    tab1, tab2 = st.tabs(["📋 Tabel Data Kendaraan", "📊 Rekap Rinci per SKPD (Dinas Pendidikan, Kesehatan, dll)"])

    with tab1:
        c_f1, c_f2 = st.columns([2, 2])
        with c_f1:
            daftar_skpd = ["Semua SKPD"] + sorted([str(x) for x in df[skpd_col].dropna().unique() if str(x).strip() != ''])
            pilih_skpd_main = st.selectbox("Filter berdasarkan SKPD:", daftar_skpd, key="skpd_main")
        with c_f2:
            search_query = st.text_input("Cari cepat (Nopol, Jenis, dll):", placeholder="Ketik No Polisi...")

        df_filtered = df.copy()
        if pilih_skpd_main != "Semua SKPD":
            df_filtered = df_filtered[df_filtered[skpd_col].astype(str).str.strip() == pilih_skpd_main]

        if search_query:
            query_parts = search_query.lower().split()
            def match_row(row):
                row_str = " ".join(row.fillna("").astype(str)).lower()
                return all(part in row_str for part in query_parts)
            df_filtered = df_filtered[df_filtered.apply(match_row, axis=1)]

        display_df = df_filtered.drop(columns=["Kategori_Jenis", "Nilai_Clean"], errors="ignore").reset_index(drop=True)
        st.dataframe(display_df, use_container_width=True, height=400)

    with tab2:
        st.subheader("Rekapitulasi Jumlah Kendaraan per SKPD")
        try:
            df_rekap = df.copy()
            df_rekap[skpd_col] = df_rekap[skpd_col].fillna("").astype(str).str.strip()
            df_rekap = df_rekap[df_rekap[skpd_col] != '']

            rekap_skpd = pd.pivot_table(
                df_rekap,
                index=skpd_col,
                columns='Kategori_Jenis',
                values=kolom_no,
                aggfunc='count',
                fill_value=0
            ).reset_index()

            for kat in ['Mobil', 'Sepeda Motor', 'Pick Up', 'Lainnya']:
                if kat not in rekap_skpd.columns:
                    rekap_skpd[kat] = 0

            cols_exist = [c for c in ['Mobil', 'Sepeda Motor', 'Pick Up', 'Lainnya'] if c in rekap_skpd.columns]
            rekap_skpd['Total'] = rekap_skpd[cols_exist].sum(axis=1)
            rekap_skpd = rekap_skpd.sort_values(by='Total', ascending=False).reset_index(drop=True)
            rekap_skpd = rekap_skpd.rename(columns={skpd_col: "NAMA SKPD / DINAS"})

            # Filter spesifik per SKPD di tabel rekap agar mudah dicek (misal Dinas Kesehatan / Pendidikan)
            pilih_rekap_filter = st.selectbox("Cari / Pilih SKPD Tertentu:", ["Semua SKPD"] + list(rekap_skpd["NAMA SKPD / DINAS"].unique()))
            if pilih_rekap_filter != "Semua SKPD":
                rekap_skpd = rekap_skpd[rekap_skpd["NAMA SKPD / DINAS"] == pilih_rekap_filter]

            st.dataframe(rekap_skpd, use_container_width=True, height=400)
        except Exception as e:
            st.error(f"Gagal memuat rekap: {e}")
else:
    st.error("File Excel tidak ditemukan.")