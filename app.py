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
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: bold;
        padding: 10px;
        background-color: #f0f2f6;
        color: #1f1f1f;
        border: 1px solid #d6d6d6;
    }
    .stButton>button:hover {
        background-color: #e0e2e6;
        border-color: #b6b6b6;
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
        if "no" in row_str and any(k in row_str for k in ["jenis", "merek", "skpd", "kode", "nopol", "polisi"]):
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

    # Deteksi Kolom SKPD yang Akurat (Menghindari kolom jenis kendaraan)
    skpd_col = None
    max_score = -1
    for col in df.columns:
        try:
            col_series = df[col].dropna().astype(str).str.lower()
            sample_str = " ".join(col_series.head(20))
            
            # Pastikan bukan kolom jenis/deskripsi kendaraan
            if any(v in sample_str for v in ['sepeda motor', 'minibus', 'pick up', 'truk', 'cc ']):
                continue
            
            score = col_series.str.contains('dinas |badan |sekretariat |kecamatan |rsud |inspektorat |biro |satpol |pemerintah', regex=True, na=False).sum()
            if score > max_score:
                max_score = score
                skpd_col = col
        except:
            pass

    if not skpd_col:
        for col in df.columns:
            if any(k in str(col).lower() for k in ['skpd', 'unit', 'opd', 'dinas', 'instansi', 'pemilik']):
                skpd_col = col
                break
    if not skpd_col:
        skpd_col = df.columns[min(2, len(df.columns)-1)]

    def deteksi_kategori(row):
        text = " ".join(row.fillna("").astype(str)).lower()
        if "pick up" in text or "pickup" in text or "bak terbuka" in text:
            return "Pick Up"
        elif any(k in text for k in ["sepeda motor", "roda dua", "trail", "matic", "klx", "crf", "bebek", "scoopy", "beat", "vario", "mio"]):
            return "Sepeda Motor"
        elif any(k in text for k in ["mobil", "minibus", "station", "jeep", "sedan", "bus", "truk", "truck", "doka", "double", "suv", "mpv", "pemadam", "ambulance"]):
            return "Mobil"
        else:
            return "Lainnya"

    df['Kategori_Jenis'] = df.apply(deteksi_kategori, axis=1)

    t_semua = len(df)
    t_motor = len(df[df['Kategori_Jenis'] == "Sepeda Motor"])
    t_mobil = len(df[df['Kategori_Jenis'] == "Mobil"])
    t_pickup = len(df[df['Kategori_Jenis'] == "Pick Up"])

    st.markdown("<h3 style='margin-bottom: 0px;'>🚗 SIMANTAP - Kendaraan Dinas</h3>", unsafe_allow_html=True)
    st.markdown("<p style='color: gray; font-size: 13px; margin-bottom: 15px;'>Sistem Informasi Manajemen Aset & Kendaraan Dinas</p>", unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if st.button(f"🔴 Semua Data\n({t_semua})", key="b_semua"):
            st.session_state.keyword = ""
            st.session_state.title = "Semua Kendaraan"
            st.rerun()
    with col2:
        if st.button(f"🟡 Mobil\n({t_mobil})", key="b_mobil"):
            st.session_state.keyword = "mobil"
            st.session_state.title = "Kendaraan Mobil"
            st.rerun()
    with col3:
        if st.button(f"🟢 Sepeda Motor\n({t_motor})", key="b_motor"):
            st.session_state.keyword = "sepeda motor"
            st.session_state.title = "Kendaraan Sepeda Motor"
            st.rerun()
    with col4:
        if st.button(f"🔵 Pick Up\n({t_pickup})", key="b_pickup"):
            st.session_state.keyword = "pick up"
            st.session_state.title = "Kendaraan Pick Up"
            st.rerun()

    st.markdown("---")

    tab1, tab2 = st.tabs(["📋 Tabel Data Kendaraan", "📊 Rekap Rinci per SKPD"])

    with tab1:
        c_f1, c_f2 = st.columns([2, 2])
        with c_f1:
            daftar_skpd = ["Semua SKPD"] + sorted([str(x) for x in df[skpd_col].dropna().unique() if str(x).strip() != ''])
            pilih_skpd_main = st.selectbox("Filter berdasarkan SKPD:", daftar_skpd, key="skpd_main_filter")
        with c_f2:
            search_query = st.text_input("Cari cepat:", value=st.session_state.keyword, placeholder="Ketik No Polisi atau Jenis...")

        df_filtered = df.copy()
        if pilih_skpd_main != "Semua SKPD":
            df_filtered = df_filtered[df_filtered[skpd_col].astype(str).str.strip() == pilih_skpd_main]

        if search_query:
            mask = df_filtered.astype(str).apply(lambda col: col.str.lower().str.contains(search_query.lower(), na=False)).any(axis=1)
            df_filtered = df_filtered[mask]

        display_df = df_filtered.drop(columns=["Kategori_Jenis"], errors="ignore").reset_index(drop=True)
        st.dataframe(display_df, use_container_width=True, height=380)

    with tab2:
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
            rekap_skpd = rekap_skpd.rename(columns={skpd_col: "NAMA SKPD / UNIT KERJA"})

            daftar_skpd_rekap = ["Semua SKPD"] + list(rekap_skpd["NAMA SKPD / UNIT KERJA"].unique())
            pilih_skpd_rekap = st.selectbox("Filter Rekap SKPD:", daftar_skpd_rekap, key="skpd_rekap_filter")

            if pilih_skpd_rekap != "Semua SKPD":
                rekap_skpd = rekap_skpd[rekap_skpd["NAMA SKPD / UNIT KERJA"] == pilih_skpd_rekap]

            st.dataframe(rekap_skpd, use_container_width=True, height=380)
        except Exception as e:
            st.error(f"Gagal memuat rekap: {e}")
else:
    st.error("File Excel tidak ditemukan.")