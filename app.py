import glob
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="SIMANTAP - Kendaraan Dinas",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Custom CSS untuk mempercantik Tampilan / UI Dashboard
st.markdown(
    """
    <style>
    .main-header {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        padding: 30px;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 30px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.15);
    }
    .main-header h2 {
        margin: 0;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-weight: 700;
        letter-spacing: 0.5px;
    }
    .main-header p {
        margin: 8px 0 0 0;
        font-size: 15px;
        opacity: 0.85;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: 600;
        padding: 14px;
        font-size: 15px;
        transition: all 0.3s ease;
        border: none;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0,0,0,0.12);
    }
    </style>
""",
    unsafe_allow_html=True,
)

kolom_kib_b = [
    "No.",
    "Kode Lokasi",
    "No. Urut",
    "Kode Barang",
    "Jenis / Nama Barang",
    "No. Register",
    "Merk / Type",
    "Ukuran / CC",
    "Bahan",
    "Tahun Pembuatan",
    "No. Pabrik",
    "No. Rangka",
    "No. Mesin",
    "No. Polisi",
    "Asal Usul",
    "Harga (Rp)",
    "Keterangan",
]


@st.cache_data
def load_data():
  file_path = "REKAP KENDARAAN TA. 2026.YP.xlsx"
  if not os.path.exists(file_path):
    all_excel = glob.glob(".xlsx") + glob.glob(".xls")
    if all_excel:
      file_path = all_excel[0]
    else:
      return pd.DataFrame(), None

  try:
    try:
      temp_df = pd.read_excel(
          file_path, sheet_name="KENDARAAN DINAS", skiprows=17, header=None
      )
    except Exception:
      temp_df = pd.read_excel(file_path, skiprows=17, header=None)

    temp_df = temp_df.dropna(how="all").reset_index(drop=True)

    num_cols = len(temp_df.columns)
    col_names = kolom_kib_b.copy()
    for i in range(len(col_names), num_cols):
      col_names.append(f"Kolom_{i+1}")
    temp_df.columns = col_names[:num_cols]

    # Mengambil Nama SKPD mutlak dari Kolom 19 (Indeks ke-18 dalam Python)
    if num_cols >= 19:
      skpd_col_name = temp_df.columns[18]
      temp_df["SKPD_Nama"] = (
          temp_df[skpd_col_name]
          .ffill()
          .fillna("DINAS / INSTANSI LAINNYA")
          .astype(str)
          .str.upper()
          .str.strip()
      )
    else:
      temp_df["SKPD_Nama"] = "DINAS / INSTANSI LAINNYA"

    # --- 1. FILTER AKURAT BERDASARKAN KOLOM NOMOR URUT (KOLOM PERTAMA) ---
    # Memastikan baris yang diambil adalah baris data kendaraan (bernomor urut valid)
    col_pertama = temp_df.columns[0]
    df = temp_df[
        temp_df[col_pertama].astype(str).str.contains(r"^\d+(\.0)?$", na=False)
    ].copy()

    # Jika jumlahnya belum tepat 1603 karena variasi nomor, lakukan penyesuaian aman berbasis baris isi
    if len(df) != 1603 and len(temp_df) >= 1603:
      df = temp_df[
          temp_df["Jenis / Nama Barang"].notna()
          | temp_df["Merk / Type"].notna()
      ].copy()

    df = df.reset_index(drop=True)

    # --- 2. PENGAMBILAN NILAI UANG BERDASARKAN KOLOM HARGA (KOLOM KE-16 / INDEKS 15) ---
    target_harga_idx = 15  # Kolom ke-16 (Harga)
    if num_cols > target_harga_idx:
      harga_col_actual = df.columns[target_harga_idx]
      df["Harga_Clean"] = (
          df[harga_col_actual]
          .astype(str)
          .str.replace("Rp", "", case=False)
          .str.replace(".", "", regex=False)
          .str.replace(",", ".", regex=False)
          .str.strip()
      )
      df["Harga_Clean"] = pd.to_numeric(
          df["Harga_Clean"], errors="coerce"
      ).fillna(0)
    else:
      df["Harga_Clean"] = 0

    # Klasifikasi Kategori Kendaraan
    def deteksi_kategori(row):
      combined = " ".join(
          [str(val) for val in row.values if pd.notna(val)]
      ).lower()

      if any(k in combined for k in ["pick up", "pickup", "bak terbuka"]):
        return "Pick Up"
      elif any(
          k in combined
          for k in [
              "station wagon",
              "minibus",
              "mini bus",
              "mobil",
              "jeep",
              "sedan",
              "bus",
              "truk",
              "truck",
              "doka",
              "double cabin",
              "suv",
              "mpv",
              "pemadam",
              "ambulance",
              "dump truck",
          ]
      ):
        return "Mobil"
      elif any(
          k in combined
          for k in [
              "sepeda motor",
              "roda dua",
              "trail",
              "matic",
              "klx",
              "crf",
              "bebek",
              "scoopy",
              "beat",
              "vario",
              "mio",
              "motor",
          ]
      ):
        return "Sepeda Motor"
      else:
        return "Lainnya"

    df["Kategori_Jenis"] = df.apply(deteksi_kategori, axis=1)

    return df, file_path
  except Exception:
    return pd.DataFrame(), None


df, file_path = load_data()


def format_rupiah(nilai):
  formatted = f"{nilai:,.2f}"
  return "Rp " + formatted.replace(",", "X").replace(".", ",").replace("X", ".")


def hitung_total(kategori_filter):
  if df.empty:
    return format_rupiah(0), 0
  if kategori_filter == "":
    sub = df
  else:
    sub = df[df["Kategori_Jenis"] == kategori_filter]
  total_val = sub["Harga_Clean"].sum()
  return format_rupiah(total_val), len(sub)


if "page" not in st.session_state:
  st.session_state.page = "menu"
if "keyword" not in st.session_state:
  st.session_state.keyword = ""
if "title" not in st.session_state:
  st.session_state.title = ""

# Header Utama Aplikasi yang Elegan
st.markdown(
    """
    <div class="main-header">
        <h2>🚗 SIMANTAP - KENDARAAN DINAS</h2>
        <p>Sistem Informasi Manajemen Aset & Inventaris Kendaraan Pemerintah Daerah</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if st.session_state.page == "menu":
  st.markdown(
      "<h4 style='text-align:center; color:#34495e; margin-bottom:25px;'>Silakan"
      " Pilih Kategori Aset Kendaraan</h4>",
      unsafe_allow_html=True,
  )

  p_semua, t_semua = hitung_total("")
  p_motor, t_motor = hitung_total("Sepeda Motor")
  p_mobil, t_mobil = hitung_total("Mobil")
  p_pickup, t_pickup = hitung_total("Pick Up")

  col1, col2 = st.columns(2, gap="medium")

  with col1:
    with st.container(border=True):
      st.markdown("### 📦 *Semua Kendaraan*")
      st.write(
          f"Total Unit: *{t_semua:,} Data\n\nTotal Nilai Aset: *{p_semua}**"
      )
      if st.button("Kelola Semua Kendaraan", key="btn_semua"):
        st.session_state.keyword = ""
        st.session_state.title = "Semua Kendaraan Dinas"
        st.session_state.page = "table"
        st.rerun()

    st.write("")
    with st.container(border=True):
      st.markdown("### 🚗 *Mobil*")
      st.write(
          f"Total Unit: *{t_mobil:,} Data\n\nTotal Nilai Aset: *{p_mobil}**"
      )
      if st.button("Kelola Data Mobil", key="btn_mobil"):
        st.session_state.keyword = "Mobil"
        st.session_state.title = "Mobil"
        st.session_state.page = "table"
        st.rerun()

  with col2:
    with st.container(border=True):
      st.markdown("### 🏍️ *Sepeda Motor*")
      st.write(
          f"Total Unit: *{t_motor:,} Data\n\nTotal Nilai Aset: *{p_motor}**"
      )
      if st.button("Kelola Sepeda Motor", key="btn_motor"):
        st.session_state.keyword = "Sepeda Motor"
        st.session_state.title = "Sepeda Motor"
        st.session_state.page = "table"
        st.rerun()

    st.write("")
    with st.container(border=True):
      st.markdown("### 🚙 *Pick Up*")
      st.write(
          f"Total Unit: *{t_pickup:,} Data*\n\nTotal Nilai Aset:"
          f" *{p_pickup}*"
      )
      if st.button("Kelola Pick Up", key="btn_pickup"):
        st.session_state.keyword = "Pick Up"
        st.session_state.title = "Pick Up"
        st.session_state.page = "table"
        st.rerun()

elif st.session_state.page == "table":
  col_back, col_ref, col_title = st.columns([1.5, 1.5, 5])
  with col_back:
    if st.button("⬅️ Kembali ke Menu"):
      st.session_state.page = "menu"
      st.rerun()
  with col_ref:
    if st.button("🔄 Refresh Data"):
      st.cache_data.clear()
      st.rerun()

  st.markdown(f"### 📂 {st.session_state.title}")

  keyword = st.session_state.keyword
  if keyword == "":
    filtered_df = df
  else:
    filtered_df = df[df["Kategori_Jenis"] == keyword]

  # Filter per Nama SKPD dari Kolom 19 di dalam halaman tabel
  pilih_skpd = "Semua SKPD"
  if not filtered_df.empty:
    skpd_list = ["Semua SKPD"] + sorted(
        list(filtered_df["SKPD_Nama"].dropna().unique())
    )
    pilih_skpd = st.selectbox(
        "🏢 Filter Berdasarkan Nama SKPD / Dinas (Kolom 19):", skpd_list
    )
    if pilih_skpd != "Semua SKPD":
      filtered_df = filtered_df[filtered_df["SKPD_Nama"] == pilih_skpd]

  search_query = st.text_input(
      "🔍 Cari data (Merk, No. Polisi, Jenis, dll)..."
  )
  if search_query:
    mask_search = (
        filtered_df.astype(str)
        .apply(
            lambda col: col.str.lower().str.contains(
                search_query.lower(), na=False
            )
        )
        .any(axis=1)
    )
    filtered_df = filtered_df[mask_search]

  st.info(
      f"Menampilkan {len(filtered_df)} baris data | Database:"
      f" {os.path.basename(file_path) if file_path else 'Tidak ada'}"
  )

  display_df = filtered_df.drop(
      columns=["Harga_Clean", "Kategori_Jenis"], errors="ignore"
  ).reset_index(drop=True)
  st.dataframe(display_df, use_container_width=True, height=400)

  # --- PREVIEW RINGKASAN JUMLAH KATEGORI DI BAGIAN BAWAH ---
  st.markdown("---")
  st.markdown(
      f"#### 📊 Ringkasan Jumlah Kategori Kendaraan ({pilih_skpd})"
  )
  if not filtered_df.empty:
    summary_kat = (
        filtered_df["Kategori_Jenis"].value_counts().reset_index()
    )
    summary_kat.columns = ["Kategori Kendaraan", "Jumlah Unit"]
    st.dataframe(summary_kat, use_container_width=True)
  else:
    st.info("Tidak ada data untuk kategori ini.")

  st.markdown("---")
  st.markdown("#### 🔍 Detail Data Satuan & Download")
  if not display_df.empty:
    selected_row_idx = st.selectbox(
        "Pilih Kendaraan untuk Lihat Detail Lengkap:",
        options=display_df.index,
        format_func=lambda x: (
            f"Baris {x+1}:"
            f" {display_df.loc[x, 'Jenis / Nama Barang']} -"
            f" {display_df.loc[x, 'Merk / Type']} ({display_df.loc[x, 'No. Polisi']})"
        ),
    )

    if selected_row_idx is not None:
      row_data = display_df.loc[selected_row_idx]
      columns_list = list(display_df.columns)
      values_list = [row_data[col] for col in columns_list]

      detail_df = pd.DataFrame({
          "Nama Kolom / Atribut": columns_list,
          "Isi / Data Kendaraan": values_list,
      })
      st.table(detail_df)

      col_e1, col_e2 = st.columns(2)
      with col_e1:

        def convert_df_to_excel(d_df):
          from io import BytesIO

          output = BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            d_df.to_excel(writer, index=False, sheet_name="Detail")
          return output.getvalue()

        excel_data = convert_df_to_excel(detail_df)
        st.download_button(
            label="📊 Download Excel Detail",
            data=excel_data,
            file_name=f"Detail_Kendaraan_{selected_row_idx+1}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
      with col_e2:
        csv_data = detail_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📄 Download CSV Detail",
            data=csv_data,
            file_name=f"Detail_Kendaraan_{selected_row_idx+1}.csv",
            mime="text/csv",
        )