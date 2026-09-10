import glob
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="SIMANTAP - Kendaraan Dinas",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: bold;
        padding: 15px;
        font-size: 14px;
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

    # Penamaan Kolom Standar
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
          .fillna("DINAS / INSTANSI LAINNYA")
          .astype(str)
          .str.upper()
          .str.strip()
      )
    else:
      temp_df["SKPD_Nama"] = "DINAS / INSTANSI LAINNYA"

    # Filter baris data valid berdasarkan nomor urut / No.
    df = temp_df[temp_df["No."].astype(str).str.contains(r"\d", na=False)].copy()
    df = df.reset_index(drop=True)

    # Pembersihan Harga
    if "Harga (Rp)" in df.columns:
      df["Harga_Clean"] = (
          df["Harga (Rp)"]
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

    # Klasifikasi Kategori Kendaraan Presisi
    def deteksi_kategori(row):
      text = str(row.get("Jenis / Nama Barang", "")).lower()
      merk = str(row.get("Merk / Type", "")).lower()
      combined = f"{text} {merk}"

      if any(k in combined for k in ["pick up", "pickup", "bak terbuka"]):
        return "Pick Up"
      elif any(
          k in combined
          for k in [
              "station wagon",
              "minibus",
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

st.markdown(
    """
    <div style="background-color:#2c3e50;padding:15px;border-radius:6px;text-align:center;margin-bottom:20px;">
        <h3 style="color:white;margin:0;font-family:Segoe UI;">SISTEM INFORMASI MANAJEMEN ASET - KENDARAAN DINAS (SIMANTAP)</h3>
    </div>
    """,
    unsafe_allow_html=True,
)

if st.session_state.page == "menu":
  st.markdown(
      "<h4 style='text-align:center; color:#2c3e50;'>Silakan Pilih Menu Kategori"
      " Aset:</h4>",
      unsafe_allow_html=True,
  )
  st.write("")

  p_semua, t_semua = hitung_total("")
  p_motor, t_motor = hitung_total("Sepeda Motor")
  p_mobil, t_mobil = hitung_total("Mobil")
  p_pickup, t_pickup = hitung_total("Pick Up")

  col1, col2 = st.columns(2)

  with col1:
    if st.button(
        f"📦 SEMUA KENDARAAN\n\nTotal: {t_semua} Data\nNilai: {p_semua}",
        key="btn_semua",
    ):
      st.session_state.keyword = ""
      st.session_state.title = "Semua Kendaraan Dinas"
      st.session_state.page = "table"
      st.rerun()

    st.write("")
    if st.button(
        f"🚗 MOBIL\n\nTotal: {t_mobil} Data\nNilai: {p_mobil}", key="btn_mobil"
    ):
      st.session_state.keyword = "Mobil"
      st.session_state.title = "Mobil"
      st.session_state.page = "table"
      st.rerun()

  with col2:
    if st.button(
        f"🏍️ SEPEDA MOTOR\n\nTotal: {t_motor} Data\nNilai: {p_motor}",
        key="btn_motor",
    ):
      st.session_state.keyword = "Sepeda Motor"
      st.session_state.title = "Sepeda Motor"
      st.session_state.page = "table"
      st.rerun()

    st.write("")
    if st.button(
        f"🚙 PICK UP\n\nTotal: {t_pickup} Data\nNilai: {p_pickup}",
        key="btn_pickup",
    ):
      st.session_state.keyword = "Pick Up"
      st.session_state.title = "Pick Up"
      st.session_state.page = "table"
      st.rerun()

  # Bagian Preview Rekapitulasi Rincian Berdasarkan Kolom 19 (Per Nama SKPD)
  st.markdown("---")
  st.markdown(
      "#### 📊 Preview Jumlah Kendaraan per Nama SKPD (Berdasarkan Kolom 19)"
  )
  if not df.empty:
    rekap_skpd = (
        pd.pivot_table(
            df,
            index="SKPD_Nama",
            columns="Kategori_Jenis",
            values="No.",
            aggfunc="count",
            fill_value=0,
        )
        .reset_index()
    )

    for kat in ["Sepeda Motor", "Mobil", "Pick Up", "Lainnya"]:
      if kat not in rekap_skpd.columns:
        rekap_skpd[kat] = 0

    cols_exist = ["Sepeda Motor", "Mobil", "Pick Up", "Lainnya"]
    rekap_skpd["Jumlah Total"] = rekap_skpd[
        [c for c in cols_exist if c in rekap_skpd.columns]
    ].sum(axis=1)
    rekap_skpd = rekap_skpd.sort_values(
        by="Jumlah Total", ascending=False
    ).reset_index(drop=True)

    rekap_skpd = rekap_skpd.rename(
        columns={
            "SKPD_Nama": "Nama SKPD / Dinas",
            "Sepeda Motor": "Jumlah Sepeda Motor",
            "Mobil": "Jumlah Mobil",
            "Pick Up": "Jumlah Pick Up",
            "Lainnya": "Jumlah Lainnya",
        }
    )

    st.dataframe(rekap_skpd, use_container_width=True, height=400)

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