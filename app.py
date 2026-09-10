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

    if len(temp_df.columns) >= len(kolom_kib_b):
      temp_df.columns = kolom_kib_b[: len(kolom_kib_b)] + [
          f"Kolom {i+1}"
          for i in range(len(kolom_kib_b), len(temp_df.columns))
      ]
    else:
      temp_df.columns = kolom_kib_b[: len(temp_df.columns)]

    df = temp_df

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

    return df, file_path
  except Exception:
    return pd.DataFrame(), None


df, file_path = load_data()


def format_rupiah(nilai):
  formatted = f"{nilai:,.2f}"
  return "Rp " + formatted.replace(",", "X").replace(".", ",").replace("X", ".")


def hitung_total(keyword):
  if df.empty:
    return format_rupiah(0), 0
  if keyword == "":
    sub = df
  else:
    mask = (
        df.astype(str)
        .apply(
            lambda x: x.str.lower().str.contains(keyword.lower(), na=False)
        )
        .any(axis=1)
    )
    sub = df[mask]
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
  p_motor, t_motor = hitung_total("sepeda motor")
  p_mobil, t_mobil = hitung_total("mobil")
  p_pickup, t_pickup = hitung_total("pick up")

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
      st.session_state.keyword = "mobil"
      st.session_state.title = "Mobil"
      st.session_state.page = "table"
      st.rerun()

  with col2:
    if st.button(
        f"🏍️ SEPEDA MOTOR\n\nTotal: {t_motor} Data\nNilai: {p_motor}",
        key="btn_motor",
    ):
      st.session_state.keyword = "sepeda motor"
      st.session_state.title = "Sepeda Motor"
      st.session_state.page = "table"
      st.rerun()

    st.write("")
    if st.button(
        f"🚙 PICK UP\n\nTotal: {t_pickup} Data\nNilai: {p_pickup}",
        key="btn_pickup",
    ):
      st.session_state.keyword = "pick up"
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
    mask = (
        df.astype(str)
        .apply(
            lambda col: col.str.lower().str.contains(keyword.lower(), na=False)
        )
        .any(axis=1)
    )
    filtered_df = df[mask]

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
      columns=["Harga_Clean"], errors="ignore"
  ).reset_index(drop=True)
  st.table(display_df)

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
