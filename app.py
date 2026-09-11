from io import BytesIO
import glob
import os
import re
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="SIMANTAP - Manajemen Aset Daerah",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(135deg, #e0c3fc 0%, #8ec5fc 100%);
    }
    .main-header {
        background: linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%);
        padding: 30px;
        border-radius: 16px;
        color: white;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 8px 25px rgba(0,0,0,0.25);
    }
    .kib-card-kendaraan {
        background: linear-gradient(135deg, #2b5876 0%, #4e4376 100%);
        padding: 22px;
        border-radius: 14px;
        color: white;
        box-shadow: 0 6px 20px rgba(43, 88, 118, 0.35);
        margin-bottom: 15px;
    }
    .kib-card-tanah {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        padding: 22px;
        border-radius: 14px;
        color: white;
        box-shadow: 0 6px 20px rgba(17, 153, 142, 0.35);
        margin-bottom: 15px;
    }
    .stButton>button {
        width: 100%;
        border-radius: 10px;
        font-weight: 600;
        padding: 10px;
        font-size: 14px;
        border: none;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    </style>
""",
    unsafe_allow_html=True,
)


@st.cache_data
def load_all_data():
  file_path = "REKAP KENDARAAN TA. 2026.YP.xlsx"
  if not os.path.exists(file_path):
    all_excel = glob.glob(".xlsx") + glob.glob(".xls")
    if all_excel:
      file_path = all_excel[0]
    else:
      return pd.DataFrame(), pd.DataFrame(), None

  # 1. Load KIB B (KENDARAAN DINAS) Sesuai Format Asli Excel & Permendagri
  df_kendaraan = pd.DataFrame()
  try:
    df_raw_k = pd.read_excel(
        file_path, sheet_name="KENDARAAN DINAS", header=None
    )
    h_idx_k = 15
    for idx, row in df_raw_k.head(25).iterrows():
      txt = " ".join(str(v) for v in row.values).lower()
      if "kode barang" in txt or "jenis" in txt or "merk" in txt:
        h_idx_k = idx
        break

    df_k = pd.read_excel(
        file_path, sheet_name="KENDARAAN DINAS", header=h_idx_k
    )
    df_k = df_k.loc[:, ~df_k.columns.astype(str).str.contains("^Unnamed")]
    df_k.columns = [str(c).strip() for c in df_k.columns]
    df_k = df_k.dropna(how="all").reset_index(drop=True)

    first_col = df_k.columns[0]

    def is_valid_row(val):
      try:
        return int(float(val)) > 0
      except:
        return False

    df_k = df_k[df_k[first_col].apply(is_valid_row)].reset_index(drop=True)

    for col in df_k.columns:
      df_k[col] = (
          df_k[col]
          .astype(str)
          .str.replace(r"\.0+$", "", regex=True)
          .replace("nan", "")
          .replace("None", "")
      )

    skpd_col_k = next(
        (c for c in df_k.columns if "skpd" in c.lower() or "dinas" in c.lower()),
        None,
    )
    df_k["SKPD_Nama"] = (
        df_k[skpd_col_k]
        .ffill()
        .fillna("DINAS / INSTANSI LAINNYA")
        .astype(str)
        .str.upper()
        .str.strip()
        if skpd_col_k
        else "DINAS / INSTANSI LAINNYA"
    )

    harga_col_k = next(
        (
            c
            for c in df_k.columns
            if "harga" in c.lower() or "rupiah" in c.lower()
        ),
        None,
    )
    df_k["Harga_Clean"] = (
        pd.to_numeric(
            df_k[harga_col_k].str.replace(r"[^\d.]", "", regex=True),
            errors="coerce",
        ).fillna(0)
        if harga_col_k
        else 0
    )

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

    df_k["Kategori_Jenis"] = df_k.apply(deteksi_kategori, axis=1)
    df_kendaraan = df_k
  except Exception as e:
    print("Error KENDARAAN:", e)

  # 2. Load KIB A (KIB A TANAH) - Dipertahankan Sesuai Script Asli Tanpa Ubah
  df_tanah = pd.DataFrame()
  try:
    df_raw_t = pd.read_excel(file_path, sheet_name="KIB A TANAH", header=None)
    h_idx_t = 0
    for idx, row in df_raw_t.head(20).iterrows():
      txt = " ".join(str(v) for v in row.values).lower()
      if "luas" in txt or "letak" in txt or "hak" in txt:
        h_idx_t = idx
        break

    df_t = pd.read_excel(
        file_path, sheet_name="KIB A TANAH", header=[h_idx_t, h_idx_t + 1]
    )
    df_t.columns = [
        " ".join([str(c) for c in col if "unnamed" not in str(c).lower()]).strip()
        for col in df_t.columns
    ]
    if len(df_t.columns) == 0 or all(c == "" for c in df_t.columns):
      df_t = pd.read_excel(
          file_path, sheet_name="KIB A TANAH", header=h_idx_t + 1
      )

    df_t = df_t.loc[:, ~df_t.columns.astype(str).str.contains("^Unnamed")]
    df_t = df_t.dropna(how="all").reset_index(drop=True)

    first_col_t = df_t.columns[0]
    df_t = df_t[
        df_t[first_col_t].apply(
            lambda x: str(x).strip().replace(".0", "").isdigit()
            if pd.notna(x)
            else False
        )
    ].reset_index(drop=True)

    rename_t = {}
    for col in df_t.columns:
      cl = col.lower()
      if cl in ["no", "no.", "nomor urut"]:
        rename_t[col] = "No. Urut"
      elif "kode" in cl:
        rename_t[col] = "Kode Barang"
      elif cl in ["nomor", "no. register", "register"]:
        rename_t[col] = "Nomor Register"
    df_t = df_t.rename(columns=rename_t)

    if "No. Urut" in df_t.columns:
      df_t["No. Urut"] = range(1, len(df_t) + 1)

    for col in df_t.columns:
      df_t[col] = (
          df_t[col]
          .astype(str)
          .str.replace(r"\.0+$", "", regex=True)
          .replace("nan", "")
          .replace("None", "")
      )

    skpd_col_t = next(
        (c for c in df_t.columns if "skpd" in c.lower() or "dinas" in c.lower()),
        None,
    )
    df_t["SKPD_Nama"] = (
        df_t[skpd_col_t]
        .ffill()
        .fillna("DINAS / INSTANSI LAINNYA")
        .astype(str)
        .str.upper()
        .str.strip()
        if skpd_col_t
        else "DINAS / INSTANSI LAINNYA"
    )

    harga_col_t = next(
        (
            c
            for c in df_t.columns
            if "harga" in c.lower() or "rupiah" in c.lower()
        ),
        None,
    )
    df_t["Harga_Clean"] = (
        pd.to_numeric(
            df_t[harga_col_t].str.replace(r"[^\d.]", "", regex=True),
            errors="coerce",
        ).fillna(0)
        if harga_col_t
        else 0
    )
    df_tanah = df_t
  except Exception as e:
    print("Error KIB A:", e)

  return df_kendaraan, df_tanah, file_path


df_kendaraan, df_tanah, file_path = load_all_data()


def format_rupiah(nilai):
  formatted = f"{nilai:,.2f}"
  return "Rp " + formatted.replace(",", "X").replace(".", ",").replace("X", ".")


if "page" not in st.session_state:
  st.session_state.page = "menu"
if "module" not in st.session_state:
  st.session_state.module = ""
if "keyword" not in st.session_state:
  st.session_state.keyword = ""
if "title" not in st.session_state:
  st.session_state.title = ""

st.markdown(
    """
    <div class="main-header">
        <h2>🏛️ SIMANTAP - MANAJEMEN ASET DAERAH</h2>
        <p>Sistem Informasi Manajemen Aset & Inventaris Pemerintah Daerah</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if st.session_state.page == "menu":
  st.markdown(
      "<h4 style='text-align:center; color:#1f2937; margin-bottom:20px;"
      " font-weight:700;'>Silakan Pilih Modul KIB Aset Daerah</h4>",
      unsafe_allow_html=True,
  )

  tot_val_k = df_kendaraan["Harga_Clean"].sum() if not df_kendaraan.empty else 0
  len_k = len(df_kendaraan)
  tot_val_t = df_tanah["Harga_Clean"].sum() if not df_tanah.empty else 0
  len_t = len(df_tanah)

  col1, col2 = st.columns(2, gap="large")

  with col1:
    st.markdown(
        f"""
        <div class="kib-card-kendaraan">
            <h3>🚗 KIB B - Kendaraan Dinas</h3>
            <p>Total Unit: <b>{len_k:,} Data</b><br>
            Total Nilai: <b>{format_rupiah(tot_val_k)}</b></p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Kelola KIB B (Kendaraan Dinas)", key="btn_kib_b"):
      st.session_state.module = "kendaraan"
      st.session_state.page = "sub_menu_kendaraan"
      st.rerun()

  with col2:
    st.markdown(
        f"""
        <div class="kib-card-tanah">
            <h3>🗺️ KIB A - Tanah</h3>
            <p>Total Bidang: <b>{len_t:,} Data</b><br>
            Total Nilai: <b>{format_rupiah(tot_val_t)}</b></p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Kelola KIB A (Tanah)", key="btn_kib_a"):
      st.session_state.module = "tanah"
      st.session_state.keyword = ""
      st.session_state.title = "KIB A - Tanah"
      st.session_state.page = "table"
      st.rerun()

elif st.session_state.page == "sub_menu_kendaraan":
  if st.button("⬅️ Kembali ke Menu Utama"):
    st.session_state.page = "menu"
    st.rerun()

  st.markdown(
      "<h4 style='text-align:center; color:#1f2937; margin-bottom:20px;"
      " font-weight:700;'>Pilih Kategori Kendaraan Dinas</h4>",
      unsafe_allow_html=True,
  )


  def hitung_sub_k(kat):
    if df_kendaraan.empty:
      return format_rupiah(0), 0
    sub = (
        df_kendaraan
        if kat == ""
        else df_kendaraan[df_kendaraan["Kategori_Jenis"] == kat]
    )
    return format_rupiah(sub["Harga_Clean"].sum()), len(sub)


  p_semua, t_semua = hitung_sub_k("")
  p_motor, t_motor = hitung_sub_k("Sepeda Motor")
  p_mobil, t_mobil = hitung_sub_k("Mobil")
  p_pickup, t_pickup = hitung_sub_k("Pick Up")

  c1, c2 = st.columns(2, gap="large")
  with c1:
    st.markdown(
        f"""
        <div class="kib-card-kendaraan" style="background: linear-gradient(135deg, #3a6073 0%, #16222a 100%);">
            <h3>📦 Semua Kendaraan Dinas</h3>
            <p>Total Unit: <b>{t_semua:,} Data</b> | Total Nilai: <b>{p_semua}</b></p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Buka Semua Kendaraan", key="sub_semua"):
      st.session_state.keyword = ""
      st.session_state.title = "Semua Kendaraan Dinas"
      st.session_state.page = "table"
      st.rerun()

    st.markdown(
        f"""
        <div class="kib-card-kendaraan" style="background: linear-gradient(135deg, #1d976c 0%, #939b62 100%); margin-top:15px;">
            <h3>🚗 Mobil Dinas</h3>
            <p>Total Unit: <b>{t_mobil:,} Data</b> | Total Nilai: <b>{p_mobil}</b></p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Buka Data Mobil", key="sub_mobil"):
      st.session_state.keyword = "Mobil"
      st.session_state.title = "Mobil Dinas"
      st.session_state.page = "table"
      st.rerun()

  with c2:
    st.markdown(
        f"""
        <div class="kib-card-kendaraan" style="background: linear-gradient(135deg, #4b6cb7 0%, #182848 100%);">
            <h3>🏍️ Sepeda Motor Dinas</h3>
            <p>Total Unit: <b>{t_motor:,} Data</b> | Total Nilai: <b>{p_motor}</b></p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Buka Sepeda Motor", key="sub_motor"):
      st.session_state.keyword = "Sepeda Motor"
      st.session_state.title = "Sepeda Motor Dinas"
      st.session_state.page = "table"
      st.rerun()

    st.markdown(
        f"""
        <div class="kib-card-kendaraan" style="background: linear-gradient(135deg, #f2994a 0%, #f2c94c 100%); margin-top:15px;">
            <h3>🚙 Pick Up / Kendaraan Khusus</h3>
            <p>Total Unit: <b>{t_pickup:,} Data</b> | Total Nilai: <b>{p_pickup}</b></p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Buka Pick Up", key="sub_pickup"):
      st.session_state.keyword = "Pick Up"
      st.session_state.title = "Pick Up / Kendaraan Khusus"
      st.session_state.page = "table"
      st.rerun()

elif st.session_state.page == "table":
  col_back, col_ref, _ = st.columns([1.5, 1.5, 7])
  with col_back:
    if st.button("⬅️ Kembali"):
      if st.session_state.module == "kendaraan":
        st.session_state.page = "sub_menu_kendaraan"
      else:
        st.session_state.page = "menu"
      st.rerun()
  with col_ref:
    if st.button("🔄 Refresh Data"):
      st.cache_data.clear()
      st.rerun()

  st.markdown(f"### 📂 {st.session_state.title}")

  if st.session_state.module == "kendaraan":
    df_active = df_kendaraan
    keyword = st.session_state.keyword
    filtered_df = (
        df_active
        if keyword == ""
        else df_active[df_active["Kategori_Jenis"] == keyword]
    )
  else:
    df_active = df_tanah
    filtered_df = df_active

  pilih_skpd = "Semua SKPD"
  if not filtered_df.empty:
    skpd_list = ["Semua SKPD"] + sorted(
        list(filtered_df["SKPD_Nama"].dropna().unique())
    )
    pilih_skpd = st.selectbox(
        "🏢 Filter Berdasarkan Nama SKPD / Dinas:", skpd_list
    )
    if pilih_skpd != "Semua SKPD":
      filtered_df = filtered_df[filtered_df["SKPD_Nama"] == pilih_skpd]

  search_query = st.text_input("🔍 Cari data berdasarkan Nomor, Merk, dll...")
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
      f"Menampilkan {len(filtered_df)} baris data (Total Keseluruhan:"
      f" {len(df_active)} baris) | File: {os.path.basename(file_path)}"
  )

  columns_to_drop = ["Harga_Clean", "Kategori_Jenis", "SKPD_Nama"]
  display_df = filtered_df.drop(
      columns=[c for c in columns_to_drop if c in filtered_df.columns],
      errors="ignore",
  ).reset_index(drop=True)

  st.dataframe(display_df, use_container_width=True, height=400)

  st.markdown("---")
  st.markdown("#### 🔍 Preview Kartu Detail & Download Laporan")
  if not display_df.empty:
    selected_row_idx = st.selectbox(
        "Pilih Data untuk Lihat Detail Lengkap:",
        options=display_df.index,
        format_func=lambda x: (
            f"Baris {x+1}: {display_df.iloc[x].values[1] if len(display_df.columns) > 1 else display_df.iloc[x].values[0]}"
        ),
    )

    if selected_row_idx is not None:
      row_data = display_df.loc[selected_row_idx]
      columns_list = list(display_df.columns)
      values_list = [row_data[col] for col in columns_list]

      detail_df = pd.DataFrame({
          "Atribut / Kolom Database Excel": columns_list,
          "Keterangan / Isi Data": values_list,
      })

      st.dataframe(detail_df, use_container_width=True, height=400)

      col_e1, col_e2, col_e3 = st.columns(3)
      with col_e1:

        def create_excel(cols, vals):
          card_df = pd.DataFrame(
              {"Atribut / Kolom Database": cols, "Keterangan": vals}
          )
          output = BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            card_df.to_excel(writer, index=False, sheet_name="Detail Aset")
          output.seek(0)
          return output.getvalue()

        st.download_button(
            label="📊 Download Excel Detail",
            data=create_excel(columns_list, values_list),
            file_name=f"Detail_Aset_{selected_row_idx+1}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

      with col_e2:
        st.download_button(
            label="📄 Download CSV Detail",
            data=detail_df.to_csv(index=False).encode("utf-8"),
            file_name=f"Detail_Aset_{selected_row_idx+1}.csv",
            mime="text/csv",
        )

      with col_e3:

        def create_pdf(cols, vals):
          pdf = FPDF(orientation="P", unit="mm", format="A4")
          pdf.add_page()
          pdf.set_font("Arial", "B", 12)
          pdf.cell(0, 10, "KARTU INVENTARIS BARANG - DETAIL ASET", 0, 1, "C")
          pdf.ln(3)
          pdf.set_font("Arial", "B", 9)
          pdf.set_fill_color(44, 62, 80)
          pdf.set_text_color(255, 255, 255)
          pdf.cell(75, 7, "Atribut / Kolom Database", 1, 0, "C", True)
          pdf.cell(115, 7, "Keterangan / Isi Data", 1, 1, "C", True)
          pdf.set_font("Arial", "", 8.5)
          pdf.set_text_color(0, 0, 0)
          fill = False
          for col, val in zip(cols, vals):
            pdf.set_fill_color(241, 245, 249) if fill else pdf.set_fill_color(
                255, 255, 255
            )
            pdf.cell(75, 6, str(col or ""), 1, 0, "L", True)
            pdf.cell(115, 6, str(val or ""), 1, 1, "L", True)
            fill = not fill
          output_pdf = pdf.output(dest="S")
          return (
              bytes(output_pdf)
              if isinstance(output_pdf, (bytes, bytearray))
              else output_pdf.encode("latin1")
          )

        st.download_button(
            label="📑 Download PDF Detail",
            data=create_pdf(columns_list, values_list),
            file_name=f"Detail_Aset_{selected_row_idx+1}.pdf",
            mime="application/pdf",
        )