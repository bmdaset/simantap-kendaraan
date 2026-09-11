from io import BytesIO
import glob
import os
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


@st.cache_data
def load_all_data():
  file_path = "REKAP KENDARAAN TA. 2026.YP.xlsx"
  if not os.path.exists(file_path):
    all_excel = glob.glob(".xlsx") + glob.glob(".xls")
    if all_excel:
      file_path = all_excel[0]
    else:
      return pd.DataFrame(), pd.DataFrame(), None

  # 1. Load KIB B (KENDARAAN DINAS)
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
    if df_k.shape[1] > 18:
      df_k = df_k.iloc[:, :18]
    df_k = df_k.dropna(how="all").reset_index(drop=True)

    first_col = df_k.columns[0]

    def is_valid_row(val):
      try:
        return int(float(val)) > 0
      except:
        return False

    df_k = df_k[df_k[first_col].apply(is_valid_row)].reset_index(drop=True)

    rename_k = {}
    for col in df_k.columns:
      cl = col.lower()
      if cl in ["no", "no.", "nomor urut"]:
        rename_k[col] = "No. Urut"
      elif "kode" in cl:
        rename_k[col] = "Kode Barang"
      elif cl in ["nomor", "no. register", "register"]:
        rename_k[col] = "Nomor Register"
    df_k = df_k.rename(columns=rename_k)
    if "No. Urut" in df_k.columns:
      df_k["No. Urut"] = range(1, len(df_k) + 1)

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

  # 2. Load KIB A (KIB A TANAH) - Robust Header & Column Detection
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
    # Flatten MultiIndex columns if any
    df_t.columns = [
        " ".join([str(c) for c in col if "unnamed" not in str(c).lower()]).strip()
        for col in df_t.columns
    ]
    # Fallback if flat columns are empty/messy, load with single header
    if len(df_t.columns) == 0 or all(c == "" for c in df_t.columns):
      df_t = pd.read_excel(
          file_path, sheet_name="KIB A TANAH", header=h_idx_t + 1
      )

    df_t = df_t.loc[:, ~df_t.columns.astype(str).str.contains("^Unnamed")]
    df_t.columns = [str(c).strip() for c in df_t.columns]
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
      "<h4 style='text-align:center; color:#34495e; margin-bottom:25px;'>Silakan"
      " Pilih Modul KIB Aset Daerah</h4>",
      unsafe_allow_html=True,
  )

  tot_val_k = df_kendaraan["Harga_Clean"].sum() if not df_kendaraan.empty else 0
  len_k = len(df_kendaraan)

  tot_val_t = df_tanah["Harga_Clean"].sum() if not df_tanah.empty else 0
  len_t = len(df_tanah)

  col1, col2 = st.columns(2, gap="medium")

  with col1:
    with st.container(border=True):
      st.markdown("### 🚗 *KIB B - Kendaraan Dinas*")
      st.write(
          f"Total Unit: *{len_k:,} Data*\n\nTotal Nilai Aset:"
          f" *{format_rupiah(tot_val_k)}*"
      )
      if st.button("Kelola KIB B (Kendaraan)", key="btn_kib_b"):
        st.session_state.module = "kendaraan"
        st.session_state.page = "sub_menu_kendaraan"
        st.rerun()

  with col2:
    with st.container(border=True):
      st.markdown("### 🗺️ *KIB A - Tanah*")
      st.write(
          f"Total Bidang: *{len_t:,} Data*\n\nTotal Nilai Aset:"
          f" *{format_rupiah(tot_val_t)}*"
      )
      if st.button("Kelola KIB A (Tanah)", key="btn_kib_a"):
        st.session_state.module = "tanah"
        st.session_state.keyword = ""
        st.session_state.title = "KIB A - Tanah"
        st.session_state.page = "table"
        st.rerun()

elif st.session_state.page == "sub_menu_kendaraan":
  col_back, _ = st.columns([1.5, 8.5])
  with col_back:
    if st.button("⬅️ Kembali ke Menu Utama"):
      st.session_state.page = "menu"
      st.rerun()

  st.markdown(
      "<h4 style='text-align:center; color:#34495e; margin-bottom:25px;'>Pilih"
      " Kategori Kendaraan Dinas</h4>",
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

  c1, c2 = st.columns(2, gap="medium")
  with c1:
    with st.container(border=True):
      st.markdown("### 📦 *Semua Kendaraan*")
      st.write(
          f"Total Unit: *{t_semua:,} Data\n\nTotal Nilai: *{p_semua}**"
      )
      if st.button("Buka Semua Kendaraan", key="sub_semua"):
        st.session_state.keyword = ""
        st.session_state.title = "Semua Kendaraan Dinas"
        st.session_state.page = "table"
        st.rerun()
    st.write("")
    with st.container(border=True):
      st.markdown("### 🚗 *Mobil*")
      st.write(
          f"Total Unit: *{t_mobil:,} Data\n\nTotal Nilai: *{p_mobil}**"
      )
      if st.button("Buka Data Mobil", key="sub_mobil"):
        st.session_state.keyword = "Mobil"
        st.session_state.title = "Mobil"
        st.session_state.page = "table"
        st.rerun()
  with c2:
    with st.container(border=True):
      st.markdown("### 🏍️ *Sepeda Motor*")
      st.write(
          f"Total Unit: *{t_motor:,} Data\n\nTotal Nilai: *{p_motor}**"
      )
      if st.button("Buka Sepeda Motor", key="sub_motor"):
        st.session_state.keyword = "Sepeda Motor"
        st.session_state.title = "Sepeda Motor"
        st.session_state.page = "table"
        st.rerun()
    st.write("")
    with st.container(border=True):
      st.markdown("### 🚙 *Pick Up*")
      st.write(
          f"Total Unit: *{t_pickup:,} Data\n\nTotal Nilai: *{p_pickup}**"
      )
      if st.button("Buka Pick Up", key="sub_pickup"):
        st.session_state.keyword = "Pick Up"
        st.session_state.title = "Pick Up"
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

  search_query = st.text_input("🔍 Cari data...")
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
  st.markdown("#### 🔍 Preview Kartu Detail Bergaris & Download")
  if not display_df.empty:
    selected_row_idx = st.selectbox(
        "Pilih Data untuk Lihat Detail Lengkap:",
        options=display_df.index,
        format_func=lambda x: f"Baris {x+1}: {display_df.iloc[x].values[0]}",
    )

    if selected_row_idx is not None:
      row_data = display_df.loc[selected_row_idx]
      columns_list = list(display_df.columns)
      values_list = [row_data[col] for col in columns_list]

      detail_df = pd.DataFrame({
          "Atribut / Kolom Data": columns_list,
          "Keterangan / Isi Data": values_list,
      })

      styled_preview = (
          detail_df.style.set_table_styles([
              {
                  "selector": "th",
                  "props": [
                      ("background-color", "#1e3c72"),
                      ("color", "white"),
                      ("font-weight", "bold"),
                      ("border", "1px solid black"),
                      ("text-align", "center"),
                  ],
              },
              {
                  "selector": "td",
                  "props": [
                      ("border", "1px solid #b0b0b0"),
                      ("padding", "6px 10px"),
                  ],
              },
          ])
          .set_properties(**{"text-align": "left"})
          .hide(axis="index")
      )

      st.dataframe(styled_preview, use_container_width=True, height=450)

      col_e1, col_e2, col_e3 = st.columns(3)
      with col_e1:

        def create_styled_vertical_excel(cols, vals):
          card_df = pd.DataFrame(
              {"Atribut / Kolom Data": cols, "Keterangan / Isi Data": vals}
          )
          output = BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            card_df.to_excel(writer, index=False, sheet_name="Detail Aset")
          output.seek(0)
          wb = openpyxl.load_workbook(output)
          ws = wb.active
          thin_border = Border(
              left=Side(style="thin", color="888888"),
              right=Side(style="thin", color="888888"),
              top=Side(style="thin", color="888888"),
              bottom=Side(style="thin", color="888888"),
          )
          header_fill = PatternFill(
              start_color="1E3C72", end_color="1E3C72", fill_type="solid"
          )
          header_font = Font(
              name="Calibri", size=11, bold=True, color="FFFFFF"
          )
          fill_even = PatternFill(
              start_color="F9FAFB", end_color="F9FAFB", fill_type="solid"
          )

          for col_idx in range(1, 3):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            cell.border = Border(
                left=Side(style="thin", color="000000"),
                right=Side(style="thin", color="000000"),
                top=Side(style="thin", color="000000"),
                bottom=Side(style="thin", color="000000"),
            )

          for row_idx in range(2, ws.max_row + 1):
            is_even = row_idx % 2 == 0
            for col_idx in range(1, 3):
              cell = ws.cell(row=row_idx, column=col_idx)
              cell.border = thin_border
              cell.alignment = Alignment(vertical="center", wrap_text=True)
              if is_even:
                cell.fill = fill_even

          ws.column_dimensions["A"].width = 30
          ws.column_dimensions["B"].width = 50
          final_output = BytesIO()
          wb.save(final_output)
          return final_output.getvalue()

        excel_data = create_styled_vertical_excel(columns_list, values_list)
        st.download_button(
            label="📊 Download Excel Bergaris",
            data=excel_data,
            file_name=f"Detail_Aset_{selected_row_idx+1}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

      with col_e2:
        csv_data = detail_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📄 Download CSV Detail",
            data=csv_data,
            file_name=f"Detail_Aset_{selected_row_idx+1}.csv",
            mime="text/csv",
        )

      with col_e3:

        def create_pdf_detail(cols, vals):
          pdf = FPDF(orientation="P", unit="mm", format="A4")
          pdf.add_page()
          pdf.set_font("Arial", "B", 14)
          pdf.set_text_color(30, 60, 114)
          pdf.cell(0, 10, "DETAIL INFORMASI ASET DAERAH", 0, 1, "C")
          pdf.ln(4)
          pdf.set_font("Arial", "B", 10)
          pdf.set_fill_color(30, 60, 114)
          pdf.set_text_color(255, 255, 255)
          pdf.cell(70, 7, "Atribut / Kolom Data", 1, 0, "C", True)
          pdf.cell(120, 7, "Keterangan / Isi Data", 1, 1, "C", True)
          pdf.set_font("Arial", "", 9)
          pdf.set_text_color(0, 0, 0)
          fill = False
          for col, val in zip(cols, vals):
            if fill:
              pdf.set_fill_color(240, 244, 248)
            else:
              pdf.set_fill_color(255, 255, 255)
            pdf.cell(70, 6, str(col or ""), 1, 0, "L", True)
            pdf.cell(120, 6, str(val or ""), 1, 1, "L", True)
            fill = not fill
          output_pdf = pdf.output(dest="S")
          if isinstance(output_pdf, (bytes, bytearray)):
            return bytes(output_pdf)
          else:
            return output_pdf.encode("latin1")

        pdf_bytes = create_pdf_detail(columns_list, values_list)
        st.download_button(
            label="📑 Download PDF Bergaris",
            data=pdf_bytes,
            file_name=f"Detail_Aset_{selected_row_idx+1}.pdf",
            mime="application/pdf",
        )