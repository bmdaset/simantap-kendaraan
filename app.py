from io import BytesIO
import glob
import os
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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
def load_data():
  file_path = "REKAP KENDARAAN TA. 2026.YP.xlsx"
  if not os.path.exists(file_path):
    all_excel = glob.glob(".xlsx") + glob.glob(".xls")
    if all_excel:
      file_path = all_excel[0]
    else:
      return pd.DataFrame(), None

  try:
    df_raw = pd.read_excel(
        file_path, sheet_name="KENDARAAN DINAS", header=None
    )

    # Definisi kolom KIB B secara akurat dan lengkap
    columns_kib_b = [
        "No. Urut",
        "Kode Barang",
        "Jenis Barang",
        "Nomor Register",
        "Merk / Type",
        "Ukuran / CC",
        "Bahan",
        "Tahun Pembelian",
        "Nomor Pabrik",
        "Nomor Rangka",
        "Nomor Mesin",
        "Nomor Polisi",
        "Nomor BPKB",
        "Asal Usul",
        "Jumlah / Satuan",
        "Harga",
        "Keterangan",
        "Pengguna / Pemakai",
        "SKPD",
    ]

    # Data aktual dimulai dari baris ke-16 (index 15)
    df = df_raw.iloc[15:].copy()

    if df.shape[1] <= len(columns_kib_b):
      df.columns = columns_kib_b[: df.shape[1]]
    else:
      extra_cols = [
          f"Kolom_{i}" for i in range(len(columns_kib_b), df.shape[1])
      ]
      df.columns = columns_kib_b + extra_cols

    df = df.reset_index(drop=True)

    # Validasi baris berdasarkan No. Urut angka valid > 0
    def is_valid_row(val):
      try:
        return int(float(val)) > 0
      except:
        return False

    if "No. Urut" in df.columns:
      df = df[df["No. Urut"].apply(is_valid_row)].reset_index(drop=True)
      df["No. Urut"] = range(1, len(df) + 1)

    # Pembersihan string & spasi ekstra di seluruh dataframe
    for col in df.columns:
      df[col] = (
          df[col]
          .astype(str)
          .str.replace(r"\.0+$", "", regex=True)
          .replace("nan", "")
          .replace("None", "")
          .str.strip()
      )

    # Penanganan SKPD
    skpd_col = next(
        (c for c in df.columns if "skpd" in c.lower() or "dinas" in c.lower()),
        None,
    )
    if skpd_col:
      df["SKPD_Nama"] = (
          df[skpd_col]
          .ffill()
          .fillna("DINAS / INSTANSI LAINNYA")
          .astype(str)
          .str.upper()
          .str.strip()
      )
    else:
      df["SKPD_Nama"] = "DINAS / INSTANSI LAINNYA"

    # Pembersihan nilai Harga
    harga_col = next(
        (c for c in df.columns if "harga" in c.lower() or "rupiah" in c.lower()),
        None,
    )
    if harga_col:
      df["Harga_Clean"] = pd.to_numeric(
          df[harga_col].str.replace(r"[^\d.]", "", regex=True), errors="coerce"
      ).fillna(0)
    else:
      df["Harga_Clean"] = 0

    # Deteksi Kategori Kendaraan
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
  except Exception as e:
    print("Error:", e)
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
      f" Pilih Kategori Aset Kendaraan (Total Database: {len(df):,} Unit)</h4>",
      unsafe_allow_html=True,
  )

  p_semua, t_semua = hitung_total("")
  p_motor, t_motor = hitung_total("Sepeda Motor")
  p_mobil, t_mobil = hitung_total("Mobil")
  p_pickup, t_pickup = hitung_total("Pick Up")

  col1, col2 = st.columns(2, gap="medium")

  with col1:
    with st.container(border=True):
      st.markdown("### 📦 Semua Kendaraan")
      st.write(
          f"Total Unit: {t_semua:,} Data\n\nTotal Nilai Aset:"
          f" {p_semua}"
      )
      if st.button("Kelola Semua Kendaraan", key="btn_semua"):
        st.session_state.keyword = ""
        st.session_state.title = "Semua Kendaraan Dinas"
        st.session_state.page = "table"
        st.rerun()

    st.write("")
    with st.container(border=True):
      st.markdown("### 🚗 Mobil")
      st.write(
          f"Total Unit: {t_mobil:,} Data\n\nTotal Nilai Aset:"
          f" {p_mobil}"
      )
      if st.button("Kelola Data Mobil", key="btn_mobil"):
        st.session_state.keyword = "Mobil"
        st.session_state.title = "Mobil"
        st.session_state.page = "table"
        st.rerun()

  with col2:
    with st.container(border=True):
      st.markdown("### 🏍️ Sepeda Motor")
      st.write(
          f"Total Unit: {t_motor:,} Data\n\nTotal Nilai Aset:"
          f" {p_motor}"
      )
      if st.button("Kelola Sepeda Motor", key="btn_motor"):
        st.session_state.keyword = "Sepeda Motor"
        st.session_state.title = "Sepeda Motor"
        st.session_state.page = "table"
        st.rerun()

    st.write("")
    with st.container(border=True):
      st.markdown("### 🚙 Pick Up")
      st.write(
          f"Total Unit: {t_pickup:,} Data\n\nTotal Nilai Aset:"
          f" {p_pickup}"
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
    if st.button("🔄 Refresh Data (Clear Cache)"):
      st.cache_data.clear()
      st.rerun()

  st.markdown(f"### 📂 {st.session_state.title}")

  keyword = st.session_state.keyword
  if keyword == "":
    filtered_df = df
  else:
    filtered_df = df[df["Kategori_Jenis"] == keyword]

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

  search_query = st.text_input(
      "🔍 Cari data (No. Polisi, No. Rangka, No. Mesin, No. BPKB, Merk,"
      " dll)..."
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
      f"Menampilkan {len(filtered_df)} baris data (Total Keseluruhan"
      f" Database: {len(df)} baris) | File:"
      f" {os.path.basename(file_path) if file_path else 'Tidak ada'}"
  )

  columns_to_drop = [
      "Harga_Clean",
      "Kategori_Jenis",
      "SKPD_Nama",
  ]
  display_df = filtered_df.drop(
      columns=[c for c in columns_to_drop if c in filtered_df.columns],
      errors="ignore",
  ).reset_index(drop=True)

  st.dataframe(display_df, use_container_width=True, height=400)

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
  st.markdown("#### 🔍 Preview Kartu Detail Bergaris & Download")
  if not display_df.empty:
    selected_row_idx = st.selectbox(
        "Pilih Kendaraan untuk Lihat Detail Lengkap:",
        options=display_df.index,
        format_func=lambda x: f"Baris {x+1}: No. Polisi: {display_df.loc[x, 'Nomor Polisi'] if 'Nomor Polisi' in display_df.columns else display_df.iloc[x].values[0]} | Merk: {display_df.loc[x, 'Merk / Type'] if 'Merk / Type' in display_df.columns else ''}",
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

      st.markdown(
          "<p style='font-weight:600; color:#1e3c72;'>Preview Tabel"
          " Bergaris:</p>",
          unsafe_allow_html=True,
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
            card_df.to_excel(writer, index=False, sheet_name="Detail Kendaraan")

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

      with col_e3:

        def create_pdf_detail(cols, vals):
          pdf = FPDF(orientation="P", unit="mm", format="A4")
          pdf.add_page()

          pdf.set_font("Arial", "B", 14)
          pdf.set_text_color(30, 60, 114)
          pdf.cell(
              0, 10, "DETAIL INFORMASI ASET KENDARAAN DINAS", 0, 1, "C"
          )
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

            col_str = str(col or "")
            val_str = str(val or "")

            pdf.cell(70, 6, col_str, 1, 0, "L", True)
            pdf.cell(120, 6, val_str, 1, 1, "L", True)
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
            file_name=f"Detail_Kendaraan_{selected_row_idx+1}.pdf",
            mime="application/pdf",
        )