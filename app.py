from io import BytesIO
import glob
import os
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="SIMANTAP - Aset Pemerintah Daerah",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    /* Global Background Tema BMD */
    .stApp {
        background: linear-gradient(135deg, #eef2f5 0%, #e2e8f0 100%);
    }
    .main-header {
        background: linear-gradient(135deg, #1b365d 0%, #3182ce 100%);
        padding: 35px;
        border-radius: 14px;
        color: white;
        text-align: center;
        margin-bottom: 30px;
        box-shadow: 0 8px 25px rgba(27,54,93,0.3);
        border: 1px solid rgba(255,255,255,0.2);
    }
    .main-header h2 {
        margin: 0;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-weight: 700;
        letter-spacing: 0.8px;
        color: #ffffff;
    }
    .main-header p {
        margin: 10px 0 0 0;
        font-size: 15px;
        color: #edf2f7;
    }
    /* Kartu Menu Kendaraan - Nuansa Cerah & Elegan */
    .card-menu-kendaraan {
        background: linear-gradient(135deg, #2b6cb0 0%, #4299e1 100%);
        padding: 28px;
        border-radius: 14px;
        color: white;
        box-shadow: 0 10px 25px rgba(43,108,176,0.35);
        margin-bottom: 15px;
        border-left: 6px solid #bee3f8;
    }
    /* Kartu Menu KIB A Tanah - Nuansa Hijau Tanah Cerah & Elegan */
    .card-menu-tanah {
        background: linear-gradient(135deg, #276749 0%, #38a169 100%);
        padding: 28px;
        border-radius: 14px;
        color: white;
        box-shadow: 0 10px 25px rgba(39,103,73,0.35);
        margin-bottom: 15px;
        border-left: 6px solid #c6f6d5;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: 600;
        padding: 12px;
        font-size: 15px;
        transition: all 0.3s ease;
        border: none;
        background-color: #ffffff;
        color: #1a365d;
        box-shadow: 0 4px 6px rgba(0,0,0,0.08);
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 15px rgba(0,0,0,0.18);
        background-color: #f7fafc;
        color: #2b6cb0;
    }
    </style>
""",
    unsafe_allow_html=True,
)


@st.cache_data
def load_kendaraan_data():
  file_path = "REKAP KENDARAAN TA. 2026.YP.xlsx"
  if not os.path.exists(file_path):
    all_excel = glob.glob("*.xlsx") + glob.glob("*.xls")
    if all_excel:
      file_path = all_excel[0]
    else:
      return pd.DataFrame()

  try:
    df_raw = pd.read_excel(
        file_path, sheet_name="KENDARAAN DINAS", header=None
    )
    header_row_idx = 13
    for idx, row in df_raw.head(25).iterrows():
      row_str = " ".join([str(val).lower() for val in row.values])
      if "jenis barang" in row_str or "merk" in row_str:
        header_row_idx = idx
        break

    h1 = (
        df_raw.iloc[header_row_idx]
        .fillna("")
        .astype(str)
        .str.strip()
        .tolist()
    )
    if header_row_idx + 1 < len(df_raw):
      h2 = (
          df_raw.iloc[header_row_idx + 1]
          .fillna("")
          .astype(str)
          .str.strip()
          .tolist()
      )
      combined_cols = []
      last_main = ""
      for m, s in zip(h1, h2):
        if m != "" and not m.startswith("Unnamed"):
          last_main = m
        if s != "" and not s.startswith("Unnamed"):
          if last_main and last_main.lower() not in s.lower():
            combined_cols.append(f"{last_main} - {s}")
          else:
            combined_cols.append(s)
        else:
          combined_cols.append(last_main if last_main else "Kolom")
      final_cols = combined_cols
      data_start_idx = header_row_idx + 2
    else:
      final_cols = h1
      data_start_idx = header_row_idx + 1

    transformed_cols = []
    for c in final_cols:
      c_lower = c.lower()
      if "pembelian" in c_lower:
        transformed_cols.append("Tahun Pembuatan")
      elif c_lower.startswith("tahun"):
        transformed_cols.append(
            c.replace("Tahun", "Nomor").replace("tahun", "Nomor")
        )
      else:
        transformed_cols.append(c)

    seen = {}
    unique_cols = []
    for c in transformed_cols:
      if c in seen:
        seen[c] += 1
        unique_cols.append(f"{c}_{seen[c]}")
      else:
        seen[c] = 0
        unique_cols.append(c)

    df = df_raw.iloc[data_start_idx:].copy()
    df.columns = unique_cols[: df.shape[1]]
    if df.shape[1] > len(unique_cols):
      for i in range(len(unique_cols), df.shape[1]):
        df.rename(columns={df.columns[i]: f"Kolom_{i}"}, inplace=True)

    drop_indices = [14, 19, 20]
    valid_drop_indices = [i for i in drop_indices if i < len(df.columns)]
    df = df.drop(df.columns[valid_drop_indices], axis=1).reset_index(drop=True)
    df = df.dropna(how="all").reset_index(drop=True)
    df = df[df.iloc[:, 0].notna()].reset_index(drop=True)

    if not df.empty:
      first_col_name = df.columns[0]
      df = df[
          ~df[first_col_name]
          .astype(str)
          .str.lower()
          .str.contains("jumlah|total|n o", na=False)
      ].reset_index(drop=True)

    df["SKPD_Nama"] = "DINAS / INSTANSI LAINNYA"
    df["Harga_Clean"] = 0.0
    df["Kategori_Jenis"] = "Mobil"

    if not df.empty:
      skpd_col = next(
          (
              c
              for c in df.columns
              if c == "SKPD"
              or "skpd" in c.lower()
              or "dinas" in c.lower()
              or "unit" in c.lower()
          ),
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
        df[skpd_col] = df["SKPD_Nama"]

      harga_col = next(
          (
              c
              for c in df.columns
              if "harga" in c.lower()
              or "rupiah" in c.lower()
              or "nilai" in c.lower()
          ),
          None,
      )
      if harga_col:
        calculated_harga = (
            pd.to_numeric(
                df[harga_col]
                .astype(str)
                .str.replace(r"[^\d.]", "", regex=True),
                errors="coerce",
            ).fillna(0)
            * 1000
        )
        df["Harga_Clean"] = calculated_harga
        df[harga_col] = calculated_harga

      def deteksi_kategori(row):
        combined = " ".join(
            [str(val) for val in row.values if pd.notna(val)]
        ).lower()
        if any(
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
        elif any(k in combined for k in ["pick up", "pickup", "bak terbuka"]):
          return "Pick Up"
        elif any(
            k in combined
            for k in [
                "excavator",
                "bulldozer",
                "loader",
                "grader",
                "crane",
                "forklift",
                "tractor",
                "traktor",
                "molen",
                "roller",
                "compactor",
                "backhoe",
                "alat berat",
            ]
        ):
          return "Alat Berat"
        else:
          return "Mobil"

      df["Kategori_Jenis"] = df.apply(deteksi_kategori, axis=1)

    return df
  except Exception as e:
    print("Error Load Kendaraan:", e)
    return pd.DataFrame()


@st.cache_data
def load_kiba_data():
  file_path = "REKAP KENDARAAN TA. 2026.YP.xlsx"
  if not os.path.exists(file_path):
    all_excel = glob.glob("*.xlsx") + glob.glob("*.xls")
    if all_excel:
      file_path = all_excel[0]
    else:
      return pd.DataFrame()

  try:
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    target_sheet = next(
        (
            s
            for s in sheet_names
            if "tanah" in s.lower()
            or "kib a" in s.lower()
            or "kiba" in s.lower()
        ),
        sheet_names[1] if len(sheet_names) > 1 else sheet_names[0],
    )

    df_raw = pd.read_excel(file_path, sheet_name=target_sheet, header=None)

    header_row_idx = 0
    for idx, row in df_raw.head(25).iterrows():
      row_str = " ".join([str(val).lower() for val in row.values])
      if (
          "luas" in row_str
          or "alamat" in row_str
          or "letak" in row_str
          or "hak" in row_str
      ):
        header_row_idx = idx
        break

    h1 = (
        df_raw.iloc[header_row_idx]
        .fillna("")
        .astype(str)
        .str.strip()
        .tolist()
    )
    if header_row_idx + 1 < len(df_raw):
      h2 = (
          df_raw.iloc[header_row_idx + 1]
          .fillna("")
          .astype(str)
          .str.strip()
          .tolist()
      )
      row_str_h2 = " ".join([str(v).lower() for v in h2])
      if (
          "m2" in row_str_h2
          or "status" in row_str_h2
          or "tanggal" in row_str_h2
          or "nomor" in row_str_h2
          or "sertifikat" in row_str_h2
      ):
        combined_cols = []
        last_main = ""
        for m, s in zip(h1, h2):
          if m != "" and not m.startswith("Unnamed"):
            last_main = m
          if s != "" and not s.startswith("Unnamed"):
            if last_main and last_main.lower() not in s.lower():
              combined_cols.append(f"{last_main} - {s}")
            else:
              combined_cols.append(s)
          else:
            combined_cols.append(last_main if last_main else "Kolom")
        final_cols = combined_cols
        data_start_idx = header_row_idx + 2
      else:
        final_cols = h1
        data_start_idx = header_row_idx + 1
    else:
      final_cols = h1
      data_start_idx = header_row_idx + 1

    cleaned_cols = []
    for idx_col, c in enumerate(final_cols):
      c_str = str(c).strip()

      if idx_col == 9 or c_str.lower() == "letak/":
        c_str = "Tanggal Sertifikat"
      else:
        if "alamat" not in c_str.lower():
          c_str = (
              c_str.replace("Letak/ - ", "")
              .replace("Letak / - ", "")
              .replace("Letak/", "")
              .replace("letak/", "")
              .strip()
          )

      if (
          not c_str
          or c_str.lower() == "none"
          or c_str.startswith("Unnamed")
      ):
        c_str = f"Kolom_{idx_col}"

      cleaned_cols.append(c_str)

    seen = {}
    unique_cols = []
    for c in cleaned_cols:
      if c in seen:
        seen[c] += 1
        unique_cols.append(f"{c}_{seen[c]}")
      else:
        seen[c] = 0
        unique_cols.append(c)

    df = df_raw.iloc[data_start_idx:].copy()
    df.columns = unique_cols[: df.shape[1]]
    if df.shape[1] > len(unique_cols):
      for i in range(len(unique_cols), df.shape[1]):
        df.rename(columns={df.columns[i]: f"Kolom_{i}"}, inplace=True)

    df = df.dropna(how="all").reset_index(drop=True)
    df = df[df.iloc[:, 0].notna()].reset_index(drop=True)

    if not df.empty:
      first_col_name = df.columns[0]
      df = df[
          ~df[first_col_name]
          .astype(str)
          .str.lower()
          .str.contains("jumlah|total|n o", na=False)
      ].reset_index(drop=True)

    df["SKPD_Nama"] = "DINAS / INSTANSI LAINNYA"
    df["Harga_Clean"] = 0.0

    if not df.empty:
      skpd_col = next(
          (
              c
              for c in df.columns
              if c == "SKPD"
              or "skpd" in c.lower()
              or "dinas" in c.lower()
              or "unit" in c.lower()
          ),
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
        df[skpd_col] = df["SKPD_Nama"]

      harga_col = next(
          (
              c
              for c in df.columns
              if "harga" in c.lower()
              or "rupiah" in c.lower()
              or "nilai" in c.lower()
          ),
          None,
      )
      if harga_col:
        raw_harga = pd.to_numeric(
            df[harga_col].astype(str).str.replace(r"[^\d.]", "", regex=True),
            errors="coerce",
        ).fillna(0)
        if raw_harga.mean() > 0 and raw_harga.mean() < 100000:
          calculated_harga = raw_harga * 1000
        else:
          calculated_harga = raw_harga

        df["Harga_Clean"] = calculated_harga
        df[harga_col] = calculated_harga

    return df
  except Exception as e:
    print("Error Load KIB A Tanah:", e)
    return pd.DataFrame()


df_kendaraan = load_kendaraan_data()
df_tanah = load_kiba_data()


if "page" not in st.session_state:
  st.session_state.page = "menu"
if "module" not in st.session_state:
  st.session_state.module = "kendaraan"
if "keyword" not in st.session_state:
  st.session_state.keyword = ""
if "title" not in st.session_state:
  st.session_state.title = ""

if st.session_state.page == "table":
  if st.session_state.module == "kendaraan":
    st.markdown(
        """
        <style>
        .stApp {
            background: linear-gradient(135deg, #ebf8ff 0%, #bee3f8 100%);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
  else:
    st.markdown(
        """
        <style>
        .stApp {
            background: linear-gradient(135deg, #f0fff4 0%, #c6f6d5 100%);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

st.markdown(
    """
    <div class="main-header">
        <h2>🏛️ SIMANTAP - ASET PEMERINTAH DAERAH</h2>
        <p>Sistem Informasi Manajemen Aset & Inventaris Kendaraan dan Tanah (BMD)</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if st.session_state.page == "menu":
  st.markdown(
      "<h4 style='text-align:center; color:#2d3748; margin-bottom:25px;"
      " font-weight:600;'>Silakan Pilih Jenis Modul Aset Daerah</h4>",
      unsafe_allow_html=True,
  )

  col1, col2 = st.columns(2, gap="medium")

  with col1:
    st.markdown(
        f"""
        <div class="card-menu-kendaraan">
            <h3 style="margin-top:0; color:#fff;">🚗 Kendaraan Dinas</h3>
            <p style="color:#f7fafc; font-size:14px;"><b>Total Unit Terdata:</b> {len(df_kendaraan):,} Data<br>Kelola data mobil, motor, pick-up, & alat berat dinas.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Kelola Kendaraan Dinas", key="btn_mod_kendaraan"):
      st.session_state.module = "kendaraan"
      st.session_state.keyword = ""
      st.session_state.title = "Semua Kendaraan Dinas"
      st.session_state.page = "table"
      st.rerun()

  with col2:
    st.markdown(
        f"""
        <div class="card-menu-tanah">
            <h3 style="margin-top:0; color:#fff;">🌍 KIB A - Tanah</h3>
            <p style="color:#f7fafc; font-size:14px;"><b>Total Bidang Terdata:</b> {len(df_tanah):,} Data<br>Kelola data aset tanah dan alamat lokasi.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Kelola KIB A Tanah", key="btn_mod_tanah"):
      st.session_state.module = "tanah"
      st.session_state.keyword = ""
      st.session_state.title = "KIB A - Tanah"
      st.session_state.page = "table"
      st.rerun()

elif st.session_state.page == "table":
  col_back, col_ref, col_title = st.columns([1.5, 1.5, 5])
  with col_back:
    if st.button("⬅️ Kembali ke Menu Utama"):
      st.session_state.page = "menu"
      st.rerun()
  with col_ref:
    if st.button("🔄 Refresh Data"):
      st.cache_data.clear()
      st.rerun()

  module_accent = (
      "#2b6cb0" if st.session_state.module == "kendaraan" else "#276749"
  )
  st.markdown(
      f"<h3 style='color:{module_accent}; font-weight:700;'>📂"
      f" {st.session_state.title}</h3>",
      unsafe_allow_html=True,
  )

  if st.session_state.module == "kendaraan":
    active_df = df_kendaraan
    keyword = st.session_state.keyword
    if keyword != "":
      active_df = active_df[active_df["Kategori_Jenis"] == keyword]
  else:
    active_df = df_tanah

  pilih_skpd = "Semua SKPD"
  if not active_df.empty:
    skpd_list = ["Semua SKPD"] + sorted(
        list(active_df["SKPD_Nama"].dropna().unique())
    )
    pilih_skpd = st.selectbox(
        "🏢 Filter Berdasarkan Nama SKPD / Dinas:", skpd_list
    )
    if pilih_skpd != "Semua SKPD":
      active_df = active_df[active_df["SKPD_Nama"] == pilih_skpd]

  if st.session_state.module == "tanah":
    search_query = st.text_input(
        "🔍 Cari KIB A Tanah (Berdasarkan Alamat, Keterangan, No. Sertifikat,"
        " Nama Barang, dll)..."
    )
  else:
    search_query = st.text_input(
        "🔍 Cari data Kendaraan (No. Polisi, Merk, No. Rangka, dll)..."
    )

  if search_query:
    mask_search = (
        active_df.astype(str)
        .apply(
            lambda col: col.str.lower().str.contains(
                search_query.lower(), na=False
            )
        )
        .any(axis=1)
    )
    active_df = active_df[mask_search]

  st.info(f"Menampilkan {len(active_df):,} baris data aset")

  columns_to_drop = ["Harga_Clean", "Kategori_Jenis", "SKPD_Nama"]
  display_df = active_df.drop(
      columns=[c for c in columns_to_drop if c in active_df.columns],
      errors="ignore",
  ).reset_index(drop=True)

  st.dataframe(display_df, use_container_width=True, height=400)

  st.markdown("---")
  if st.session_state.module == "kendaraan":
    rekap_title = (
        f"📊 Rekapitulasi Kategori Kendaraan: {pilih_skpd}"
        if pilih_skpd != "Semua SKPD"
        else "📊 Rekapitulasi Keseluruhan Kategori Kendaraan per SKPD"
    )
    st.markdown(f"#### {rekap_title}")

    if not active_df.empty:
      summary_pivot = (
          active_df.pivot_table(
              index="SKPD_Nama",
              columns="Kategori_Jenis",
              values=active_df.columns[0],
              aggfunc="count",
              fill_value=0,
          )
          .reset_index()
      )
      kat_cols = [c for c in summary_pivot.columns if c != "SKPD_Nama"]
      summary_pivot["Total Unit"] = summary_pivot[kat_cols].sum(axis=1)
      st.dataframe(summary_pivot, use_container_width=True)
    else:
      st.info("Tidak ada data kendaraan untuk ditampilkan.")
  else:
    st.markdown("#### 📊 Ringkasan Jumlah Bidang Tanah per SKPD")
    if not active_df.empty and "SKPD_Nama" in active_df.columns:
      summary_tanah = active_df["SKPD_Nama"].value_counts().reset_index()
      summary_tanah.columns = ["Nama SKPD / Dinas", "Jumlah Bidang Tanah"]
      st.dataframe(summary_tanah, use_container_width=True)
    else:
      st.info("Tidak ada data ringkasan tanah.")

  st.markdown("---")
  st.markdown("#### 🔍 Preview Kartu Detail Bergaris & Download")
  if not display_df.empty:
    id_col = (
        display_df.columns[1]
        if len(display_df.columns) > 1
        else display_df.columns[0]
    )
    name_col = (
        display_df.columns[2]
        if len(display_df.columns) > 2
        else display_df.columns[0]
    )

    selected_row_idx = st.selectbox(
        "Pilih Item Aset untuk Lihat Detail Lengkap:",
        options=display_df.index,
        format_func=lambda x: f"Baris {x+1}: {display_df.loc[x, id_col]} | {display_df.loc[x, name_col]}",
    )

    if selected_row_idx is not None:
      row_data = display_df.loc[selected_row_idx]
      columns_list = list(display_df.columns)

      cleaned_values = []
      for idx_col, col in enumerate(columns_list):
        val = row_data[col]
        if pd.isna(val):
          cleaned_values.append("None")
        elif isinstance(val, (int, float)):
          if "harga" in col.lower():
            cleaned_values.append(f"{int(val):,}".replace(",", "."))
          else:
            if val == int(val):
              cleaned_values.append(str(int(val)))
            else:
              cleaned_values.append(str(val))
        else:
          val_str = str(val).strip()
          # Hilangkan 00:00:00 pada format tanggal/timestamp
          if " 00:00:00" in val_str:
            val_str = val_str.replace(" 00:00:00", "")
          cleaned_values.append(val_str)

      detail_df = pd.DataFrame({
          "Atribut / Kolom Data": columns_list,
          "Keterangan / Isi Data": cleaned_values,
      })

      table_bg_color = (
          "2b6cb0" if st.session_state.module == "kendaraan" else "276749"
      )
      styled_preview = (
          detail_df.style.set_table_styles([
              {
                  "selector": "th",
                  "props": [
                      ("background-color", f"#{table_bg_color}"),
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
          "<p style='font-weight:600; color:#2d3748;'>Preview Tabel Bergaris:</p>",
          unsafe_allow_html=True,
      )
      st.dataframe(styled_preview, use_container_width=True, height=450)

      col_e1, col_e2, col_e3 = st.columns(3)

      with col_e1:

        def create_styled_vertical_excel(cols, vals, bg_hex):
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
              start_color=bg_hex, end_color=bg_hex, fill_type="solid"
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

        excel_data = create_styled_vertical_excel(
            columns_list, cleaned_values, table_bg_color
        )
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

        def create_pdf_detail(cols, vals, module_type):
          pdf = FPDF(orientation="P", unit="mm", format="A4")
          pdf.add_page()
          pdf.set_font("Arial", "B", 14)
          if module_type == "kendaraan":
            pdf.set_text_color(43, 108, 176)
          else:
            pdf.set_text_color(39, 103, 73)
          pdf.cell(0, 10, "DETAIL INFORMASI ASET PEMERINTAH DAERAH", 0, 1, "C")
          pdf.ln(4)
          pdf.set_font("Arial", "B", 10)
          if module_type == "kendaraan":
            pdf.set_fill_color(43, 108, 176)
          else:
            pdf.set_fill_color(39, 103, 73)
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

        pdf_bytes = create_pdf_detail(
            columns_list, cleaned_values, st.session_state.module
        )
        st.download_button(
            label="📑 Download PDF Bergaris",
            data=pdf_bytes,
            file_name=f"Detail_Aset_{selected_row_idx+1}.pdf",
            mime="application/pdf",
        )