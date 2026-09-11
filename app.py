from io import BytesIO
import glob
import os
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
    .stApp { background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); }
    .main-header {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        padding: 28px;
        border-radius: 16px;
        color: white;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 8px 25px rgba(30, 60, 114, 0.3);
    }
    .kib-card {
        padding: 24px;
        border-radius: 14px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 6px 18px rgba(0,0,0,0.15);
    }
    .kib-b { background: linear-gradient(135deg, #2b5876 0%, #4e4376 100%); }
    .kib-a { background: linear-gradient(135deg, #134e5e 0%, #71b280 100%); }
    .sub-card { padding: 18px; border-radius: 12px; color: white; margin-bottom: 12px; }
    .sub-1 { background: linear-gradient(135deg, #4e54c8 0%, #8f94fb 100%); }
    .sub-2 { background: linear-gradient(135deg, #ff7e5f 0%, #feb47b 100%); }
    .stButton>button { width: 100%; border-radius: 8px; font-weight: 600; padding: 10px; }
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

  # 1. LOAD KIB B (KENDARAAN DINAS) DENGAN PEMetaan JUDUL KOLOM RESMI
  df_kendaraan = pd.DataFrame()
  try:
    df_raw_k = pd.read_excel(
        file_path, sheet_name="KENDARAAN DINAS", header=None
    )
    h_idx_k = 12
    for idx, row in df_raw_k.head(20).iterrows():
      txt = " ".join(str(v) for v in row.values).lower()
      if (
          "kode barang" in txt
          or "merk" in txt
          or "nomor" in txt
          or "jenis" in txt
      ):
        h_idx_k = idx
        break

    df_k = pd.read_excel(file_path, sheet_name="KENDARAAN DINAS", header=h_idx_k)
    df_k = df_k.loc[:, ~df_k.columns.astype(str).str.contains("^Unnamed")]
    df_k = df_k.dropna(how="all").reset_index(drop=True)

    # Standarisasi & Pemetaan Judul Kolom agar tidak ada yang terpotong/kosong
    raw_cols = [str(c).strip() for c in df_k.columns]
    mapped_cols = []
    for c in raw_cols:
      cl = c.lower()
      if "no" in cl and ("urut" in cl or cl == "no"):
        mapped_cols.append("No. Urut")
      elif "kode" in cl and "barang" in cl:
        mapped_cols.append("Kode Barang")
      elif (
          "nama" in cl
          or "jenis" in cl
          or "barang" in cl
          or "nama barang" in cl
      ):
        mapped_cols.append("Nama Barang / Jenis Barang")
      elif "register" in cl:
        mapped_cols.append("Nomor Register")
      elif "merk" in cl or "type" in cl:
        mapped_cols.append("Merk / Type")
      elif "ukuran" in cl or "cc" in cl:
        mapped_cols.append("Ukuran / CC")
      elif "bahan" in cl:
        mapped_cols.append("Bahan")
      elif "tahun" in cl or "pembelian" in cl:
        mapped_cols.append("Tahun Pembelian")
      elif "pabrik" in cl:
        mapped_cols.append("Nomor Pabrik")
      elif "rangka" in cl:
        mapped_cols.append("Nomor Rangka")
      elif "mesin" in cl:
        mapped_cols.append("Nomor Mesin")
      elif "polisi" in cl:
        mapped_cols.append("Nomor Polisi")
      elif "bpkb" in cl:
        mapped_cols.append("Nomor BPKB")
      elif "asal" in cl or "perolehan" in cl:
        mapped_cols.append("Asal-usul Cara Perolehan")
      elif "harga" in cl or "rupiah" in cl:
        mapped_cols.append("Harga (Rp)")
      elif "ket" in cl:
        mapped_cols.append("Keterangan")
      else:
        mapped_cols.append(c if c and c != "nan" else "Atribut Lainnya")

    df_k.columns = mapped_cols
    df_k = df_k[
        df_k.iloc[:, 0].notna()
        & (df_k.iloc[:, 0].astype(str).str.strip() != "")
    ].reset_index(drop=True)

    # Filter baris rekap / total agar jumlah pas (1,603 data)
    mask_valid_k = ~df_k.astype(str).apply(
        lambda col: col.str.lower().str.contains(
            r"\bjumlah\b|\btotal\b|sub total|r a p i t|h a r g a"
        )
    ).any(axis=1)
    df_k = df_k[mask_valid_k].reset_index(drop=True)

    # Pembersihan kolom Harga untuk kalkulasi nilai uang
    harga_col_k = (
        "Harga (Rp)"
        if "Harga (Rp)" in df_k.columns
        else next(
            (c for c in df_k.columns if "harga" in c.lower()), df_k.columns[-2]
        )
    )

    def parse_rupiah(val):
      try:
        s = (
            str(val)
            .replace("Rp", "")
            .replace("RP", "")
            .replace(" ", "")
            .strip()
        )
        if not s or s.lower() == "nan":
          return 0.0
        if "," in s and "." in s:
          s = (
              s.replace(".", "").replace(",", ".")
              if s.find(",") > s.find(".")
              else s.replace(",", "")
          )
        elif "," in s:
          s = s.replace(".", "").replace(",", ".")
        return float(s)
      except:
        return 0.0

    df_k["Harga_Clean"] = df_k[harga_col_k].apply(parse_rupiah).fillna(0)

    # Kategori Jenis Kendaraan untuk Sub-Menu
    def deteksi_kategori(row):
      combined = " ".join(
          [str(val) for val in row.values if pd.notna(val)]
      ).lower()
      if any(k in combined for k in ["pick up", "pickup", "bak terbuka"]):
        return "Pick Up"
      elif any(
          k in combined
          for k in ["motor", "roda dua", "trail", "matic", "bebek", "scoopy"]
      ):
        return "Sepeda Motor"
      else:
        return "Mobil"

    df_k["Kategori_Jenis"] = df_k.apply(deteksi_kategori, axis=1)
    df_kendaraan = df_k
  except Exception as e:
    print("Error KENDARAAN:", e)

  # 2. LOAD KIB A (TANAH)
  df_tanah = pd.DataFrame()
  try:
    df_raw_t = pd.read_excel(file_path, sheet_name="KIB A TANAH", header=None)
    h_idx_t = 0
    for idx, row in df_raw_t.head(20).iterrows():
      if "luas" in " ".join(str(v) for v in row.values).lower():
        h_idx_t = idx
        break
    df_t = pd.read_excel(file_path, sheet_name="KIB A TANAH", header=h_idx_t)
    df_t = df_t.loc[:, ~df_t.columns.astype(str).str.contains("^Unnamed")].dropna(
        how="all"
    )
    df_t = df_t[
        ~df_t.astype(str)
        .apply(lambda col: col.str.lower().str.contains(r"\bjumlah\b|\btotal\b"))
        .any(axis=1)
    ].reset_index(drop=True)
    harga_col_t = next(
        (c for c in df_t.columns if "harga" in c.lower()), df_t.columns[-1]
    )
    df_t["Harga_Clean"] = pd.to_numeric(
        df_t[harga_col_t].astype(str).str.replace(r"[^\d.]", "", regex=True),
        errors="coerce",
    ).fillna(0.0)
    df_tanah = df_t
  except Exception as e:
    print("Error KIB A:", e)

  return df_kendaraan, df_tanah, file_path


df_kendaraan, df_tanah, file_path = load_all_data()


def format_rupiah(nilai):
  return "Rp " + f"{nilai:,.2f}".replace(",", "X").replace(".", ",").replace(
      "X", "."
  )


if "page" not in st.session_state:
  st.session_state.page = "menu"
if "module" not in st.session_state:
  st.session_state.module = ""
if "keyword" not in st.session_state:
  st.session_state.keyword = ""

st.markdown(
    """<div class="main-header"><h2>🏛️ SIMANTAP - MANAJEMEN ASET DAERAH</h2><p>Sistem Informasi Inventaris Barang Daerah Berbasis Database Excel</p></div>""",
    unsafe_allow_html=True,
)

# --- MENU UTAMA ---
if st.session_state.page == "menu":
  c1, c2 = st.columns(2, gap="large")
  with c1:
    st.markdown(
        f"""<div class="kib-card kib-b"><h3>🚗 KIB B (Peralatan & Mesin)</h3>
        <p>Total Unit: <b>{len(df_kendaraan):,} Data</b><br>Total Nilai Aset: <b>{format_rupiah(df_kendaraan['Harga_Clean'].sum() if not df_kendaraan.empty else 0)}</b></p></div>""",
        unsafe_allow_html=True,
    )
    if st.button("Kelola KIB B (Kendaraan Dinas)"):
      st.session_state.module = "kendaraan"
      st.session_state.page = "sub_menu_kendaraan"
      st.rerun()
  with c2:
    st.markdown(
        f"""<div class="kib-card kib-a"><h3>🗺️ KIB A (Tanah)</h3>
        <p>Total Bidang: <b>{len(df_tanah):,} Data</b><br>Total Nilai Aset: <b>{format_rupiah(df_tanah['Harga_Clean'].sum() if not df_tanah.empty else 0)}</b></p></div>""",
        unsafe_allow_html=True,
    )
    if st.button("Kelola KIB A (Tanah)"):
      st.session_state.module = "tanah"
      st.session_state.page = "sub_menu_tanah"
      st.rerun()

# --- SUB-MENU KIB B ---
elif st.session_state.page == "sub_menu_kendaraan":
  if st.button("⬅️ Kembali ke Menu Utama"):
    st.session_state.page = "menu"
    st.rerun()
  st.markdown("### 📋 Pilih Kategori Kendaraan Dinas (KIB B)")

  c1, c2 = st.columns(2)
  with c1:
    st.markdown(
        """<div class="sub-card sub-1"><h4>Semua Kendaraan Dinas</h4></div>""",
        unsafe_allow_html=True,
    )
    if st.button("Buka Semua Kendaraan"):
      st.session_state.keyword = ""
      st.session_state.page = "table"
      st.rerun()
    st.markdown(
        """<div class="sub-card sub-2"><h4>Mobil Dinas</h4></div>""",
        unsafe_allow_html=True,
    )
    if st.button("Buka Mobil Dinas"):
      st.session_state.keyword = "Mobil"
      st.session_state.page = "table"
      st.rerun()
  with c2:
    st.markdown(
        """<div class="sub-card sub-2"><h4>Sepeda Motor Dinas</h4></div>""",
        unsafe_allow_html=True,
    )
    if st.button("Buka Sepeda Motor"):
      st.session_state.keyword = "Sepeda Motor"
      st.session_state.page = "table"
      st.rerun()
    st.markdown(
        """<div class="sub-card sub-1"><h4>Pick Up / Truk</h4></div>""",
        unsafe_allow_html=True,
    )
    if st.button("Buka Pick Up"):
      st.session_state.keyword = "Pick Up"
      st.session_state.page = "table"
      st.rerun()

# --- SUB-MENU KIB A ---
elif st.session_state.page == "sub_menu_tanah":
  if st.button("⬅️ Kembali ke Menu Utama"):
    st.session_state.page = "menu"
    st.rerun()
  st.markdown("### 🗺️ Pilih Kategori KIB A - Tanah")
  if st.button("Buka Semua Data Tanah"):
    st.session_state.keyword = ""
    st.session_state.page = "table"
    st.rerun()

# --- HALAMAN TABEL DATA LENGKAP DENGAN JUDUL KOLOM EXCEL ---
elif st.session_state.page == "table":
  if st.button("⬅️ Kembali"):
    st.session_state.page = (
        "sub_menu_kendaraan"
        if st.session_state.module == "kendaraan"
        else "sub_menu_tanah"
    )
    st.rerun()

  df_active = (
      df_kendaraan
      if st.session_state.module == "kendaraan"
      else df_tanah
  )
  kategori_col = (
      "Kategori_Jenis" if st.session_state.module == "kendaraan" else None
  )

  filtered_df = (
      df_active
      if st.session_state.keyword == "" or not kategori_col
      else df_active[df_active[kategori_col] == st.session_state.keyword]
  )

  search_q = st.text_input(
      "🔍 Cari data secara global (Ketik Nomor Rangka, Nomor Mesin, No. Polisi,"
      " BPKB, dll):"
  )
  if search_q:
    filtered_df = filtered_df[
        filtered_df.astype(str)
        .apply(
            lambda col: col.str.lower().str.contains(
                search_q.lower(), na=False
            )
        )
        .any(axis=1)
    ]

  st.info(
      f"Menampilkan {len(filtered_df):,} baris data dari total keseluruhan"
      f" {len(df_active):,} unit | File: {os.path.basename(file_path)}"
  )

  # Drop kolom sistem internal, pertahankan seluruh kolom database Excel (Nomor Rangka, Mesin, Polisi, BPKB, dll)
  display_df = filtered_df.drop(
      columns=["Harga_Clean", "Kategori_Jenis"], errors="ignore"
  )

  # Menampilkan dataframe dengan header kolom lengkap dan fitur scroll horizontal
  st.dataframe(display_df, use_container_width=True, height=550)

  st.markdown("---")
  csv_bytes = display_df.to_csv(index=False).encode("utf-8")
  st.download_button(
      label="📥 Download Database Excel Lengkap (CSV)",
      data=csv_bytes,
      file_name="Database_Excel_KIB_B_Lengkap.csv",
      mime="text/csv",
  )