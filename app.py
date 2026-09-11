from io import BytesIO
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="SIMANTAP - KIB B Kendaraan Dinas",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .stApp { background: #f8fafc; }
    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
        padding: 20px;
        border-radius: 12px;
        color: white;
        margin-bottom: 20px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data
def load_excel_kib_b():
  file_path = "REKAP KENDARAAN TA. 2026.YP.xlsx"
  if not os.path.exists(file_path):
    return pd.DataFrame(), file_path

  try:
    df_raw = pd.read_excel(
        file_path, sheet_name="KENDARAAN DINAS", header=None
    )
    h_idx = 12
    for idx, row in df_raw.head(20).iterrows():
      txt = " ".join(str(v) for v in row.values).lower()
      if (
          "kode barang" in txt
          or "merk" in txt
          or "nomor" in txt
          or "jenis" in txt
      ):
        h_idx = idx
        break

    df = pd.read_excel(file_path, sheet_name="KENDARAAN DINAS", header=h_idx)
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    # Hapus baris kosong pada kolom pertama
    df = df[
        df.iloc[:, 0].notna() & (df.iloc[:, 0].astype(str).str.strip() != "")
    ].reset_index(drop=True)

    # Filter baris rekap/total agar jumlah pas dan bersih
    mask_valid = ~df.astype(str).apply(
        lambda col: col.str.lower().str.contains(
            r"\bjumlah\b|\btotal\b|sub total|r a p i t|h a r g a"
        )
    ).any(axis=1)
    df = df[mask_valid].reset_index(drop=True)

    return df, file_path
  except Exception as e:
    return pd.DataFrame(), file_path


df, file_path = load_excel_kib_b()

st.markdown(
    """
    <div class="main-header">
        <h3>🚗 KIB B - KENDARAAN DINAS (PERALATAN & MESIN)</h3>
        <p>Tampilan Presisi Sesuai Database Excel Master (Menampilkan Seluruh Kolom: No. Rangka, No. Mesin, No. BPKB, No. Polisi, dll)</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if not df.empty:
  total_unit = len(df)

  # Cari kolom harga untuk kalkulasi total nilai uang secara akurat
  harga_col = next(
      (c for c in df.columns if "harga" in c.lower() or "rupiah" in c.lower()),
      None,
  )
  total_nilai = 0.0
  if harga_col:

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
          if s.find(",") > s.find("."):
            s = s.replace(".", "").replace(",", ".")
          else:
            s = s.replace(",", "")
        elif "," in s:
          s = s.replace(".", "").replace(",", ".")
        return float(s)
      except:
        return 0.0

    total_nilai = df[harga_col].apply(parse_rupiah).sum()

  def format_rupiah(nilai):
    formatted = f"{nilai:,.2f}"
    return "Rp " + formatted.replace(",", "X").replace(".", ",").replace("X", ".")

  col1, col2, col3 = st.columns(3)
  col1.metric("Total Unit Data", f"{total_unit:,} Data")
  col2.metric("Total Nilai Aset", format_rupiah(total_nilai))
  col3.metric("Sumber File", os.path.basename(file_path))

  st.divider()

  # Fitur Pencarian Global di Semua Kolom (Nomor Rangka, Mesin, Polisi, dll)
  search_query = st.text_input(
      "🔍 Cari data (Ketik Nomor Rangka, Nomor Mesin, No. Polisi, Merk, dll):"
  )
  if search_query:
    mask_search = (
        df.astype(str)
        .apply(
            lambda col: col.str.lower().str.contains(
                search_query.lower(), na=False
            )
        )
        .any(axis=1)
    )
    display_df = df[mask_search]
  else:
    display_df = df

  st.info(
      f"Menampilkan {len(display_df)} baris data dari total keseluruhan"
      f" {total_unit} unit."
  )

  # Menampilkan tabel secara penuh dengan scroll horizontal agar persis format Excel
  st.dataframe(display_df, use_container_width=True, height=520)

  st.markdown("---")
  csv_data = display_df.to_csv(index=False).encode("utf-8")
  st.download_button(
      label="📥 Download Database Excel KIB B Lengkap (CSV)",
      data=csv_data,
      file_name="Database_KIB_B_Kendaraan_Dinas_Lengkap.csv",
      mime="text/csv",
  )
else:
  st.warning(
      "File Excel tidak ditemukan atau sheet KENDARAAN DINAS tidak terbaca."
  )