from io import BytesIO
import glob
import os
import openpyxl
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
    .card-menu-kendaraan {
        background: linear-gradient(135deg, #2b6cb0 0%, #4299e1 100%);
        padding: 24px;
        border-radius: 14px;
        color: white;
        box-shadow: 0 10px 25px rgba(43,108,176,0.35);
        margin-bottom: 15px;
        border-left: 6px solid #bee3f8;
    }
    .card-menu-tanah {
        background: linear-gradient(135deg, #276749 0%, #38a169 100%);
        padding: 24px;
        border-radius: 14px;
        color: white;
        box-shadow: 0 10px 25px rgba(39,103,73,0.35);
        margin-bottom: 15px;
        border-left: 6px solid #c6f6d5;
    }
    .card-menu-gedung {
        background: linear-gradient(135deg, #975a16 0%, #d69e2e 100%);
        padding: 24px;
        border-radius: 14px;
        color: white;
        box-shadow: 0 10px 25px rgba(151,90,22,0.35);
        margin-bottom: 15px;
        border-left: 6px solid #feebc8;
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


def get_excel_filename():
  file_path = "REKAP KENDARAAN TA. 2026.YP .xlsx"
  if not os.path.exists(file_path):
    all_excel = glob.glob("*.xlsx") + glob.glob("*.xls")
    if all_excel:
      return all_excel[0]
  return file_path


@st.cache_data
def load_kendaraan_data():
  file_path = get_excel_filename()
  if not os.path.exists(file_path):
    return pd.DataFrame()
  try:
    df = pd.read_excel(file_path, sheet_name="KENDARAAN DINAS", header=0)
    df.columns = [str(c).strip() for c in df.columns]

    if "no_urut" in df.columns:
      df = df.dropna(subset=["no_urut"]).reset_index(drop=True)
    if len(df) > 1609:
      df = df.iloc[:1609].reset_index(drop=True)

    if "pengguna_pemakai" not in df.columns:
      df["pengguna_pemakai"] = ""
    else:
      df["pengguna_pemakai"] = (
          df["pengguna_pemakai"]
          .fillna("")
          .astype(str)
          .replace(["nan", "None", "NONE"], "")
      )

    df["SKPD_Nama"] = "DINAS / INSTANSI LAINNYA"
    df["Harga_Clean"] = 0.0
    df["Kategori_Jenis"] = "Mobil"

    skpd_col = next(
        (
            c
            for c in df.columns
            if "skpd" in c.lower() or "dinas" in c.lower() or "unit" in c.lower()
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

      if raw_harga.median() < 100000 and raw_harga.median() > 0:
        calculated_harga = raw_harga * 1000
      else:
        calculated_harga = raw_harga

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
    return pd.DataFrame()


@st.cache_data
def load_kiba_data():
  file_path = get_excel_filename()
  if not os.path.exists(file_path):
    return pd.DataFrame()
  try:
    df = pd.read_excel(file_path, sheet_name="KIB A Tanah", header=0)
    df.columns = [str(c).strip() for c in df.columns]
    if "no" in df.columns:
      df["no"] = range(1, len(df) + 1)
    df = df.reset_index(drop=True)
    return df
  except Exception as e:
    return pd.DataFrame()


@st.cache_data
def load_kibc_data():
  file_path = get_excel_filename()
  if not os.path.exists(file_path):
    return pd.DataFrame()
  try:
    df = pd.read_excel(
        file_path, sheet_name="KIB C - Gedung dan Bangunan", header=0
    )
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    if len(df) > 3522:
      df = df.iloc[:3522].reset_index(drop=True)
    return df
  except Exception as e:
    return pd.DataFrame()


# Inisialisasi Session State
if "df_kendaraan_live" not in st.session_state:
  st.session_state.df_kendaraan_live = load_kendaraan_data()
if "df_tanah_live" not in st.session_state:
  st.session_state.df_tanah_live = load_kiba_data()
if "df_gedung_live" not in st.session_state:
  st.session_state.df_gedung_live = load_kibc_data()

if "page" not in st.session_state:
  st.session_state.page = "menu"
if "module" not in st.session_state:
  st.session_state.module = "kendaraan"
if "keyword" not in st.session_state:
  st.session_state.keyword = ""
if "title" not in st.session_state:
  st.session_state.title = ""
if "selected_original_index" not in st.session_state:
  st.session_state.selected_original_index = None

st.markdown(
    """
    <div class="main-header">
        <h2>🏛️ SIMANTAP - ASET PEMERINTAH DAERAH</h2>
        <p>Sistem Informasi Manajemen Aset & Inventaris Kendaraan, Tanah, dan Gedung (BMD)</p>
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

  col1, col2, col3 = st.columns(3, gap="medium")

  with col1:
    st.markdown(
        f"""
        <div class="card-menu-kendaraan">
            <h3 style="margin-top:0; color:#fff;">🚗 Kendaraan Dinas</h3>
            <p style="color:#f7fafc; font-size:14px;"><b>Total Unit:</b> {len(st.session_state.df_kendaraan_live):,} Data<br>Kelola mobil, motor, pick-up, & alat berat.</p>
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
            <p style="color:#f7fafc; font-size:14px;"><b>Total Bidang:</b> {len(st.session_state.df_tanah_live):,} Data<br>Kelola data aset tanah dan alamat lokasi.</p>
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

  with col3:
    st.markdown(
        f"""
        <div class="card-menu-gedung">
            <h3 style="margin-top:0; color:#fff;">🏢 KIB C - Gedung</h3>
            <p style="color:#744210; font-size:14px;"><b>Total Gedung:</b> {len(st.session_state.df_gedung_live):,} Data<br>Kelola gedung & bangunan permanen.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Kelola KIB C Gedung", key="btn_mod_gedung"):
      st.session_state.module = "gedung"
      st.session_state.keyword = ""
      st.session_state.title = "KIB C - Gedung & Bangunan"
      st.session_state.page = "table"
      st.rerun()

elif st.session_state.page == "table":
  col_back, col_ref = st.columns([1.5, 1.5])
  with col_back:
    if st.button("⬅️ Kembali ke Menu Utama"):
      st.session_state.page = "menu"
      st.rerun()
  with col_ref:
    if st.button("🔄 Refresh Data"):
      st.cache_data.clear()
      st.session_state.df_kendaraan_live = load_kendaraan_data()
      st.session_state.df_tanah_live = load_kiba_data()
      st.session_state.df_gedung_live = load_kibc_data()
      st.rerun()

  module_accent = (
      "#2b6cb0"
      if st.session_state.module == "kendaraan"
      else ("#276749" if st.session_state.module == "tanah" else "#975a16")
  )
  st.markdown(
      f"<h3 style='color:{module_accent}; font-weight:700;'>📂"
      f" {st.session_state.title}</h3>",
      unsafe_allow_html=True,
  )

  if st.session_state.module == "kendaraan":
    active_df = st.session_state.df_kendaraan_live
    if st.session_state.keyword != "":
      active_df = active_df[
          active_df["Kategori_Jenis"] == st.session_state.keyword
      ]
  elif st.session_state.module == "tanah":
    active_df = st.session_state.df_tanah_live
  else:
    active_df = st.session_state.df_gedung_live

  pilih_skpd = "Semua SKPD"
  if not active_df.empty and "skpd" in active_df.columns:
    skpd_list = ["Semua SKPD"] + sorted(
        list(active_df["skpd"].dropna().unique())
    )
    pilih_skpd = st.selectbox(
        "🏢 Filter Berdasarkan Nama SKPD / Dinas:", skpd_list
    )

    # --- REKAPITULASI KATEGORI PER SKPD (JUMLAH BARANG & TOTAL NILAI) ---
    if pilih_skpd != "Semua SKPD" and st.session_state.module == "kendaraan":
      df_skpd_selected = active_df[active_df["SKPD_Nama"] == pilih_skpd]

      st.markdown("---")
      st.markdown(f"##### 📊 Rekapitulasi Kategori untuk: **{pilih_skpd}**")

      if (
          "Kategori_Jenis" in df_skpd_selected.columns
          and "Harga_Clean" in df_skpd_selected.columns
      ):
        # Tambahan Informasi Total Keseluruhan Barang & Nilai SKPD Terpilih
        total_item_skpd = len(df_skpd_selected)
        total_nilai_skpd = df_skpd_selected["Harga_Clean"].sum()
        formatted_total_nilai_skpd = (
            f"Rp {total_nilai_skpd:,.0f}".replace(",", ".")
            if total_nilai_skpd > 0
            else "Rp 0"
        )

        col_tot1, col_tot2 = st.columns(2)
        with col_tot1:
          st.metric(
              label="📦 Total Keseluruhan Barang SKPD",
              value=f"{total_item_skpd:,} Unit",
          )
        with col_tot2:
          st.metric(
              label="💰 Total Keseluruhan Nilai Aset SKPD",
              value=formatted_total_nilai_skpd,
          )

        st.markdown("")  # Spasi pemisah

        rekap_df = (
            df_skpd_selected.groupby("Kategori_Jenis")
            .agg(
                Jumlah_Barang=("Kategori_Jenis", "count"),
                Total_Nilai=("Harga_Clean", "sum"),
            )
            .reset_index()
        )

        kategori_standar = ["Alat Berat", "Mobil", "Pick Up", "Sepeda Motor"]
        df_standar = pd.DataFrame({"Kategori_Jenis": kategori_standar})
        rekap_df = (
            pd.merge(df_standar, rekap_df, on="Kategori_Jenis", how="left")
            .fillna(0)
        )

        rekap_df["Total_Nilai_Format"] = rekap_df["Total_Nilai"].apply(
            lambda x: f"Rp {x:,.0f}".replace(",", ".") if x > 0 else "Rp 0"
        )

        cols_rekap = st.columns(len(kategori_standar))
        for idx, row in rekap_df.iterrows():
          with cols_rekap[idx]:
            st.metric(
                label=f"{row['Kategori_Jenis']}",
                value=f"{int(row['Jumlah_Barang'])} Barang",
                delta=row["Total_Nilai_Format"],
                delta_color="off",
            )
      st.markdown("---")
    # -----------------------------------------------------------------

    if pilih_skpd != "Semua SKPD":
      if st.session_state.module == "tanah":
        active_df = active_df[active_df["skpd"] == pilih_skpd]
      else:
        active_df = active_df[active_df["SKPD_Nama"] == pilih_skpd]

  search_query = st.text_input("🔍 Cari Data Berdasarkan Kata Kunci...")
  if search_query:
    mask = (
        active_df.astype(str)
        .apply(
            lambda col: col.str.lower().str.contains(
                search_query.lower(), na=False
            )
        )
        .any(axis=1)
    )
    active_df = active_df[mask]

  st.info(f"Menampilkan {len(active_df):,} baris data aset")

  if st.session_state.module == "kendaraan" and "Harga_Clean" in active_df.columns:
    total_nilai_aset = active_df["Harga_Clean"].sum()
    st.metric(
        "Total Keseluruhan Nilai Aset Terfilter",
        f"Rp {total_nilai_aset:,.0f}".replace(",", "."),
    )

  display_df = active_df.drop(
      columns=["Harga_Clean", "Kategori_Jenis", "SKPD_Nama"], errors="ignore"
  ).reset_index(drop=True)

  event = st.dataframe(
      display_df,
      use_container_width=True,
      height=400,
      on_select="rerun",
      selection_mode="single-row",
  )
  selected_rows = event.selection.rows if event and event.selection else []
  if selected_rows:
    st.session_state.selected_original_index = active_df.index[
        selected_rows[0]
    ]
    st.session_state.page = "detail"
    st.rerun()

elif st.session_state.page == "detail":
  if st.session_state.module == "kendaraan":
    active_df = st.session_state.df_kendaraan_live
    table_bg_color = "2b6cb0"
  elif st.session_state.module == "tanah":
    active_df = st.session_state.df_tanah_live
    table_bg_color = "276749"
  else:
    active_df = st.session_state.df_gedung_live
    table_bg_color = "975a16"

  if st.button("⬅️ Kembali ke Tabel Utama"):
    st.session_state.page = "table"
    st.rerun()

  st.markdown(
      f"<h3 style='color:#{table_bg_color}; font-weight:700;'>📋 Preview"
      " Kartu Detail Aset Terpilih</h3>",
      unsafe_allow_html=True,
  )

  selected_idx = st.session_state.get("selected_original_index", None)
  if selected_idx is not None and selected_idx in active_df.index:
    row_full = active_df.loc[selected_idx]
    columns_to_drop = ["Harga_Clean", "Kategori_Jenis", "SKPD_Nama"]
    columns_list = [c for c in active_df.columns if c not in columns_to_drop]

    cleaned_values = []
    for col in columns_list:
      val = row_full[col]
      if pd.isna(val) or str(val).lower() in ["nan", "none", ""]:
        cleaned_values.append("-")
      elif isinstance(val, (int, float)):
        if "harga" in col.lower() or "nilai" in col.lower():
          cleaned_values.append(f"{int(val):,}".replace(",", "."))
        else:
          cleaned_values.append(
              str(int(val)) if val == int(val) else str(val)
          )
      else:
        cleaned_values.append(str(val))

    detail_df = pd.DataFrame({
        "Atribut / Kolom Data": columns_list,
        "Keterangan / Isi Data": cleaned_values,
    })

    st.dataframe(detail_df, use_container_width=True, height=350)
    st.markdown("---")

    tab_unduh, tab_edit = st.tabs(
        ["📥 Unduh Dokumen Detail", "✏️ Menu Edit & Simpan ke Excel"]
    )

    with tab_unduh:
      col_e1, col_e2, col_e3 = st.columns(3)
      with col_e1:
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          detail_df.to_excel(writer, index=False, sheet_name="Detail Aset")
        st.download_button(
            label="📊 Download Excel",
            data=output.getvalue(),
            file_name=f"Detail_Aset_{selected_idx+1}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
      with col_e2:
        st.download_button(
            label="📄 Download CSV",
            data=detail_df.to_csv(index=False).encode("utf-8"),
            file_name=f"Detail_Aset_{selected_idx+1}.csv",
            mime="text/csv",
        )

    with tab_edit:
      st.markdown(
          "<p style='color:#2d3748;'>Perbarui atribut di bawah ini. Tombol"
          " simpan akan langsung memperbarui database memori sekaligus"
          " menuliskan perubahan secara permanen ke file Excel asli.</p>",
          unsafe_allow_html=True,
      )

      with st.form(key="form_edit_aset"):
        updated_inputs = {}
        fcol1, fcol2 = st.columns(2)
        for i, col_name in enumerate(columns_list):
          current_val = row_full[col_name]
          val_str = (
              ""
              if pd.isna(current_val)
              or str(current_val).lower() in ["nan", "none", ""]
              else str(current_val)
          )
          target_col = fcol1 if i % 2 == 0 else fcol2
          with target_col:
            updated_inputs[col_name] = st.text_input(
                label=col_name,
                value=val_str,
                key=f"input_edit_{selected_idx}_{col_name}",
            )

        submit_edit = st.form_submit_button(
            "💾 Simpan Perubahan ke Database Excel"
        )

        if submit_edit:
          for col_name, new_val in updated_inputs.items():
            original_dtype = active_df[col_name].dtype
            cleaned_input_val = new_val.strip()
            try:
              if pd.api.types.is_integer_dtype(original_dtype):
                converted_val = int(
                    pd.to_numeric(cleaned_input_val, errors="coerce") or 0
                )
              elif pd.api.types.is_float_dtype(original_dtype):
                converted_val = float(
                    pd.to_numeric(cleaned_input_val, errors="coerce") or 0.0
                )
              else:
                converted_val = (
                    "" if cleaned_input_val.lower() == "none" else cleaned_input_val
                )
              active_df.at[selected_idx, col_name] = converted_val
            except:
              active_df.at[selected_idx, col_name] = cleaned_input_val

          if (
              st.session_state.module == "kendaraan"
              and "Harga_Clean" in active_df.columns
          ):
            harga_col = next(
                (
                    c
                    for c in active_df.columns
                    if "harga" in c.lower()
                    or "rupiah" in c.lower()
                    or "nilai" in c.lower()
                ),
                None,
            )
            if harga_col and harga_col in updated_inputs:
              cleaned_harga = (
                  pd.to_numeric(
                      pd.Series([updated_inputs[harga_col]])
                      .astype(str)
                      .str.replace(r"[^\d.]", "", regex=True),
                      errors="coerce",
                  ).fillna(0)
                  .iloc[0]
              )
              active_df.at[selected_idx, "Harga_Clean"] = cleaned_harga

          if st.session_state.module == "kendaraan":
            st.session_state.df_kendaraan_live = active_df
          elif st.session_state.module == "tanah":
            st.session_state.df_tanah_live = active_df
          else:
            st.session_state.df_gedung_live = active_df

          excel_path = get_excel_filename()
          try:
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
              st.session_state.df_kendaraan_live.to_excel(
                  writer, sheet_name="KENDARAAN DINAS", index=False
              )
              st.session_state.df_tanah_live.to_excel(
                  writer, sheet_name="KIB A Tanah", index=False
              )
              st.session_state.df_gedung_live.to_excel(
                  writer,
                  sheet_name="KIB C - Gedung dan Bangunan",
                  index=False,
              )
            st.success(
                "✅ Perubahan berhasil disimpan ke memori DAN langsung"
                " tersimpan permanen di file Excel fisik!"
            )
          except Exception as write_err:
            st.warning(
                f"⚠️ Data berhasil diperbarui di memori aplikasi, namun penulisan file fisik mengalami kendala: {write_err}"
            )

          st.rerun()
  else:
    st.warning("Data item aset tidak ditemukan.")