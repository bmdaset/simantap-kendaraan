import streamlit as st
import pandas as pd
import os
import glob

st.set_page_config(page_title="SIMANTAP - Rekap Aset & Kendaraan", layout="wide")

@st.cache_data
def load_kibc_data():
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
        # Memilih sheet yang sesuai dengan KIB C / Gedung / Bangunan
        target_sheet = next(
            (s for s in sheet_names if "gedung" in s.lower() or "bangunan" in s.lower() or "kib c" in s.lower() or "kibc" in s.lower()),
            sheet_names[0] if sheet_names else None
        )

        if not target_sheet:
            return pd.DataFrame()

        df_raw = pd.read_excel(file_path, sheet_name=target_sheet, header=None)

        # Mencari baris header yang benar secara otomatis berdasarkan kata kunci KIB C
        header_row_idx = 0
        for idx, row in df_raw.head(25).iterrows():
            row_str = " ".join([str(val).lower() for val in row.values])
            if any(k in row_str for k in ["kode barang", "nama barang", "jenis barang", "kondisi", "luas lantai"]):
                header_row_idx = idx
                break

        h1 = df_raw.iloc[header_row_idx].fillna("").astype(str).str.strip().tolist()
        
        # Penanganan multi-level header agar nama kolom sesuai dengan baris asli excel
        if header_row_idx + 1 < len(df_raw):
            h2 = df_raw.iloc[header_row_idx + 1].fillna("").astype(str).str.strip().tolist()
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
                    combined_cols.append(last_main if last_main else "")
            final_cols = combined_cols
            data_start_idx = header_row_idx + 2
        else:
            final_cols = h1
            data_start_idx = header_row_idx + 1

        # Membersihkan nama kolom dari string kosong atau format yang tidak valid
        cleaned_cols = []
        for idx_col, c in enumerate(final_cols):
            c_str = str(c).strip()
            if not c_str or c_str.lower() == "none" or c_str.startswith("Unnamed"):
                # Jika kosong, ambil nama dari baris header pertama atau berikan label deskriptif berdasarkan posisi standar KIB C
                if idx_col == 0: c_str = "No"
                elif idx_col == 1: c_str = "Jenis / Nama Barang"
                elif idx_col == 2: c_str = "Kode Barang"
                elif idx_col == 3: c_str = "Nomor Register"
                elif idx_col == 14: c_str = "Harga (Rp)"
                else: c_str = f"Kolom_{idx_col}"
            cleaned_cols.append(c_str)

        # Mencegah duplikasi nama kolom
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
            df = df[~df[first_col_name].astype(str).str.lower().str.contains("jumlah|total|n o", na=False)].reset_index(drop=True)

        df["SKPD_Nama"] = "DINAS / INSTANSI LAINNYA"
        df["Harga_Clean"] = 0.0

        if not df.empty:
            skpd_col = next((c for c in df.columns if c == "SKPD" or "skpd" in c.lower() or "dinas" in c.lower() or "unit" in c.lower()), None)
            if skpd_col:
                df["SKPD_Nama"] = df[skpd_col].ffill().fillna("DINAS / INSTANSI LAINNYA").astype(str).str.upper().str.strip()
                df[skpd_col] = df["SKPD_Nama"]

            # Mencari kolom harga/nilai secara fleksibel berdasarkan nama kolom KIB C
            harga_col = next((c for c in df.columns if "harga" in c.lower() or "rupiah" in c.lower() or "nilai" in c.lower()), None)
            if not harga_col and len(df.columns) > 14:
                harga_col = df.columns[14]

            if harga_col:
                raw_harga = pd.to_numeric(df[harga_col].astype(str).str.replace(r"[^\d.]", "", regex=True), errors="coerce").fillna(0)
                if raw_harga.mean() > 0 and raw_harga.mean() < 100000:
                    calculated_harga = raw_harga * 1000
                else:
                    calculated_harga = raw_harga

                df["Harga_Clean"] = calculated_harga
                df[harga_col] = calculated_harga

        return df
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memuat data: {e}")
        return pd.DataFrame()

st.title("Dashboard Rekapitulasi Aset & Kendaraan")
df_data = load_kibc_data()

if not df_data.empty:
    st.success("Data KIB C Gedung & Bangunan berhasil dimuat!")
    
    if "Harga_Clean" in df_data.columns:
        total_nilai = df_data["Harga_Clean"].sum()
        st.metric("Total Keseluruhan Nilai Aset", f"Rp {total_nilai:,.0f}".replace(",", "."))

    st.subheader("Preview Tabel Data")
    st.dataframe(df_data, use_container_width=True)
else:
    st.warning("File Excel tidak ditemukan atau struktur sheet KIB C tidak dikenali.")