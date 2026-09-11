import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="SIMANTAP - KIB B Mendagri", layout="wide")

# Generator data konsisten dengan total persis 1,605 item
@st.cache_data
def load_data():
    np.random.seed(42)
    n = 1605
    
    jenis_list = [
        ("Laptop / Notebook", "02.06.01.04.001", "Plastik/Metal", "Lenovo ThinkPad", "14 inch", "Pembelian APBD"),
        ("Printer Laserjet", "02.06.02.01.002", "Plastik", "HP LaserJet Pro", "Standard", "Hibah"),
        ("Kendaraan Dinas Roda 4", "02.02.01.02.001", "Besi", "Toyota Avanza", "1300 CC", "Pembelian APBD"),
        ("Sepeda Motor Dinas", "02.02.01.01.004", "Besi", "Honda Supra X", "125 CC", "Pembelian APBD"),
        ("AC Split 2 PK", "02.05.01.02.001", "Besi/Plastik", "Daikin", "2 PK", "Pembelian APBD"),
        ("Meja Kerja Resepsionis", "02.04.01.01.001", "Kayu Jati", "Lokal / Custom", "120x60cm", "Pembelian APBD"),
        ("Kursi Direktur", "02.04.01.02.003", "Kulit/Metal", "Chitol", "Standard", "Pembelian APBD"),
        ("Server Jaringan", "02.06.01.05.001", "Metal", "Dell PowerEdge", "Rackmount", "Pembelian APBD"),
        ("Proyektor LCD", "02.06.03.01.001", "Plastik", "Epson EB-S400", "Standard", "Pembelian APBD"),
        ("Lemari Arsip Besi", "02.04.02.01.001", "Besi", "Chitatop", "2 Pintu", "Pembelian APBD")
    ]
    
    data = []
    for i in range(1, n + 1):
        item = jenis_list[(i - 1) % len(jenis_list)]
        reg = f"{i:04d}"
        tahun = np.random.choice([2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025])
        
        if "Kendaraan" in item[0]:
            harga = int(np.random.randint(180, 350) * 1000000)
            no_polisi = f"B {np.random.randint(1000, 9999)} ABC"
            no_bpkb = f"M-{np.random.randint(100000, 999999)}"
            no_rangka = f"RNK{np.random.randint(10000,99999)}"
            no_mesin = f"MSN{np.random.randint(10000,99999)}"
        else:
            harga = int(np.random.randint(2, 25) * 1000000)
            no_polisi = "-"
            no_bpkb = "-"
            no_rangka = "-"
            no_mesin = "-"
            
        no_pabrik = f"PBK-{np.random.randint(100, 999)}" if np.random.rand() > 0.3 else "-"
        
        # 16 Kolom Standar KIB B Mendagri
        data.append({
            "No. Urut": i,                               # Kolom 1
            "Nama Barang / Jenis Barang": item[0],       # Kolom 2
            "Nomor Kode Barang": item[1],                # Kolom 3
            "Nomor Register": reg,                       # Kolom 4
            "Merk / Type": item[3],                      # Kolom 5
            "Ukuran / CC": item[4],                      # Kolom 6
            "Bahan": item[2],                            # Kolom 7
            "Tahun Pembelian": tahun,                    # Kolom 8
            "Nomor Pabrik": no_pabrik,                   # Kolom 9
            "Nomor Rangka": no_rangka,                   # Kolom 10
            "Nomor Mesin": no_mesin,                     # Kolom 11
            "Nomor Polisi": no_polisi,                   # Kolom 12
            "Nomor BPKB": no_bpkb,                       # Kolom 13
            "Asal-usul / Cara Perolehan": item[5],       # Kolom 14
            "Harga (Rp)": harga,                         # Kolom 15
            "Keterangan": "Baik (B)" if i % 10 != 0 else "Rusak Ringan (RR)" # Kolom 16
        })
        
    return pd.DataFrame(data)

df = load_data()

st.title("📊 SIMANTAP - Manajemen Aset Daerah")
st.markdown("### Modul Penatausahaan KIB B (Peralatan dan Mesin)")
st.markdown("Sinkronisasi format berdasarkan *Permendagri No. 19 Tahun 2016 & Permendagri No. 47 Tahun 2021*.")

# Summary Metrics
total_items = len(df)
total_val = df["Harga (Rp)"].sum()

col1, col2, col3 = st.columns(3)
col1.metric("Total Item Aset (KIB B)", f"{total_items:,} Unit")
col2.metric("Total Nilai Perolehan", f"Rp {total_val:,.0f}".replace(",", "."))
col3.metric("Standar Format", "16 Kolom Mendagri & Excel")

st.divider()

# Tab Navigation for Views
tab1, tab2 = st.tabs(["📋 Format Resmi KIB B (16 Kolom Mendagri)", "📁 Tampilan Database Excel Master"])

with tab1:
    st.markdown("#### Kartu Inventaris Barang (KIB) B - Peralatan dan Mesin")
    st.info("Menampilkan seluruh 16 atribut kolom baku sesuai ketentuan pengelolaan barang milik daerah.")
    
    search_kib = st.text_input("🔍 Cari berdasarkan Nama Barang, Kode, atau Merk (KIB B):", "")
    if search_kib:
        filtered_kib = df[
            df["Nama Barang / Jenis Barang"].str.contains(search_kib, case=False, na=False) |
            df["Nomor Kode Barang"].str.contains(search_kib, case=False, na=False) |
            df["Merk / Type"].str.contains(search_kib, case=False, na=False)
        ]
    else:
        filtered_kib = df
        
    # Format Rupiah for display
    view_kib = filtered_kib.copy()
    view_kib["Harga (Rp)"] = view_kib["Harga (Rp)"].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))
    
    st.dataframe(view_kib, use_container_width=True, height=520)
    
    csv_kib = filtered_kib.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download KIB B Resmi (CSV)",
        data=csv_kib,
        file_name='KIB_B_Peralatan_dan_Mesin_Mendagri.csv',
        mime='text/csv',
    )

with tab2:
    st.markdown("#### Database Excel Master Aset Daerah")
    st.info("Tampilan ringkas yang disesuaikan dengan struktur sheet spreadsheet Excel internal instansi.")
    
    excel_cols = [
        "No. Urut", "Nomor Kode Barang", "Nama Barang / Jenis Barang", "Nomor Register", 
        "Merk / Type", "Tahun Pembelian", "Asal-usul / Cara Perolehan", "Harga (Rp)", "Keterangan"
    ]
    view_excel = df[excel_cols].copy()
    
    search_excel = st.text_input("🔍 Cari di Database Excel:", "")
    if search_excel:
        view_excel = view_excel[
            view_excel["Nama Barang / Jenis Barang"].str.contains(search_excel, case=False, na=False) |
            view_excel["Nomor Kode Barang"].str.contains(search_excel, case=False, na=False)
        ]
        
    view_excel["Harga (Rp)"] = view_excel["Harga (Rp)"].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))
    
    st.dataframe(view_excel, use_container_width=True, height=520)
    
    csv_excel = view_excel.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Database Excel Master (CSV)",
        data=csv_excel,
        file_name='Database_Excel_Master_Aset.csv',
        mime='text/csv',
    )