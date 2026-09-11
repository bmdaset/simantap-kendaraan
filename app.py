def deteksi_kategori(row):
        combined = " ".join(
            [str(val) for val in row.values if pd.notna(val)]
        ).lower()

        # 1. Kategori Sepeda Motor / Roda Dua
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

        # 2. Kategori Pick Up / Bak Terbuka
        elif any(k in combined for k in ["pick up", "pickup", "bak terbuka"]):
          return "Pick Up"

        # 3. Kategori Alat Berat
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

        # 4. Kategori Mobil (Minibus, Station Wagon, Dump Truck, Truk, Sedan, Jeep, dll.)
        else:
          return "Mobil"