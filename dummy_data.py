# ===============================
# 1. Install & import libraries
# ===============================
# (di Colab tidak perlu pip install pandas)
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from google.colab import files   # untuk download otomatis

# ===============================
# 2. Parameter
# ===============================
yesterday = datetime.utcnow().date() - timedelta(days=1)
file_name = f"{yesterday:%Y%m%d}_WITSML.csv"

# 20 nama sumur fiktif lepas-pantai Indonesia
wells = [
    "Natuna-A1", "Natuna-A2", "Natuna-B1", "Natuna-B2", "Jawa-C1",
    "Jawa-C2", "Jawa-D1", "Jawa-D2", "Maluku-E1", "Maluku-E2",
    "Maluku-F1", "Maluku-F2", "Sumatra-G1", "Sumatra-G2", "Sumatra-H1",
    "Sumatra-H2", "Kalimantan-I1", "Kalimantan-I2", "Papua-J1", "Papua-J2"
]

# Lokasi acak di sekitar blok Indonesia
np.random.seed(42)  # agar hasil konsisten di demo
lats  = np.round(np.random.uniform(-6.5, 4.5, 20), 4)
lons  = np.round(np.random.uniform(105.5, 119.5, 20), 4)

# ===============================
# 3. Buat DataFrame dummy
# ===============================
data_rows = []
for idx, w in enumerate(wells):
    gross = np.round(np.random.uniform(200, 2500), 1)
    water = np.round(gross * np.random.uniform(0.05, 0.55), 1)   # water-cut 5–55 %
    gas   = np.round(gross * np.random.uniform(600, 1500), 0)    # GOR 600–1500
    choke = f"{np.random.randint(30, 90)}/{np.random.randint(30, 90)}"
    press = np.round(np.random.uniform(750, 2800), 0)
    temp  = np.round(np.random.uniform(160, 220), 1)
    
    data_rows.append({
        "Date"               : yesterday.strftime("%Y-%m-%d"),
        "Well_ID"            : w,
        "Latitude"           : lats[idx],
        "Longitude"          : lons[idx],
        "ChokeSize"          : choke,
        "Gross_Oil_bbl"      : gross,
        "Water_bbl"          : water,
        "Gas_Mscf"           : gas,
        "WellHead_Pressure_psi": press,
        "WellHead_Temp_F"    : temp
    })

df = pd.DataFrame(data_rows)

# ===============================
# 4. Simpan & download
# ===============================
df.to_csv(file_name, index=False)
print("File berhasil dibuat:", file_name)
files.download(file_name)   # otomatis muncul di laptop Anda