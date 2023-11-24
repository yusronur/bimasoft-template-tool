import pandas as pd
import os
import re
from tqdm import tqdm

#ambil path tempat excel disimpan
folder_path = os.path.dirname(os.path.abspath(__file__))
all_files = os.listdir(folder_path)

xlsm_files = [file for file in all_files if file.endswith('.xlsm')]
# print(all_files)
combined_data = pd.DataFrame()

for file_name in tqdm(xlsm_files, desc="Processing files", unit="file"):

    # pisahkan nama mapel dan kelas
    # dalam hal ini format nama mapel yang digunakan
    # cth : 102023-PRAKARYA-KELAS 9.xlsm
    # hapus jika filename tidak sama formatnya dengan di atas
    match = re.match(r'\d{6}-(.+?)-KELAS (\d+)', file_name)
    if match:
        mapel = match.group(1)
        kelas = f"KELAS {match.group(2)}"
    else:
        mapel = "Unknown"
        kelas = "Unknown"

    # baca file excel file dan ekstrak data pada sheet "HASIL" <= sesuai template excel bimasoft
    df = pd.read_excel(os.path.join(folder_path, file_name), sheet_name='HASIL', engine='openpyxl')
    
    # Ekstrak data pada kolom "Nilai" dan "No. Peserta"
    nilai = df['Nilai']
    username = df['No. Peserta']

    file_data = pd.DataFrame({
        'File Name': [file_name] * len(df), 
        'Username': username,
        'Nilai': nilai,
        'Mapel':mapel,
        'Kelas':kelas
    })
    
    # gabungkan semua data pada masing masing file dan kolom.
    combined_data = pd.concat([combined_data, file_data], ignore_index=True)

# buat file baru dan masukkan hasil gabungan dari masing-masing file ke file tsb
combined_data.to_excel('combine.xlsx', index=False)