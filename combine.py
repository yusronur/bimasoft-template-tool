import pandas as pd
import os
import re
from tqdm import tqdm

#ambil path tempat excel disimpan
folder_path = "G:/My Drive/A/B/!CLOUD/2.KESISWAAN_NEW/2023-2024/UJIAN/1. PTS GANJIL 2023-2024/TEMPLATE/excel"
all_files = os.listdir(folder_path)

xlsm_files = [file for file in all_files if file.endswith('.xlsm')]
# print(all_files)
combined_data = pd.DataFrame()

for file_name in tqdm(xlsm_files, desc="Processing files", unit="file"):

    # pisahkan nama mapel dan kelas
    # dalam hal ini format nama mapel yang digunakan
    # cth : 102023-PRAKARYA-KELAS 9.xlsm
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
        'File Name': [file_name] * len(df),  # Repeat the file name for each row
        'Username': username,
        'Nilai': nilai,
        'Mapel':mapel,
        'Kelas':kelas
    })
    
    # Append the column data to the combined_data DataFrame
    combined_data = pd.concat([combined_data, file_data], ignore_index=True)

# Create a new Excel file and write the combined data to it
combined_data.to_excel('combine.xlsx', index=False)
