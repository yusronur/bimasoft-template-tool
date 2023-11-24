import pyperclip
import re

def ekstrak(file_path):
    nomor = []

    with open(file_path, 'r') as file:
        for line in file:
            # cari menggunakan regex "no=xxxx"
            matches = re.findall(r'no=(\d+)', line)
            nomor.extend(matches)
    
    #hapus yang ganda dan urutkan
    unik = list(dict.fromkeys(nomor))
    return unik

#ambil data dari ceknomor.txt
docx_path = 'ceknomor.txt'
result = ekstrak(docx_path)

a = ''
for no in result:
    a+=f'{no}\n'

print(f'{a}')
pyperclip.copy(a)
#hasil nomor otomatis masuk clipboard, tinggal paste :)
