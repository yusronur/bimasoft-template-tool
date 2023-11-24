import docx
import pyperclip
import re

def ekstrak(txt_path):
    numbers = []

    with open(txt_path, 'r') as file:
        for line in file:
            # Use regular expression to find matches for the pattern "no=xxxx"
            matches = re.findall(r'no=(\d+)', line)
            numbers.extend(matches)

    return numbers

# Replace 'your_document.docm' with the actual path to your Word document
docx_path = 'ceknomor.txt'
result = ekstrak(docx_path)

a = ''
for no in result:
    a+=f'{no}\n'

print(f'{a}')
pyperclip.copy(a)