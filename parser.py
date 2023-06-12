import os
import numpy as np
import pandas as pd
import fitz

def extract_text(path_pdf):
    doc = fitz.open(path_pdf)
    
    page = doc.load_page(1)
    text = page.get_text()
    
    list_text = []
    list_char = []
    
    for i in text:
       if i == '\n':
           txt = ''.join(list_char)
           txt = txt.strip()
           list_text.append(txt)
           list_char = []
       else:
           list_char.append(i)
    
    try:
        anchor = list_text.index('NIM')
        list_text = list_text[anchor:]
    except:
        return ['' for i in range(7)]

    return list_text

def select_attributes(list_attribute, list_text):
    list_value = []
    list_check = ['Nama', 'Program', 'Program Sertifikasi', 'Penyelenggara',
                  'Homepage', 'Status', 'Status Akhir', 'Status Akhir Kelulusan',
                  'Status Akhir Kelulusan*']

    for attribute in list_attribute:
        try:
            ida = list_text.index(attribute)
            if list_text[ida+2] in list_check:
                list_value.append(list_text[ida+1].replace(': ', '').strip())
            else:
                list_value.append(list_text[ida+2].strip())
        except:
            list_value.append('')
    
    if list_value[0] == 'Nama':
        return ['' for i in range(7)]
    elif list_value == []:
        return ['' for i in range(7)]

    return list_value

# folder path
dir_path = r'laporan-akhir'

# columns
list_attribute = ['NIM', 'Nama', 'Program', 'Program Sertifikasi',
                  'Penyelenggara', 'Nilai Akhir', 'Hasil Akhir',
                  'Nilai Akhir *']

# rows
list_rows = []

# iterate directory
for path in os.listdir(dir_path):
    # check if current path is a file
    if os.path.isfile(os.path.join(dir_path, path)):
        list_text = extract_text(dir_path + '/' + path)
        list_value = select_attributes(list_attribute, list_text)
        list_rows.append(list_value)
        # print(list_value)

df_summary = pd.DataFrame(list_rows)
df_summary.columns = list_attribute
df_summary['Nilai'] = df_summary['Nilai Akhir'] + df_summary['Hasil Akhir'] + df_summary['Nilai Akhir *']
df_summary.drop(['Nilai Akhir', 'Hasil Akhir', 'Nilai Akhir *'], inplace=True, axis=1)

df_summary['Program'] = np.where(df_summary['Program'] == df_summary['Program Sertifikasi'],
                                 df_summary['Program'],
                                 df_summary['Program'] + df_summary['Program Sertifikasi'])
df_summary.drop(['Program Sertifikasi'], inplace=True, axis=1)

try:
    # df_summary.to_excel('summary.xlsx',
    #                     sheet_name='summary')
    print(df_summary)
    print("Success")
except:
    print("Failed")