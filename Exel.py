from os import remove
from openpyxl import load_workbook
import openpyxl as xl
import re
import docx


def stolb(to_chto_nushno):
    global s
    if to_chto_nushno == 1:s = f'A{srch}'
    if to_chto_nushno == 2:s = f'B{srch}'
    if to_chto_nushno == 3:s = f'C{srch}'
    if to_chto_nushno == 4:s = f'D{srch}'
    if to_chto_nushno == 5:s = f'I{srch}'
    if to_chto_nushno == 6:s = f'F{srch}'
    if to_chto_nushno == 7:s = f'G{srch}'
    if to_chto_nushno == 8:s = f'H{srch}'
    if to_chto_nushno == 9:s = f'I{srch}'
    if to_chto_nushno == 10:s = f'J{srch}'
    if to_chto_nushno == 11:s = f'K{srch}'
    if to_chto_nushno == 12:s = f'L{srch}'
    if to_chto_nushno == 13:s = f'M{srch}'
    if to_chto_nushno == 14:s = f'N{srch}'
    if to_chto_nushno == 15:s = f'O{srch}'
    if to_chto_nushno == 16:s = f'P{srch}'
    if to_chto_nushno == 17:s = f'Q{srch}'
    if to_chto_nushno == 18:s = f'R{srch}'
    if to_chto_nushno == 19:s = f'S{srch}'
    if to_chto_nushno == 20:s = f'T{srch}'
    if to_chto_nushno == 21:s = f'U{srch}'
    if to_chto_nushno == 22:s = f'V{srch}'
    if to_chto_nushno == 23:s = f'W{srch}'
    if to_chto_nushno == 24:s = f'X{srch}'

# Из документа в обычный txt
doc = docx.Document('Files/Mission.docx')
file = open('Files/tmp.txt','w+')
all_paras = doc.paragraphs
for x in all_paras:
    file.write(f'{x.text}\n') 
file.close()
file = open('Files/tmp.txt',"r")

# Сплит по разделителям
tmp_array = []
for i in re.split(', | |\n',file.read()):
    tmp_array.append(i)
ind = tmp_array.index('Столбцы:')
searname = tmp_array[:ind]
tmp_array = tmp_array[ind+1:]
stolbs = tmp_array[::2]
formuls = tmp_array[1::2]

def del_space():
    for i in searname:
        if i == '':
            searname.remove(i)
    for i in stolbs:
        if i == '':
            stolbs.remove(i)
    for i in formuls:
        if i == '':
            formuls.remove(i)
del_space()

path = r'Files/idk.xlsx' #путь, где к Excel файлу
wb = xl.load_workbook(filename=path)
ws = wb['ОУ'] #Название листа с данными

dont_have = []
chet_formuls=-1
while True:
        for i in stolbs:
            chet_formuls+=1
            for jk in searname:
                break_out_flag = False
                break_out_flag2 = False
                for row in ws.rows:
                    if break_out_flag2:
                        break
                    for cell in row:
                        if break_out_flag:
                            break
                        if re.match(jk, str(cell.value)): #вместо test ввести искомое значение
                            find_str = cell.column_letter + str(cell.row) # cell.value - само слово
                            break_out_flag = True
                            break_out_flag2 = True
                            srch = find_str
                            srch = re.sub('[\D]', '', srch) # взять из него только цифры
                            formula = formuls[chet_formuls].replace('{srch}',f'{srch}')
                            stolb(int(i))
                            sheet_ranges = wb['ОУ']
                            ws = wb ['ОУ'] # Получить лист в соответствии с именем листа
                            ws [s] = formula
                            break
        if chet_formuls==len(formuls)-1:
            break
wb.save ('Files/Сompleted.xlsx')
file.close()
remove('Files/tmp.txt')