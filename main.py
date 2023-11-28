import re
import csv
import openpyxl
from utils import cleanhtml, desc_identificate

# регулярное выражение для очистки строк от html разметки
CLEANR = re.compile('<.*?>|&([a-z0-9]+|#[0-9]{1,6}|#x[0-9a-f]{1,6});|_x\\S{1,5}_')

# путь к файлу с данными
PATH = "раковины.csv"

# Путь к файлу для записи совпавших id (создать заранее)
SAVE_ID = "rakoviny_id.xlsx"

# Путь для записи совпаших названий (создать заранее)
SAVE_INFO = "rakoviny_info.xlsx"

# Чтение csv файла и его запись в переменную в виде списка
with open(PATH, "r", encoding="utf-8") as file:
    reader = csv.reader(file, delimiter=";")
    data = list(reader)

# номера строк для записи в Excel файлы
save_id_row = 1
save_info_row = 1
# Список id совпавших товаров
working_id = []
compare_count = 0
for i, row in enumerate(data):
    # Проверка строки на валидность
    if i == 0 or len(row) != 5:
        continue
    product_id, product_name, description, product_brand, product_collection = row
    # Проверка есть ли id в списке отработанных id
    if product_id in working_id:
        continue
    clean_description = cleanhtml(CLEANR, description)
    identification_list = []
    for n, compare_row in enumerate(data):
        if n == 0 or i == n or len(compare_row) != 5:
            continue
        compare_id, compare_name, compare_description, compare_brand, compare_collection = compare_row
        clean_compare_description = cleanhtml(CLEANR, compare_description)
        if len(product_collection) > 1:
            if product_collection == compare_collection:
                ident_percent = desc_identificate(clean_description, clean_compare_description)
            else:
                continue
        else:
            if product_brand == compare_brand:
                ident_percent = desc_identificate(clean_description, clean_compare_description)
            else:
                continue
        if ident_percent > 80 and compare_id not in working_id:
            print(f"Установлено совпадение: {product_name} == {compare_name}")
            identification_list.append([compare_id, compare_name, int(ident_percent)])
            working_id.append(compare_id)
    if len(identification_list) > 0:
        workbook = openpyxl.load_workbook(SAVE_ID)
        sheet = workbook.active
        sheet.cell(row=save_id_row, column=1, value=product_id)
        column = 2
        for value_id in identification_list:
            sheet.cell(row=save_id_row, column=column, value=value_id[0])
            column += 1
        workbook.save(SAVE_ID)
        workbook.close()
        info_workbook = openpyxl.load_workbook(SAVE_INFO)
        info_sheet = info_workbook.active
        info_sheet.cell(row=save_info_row, column=1, value=product_id)
        info_sheet.cell(row=save_info_row, column=2, value=product_name)
        info_sheet.cell(row=save_info_row, column=3, value=clean_description)
        info_sheet.cell(row=save_info_row, column=4, value=product_brand)
        info_sheet.cell(row=save_info_row, column=5, value=product_collection)
        save_info_row += 1
        for values in identification_list:
            compare_id, compare_name, ident_percent = values
            info_sheet.cell(row=save_info_row, column=2, value=compare_name)
            info_sheet.cell(row=save_info_row, column=3, value=compare_id)
            info_sheet.cell(row=save_info_row, column=4, value=ident_percent)
            save_info_row += 1
        info_workbook.save(SAVE_INFO)
        info_workbook.close()
        print(f"Отработан столбец № {save_id_row}. Установлено совпадений: {len(identification_list)}")
        print(i)
        working_id.append(product_id)
        save_id_row += 1
