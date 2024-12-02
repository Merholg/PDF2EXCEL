#!/usr/bin/env python3

import openpyxl
import sys
import pdfplumber
import pandas as pd


def pdf_to_excel(pdf_file, excel_file):
    title_list = ['Дата', 'Плательщик / Получатель', 'Операция', 'Сумма (RUB)']
    common_tab = [title_list]
    fields_index = []
    common_tab_string = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table:
                    flag = True
                    for tab_string in table:
                        if flag:
                            # поиск строки текущей таблицы являющейся заголовком
                            fields_index.clear()
                            for title_element in title_list:
                                try:
                                    idx = tab_string.index(title_element)
                                except ValueError:
                                    fields_index.append(-1)
                                else:
                                    fields_index.append(idx)
                            if all(number >= 0 for number in fields_index): # если все поля присутствуют - заголовок таблицы найден
                                flag = False  # следующие строки этой таблицы будут внесены в общую таблицу
                        else:
                            common_tab_string.clear()
                            for field_idx in fields_index:
                                common_tab_string.append(tab_string[field_idx])
                            common_tab.append(common_tab_string.copy())
                            #print(common_tab_string)
        #print(common_tab)

        df = pd.DataFrame(common_tab)
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Transactions', index=False)

if __name__ == '__main__':
    """
    Чтение таблиц из PDF файла постранично и их запись в файл XLSX одной таблицей
    Usage: pdf2excel.py file.pdf file.xlsx
    """
    if len(sys.argv) > 2:
        try:
            pdf_to_excel(sys.argv[1], sys.argv[2])
            sys.exit(0)
        except Exception as error:
            print(f"Unexpected error: {error}")
            sys.exit(1)
    else:
        print("Usage: pdf2excel.py file.pdf file.xlsx")
        sys.exit(1)
