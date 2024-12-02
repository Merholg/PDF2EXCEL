#!/usr/bin/env python3

import openpyxl
import sys
import pdfplumber
import pandas as pd


def pdf_to_excel(pdf_file, excel_file):
    with pdfplumber.open(pdf_file) as pdf:
        all_tables = []
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table:
                    df = pd.DataFrame(table)
                    all_tables.append(df)

        if not all_tables:
            all_tables.append(pd.DataFrame([["No tables found"]]))

        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            for idx, df in enumerate(all_tables):
                df.to_excel(writer, sheet_name=f'Sheet{idx}', index=False)


if __name__ == '__main__':
    """
    Чтение таблиц из PDF файла и их запись в файл XLSX
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
