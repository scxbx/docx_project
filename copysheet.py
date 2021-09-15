import os
import win32com.client

import openpyxl

book_in = openpyxl.load_workbook(r'C:\Users\sc\Desktop\test_docx\copy.xlsx', data_only=True)

source = book_in.active
for i in range(10):
    target = book_in.copy_worksheet(source)

book_in.save(r'C:\Users\sc\Desktop\test_docx\copy_new.xlsx')