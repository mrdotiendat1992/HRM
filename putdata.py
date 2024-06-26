# -*- coding: utf-8 -*-

file1 = r"c:\Users\Khanh\Downloads\Book1.xlsx"

import openpyxl
import pyodbc


book = openpyxl.load_workbook(file1)
sheet = book.active

for row in sheet.iter_rows(values_only=True):
    conn = pyodbc.connect("Driver={SQL Server};Server=172.16.60.100;Database=HR;UID=huynguyen;PWD=Namthuan@123;")
    cursor = conn.cursor()
    cursor.execute(f"""
    UPDATE HR.dbo.Danh_sach_CBCNV SET Loai_hop_dong = N'{row[1]}' WHERE MST = '{row[0]}' AND Factory = 'NT1'
                   """)
    conn.commit()
conn.close()