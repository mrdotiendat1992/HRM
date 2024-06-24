import pyodbc
conn = pyodbc.connect("Driver={SQL Server};Server=172.16.60.100;Database=HR;UID=huynguyen;PWD=Namthuan@123;")
cursor = conn.cursor()
query = f"""
SELECT 
    Danh_sach_CBCNV.Factory AS Nha_may,
    Danh_sach_CBCNV.MST,
    Danh_sach_CBCNV.Line AS Line,
    Danh_sach_CBCNV.Department AS Department,
    Check_In_Out.NgayCham AS Ngay,
    Check_In_Out.GioCham
FROM 
    Danh_sach_CBCNV
JOIN 
    Check_In_Out 
ON 
    Danh_sach_CBCNV.Factory = Check_In_Out.Nha_may
AND
    Danh_sach_CBCNV.MST = Check_In_Out.MaChamCong
WHERE
    Nha_may = 'NT1'
ORDER BY Ngay DESC, CAST(MST AS INT) ASC, Line ASC, Department ASC
"""
# print(query)
rows = cursor.execute(query).fetchall()
conn.close()
for row in rows:
    print(row)