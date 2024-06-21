import pyodbc


used_db = "Driver={SQL Server};Server=172.16.60.100;Database=HR;UID=huynguyen;PWD=Namthuan@123;"
conn = pyodbc.connect(used_db)
cursor = conn.cursor()

for row in cursor.execute("SELECT MST FROM dbo.DANH_SACH_CBCNV WHERE Factory = 'NT2'").fetchall():
    # print(row)
    cursor.execute(f"INSERT INTO dbo.Dang_ky_ca_lam_viec VALUES ('{row[0]}','NT2','2024-05-01','2054-12-31','A1')")
conn.commit()
conn.close()