import requests
import pyodbc

conn = pyodbc.connect("Driver={SQL Server};"
    "Server=172.16.60.100;"
    "Database=HR;"
    "UID=huynguyen;"
    "PWD=Namthuan@123;")

cursor = conn.cursor()
rows = cursor.execute("SELECT MST,Ho_ten,Department,Factory,Mat_khau FROM dbo.DANH_SACH_CBCNV").fetchall()
conn.close()

for row in rows:
    if row[1] and row[2] and row[3] and row[4]:
        requests.post(f"http://127.0.0.1:5000/register?masothe={row[0]}&hoten={row[1]}&phongban={row[2]}&macongty={row[3]}&matkhau={row[4]}")

# conn1 = pyodbc.connect('DRIVER={SQL Server};SERVER=10.0.0.252/SQLEXPRESS;DATABASE=MITACOSQL;UID=sa;PWD=Namthuan1;')
# cursor1 = conn1.cursor()
# rows1 = cursor1.execute("SELECT MaChamCong,MaThe FROM dbo.NHANVIEN").fetchall()
# conn1.close()

# conn2 = pyodbc.connect("Driver={SQL Server};"
#     "Server=172.16.60.100;"
#     "Database=HR;"
#     "UID=huynguyen;"
#     "PWD=Namthuan@123;")

# cursor2 = conn2.cursor()
    
# for row1 in rows1:
#     cursor2.execute(f"UPDATE Danh_sach_CBCNV SET The_cham_cong = '{row1[1]}' WHERE MST = '{row1[0]}' AND Factory = 'NT1'")
#     conn2.commit()
    
# conn2.close()

# import sqlite3

# # Kết nối đến cơ sở dữ liệu SQLite
# conn = sqlite3.connect('instance/db.sqlite')
# cursor = conn.cursor()

# # Thêm cột 'role' vào bảng 'users'
# rows = cursor.execute('SELECT * FROM users').fetchall()

# for row in rows:
#     cursor.execute(f"UPDATE users SET role = 'sa' WHERE id = '{row[0]}'")
# conn.commit()
