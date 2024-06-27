import sqlite3
import pyodbc

# conn = pyodbc.connect("Driver={SQL Server};"
#     "Server=172.16.60.100;"
#     "Database=HR;"
#     "UID=huynguyen;"
#     "PWD=Namthuan@123;")

# cursor = conn.cursor()
# rows = cursor.execute("SELECT MST,Ho_ten,Department,Factory,Mat_khau FROM dbo.DANH_SACH_CBCNV").fetchall()
# conn.close()
# x=1
# for row in rows:
#     mst = int(row[0])
#     hoten = row[1]
#     phongban = row[2]
#     macongty = row[3]
#     if macongty == "NT1":
#         tencongty = "Công ty cổ phần sản xuất Nam Thuận"
#     elif macongty == "NT2":
#         tencongty = "Công ty cổ phần Nam Thuận Nghệ An"
#     elif macongty == "NT0":
#         tencongty = "Công ty cổ phần tập đoàn Nam Thuận"
#     matkhau = "1"
#     role = "sa"
#     conn1 = sqlite3.connect('instance/db.sqlite')
#     cursor1 = conn1.cursor()

#     rows = cursor1.execute(f"""INSERT INTO users VALUES ({x}, '{mst}','{hoten}','{macongty}','{tencongty}','{phongban}','{matkhau}','{role}')""")
#     conn1.commit()
#     conn1.close()
#     x+=1


import sqlite3

# Kết nối đến cơ sở dữ liệu SQLite
conn = sqlite3.connect('instance/db.sqlite')
cursor = conn.cursor()

cursor.execute("UPDATE users SET role = 'sa' WHERE masothe='12579' AND macongty='NT1'")
# cursor.execute("UPDATE users SET role = 'user'").fetchall()
conn.commit()
conn.close()