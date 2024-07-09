import sqlite3
import pyodbc

# conn = pyodbc.connect("Driver={SQL Server};"
#     "Server=172.16.60.100;"
#     "Database=HR;"
#     "UID=huynguyen;"
#     "PWD=Namthuan@123;")

# cursor = conn.cursor()
# rows = cursor.execute("SELECT distinct MST from Phan_quyen_thu_ky").fetchall()
# conn.close()
# for row in rows:
#     mst = int(row[0])

#     conn1 = sqlite3.connect('instance/db.sqlite')
#     cursor1 = conn1.cursor()

#     rows = cursor1.execute(f"UPDATE users SET role = 'tk' WHERE masothe = {mst}")
#     conn1.commit()
#     conn1.close()


import sqlite3

conn = sqlite3.connect('instance/db.sqlite')
cursor = conn.cursor()
count  = cursor.execute("SELECT count(*) FROM users").fetchone()[0]
cursor.execute(f"INSERT INTO users VALUES ('{count+1}','25',N'Nguyễn Xuân Hoàng Phúc','NT2',N'Công ty cổ phần Nam Thuận Nghệ An','MGT','1','gd')")
cursor.execute(f"INSERT INTO users VALUES ('{count+2}','28',N'Trần Lê Đại Dương','NT2',N'Công ty cổ phần Nam Thuận Nghệ An','MGT','1','gd')")
conn.commit()
conn.close()