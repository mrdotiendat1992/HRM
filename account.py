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

cursor.execute("INSERT INTO users VALUES ('6569','3',N'Nguyễn Thị Kim Anh','NT1',N'Công ty cổ phần sản xuất Nam Thuận','1IED','1','user')")

conn.commit()
conn.close()