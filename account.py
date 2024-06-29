import sqlite3
import pyodbc

conn = pyodbc.connect("Driver={SQL Server};"
    "Server=172.16.60.100;"
    "Database=HR;"
    "UID=huynguyen;"
    "PWD=Namthuan@123;")

cursor = conn.cursor()
rows = cursor.execute("SELECT MST,Factory,Grade_code FROM dbo.DANH_SACH_CBCNV").fetchall()
conn.close()
for row in rows:
    mst = int(row[0])
    macongty = row[1]
    capbac = row[2]

    conn1 = sqlite3.connect('instance/db.sqlite')
    cursor1 = conn1.cursor()

    rows = cursor1.execute(f"UPDATE users SET role = 'tbp' WHERE capbac = """)
    conn1.commit()
    conn1.close()


# import sqlite3

# conn = sqlite3.connect('instance/db.sqlite')
# cursor = conn.cursor()

# cursor.execute("ALTER TABLE users ADD COLUMN capbac VARCHAR(50)")

# conn.commit()
# conn.close()