local_db = "Driver={SQL Server}; Server=DESKTOP-G635SF6; Database=HR; Trusted_Connection=yes;"

test_db = "Driver={SQL Server}; Server=172.16.60.100; Database=HR; UID=huynguyen; PWD=Namthuan@123;"

mcc_db = "Driver={SQL Server}; Server=10.0.0.252\SQLEXPRESS; Database=MITACOSQL; UID=sa; PWD=Namthuan1;"

# import pyodbc

# conn = pyodbc.connect(mcc_db)

# cursor = conn.cursor()

# cursor.execute("""SELECT [MaChamCong],[NgayCham],[GioCham],[TenMay]
# FROM [MITACOSQL].[dbo].[CheckInOut]
# WHERE NgayCham = '2024-06-23'
# """)

# rows = cursor.fetchall()

# conn.close()

# for row in rows:
#     conn_db = pyodbc.connect(test_db)
    
#     cursor_db = conn_db.cursor()
    
#     cursor_db.execute(f"INSERT INTO HR.dbo.Check_In_Out VALUES('NT1','{row[0]}','{row[1]}','{row[2]}','{row[3]}')")
    
#     conn_db.commit()
    
#     conn_db.close()