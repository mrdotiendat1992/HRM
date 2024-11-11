# Tham số kết nối đển database (trong config flask, dùng cho phần đăng nhập sử dụng ORM với flask_sqlalchemy) 
import os
from dotenv import load_dotenv

load_dotenv() 

odbc_driver = "{ODBC Driver 17 for SQL Server}"
database_server = os.environ.get("DATABASE_SERVER")
database_name = os.environ.get("DATABASE_NAME")
database_user = os.environ.get("DATABASE_USER")
database_password = os.environ.get("DATABASE_PASSWORD")

# url kết nối với database (sử dụng khi dùng0 pyodbc)
url_database_pyodbc = f"Driver={{SQL Server}};Server={database_server};Database={database_name};UID={database_user};PWD={database_password};"
