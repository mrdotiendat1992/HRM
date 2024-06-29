# import pyodbc

# # Đường dẫn đến tệp Access
# database_path = r'c:\Users\Khanh\Documents\Database1.accdb'

# # Chuỗi kết nối
# conn_str = ("Driver={SQL Server};Server=172.16.60.100;Database=HR;UID=huynguyen;PWD=Namthuan@123;")

# # Kết nối đến cơ sở dữ liệu
# conn = pyodbc.connect(conn_str)
# cursor = conn.cursor()

# # Thực hiện một truy vấn
# cursor.execute("""
# INSERT INTO dbo.DANH_SACH_CBCNV VALUES
# (
#     '12991',
#     '12991',
#     N'Lý Quang Hợp',
#     '0901216365',
#     '1998-04-16',
#     N'Nam',
#     '002098006122',
#     '2021-06-25',
#     N'Cục cảnh sát',
#     '',
#     N'Thôn Vĩnh Ban, Vĩnh Phúc, Bắc Quang, Hà Giang',
#     N'Thôn Vĩnh Ban',
#     N'Vĩnh Phúc',
#     N'Bắc Quang',
#     N'Hà Giang',
#     N'Kinh',
#     N'Việt Nam',
#     N'Không',
#     N'Cấp 3',
#     N'Vĩnh Phúc, Bắc Quang, Hà Giang',
#     N'Thôn Vĩnh Ban, Vĩnh Phúc, Bắc Quang, Hà Giang',
#     '0220686323',
#     '8738900105',
#     'VTB',
#     '100879046497',
#     N'Không',
#     NULL,
#     NULL,
#     NULL,
#     NULL,
#     NULL,
#     NULL,
#     NULL,
#     NULL,
#     NULL,
#     NULL,
#     NULL,
#     N'Mẹ',
#     '0984143394',
#     N'HĐ xác định thời hạn lần 1',
#     '2024-02-09',
#     '2025-02-08',
#     N'Công nhân may công nghiệp ',
#     'K',
#     'W1',
#     'NT1',
#     '1PDN',
#     'Công nhân',
#     '1P02',
#     '1SEW2', 
#     '12S09',
#     'Incentive Worker',
#     'Sewer',
#     'SWR',
#     ' Skilled worker ',
#     NULL,
#     NULL,
#     NULL,
#     '2024-01-11',
#     NULL,
#     N'Đang làm việc',
#     '2024-01-11',
#     '1',
#     NULL,
#     NULL,
#     '2024-02-09',
#     '2025-02-08',
#     NULL,
#     NULL,
#     NULL,
#     'N',
#     NULL
# )""")

# # Lấy dữ liệu từ truy vấn
# # rows = cursor.fetchall()
# # for row in rows:
# #     print(row)

# # Đóng kết nối
# conn.commit()
# cursor.close()
# conn.close()

from openpyxl import Workbook
from openpyxl.styles import NamedStyle

# Tạo một workbook và worksheet mới
wb = Workbook()
ws = wb.active

# Định dạng ô như là văn bản
text_format = NamedStyle(name="text_format", number_format="@")
wb.add_named_style(text_format)

# Dữ liệu với số 0 đầu tiên
data = ["012345", "001234", "000123"]

# Thêm dữ liệu vào worksheet
for row_num, value in enumerate(data, 1):
    cell = ws.cell(row=row_num, column=1, value=value)
    cell.style = text_format

# Lưu file
wb.save("data.xlsx")