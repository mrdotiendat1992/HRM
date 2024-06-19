HR_DB = r'DRIVER={SQL Server};SERVER=172.16.60.100;DATABASE=HR;UID=huynguyen;PWD=Namthuan@123'
MITA_DB = r'DRIVER={SQL Server};SERVER=10.0.0.252\SQLEXPRESS;DATABASE=MITACOSQL;UID=sa;PWD=Namthuan1'

# -*- coding: utf-8 -*-
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
import time
import schedule
import pyodbc
from datetime import datetime, timedelta
# from datetime import datetime, timedelta
# from office365.runtime.auth.user_credential  import UserCredential
# from office365.sharepoint.client_context import ClientContext
# from office365.sharepoint.listitems.caml.query import CamlQuery


# def daydulieuloithe_lensharepoint():

#     # Thông tin kết nối
#     site_url = 'https://namthuanvn-my.sharepoint.com/personal/huy_nguyen_namthuan_vn'
#     username = 'huy.nguyen@namthuan.vn'
#     password = 'Namthuan@2024#'
#     list_title = "CBCNV_loi_the"

#     # Kết nối tới SharePoint
#     # Authentication
#     print("Connect to sharepoint ... ")
#     credentials = UserCredential(username, password)
#     ctx = ClientContext(site_url).with_credentials(credentials)
#     sp_list = ctx.web.lists.get_by_title(list_title)
#     print(f"Connect to {list_title} ... ")
#     # Lấy tất cả các mục trong danh sách
#     items = sp_list.get_items()
#     ctx.load(items)
#     ctx.execute_query()
    
#     print("Delete old data ... ")
#     # Duyệt qua các mục và xóa từng mục
#     for item in items:
#         try:
#             item.delete_object()
#             ctx.execute_query()
#         except Exception as ex:
#             print(ex)
#             print("Error deleting item. Skipping.")
            
#     print("Connect to Database ...")
#     conn = pyodbc.connect(HR_DB)
#     cursor = conn.cursor()
#     data=[]
#     rows = cursor.execute("SELECT * FROM HR.dbo.Danh_sach_CBCNV_Email").fetchall()
#     print(f"Found {len(rows)} rows, put to sharepoint ...")
#     for row in rows:
#         data.append({
#             "MST": f"{row[0]}",
#             "Ho_ten": row[1],
#             "Chuc_vu": row[2],
#             "Chuyen": row[3],
#             "Bo_phan": row[4],
#             "NgayCham": row[5],
#             "Email_thu_ky": row[6],
#             "Email_truong_bo_phan": row[7],
#         })
#         # break
#     for d in data:
#         sp_list.add_item(d)
#         ctx.execute_query()

#     print("All items have been added.")

def cao_xinnghiphep():
    conn = pyodbc.connect('Driver=SQL Server; Server=172.16.60.100; Database=HR; UID=IT; PWD=Namthuan@123;')
    cursor = conn.cursor()

    # Thông tin để đăng nhập SharePoint
    authcookie = Office365('https://namthuanvn-my.sharepoint.com', username='huy.nguyen@namthuan.vn', password='Namthuan@2024#').GetCookies()
    site = Site('https://namthuanvn-my.sharepoint.com/personal/huy_nguyen_namthuan_vn', version=Version.v365, authcookie=authcookie)

    # Truy cập list
    sp_list = site.List('Form xin nghỉ phép')
    data = sp_list.GetListItems('All Items')
    
    for row in data:
        nhamay = f"N'{row['Nha_may']}'"
        mst = f"'{row['MST'].replace('_NT1', '')}'"
        hoten = f"N'{row['Ho_ten']}'"
        chucvu = f"N'{row['Chuc_vu']}'"
        line = f"N'{row['Line']}'"
        bophan = f"N'{row['Bo_phan']}'"
        ngaynghi = f"'{row['Ngay_nghi_phep'].date()}'"
        tongsophut = f"'{int(row['Tong_so_phut'])}'"
        lido = f"N'{row['Ly_do']}'"
        trangthai = f"N'{row['Trạng thái']}'"
        
        query = f"INSERT INTO dbo.Xin_nghi_phep VALUES({nhamay}, {mst}, {hoten}, {chucvu}, {line}, {bophan}, {ngaynghi}, {tongsophut}, {lido}, {trangthai})"
        print(query)
        cursor.execute(query)

    conn.commit()
    conn.close()

def laydulieuinout_tumita():
    data = []
    conn = pyodbc.connect(MITA_DB)
    cursor = conn.cursor()
    ngaykiemtra = datetime.strftime(datetime.now().date() - timedelta(days=1), '%Y-%m-%d')
    ngaytruockiemtra = datetime.strftime(datetime.now().date() - timedelta(days=2), '%Y-%m-%d')
    print(f"Get data on {ngaykiemtra} from 10.0.0.252/SQLEXPRESS/MITACOSQL/CheckInOut ...")
    cursor.execute(f"SELECT * FROM CheckInOut WHERE NgayCham = '{ngaykiemtra}' OR NgayCham = '{ngaytruockiemtra}'")
    rows = cursor.fetchall()
    print(r"Get data success ...")
    for row in rows:
        data.append(
            ['NT1',
            row[1],
            row[2],
            row[3],
            row[-1]]
        )
    return data
    
def xoa_inout_trung_tuhr():
    conn = pyodbc.connect(HR_DB)
    cursor = conn.cursor()
    ngaykiemtra = datetime.strftime(datetime.now().date() - timedelta(days=1), '%Y-%m-%d')
    cursor.execute(f"SELECT * FROM Check_In_Out WHERE NgayCham = '{ngaykiemtra}'")
    rows = cursor.fetchall()
    if rows:
        print(f"Delete data on {ngaykiemtra} from 172.16.60.100/HR/Check_In_Out ... ")
        for row in rows:
            cursor.execute(f"DELETE FROM Check_In_Out WHERE MaChamCong = {row[1]} AND NgayCham = '{ngaykiemtra}'")
        print("Delete success")
        conn.commit()

def duadulieuinout_lenhr():
    xoa_inout_trung_tuhr()
    data = laydulieuinout_tumita()
    conn = pyodbc.connect(HR_DB)
    cursor = conn.cursor()
    print(f"Insert data to 172.16.60.100/HR/Check_In_Out ... ")
    for row in data:
        cursor.execute('INSERT INTO Check_In_Out VALUES(?,?,?,?,?)', row[0], row[1], row[2], row[3], row[4])
    conn.commit()
    print("Insert success")

def cao_diemdanhbu():
    conn = pyodbc.connect(HR_DB)
    cursor = conn.cursor()

    # Thông tin để đăng nhập SharePoint
    authcookie = Office365('https://namthuanvn-my.sharepoint.com', username='huy.nguyen@namthuan.vn', password='Namthuan@2024#').GetCookies()
    site = Site('https://namthuanvn-my.sharepoint.com/personal/huy_nguyen_namthuan_vn', version=Version.v365, authcookie=authcookie)

    # Truy cập list
    sp_list = site.List('Form đăng ký điểm danh bù')
    data = sp_list.GetListItems('All Items')
    
    for row in data:
        nhamay = f"N'{row['Nhà máy']}'"
        mst = f"'{row['Mã số thẻ']}'"
        hoten = f"N'{row['Họ tên']}'"
        chucvu = f"N'{row['Chức vụ']}'"
        line = f"N'{row['Tổ']}'"
        bophan = f"N'{row['Bộ phận']}'"
        loaidiemdanh = f"N'{row['Loại điểm danh']}'"
        ngaydiemdanh = f"'{row['Ngày điểm danh bù'].date()}'"
        time_str = row['Giờ điểm danh']
        giodiemdanh = f"'{datetime.strptime(time_str, "%H:%M").time()}'"
        lido = f"N'{row['Lý do']}'"
        trangthai = f"N'{row['Trạng thái']}'"
        
        query = f"INSERT INTO dbo.Diem_danh_bu VALUES({nhamay}, {mst}, {hoten}, {chucvu}, {line}, {bophan}, {loaidiemdanh}, {ngaydiemdanh}, {giodiemdanh}, {lido}, {trangthai})"
        if row['Trạng thái'] in ["Đã phê duyệt", "Đã điểm tra"]:
            cursor.execute(query)

    conn.commit()
    conn.close()

def cao_tuyendung():
    conn = pyodbc.connect(HR_DB)
    cursor = conn.cursor()

    # Thông tin để đăng nhập SharePoint
    authcookie = Office365('https://namthuanvn-my.sharepoint.com', username='huy.nguyen@namthuan.vn', password='Namthuan@2024#').GetCookies()
    site = Site('https://namthuanvn-my.sharepoint.com/personal/huy_nguyen_namthuan_vn', version=Version.v365, authcookie=authcookie)

    # Truy cập list
    sp_list = site.List('Form điền thông tin cá nhân')
    data = sp_list.GetListItems('All Items')
    print(f"Found {len(data)} records")
    print("Insert data to 172.16.60.100/HR/Dang_ky_thong_tin ... ")
    for row in data:
        nhamay = f"N'{row['Nhà máy']}'" if "Nhà máy" in row else "NULL"
        vitri = f"N'{row['Vị trí ứng tuyển']}'" if "Vị trí ứng tuyển" in row else "NULL"
        hoten = f"N'{row['Họ tên']}'" if "Họ tên" in row else "NULL"
        sdt = f"N'{row['Số điện thoại']}'" if "Số điện thoại" in row else "NULL"
        cccd = f"N'{row['Số CCCD']}'" if "Số CCCD" in row else "NULL"
        dantoc = f"N'{row['Dân tộc']}'" if "Dân tộc" in row else "NULL"
        tongiao = f"N'{row['Tôn giáo']}'" if "Tôn giáo" in row else "NULL"
        quoctich = f"N'{row['Quốc tịch']}'" if "Quốc tịch" in row else "NULL"
        trinhdo = f"N'{row['Trình độ học vấn']}'" if "Trình độ học vấn" in row else "NULL"
        noisinh = f"N'{row['Nơi sinh']}'" if "Nơi sinh" in row else "NULL"
        tamtru = f"N'{row['Địa chỉ tạm trú']}'" if "Địa chỉ tạm trú" in row else "NULL"
        nganhang = f"N'{row['Bạn có tài khoản ngân hàng VP Bank hoặc VietinBank chưa?']}'"if "Bạn có tài khoản ngân hàng VP Bank hoặc VietinBank chưa?" in row else "NULL"
        sotaikhoan = f"N'{row['Số tài khoản ngân hàng VP Bank hoặc VietinBank (nếu có)']}'" if "Số tài khoản ngân hàng VP Bank hoặc VietinBank (nếu có)" in row else "NULL"
        nguoithan = f"N'{row['Tên người thân']}'" if "Tên người thân" in row else "NULL" 
        sdtnguoithan = f"N'{row['SĐT người thân']}'" if "SĐT người thân" in row else "NULL"
        kenhtuyendung = f"N'{row['Bạn biết thông tin tuyển dụng qua kênh nào']}'" if "Bạn biết thông tin tuyển dụng qua kênh nào" in row else "NULL"
        kinhnghiem = f"N'{row['Kinh nghiệm làm việc']}'" if "Kinh nghiệm làm việc" in row else "NULL"
        mucluong = f"N'{row['Mức lương mong muốn']}'" if "Mức lương mong muốn" in row else "NULL"
        thoigiannhanviec = f"'{row['Thời gian có thể nhận việc'].date()}'" if "Thời gian có thể nhận việc" in row else "NULL"
        connho = f"N'{row['Bạn có con nhỏ dưới 6 tuổi không?']}'" if "Bạn có con nhỏ dưới 6 tuổi không?" in row else "NULL"
        ngaygui = f"'{row['Ngày gửi thông tin'].date()}'" if "Ngày gửi thông tin" in row else "NULL"
        sobhxh = f"N'{row['Số bảo hiểm xã hội (nếu có)']}'" if "Số bảo hiểm xã hội (nếu có)" in row else "NULL"
        masothue = f"N'{row['Mã số thuế cá nhân (nếu có)']}'" if "Mã số thuế cá nhân (nếu có)" in row else "NULL"
        tencon1 = f"N'{row['Họ tên con 1']}'" if "Họ tên con 1" in row else "NULL"
        tencon2 = f"N'{row['Họ tên con 2']}'" if "Họ tên con 2" in row else "NULL"
        tencon3 = f"N'{row['Họ tên con 3']}'" if "Họ tên con 3" in row else "NULL"
        tencon4 = f"N'{row['Họ tên con 4']}'" if "Họ tên con 4" in row else "NULL"
        tencon5 = f"N'{row['Họ tên con 5']}'" if "Họ tên con 5" in row else "NULL"
        ngaysinhcon1 = f"N'{row['Ngày sinh của con 1']}'" if "Ngày sinh của con 1" in row else "NULL"
        ngaysinhcon2 = f"N'{row['Ngày sinh của con 2']}'" if "Ngày sinh của con 2" in row else "NULL"
        ngaysinhcon3 = f"N'{row['Ngày sinh của con 3']}'" if "Ngày sinh của con 3" in row else "NULL"
        ngaysinhcon4 = f"N'{row['Ngày sinh của con 4']}'" if "Ngày sinh của con 4" in row else "NULL"
        ngaysinhcon5 = f"N'{row['Ngày sinh của con 5']}'" if "Ngày sinh của con 5" in row else "NULL"  
        
        query = f"INSERT INTO HR.dbo.Dang_ky_thong_tin VALUES ({nhamay},{vitri},{hoten},{sdt},{cccd},{dantoc},{tongiao},{quoctich},{trinhdo},{noisinh},{tamtru},{sobhxh},{masothue},{nganhang},{sotaikhoan},{nguoithan},{sdtnguoithan},{kenhtuyendung},{kinhnghiem},{mucluong},{thoigiannhanviec},{connho},{tencon1},{ngaysinhcon1},{tencon2},{ngaysinhcon2},{tencon3},{ngaysinhcon3},{tencon4},{ngaysinhcon4},{tencon5},{ngaysinhcon5},{ngaygui},NULL,NULL,NULL,NULL,NULL);"
        print(query)
        cursor.execute(query)
        conn.commit()
    conn.close()
    print("Insert success")
    
if __name__ == "__main__":
    cao_tuyendung()