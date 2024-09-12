# -*- encoding: utf-8 -*-

from const import *

app = Flask(__name__)

f12 = True    

# Cấu hình kết nối SQL Server
params = urllib.parse.quote_plus(
                "DRIVER={ODBC Driver 17 for SQL Server};"
                "SERVER=172.16.60.100;"
                "DATABASE=HR;"
                "UID=huynguyen;"
                "PWD=Namthuan@123;"
            )
app.config["SQLALCHEMY_DATABASE_URI"] = f"mssql+pyodbc:///?odbc_connect={params}"

# Cấu hình kết nối SQL Server
# params = urllib.parse.quote_plus(
#             "DRIVER={ODBC Driver 17 for SQL Server};"
#             "SERVER=DESKTOP-G635SF6;"
#             "DATABASE=HR;"
#             "Trusted_Connection=yes;"
#         )
# app.config["SQLALCHEMY_DATABASE_URI"] = f"mssql+pyodbc:///?odbc_connect={params}"
    
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config["SECRET_KEY"] = "hrm_system_NT"

db = SQLAlchemy(app)

handler = RotatingFileHandler('app.log', maxBytes=10000, backupCount=1, encoding='utf-8')
handler.setLevel(logging.INFO)
formatter = logging.Formatter(
    '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
)
handler.setFormatter(formatter)
app.logger.addHandler(handler)
app.logger.setLevel(logging.INFO)

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

class Nhanvien(UserMixin, db.Model):
    __tablename__ = 'Nhanvien'
    id = db.Column(db.Integer, primary_key=True)
    macongty = db.Column(db.String(10), nullable=False)
    masothe = db.Column(db.Integer, nullable=False)
    hoten = db.Column(db.Unicode(50), nullable=False)
    phongban = db.Column(db.String(10), nullable=False)
    capbac = db.Column(db.String(10), nullable=False)
    phanquyen = db.Column(db.String(10), nullable=False)
    matkhau = db.Column(db.String(10), nullable=False)

    def __repr__(self):
        return f"<User {self.hoten}>"
    
@login_manager.user_loader
def load_user(user_id):
    return Nhanvien.query.get(int(user_id))
  
def doimatkhautaikhoan(macongty,mst,matkhau):
    try:
        current_user.matkhau = matkhau
        db.session.commit()
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE Danh_sach_CBCNV SET Mat_khau='{matkhau}' WHERE Factory='{macongty}' AND The_cham_cong='{mst}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
        print("Đổi mật khẩu thành công !!!")
        return True        
    except Exception as e:
        print(f"Doi mat khau that bai {e} !!!")
        print(e)
        return False

def checkformatmst(mst):
    "Thêm số 0 đằng trước mã số thẻ nếu là mã số thẻ ở Nghệ An"
    if current_user.macongty == 'NT2':
        maxwidth = 5
        soluong = maxwidth-len(str(mst))
        if soluong >0:
            return '0'*soluong+str(mst)
    return mst
    
def laydanhsachsaphethanhopdong():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM Sap_het_han_HDLD WHERE Factory='{current_user.macongty}' ORDER BY Ngay_het_han_HD ASC, Department ASC, Line ASC, CAST(MST AS INT) ASC"
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(f"Lay danh sach sap het han hop dong that bai: {e}")
        return []

def laycatheochuyen(chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT CA FROM GOI_Y_CA WHERE LINE = '{chuyen}'"
        ##
        row = cursor.execute(query).fetchone()
        conn.close()
        if not row:
            if current_user.macongty=="NT2":
                return "A4-01"
            else:
                return "A1-02"
        else:
            return row[0]
    except Exception as e:
        print(f"Loi kiem tra ca theo chuyen {e} !!!")
        if current_user.macongty=="NT2":
            return "A4-01"
        else:
            return "A1-02"
  
def dieuchuyennhansu(mst,
                    loaidieuchuyen,
                    vitricu,
                    vitrimoi,
                    chuyencu,
                    chuyenmoi,
                    gradecodecu,
                    gradecodemoi,
                    sectioncodecu,
                    sectioncodemoi,
                    hccategorycu,
                    hccategorymoi,
                    departmentcu,
                    departmentmoi,
                    sectiondescriptioncu,
                    sectiondescriptionmoi,
                    employeetypecu,
                    employeetypemoi,
                    positioncodedescriptioncu,
                    positioncodedescriptionmoi,
                    positioncodecu,
                    positioncodemoi,
                    vitriencu,
                    vitrienmoi,
                    ngaydieuchuyen,
                    ghichu,
                    khongdoica
                   ):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        ghichu = "" if str(ghichu)=='nan' else ghichu
        query1 = f"INSERT INTO HR.dbo.Lich_su_cong_tac VALUES ('{current_user.macongty}','{mst}','{chuyencu}',N'{vitricu}','{chuyenmoi}',N'{vitrimoi}',N'{loaidieuchuyen}','{ngaydieuchuyen}',N'{ghichu}','{gradecodecu}','{gradecodemoi}','{hccategorycu}','{hccategorymoi}',GETDATE())"
        try:
            cursor.execute(query1)
            conn.commit()
        except Exception as e:
            return {
                "ketqua": False,
                "lido":e,
                "query":query1
            }
        
        if not khongdoica:
            camoi = laycatheochuyen(chuyenmoi)
            query3 = f"""
            UPDATE HR.dbo.Dang_ky_ca_lam_viec SET Den_ngay = '{datetime.strptime(str(ngaydieuchuyen)[:10], '%Y-%m-%d') - timedelta(days=1)}'  WHERE MST = '{int(mst)}' AND Factory = '{current_user.macongty}' AND Den_ngay='2054-12-31'
            INSERT INTO HR.dbo.Dang_ky_ca_lam_viec VALUES ('{int(mst)}','{current_user.macongty}','{ngaydieuchuyen}','2054-12-31','{camoi}')
            """
            try:
                cursor.execute(query3)
                conn.commit()
            except Exception as e:
                return {
                    "ketqua": False,
                    "lido":e,
                    "query":query3
                    }
            conn.close()
        return {"ketqua":True}
    except Exception as e:
        return {
                "ketqua": False,
                "lido":e,
                "query":""
                }
    
def laydanhsachca(mst):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT Ca,Tu_ngay,Den_ngay FROM Dang_ky_ca_lam_viec WHERE MST = '{mst}' AND Factory = '{current_user.macongty}' ORDER BY Tu_ngay DESC"
        cursor.execute(query)
        rows = cursor.fetchall()
        conn.close()
        return rows
    except Exception as e:
            print(e)
            return
    
def dichuyennghiviec(mst,
                    vitricu,
                    chuyencu,
                    gradecodecu,
                    hccategorycu,
                    ngaydieuchuyen,
                    ghichu
                   ):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"""
        INSERT INTO HR.dbo.Lich_su_cong_tac 
        VALUES ('{current_user.macongty}','{mst}',N'{chuyencu}',N'{vitricu}',
        NULL,NULL,N'Nghỉ việc','{ngaydieuchuyen}',N'{ghichu}','{gradecodecu}',
        NULL,'{hccategorycu}',NULL,GETDATE())
            """
        try:
            cursor.execute(query)
            conn.commit()
            conn.close()
            return {"ketqua":True} 
        except Exception as e:
            return {
                "ketqua": False,
                "lido":e,
                "query":query
                }
    except Exception as e:
        return {
                "ketqua": False,
                "lido":e,
                "query":""
                }

def dichuyennghithaisan(mst,
                        vitricu,
                        chuyencu,
                        gradecodecu,
                        hccategorycu,
                        ngaydieuchuyen
                            ):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"""
            INSERT INTO HR.dbo.Lich_su_cong_tac VALUES ('{current_user.macongty}','{mst}','{chuyencu}',N'{vitricu}',NULL,NULL,N'Nghỉ thai sản','{ngaydieuchuyen}',NULL,'{gradecodecu}',NULL,'{hccategorycu}',NULL,GETDATE())
            """    
        try:
            cursor.execute(query)
            conn.commit()
            conn.close()
            return {"ketqua":True} 
        except Exception as e:
            return {
                "ketqua": False,
                "lido":e,
                "query":query
                }
    except Exception as e:
        return {
                "ketqua": False,
                "lido":e,
                "query":""
                }

def dichuyenthaisandilamlai(mst,
                            vitricu,
                            vitrimoi,
                            chuyencu,
                            chuyenmoi,
                            gradecodecu,
                            gradecodemoi,
                            sectioncodecu,
                            sectioncodemoi,
                            hccategorycu,
                            hccategorymoi,
                            departmentcu,
                            departmentmoi,
                            sectiondescriptioncu,
                            sectiondescriptionmoi,
                            employeetypecu,
                            employeetypemoi,
                            positioncodedescriptioncu,
                            positioncodedescriptionmoi,
                            positioncodecu,
                            positioncodemoi,
                            vitriencu,
                            vitrienmoi,
                            ngaydieuchuyen
                            ):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"""
            INSERT INTO HR.dbo.Lich_su_cong_tac VALUES ('{current_user.macongty}','{mst}','{chuyencu}',N'{vitricu}','{chuyenmoi}',N'{vitrimoi}',N'Thai sản đi làm lại','{ngaydieuchuyen}',NULL,'{gradecodecu}','{gradecodemoi}','{hccategorycu}','{hccategorymoi}',GETDATE())
            """ 
        try:   
            cursor.execute(query)
            conn.commit()
            conn.close()
            return {"ketqua":True} 
        except Exception as e:
            return {
                "ketqua": False,
                "lido":e,
                "query":query
                }
    except Exception as e:
        return {
                "ketqua": False,
                "lido":e,
                "query":""
                }
    
def inhopdongtheomau(macongty,masothe,hoten,gioitinh,ngaysinh,thuongtru,tamtru,cccd,ngaycapcccd,capbac,loaihopdong,chucdanh,phongban,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc):   
    
    try:
        sodienthoai = lay_sodienthoai_theo_mst(masothe)
        ngaylamhopdong = ngaybatdau.split("/")[0]
        thanglamhopdong = ngaybatdau.split("/")[1]
        namlamhopdong = ngaybatdau.split("/")[2]
        ngayketthuchopdong = ngayketthuc.split("/")[0] if ngayketthuc else None
        thangketthuchopdong = ngayketthuc.split("/")[1] if ngayketthuc else None 
        namketthuchopdong = ngayketthuc.split("/")[2] if ngayketthuc else None
        if loaihopdong == "Hợp đồng thử việc":
            if macongty == "NT1":
                if capbac in ["O2","O1","M3","M2","M1"]:
                    try:
                        songaythuviec = (datetime.strptime(ngayketthuc,"%d/%m/%Y")-datetime.strptime(ngaybatdau,"%d/%m/%Y")).days+1
                        workbook = openpyxl.load_workbook(FILE_MAU_HDTV_NT1_O2_TROLEN)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: PC/{masothe}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = hoten.upper()
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['B21'] = f"Số CCCD: {cccd}"
                        sheet['L21'] = ngaycapcccd
                        sheet['E22'] = sodienthoai
                        sheet['H25'] = f"Thử việc {songaythuviec} ngày"
                        sheet['B26'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G32'] = chucdanh
                        sheet['G43'] = f"{int(luongcoban):,} VNĐ/tháng"   
                        if phucap > 0:
                            sheet['G44'] = f"{phucap} VNĐ/tháng"
                        else:
                            sheet['G44'] = "Không"
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDTV_TO2_{masothe}{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        print(e)
                        return None
                else:
                    try:
                        songaythuviec = (datetime.strptime(ngayketthuc,"%d/%m/%Y")-datetime.strptime(ngaybatdau,"%d/%m/%Y")).days+1
                        workbook = openpyxl.load_workbook(FILE_MAU_HDTV_NT1_DUOI_O2)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: PC/{masothe}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = hoten.upper()
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['D21'] = cccd
                        sheet['L21'] = ngaycapcccd
                        sheet['E22'] = sodienthoai
                        sheet['H25'] = f"Thử việc {songaythuviec} ngày"
                        sheet['B26'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G32'] = chucdanh
                        sheet['G43'] = f"{int(luongcoban):,} VNĐ/tháng"  
                        if phucap > 0:
                            sheet['G44'] = f"{phucap} VNĐ/tháng"
                        else:
                            sheet['G44'] = "Không" 
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDTV_DO2_{masothe}{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        print(e)
                        return None
            elif macongty == "NT2":
                try:
                    songaythuviec = (datetime.strptime(ngayketthuc,"%d/%m/%Y")-datetime.strptime(ngaybatdau,"%d/%m/%Y")).days+1
                    workbook = openpyxl.load_workbook(FILE_MAU_HDTV_NT2)
                    sheet = workbook.active
                    sheet['E4'] = f'Số: PC/{masothe}'
                    sheet['M4'] = f'Nghệ An, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                    sheet['D19'] = hoten.upper()
                    sheet['E21'] = ngaysinh
                    sheet['E20'] = gioitinh
                    sheet['F22'] = thuongtru
                    sheet['F23'] = tamtru
                    sheet['D24'] = f"{cccd}"
                    sheet['L24'] = ngaycapcccd
                    sheet['H27'] = f"{songaythuviec} ngày"
                    sheet['B28'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                    sheet['G46'] = f"{int(luongcoban):,} VNĐ/tháng"    
                    sheet['F33'] = chucdanh  
                    sheet['F34'] = phongban
                    sheet['F35'] = f"{masothe}"
                    if phucap > 0 :
                        sheet['G47'] = f"{phucap} VNĐ/tháng"
                    else:
                        sheet['G47'] = "Không" 
                    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                    filepath = os.path.join(FOLDER_XUAT, f'NT2_HDTV_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                    workbook.save(filepath)
                    return filepath
                except Exception as e:
                    print(e)
                    return None
        elif loaihopdong == "Hợp đồng có thời hạn 28 ngày":
            if macongty == "NT1":
                if capbac in ["O2","O1","M3","M2","M1"]:
                    try:
                        workbook = openpyxl.load_workbook(FILE_MAU_HDNH_NT1_O2_TROLEN)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: LC28D/{masothe}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = hoten.upper()
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['B21'] = f"Số CCCD:{cccd}"
                        sheet['L21'] = ngaycapcccd
                        sheet['B25'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G28'] = chucdanh
                        sheet['G38'] = f"{luongcoban} VNĐ/tháng"
                        if phucap > 0 :
                            sheet['G39'] = f"{phucap} VNĐ/tháng"
                        else:
                            sheet['G39'] = "Không"
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDNH_TO2_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        print(e)
                        return None
                else:
                    try:
                        workbook = openpyxl.load_workbook(FILE_MAU_HDNH_NT1_DUOI_O2)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: LC28D/{masothe}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = hoten.upper()
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['B21'] = f"Số CCCD: {cccd}"
                        sheet['L21'] = ngaycapcccd
                        sheet['B25'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G28'] = f"{chucdanh}"
                        sheet['G38'] = f"{int(luongcoban):,} VNĐ/tháng"
                        if phucap > 0 :
                            sheet['G39'] = f"{phucap} VNĐ/tháng"
                        else:
                            sheet['G39'] = "Không"
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDNH_DO2_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        print(e)
                        return None
            elif macongty == "NT2":
                try:
                    workbook = openpyxl.load_workbook(FILE_MAU_HDCTH_NT2)
                    sheet = workbook.active
                    sheet['E4'] = f'Số: LC28D/{masothe}'
                    sheet['M4'] = f'Nghệ An, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                    sheet['D19'] = hoten.upper()
                    sheet['E21'] = ngaysinh
                    sheet['E20'] = gioitinh
                    sheet['F22'] = thuongtru
                    sheet['F23'] = tamtru
                    sheet['D24'] = cccd
                    sheet['L24'] = ngaycapcccd
                    sheet['F32'] = phongban
                    sheet['F33'] = masothe
                    sheet['G43'] = f"{int(luongcoban):,} VNĐ/tháng"
                    if phucap > 0 :
                        sheet['G44'] = f"{phucap} VNĐ/tháng"
                    else:
                        sheet['G44'] = "Không" 
                    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                    filepath = os.path.join(FOLDER_XUAT, f'NT2_HDNH_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                    workbook.save(filepath)
                    return filepath
                except Exception as e:
                    print(e)
                    return None
        elif loaihopdong == "Hợp đồng có thời hạn 1 năm":
            if macongty == "NT1":
                if capbac in ["O2","O1","M3","M2","M1"]:
                    try:
                        workbook = openpyxl.load_workbook(FILE_MAU_HDCTH_NT1_O2_TROLEN)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: LC12/{masothe}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = hoten.upper()
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['B21'] = f"Số CCCD:{cccd}"
                        sheet['L21'] = ngaycapcccd
                        sheet['E22'] = sodienthoai
                        sheet['B26'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G29'] = f"{chucdanh}"
                        sheet['G39'] = f"{luongcoban} VNĐ/tháng"
                        if phucap > 0 :
                            sheet['G40'] = f"{phucap} VNĐ/tháng"
                        else:
                            sheet['G40'] = "Không"
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDCTH_TO3_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        print(e)
                        return None
                else:
                    try:
                        workbook = openpyxl.load_workbook(FILE_MAU_HDCTH_NT1_DUOI_O2)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: LC12/{masothe}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = hoten.upper()
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['B21'] = f"Số CCCD: {cccd}"
                        sheet['L21'] = ngaycapcccd
                        sheet['E22'] = sodienthoai
                        sheet['B26'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G29'] = f"{chucdanh}"
                        sheet['G39'] = f"{int(luongcoban):,} VNĐ/tháng"
                        if phucap > 0 :
                            sheet['G40'] = f"{phucap} VNĐ/tháng"
                        else:
                            sheet['G40'] = "Không"
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDCTH_DO3_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        print(e)
                        return None
            elif macongty == "NT2":
                try:
                    workbook = openpyxl.load_workbook(FILE_MAU_HDCTH_NT2)
                    sheet = workbook.active
                    sheet['E4'] = f'Số: LC12/{masothe}'
                    sheet['M4'] = f'Nghệ An, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                    sheet['D19'] = hoten.upper()
                    sheet['E21'] = ngaysinh
                    sheet['E20'] = gioitinh
                    sheet['F22'] = thuongtru
                    sheet['F23'] = tamtru
                    sheet['D24'] = cccd
                    sheet['L24'] = ngaycapcccd
                    sheet['F32'] = phongban
                    sheet['F33'] = masothe
                    sheet['G43'] = f"{int(luongcoban):,} VNĐ/tháng"
                    if phucap > 0 :
                        sheet['G44'] = f"{phucap} VNĐ/tháng"
                    else:
                        sheet['G44'] = "Không" 
                    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                    filepath = os.path.join(FOLDER_XUAT, f'NT2_HDCTH_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                    workbook.save(filepath)
                    return filepath
                except Exception as e:
                    print(e)
                    return None
        elif loaihopdong == "Hợp đồng vô thời hạn":
            if macongty == "NT1":
                if capbac in ["O2","O1","M3","M2","M1"]:
                    try:
                        workbook = openpyxl.load_workbook(FILE_MAU_HDVTH_NT1_O2_TROLEN)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: LC/{masothe}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = hoten.upper()
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['B21'] = f"Số CCCD: {cccd}"
                        sheet['L21'] = ngaycapcccd
                        sheet['E22'] = sodienthoai
                        sheet['B26'] = f"Kể từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}"
                        sheet['G39'] = f"{luongcoban} VNĐ/tháng"        
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT2_HDVTH_T03_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        print(e)
                        return None   
                else:
                    try:
                        workbook = openpyxl.load_workbook(FILE_MAU_HDVTH_NT1_DUOI_O2)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: LC/{masothe}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = hoten.upper()
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['B21'] = f"Số CCCD: {cccd}"
                        sheet['L21'] = ngaycapcccd
                        sheet['E22'] = sodienthoai
                        sheet['B26'] = f"Kể từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}"   
                        sheet['G29'] = f"{chucdanh}"
                        sheet['G39'] = f"{int(luongcoban):,} VNĐ/tháng"
                        if phucap > 0 :
                            sheet['G40'] = f"{phucap} VNĐ/tháng"
                        else:
                            sheet['G40'] = "Không"    
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT2_HDVTH_D03_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        
                        return filepath
                    except Exception as e:
                        print(e)
                        return None  
            if macongty == "NT2":
                try:
                    workbook = openpyxl.load_workbook(FILE_MAU_HDVTH_NT2)
                    sheet = workbook.active
                    sheet['E4'] = f'Số: LC/{masothe}'
                    sheet['M4'] = f'Nghệ An, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                    sheet['D19'] = hoten.upper()
                    sheet['E21'] = ngaysinh
                    sheet['E20'] = gioitinh
                    sheet['F22'] = thuongtru
                    sheet['F23'] = tamtru
                    sheet['D24'] = cccd
                    sheet['L24'] = ngaycapcccd
                    sheet['B28'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}"
                    sheet['G43'] = f"{int(luongcoban):,} VNĐ/tháng"  
                    if phucap > 0 :
                        sheet['G44'] = f"{phucap} VNĐ/tháng"
                    else:
                        sheet['G44'] = "Không"      
                    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                    filepath = os.path.join(FOLDER_XUAT, f'NT2_HDVTH_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                    workbook.save(filepath)
                    return filepath
                except Exception as e:
                    print(e)
                    return None  
    except Exception as e:
        print(e)
        return None
    
def inchamduthd(mst,
                ngaylamhopdong,
                thanglamhopdong,
                namlamhopdong,
                tennhanvien,
                chucvu,
                ngaynghi,
                ngaysinh,
                diachi,
                bophan,
                lydo):
    if current_user.macongty == "NT1":
        try:
            
            workbook = openpyxl.load_workbook(FILE_MAU_CDHD_NT1)
            sheet = workbook.active
            sheet['C4'] = f'{mst}'
            sheet['H5'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
            sheet['I12'] = f'{mst}'
            sheet['I17'] = f'{mst}'
            sheet['G16'] = tennhanvien
            sheet['C21'] = tennhanvien
            sheet['B26'] = tennhanvien
            sheet['D18'] = lydo
            sheet['D19'] = ngaynghi
            sheet['E22'] = ngaynghi
            sheet['D17'] = chucvu  
            thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
            filepath = os.path.join(FOLDER_XUAT, f'NT_CDHĐ_{mst}_{thoigian}.xlsx')
            workbook.save(filepath)
            
            return filepath
        except Exception as e:
            print(e)
            return None
    elif current_user.macongty == "NT2":
        try:
            thangnghi=str(ngaynghi[3:])
            workbook = openpyxl.load_workbook(FILE_MAU_CDHD_NT2)
            sheet = workbook.active
            sheet['A3'] = f'     Số: {mst}/QĐ-NTNA'
            sheet['G3'] = f'          Nghệ An, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}     '
            sheet['C17'] = tennhanvien
            sheet['C18'] = ngaysinh
            sheet['C19'] = diachi
            sheet['C20'] = mst
            sheet['C21'] = chucvu
            sheet['C22'] = bophan
            sheet['C23'] = ngaynghi 
            sheet['C24'] = lydo
            sheet['F25'] = thangnghi
            sheet['B26'] = f"Ông/Bà: {tennhanvien} có trách nhiệm bàn giao toàn bộ công việc, tài liệu (nếu có) và các giấy tờ liên quan cho phòng HCNS."
            sheet['B28'] = f"Quyết định có hiệu lực kể từ ngày {ngaynghi}, Phòng Hành chính nhân sự, Phòng Tài chính Kế toán, các Phòng/Bộ phận có liên quan và Ông/Bà {tennhanvien} chịu trách nhiệm thi hành quyết định này."
            thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
            filepath = os.path.join(FOLDER_XUAT, f'NT2_CDHĐ_{mst}_{thoigian}.xlsx')
            workbook.save(filepath)
            
            return filepath
        except Exception as e:
            print(e)
            return None
   
def laylichsucongtac(mst,hoten,ngay,kieudieuchuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query= f"""SELECT * FROM Lich_su_cong_tac_OK WHERE Nha_may = '{current_user.macongty}' """
        if mst:
            query += f"AND MST = '{mst}' "
        if ngay:
            query += f"AND Ngay_thuc_hien = '{ngay}' "
        if kieudieuchuyen:
            query += f"AND Phan_loai LIKE N'%{kieudieuchuyen}%' "
        if hoten:
            query += f"AND Ho_ten LIKE N'%{hoten}%' "
        query += "ORDER BY Ngay_thuc_hien DESC, CAST(MST AS INT) ASC"
        ##
        rows = cursor.execute(query)
        result = []
        for row in rows:
            result.append({
                "Nhà máy": row[12],
                "ID": row[13],
                "MST": row[1],
                "Họ tên": row[0],
                "Ngày chính thức": row[2] if row[2] else '',
                "Chuyền cũ": row[3],
                "Chuyền mới": row[5] if row[5] else '',
                "Vị trí cũ": row[8],
                "Vị trí mới": row[9] if row[9] else '',
                "Phòng ban cũ": row[4],
                "Phòng ban mới": row[6] if row[6] else '',
                "Phân loại": row[10],
                "Ngày thực hiện": row[7] if row[7] else '',
                "Ghi chú": row[11] if row[1] else ''
            })
        conn.commit()
        conn.close()
        return result
    except Exception as e:
        print(e)
        return []
                     
def laydanhsachlinetheovitri(vitri):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT DISTINCT Line FROM HR.dbo.HC_Name WHERE Factory = '{current_user.macongty}' AND Detail_job_title_VN = N'{vitri}'"
        
        rows = cursor.execute(query).fetchall()
        conn.close()
        result = []
        for row in rows:
            result.append(row[0])
        return result
    except Exception as e:
        print(e)
        return []
    
def lay_user(user):
    if user:
        return {
            "MST": user[0],
            "Thẻ chấm công": user[1],
            "Họ tên": user[2],
            "Số điện thoại": user[3],
            "Ngày sinh": datetime.strptime(user[4],"%Y-%m-%d").strftime("%d/%m/%Y") if user[4] else "",
            "Giới tính": user[5],
            "CCCD": user[6],
            "Ngày cấp CCCD": datetime.strptime(user[7],"%Y-%m-%d").strftime("%d/%m/%Y") if user[7] else "",
            "Nơi cấp": user[8],
            "CMT": user[9],
            "Thường trú": user[10],
            "Thôn xóm": user[11],
            "Phường xã": user[12],
            "Quận huyện": user[13],
            "Tỉnh thành phố": user[14],
            "Dân tộc": user[15],
            "Quốc tịch": user[16],
            "Tôn giáo": user[17],
            "Học vấn": user[18],
            "Nơi sinh": user[19],
            "Tạm trú": user[20],
            "Số BHXH": user[21],
            "Mã số thuế": user[22],
            "Ngân hàng": user[23],
            "Số tài khoản": user[24],
            "Con nhỏ": user[25],
            "Tên con 1": user[26],
            "Ngày sinh con 1": user[27],
            "Tên con 2": user[28],
            "Ngày sinh con 2": user[29],
            "Tên con 3": user[30],
            "Ngày sinh con 3": user[31],
            "Tên con 4": user[32],
            "Ngày sinh con 4": user[33],
            "Tên con 5": user[34],
            "Ngày sinh con 5": user[35],
            "Ảnh chân dung": user[36],
            "Người thân": user[37],
            "SĐT liên hệ": user[38],
            "Loại hợp đồng": user[39],
            "Ngày ký HĐ": datetime.strptime(user[40],"%Y-%m-%d").strftime("%d/%m/%Y") if user[40] else "",
            "Ngày hết hạn": datetime.strptime(user[41],"%Y-%m-%d").strftime("%d/%m/%Y") if user[41] else "",
            "Job title VN": user[42],
            "HC category": user[43],
            "Gradecode": user[44],
            "Factory": user[45],
            "Department": user[46],
            "Chức vụ": user[47],
            "Section code": user[48],
            "Section description": user[49],
            "Line": user[50],
            "Employee type": user[51],
            "Job title EN": user[52],
            "Position code": user[53],
            "Position description": user[54],
            "Lương cơ bản": user[55],
            "Phụ cấp": user[56],
            "Tiền phụ cấp": user[57],
            "Ngày vào": datetime.strptime(user[58],"%Y-%m-%d").strftime("%d/%m/%Y") if user[58] else "",
            "Ngày nghỉ": datetime.strptime(user[59],"%Y-%m-%d").strftime("%d/%m/%Y") if user[59] else "",
            "Trạng thái": user[60],
            "Ngày vào nối thâm niên": datetime.strptime(user[61],"%Y-%m-%d").strftime("%d/%m/%Y") if user[61] else "",
            "Ngày kí HĐ Thử việc": datetime.strptime(user[63],"%Y-%m-%d").strftime("%d/%m/%Y") if user[63] else "",
            "Ngày hết hạn HĐ Thử việc": datetime.strptime(user[64],"%Y-%m-%d").strftime("%d/%m/%Y") if user[64] else "",
            "Ngày kí HĐ xác định thời hạn lần 1": datetime.strptime(user[65],"%Y-%m-%d").strftime("%d/%m/%Y") if user[65] else "",
            "Ngày hết hạn HĐ xác định thời hạn lần 1": datetime.strptime(user[66],"%Y-%m-%d").strftime("%d/%m/%Y") if user[66] else "",
            "Ngày kí HĐ xác định thời hạn lần 2": datetime.strptime(user[67],"%Y-%m-%d").strftime("%d/%m/%Y") if user[67] else "",
            "Ngày hết hạn HĐ xác định thời hạn lần 2": datetime.strptime(user[68],"%Y-%m-%d").strftime("%d/%m/%Y") if user[68] else "",
            "Ngày kí HĐ không thời hạn": datetime.strptime(user[69],"%Y-%m-%d").strftime("%d/%m/%Y") if user[69] else "",
            "Ghi chú": user[71] if user[71] else ""
        }
    else:
        return {}

def laydanhsachuser(mst, hoten, sdt, cccd, gioitinh, vaotungay, vaodenngay, nghitungay, nghidenngay, phongban, trangthai, hccategory,chucvu, ghichu, chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Danh_sach_CBCNV WHERE Factory = '{current_user.macongty}'"
        if mst:
            query += f" AND MST = '{mst}'"
        if hoten:
            query += f" AND Ho_ten LIKE N'%{hoten}%'"
        if sdt:
            query += f" AND SDT = '{sdt}'"
        if cccd:
            query += f" AND CCCD = '{cccd}'"
        if gioitinh:
            query += f" AND Gioi_tinh LIKE N'%{gioitinh}%'"
        if vaotungay:
            query += f" AND Ngay_vao >= '{vaotungay}'"
        if vaodenngay:
            query += f" AND Ngay_vao <= '{vaodenngay}'"
        if nghitungay:
            query += f" AND Ngay_nghi >= '{nghitungay}'"
        if nghidenngay:
            query += f" AND Ngay_nghi <= '{nghidenngay}'"
        if phongban:
            query += f" AND Department = N'{phongban}'"
        if trangthai:
            query += f" AND Trang_thai_lam_viec LIKE N'%{trangthai}%'"
        if hccategory:
            query += f" AND Headcount_category = '{hccategory}'"
        if chucvu:
            query += f" AND Chuc_vu LIKE N'%{chucvu}%'"
        if ghichu:
            query += f" AND Ghi_chu LIKE N'%{ghichu}%'"
        if chuyen:
            query += f" AND Line = N'{chuyen}'"
        
        query += " ORDER BY CAST(mst AS INT) ASC"
        users = cursor.execute(query).fetchall()
        conn.close()
        result = []
        for user in users:
            result.append(lay_user(user))
        if not mst and not hoten and not sdt and not cccd and not gioitinh and not vaotungay and not vaodenngay and not nghitungay and not nghidenngay and not phongban and not trangthai and not hccategory and not chucvu and not ghichu and not chuyen:
            return []
        return result
    except Exception as e:
        flash(f"Lỗi khi lấy danh sách nhân viên: {e}")
        return []

def laydanhsachuserhientai():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Danh_sach_CBCNV WHERE Factory = '{current_user.macongty}' ORDER BY CAST(mst AS INT) ASC"
        users = cursor.execute(query).fetchall()
        conn.close()
        return [lay_user(user) for user in users]
    except Exception as e:
        flash(f"Lỗi khi lấy danh sách nhân viên hiện tại: {e}")
        return []
   
def laycacphongban():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT DISTINCT Department FROM HR.dbo.HC_Name WHERE Factory = '{current_user.macongty}'"
        
        cacphongban =  cursor.execute(query).fetchall()
        conn.close()
        result = []
        for x in cacphongban:
            result.append(x[0])
        return result
    except Exception as e:
        print(e)
        return []

def laycacto():
    try:    
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT DISTINCT Line FROM HR.dbo.HC_Name WHERE Factory = '{current_user.macongty}'"
        
        cacto =  cursor.execute(query).fetchall()
        conn.close()
        return [x[0] for x in cacto]
    except Exception as e:
        print(e)
        return []

def laycachccategory():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT DISTINCT HC_category FROM HR.dbo.HC_Name WHERE Factory = '{current_user.macongty}'"
        
        cachccategory =  cursor.execute(query).fetchall()
        conn.close()
        return [x[0] for x in cachccategory]
    except Exception as e:
        print(e)
        return []
    
def laydanhsachtheomst(mst):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Danh_sach_CBCNV WHERE MST = '{mst}' AND Factory = '{current_user.macongty}'"
        
        users = cursor.execute(query).fetchall()
        conn.close()
        result = []
        for user in users:
            result.append(lay_user(user))
        return result
    except Exception as e:
        print(e)
        return []
    
def laydanhsachtheothechamcong(mst):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Danh_sach_CBCNV WHERE The_cham_cong = '{mst}' AND Factory = '{current_user.macongty}'"
        
        users = cursor.execute(query).fetchall()
        conn.close()
        result = []
        for user in users:
            result.append(lay_user(user))
        return result
    except Exception as e:
        print(e)
        return []

def laydanhsachusercacongty(macongty):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Danh_sach_CBCNV WHERE Factory = '{macongty}' ORDER BY CAST(mst AS INT) ASC"
        
        users = cursor.execute(query).fetchall()
        conn.close()
        result = []
        for user in users:
            result.append(lay_user(user))
        return result
    except Exception as e:
        print(e)
        return []
    
def laydanhsachusertheophongban(phongban):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Danh_sach_CBCNV WHERE Department = '{phongban}' AND Factory = '{current_user.macongty}' ORDER BY CAST(mst AS INT) ASC"
        
        users = cursor.execute(query).fetchall()
        conn.close()
        result = []
        for user in users:
            result.append(lay_user(user))
        return result
    except Exception as e:
        print(e)
        return []
    
def laydanhsachusertheogioitinh(gioitinh):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Danh_sach_CBCNV WHERE Gioi_tinh = N'{gioitinh}' AND Factory = '{current_user.macongty}' ORDER BY CAST(mst AS INT) ASC"
        
        users = cursor.execute(query).fetchall()
        conn.close()
        result = []
        for user in users:
            result.append(lay_user(user))
        return result
    except Exception as e:
        print(e)
        return []

def laydanhsachusertheoline(line):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Danh_sach_CBCNV WHERE Line = '{line} 'AND Factory = '{current_user.macongty}' ORDER BY CAST(mst AS INT) ASC"
        
        users = cursor.execute(query).fetchall()
        conn.close()
        result = []
        for user in users:
            result.append(lay_user(user))
        return result
    except Exception as e:
        print(e)
        return []

def laydanhsachusertheostatus(status):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Danh_sach_CBCNV WHERE Trang_thai_lam_viec = N'{status}' AND Factory = '{current_user.macongty}' ORDER BY CAST(mst AS INT) ASC"
        
        users = cursor.execute(query).fetchall()
        conn.close()
        result = []
        for user in users:
            result.append(lay_user(user))
        return result
    except Exception as e:
        print(e)
        return []

def laycactrangthai():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT DISTINCT Trang_thai_lam_viec FROM HR.dbo.Danh_sach_CBCNV WHERE Factory = '{current_user.macongty}'"
        
        cactrangtha =  cursor.execute(query).fetchall()
        conn.close()
        return [x[0] for x in cactrangtha]
    except Exception as e:
        print(e)
        return []

def laycacvitri():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT DISTINCT Detail_job_title_VN FROM HR.dbo.HC_Name WHERE Factory = '{current_user.macongty}'"
        
        cacvitri =  cursor.execute(query).fetchall()
        conn.close()
        return [x[0] for x in cacvitri]
    except Exception as e:
        print(e)
        return []
    
def laycacca():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT DISTINCT Ten_ca FROM HR.dbo.Ca_lam_viec"
        
        cacca =  cursor.execute(query).fetchall()
        conn.close()
        return [x[0] for x in cacca]
    except Exception as e:
        print(e)
        return []

def layhcname(jobtitle,line):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query =  f"SELECT * FROM HR.dbo.HC_Name WHERE Detail_job_title_VN = N'{jobtitle}' AND Line = N'{line}' AND Factory = N'{current_user.macongty}'"
        
        result = cursor.execute(query).fetchone()
        conn.close()
        return result
    except Exception as e:
        print(e)
        return []

def laydanhsachdangkytuyendung(sdt=None, cccd=None, ngaygui=None):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Dang_ky_thong_tin_OK WHERE Nha_may = '{current_user.macongty}'"
        if sdt:
            query += f"AND Sdt LIKE '{sdt}'"
        if cccd:
            query += f"AND CCCD LIKE '{cccd}'"
        if ngaygui:
            query += f"AND Ngay_gui_thong_tin =  '{ngaygui}'"
            
        query+= " ORDER BY Ngay_gui_thong_tin DESC"
        
        rows =  cursor.execute(query).fetchall()
        conn.close()
        result = []
        for row in rows:
            result.append({
                "ID": row[39],
                "Nhà máy": row[0],
                "Vị trí tuyển dụng": row[1],
                "Họ tên": row[2].upper(),
                "Số điện thoại": row[3],
                "CCCD": row[4],
                "Dân tộc": row[5],
                "Tôn giáo": row[6],
                "Quốc tịch": row[7],
                "Học vấn": row[8],
                "Nơi sinh": row[9],
                "Tạm trú": row[10],
                "Số BHXH": row[11],
                "Mã số thuế": row[12],
                "Ngân hàng": row[13],
                "Số tài khoản": row[14],
                "Tên người thân": row[15],
                "SĐT người thân": row[16],
                "Kênh tuyển dụng": row[17],
                "Kinh nghiệm": row[18],
                "Mức lương": row[19],
                "Ngày nhận việc": datetime.strptime(row[20], '%Y-%m-%d').strftime("%d/%m/%Y") if (row[20] and "-" in row[20]) else "",
                "Có con nhỏ": row[21],
                "Tên con 1": row[22],
                "Ngày sinh con 1": row[23] if row[23] else "",
                "Tên con 2": row[24],
                "Ngày sinh con 2": row[25] if row[25] else "",
                "Tên con 3": row[26],
                "Ngày sinh con 3": row[27] if row[27] else "",
                "Tên con 4": row[28],
                "Ngày sinh con 4": row[29] if row[29] else "",
                "Tên con 5": row[30],
                "Ngày sinh con 5": row[31] if row[31] else "",
                "Ngày gửi": datetime.strptime(row[32], '%Y-%m-%d').strftime("%d/%m/%Y") if (row[32] and "-" in row[32]) else "",
                "Trạng thái": row[33],
                "Ngày cập nhật": datetime.strptime(row[34], '%Y-%m-%d').strftime("%d/%m/%Y") if (row[34] and "-" in row[34]) else "",
                "Ngày hẹn đi làm": datetime.strptime(row[35], '%Y-%m-%d').strftime("%d/%m/%Y") if (row[35] and "-" in row[35]) else "",
                "Hiệu suất": row[36],
                "Loại máy": row[37],
                "Ghi chú": row[38], 
                "Lưu hồ sơ": row[40]
            })
        return result
    except Exception as e:
        print(e)
        return []
    
def capnhattrangthaimoiungvien(sdt, trangthai, luuhoso):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        ngaythuchien = datetime.now().date()
        query = f"UPDATE HR.Dbo.Dang_ky_thong_tin SET Trang_thai = N'{trangthai}', Luu_ho_so='{luuhoso}', Ngay_cap_nhat = '{ngaythuchien}' WHERE Sdt = '{sdt}' AND Nha_may = N'{current_user.macongty}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        return False
    
def capnhatthongtinungvien(id,
                        sdt,
                        ngayhendilam,
                        hieusuat,
                        loaimay,
                        vitrituyendung,
                        hocvan,
                        diachi,
                        dantoc,
                        connho,
                        tencon1,
                        ngaysinhcon1,
                        tencon2,
                        ngaysinhcon2,
                        tencon3,
                        ngaysinhcon3,
                        tencon4,
                        ngaysinhcon4,
                        tencon5,
                        ngaysinhcon5,
                        nguoithan,
                        sdtnguoithan,
                        luuhoso,
                        ghichu,
                        cccd
                        ):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if not hieusuat.isdigit():
            hieusuat = 0 
        query = f"""
        UPDATE HR.Dbo.Dang_ky_thong_tin 
        SET 
        Sdt = '{sdt}',
        Ngay_hen_di_lam = '{ngayhendilam}',
        Hieu_suat = '{hieusuat}',
        Loai_may = N'{loaimay}',
        Vi_tri_ung_tuyen=N'{vitrituyendung}',
        Trinh_do=N'{hocvan}',
        Dia_chi_tam_tru=N'{diachi}', 
        Dan_toc = N'{dantoc}',
        Con_nho = N'{connho}',
        Ho_ten_con_1 = N'{tencon1}',
        Ngay_sinh_con_1 = '{ngaysinhcon1}',
        Ho_ten_con_2 = N'{tencon2}',
        Ngay_sinh_con_2 = '{ngaysinhcon2}',
        Ho_ten_con_3 = N'{tencon3}',
        Ngay_sinh_con_3 = '{ngaysinhcon3}',
        Ho_ten_con_4 = N'{tencon4}',
        Ngay_sinh_con_4 = '{ngaysinhcon4}',
        Ho_ten_con_5 = N'{tencon5}',
        Ngay_sinh_con_5 = '{ngaysinhcon5}',
        Nguoi_than = N'{nguoithan}',
        SDT_nguoi_than = '{sdtnguoithan}',
        Luu_ho_so = N'{luuhoso}',
        Ghi_chu = N'{ghichu}',
        CCCD = '{cccd}'
        WHERE 
        ID = '{id}' AND Nha_may = N'{current_user.macongty}'"""
        
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        return False
    
def themnhanvienmoi(nhanvienmoi):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"INSERT INTO HR.Dbo.Danh_sach_CBCNV VALUES {nhanvienmoi}"
        # 
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi insert nhan vien moi: {e}")
        return False

def xoadautrongten(s):
    s = re.sub(r'[àáạảãâầấậẩẫăằắặẳẵ]', 'a', s)
    s = re.sub(r'[ÀÁẠẢÃĂẰẮẶẲẴÂẦẤẬẨẪ]', 'A', s)
    s = re.sub(r'[èéẹẻẽêềếệểễ]', 'e', s)
    s = re.sub(r'[ÈÉẸẺẼÊỀẾỆỂỄ]', 'E', s)
    s = re.sub(r'[òóọỏõôồốộổỗơờớợởỡ]', 'o', s)
    s = re.sub(r'[ÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠ]', 'O', s)
    s = re.sub(r'[ìíịỉĩ]', 'i', s)
    s = re.sub(r'[ÌÍỊỈĨ]', 'I', s)
    s = re.sub(r'[ùúụủũưừứựửữ]', 'u', s)
    s = re.sub(r'[ƯỪỨỰỬỮÙÚỤỦŨ]', 'U', s)
    s = re.sub(r'[ỳýỵỷỹ]', 'y', s)
    s = re.sub(r'[ỲÝỴỶỸ]', 'Y', s)
    s = re.sub(r'[Đ]', 'D', s)
    s = re.sub(r'[đ]', 'd', s)
    return s
    
def xoanhanvien(MST):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query=f"DELETE FROM HR.Dbo.Danh_sach_CBCNV WHERE MST = '{MST}' AND Factory = N'{current_user.macongty}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
        return f"{MST} đã xoá thành công"
    except Exception as e:
        print(e)
        return f"{MST} đã xoá thất bại"
    
def laymasothemoi():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT TOP 1 MST FROM HR.dbo.Danh_sach_CBCNV WHERE Factory = '{current_user.macongty}' ORDER BY CAST(MST AS INT) DESC"
        
        result =  cursor.execute(query).fetchone()
        conn.close()
        if result:
            return result[0]
        return 0
    except Exception as e:
        print(e)
        return 0

def laydanhsachloithe(mst=None,chuyen=None, bophan=None, ngay=None, mstthuky=None):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if mstthuky:
            query = f"""
                    SELECT  *
                    FROM 
                        Danh_sach_loi_the_3
                    INNER JOIN 
                        Phan_quyen_thu_ky
                    ON
                        Danh_sach_loi_the_3.Nha_may= Phan_quyen_thu_ky.Nha_may and Danh_sach_loi_the_3.Chuyen_to=Phan_quyen_thu_ky.Chuyen_to
                    WHERE 
                        Phan_quyen_thu_ky.MST='{mstthuky}' and Trang_thai is null """
        else:
            query = f"SELECT * FROM HR.dbo.Danh_sach_loi_the_3 WHERE Nha_may = '{current_user.macongty}'"
            if mst:
                query += f"AND MST LIKE '%{mst}%' "
            if chuyen:
                query += f"AND Chuyen_to LIKE '%{chuyen}%' "
            if bophan:
                query += f"AND Bo_phan LIKE '%{bophan}%' "
            if ngay:
                query += f"AND Ngay = '{ngay}' "
            
            query += "ORDER BY CAST(MST AS INT) ASC, Ngay DESC"
        
        rows = cursor.execute(query).fetchall()
        conn.close()
        result = []
        for row in rows:
            result.append({
            "Nhà máy": row[0],
            "MST": row[1],
            "Họ tên": row[2],
            "Chức danh": row[3],
            "Chuyền tổ": row[4],
            "Bộ phận": row[5],
            "Cấp bậc": row[6],
            "Ngày": datetime.strptime(row[7], "%Y-%m-%d").strftime("%d/%m/%Y"),
            "Ca": row[8],
            "Giờ vào": row[9],
            "Giờ ra": row[10],
            "Phút HC": row[11],
            "Phút nghỉ phép": row[12],
            "Số phút thiếu": row[13],
            "Phép tồn": row[14],
            "Phút nghỉ không lương": row[15],
            "Phút nghỉ khác": row[16],
            "Trạng thái": row[18]
            })
        return result
    except Exception as e:
        print(e)
        return []

def laydanhsachchuyen():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        # ngaykiemtra= (datetime.now()-timedelta(days=1)).date()
        query = f"SELECT DISTINCT Line FROM HR.dbo.Danh_sach_CBCNV WHERE Factory = '{current_user.macongty}'"
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
            print(e)
            return []

def laydanhsachbophan():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        # ngaykiemtra= (datetime.now()-timedelta(days=1)).date()
        query = f"SELECT DISTINCT Department FROM HR.dbo.Danh_sach_CBCNV WHERE Factory = '{current_user.macongty}'"
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
            print(e)
            return []

def laydanhsachchamcong(mst=None, chuyen=None, phongban=None, tungay=None, denngay=None, phanloai=None):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Bang_cham_cong_tu_dong WHERE Nha_may = '{current_user.macongty}'"
        if mst: 
            query += f" AND MST = '{mst}'"
        if chuyen: 
            query += f" AND Chuyen_to = '{chuyen}'"
        if phongban:
            query += f" AND Bo_phan LIKE N'%{phongban}%'"
        if tungay:
            query += f" AND '{tungay}' <= Ngay"
        if denngay:
            query += f" AND Ngay <= '{denngay}'"
        if phanloai:
            query += f" AND Phan_loai LIKE N'%{phanloai}%'"
        query +=" ORDER BY Ngay DESC, Bo_phan ASC, Chuyen_to ASC, MST ASC"
        
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
            print(e)
            return []

def laydanhsachchamcongchot(mst=None, chuyen=None, phongban=None, tungay=None, denngay=None, phanloai=None):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Bang_cham_cong WHERE Nha_may = '{current_user.macongty}'"
        if mst: 
            query += f" AND MST LIKE '%{mst}%'"
        if chuyen: 
            query += f" AND Chuyen_to = '{chuyen}'"
        if phongban:
            query += f" AND Bo_phan LIKE '%{phongban}%'"
        if tungay:
            query += f" AND '{tungay}' <= Ngay"
        if denngay:
            query += f" AND Ngay <= '{denngay}'"
        if phanloai:
            query += f" AND Phan_loai LIKE N'%{phanloai}%'"
        query +=" ORDER BY Ngay DESC, Bo_phan ASC, Chuyen_to ASC, MST ASC"
        
        rows = cursor.execute(query).fetchall()
        conn.close()
        result = []
        for row in rows:
            result.append(row)
        return result
    except Exception as e:
        print(e)
        return []

def laydanhsachdiemdanhbu(mst=None,hoten=None,chucvu=None,chuyen=None,bophan=None,loaidiemdanh=None,ngaydiemdanh=None,lido=None,trangthai=None,mstquanly=None,mstthuky=None):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if mstthuky:
            query = f"""
            SELECT  DISTINCT Diem_danh_bu.*
            FROM 
                Diem_danh_bu 
            INNER JOIN 
                Phan_quyen_thu_ky
            ON
                Diem_danh_bu.Nha_may= Phan_quyen_thu_ky.Nha_may and Diem_danh_bu.Line=Phan_quyen_thu_ky.Chuyen_to
            WHERE 
                Diem_danh_bu.Trang_thai=N'Chờ kiểm tra' and Phan_quyen_thu_ky.MST='{mstthuky}' and Diem_danh_bu.Nha_may = '{current_user.macongty}' """
        else:
            if mstquanly:
                query = f"""
                SELECT DISTINCT Diem_danh_bu.*
                FROM 
                    Diem_danh_bu 
                INNER JOIN 
                    Phan_quyen_thu_ky
                ON
                    Diem_danh_bu.Nha_may= Phan_quyen_thu_ky.Nha_may and Diem_danh_bu.Line=Phan_quyen_thu_ky.Chuyen_to
                WHERE 
                    Diem_danh_bu.Trang_thai=N'Đã kiểm tra' and Phan_quyen_thu_ky.MST_QL='{mstquanly}' and Diem_danh_bu.Nha_may = '{current_user.macongty}'"""
            else:
                query = f"SELECT * FROM HR.dbo.Diem_danh_bu WHERE Nha_may = '{current_user.macongty}' "   
                if mst:
                    query += f"AND MST = '{mst}' "
                if hoten:
                    query += f"AND Ho_ten LIKE N'%{hoten}%' "
                if chucvu:
                    query += f"AND Chuc_vu LIKE N'%{chucvu}%' "
                if chuyen:
                    query += f"AND Line LIKE N'%{chuyen}%' "
                if bophan:
                    query += f"AND Bo_phan LIKE N'%{bophan}%' "
                if loaidiemdanh:
                    query += f"AND Loai_diem_danh LIKE N'%{loaidiemdanh}%' "
                if ngaydiemdanh:
                    query += f"AND Ngay_diem_danh = '{ngaydiemdanh}' "    
                if lido:
                    query += f"AND Ly_do LIKE N'%{lido}%' "   
                if trangthai:
                    query += f"AND Trang_thai LIKE N'%{trangthai}%' "    
                query += "ORDER BY Ngay_diem_danh DESC, Bo_phan ASC, Line ASC, MST ASC"
        print(query)
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(e)
        return []

def laydanhsachxinnghiphep(mst,hoten,chucvu,chuyen,bophan,ngaynghi,lydo,trangthai,mstquanly,mstthuky):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if mstthuky:
            query = f"""
            SELECT  *
            FROM 
                DS_Xin_nghi_phep 
            INNER JOIN 
                Phan_quyen_thu_ky
            ON
                DS_Xin_nghi_phep.Nha_may= Phan_quyen_thu_ky.Nha_may and DS_Xin_nghi_phep.Chuyen=Phan_quyen_thu_ky.Chuyen_to
            WHERE 
                DS_Xin_nghi_phep.Trang_thai=N'Chờ kiểm tra' and Phan_quyen_thu_ky.MST='{mstthuky}'"""
        else:
            if mstquanly:
                query = f"""
                SELECT 
                    *
                FROM 
                    DS_Xin_nghi_phep 
                INNER JOIN 
                    Phan_quyen_thu_ky
                ON
                    DS_Xin_nghi_phep.Nha_may= Phan_quyen_thu_ky.Nha_may and DS_Xin_nghi_phep.Chuyen=Phan_quyen_thu_ky.Chuyen_to
                WHERE 
                    DS_Xin_nghi_phep.Trang_thai=N'Đã kiểm tra' and MST_QL='{mstquanly}'"""
            else:            
                query = f"SELECT * FROM HR.dbo.DS_Xin_nghi_phep WHERE Nha_may = '{current_user.macongty}' "
                if mst:
                    query += f"AND MST = '{mst}'"
                if hoten:
                    query += f"AND Ho_ten LIKE N'%{hoten}%'"
                if chucvu:
                    query += f"AND Chuc_vu LIKE N'%{chucvu}%'"
                if chuyen:
                    query += f"AND Chuyen LIKE N'%{chuyen}%'"
                if bophan:
                    query += f"AND Bo_phan LIKE N'%{bophan}%'"
                if ngaynghi:
                    query += f"AND Ngay_nghi_phep = '{ngaynghi}'"    
                if trangthai:
                    query += f"AND Trang_thai LIKE N'%{trangthai}%'"
                query += " ORDER BY Ngay_nghi_phep DESC, MST ASC"
        ##
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(e)
        return []

def laydanhsachxinnghikhongluong(mst,hoten,chucvu,chuyen,bophan,ngay,lydo,trangthai,mstquanly,mstthuky):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if mstthuky:
            query = f"""
            SELECT  *
            FROM 
                Xin_nghi_khong_luong 
            INNER JOIN 
                Phan_quyen_thu_ky
            ON
                Xin_nghi_khong_luong.Nha_may= Phan_quyen_thu_ky.Nha_may and Xin_nghi_khong_luong.Chuyen=Phan_quyen_thu_ky.Chuyen_to
            WHERE 
                Xin_nghi_khong_luong.Trang_thai=N'Chờ kiểm tra' and Phan_quyen_thu_ky.MST='{mstthuky}'"""
        else:
            if mstquanly:
                query = f"""
                SELECT *
                FROM 
                    Xin_nghi_khong_luong 
                INNER JOIN 
                    Phan_quyen_thu_ky
                ON
                    Xin_nghi_khong_luong.Nha_may= Phan_quyen_thu_ky.Nha_may and Xin_nghi_khong_luong.Chuyen=Phan_quyen_thu_ky.Chuyen_to
                WHERE 
                    Xin_nghi_khong_luong.Trang_thai=N'Đã kiểm tra' and Phan_quyen_thu_ky.MST_QL='{mstquanly}'"""
            else:
                
                query = f"SELECT * FROM HR.dbo.Xin_nghi_khong_luong WHERE Nha_may = '{current_user.macongty}' "
                if mst:
                    query += f"AND MST LIKE '{mst}'"
                if hoten:
                    query += f"AND Ho_ten LIKE N'%{hoten}%'"
                if chucvu:
                    query += f"AND Chuc_vu LIKE N'%{chucvu}%'"
                if chuyen:
                    query += f"AND Chuyen LIKE N'%{chuyen}%'"
                if bophan:
                    query += f"AND Bo_phan LIKE N'%{bophan}%'"
                if ngay:
                    query += f"AND Ngay_xin_phep = '{ngay}'"    
                if lydo:
                    query += f"AND Ly_do LIKE N'%{lydo}%'"
                if trangthai:
                    query += f"AND Trang_thai LIKE N'%{trangthai}%'"
                query += " ORDER BY Ngay_xin_phep DESC, Bo_phan ASC, Chuyen ASC, MST ASC"
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(e)
        return []

def thuky_duoc_phanquyen(mst,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT COUNT(*) FROM HR.dbo.Phan_quyen_thu_ky WHERE Nha_may = '{current_user.macongty}' AND MST = '{mst}' AND Chuyen_to = '{chuyen}'"
        result = cursor.execute(query).fetchone()
        conn.close()
        if result[0] > 0:
            return True
        else:
            return False
    except Exception as e:
        print(e)
        return False

def quanly_duoc_phanquyen(mst,chuyen):
    try:
        if not chuyen:
            return False
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT COUNT(*) FROM HR.dbo.Phan_quyen_thu_ky WHERE Nha_may = '{current_user.macongty}' AND MST_QL = '{mst}' AND Chuyen_to = '{chuyen}'"
        result = cursor.execute(query).fetchone()
        conn.close()
        if result[0] > 0:
            return True
        else:
            return False
    except Exception as e:
        print(e)
        return False
    
def kiemtrathuki(mst,chuyen):
    try:
        if not chuyen:
            return False
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT COUNT(*) FROM HR.dbo.Phan_quyen_thu_ky WHERE Nha_may = '{current_user.macongty}' AND MST = '{mst}' AND Chuyen_to = '{chuyen}'"
        result = cursor.execute(query).fetchone()
        conn.close()

        if result[0] > 0:
            return True
        else:
            return False
    except Exception as e:
        print(e)
        return False

def capnhat_xinnghiphep(mst,ngay):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_phep SET Trang_thai = N'Đã phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND MST = '{mst}' AND Ngay_nghi_phep = '{ngay}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)

def insert_tangca(nhamay,mst,hoten,chucvu,chuyen,phongban,ngay,giotangca):
    try:
        if chucvu=='nan':
            chucvu = 'Không'
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"INSERT INTO HR.dbo.Dang_ky_tang_ca VALUES (N'{nhamay}','{mst}',N'{hoten}',N'{chucvu}',N'{chuyen}',N'{phongban}','{ngay}','{giotangca}',NULL, NULL, NULL, NULL, NULL, NULL)"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        conn.close()
        return False
        
def laydanhsachtangca(mst=None,phongban=None,chuyen=None,ngayxem=None,tungay=None,denngay=None):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        
        query = f"SELECT * FROM HR.dbo.Dang_ky_tang_ca WHERE Nha_may = '{current_user.macongty}'"
        if mst:
            query += f"AND MST = '{mst}' "
        if phongban:
            query += f"AND Bo_phan = '{phongban}' "
        if chuyen:
            query += f"AND Chuyen_to = '{chuyen}' "
        if ngayxem:
            query += f"AND Ngay_dang_ky = '{ngayxem}'"
        if tungay:
            query += f"AND Ngay_dang_ky >= '{tungay}'"
        if denngay:
            query += f"AND Ngay_dang_ky <= '{denngay}'"
        query += f" ORDER BY Ngay_dang_ky desc, CAST(MST as INT) asc"
        
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(e)  
        return [] 
    
def laydanhsachphepton(mst=None):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Phep_ton_chi_tiet WHERE Nha_may = '{current_user.macongty}'"
        if mst:
            query += f" AND MST = '{mst}'"

        query += " ORDER BY CAST(MST AS INT) asc"
        rows = cursor.execute(query).fetchall()
        conn.close()
        result = []
        return rows 
    except Exception as e:
        print(e)   
        return [] 
    
def laydanhsachkyluat():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Xu_ly_ky_luat WHERE Nha_may = '{current_user.macongty}'"
        
        rows = cursor.execute(query).fetchall()
        conn.close()
        result = []
        return rows 
    except Exception as e:
        print(e)  
        return [] 

def themdanhsachkyluat(mst,hoten,chucvu,bophan,chuyento,ngayvao,ngayvipham,diadiem,ngaylapbienban,noidung,bienphap):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"INSERT INTO HR.dbo.Xu_ly_ky_luat VALUES('{diadiem}','{mst}',N'{hoten}',N'{chucvu}','{chuyento}','{bophan}','{ngayvao}','{ngayvipham}','{ngaylapbienban}','{diadiem}',N'{noidung}',N'{bienphap}')"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)

def chuadangkycalamviec(mst):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select count(*) from Dang_ky_ca_lam_viec where Factory='{current_user.macongty}' and MST='{mst}'"
        row = cursor.execute(query).fetchone()
        conn.close()
        if row[0]==0:
            return True
    except Exception as e:
        print(f"Kiem tra da dang ky ca lam viec chua loi: {e} !!!")
        return False
        
def themdoicamoi(mst,cacu,camoi,ngaybatdau,ngayketthuc):
    try:
        if type(ngayketthuc) == str:
            ngayketthuc = datetime.strptime(ngayketthuc, '%Y-%m-%d')
        if type(ngaybatdau) == str:
            ngaybatdau = datetime.strptime(ngaybatdau, '%Y-%m-%d')
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if pd.notna(ngayketthuc):
            print("Co ngay ket thuc")
            ngayketthuccacu = ngaybatdau - timedelta(days=1)
            ngayvecamacdinh = ngayketthuc + timedelta(days=1)
            if ngaybatdau <= ngayketthuc:
                ngaymoc = datetime(2054,12,31)
                if chuadangkycalamviec(mst):
                    query = f"INSERT INTO HR.dbo.Dang_ky_ca_lam_viec VALUES ('{int(mst)}','{current_user.macongty}','{ngaybatdau}','{ngaymoc}','{camoi}')"
                else:
                    query = f"""
                            UPDATE HR.dbo.Dang_ky_ca_lam_viec
                            SET Den_ngay = '{ngayketthuccacu}'
                            WHERE MST='{int(mst)}' AND Den_ngay = '{ngaymoc}' AND Factory = '{current_user.macongty}'

                            INSERT INTO HR.dbo.Dang_ky_ca_lam_viec
                            VALUES ('{int(mst)}','{current_user.macongty}','{ngaybatdau}','{ngayketthuc}','{camoi}')

                            INSERT INTO HR.dbo.Dang_ky_ca_lam_viec
                            VALUES ('{int(mst)}','{current_user.macongty}','{ngayvecamacdinh}','{ngaymoc}','{cacu}')
                        """
                cursor.execute(query)
                conn.commit()
                conn.close()
                print("Đổi ca thành công", "success")
                return True
            else:
                print("Đổi ca thất bại, ngày bắt đầu lớn hơn ngày kết thúc")
        else:
            print("Khong co ngay ket thuc")
            ngayketthuccacu = ngaybatdau - timedelta(days=1)
            ngaymoc = datetime(2054,12,31)
            query = f"""
                        UPDATE HR.dbo.Dang_ky_ca_lam_viec
                        SET Den_ngay = '{ngayketthuccacu}'
                        WHERE MST='{int(mst)}' AND Den_ngay = '{ngaymoc}' AND Factory = '{current_user.macongty}'

                        INSERT INTO HR.dbo.Dang_ky_ca_lam_viec
                        VALUES ('{int(mst)}','{current_user.macongty}','{ngaybatdau}','{ngaymoc}','{camoi}')
                    """
            cursor.execute(query)
            conn.commit()
            conn.close()
            print("Đổi ca thành công", "success")
            return True
    except Exception as e:
        print(e)   
        return False 
        
def laycahientai(mst):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        ngayketthuc = datetime(2054,12,31)
        query = f"SELECT * FROM HR.dbo.Dang_ky_ca_lam_viec WHERE MST = '{mst}' AND Factory = '{current_user.macongty}' AND Den_ngay = '{ngayketthuc}'"
        
        row = cursor.execute(query).fetchone()
        if row:
            return row[-2]
        return None
    except Exception as e:
        print(e)
        return None

def laydanhsachyeucautuyendung(maso):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Yeu_cau_tuyen_dung WHERE Bo_phan LIKE '{maso}%'"
        
        rows = cursor.execute(query).fetchall()
        result =[]
        for row in rows:
            result.append(row)
        return result 
    except Exception as e:
        print(e)
        return []
    
def themyeucautuyendungmoi(bophan,vitri,soluong,mota,thoigiandukien,phanloai, khoangluong,capbac,bacluong):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"""INSERT INTO HR.dbo.Yeu_cau_tuyen_dung 
        (Bo_phan,Vi_tri,Grade_code,Bac_luong,Khoang_luong,So_luong,JD,Thoi_gian_du_kien,Phan_loai,Trang_thai_yeu_cau,Trang_thai_thuc_hien,Ghi_chu,MST,HO_TEN,NHA_MAY)
        VALUES
        ('{bophan}',N'{vitri}','{capbac}',N'{bacluong}',N'{khoangluong}','{soluong}',N'{mota}','{thoigiandukien}',N'{phanloai}',N'Chưa phê duyệt',N'Chưa tuyển',NULL,'{current_user.masothe}',N'{current_user.hoten}','{current_user.macongty}')"""
        cursor.execute(query)
        conn.commit()
        return True
    except Exception as e:
        print(e)
        return False
    
def laydanhsachxinnghikhac(mst,chuyen,bophan,ngaynghi,loainghi,trangthai,nhangiayto):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"""
        SELECT  Xin_nghi_khac.*,
                Danh_sach_CBCNV.Ho_ten,
                Danh_sach_CBCNV.Line,
                Danh_sach_CBCNV.Department
        FROM Xin_nghi_khac
        JOIN Danh_sach_CBCNV
        ON Xin_nghi_khac.MST = Danh_sach_CBCNV.The_cham_Cong and Xin_nghi_khac.Nha_may = Danh_sach_CBCNV.Factory
        where Xin_nghi_khac.Nha_may='{current_user.macongty}' """
        
        if mst:
            query += f" AND Xin_nghi_khac.MST='{mst}'" 
        if chuyen:
            query += f" AND Danh_sach_CBCNV.Line='{chuyen}'" 
        if bophan:
            query += f" AND Danh_sach_CBCNV.Department='{bophan}'" 
        if ngaynghi:
            query += f" AND Xin_nghi_khac.Ngay_nghi='{ngaynghi}'" 
        if loainghi:
            query += f" AND Xin_nghi_khac.Loai_nghi='{loainghi}'" 
        if trangthai:
            if trangthai=="Chưa kiểm tra":
                query += f" AND (Xin_nghi_khac.Trang_thai=N'{trangthai}' or Xin_nghi_khac.Trang_thai is NULL)"
            else:
                query += f" AND Xin_nghi_khac.Trang_thai=N'{trangthai}'"
        if nhangiayto:
            if nhangiayto=="Chưa nhận":
                query += f" AND (Xin_nghi_khac.Giay_to=N'{nhangiayto}' or Xin_nghi_khac.Giay_to is NULL)"
            else:
                query += f" AND Xin_nghi_khac.Giay_to=N'{nhangiayto}'" 
        query += " ORDER BY Xin_nghi_khac.Ngay_nghi DESC, Xin_nghi_khac.MST ASC"
        ##
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows 
    except Exception as e:
        print(e)
        return []

# def xoadulieuchamcong2ngay():
#     try:
#         conn = pyodbc.connect(used_db)
#         cursor = conn.cursor()
#         query = f"Delete from Check_in_out where CONVERT(varchar, NgayCham, 23) >= CONVERT(varchar, DATEADD(DAY, -2, GETDATE()), 23) and Nha_may = 'NT1';"
#         cursor.execute(query)
#         conn.commit()
#         conn.close()
#         return True
#     except Exception as e:
#         print(e)
#         return False

# def themdulieuchamcong2ngay():
#     try:
#         xoadulieuchamcong2ngay()
#         conn = pyodbc.connect(mccdb)
#         cursor = conn.cursor()
#         query = f"""
# 	    select 'NT1',MaChamCong,NgayCham,GioCham,TenMay from checkinout 
# 	    where CONVERT(varchar, NgayCham, 23) >= CONVERT(varchar, DATEADD(DAY, -2, GETDATE()), 23)"""
#         rows = cursor.execute(query).fetchall()
#         conn.close()
#         conn1 = pyodbc.connect(used_db)
#         cursor1 = conn1.cursor()
#         query1 = "insert into Check_in_out(Nha_may,Machamcong,NgayCham,GioCham,TenMay) values(?,?,?,?,?)"
#         for row in rows:
#             cursor1.execute(query1, row)
#         conn1.commit()
#         conn1.close()
#         return True
#     except Exception as e:
#         print(e)
#         return False
    
def thuky_dakiemtra_diemdanhbu(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Diem_danh_bu SET Trang_thai = N'Đã kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
    
def thuky_tuchoi_diemdanhbu(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Diem_danh_bu SET Trang_thai = N'Bị từ chối bởi người kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)

def quanly_pheduyet_diemdanhbu(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Diem_danh_bu SET Trang_thai = N'Đã phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
    
def quanly_tuchoi_diemdanhbu(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Diem_danh_bu SET Trang_thai = N'Bị từ chối bởi người phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)

def thuky_dakiemtra_xinnghiphep(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_phep SET Trang_thai = N'Đã kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
    
def thuky_tuchoi_xinnghiphep(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_phep SET Trang_thai = N'Bị từ chối bởi người kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)

def quanly_pheduyet_xinnghiphep(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_phep SET Trang_thai = N'Đã phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
    
def quanly_tuchoi_xinnghiphep(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_phep SET Trang_thai = N'Bị từ chối bởi người phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
        
def thuky_dakiemtra_xinnghikhongluong(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khong_luong SET Trang_thai = N'Đã kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
    
def thuky_tuchoi_xinnghikhongluong(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khong_luong SET Trang_thai = N'Bị từ chối bởi người kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)

def quanly_pheduyet_xinnghikhongluong(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khong_luong SET Trang_thai = N'Đã phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
    
def quanly_tuchoi_xinnghikhongluong(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khong_luong SET Trang_thai = N'Bị từ chối bởi người phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        # 
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)

def thuky_dakiemtra_xinnghikhac(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khac SET Trang_thai = N'Đã kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
    
def thuky_tuchoi_xinnghikhac(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khac SET Trang_thai = N'Bị từ chối bởi người kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)

def quanly_pheduyet_xinnghikhac(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khac SET Trang_thai = N'Đã phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
    
def quanly_tuchoi_xinnghikhac(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khac SET Trang_thai = N'Bị từ chối bởi người phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)

def nhansu_nhangiayto_xinnghikhac(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khac SET Giay_to = N'Đã nhận' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
        
def nhansu_khongnhangiayto_xinnghikhac(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khac SET Giay_to = N'Không có' WHERE Nha_may = '{current_user.macongty}' AND ID = '{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
        
def laydanhsachcahientai(mst,chuyen, phongban):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"""
        SELECT 
            Dang_ky_ca_lam_viec.Factory,
            Dang_ky_ca_lam_viec.MST, 
            Danh_sach_CBCNV.Ho_ten,
            Danh_sach_CBCNV.Line,
            Danh_sach_CBCNV.Department,
            Dang_ky_ca_lam_viec.Ca,
            Dang_ky_ca_lam_viec.Tu_ngay,
            Dang_ky_ca_lam_viec.Den_ngay,
            Dang_ky_ca_lam_viec.ID
        FROM 
            Dang_ky_ca_lam_viec
        INNER JOIN 
            Danh_sach_CBCNV 
        ON 
            Dang_ky_ca_lam_viec.MST = Danh_sach_CBCNV.The_cham_cong AND Dang_ky_ca_lam_viec.Factory = Danh_sach_CBCNV.Factory
        WHERE 
            Dang_ky_ca_lam_viec.Factory = '{current_user.macongty}'
        """
        if mst:
            query += f" AND Dang_ky_ca_lam_viec.MST = '{mst}'"
        if chuyen:
            query += f" AND Danh_sach_CBCNV.Line LIKE '%{chuyen}%'"
        if phongban:
            query += f" AND Danh_sach_CBCNV.Department LIKE '%{phongban}%'"
        query += "ORDER BY Dang_ky_ca_lam_viec.Tu_ngay desc, Dang_ky_ca_lam_viec.Den_ngay desc, MST asc"
        # print(query)
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(e)
        return []

def laydanhsachkpichuaduyet(mst,macongty):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from KPI_Data where Status = 'Waiting for approval' "
        if mst:
            query += f"and MST='{mst}'"
        if macongty:
            query += f"and Nha_may='{macongty}' " 
        # 
        rows = cursor.execute(query).fetchall()
        conn.close()
        if not mst:
            return []
        return rows
    except Exception as e:
        print(e)
        return []  
   
def laydanhsachkpidaduyet(mst,macongty):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from KPI_Data where Status = 'Approved' "
        if mst:
            query += f"and MST='{mst}'"
        if macongty:
            query += f"and Nha_may='{macongty}' " 
        query += f"order by Nha_may asc, MST asc"
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        print
        return []
      
def roles_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if not current_user.is_authenticated or current_user.phanquyen not in roles:
                return redirect(url_for('unauthorized'))
            return f(*args, **kwargs)
        return decorated_function
    return decorator

def delete_kpidata(masothe,macongty):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"delete KPI_Data where MST='{masothe}' and Nha_may='{macongty}'"

        cursor.execute(query)
        conn.commit()
        conn.close()
        return True    
    except Exception as e:
        print(e)
        return False
    
def insert_kpidata(masothe:str,macongty:str,values:list):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if values[0]!=macongty or values[1]!=str(masothe):
            return False
        query = f"insert into KPI_Data values ("
        for value in values:
            if values.index(value)==2: 
                value = f"N'{value}',"
            else:
                if value != np.nan:
                    value = f"'{value}',"
                else:
                    value = 'NULL,'
            query += value
        query += "'Waiting for approval',GETDATE())" 
        query = query.replace("'nan'","NULL") 
        # 
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        return False
    
def guimailthongbaodaguikpi(nhamay,mst,hoten):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into KPI_Cho_phe_duyet values ('{nhamay}','{mst}',N'{hoten}',GETDATE())"
        # #   
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
        return False
    
def guimailthongbaodapheduyetkpi(nhamay,mst,hoten,email):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into KPI_Da_phe_duyet values ('{nhamay}','{mst}',N'{hoten}','{email}',GETDATE())"  
        # # 
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
        return False
    
def guimailthongbaodatuchoikpi(nhamay,mst,hoten,email):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into KPI_Bi_tu_choi values ('{nhamay}','{mst}',N'{hoten}','{email}',GETDATE())"  
        # # 
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
        return False
    
def laydanhsachquanly(macongty):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if macongty:
            query = f"select distinct MST,Ho_ten from KPI_Data where Nha_may='{macongty}'"  
        else:
            query = "select distinct MST,Ho_ten from KPI_Data"
        rows = cursor.execute(query).fetchall()
        conn.close()
        result=[]
        for row  in rows:
            result.append({"mst":row[0],"hoten":row[1]})
        return result        
    except Exception as e:
        print(e)
        return []
    
def pheduyetkpi(masothe,macongty):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"update KPI_Data set Status='Approved',Time_stamp=GETDATE() where Nha_may='{macongty}' and MST='{masothe}'"  
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        return False
    
def tuchoikpi(masothe,macongty):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"update KPI_Data set Status='Denied',Time_stamp=GETDATE() where Nha_may='{macongty}' and MST='{masothe}'"  
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        return False
    
def layemailquanly(macongty,masothe):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select Email from KPI_DS_Email where Nha_may='{macongty}' and MST='{masothe}'"  
        row = cursor.execute(query).fetchone()
        conn.close()
        return row[0]
    except Exception as e:
        print(e)
        return False
    
def laydanhsach_chonghiviec(mst,hoten,chuyen,phongban,ngaynopdon,ngaynghi,saphethan):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"""
        SELECT Cho_nghi_viec.*, Danh_sach_CBCNV.Trang_thai_lam_viec
        FROM Cho_nghi_viec
        LEFT JOIN Danh_sach_CBCNV
        ON Cho_nghi_viec.MST = Danh_sach_CBCNV.MST 
        AND Cho_nghi_viec.Nha_may = Danh_sach_CBCNV.Factory
        WHERE Nha_may = '{current_user.macongty}'
        """
        if mst:
            query += f" AND Cho_nghi_viec.MST = '{mst}'"
        if hoten:
            query += f" AND Cho_nghi_viec.Ho_ten LIKE N'%{hoten}%'"
        if chuyen:
            query += f" AND Cho_nghi_viec.Chuyen_to LIKE N'%{chuyen}%'"
        if phongban:
            query += f" AND Cho_nghi_viec.Bo_phan LIKE '%{phongban}%'"
        if ngaynopdon:
            query += f" AND Cho_nghi_viec.Ngay_nop_don = '{ngaynopdon}'"
        if ngaynghi:
            query += f" AND Cho_nghi_viec.Ngay_nghi_du_kien = '{ngaynghi}'" 
        if saphethan:
            ngayhientai = datetime.now().date()
            ngaynghisapden = datetime.now().date() + timedelta(days=7)
            query += f" AND Cho_nghi_viec.Ngay_nghi_du_kien >= '{ngayhientai}' AND Cho_nghi_viec.Ngay_nghi_du_kien <= '{ngaynghisapden}'"   
        query += "ORDER BY Cho_nghi_viec.Ngay_nghi_du_kien DESC,Cho_nghi_viec.Ngay_nop_don DESC"  
        # # 
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(e)
        return []

def themdonxinnghi(mst,hoten,chucdanh,chuyen,phongban,ngaynopdon,ngaynghi,ghichu):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if ghichu:
            query = f"INSERT INTO Cho_nghi_viec VALUES ('{current_user.macongty}','{mst}',N'{hoten}',N'{chucdanh}','{chuyen}','{phongban}','{ngaynopdon}','{ngaynghi}',N'{ghichu}')"
        else:
            query = f"INSERT INTO Cho_nghi_viec VALUES ('{current_user.macongty}','{mst}',N'{hoten}',N'{chucdanh}','{chuyen}','{phongban}','{ngaynopdon}','{ngaynghi}',NULL)"
        # # 
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        return False
    
def rutdonnghiviec(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"DELETE Cho_nghi_viec WHERE ID='{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        return False

def laydanhsach_hopdong_theomst(mst):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM [HR].[dbo].[QUAN_LY_HD] WHERE MST='{mst}' AND NHA_MAY='{current_user.macongty}'"
        rows = cursor.execute(query).fetchall()
        conn.close()
        return [{
            "Số thứ tự": row[0],
            "Nhà máy": row[1],
            "Mã số thẻ": row[2],
            "Họ tên": row[3],
            "Giới tính": row[4],
            "Ngày sinh": row[5],
            "Địa chỉ": row[6],
            "Tạm trú": row[7],
            "CCCD": row[8],
            "Ngày cấp CCCD": row[9],
            "Cấp bậc": row[10],
            "Loại hợp đồng": row[11],
            "Chức danh": row[12],
            "Phòng ban": row[13],
            "Chuyền": row[14],
            "Lương cơ bản": row[15],
            "Phụ cấp": row[16],
            "Ngày ký hợp đồng": row[17],
            "Ngày hết hạn hợp đồng": row[18]
        } for row in rows]
    except Exception as e:
        print(e)
        return []  

def capnhat_stk(mst, stk, macongty):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE Danh_sach_CBCNV SET So_tai_khoan=N'{stk}' WHERE MST='{mst}' AND Factory='{macongty}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Cap nhat so tai khoan loi {e}")
        return False 
    
def lay_thongtin_hopdong_theo_id(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from QUAN_LY_HD WHERE ID='{id}'"
        result = cursor.execute(query).fetchone()
        conn.close()
        return result
    except Exception as e:
        flash(f"Lay thong tin hop dong loi {e} !!!")
        return ()
    
def timkiemchucdanh(tutimkiem):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select DISTINCT Detail_job_title_VN from HC_Name WHERE Detail_job_title_VN COLLATE Latin1_General_CI_AI LIKE N'%{tutimkiem}%'"
        result = cursor.execute(query).fetchall()
        conn.close()
        return list(x[0] for x in result)
    except Exception as e:
        print(f"Loi khi tim kiem cac chuc danh: {e} !!!")
        return []
    
def themhopdongmoi(nhamay,mst,hoten,gioitinh,ngaysinh,thuongtru,tamtru,cccd,ngaycapcccd,capbac,loaihopdong,chucdanh,phongban,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc):
    
    try:
        try:
            ngaybatdau=datetime.strptime(ngaybatdau,"%d/%m/%Y").strftime("%Y-%m-%d")
        except:
            ngaybatdau=ngaybatdau
        try:
            ngayketthuc=datetime.strptime(ngayketthuc,"%d/%m/%Y").strftime("%Y-%m-%d")
        except:
            ngayketthuc=ngayketthuc
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"""
        INSERT INTO QUAN_LY_HD VALUES (
            '{nhamay}', '{int(mst)}', N'{hoten}', N'{gioitinh}', '{ngaysinh}', N'{thuongtru}', N'{tamtru}', '{cccd}', '{ngaycapcccd}', '{capbac}',
            N'{loaihopdong}', N'{chucdanh}', '{phongban}', '{chuyen}', '{int(luongcoban)}', '0', '{ngaybatdau}', '{ngayketthuc}')
        """
        
        cursor.execute(query)
        conn.commit()
        conn.close()
        return {"ketqua":True}
    except Exception as e:
        print(f"Loi khi them hop dong co ngay ket thuc: {e} !!!\nQuery:{query}")
        query = f"""
        INSERT INTO QUAN_LY_HD VALUES (
            '{nhamay}', '{int(mst)}', N'{hoten}', N'{gioitinh}', '{ngaysinh}', N'{thuongtru}', N'{tamtru}', '{cccd}', '{ngaycapcccd}', '{capbac}',
            N'{loaihopdong}', N'{chucdanh}', '{phongban}', '{chuyen}', '{int(luongcoban)}', '0', '{ngaybatdau}', NULL )
        """
        cursor.execute(query)
        try:
            conn.commit()
            conn.close()
            return {"ketqua":True}
        except Exception as e:  
            return {"ketqua":False,"lido":e,"query":query}
    
def capnhatthongtinhopdong(nhamay,mst,loaihopdong,chucdanh,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc,vitrien,employeetype,posotioncode,postitioncodedescription,hccategory,sectioncode,sectiondescription):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if loaihopdong == "Phụ lục hợp đồng":
            query = f"""
                UPDATE Danh_sach_CBCNV SET Luong_co_ban='{luongcoban}', Phu_cap='{phucap}',
                Job_title_VN=N'{chucdanh}', Job_title_EN='{vitrien}', Emp_type='{employeetype}', Position_code='{posotioncode}', Position_code_description='{postitioncodedescription}',
                Headcount_category='{hccategory}', Section_code='{sectioncode}', Section_description='{sectiondescription}', Line=N'{chuyen}'
                WHERE Factory='{nhamay}' AND The_cham_cong='{mst}'
                """
        elif loaihopdong == "Hợp đồng có thời hạn 28 ngày" or loaihopdong == "Hợp đồng có thời hạn 1 năm":
            query = f"""
                UPDATE Danh_sach_CBCNV SET Luong_co_ban='{luongcoban}', Phu_cap='{phucap}', Ngay_ky_HDXDTH_Lan1='{ngaybatdau}', Ngay_het_han_HDXDTH_Lan1='{ngayketthuc}',
                Job_title_VN=N'{chucdanh}', Job_title_EN='{vitrien}', Emp_type='{employeetype}', Position_code='{posotioncode}', Position_code_description='{postitioncodedescription}',
                Headcount_category='{hccategory}', Section_code='{sectioncode}', Section_description='{sectiondescription}', Loai_hop_dong=N'{loaihopdong}', Line=N'{chuyen}',
                Ngay_ky_HD='{ngaybatdau}', Ngay_het_han_HD ='{ngayketthuc}'
                WHERE Factory='{nhamay}' AND The_cham_cong='{mst}'
                """
        elif loaihopdong == "Hợp đồng vô thời hạn":
            query = f"""
                UPDATE Danh_sach_CBCNV SET Luong_co_ban='{luongcoban}', Phu_cap='{phucap}', Ngay_ky_HDKXDTH='{ngaybatdau}',
                Job_title_VN=N'{chucdanh}', Job_title_EN='{vitrien}', Emp_type='{employeetype}', Position_code='{posotioncode}', Position_code_description='{postitioncodedescription}',
                Headcount_category='{hccategory}', Section_code='{sectioncode}', Section_description='{sectiondescription}', Loai_hop_dong=N'{loaihopdong}', Line=N'{chuyen}',
                Ngay_ky_HD='{ngaybatdau}', Ngay_het_han_HD =NULL
                WHERE Factory='{nhamay}' AND The_cham_cong='{mst}'"""
        elif loaihopdong == "Hợp đồng thử việc":
            query = f"""UPDATE Danh_sach_CBCNV SET Ngay_ky_HDTV='{ngaybatdau}', Ngay_het_han_HDTV='{ngayketthuc}', Loai_hop_dong=N'{loaihopdong}',
                Ngay_ky_HD='{ngaybatdau}', Ngay_het_han_HD ='{ngayketthuc}'
                WHERE Factory='{nhamay}' AND The_cham_cong='{mst}'
                """
        else:
            query = ""
        if query:
            try:
                cursor.execute(query)
                conn.commit()
                conn.close()  
                return {"ketqua":True}
            except Exception as e:
                return {"ketqua":False,"lido":e, "query":query}
        else:   
            return {"ketqua":False,"lido":"Khong hieu loai hop dong", "query":query}
    except Exception as e:
        return {"ketqua":False,"lido":e, "query":query}
    
def thaydoithongtinhopdong(id,masothe,hoten,gioitinh,ngaysinh,thuongtru,tamtru,cccd,ngaycapcccd,
                           loaihopdong,ngaybatdau,ngayketthuc,chuyen,capbac,chucdanh,phongban,luongcoban,phucap):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        ngayketthuc = f"'{ngayketthuc}'" if ngayketthuc else 'NULL'
        query = f"""
        update QUAN_LY_HD set MST='{masothe}',HO_TEN=N'{hoten}',GIOI_TINH=N'{gioitinh}',NGAY_SINH='{ngaysinh}',DIA_CHI=N'{thuongtru}',TAM_TRU=N'{tamtru}',
        CCCD='{cccd}',NGAY_CAP='{ngaycapcccd}',LOAI_HD=N'{loaihopdong}',CHUC_DANH=N'{chucdanh}',PHONG_BAN='{phongban}',CHUYEN='{chuyen}',CAP_BAC='{capbac}',
        LCB='{luongcoban}',PHU_CAP='{phucap}',NGAY_KY='{ngaybatdau}',NGAY_HET_HAN={ngayketthuc} where ID='{id}'
        """
        # 
        cursor.execute(query)
        conn.commit()
        conn.close()  
        return True
    except Exception as e:
        flash(f"Loi khi cap nhat thong tin hop dong: {e} !!!")
        return False
    
def xoa_hopdong(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"delete QUAN_LY_HD where ID='{id}'"
        ##
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        flash(f"Loi xoa hop dong: {e} !!!")
        return False
    
def them_diemdanhbu(masothe,hoten,chucdanh,chuyen,phongban,loaidiemdanh,ngay,giovao,lydo,trangthai):
    try:
        ngay = ngay.split("/")[2] + "-" + ngay.split("/")[1] + "-" + ngay.split("/")[0]
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into Diem_danh_bu values ('{current_user.macongty}','{masothe}',N'{hoten}',N'{chucdanh}','{chuyen}','{phongban}',N'{loaidiemdanh}','{ngay}','{giovao}',N'{lydo}',N'{trangthai}')"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        flash(f"Loi khi them diem danh bu: {e} !!!")
        return False
    
def them_xinnghiphep(masothe,hoten,chucdanh,chuyen,phongban,ngay,sophut,trangthai):
    try:
        ngay = ngay.split("/")[2] + "-" + ngay.split("/")[1] + "-" + ngay.split("/")[0]
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into Xin_nghi_phep values ('{current_user.macongty}','{masothe}',N'{hoten}',N'{chucdanh}','{chuyen}','{phongban}','{ngay}','{sophut}',NULL,N'{trangthai}')"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        flash(f"Loi khi them xin nghi phep: {e} !!!")
        return False
    
def them_xinnghikhongluong(masothe,hoten,chucdanh,chuyen,phongban,ngay,sophut,lydo,trangthai):
    try:
        ngay = ngay.split("/")[2] + "-" + ngay.split("/")[1] + "-" + ngay.split("/")[0]
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into Xin_nghi_khong_luong values ('{current_user.macongty}','{masothe}',N'{hoten}',N'{chucdanh}','{chuyen}','{phongban}','{ngay}','{sophut}',N'{lydo}',N'{trangthai}')"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        flash(f"Loi khi them xin nghi khong luong: {e} !!!")
        return False

def them_xinnghikhac(masothe,ngay,sophut,lydo,trangthai,nhangiayto):
    try:
        ngay = ngay.split("/")[2] + "-" + ngay.split("/")[1] + "-" + ngay.split("/")[0] if "/" in ngay else ngay 
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if nhangiayto == "Chưa":
            query = f"INSERT INTO Xin_nghi_khac VALUES ('{current_user.macongty}','{masothe}','{ngay}','{sophut}',N'{lydo}',N'{trangthai}',N'Chưa')"
        else:
            query = f"INSERT INTO Xin_nghi_khac VALUES ('{current_user.macongty}','{masothe}','{ngay}','{sophut}',N'{lydo}',N'{trangthai}',N'Đã nhận')"
        ##
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        flash(f"Loi khi them xin nghi khac: {e} !!!")
        return False

def lay_chuyen_va_capbac(macongty, mst):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select Line ,Grade_code from Danh_sach_CBCNV where The_cham_cong='{mst} 'and Factory='{macongty}'"
        row = cursor.execute(query).fetchone()
        conn.close()
        if row:
            return row[0],row[1]
        else:
            return None
    except Exception as e:
        print(f"Loi khi kiem tra co phai to truong khong: {e} !!!")
        return None
    
def capnhat_ghichu_lsct(mst,ngaythuchien,phanloai,ghichumoi):
    try:
        
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE Lich_su_Cong_tac set Ghi_chu = N'{ghichumoi}' where MST='{int(mst)}' and Phan_loai=N'{phanloai}' and Nha_may='{current_user.macongty}' "
        if ngaythuchien == 'None':
            query+= "and Ngay_thuc_hien is null "
        else:
            query+= f"and Ngay_thuc_hien = '{ngaythuchien}'"
        ##
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        return False

def capnhat_ngaythuchien_lsct(mst,ngaythuchienmoi,phanloai,chuyencu,chuyenmoi):
    try:
        
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"""UPDATE Lich_su_Cong_tac 
        set Ngay_thuc_hien = N'{ngaythuchienmoi}' 
        where MST='{int(mst)}' and Phan_loai=N'{phanloai}' 
        and Nha_may='{current_user.macongty}' 
        and Line_cu='{chuyencu}'
        and Line_moi='{chuyenmoi}'"""
        ##
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        return False

def sua_dangky_ca(id,camoi):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE Dang_ky_ca_lam_viec set Ca = N'{camoi}' where ID ='{id}' "
        ##
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        return False   


def suadoi_ngaybatdau_ca_dangky_ca(id,ngaybatdau_camoi):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE Dang_ky_ca_lam_viec set Tu_ngay = '{ngaybatdau_camoi}' where ID ='{id}' "
        ##
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        return False   
    
def suadoi_ngayketthuc_ca_dangky_ca(id,ngayketthuc_camoi):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE Dang_ky_ca_lam_viec set Den_ngay = '{ngayketthuc_camoi}' where ID ='{id}' "
        ##
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(e)
        return False  
    
def lay_chuyen_theo_mst(mst):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select Line from Danh_sach_CBCNV where The_cham_cong='{mst}' and Factory='{current_user.macongty}'"
        ##
        row = cursor.execute(query).fetchone()
        conn.close()
        if row:
            return row[0]
        else:
            return None
    except Exception as e:
        print(e)
        return None 
    
def danhsach_tangca(chuyen:list,ngay,pheduyet):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[Dang_ky_tang_ca] where ("
        for ch in chuyen:
            query += f" Chuyen_to='{ch}' or"
        query = query[:-2] + " ) "
        if pheduyet == "ok" or pheduyet== "notok":
            query += f" and HR IS NOT NULL " if pheduyet == "ok" else f" and HR IS NULL "
        query += f" and Ngay_dang_ky = '{ngay}' and Nha_may='{current_user.macongty}' ORDER BY CAST(MST AS INT) ASC, GIO_VAO ASC"
        # print(query)
        cursor = cursor.execute(query)
        rows = cursor.fetchall()
        result = [{
            "ID": row[17],
            "Nhà máy": row[0],
            "Mã số thẻ": row[1],
            "Họ tên": row[2],
            "Chức danh": row[3],
            "Chuyền": row[4],
            "Phòng ban": row[5],
            "Ngày": row[6],
            "Tăng ca sáng": row[7][:5] if row[7] else "",
            "Tăng ca sáng thực tế": row[8][:5] if row[8] else "",
            "Giờ tăng ca": row[9][:5] if row[9] else "",
            "Giờ tăng ca thực tế": row[10][:5] if row[10] else "",
            "Tăng ca đêm": row[11][:5] if row[11] else "",
            "Tăng ca đêm thực tế": row[12][:5] if row[12] else "",
            "Ca": row[13],
            "Giờ vào": row[14][:5] if row[14] else "",
            "Giờ ra": row[15][:5] if row[15] else "",
            "HR phê duyệt": row[16] if row[16] else ""     
            } for row in rows]
        # print(result)
        conn.close()
        return result
    except Exception as e:
        flash(f"Lỗi lấy bảng đăng ký tang ca: ({e})")
        return []

def laychuyen_quanly(masothe,macongty):
    try:
        if "HRD" in current_user.phongban:
            conn = pyodbc.connect(used_db)
            cursor = conn.cursor()
            query = f"select distinct Chuyen_to from [HR].[dbo].[Phan_quyen_thu_ky] where Chuyen_to LIKE '{macongty[2]}%'"
            ##
            cursor = cursor.execute(query)
            rows = cursor.fetchall()
            result = [row[0] for row in rows]
            conn.close()
        else:
            conn = pyodbc.connect(used_db)
            cursor = conn.cursor()
            query = f"select Chuyen_to from [HR].[dbo].[Phan_quyen_thu_ky] where MST='{masothe}' and NHA_MAY='{macongty}'"
            
            cursor = cursor.execute(query)
            rows = cursor.fetchall()
            result = [row[0] for row in rows]
            # print(result)
            conn.close()
        return result
    except:
        return []

def capnhat_tangca_thanhcong(id,tangcasang,tangcasangthucte,tangca,tangcathucte,tangcadem,tangcademthucte):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE Dang_ky_tang_ca SET "
        if tangcasang:
            query += f"Tang_ca_sang = '{tangcasang}',"
        else:
            query += "Tang_ca_sang = NULL,"
        if tangcasangthucte:
            query += f"Tang_ca_sang_thuc_te = '{tangcasangthucte}',"
        else:
            query += "Tang_ca_sang_thuc_te = NULL,"
        if tangca:
            query += f"Gio_tang_ca = '{tangca}',"
        else:
            query += "Gio_tang_ca = NULL,"
        if tangcathucte:
            query += f"Gio_tang_ca_thuc_te = '{tangcathucte}',"
        else:
            query += "Gio_tang_ca_thuc_te = NULL,"
        if tangcadem:
            query += f"Tang_ca_dem = '{tangcadem}',"
        else:
            query += "Tang_ca_dem = NULL,"
        if tangcademthucte:
            query += f"Tang_ca_dem_thuc_te = '{tangcadem}'"
        else:
            query += "Tang_ca_dem_thuc_te = NULL"
        query += f" WHERE ID='{id}'"
        
        cursor = cursor.execute(query)
        conn.commit()
        conn.close()
        return cursor
    except:
        return False
    
def nhansu_bopheduyet_tangca(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE Dang_ky_tang_ca SET HR=NULL WHERE ID='{id}'"
        cursor = cursor.execute(query)
        conn.commit()
        conn.close()
        return cursor
    except:
        return False 
    
def nhansu_pheduyet_tangca(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE Dang_ky_tang_ca SET HR='OK' WHERE ID='{id}'"
        cursor = cursor.execute(query)
        conn.commit()
        conn.close()
        return cursor
    except:
        return False 
        
def tat_function_12():
    try:
        config = ConfigParser()
        config.read("f12.ini")
        config["F12"]={"ON":0}
        with open("f12.ini", "w") as configfile:
            config.write(configfile)
    except Exception as e:
        return False
    
def bat_function_12():
    try:
        config = ConfigParser()
        config.read("f12.ini")
        config["F12"]={"ON":1}
        with open("f12.ini", "w") as configfile:
            config.write(configfile)
        return config.get("F12","ON")
    except Exception as e:
        return False

def trang_thai_function_12():
    try:
        config = ConfigParser()
        config.read("f12.ini")
        return config.get("F12","ON")
    except Exception as e:
        return None
    
def danhsach_chamcong_sang(chuyen,bophan,cochamcong):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"""Select * from Cham_cong_sang where Factory='{current_user.macongty}'"""
        if chuyen:
            query += f" and Chuyen_to = '{chuyen}'"
        if bophan:
            query += f" and Bo_phan = '{bophan}'"
        if cochamcong:
            if cochamcong=="co":
                query += f" and Gio_vao is not null"
            if cochamcong=="khong":
                query += f" and Gio_vao is null"
        query += " and Trang_thai_lam_viec=N'Đang làm việc' order by The_cham_cong asc"
        
        cursor = cursor.execute(query)
        rows = cursor.fetchall()
        result = [row for row in rows]
        return result
    except Exception as e:
        return []
    
def them_dangky_tangca(nhamay, mst, hoten, chucdanh, chuyen, 
                       phongban, ngay, giotangcasang, giotangcasangthucte, 
                       giotangca, giotangcathucte, giotangcadem, giotangcademthucte, 
                       ca, giovao, giora, hrpheduyet):
    query = f"""INSERT INTO Dang_ky_tang_ca VALUES ('{nhamay}','{mst}',N'{hoten}',N'{chucdanh}','{chuyen}','{phongban}','{ngay}',"""
    if giotangcasang:
        query += f"'{giotangcasang}',"
    else:
        query += 'NULL,'
    if giotangcasangthucte:
        query += f"'{giotangcasangthucte}',"
    else:
        query += 'NULL,'
    if giotangca:
        query += f"'{giotangca}',"
    else:
        query += 'NULL,'
    if giotangcathucte:
        query += f"'{giotangcathucte}',"
    else:
        query += 'NULL,'
    if giotangcadem:
        query += f"'{giotangcadem}',"
    else:
        query += 'NULL,'
    if giotangcademthucte:
        query += f"'{giotangcademthucte}',"
    else:
        query += 'NULL,'
    if ca:
        query += f"'{ca}',"
    else:
        query += 'NULL,'
    if giovao:
        query += f"'{giovao}',"
    else:
        query += 'NULL,'
    if giora:
        query += f"'{giora}',"
    else:
        query += 'NULL,'
    if hrpheduyet:
        query += f"'{hrpheduyet}')"
    else:
        query += 'NULL)'
    ##
    conn = pyodbc.connect(used_db)
    cursor = conn.cursor()
    try:
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi them dang ky tang ca: {e}")
        return False
    
def lay_bangcong_thucte(thang,nam,mst,bophan,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[BANG_CONG_TONG_THUC_TE] where Nha_may='{current_user.macongty}'"
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay bang cong tong thuc te: {e}")
        return []

def lay_bangcong_kx(thang,nam,mst,bophan,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[BANG_CONG_TONG_KIEM_XUONG] where Nha_may='{current_user.macongty}'"
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay bang cong tong thuc te: {e}")
        return []
 
def lay_tangcachedo_web(thang,nam,mst,bophan,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[TANG_CA_CHE_DO_THUC_TE] where Nha_may='{current_user.macongty}'"
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay bang tang ca che do: {e}")
        return []
    
def lay_tangcachedo(thang,nam,mst,bophan,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[TANG_CA_CHE_DO] where Nha_may='{current_user.macongty}'"
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay bang tang ca che do: {e}")
        return []

def lay_tangcangay(thang,nam,mst,bophan,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[TANG_CA_NGAY] where Nha_may='{current_user.macongty}'"
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay bang tang ca ngay: {e}")
        return []
       
def lay_tangcangay_web(thang,nam,mst,bophan,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[TANG_CA_NGAY_THUC_TE] where Nha_may='{current_user.macongty}'"
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay bang tang ca ngay: {e}")
        return []
   
def lay_tangcadem(thang,nam,mst,bophan,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[TANG_CA_DEM] where Nha_may='{current_user.macongty}'"
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay bang tang ca chu nhat: {e}")
        return []
    
def lay_tangcadem_web(thang,nam,mst,bophan,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[TANG_CA_DEM_THUC_TE] where Nha_may='{current_user.macongty}'"
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay bang tang ca chu nhat: {e}")
        return []

def lay_tangcangayle(thang,nam,mst,bophan,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[TANG_CA_NGAY_LE] where Nha_may='{current_user.macongty}'"
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay bang tang ca ngay le: {e}")
        return []

def lay_tangcangayle_web(thang,nam,mst,bophan,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[TANG_CA_NGAY_LE_THUC_TE] where Nha_may='{current_user.macongty}'"
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay bang tang ca ngay le: {e}")
        return []

def lay_tangcachunhat(thang,nam,mst,bophan,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[TANG_CA_CHU_NHAT] where Nha_may='{current_user.macongty}'"
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay bang tang ca dem: {e}")
        return []
       
def lay_tangcachunhat_web(thang,nam,mst,bophan,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[TANG_CA_CHU_NHAT_THUC_TE] where Nha_may='{current_user.macongty}'"
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay bang tang ca dem: {e}")
        return []
    
def lay_dulieu_chamcong_web(mst,chuyen, bophan,ngay):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[NOT_DAP_THE_GOC] where Nha_may='{current_user.macongty}' "
        if mst:
            query += f" and MST = '{mst}'"
        if chuyen:
            query += f" and Chuyen = '{chuyen}'"
        if bophan:
            query += f" and Bo_phan = '{bophan}'"
        if ngay:
            query += f" and Ngay = '{ngay}'"
        query += " order by MST asc,Ngay desc, Bo_phan asc, Chuyen asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay cham cong goc: {e}")
        return []
    
def lay_bangcong5ngay_web(masothe,chuyen,bophan,phanloai,ngay,tungay,denngay):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[BANG_CHAM_CONG_TU_DONG_THUC_TE] where Nha_may='{current_user.macongty}' "
        if masothe:
            query += f" and MST = '{masothe}'"
        if chuyen:
            query += f" and Chuyen_to = '{chuyen}'"   
        if bophan:
            query += f" and Bo_phan = '{bophan}'"
        if phanloai:
            query += f" and phan_loai = N'{phanloai}'"  
        if ngay:
            query += f" and Ngay = '{ngay}'"   
        if tungay:
            query += f" and Ngay >= '{tungay}'"   
        if denngay:
            query += f" and Ngay <= '{denngay}'"       
        query += " order by Ngay desc"
        ## 
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay cham cong goc: {e}")
        return []
    
def lay_bangcongchot_web(masothe,chuyen,bophan,phanloai,ngay,tungay,denngay):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[BANG_CHAM_CONG_THUC_TE] where Nha_may='{current_user.macongty}' "
        if masothe:
            query += f" and MST = '{masothe}'"
        if chuyen:
            query += f" and Chuyen_to = '{chuyen}'"   
        if bophan:
            query += f" and Bo_phan = '{bophan}'"
        if phanloai:
            query += f" and phan_loai = N'{phanloai}'"  
        if ngay:
            query += f" and Ngay = '{ngay}'"   
        if tungay:
            query += f" and Ngay >= '{tungay}'"   
        if denngay:
            query += f" and Ngay <= '{denngay}'"       
        query += " order by Ngay desc"
        ## 
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay cham cong goc: {e}")
        return []
    
def kiemtra_masothe(masothe):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select Count(*) from Danh_sach_CBCNV where The_cham_cong='{masothe}' and Factory='{current_user.macongty}'"
        result = cursor.execute(query).fetchone()
        return True if result[0] > 0 else False
    except Exception as e:
        print(f"Loi kiem tra ma so the hop le: {e}")
        return False

def chucdanh_chuyen_hople(chucdanhmoi,chuyenmoi):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select Count(*) from HC_name where Detail_job_title_VN=N'{chucdanhmoi}' and Line='{chuyenmoi}'"
        result = cursor.execute(query).fetchone()
        print(result)
        return True if result[0] > 0 else False
    except Exception as e:
        print(f"Loi kiem tra hc name: {e}")
        return False

def kiemtra_thongtin_dieuchuyen(dong,masothe,chucdanhmoi,chuyenmoi,loaidieuchuyen):
    masothe_hople = kiemtra_masothe(masothe)
    print(masothe_hople)
    if not masothe_hople:
        return {"ketqua":False,
                    "dong":dong,
                    "lydo": "Mã số thẻ không hợp lệ !!!"}
    if loaidieuchuyen == "Chuyển vị trí":
        if not chucdanh_chuyen_hople(chucdanhmoi,chuyenmoi):
            return {"ketqua":False,
                    "dong":dong,
                    "lydo": "Không tìm thấy thông tin chuyền, chức danh mới trong danh sách HC Name !!!"}
    return {"ketqua":True}

def laylichsucongviec(mst,chuyen,bophan):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[LICH_SU_CONG_VIEC] where Nha_may='{current_user.macongty}' "
        if mst:
            query += f" and MST = '{mst}'"
        if chuyen:
            query += f" and Chuyen = '{chuyen}'"
        if bophan:
            query += f" and Bo_phan = '{bophan}'"
        query += " order by Tu_ngay desc, MST asc"
        data = cursor.execute(query)
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay lich su cong viec: {e}")
        return []
    
def hr_pheduyet_tangca(id,hrpheduyet):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if hrpheduyet and hrpheduyet.upper() =="OK":
            query = f"UPDATE Dang_ky_tang_ca SET HR=N'{hrpheduyet}' WHERE ID='{id}'"
        else:
            query = f"UPDATE Dang_ky_tang_ca SET HR=NULL WHERE ID='{id}'"
        cursor = cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi hr phe duyet tang ca ({e})")
        return False 
    
def lay_cacca_theobang():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[ca_lam_viec] order by Ten_ca"
        data = cursor.execute(query).fetchall()
        return [x for x in data]
    except Exception as e:
        print(f"Loi lay cac ca: {e}")
        return []

def sua_ngaybatdau_lichsu_congviec(id,ngaybatdau):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"update LICH_SU_CONG_VIEC set Tu_ngay='{ngaybatdau}' where ID='{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi cap nhat ngay bat dau lich su cong viec: {e}")
        return False
    
def sua_ngayketthuc_lichsu_congviec(id,ngayketthuc):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"update LICH_SU_CONG_VIEC set Den_ngay='{ngayketthuc}' where ID='{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi cap nhat ngay ket thuc lich su cong viec: {e}")
        return False

def sua_chuyen_lichsu_congviec(id,chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"update LICH_SU_CONG_VIEC set Chuyen='{chuyen}' where ID='{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi cap nhat chuyen lich su cong viec: {e}")
        return False
    
def sua_bophan_lichsu_congviec(id,bophan):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"update LICH_SU_CONG_VIEC set Bo_phan='{bophan}' where ID='{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi cap nhat bo phan lich su cong viec: {e}")
        return False
    
def sua_chucdanh_lichsu_congviec(id,chucdanh):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"update LICH_SU_CONG_VIEC set Chuc_danh=N'{chucdanh}' where ID='{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi cap nhat chuc danh lich su cong viec: {e}")
        return False
    
def sua_hccategory_lichsu_congviec(id,hccategory):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"update LICH_SU_CONG_VIEC set HC_CATEGORY='{hccategory}' where ID='{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi cap nhat HC_CATEGORY lich su cong viec: {e}")
        return False
    
def sua_capbac_lichsu_congviec(id,capbac):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"update LICH_SU_CONG_VIEC set Grade_code=N'{capbac}' where ID='{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi cap nhat cap bac lich su cong viec: {e}")
        return False
    
def themtaikhoanmoi(masothe,hoten,department,gradecode):
    try: 
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into Nhanvien values ('{current_user.macongty}','{masothe}',N'{hoten}','{department}','{gradecode}','user','1')"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi them tai khoan moi: {e}")
        
def lay_sodienthoai_theo_mst(mst):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select Sdt from Danh_sach_CBCNV where The_cham_cong='{mst}' and Factory='{current_user.macongty}'"
        data = cursor.execute(query).fetchone()
        return data[0]
    except Exception as e:
        print(f"Loi lay so dien thoai: {e}")
        return 0
    
def get_thongtin_vitri(vitri):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT Detail_job_title_EN,Grade_code FROM HC_Name WHERE Factory = '{current_user.macongty}' AND Detail_job_title_VN=N'{vitri}' "
        data = cursor.execute(query).fetchone()
        query1 = f"SELECT Bac_luong,Luong_co_ban FROM Luong_co_ban Where Grade_code = '{data[1]}' and Nha_may='{current_user.macongty}'"
        data1 = cursor.execute(query1).fetchall()
        cacbacluong = []
        for x in data1:
            cacbacluong.append([x[0],x[1]])
        return {'Detail_job_title_EN':data[0],'Grade_code':data[1],'Bac_luong': cacbacluong}
    except Exception as e:
        print(f"Loi lay thong tin vi tri: {e}")
        return {'Detail_job_title_EN':"",'Grade_code':"",'Bac_luong': []}
    
def lay_cac_vitri_trong_phong(phongban):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT DISTINCT Detail_job_title_VN FROM HC_Name WHERE Factory = '{current_user.macongty}' AND Department='{phongban}' "
        ##
        data = cursor.execute(query).fetchall()
        return [x[0] for x in data ]
    except Exception as e:
        print(f"Loi lay cac vi tri: {e}")
        return []

def lay_bangcongthang_kx(mst,bophan,chuyen,thang,nam):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[BANG_TONG_CONG_CA_THANG] where Nha_may='{current_user.macongty}' "
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        rows =  cursor.execute(query).fetchall()
        # print(len(rows))
        return [x for x in rows]
    except Exception as e:
        print(f"Loi lay bang cong thang: {e}")
        return []
    
def lay_bangcongthang_web(mst,bophan,chuyen,thang,nam):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[BANG_TONG_CONG_CA_THANG_THUC_TE] where Nha_may='{current_user.macongty}' "
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and Thang={thang} and Nam={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        rows =  cursor.execute(query).fetchall()
        # print(len(rows))
        return [x for x in rows]
    except Exception as e:
        print(f"Loi lay bang cong thang: {e}")
        return []
    
def capnhat_trangthai_yeucau_tuyendung(id,trangthaimoi):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung set Trang_thai_yeu_cau=N'{trangthaimoi}' where ID={id}"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return {"ketqua":True}
    except Exception as e:
        print(f"Loi cap nhat trang thai tuyen dung: {e}")
        return {"ketqua":True,"lido":e}
    
def capnhat_trangthai_tuyendung(id,trangthaimoi):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung set Trang_thai_thuc_hien=N'{trangthaimoi}' where ID={id}"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return {"ketqua":True}
    except Exception as e:
        print(f"Loi cap nhat trang thai thuc hien tuyen dung: {e}")
        return {"ketqua":True,"lido":e}
    
def capnhat_ghichu_tuyendung(id,ghichu):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung set Ghi_chu=N'{ghichu}' where ID={id}"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return {"ketqua":True}
    except Exception as e:
        print(f"Loi cap nhat trang thai thuc hien tuyen dung: {e}")
        return {"ketqua":True,"lido":e}
    
def lay_danhsach_dangky_ngayle(mst, chuyen, bophan, ngay):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[DANG_KY_LAM_VIEC_NGAY_LE] where NHA_MAY='{current_user.macongty}' "
        if mst:
            query += f" AND MST='{mst}'"
        if chuyen:
            query += f" AND CHUYEN='{chuyen}'"
        if bophan:
            query += f" AND BO_PHAN='{bophan}'"
        if ngay:
            query += f" AND NGAY_DANG_KY='{ngay}'"
        query += " order by NGAY_DANG_KY desc, MST asc"
        rows =  cursor.execute(query).fetchall()
        return [x for x in rows]
    except Exception as e:
        print(f"Loi lay bang dang ky ngay le: {e}")
        return []
    
def lay_danhsach_dangky_chunhat(mst, chuyen, bophan, ngay):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[DANG_KY_LAM_VIEC_CHU_NHAT] where NHA_MAY='{current_user.macongty}' "
        if mst:
            query += f" AND MST='{mst}'"
        if chuyen:
            query += f" AND CHUYEN='{chuyen}'"
        if bophan:
            query += f" AND BO_PHAN='{bophan}'"
        if ngay:
            query += f" AND NGAY_DANG_KY='{ngay}'"
        query += " order by NGAY_DANG_KY desc, MST asc"
        
        rows =  cursor.execute(query).fetchall()
        return [x for x in rows]
    except Exception as e:
        print(f"Loi lay bang dang ky Chu nhat: {e}")
        return []
    
def hr_pheduyet_dilam_ngayle(id,hrpheduyet:str,congkhai):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE DANG_KY_LAM_VIEC_NGAY_LE SET"
        if hrpheduyet and hrpheduyet.upper() =="OK":
            query += f" HR='OK',"
        else:
            query += f" HR=NULL,"
        if congkhai and congkhai.upper() =="OK":
            query += f" CONG_KHAI='OK' "
        else:
            query += f" CONG_KHAI=NULL "
        query += f"WHERE ID='{id}'"

        cursor = cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi hr phe duyet di lam ngay le: ({e})")
        return False
    
def hr_pheduyet_dilam_chunhat(id,hrpheduyet:str,congkhai):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE DANG_KY_LAM_VIEC_CHU_NHAT SET"
        if hrpheduyet and hrpheduyet.upper() =="OK":
            query += f" HR='OK',"
        else:
            query += f" HR=NULL,"
        if congkhai and congkhai.upper() =="OK":
            query += f" CONG_KHAI='OK' "
        else:
            query += f" CONG_KHAI=NULL "
        query += f"WHERE ID='{id}'"

        cursor = cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi hr phe duyet di lam chu nhat: ({e})")
        return False
    
def them_dangky_dilam_ngayle(nhamay,mst,hoten,chuyen,bophan,vitri,ngay):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into DANG_KY_LAM_VIEC_NGAY_LE values ('{nhamay}','{mst}',N'{hoten}','{chuyen}','{bophan}',N'{vitri}','{ngay}',NULL,NULL)"
        #
        cursor = cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi them dong di lam ngay le: ({e})")
        return False

def them_dangky_dilam_chunhat(nhamay,mst,hoten,chuyen,bophan,vitri,ngay):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into DANG_KY_LAM_VIEC_CHU_NHAT values ('{nhamay}','{mst}',N'{hoten}','{chuyen}','{bophan}',N'{vitri}','{ngay}',NULL,NULL)"
        #
        cursor = cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi them dong di lam chu nhat: ({e})")
        return False
    
def them_thongbao_co_yeucautuyendung(masothe,hoten):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into YEU_CAU_TUYEN_DUNG_CHO_PHE_DUYET values ('{current_user.macongty}','{masothe}',N'{hoten}',GETDATE())"
        
        cursor = cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi them dong di lam chu nhat: ({e})")
        return False
    
def them_yeucau_tuyendung_duoc_pheduyet(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        data = cursor.execute(f"select MST,HO_TEN,NHA_MAY from YEU_CAU_TUYEN_DUNG where ID='{id}'").fetchone()
        mst = data[0]
        hoten = data[1]
        nhamay = data[2]
        email = cursor.execute(f"select Email from KPI_DS_Email where Nha_may='{nhamay}' and MST='{mst}'").fetchone()[0]
        query = f"insert into YEU_CAU_TUYEN_DUNG_DA_PHE_DUYET values ('{nhamay}','{mst}',N'{hoten}','{email}',GETDATE())"
        
        cursor = cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi them dong phe duyet yeu cau tuyen dung: ({e})")
        return False
    
def them_yeucau_tuyendung_bi_tuchoi(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        data = cursor.execute(f"select MST,HO_TEN,NHA_MAY from YEU_CAU_TUYEN_DUNG where ID='{id}'").fetchone()
        mst = data[0]
        hoten = data[1]
        nhamay = data[2]
        email = cursor.execute(f"select Email from KPI_DS_Email where Nha_may='{nhamay}' and MST='{mst}'").fetchone()[0]
        query = f"insert into YEU_CAU_TUYEN_DUNG_BI_TU_CHOI values ('{nhamay}','{mst}',N'{hoten}','{email}',GETDATE())"
        
        cursor = cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Loi them dong tu choi yeu cau tuyen dung: ({e})")
        return False

def lay_bangcongtrangoai_web(mst,chuyen,bophan,thang,nam):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from [HR].[dbo].[BANG_CONG_TRA_NGOAI] where NHA_MAY='{current_user.macongty}' "
        if not thang:
            thang = datetime.now().month
        if not nam:
            nam =  datetime.now().year
        query += f" and THANG={thang} and NAM={nam}"
        if mst:
            query += f" and MST='{mst}'"
        if bophan:
            query += f" and Bo_phan='{bophan}'"
        if chuyen:
            query += f" and Chuyen='{chuyen}'"
        query += " order by MST asc"
        rows =  cursor.execute(query).fetchall()
        # print(len(rows))
        return [x for x in rows]
    except Exception as e:
        flash(f"Loi lay bang cong thang: {e}")
        return []
    
def lay_soluong_danglamviec():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select count(*) from [HR].[dbo].[Danh_sach_CBCNV] where Factory='{current_user.macongty}' and Trang_thai_lam_viec=N'Đang làm việc' "
        count = cursor.execute(query).fetchone()
        return count[0]
    except Exception as e:
        flash(f"Lỗi lấy số người đang làm việc: {e}")
        return 0
    
def lay_soluong_dangnghithaisan():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select count(*) from [HR].[dbo].[Danh_sach_CBCNV] where Factory='{current_user.macongty}' and Trang_thai_lam_viec=N'Nghỉ thai sản' "
        count = cursor.execute(query).fetchone()
        return count[0]
    except Exception as e:
        flash(f"Lỗi lấy số người đang nghỉ thai sản: {e}")
        return 0
    
def thaydoi_chuyen_lichsu_congtac(id,chuyenmoi):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"update Lich_su_cong_tac set Line_moi='{chuyenmoi}' where id={id}"
        cursor.execute(query)
        cursor.commit()
        conn.close()
        return {"ketqua":True}
    except Exception as e:
        return {
            "ketqua":False,
            "lido":e,
            "query":query
        }
        
def thaydoi_vitri_lichsu_congtac(id,vitrimoi):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"update Lich_su_cong_tac set Chuc_vu_moi=N'{vitrimoi}' where id={id}"
        cursor.execute(query)
        cursor.commit()
        conn.close()
        return {"ketqua":True}
    except Exception as e:
        return {
            "ketqua":False,
            "lido":e,
            "query":query
        }
        
def thaydoi_phanloai_lichsu_congtac(id,phanloaimoi):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"update Lich_su_cong_tac set Phan_loai=N'{phanloaimoi}' where id={id}"
        cursor.execute(query)
        cursor.commit()
        conn.close()
        return {"ketqua":True}
    except Exception as e:
        return {
            "ketqua":False,
            "lido":e,
            "query":query
        }
        
def thaydoi_ngaythuchien_lichsu_congtac(id,ngaythuchienmoi):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"update Lich_su_cong_tac set Ngay_thuc_hien='{ngaythuchienmoi}' where id={id}"
        cursor.execute(query)
        cursor.commit()
        conn.close()
        return {"ketqua":True}
    except Exception as e:
        return {
            "ketqua":False,
            "lido":e,
            "query":query
        }
        
def xoabo_lichsu_congtac(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"delete Lich_su_cong_tac where id={id}"
        cursor.execute(query)
        cursor.commit()
        conn.close()
        return {"ketqua":True}
    except Exception as e:
        return {
            "ketqua":False,
            "lido":e,
            "query":query
        }