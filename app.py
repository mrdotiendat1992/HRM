from const import *

# Cấu hình kết nối SQL Server
params = urllib.parse.quote_plus(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=172.16.60.100;"
    "DATABASE=HR;"
    "UID=huynguyen;"
    "PWD=Namthuan@123;"
)

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = f"mssql+pyodbc:///?odbc_connect={params}"
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
        flash("Đổi mật khẩu thành công !!!")
        return True        
    except Exception as e:
        flash(f"Doi mat khau that bai {e} !!!")
        app.logger.info(e)
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
        app.logger.info(f"Lay danh sach sap het han hop dong that bai: {e}")
        return []
    
def capnhattrangthaiyeucautuyendung(bophan,vitri,soluong,mota,thoigian,phanloai,trangthaiyeucau,trangthaithuchien,ghichu):
    try:
        trangthaiyeucau = "NULL" if not trangthaiyeucau else f"N'{trangthaiyeucau}'"            
        trangthaithuchien = "NULL" if not trangthaithuchien else f"N'{trangthaithuchien}'"
        ghichu = "NULL" if not ghichu else f"N'{ghichu}'"
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"""
            UPDATE HR.dbo.Yeu_cau_tuyen_dung
            SET Trang_thai_yeu_cau = {trangthaiyeucau}, Trang_thai_thuc_hien = {trangthaithuchien}, Ghi_chu = {ghichu}
            WHERE Bo_phan = '{bophan}' AND Vi_tri = N'{vitri}' AND So_luong = '{soluong}' AND JD = N'{mota}' AND Thoi_gian_du_kien = '{thoigian}' AND Phan_loai = N'{phanloai}'
        """ 
        cursor.execute(query)
        conn.commit()
        conn.close()
        flash("Cập nhật trạng thái ứng viên thành công !!!")
        return True
    except Exception as e:
        app.logger.info(f"Cap nhat trang thai ung vien that bai {e} !!!")
        return False

def laycatheochuyen(chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        row = cursor.execute(f"SELECT CA FROM GOI_Y_CA WHERE LINE = '{chuyen}'").fetchone()
        conn.close()
        return row[0]
    except Exception as e:
        app.logger.info(f"Loi kiem tra ca theo chuyen {e} !!!")
        return None
  
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
                    ghichu
                   ):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query1 = f"INSERT INTO HR.dbo.Lich_su_cong_tac VALUES ('{current_user.macongty}','{mst}','{chuyencu}',N'{vitricu}','{chuyenmoi}',N'{vitrimoi}',N'{loaidieuchuyen}','{ngaydieuchuyen}',N'{ghichu}','{gradecodecu}','{gradecodemoi}')"
        app.logger.info(query1)
        cursor.execute(query1)
        query2 = f"UPDATE HR.dbo.Danh_sach_CBCNV SET Job_title_VN = N'{vitrimoi}', Line = '{chuyenmoi}', Headcount_category = '{hccategorymoi}', Department = '{departmentmoi}', Section_description = '{sectiondescriptionmoi}', Emp_type = '{employeetypemoi}', Position_code_description = '{positioncodedescriptionmoi}', Section_code = '{sectioncodemoi}', Grade_code = '{gradecodemoi}', Position_code = '{positioncodemoi}', Job_title_EN = N'{vitrienmoi}', Ghi_chu = N'{ghichu}' WHERE MST = '{mst}' AND Factory = '{current_user.macongty}'"
        app.logger.info(query2)
        cursor.execute(query2)
        camoi = laycatheochuyen(chuyenmoi)
        query3 = f"""
        UPDATE HR.dbo.Dang_ky_ca_lam_viec SET Den_ngay = '{datetime.strptime(ngaydieuchuyen, '%Y-%m-%d') - timedelta(days=1)}'  WHERE MST = '{mst}' AND Factory = '{current_user.macongty}' AND Den_ngay='2054-12-31'
        INSERT INTO HR.dbo.Dang_ky_ca_lam_viec VALUES ('{mst}','{current_user.macongty}','{ngaydieuchuyen}','2054-12-31','{camoi}')
        """
        app.logger.info(query3)
        cursor.execute(query3)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
        return
    
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
            app.logger.info(e)
            return
    
def dichuyennghiviec(mst,
                    vitricu,
                    chuyencu,
                    gradecodecu,
                    ngaydieuchuyen,
                    ghichu
                   ):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        truocngaynghiviec = datetime.strptime(ngaydieuchuyen, '%Y-%m-%d') - timedelta(days=1)
        query = f"""
INSERT INTO HR.dbo.Lich_su_cong_tac VALUES ('{current_user.macongty}','{mst}',N'{chuyencu}',N'{vitricu}',NULL,NULL,N'Nghỉ việc','{ngaydieuchuyen}',N'{ghichu}','{gradecodecu}',NULL)
UPDATE HR.dbo.Danh_sach_CBCNV SET Trang_thai_lam_viec = N'Nghỉ việc', Ngay_nghi = '{ngaydieuchuyen}', Ghi_chu = N'{ghichu}' WHERE MST = '{mst}' AND Factory = '{current_user.macongty}'
            """
        # app.logger.info(query)
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
        return

def dichuyennghithaisan(mst,
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
                            ghichu
                            ):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        truocngaynghiviec = datetime.strptime(ngaydieuchuyen, '%Y-%m-%d') - timedelta(days=1)
        query = f"""
            INSERT INTO HR.dbo.Lich_su_cong_tac VALUES ('{current_user.macongty}','{mst}','{chuyencu}',N'{vitricu}',NULL,NULL,N'Nghỉ thai sản','{ngaydieuchuyen}',NULL,'{gradecodecu}',NULL)
            UPDATE HR.dbo.Danh_sach_CBCNV SET Trang_thai_lam_viec = N'Nghỉ thai sản' WHERE MST = '{mst}' AND Factory = '{current_user.macongty}'
            """    
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
        return

def dichuyenthaisandilamlai(mst,
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
                            ghichu
                            ):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        truocngaydilamlai = datetime.strptime(ngaydieuchuyen, '%Y-%m-%d') - timedelta(days=1)
        query = f"""
            INSERT INTO HR.dbo.Lich_su_cong_tac VALUES ('{current_user.macongty}','{mst}','{chuyencu}',N'{vitricu}','{chuyencu}',N'{vitricu}',N'Thai sản đi làm lại','{ngaydieuchuyen}',NULL,'{gradecodecu}','{gradecodecu}')
            UPDATE HR.dbo.Danh_sach_CBCNV SET Trang_thai_lam_viec = N'Đang làm việc',Ghi_chu=NULL WHERE MST = '{mst}' AND Factory = '{current_user.macongty}'
            """    
        # # app.logger.info(query)
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
        return
    
def inhopdongtheomau(macongty,masothe,hoten,gioitinh,ngaysinh,thuongtru,tamtru,cccd,ngaycapcccd,capbac,loaihopdong,chucdanh,phongban,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc):   
    
    try:
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
                        sheet['H24'] = f"Thử việc {songaythuviec} ngày"
                        sheet['B25'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G31'] = chucdanh
                        sheet['G42'] = f"{int(luongcoban):,} VNĐ/tháng"   
                        if phucap > 0:
                            sheet['G43'] = f"{phucap} VNĐ/tháng"
                        else:
                            sheet['G43'] = "Không"
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDTV_TO2_{masothe}{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        app.logger.info(e)
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
                        sheet['H24'] = f"Thử việc {songaythuviec} ngày"
                        sheet['B25'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G31'] = chucdanh
                        sheet['G42'] = f"{int(luongcoban):,} VNĐ/tháng"  
                        if phucap > 0:
                            sheet['G43'] = f"{phucap} VNĐ/tháng"
                        else:
                            sheet['G43'] = "Không" 
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDTV_DO2_{masothe}{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        app.logger.info(e)
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
                    app.logger.info(e)
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
                        app.logger.info(e)
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
                        app.logger.info(e)
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
                    app.logger.info(e)
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
                        sheet['B25'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G38'] = f"{luongcoban} VNĐ/tháng"
                        if phucap > 0 :
                            sheet['G39'] = f"{phucap} VNĐ/tháng"
                        else:
                            sheet['G39'] = "Không"
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDCTH_TO3_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        app.logger.info(e)
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
                        sheet['B25'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G28'] = f"{chucdanh}"
                        sheet['G38'] = f"{int(luongcoban):,} VNĐ/tháng"
                        if phucap > 0 :
                            sheet['G39'] = f"{phucap} VNĐ/tháng"
                        else:
                            sheet['G39'] = "Không"
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDCTH_DO3_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        app.logger.info(e)
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
                    app.logger.info(e)
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
                        sheet['B25'] = f"Kể từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}"
                        sheet['G38'] = f"{luongcoban} VNĐ/tháng"        
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT2_HDVTH_T03_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        app.logger.info(e)
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
                        sheet['B25'] = f"Kể từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}"   
                        sheet['G28'] = f"{chucdanh}"
                        sheet['G38'] = f"{int(luongcoban):,} VNĐ/tháng"
                        if phucap > 0 :
                            sheet['G39'] = f"{phucap} VNĐ/tháng"
                        else:
                            sheet['G39'] = "Không"    
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT2_HDVTH_D03_{masothe}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        
                        return filepath
                    except Exception as e:
                        app.logger.info(e)
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
                    app.logger.info(e)
                    return None  
    except Exception as e:
        app.logger.info(e)
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
            app.logger.info(e)
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
            app.logger.info(e)
            return None
   
def laylichsucongtac(mst,hoten,ngay,kieudieuchuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query= f"""SELECT * FROM Lich_su_cong_tac_OK WHERE Nha_may = '{current_user.macongty}' """
        if mst:
            query += f"AND MST LIKE '%{mst}%' "
        if ngay:
            query += f"AND Ngay_thuc_hien = '{ngay}' "
        if kieudieuchuyen:
            query += f"AND Phan_loai LIKE N'%{kieudieuchuyen}%' "
        if hoten:
            query += f"AND Ho_ten LIKE N'%{hoten}%' "
        query += "ORDER BY Ngay_thuc_hien DESC, CAST(MST AS INT) ASC"
        # # app.logger.info(query)
        rows = cursor.execute(query)
        result = []
        for row in rows:
            result.append({
                "Mã công ty": row[10],
                "MST": row[1],
                "Họ tên": row[0],
                "Ngày chính thức": row[2],
                "Chuyền cũ": row[3],
                "Chuyền mới": row[5] if row[5] else '',
                "Vị trí cũ": row[8],
                "Vị trí mới": row[9] if row[9] else '',
                "Phòng ban cũ": row[4],
                "Phòng ban mới": row[6] if row[6] else '',
                "Phân loại": row[10],
                "Ngày thực hiện": row[7],
                "Ghi chú": row[11] if row[1] else '',
            })
        conn.commit()
        conn.close()
        return result
    except Exception as e:
        app.logger.info(e)
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
        app.logger.info(e)
        return []
    
def lay_user(user):
    if user:
        return {
            "MST": user[0],
            "Thẻ chấm công": user[1],
            "Họ tên": user[2],
            "Số điện thoại": user[3],
            "Ngày sinh": datetime.strptime(user[4], '%Y-%m-%d').strftime("%d/%m/%Y") if user[4] else None,
            "Giới tính": user[5],
            "CCCD": user[6],
            "Ngày cấp CCCD": datetime.strptime(user[7], '%Y-%m-%d').strftime("%d/%m/%Y") if user[7] else None ,
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
            "Ngày ký HĐ": datetime.strptime(user[40], '%Y-%m-%d').strftime("%d/%m/%Y") if user[40] else None,
            "Ngày hết hạn": datetime.strptime(user[41], '%Y-%m-%d').strftime("%d/%m/%Y") if user[41] else None,
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
            "Ngày vào": datetime.strptime(user[58], '%Y-%m-%d').strftime("%d/%m/%Y"),
            "Ngày nghỉ": datetime.strptime(user[59], '%Y-%m-%d').strftime("%d/%m/%Y") if user[59] else None,
            "Trạng thái": user[60],
            "Ngày vào nối thâm niên": datetime.strptime(user[61], '%Y-%m-%d').strftime("%d/%m/%Y") if user[61] else None,
            "Mật khẩu": user[62],
            "Ngày kí HĐ Thử việc": datetime.strptime(user[63], '%Y-%m-%d').strftime("%d/%m/%Y") if user[63] else None,
            "Ngày hết hạn HĐ Thử việc": datetime.strptime(user[64], '%Y-%m-%d').strftime("%d/%m/%Y") if user[64] else None,
            "Ngày kí HĐ xác định thời hạn lần 1": datetime.strptime(user[65], '%Y-%m-%d').strftime("%d/%m/%Y") if user[65] else None,
            "Ngày hết hạn HĐ xác định thời hạn lần 1": datetime.strptime(user[66], '%Y-%m-%d').strftime("%d/%m/%Y") if user[66] else None,
            "Ngày kí HĐ HĐ xác định thời hạn lần 2": datetime.strptime(user[67], '%Y-%m-%d').strftime("%d/%m/%Y") if user[67] else None,
            "Ngày hết hạn HĐ xác định thời hạn lần 2": datetime.strptime(user[68], '%Y-%m-%d').strftime("%d/%m/%Y") if user[68] else None,
            "Ngày kí HĐ không thời hạn": datetime.strptime(user[69], '%Y-%m-%d').strftime("%d/%m/%Y") if user[69] else None,
            "Ghi chú": user[71] if user[71] else None
        }
    else:
        return None

def laydanhsachuser(mst, hoten, sdt, cccd, gioitinh, vaotungay, vaodenngay, nghitungay, nghidenngay, phongban, trangthai, hccategory,chucvu, ghichu):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Danh_sach_CBCNV WHERE Factory = '{current_user.macongty}'"
        if mst:
            query += f" AND MST LIKE '%{mst}%'"
        if hoten:
            query += f" AND Ho_ten LIKE N'%{hoten}%'"
        if sdt:
            query += f" AND SDT LIKE '%{sdt}%'"
        if cccd:
            query += f" AND CCCD LIKE '%{cccd}%'"
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
            query += f" AND Department LIKE N'%{phongban}%'"
        if trangthai:
            query += f" AND Trang_thai_lam_viec LIKE N'%{trangthai}%'"
        if hccategory:
            query += f" AND Headcount_category = '{hccategory}'"
        if chucvu:
            query += f" AND Chuc_vu LIKE N'%{chucvu}%'"
        if ghichu:
            query += f" AND Ghi_chu LIKE N'%{ghichu}%'"
        
        query += " ORDER BY CAST(mst AS INT) ASC"
        users = cursor.execute(query).fetchall()
        conn.close()
        result = []
        for user in users:
            result.append(lay_user(user))
        return result
    except Exception as e:
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
                "Ngày nhận việc": datetime.strptime(row[20], '%Y-%m-%d').strftime("%d/%m/%Y") if row[20] else None,
                "Có con nhỏ": row[21],
                "Tên con 1": row[22],
                "Ngày sinh con 1": row[23],
                "Tên con 2": row[24],
                "Ngày sinh con 2": row[25],
                "Tên con 3": row[26],
                "Ngày sinh con 3": row[27],
                "Tên con 4": row[28],
                "Ngày sinh con 4": row[29],
                "Tên con 5": row[30],
                "Ngày sinh con 5": row[31],
                "Ngày gửi": datetime.strptime(row[32], '%Y-%m-%d').strftime("%d/%m/%Y") if row[32] else None,
                "Trạng thái": row[33],
                "Ngày cập nhật": datetime.strptime(row[34], '%Y-%m-%d').strftime("%d/%m/%Y") if row[34] else None,
                "Ngày hẹn đi làm": datetime.strptime(row[35], '%Y-%m-%d').strftime("%d/%m/%Y") if row[35] else None,
                "Hiệu suất": row[36],
                "Loại máy": row[37],
                "Ghi chú": row[38], 
                "Lưu hồ sơ": row[40]
            })
        return result
    except Exception as e:
        app.logger.info(e)
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
                        ghichu
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
        Ghi_chu = N'{ghichu}'
        WHERE 
        ID = '{id}' AND Nha_may = N'{current_user.macongty}'"""
        
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        app.logger.info(e)
        return False
    
def themnhanvienmoi(nhanvienmoi):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"INSERT INTO HR.Dbo.Danh_sach_CBCNV VALUES {nhanvienmoi}"
        # app.logger.info(query)
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        app.logger.info(f"Loi insert nhan vien moi: {e}")
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
        app.logger.info(e)
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
        app.logger.info(e)
        return 0

def laydanhsachloithe(mst=None,chuyen=None, bophan=None, ngay=None):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        
        query = f"SELECT * FROM HR.dbo.Danh_sach_loi_the WHERE Nha_may = '{current_user.macongty}'"
        if mst:
            query += f"AND MST LIKE '%{mst}%' "
        if chuyen:
            query += f"AND Chuyen_to LIKE '%{chuyen}%' "
        if bophan:
            query += f"AND Bo_phan LIKE '%{bophan}%' "
        if ngay:
            query += f"AND Ngay = '{ngay}' "
        
        query += "ORDER BY CAST(MST AS INT) ASC, Ngay DESC"
        # if not chuyen and not bophan and not ngay:
        #     query = f"""
        #                 SELECT *
        #                 FROM HR.dbo.Danh_sach_loi_the
        #                 WHERE Nha_may = '{current_user.macongty}'
        #                 AND NgayCham = (
        #                     SELECT MAX(NgayCham)
        #                     FROM HR.dbo.Danh_sach_loi_the
        #                     WHERE Nha_may = '{current_user.macongty}'
        #                 )
        #                 ORDER BY NgayCham DESC;
        #             """
        
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
            "Email thư ký": row[17],
            "Email trưởng bộ phận": row[18]
            })
        return result
    except Exception as e:
        app.logger.info(e)
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
            app.logger.info(e)
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
            app.logger.info(e)
            return []

def laydanhsachchamcong(mst=None,  phongban=None, tungay=None, denngay=None, phanloai=None):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Bang_cham_cong_tu_dong WHERE Nha_may = '{current_user.macongty}'"
        if mst: 
            query += f" AND MST LIKE '%{mst}%'"
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
            app.logger.info(e)
            return []

def laydanhsachchamcongchot(mst=None, phongban=None, tungay=None, denngay=None, phanloai=None):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Bang_cham_cong WHERE Nha_may = '{current_user.macongty}'"
        if mst: 
            query += f" AND MST LIKE '%{mst}%'"
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
        app.logger.info(e)
        return []

def laydanhsachdiemdanhbu(mst=None,hoten=None,chucvu=None,chuyen=None,bophan=None,loaidiemdanh=None,ngaydiemdanh=None,lido=None,trangthai=None,mstquanly=None):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        
        if mstquanly:
            query = f"""
            SELECT  *
            FROM 
                Diem_danh_bu 
            INNER JOIN 
                Phan_quyen_thu_ky
            ON
                Diem_danh_bu.Nha_may= Phan_quyen_thu_ky.Nha_may and Diem_danh_bu.Line=Phan_quyen_thu_ky.Chuyen_to
            WHERE 
                Diem_danh_bu.Trang_thai=N'Đã kiểm tra' and MST_QL='{mstquanly}'"""
        else:
            query = f"SELECT * FROM HR.dbo.Diem_danh_bu WHERE Nha_may = '{current_user.macongty}' "   
            if mst:
                query += f"AND MST LIKE '%{mst}%' "
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
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        app.logger.info(e)
        return []

def laydanhsachxinnghiphep(mst,hoten,chucvu,chuyen,bophan,ngaynghi,lydo,trangthai,mstquanly):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if mstquanly:
            query = f"""
            SELECT 
                *
            FROM 
                Xin_nghi_phep 
            INNER JOIN 
                Phan_quyen_thu_ky
            ON
                Xin_nghi_phep.Nha_may= Phan_quyen_thu_ky.Nha_may and Xin_nghi_phep.Line=Phan_quyen_thu_ky.Chuyen_to
            WHERE 
                Xin_nghi_phep.Trang_thai=N'Đã kiểm tra' and MST_QL='{mstquanly}'"""
        else:            
            query = f"SELECT * FROM HR.dbo.DS_Xin_nghi_phep WHERE Nha_may = '{current_user.macongty}' "
            if mst:
                query += f"AND MST LIKE '%{mst}%'"
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
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        app.logger.info(e)
        return []

def laydanhsachxinnghikhongluong(mst,hoten,chucvu,chuyen,bophan,ngay,lydo,trangthai,mstquanly):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
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
                Xin_nghi_khong_luong.Trang_thai=N'Đã kiểm tra' and MST_QL='{mstquanly}'"""
        else:
            
            query = f"SELECT * FROM HR.dbo.Xin_nghi_khong_luong WHERE Nha_may = '{current_user.macongty}' "
            if mst:
                query += f"AND MST LIKE '%{mst}%'"
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
        app.logger.info(e)
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
        app.logger.info(e)
        return False

def quanly_duoc_phanquyen(mst,chuyen):
    try:
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
        app.logger.info(e)
        return False
    
def kiemtrathuki(mst,chuyen):
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
        app.logger.info(e)
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
        app.logger.info(e)

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
        app.logger.info(e)
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
       app.logger.info(e)  
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
        app.logger.info(e)   
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
       app.logger.info(e)  
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
        app.logger.info(e)

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
        app.logger.info(f"Kiem tra da dang ky ca lam viec chua loi: {e} !!!")
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
            app.logger.info("Co ngay ket thuc")
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
                flash("Đổi ca thành công", "success")
                return True
            else:
                flash("Đổi ca thất bại, ngày bắt đầu lớn hơn ngày kết thúc")
        else:
            app.logger.info("Khong co ngay ket thuc")
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
            flash("Đổi ca thành công", "success")
            return True
    except Exception as e:
       app.logger.info(e)   
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
        app.logger.info(e)
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
        app.logger.info(e)
        return []
    
def themyeucautuyendungmoi(bophan,vitri,soluong,mota,thoigiandukien,phanloai, mucluong):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"INSERT INTO HR.dbo.Yeu_cau_tuyen_dung VALUES('{bophan}',N'{vitri}','{soluong}',N'{mota}','{thoigiandukien}',N'{phanloai}',N'{mucluong}',NULL,NULL,NULL)"
        
        cursor.execute(query)
        conn.commit()
        return True
    except Exception as e:
        app.logger.info(e)
        return False
    
def laydanhsachxinnghikhac(mst=None,ngaynghi=None,loainghi=None):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT *  FROM [HR].[dbo].Xin_nghi_khac where Nha_may='{current_user.macongty}' "
        
        if mst:
            query += f"AND MST='{mst}'" 
        if ngaynghi:
            query += f"AND Ngay_nghi = '{ngaynghi}'" 
        if loainghi:
            query += f"AND Loai_nghi = N'{loainghi}'"
            
        query += " ORDER BY Ngay_nghi DESC, MST ASC"
        
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows 
    except Exception as e:
        app.logger.info(e)

def themxinnghikhac(macongty,mst,ngaynghi,tongsophut,loainghi):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"INSERT INTO [HR].[dbo].Xin_nghi_khac VALUES ('{macongty}','{mst}','{ngaynghi}','{tongsophut}',N'{loainghi}')"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)

def xoadulieuchamcong2ngay():
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"Delete from Check_in_out where CONVERT(varchar, NgayCham, 23) >= CONVERT(varchar, DATEADD(DAY, -2, GETDATE()), 23) and Nha_may = 'NT1';"
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        app.logger.info(e)
        return False

def themdulieuchamcong2ngay():
    try:
        xoadulieuchamcong2ngay()
        conn = pyodbc.connect(mccdb)
        cursor = conn.cursor()
        query = f"""
	    select 'NT1',MaChamCong,NgayCham,GioCham,TenMay from checkinout 
	    where CONVERT(varchar, NgayCham, 23) >= CONVERT(varchar, DATEADD(DAY, -2, GETDATE()), 23)"""
        rows = cursor.execute(query).fetchall()
        conn.close()
        conn1 = pyodbc.connect(used_db)
        cursor1 = conn1.cursor()
        query1 = "insert into Check_in_out(Nha_may,Machamcong,NgayCham,GioCham,TenMay) values(?,?,?,?,?)"
        for row in rows:
            cursor1.execute(query1, row)
        conn1.commit()
        conn1.close()
        return True
    except Exception as e:
        app.logger.info(e)
        return False
    
def thuky_dakiemtra_diemdanhbu(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Diem_danh_bu SET Trang_thai = N'Đã kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = N'{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
    
def thuky_tuchoi_diemdanhbu(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Diem_danh_bu SET Trang_thai = N'Bị từ chối bởi người kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = N'{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)

def quanly_pheduyet_diemdanhbu(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Diem_danh_bu SET Trang_thai = N'Đã phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = N'{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
    
def quanly_tuchoi_diemdanhbu(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Diem_danh_bu SET Trang_thai = N'Bị từ chối bởi người phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = N'{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)

def thuky_dakiemtra_xinnghiphep(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_phep SET Trang_thai = N'Đã kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = N'{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
    
def thuky_tuchoi_xinnghiphep(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_phep SET Trang_thai = N'Bị từ chối bởi người kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = N'{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)

def quanly_pheduyet_xinnghiphep(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_phep SET Trang_thai = N'Đã phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = N'{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
    
def quanly_tuchoi_xinnghiphep(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_phep SET Trang_thai = N'Bị từ chối bởi người phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = N'{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
        
def thuky_dakiemtra_xinnghikhongluong(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khong_luong SET Trang_thai = N'Đã kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = N'{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
    
def thuky_tuchoi_xinnghikhongluong(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khong_luong SET Trang_thai = N'Bị từ chối bởi người kiểm tra' WHERE Nha_may = '{current_user.macongty}' AND ID = N'{id}'"
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)

def quanly_pheduyet_xinnghikhongluong(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khong_luong SET Trang_thai = N'Đã phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = N'{id}'"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
    
def quanly_tuchoi_xinnghikhongluong(id):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"UPDATE HR.dbo.Xin_nghi_khong_luong SET Trang_thai = N'Bị từ chối bởi người phê duyệt' WHERE Nha_may = '{current_user.macongty}' AND ID = N'{id}'"
        # app.logger.info(query)
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.error(e)
        
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
            Dang_ky_ca_lam_viec.Den_ngay
        FROM 
            Dang_ky_ca_lam_viec
        INNER JOIN 
            Danh_sach_CBCNV 
        ON 
            Dang_ky_ca_lam_viec.MST = Danh_sach_CBCNV.The_cham_cong AND Dang_ky_ca_lam_viec.Factory = Danh_sach_CBCNV.Factory
        WHERE 
            Dang_ky_ca_lam_viec.Factory = '{current_user.macongty}' AND Danh_sach_CBCNV.Trang_thai_lam_viec=N'Đang làm việc'
        """
        if mst:
            query += f" AND Dang_ky_ca_lam_viec.MST = '{mst}'"
        if chuyen:
            query += f" AND Danh_sach_CBCNV.Line LIKE '%{chuyen}%'"
        if phongban:
            query += f" AND Danh_sach_CBCNV.Department LIKE '%{phongban}%'"
        query += "ORDER BY Dang_ky_ca_lam_viec.Tu_ngay desc, Dang_ky_ca_lam_viec.Den_ngay desc, MST asc"
        # app.logger.info(query)
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        app.logger.error(e)
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
        # app.logger.info(query)
        rows = cursor.execute(query).fetchall()
        conn.close()
        if not mst:
            return []
        return rows
    except Exception as e:
        app.logger.error(e)
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
        app.logger.inerrorfo(e)
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
        app.logger.error(e)
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
        # app.logger.info(query)
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        app.logger.error(e)
        return False
    
def guimailthongbaodaguikpi(nhamay,mst,hoten):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into KPI_Cho_phe_duyet values ('{nhamay}','{mst}',N'{hoten}',GETDATE())"
        # # app.logger.info(query)  
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.error(e)
        return False
    
def guimailthongbaodapheduyetkpi(nhamay,mst,hoten,email):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into KPI_Da_phe_duyet values ('{nhamay}','{mst}',N'{hoten}','{email}',GETDATE())"  
        # # app.logger.info(query)
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
        return False
    
def guimailthongbaodatuchoikpi(nhamay,mst,hoten,email):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into KPI_Bi_tu_choi values ('{nhamay}','{mst}',N'{hoten}','{email}',GETDATE())"  
        # # app.logger.info(query)
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        # # app.logger.info(query)
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        app.logger.info(e)
        return []

def themdonxinnghi(mst,hoten,chucdanh,chuyen,phongban,ngaynopdon,ngaynghi,ghichu):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        if ghichu:
            query = f"INSERT INTO Cho_nghi_viec VALUES ('{current_user.macongty}','{mst}',N'{hoten}',N'{chucdanh}','{chuyen}','{phongban}','{ngaynopdon}','{ngaynghi}',N'{ghichu}')"
        else:
            query = f"INSERT INTO Cho_nghi_viec VALUES ('{current_user.macongty}','{mst}',N'{hoten}',N'{chucdanh}','{chuyen}','{phongban}','{ngaynopdon}','{ngaynghi}',NULL)"
        # # app.logger.info(query)
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(e)
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
        app.logger.info(f"Lay thong tin hop dong loi {e} !!!")
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
        app.logger.info(f"Loi khi tim kiem cac chuc danh: {e} !!!")
        return []
    
def themhopdongmoi(nhamay,mst,hoten,gioitinh,ngaysinh,thuongtru,tamtru,cccd,ngaycapcccd,capbac,loaihopdong,chucdanh,phongban,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"""
        INSERT INTO QUAN_LY_HD VALUES (
            '{nhamay}', '{int(mst)}', N'{hoten}', N'{gioitinh}', '{ngaysinh}', N'{thuongtru}', N'{tamtru}', '{cccd}', '{ngaycapcccd}', '{capbac}',
            N'{loaihopdong}', N'{chucdanh}', '{phongban}', '{chuyen}', '{luongcoban}', '{phucap}', '{ngaybatdau}', '{ngayketthuc}')
        """
        print(query)
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        app.logger.info(f"Loi khi them hop dong: {e} !!!")
        query = f"""
        INSERT INTO QUAN_LY_HD VALUES (
            '{nhamay}', '{int(mst)}', N'{hoten}', N'{gioitinh}', '{ngaysinh}', N'{thuongtru}', N'{tamtru}', '{cccd}', '{ngaycapcccd}', '{capbac}',
            N'{loaihopdong}', N'{chucdanh}', '{phongban}', '{chuyen}', '{luongcoban}', '{phucap}', '{ngaybatdau}', NULL )
        """
        print(query)
        cursor.execute(query)
        try:
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            app.logger.info(f"Loi khi them hop dong: {e} !!!")   
            return False
    
def capnhatthongtinhopdong(nhamay,mst,loaihopdong,chucdanh,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc,vitrien,employeetype,posotioncode,postitioncodedescription,hccategory,sectioncode,sectiondescription):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        
        # if loaihopdong == "Hợp đồng thử việc":
        #     query = f"""
        #         UPDATE Danh_sach_CBCNV SET Loai_hop_dong=N'{loaihopdong}', Luong_co_ban='{luongcoban}', Phu_cap='{phucap}', Ngay_ky_HDTV='{ngaybatdau}', Ngay_het_han_HDTV='{ngayketthuc}',
        #         Job_title_VN=N'{chucdanh}', Job_title_EN='{vitrien}', Emp_type='{employeetype}', Position_code='{posotioncode}', Position_code_description='{postitioncodedescription}',
        #         Headcount_category='{hccategory}', Section_code='{sectioncode}', Section_description='{sectiondescription}', Line=N'{chuyen}'
        #         WHERE Factory='{nhamay}' AND MST='{mst}'
        #         """
        # el
        if loaihopdong == "Phụ lục hợp đồng":
            query = f"""
                UPDATE Danh_sach_CBCNV SET Luong_co_ban='{luongcoban}', Phu_cap='{phucap}',
                Job_title_VN=N'{chucdanh}', Job_title_EN='{vitrien}', Emp_type='{employeetype}', Position_code='{posotioncode}', Position_code_description='{postitioncodedescription}',
                Headcount_category='{hccategory}', Section_code='{sectioncode}', Section_description='{sectiondescription}'
                WHERE Factory='{nhamay}' AND MST='{mst}'
                """
        elif loaihopdong == "Hợp đồng có thời hạn 28 ngày" or loaihopdong == "Hợp đồng có thời hạn 1 năm":
            query = f"""
                UPDATE Danh_sach_CBCNV SET Luong_co_ban='{luongcoban}', Phu_cap='{phucap}', Ngay_ky_HDXDTH_Lan1='{ngaybatdau}', Ngay_het_han_HDXDTH_Lan1='{ngayketthuc}',
                Job_title_VN=N'{chucdanh}', Job_title_EN='{vitrien}', Emp_type='{employeetype}', Position_code='{posotioncode}', Position_code_description='{postitioncodedescription}',
                Headcount_category='{hccategory}', Section_code='{sectioncode}', Section_description='{sectiondescription}', Loai_hop_dong=N'{loaihopdong}'
                WHERE Factory='{nhamay}' AND MST='{mst}'
                """
        elif loaihopdong == "Hợp đồng vô thời hạn":
            query = f"""
                UPDATE Danh_sach_CBCNV SET Luong_co_ban='{luongcoban}', Phu_cap='{phucap}', Ngay_ky_HDKXDTH='{ngaybatdau}',
                Job_title_VN=N'{chucdanh}', Job_title_EN='{vitrien}', Emp_type='{employeetype}', Position_code='{posotioncode}', Position_code_description='{postitioncodedescription}',
                Headcount_category='{hccategory}', Section_code='{sectioncode}', Section_description='{sectiondescription}', Loai_hop_dong=N'{loaihopdong}'
                WHERE Factory='{nhamay}' AND MST='{mst}'"""
        else:
            query = None
        if query:
            print(query)
            cursor.execute(query)
            conn.commit()
            conn.close()  
            return True
        else:
            print("Khong hieu loai hop dong")    
            return False
    except Exception as e:
        app.logger.info(f"Loi khi cap nhat thong tin hop dong: {e} !!!")
        return False
    
def thaydoithongtinhopdong(id,masothe,hoten,gioitinh,ngaysinh,thuongtru,tamtru,cccd,ngaycapcccd,
                           loaihopdong,ngaybatdau,ngayketthuc,chuyen,capbac,chucdanh,phongban,luongcoban,phucap):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        ngayketthuc = f"'{ngayketthuc}'" if ngayketthuc else 'NULL'
        query = f"""
        update QUAN_LY_HD set MST='{masothe}',HO_TEN=N'{hoten}',GIOI_TINH='{gioitinh}',NGAY_SINH='{ngaysinh}',DIA_CHI=N'{thuongtru}',TAM_TRU=N'{tamtru}',
        CCCD='{cccd}',NGAY_CAP='{ngaycapcccd}',LOAI_HD=N'{loaihopdong}',CHUC_DANH=N'{chucdanh}',PHONG_BAN='{phongban}',CHUYEN='{chuyen}',CAP_BAC='{capbac}',
        LCB='{luongcoban}',PHU_CAP='{phucap}',NGAY_KY='{ngaybatdau}',NGAY_HET_HAN={ngayketthuc} where ID='{id}'
        """
        # print(query)
        cursor.execute(query)
        conn.commit()
        conn.close()  
        return True
    except Exception as e:
        app.logger.info(f"Loi khi cap nhat thong tin hop dong: {e} !!!")
        return False
    
def them_diemdanhbu(masothe,hoten,chucdanh,chuyen,phongban,loaidiemdanh,ngay,giovao,lydo,trangthai):
    try:
        ngay = ngay.split("/")[2] + "-" + ngay.split("/")[1] + "-" + ngay.split("/")[0]
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into Diem_danh_bu values ('{current_user.macongty}','{masothe}',N'{hoten}',N'{chucdanh}','{chuyen}','{phongban}',N'{loaidiemdanh}','{ngay}','{giovao}',N'{lydo}',N'{trangthai}')"
        print(query)
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        app.logger.info(f"Loi khi them diem danh bu: {e} !!!")
        return False
    
def them_xinnghiphep(masothe,hoten,chucdanh,chuyen,phongban,ngay,sophut,trangthai):
    try:
        ngay = ngay.split("/")[2] + "-" + ngay.split("/")[1] + "-" + ngay.split("/")[0]
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into Xin_nghi_phep values ('{current_user.macongty}','{masothe}',N'{hoten}',N'{chucdanh}','{chuyen}','{phongban}','{ngay}','{sophut}',NULL,N'{trangthai}')"
        print(query)
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        app.logger.info(f"Loi khi them xin nghi phep: {e} !!!")
        return False
    
def them_xinnghikhongluong(masothe,hoten,chucdanh,chuyen,phongban,ngay,sophut,lydo,trangthai):
    try:
        ngay = ngay.split("/")[2] + "-" + ngay.split("/")[1] + "-" + ngay.split("/")[0]
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into Xin_nghi_khong_luong values ('{current_user.macongty}','{masothe}',N'{hoten}',N'{chucdanh}','{chuyen}','{phongban}','{ngay}','{sophut}',N'{lydo}',N'{trangthai}')"
        print(query)
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        app.logger.info(f"Loi khi them xin nghi khong luong: {e} !!!")
        return False
laydanhsach_chonghiviec
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
        app.logger.info(f"Loi khi kiem tra co phai to truong khong: {e} !!!")
        return None