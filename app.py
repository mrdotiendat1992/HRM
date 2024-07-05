from const import *

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///db.sqlite"
app.config["SECRET_KEY"] = "hrm_system_NT"
app.config['UPLOAD_FOLDER'] = r'./static/uploads'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SQLALCHEMY_POOL_SIZE'] = 10
app.config['SQLALCHEMY_MAX_OVERFLOW'] = 20
app.config['SQLALCHEMY_POOL_TIMEOUT'] = 60
db = SQLAlchemy()

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

test_db = r"Driver={SQL Server};Server=DESKTOP-G635SF6;Trusted_Connection=yes;"
used_db = r"Driver={SQL Server};Server=172.16.60.100;Database=HR;UID=huynguyen;PWD=Namthuan@123;"
mccdb = r"Driver={SQL Server}; Server=10.0.0.252\SQLEXPRESS; Database=MITACOSQL; UID=sa;PWD=Namthuan1;"

class Users(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    masothe = db.Column(db.String(250), nullable=False)
    hoten = db.Column(db.String(250), nullable=False)
    macongty = db.Column(db.String(250), nullable=False)
    tencongty = db.Column(db.String(250), nullable=False)
    phongban = db.Column(db.String(250), nullable=False)
    matkhau = db.Column(db.String(250), nullable=False)      
    role = db.Column(db.String(250), nullable=False)  
 
db.init_app(app)
 
with app.app_context():
    try:
        db.create_all()
    except Exception as e:
        app.logger.info(e)
    
@login_manager.user_loader
def loader_user(user_id):
    try:
        return db.session.get(Users, int(user_id))
    except Exception as e:
        app.logger.info(e)
        return None
  
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
        return True
    except Exception as e:
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
        app.logger.info(e)
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
    except Exception as e:
        app.logger.info(e)
        return

def laycatheochuyen(chuyen):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select Ca_mac_dinh from Ca_theo_chuyen where Chuyen='{chuyen}'"
        row = cursor.execute(query).fetchone()
        conn.close()
        return row[0]
    except Exception as e:
        app.logger.info(e)
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
        query1 = f"INSERT INTO HR.dbo.Lich_su_cong_tac VALUES ('{current_user.macongty}','{mst}','{chuyencu}',N'{vitricu}','{chuyenmoi}',N'{vitrimoi}',N'{loaidieuchuyen}','{ngaydieuchuyen}',N'{ghichu}')"
        app.logger.info(query1)
        cursor.execute(query1)
        query2 = f"UPDATE HR.dbo.Danh_sach_CBCNV SET Job_title_VN = N'{vitrimoi}', Line = '{chuyenmoi}', Headcount_category = '{hccategorymoi}', Department = '{departmentmoi}', Section_description = '{sectiondescriptionmoi}', Emp_type = '{employeetypemoi}', Position_code_description = '{positioncodedescriptionmoi}', Section_code = '{sectioncodemoi}', Grade_code = '{gradecodemoi}', Position_code = '{positioncodemoi}', Job_title_EN = N'{vitrienmoi}', Ghi_chu = N'{ghichu}' WHERE MST = '{mst}' AND Factory = '{current_user.macongty}'"
        app.logger.info(query2)
        cursor.execute(query2)
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
                    ngaydieuchuyen,
                    ghichu
                   ):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        ngaynghiviec = datetime.strptime(ngaydieuchuyen, '%Y-%m-%d') + timedelta(days=1)
        query = f"""
INSERT INTO HR.dbo.Lich_su_cong_tac VALUES ('{current_user.macongty}','{mst}',N'{chuyencu}',N'{vitricu}',NULL,NULL,N'Nghỉ việc','{ngaydieuchuyen}',N'{ghichu}')
UPDATE HR.dbo.Danh_sach_CBCNV SET Trang_thai_lam_viec = N'Nghỉ việc', Ngay_nghi = '{ngaydieuchuyen}', Ghi_chu = N'{ghichu}' WHERE MST = '{mst}' AND Factory = '{current_user.macongty}'
UPDATE HR.dbo.Lich_su_trang_thai_lam_viec SET Den_ngay = '{ngaydieuchuyen}' WHERE MST = '{mst}' AND Nha_may = '{current_user.macongty}' AND Den_ngay = '2054-12-31'
INSERT INTO HR.dbo.Lich_su_trang_thai_lam_viec VALUES ('{mst}','{current_user.macongty}','{ngaynghiviec}','2054-12-31',N'Nghỉ việc')
            """
        
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
        ngaynghiviec = datetime.strptime(ngaydieuchuyen, '%Y-%m-%d') + timedelta(days=1)
        query = f"""
            INSERT INTO HR.dbo.Lich_su_cong_tac VALUES ('{current_user.macongty}','{mst}','{chuyencu}',N'{vitricu}',NULL,NULL,N'Nghỉ thai sản','{ngaydieuchuyen}',NULL)
            UPDATE HR.dbo.Danh_sach_CBCNV SET Trang_thai_lam_viec = N'Nghỉ thai sản' WHERE MST = '{mst}' AND Factory = '{current_user.macongty}'
            UPDATE HR.dbo.Lich_su_trang_thai_lam_viec SET Den_ngay = '{ngaydieuchuyen}' WHERE MST = '{mst}' AND Nha_may = '{current_user.macongty}' AND Den_ngay = '2054-12-31'
            INSERT INTO HR.dbo.Lich_su_trang_thai_lam_viec VALUES ('{mst}','{current_user.macongty}','{ngaynghiviec}','2054-12-31',N'Nghỉ thai sản')
            """    
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
        return
    
def inhopdongtheomau(kieuhopdong,
                           mst,
                           ngaylamhopdong,
                           thanglamhopdong,
                           namlamhopdong,
                           ngayketthuchopdong,
                           thangketthuchopdong,
                           namketthuchopdong,
                           tennhanvien,
                           ngaysinh,
                           gioitinh,
                           thuongtru,
                           tamtru,
                           cccd,
                           ngaycapcccd,
                           mucluong,
                           phucap,
                           tongphucap,
                           chucvu,
                           bophan,
                           capbac,
                           songaythuviec):   
    
    try:
        if kieuhopdong == "HĐ thử việc":
            if current_user.macongty == "NT1":
                if capbac in ["O2","O1","M3","M2","M1"]:
                    try:
                        workbook = openpyxl.load_workbook(FILE_MAU_HDTV_NT1_O2_TROLEN)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: PC/{mst}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = tennhanvien
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['B21'] = f"Số CCCD: {cccd}"
                        sheet['L21'] = ngaycapcccd
                        sheet['H24'] = f"Thử việc {songaythuviec} ngày"
                        sheet['B25'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G31'] = chucvu
                        sheet['G42'] = f"{int(mucluong):,} VNĐ/tháng"   
                        if phucap == "Có":
                            sheet['G43'] = f"{tongphucap} VNĐ/tháng"
                        else:
                            sheet['G43'] = "Không"
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDTV_TO3_{mst}{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        app.logger.info(e)
                        return None
                else:
                    try:
                        workbook = openpyxl.load_workbook(FILE_MAU_HDTV_NT1_DUOI_O2)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: PC/{mst}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = tennhanvien
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['E21'] = f"{cccd}"
                        sheet['L21'] = ngaycapcccd
                        sheet['H24'] = f"Thử việc {songaythuviec} ngày"
                        sheet['B25'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G31'] = chucvu
                        sheet['G42'] = f"{int(mucluong):,} VNĐ/tháng"  
                        if phucap == "Có":
                            sheet['G43'] = f"{tongphucap} VNĐ/tháng"
                        else:
                            sheet['G43'] = "Không" 
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDTV_DO3_{mst}{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        app.logger.info(e)
                        return None
            elif current_user.macongty == "NT2":
                try:
                    workbook = openpyxl.load_workbook(FILE_MAU_HDTV_NT2)
                    sheet = workbook.active
                    sheet['E4'] = f'Số: PC/{mst}'
                    sheet['M4'] = f'Nghệ An, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                    sheet['D19'] = tennhanvien
                    sheet['E21'] = ngaysinh
                    sheet['E20'] = gioitinh
                    sheet['F22'] = thuongtru
                    sheet['F23'] = tamtru
                    sheet['D24'] = f"{cccd}"
                    sheet['L24'] = ngaycapcccd
                    sheet['H27'] = f"{songaythuviec} ngày"
                    sheet['B28'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                    sheet['G46'] = f"{int(mucluong):,} VNĐ/tháng"    
                    sheet['F33'] = chucvu  
                    sheet['F34'] = bophan
                    sheet['F35'] = f"{mst}"
                    if phucap == "Có":
                        sheet['G47'] = f"{tongphucap} VNĐ/tháng"
                    else:
                        sheet['G47'] = "Không" 
                    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                    filepath = os.path.join(FOLDER_XUAT, f'NT2_HDTV_{mst}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                    workbook.save(filepath)
                    return filepath
                except Exception as e:
                    app.logger.info(e)
                    return None
        elif kieuhopdong == "HĐ có thời hạn lần 1" or kieuhopdong == "HĐ có thời hạn lần 2":
            if current_user.macongty == "NT1":
                if capbac in ["O2","O1","M3","M2","M1"]:
                    try:
                        workbook = openpyxl.load_workbook(FILE_MAU_HDCTH_NT1_O2_TROLEN)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: LC12/{mst}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = tennhanvien
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['B21'] = f"Số CCCD:{cccd}"
                        sheet['L21'] = ngaycapcccd
                        sheet['B25'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G38'] = f"{mucluong} VNĐ/tháng"
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDCTH_TO3_{mst}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        app.logger.info(e)
                        return None
                else:
                    try:
                        workbook = openpyxl.load_workbook(FILE_MAU_HDCTH_NT1_DUOI_O2)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: LC12/{mst}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = tennhanvien
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['B21'] = f"Số CCCD: {cccd}"
                        sheet['L21'] = ngaycapcccd
                        sheet['B25'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong} đến hết ngày {ngayketthuchopdong} tháng {thangketthuchopdong} năm {namketthuchopdong}"
                        sheet['G28'] = f"{chucvu}"
                        sheet['G38'] = f"{int(mucluong):,} VNĐ/tháng"
                        if phucap == "Có":
                            sheet['G39'] = f"{tongphucap} VNĐ/tháng"
                        else:
                            sheet['G39'] = "Không"
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT1_HDCTH_DO3_{mst}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        app.logger.info(e)
                        return None
            elif current_user.macongty == "NT2":
                try:
                    workbook = openpyxl.load_workbook(FILE_MAU_HDCTH_NT2)
                    sheet = workbook.active
                    sheet['E4'] = f'Số: LC12/{mst}'
                    sheet['M4'] = f'Nghệ An, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                    sheet['D19'] = tennhanvien
                    sheet['E21'] = ngaysinh
                    sheet['E20'] = gioitinh
                    sheet['F22'] = thuongtru
                    sheet['F23'] = tamtru
                    sheet['D24'] = cccd
                    sheet['L24'] = ngaycapcccd
                    sheet['F32'] = bophan
                    sheet['F33'] = mst
                    sheet['G43'] = f"{int(mucluong):,} VNĐ/tháng"
                    if phucap == "Có":
                        sheet['G44'] = f"{tongphucap} VNĐ/tháng"
                    else:
                        sheet['G44'] = "Không" 
                    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                    filepath = os.path.join(FOLDER_XUAT, f'NT2_HDCTH_{mst}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                    workbook.save(filepath)
                    return filepath
                except Exception as e:
                    app.logger.info(e)
                    return None
                
        elif kieuhopdong == "HĐ vô thời hạn":
            if current_user.macongty == "NT1":
                if capbac in ["O2","O1","M3","M2","M1"]:
                    try:
                        workbook = openpyxl.load_workbook(FILE_MAU_HDVTH_NT1_O2_TROLEN)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: LC/{mst}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = tennhanvien
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['B21'] = f"Số CCCD: {cccd}"
                        sheet['L21'] = ngaycapcccd
                        sheet['B25'] = f"Kể từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}"
                        sheet['G38'] = f"{mucluong} VNĐ/tháng"        
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT2_HDVTH_T03_{mst}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        return filepath
                    except Exception as e:
                        app.logger.info(e)
                        return None   
                else:
                    try:
                        workbook = openpyxl.load_workbook(FILE_MAU_HDVTH_NT1_DUOI_O2)
                        sheet = workbook.active
                        sheet['E4'] = f'Số: LC/{mst}'
                        sheet['M4'] = f'Hải Phòng, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                        sheet['D18'] = tennhanvien
                        sheet['E19'] = ngaysinh
                        sheet['Q19'] = gioitinh
                        sheet['F20'] = thuongtru
                        sheet['B21'] = f"Số CCCD: {cccd}"
                        sheet['L21'] = ngaycapcccd
                        sheet['B25'] = f"Kể từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}"   
                        sheet['G28'] = f"{chucvu}"
                        sheet['G38'] = f"{int(mucluong):,} VNĐ/tháng"
                        if phucap == "Có":
                            sheet['G39'] = f"{tongphucap} VNĐ/tháng"
                        else:
                            sheet['G39'] = "Không"    
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                        filepath = os.path.join(FOLDER_XUAT, f'NT2_HDVTH_D03_{mst}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
                        workbook.save(filepath)
                        
                        return filepath
                    except Exception as e:
                        app.logger.info(e)
                        return None  
            if current_user.macongty == "NT2":
                try:
                    workbook = openpyxl.load_workbook(FILE_MAU_HDVTH_NT2)
                    sheet = workbook.active
                    sheet['E4'] = f'Số: LC/{mst}'
                    sheet['M4'] = f'Nghệ An, ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}'
                    sheet['D19'] = tennhanvien
                    sheet['E21'] = ngaysinh
                    sheet['E20'] = gioitinh
                    sheet['F22'] = thuongtru
                    sheet['F23'] = tamtru
                    sheet['D24'] = cccd
                    sheet['L24'] = ngaycapcccd
                    sheet['B28'] = f"Từ ngày {ngaylamhopdong} tháng {thanglamhopdong} năm {namlamhopdong}"
                    sheet['G43'] = f"{int(mucluong):,} VNĐ/tháng"  
                    if phucap == "Có":
                        sheet['G44'] = f"{tongphucap} VNĐ/tháng"
                    else:
                        sheet['G44'] = "Không"      
                    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")     
                    filepath = os.path.join(FOLDER_XUAT, f'NT2_HDVTH_{mst}_{ngaylamhopdong}{thanglamhopdong}{namlamhopdong}_{thoigian}.xlsx')
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
        query= f"""SELECT 
                Lich_su_cong_tac.Nha_may,
                Lich_su_cong_tac.MST,
                Danh_sach_CBCNV.Ho_ten,
                Lich_su_cong_tac.Chuc_vu_cu,
                Lich_su_cong_tac.Line_cu,
                Lich_su_cong_tac.Chuc_vu_moi,
                Lich_su_cong_tac.Line_moi,
                Lich_su_cong_tac.Phan_loai,
                Lich_su_cong_tac.Ngay_thuc_hien,
                Lich_su_cong_tac.Ghi_chu
            FROM 
                Lich_su_cong_tac
            INNER JOIN 
                Danh_sach_CBCNV 
            ON 
                Lich_su_cong_tac.MST = Danh_sach_CBCNV.MST
            WHERE 
                Lich_su_cong_tac.Nha_may = '{current_user.macongty}' """
        if mst:
            query += f"AND Lich_su_cong_tac.MST LIKE '%{mst}%' "
        if ngay:
            query += f"AND Lich_su_cong_tac.Ngay_thuc_hien = '{ngay}' "
        if kieudieuchuyen:
            query += f"AND Lich_su_cong_tac.Phan_loai LIKE N'%{kieudieuchuyen}%' "
        if hoten:
            query += f"AND Danh_sach_CBCNV.Ho_ten LIKE N'%{hoten}%' "
        query += "ORDER BY Lich_su_cong_tac.Ngay_thuc_hien DESC, CAST(Lich_su_cong_tac.MST AS INT) ASC, Lich_su_cong_tac.Line_moi ASC"
        # if not mst and not ngay and not kieudieuchuyen:
        #     query = f"""
        #         SELECT * FROM HR.dbo.Lich_su_cong_tac WHERE Nha_may = '{current_user.macongty}' 
        #         ORDER BY Ngay_thuc_hien DESC, CAST(MST AS INT) ASC, Line_moi ASC
        #         """
        
        rows = cursor.execute(query)
        result = []
        for row in rows:
            result.append({
                "MST": row[1],
                "Họ tên": row[2],
                "Chuyền cũ": row[3],
                "Chuyền mới": row[5],
                "Vị trí cũ": row[4],
                "Vị trí mới": row[6],
                "Phân loại": row[7],
                "Ngày thực hiện": row[8],
                "Ghi chú": row[9]
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
        query = f"SELECT * FROM HR.dbo.Dang_ky_thong_tin WHERE Nha_may = '{current_user.macongty}'"
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
                "Họ tên": row[2],
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
        
        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        app.logger.info(e)
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

# def themnhanvienvaomita(mst,hoten):
#     conn = pyodbc.connect(mcc_db)
#     cursor = conn.cursor()
#     tenchamcong = xoadautrongten(hoten)
#     query = f"INSERT INTO MITACOSQL.Dbo.NHANVIEN (MaNhanVien,TenNhanVien,MaChamCong,TenChamCong) VALUES ('{mst}','{tenchamcong}','{mst}','{tenchamcong}')"
#     
#     cursor.execute(query)
#     conn.commit()
#     conn.close()

def themlichsutrangthai(mst,tungay,denngay,trangthai):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"INSERT INTO [HR].[dbo].[Lich_su_trang_thai_lam_viec] VALUES ('{mst}','{current_user.macongty}','{tungay}','{denngay}',N'{trangthai}')"
        
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
        return 
    
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

def laydanhsachdiemdanhbu(mst=None,hoten=None,chucvu=None,chuyen=None,bophan=None,loaidiemdanh=None,ngaydiemdanh=None,lido=None,trangthai=None):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"SELECT * FROM HR.dbo.Diem_danh_bu WHERE Nha_may = '{current_user.macongty}' "
        if mst:
            query += f"AND MST LIKE '%{mst}%'"
        if hoten:
            query += f"AND Ho_ten LIKE N'%{hoten}%'"
        if chucvu:
            query += f"AND Chuc_vu LIKE N'%{chucvu}%'"
        if chuyen:
            query += f"AND Line LIKE N'%{chuyen}%'"
        if bophan:
            query += f"AND Bo_phan LIKE N'%{bophan}%'"
        if loaidiemdanh:
            query += f"AND Loai_diem_danh LIKE N'%{loaidiemdanh}%'"
        if ngaydiemdanh:
            query += f"AND Ngay_diem_danh = '{ngaydiemdanh}'"    
        if lido:
            query += f"AND Ly_do LIKE N'%{lido}%'"   
        if trangthai:
            query += f"AND Trang_thai LIKE N'%{trangthai}%'"   
            
        query += " ORDER BY Ngay_diem_danh DESC, Bo_phan ASC, Line ASC, MST ASC"
        
        
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        app.logger.info(e)
        return []

def laydanhsachxinnghiphep(mst,hoten,chucvu,chuyen,bophan,ngaynghi,lydo,trangthai):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
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
        if not mst and not hoten and not chucvu and not chuyen and not bophan and not ngaynghi and not lydo and not trangthai:
            query = f"""
                        SELECT *
                        FROM HR.dbo.DS_Xin_nghi_phep
                        WHERE Nha_may = '{current_user.macongty}'
                        ORDER BY Ngay_nghi_phep DESC, MST ASC;
                    """
        
        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        app.logger.info(e)
        return []

def laydanhsachxinnghikhongluong(mst,hoten,chucvu,chuyen,bophan,ngay,lydo,trangthai):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
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
        app.logger.info(result[0])
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

def themdoicamoi(mst,cacu,camoi,ngaybatdau,ngayketthuc):
    try:
        if cacu != camoi and ngaybatdau < ngayketthuc:
            conn = pyodbc.connect(used_db)
            cursor = conn.cursor()
            if ngayketthuc:
                if type(ngaybatdau) == str:
                    ngayketthuccacu = datetime.strptime(ngaybatdau, '%Y-%m-%d') - timedelta(days=1)
                else:
                    ngayketthuccacu = ngaybatdau - timedelta(days=1)
                if type(ngayketthuc) == str:
                    ngayvecamacdinh = datetime.strptime(ngayketthuc, '%Y-%m-%d') + timedelta(days=1)
                else:
                    ngayvecamacdinh = ngayketthuc + timedelta(days=1)
                ngaymoc = datetime(2054,12,31)
                query = f"""
                            UPDATE HR.dbo.Dang_ky_ca_lam_viec
                            SET Den_ngay = '{ngayketthuccacu}'
                            WHERE MST='{mst}' AND Den_ngay = '{ngaymoc}' AND Factory = '{current_user.macongty}'

                            INSERT INTO HR.dbo.Dang_ky_ca_lam_viec
                            VALUES ('{mst}','{current_user.macongty}','{ngaybatdau}','{ngayketthuc}','{camoi}')

                            INSERT INTO HR.dbo.Dang_ky_ca_lam_viec
                            VALUES ('{mst}','{current_user.macongty}','{ngayvecamacdinh}','{ngaymoc}','{cacu}')
                        """
                
                cursor.execute(query)
                conn.commit()
                conn.close()
            else:
                if type(ngaybatdau) == str:
                    ngayketthuccacu = datetime.strptime(ngaybatdau, '%Y-%m-%d') - timedelta(days=1)
                else:
                    ngayketthuccacu = ngaybatdau - timedelta(days=1)
                ngaymoc = datetime(2054,12,31)
                query = f"""
                            UPDATE HR.dbo.Dang_ky_ca_lam_viec
                            SET Den_ngay = '{ngayketthuccacu}'
                            WHERE MST='{mst}' AND Den_ngay = '{ngaymoc}' AND Factory = '{current_user.macongty}'

                            INSERT INTO HR.dbo.Dang_ky_ca_lam_viec
                            VALUES ('{mst}','{current_user.macongty}','{ngaybatdau}','{ngaymoc}','{camoi}')
                        """
                
                cursor.execute(query)
                conn.commit()
                conn.close()
    except Exception as e:
        app.logger.info(e)     
        
def laycahientai(mst):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        ngayketthuc = datetime(2054,12,31)
        query = f"SELECT * FROM HR.dbo.Dang_ky_ca_lam_viec WHERE MST = '{mst}' AND Factory = '{current_user.macongty}' AND Den_ngay = '{ngayketthuc}'"
        
        row = cursor.execute(query).fetchone()
        if row:
            return row[-1]
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
    except Exception as e:
        app.logger.info(e)

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
        cursor.execute(query)
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.info(e)
        
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
            Dang_ky_ca_lam_viec.MST = Danh_sach_CBCNV.MST AND Dang_ky_ca_lam_viec.Factory = Danh_sach_CBCNV.Factory
        WHERE 
            Dang_ky_ca_lam_viec.Factory = '{current_user.macongty}' AND Danh_sach_CBCNV.Trang_thai_lam_viec=N'Đang làm việc'
        """
        if mst:
            query += f" AND Dang_ky_ca_lam_viec.MST LIKE '%{mst}%'"
        if chuyen:
            query += f" AND Danh_sach_CBCNV.Line LIKE '%{chuyen}%'"
        if phongban:
            query += f" AND Danh_sach_CBCNV.Department LIKE '%{phongban}%'"
        query += "ORDER BY Dang_ky_ca_lam_viec.Tu_ngay desc, Dang_ky_ca_lam_viec.Den_ngay desc, MST asc"

        rows = cursor.execute(query).fetchall()
        conn.close()
        return rows
    except Exception as e:
        app.logger.info(e)
        return []

def laydanhsachkpi(mst,macongty):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"select * from KPI_Name where MST='{mst}' and  Nha_may='{macongty}'"
        rows = cursor.execute(query).fetchall()
        result = []
        for row in rows:
            result.append({
                "Department": row[3],
                "Objective": row[4],
                "Possible measures": row[5],
                "Unit of measurement": row[6],
                "Measurement source": row[7],
                "C target": row[8],
                "B target": row[9],
                "A target": row[10],
            })
        return result
    except Exception as e:
        app.logger.info(e)
        return None  

def inputkpi(macongty,
                masothe,
                hoten,
                department,
                objectives,
                possibleMeasures,
                unitOfMeasurement,
                measurementSource,
                aTarget,
                bTarget,
                cTarget,
                month,
                year,
                actualResult):
    try:
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        query = f"insert into KPI_Data values ('{macongty}','{masothe}',N'{hoten}','{department}','{objectives}','{possibleMeasures}','{unitOfMeasurement}','{measurementSource}','{cTarget}','{bTarget}','{aTarget}','{actualResult}','{month}','{year}')"

        cursor.execute(query)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        app.logger.info(e)
        return False  

   
def roles_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if not current_user.is_authenticated or current_user.role not in roles:
                return redirect(url_for('unauthorized'))
            return f(*args, **kwargs)
        return decorated_function
    return decorator

##################################
#          MAIN ROUTES           #
##################################

if __name__ != '__main__':
    gunicorn_logger = logging.getLogger('gunicorn.error')
    app.logger.handlers = gunicorn_logger.handlers
    app.logger.setLevel(gunicorn_logger.level)

if __name__ == '__main__':
    handler = RotatingFileHandler('app.log', maxBytes=10000, backupCount=1, encoding='utf-8')
    handler.setLevel(logging.INFO)
    formatter = logging.Formatter(
        '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
    )
    handler.setFormatter(formatter)
    app.logger.addHandler(handler)
    app.logger.setLevel(logging.INFO)


@app.route('/unauthorized')
def unauthorized():
    return "Bạn không có quyền truy cập vui lòng chọn mục khác", 401

@app.errorhandler(404)
def page_not_found(e):
    return render_template('blank.html'), 404

@app.route('/admin', methods=["GET"])
@login_required
@roles_required('sa')
def admin_template():
    page = request.args.get('page', 1, type=int)
    per_page = 10
    mst = request.args.get('mst', None, type=str)
    if mst:
        users_paginated = Users.query.filter(Users.masothe.like(f"%{mst}%")).paginate(page=page, per_page=per_page, error_out=False)
    else:
        hoten = request.args.get('hoten', None, type=str)
        if hoten:
            search_pattern = f"%{hoten}%"
            users_paginated = Users.query.filter(Users.hoten.like(search_pattern)).paginate(page=page, per_page=per_page, error_out=False)
        else:
            users_paginated = Users.query.paginate(page=page, per_page=per_page, error_out=False)
    cacrole= ['sa','user','hr','gd','luong','tnc','td']
    return render_template('admin.html', users=users_paginated,cacrole=cacrole)

@app.route('/register', methods=["POST"])
def register():
    if request.method == "POST":
        if request.args.get("macongty") == "NT1":
            tencty = "Công ty cổ phần sản xuất Nam Thuận"
        elif request.args.get("macongty") == "NT2":
            tencty = "Công ty cổ phần Nam Thuận Nghệ An"
        elif request.args.get("macongty") == "NT0":    
            tencty = "Công ty cổ phần tập đoàn Nam Thuận"
        user = Users(masothe=request.args.get("masothe"),
                     hoten=request.args.get("hoten"),
                     phongban=request.args.get("phongban"),
                     macongty=request.args.get("macongty"),
                     matkhau=request.args.get("matkhau"),
                     tencongty = tencty
                     )
        try:
            db.session.add(user)
            db.session.commit()
        except Exception as e:
            return str(e)
        return "OK"
 
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = Users.query.filter_by(
            masothe=request.form.get("masothe"),
            macongty=request.form.get("congty")).first()
        if not user:
            return redirect(url_for("login"))
        if user.matkhau == request.form.get("matkhau"):
            login_user(user)
            app.logger.info(f"User {user.masothe} {user.macongty} logged in")
            return redirect(url_for("home"))
        else:
            return redirect(url_for("login"))
    return render_template("login.html")

@app.route("/logout", methods=["POST"])
def logout():
    logout_user()
    return redirect(url_for("login"))

@app.route("/doimatkhau", methods=['POST'])
def doimatkhau():
    macongty = request.form.get("macongty")
    masothe = request.form.get("masothe")
    matkhaumoi = request.form.get("matkhaumoi")
    try:
        if doimatkhautaikhoan(macongty,masothe,matkhaumoi):
            flash("Đổi mật khẩu thành công")
    except Exception as e:
        flash(f"Đổi mật khẩu không thành công: {e}")
    return redirect(url_for("home"))


@app.route("/home", methods=['GET','POST'])
@login_required
def index():
    return redirect("/")

@app.route("/", methods=['GET','POST'])
@login_required
def home():
    if request.method == "GET":
        mst = request.args.get("Mã số thẻ")
        hoten = request.args.get("Họ tên")
        sdt = request.args.get("Số điện thoại")
        cccd = request.args.get("Căn cước công dân")
        gioitinh = request.args.get("Giới tính")
        vaotungay = request.args.get("Vào từ ngày")
        vaodenngay = request.args.get("Vào đến ngày")
        nghitungay = request.args.get("Nghỉ từ ngày")
        nghidenngay = request.args.get("Nghỉ đến ngày")
        phongban = request.args.get("Phòng ban")
        chucvu = request.args.get("Chức danh")
        trangthai = request.args.get("Trạng thái")
        hccategory = request.args.get("Headcount Category")
        ghichu = request.args.get("Ghi chú")
        users = laydanhsachuser(mst, hoten, sdt, cccd, gioitinh, vaotungay, vaodenngay, nghitungay, nghidenngay, phongban, trangthai, hccategory, chucvu, ghichu)   
        count = len(users)
        cacphongban = laycacphongban()
        cacto = laycacto()
        cactrangthai = laycactrangthai()
        cachccategory = laycachccategory()
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(users)
        start = (page - 1) * per_page
        end = start + per_page
        paginated_users = users[start:end]
        
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        flash(f"Chào mừng {current_user.hoten} quay trở lại!")
        return render_template("home.html", users=paginated_users, 
                            cacphongban=cacphongban, cacto=cacto,
                            page="Trang chủ", pagination=pagination,
                            cactrangthai=cactrangthai,count=count,
                            cachccategory=cachccategory)
    else:
        mst = request.form.get("Mã số thẻ")
        hoten = request.form.get("Họ tên")
        sdt = request.form.get("Số điện thoại")
        cccd = request.form.get("Căn cước công dân")
        gioitinh = request.form.get("Giới tính")
        vaotungay = request.form.get("Vào từ ngày")
        vaodenngay = request.form.get("Vào đến ngày")
        nghitungay = request.form.get("Nghỉ từ ngày")
        nghidenngay = request.form.get("Nghỉ đến ngày")
        phongban = request.form.get("Phòng ban")
        chucvu = request.form.get("Chức danh")
        trangthai = request.form.get("Trạng thái")
        hccategory = request.form.get("Headcount Category")
        ghichu = request.form.get("Ghi chú")
            
        users = laydanhsachuser(mst, hoten, sdt, cccd, gioitinh, vaotungay, vaodenngay, nghitungay, nghidenngay, phongban, trangthai, hccategory, chucvu,ghichu)      
        df = pd.DataFrame(users)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT, f"danhsachnhanvien_{thoigian}.xlsx"), index=False)
        
        return send_file(os.path.join(FOLDER_XUAT, f"danhsachnhanvien_{thoigian}.xlsx"), as_attachment=True)

@app.route("/muc2_1", methods=["GET","POST"])
@login_required
@roles_required('hr','tnc','sa','gd','td')
def danhsachdangkytuyendung():
    if request.method == "GET":
        sdt = request.args.get("sdt")
        cccd = request.args.get("cccd")
        ngaygui = request.args.get("ngaygui")
        rows = laydanhsachdangkytuyendung(sdt,cccd,ngaygui)
        for row in rows:
            row['Ngày hẹn đi làm'] = datetime.strptime(row['Ngày hẹn đi làm'], '%d/%m/%Y').strftime('%Y-%m-%d') if row['Ngày hẹn đi làm'] else None
        count=len(rows)
        count = len(rows)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(rows)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = rows[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')

        return render_template("2_1.html", 
                            page="2.1 Danh sách ứng viên",
                            danhsach=paginated_rows, 
                            pagination=pagination,
                            count=count)
        
    if request.method == "POST":
        id = request.form.get("id")
        sdt = request.form.get("sdt")
        vitrituyendung = request.form.get("vitrituyendung")
        hocvan = request.form.get("hocvan")
        diachi = request.form.get("diachi")
        ngayhendilam = request.form.get("ngayhendilam")
        hieusuat = request.form.get("hieusuat")
        loaimay = request.form.get("loaimay")
        cccd = request.form.get("cccd")
        dantoc = request.form.get("dantoc")
        connho = request.form.get("connho")
        tencon1 = request.form.get("tenconnho1")
        ngaysinhcon1 = request.form.get("ngaysinhcon1")
        tencon2 = request.form.get("tenconnho2")
        ngaysinhcon2 = request.form.get("ngaysinhcon2")
        tencon3 = request.form.get("tenconnho3")
        ngaysinhcon3 = request.form.get("ngaysinhcon3")
        tencon4 = request.form.get("tenconnho4")
        ngaysinhcon4 = request.form.get("ngaysinhcon4")
        tencon5 = request.form.get("tenconnho5")
        ngaysinhcon5 = request.form.get("ngaysinhcon5")
        nguoithan = request.form.get("nguoithan")
        sdtnguoithan = request.form.get("sdtnguoithan")
        ngayhendilam = request.form.get("ngayhendilam")
        luuhoso = request.form.get("luuhoso")
        ghichu = request.form.get("ghichu")
        capnhatthongtinungvien(id,
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
                               )
        return redirect(f"muc2_1?sdt={sdt}")

@app.route("/muc2_2_1", methods=["GET","POST"])
@login_required
@roles_required('tbp','gd','sa')
def dangkytuyendung():
    if request.method == "GET":
        maso = current_user.macongty[-1]
        danhsach = laydanhsachyeucautuyendung(maso)
        
        return render_template("2_2_1.html", page="2.2.1 Thêm yêu cầu tuyển dụng",danhsach=danhsach)
    elif request.method == "POST":
        bophan = request.form.get("bophan")
        vitri = request.form.get("vitri")
        soluong = request.form.get("soluong")
        mota = request.form.get("mota")
        thoigiandukien = request.form.get("thoigiandukien")
        phanloai = request.form.get("phanloai")
        mucluongtu = request.form.get("mucluongtu")
        mucluongden = request.form.get("mucluongden")
        mucluong = f"{mucluongtu} - {mucluongden} triệu VNĐ"
        themyeucautuyendungmoi(bophan,vitri,soluong,mota,thoigiandukien,phanloai,mucluong)
        return redirect("muc2_2_1")
    
@app.route("/muc2_2_2", methods=["GET","POST"])
@login_required
@roles_required('tbp','gd','sa','nhansu')
def pheduyettuyendung():   
    if request.method == "GET":
        maso = current_user.macongty[-1]
        danhsach = laydanhsachyeucautuyendung(maso)
        return render_template("2_2_2.html", page="2.2.2 Phê duyệt yêu cầu tuyển dụng",danhsach=danhsach)
    
    elif request.method == "POST":
        bophan = request.form.get("bophan")
        vitri = request.form.get("vitri")
        soluong = request.form.get("soluong")
        mota = request.form.get("mota")
        thoigian = request.form.get("thoigian")
        phanloai = request.form.get("phanloai")
        trangthaiyeucau = request.form.get("trangthaiyeucau") if request.form.get("trangthaiyeucau") else None
        trangthaithuchien = request.form.get("trangthaithuchien") if request.form.get("trangthaithuchien") else None
        ghichu = request.form.get("ghichu") if request.form.get("ghichu") else None
        capnhattrangthaiyeucautuyendung(bophan,vitri,soluong,mota,thoigian,phanloai,trangthaiyeucau,trangthaithuchien,ghichu)
        return redirect("/muc2_2_2")
    
@app.route("/muc3_1", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def nhapthongtinlaodongmoi():
    
    if request.method == "GET":
        data= request.args.get("data")
        masothe = checkformatmst(int(laymasothemoi())+1)
        cacvitri= laycacvitri()
        cacto = laycacto()
        cacca = laycacca()
        data=request.args.get("scan-qrcode")
        datenow = datetime.now()
        macongty = current_user.macongty 
        return render_template("3_1.html", 
                                page="3.1 Nhập thông tin lao động mới",
                                qrcccd=data,
                                masothe=masothe,
                                ngaybatdau=datenow,
                                cacvitri=cacvitri,
                                cacto=cacto,
                                cacca=cacca,
                                macongty=macongty)
    elif request.method == "POST":

        anh = f"N'{request.form.get("anh")}'"
        masothe = f"'{request.form.get("masothe")}'"
        thechamcong = f"'{int(request.form.get("masothe"))}'"
        hoten = f"N'{request.form.get("hoten")}'"
        ngaysinh = f"'{request.form.get("ngaysinh")}'" if request.form.get("ngaysinh") else "NULL"
        gioitinh = f"N'{request.form.get("gioitinh")}'"
        cmt = f"'{request.form.get("cmt")}'"
        cccd = f"'{request.form.get("cccd")}'"
        ngaycapcccd = f"'{request.form.get("ngaycap")}'" if request.form.get("ngaycap") else "NULL"
        thuongtru = f"N'{request.form.get("thuongtru")}'" if request.form.get("thuongtru") else "NULL"
        noisinh = f"N'{request.form.get("noisinh")}'" if request.form.get("noisinh") else "NULL"
        tamtru = f"N'{request.form.get("tamtru")}'" if request.form.get("tamtru") else "NULL"
        quoctich = f"N'{request.form.get("quoctich")}'" if request.form.get("quoctich") else "NULL"
        dantoc = f"N'{request.form.get("dantoc")}'" if request.form.get("dantoc") else "NULL"
        tongiao = f"N'{request.form.get("tongiao")}'" if request.form.get("tongiao") else "NULL"
        hocvan = f"N'{request.form.get("hocvan")}'" if request.form.get("hocvan") else "NULL"
        thonxom = f"N'{request.form.get("thonxom")}'" if request.form.get("thonxom") else "NULL"
        phuongxa = f"N'{request.form.get("phuongxa")}'" if request.form.get("phuongxa") else "NULL"
        quanhuyen = f"N'{request.form.get("quanhuyen")}'" if request.form.get("quanhuyen") else "NULL"
        tinhthanhpho = f"N'{request.form.get("tinhthanhpho")}'" if request.form.get("tinhthanhpho") else "NULL"
        nganhang = f"N'{request.form.get("nganhang")}'" if request.form.get("nganhang") else "NULL"
        sotaikhoan = f"'{request.form.get("sotaikhoan")}'" if request.form.get("sotaikhoan") else "NULL"
        dienthoai = f"'{request.form.get("dienthoai")}'" if request.form.get("dienthoai") else "NULL"
        sobhxh = f"'{request.form.get("sobhxh")}'" if request.form.get("sobhxh") else 'NULL'
        masothue = f"'{request.form.get("masothue")}'" if request.form.get("masothue") else 'NULL'
        connho = f"N'{request.form.get("connho")}'" if request.form.get("connho") else 'NULL'
        tencon1 = f"N'{request.form.get("tenconmnho1")}'" if request.form.get("tenconmnho1") else 'NULL'
        ngaysinhcon1 = f"'{request.form.get("ngaysinhcon1")}'" if request.form.get("ngaysinhcon1") else 'NULL'
        tencon2 = f"N'{request.form.get("tenconmnho2")}'" if request.form.get("tenconmnho2") else 'NULL'
        ngaysinhcon2 = f"'{request.form.get("ngaysinhcon2")}'" if request.form.get("ngaysinhcon2") else 'NULL'
        tencon3 = f"N'{request.form.get("tenconmnho3")}'" if request.form.get("tenconmnho3") else 'NULL'
        ngaysinhcon3 = f"'{request.form.get("ngaysinhcon3")}'" if request.form.get("ngaysinhcon3") else 'NULL'
        tencon4 = f"N'{request.form.get("tenconmnho4")}'" if request.form.get("tenconmnho4") else 'NULL'
        ngaysinhcon4 = f"'{request.form.get("ngaysinhcon4")}'" if request.form.get("ngaysinhcon4") else 'NULL'
        tencon5 = f"N'{request.form.get("tenconmnho5")}'" if request.form.get("tenconmnho5") else 'NULL'
        ngaysinhcon5 = f"'{request.form.get("ngaysinhcon5")}'" if request.form.get("ngaysinhcon5") else 'NULL'
        jobdetailvn = f"N'{request.form.get("vitri")}'"
        line = f"'{request.form.get("line")}'"
        calamviec = request.form.get("calamviec")
        factory = f"'{current_user.macongty}'"
        hccategory = f"N'{request.form.get("hccategory")}'"
        gradecode = f"N'{request.form.get("gradecode")}'"
        department = f"N'{request.form.get("phongban")}'"
        chucvu = f"N'{request.form.get("chucvu")}'"
        employeetype = f"N'{request.form.get("loailaodong")}'"
        sectioncode = f"N'{request.form.get("mabophan")}'"
        sectiondescription = f"N'{request.form.get("bophan")}'"
        jobdetailen = f"N'{request.form.get("vitrien")}'"
        positioncode = f"N'{request.form.get("mavitri")}'"
        positioncodedescription = f"N'{request.form.get("tenvitri")}'"
        nguoithan = f"N'{request.form.get("nguoithan")}'" if request.form.get("nguoithan") else 'NULL'
        sdtnguoithan = f"N'{request.form.get("sodienthoainguoithan")}'" if request.form.get("sodienthoainguoithan") else 'NULL'
        luongcoban = f"'{request.form.get("luongcoban").replace(',','')}'" if request.form.get("luongcoban") else 'NULL'
        tongphucap = f"'{request.form.get("tongphucap").replace(',','')}'" if request.form.get("tongphucap") else 'NULL'
        kieuhopdong = request.form.get("kieuhopdong")
        if kieuhopdong == "HĐ thử việc":
            kieuhopdong = "N'HĐ thử việc'"
            ngaybatdauthuviec = f"'{request.form.get("ngayBatDau")}'" if request.form.get("ngayBatDau") else 'NULL'
            ngayvao = f"'{request.form.get("ngayBatDau")}'" if request.form.get("ngayBatDau") else 'NULL'
            ngayketthuc = f"'{request.form.get("ngayKetThuc")}'" if request.form.get("ngayKetThuc") else 'NULL'
            ngayketthucthuviec = f"'{request.form.get("ngayKetThuc")}'" if request.form.get("ngayKetThuc") else 'NULL'
            ngaybatdauhdcthl1 = "NULL"
            ngayketthuchdcthl1 = "NULL"
            ngaybatdauhdcthl2 = "NULL"
            ngayketthuchdcthl2 = "NULL"
            ngaybatdauhdvth = "NULL"
        elif kieuhopdong == "HĐ có thời hạn lần 1":
            kieuhopdong = "N'HĐ có thời hạn lần 1'"
            ngayvao = f"'{request.form.get("ngayBatDau")}'" if request.form.get("ngayBatDau") else 'NULL'
            ngayketthuc = f"'{request.form.get("ngayKetThuc")}'" if request.form.get("ngayKetThuc") else 'NULL'
            ngaybatdauhdcthl1 = f"'{request.form.get("ngayBatDau")}'" if request.form.get("ngayBatDau") else 'NULL'
            ngayketthuchdcthl1 = f"'{request.form.get("ngayKetThuc")}'" if request.form.get("ngayKetThuc") else 'NULL'
            ngaybatdauthuviec = "NULL"
            ngayketthucthuviec = "NULL"
            ngaybatdauhdcthl2 = "NULL"
            ngayketthuchdcthl2 = "NULL"
            ngaybatdauhdvth = "NULL"
        elif kieuhopdong == "HĐ có thời hạn lần 2":
            kieuhopdong = "N'HĐ có thời hạn lần 2'"
            ngayvao = f"'{request.form.get("ngayBatDau")}'" if request.form.get("ngayBatDau") else 'NULL'
            ngayketthuc = f"'{request.form.get("ngayKetThuc")}'" if request.form.get("ngayKetThuc") else 'NULL'
            ngaybatdauhdcthl2 = f"'{request.form.get("ngayBatDau")}'" if request.form.get("ngayBatDau") else 'NULL'
            ngayketthuchdcthl2 = f"'{request.form.get("ngayKetThuc")}'" if request.form.get("ngayKetThuc") else 'NULL'
            ngaybatdauthuviec = "NULL"
            ngayketthucthuviec = "NULL"
            ngaybatdauhdcthl1 = "NULL"
            ngayketthuchdcthl1 = "NULL"
            ngaybatdauhdvth = "NULL"
        elif kieuhopdong == "HĐ vô thời hạn":
            kieuhopdong = "N'HĐ vô thời hạn'"
            ngayvao = f"'{request.form.get("ngayBatDau")}'" if request.form.get("ngayBatDau") else 'NULL'
            ngayketthuc = 'NULL'
            ngaybatdauhdvth = f"'{request.form.get("ngayBatDau")}'" if request.form.get("ngayBatDau") else 'NULL'
            ngaybatdauthuviec = "NULL"
            ngayketthucthuviec = "NULL"
            ngaybatdauhdcthl1 = "NULL"
            ngayketthuchdcthl1 = "NULL"
            ngaybatdauhdcthl2 = "NULL"
            ngayketthuchdcthl2 = "NULL"
        else:
            ngayvao = f"'{request.form.get("ngayBatDau")}'" if request.form.get("ngayBatDau") else 'NULL'
            kieuhopdong = "NULL"
            ngayketthuc = 'NULL'
            ngaybatdauthuviec = "NULL"
            ngayketthucthuviec = "NULL"
            ngaybatdauhdcthl1 = "NULL"
            ngayketthuchdcthl1 = "NULL"
            ngaybatdauhdcthl2 = "NULL"
            ngayketthuchdcthl2 = "NULL"
            ngaybatdauhdvth = "NULL"
        nhanvienmoi = f"({masothe},{thechamcong},{hoten},{dienthoai},{ngaysinh},{gioitinh},{cccd},{ngaycapcccd},N'Cục cảnh sát',{cmt},{thuongtru},{thonxom},{phuongxa},{quanhuyen},{tinhthanhpho},{dantoc},{quoctich},{tongiao},{hocvan},{noisinh},{tamtru},{sobhxh},{masothue},{nganhang},{sotaikhoan},{connho},{tencon1},{ngaysinhcon1},{tencon2},{ngaysinhcon2},{tencon3},{ngaysinhcon3},{tencon4},{ngaysinhcon4},{tencon5},{ngaysinhcon5},{anh},{nguoithan}, {sdtnguoithan},{kieuhopdong},{ngayvao},{ngayketthuc},{jobdetailvn},{hccategory},{gradecode},{factory},{department},{chucvu},{sectioncode},{sectiondescription},{line},{employeetype},{jobdetailen},{positioncode},{positioncodedescription},{luongcoban},N'Không',{tongphucap},{ngayvao},NULL,N'Đang làm việc',{ngayvao},'1',{ngaybatdauthuviec},{ngayketthucthuviec},{ngaybatdauhdcthl1},{ngayketthuchdcthl1},{ngaybatdauhdcthl2},{ngayketthuchdcthl2},{ngaybatdauhdvth},'N', '')"             
        if themnhanvienmoi(nhanvienmoi):
            themdoicamoi(request.form.get("masothe"),"A1-01",calamviec,ngayvao.replace("'",""),datetime(2054,12,31))
            themlichsutrangthai(request.form.get("masothe"),request.form.get("ngayBatDau"),datetime(2054,12,31),'Đang làm việc')
            return redirect("/muc3_1")
        else:
            masothe = int(laymasothemoi())+1
            cacvitri= laycacvitri()
            cacto = laycacto()
            cacca = laycacca()
            return redirect("/muc3_1")
        
@app.route("/muc3_2", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def thaydoithongtinlaodong():
    
    if request.method == "GET":
        return render_template("3_2.html", page="3.2 Thay đổi thông tin người lao động")
    else:
        trangthailamviec = request.form.get("trangthai")
        thechamcong = request.form.get("thechamcong")
        cccd = request.form.get("cccd")
        ngaycapcccd = request.form.get("ngaycapcccd")
        hoten = request.form.get("hoten")
        ngaysinh = request.form.get("ngaysinh")
        gioitinh = request.form.get("gioitinh")
        cmt = request.form.get("cmt")
        quoctich = request.form.get("quoctich")
        dienthoai = request.form.get("dienthoai")
        thonxom = request.form.get("thonxom")
        phuongxa = request.form.get("phuongxa")
        quanhuyen = request.form.get("quanhuyen")
        tinhthanhpho = request.form.get("tinhthanhpho")
        dantoc = request.form.get("dantoc")
        tongiao = request.form.get("tongiao")
        hocvan = request.form.get("hocvan")
        masothue = request.form.get("masothue")
        nganhang = request.form.get("nganhang")
        sotaikhoan = request.form.get("sotaikhoan")
        connho = request.form.get("connho")
        mst = request.form.get("mst")
        tenconnho1 = request.form.get("tenconnho1")
        tenconnho2 = request.form.get("tenconnho2")
        tenconnho3 = request.form.get("tenconnho3")
        tenconnho4 = request.form.get("tenconnho4")
        tenconnho5 = request.form.get("tenconnho5")
        ngaysinhcon1 = request.form.get("ngaysinhcon1")
        ngaysinhcon2 = request.form.get("ngaysinhcon2")
        ngaysinhcon3 = request.form.get("ngaysinhcon3")
        ngaysinhcon4 = request.form.get("ngaysinhcon4")
        ngaysinhcon5 = request.form.get("ngaysinhcon5")
        jobtitlevn = request.form.get("jobtitlevn") 
        jobtitleen = request.form.get("jobtitleen") 
        positioncode = request.form.get("positioncode")
        positioncodedescription = request.form.get("positioncodedescription")
        chucvu =  request.form.get("chucvu")
        line = request.form.get("line")
        department = request.form.get("department")
        sectioncode = request.form.get("sectioncode")
        sectiondescription = request.form.get("sectiondescription")
        hccategory = request.form.get("hccategory")
        employeetype = request.form.get("employeetype")
        gradecode = request.form.get("gradecode")
        factory = request.form.get("factory")
        kieuhopdong = request.form.get("kieuhopdong")
        ngaybatdau = request.form.get("ngaybatdau")
        ngayketthuc = request.form.get("ngayketthuc")
        mucluong = request.form.get("mucluong").replace(',','')
        phucap = request.form.get("phucap").replace(',','')
        tienphucap = request.form.get("tienphucap")
        
        query = f"UPDATE HR.dbo.Danh_sach_CBCNV SET "
        if thechamcong:
            query += f"The_cham_cong = '{thechamcong}',"
        else:
            query += f"The_cham_cong = NULL,"
            
        if cccd: 
            query += f"CCCD = '{cccd}',"
        else:
            query += f"CCCD = NULL,"
            
        if ngaycapcccd: 
            query += f"Ngay_cap = '{ngaycapcccd}',"
        else:
            query += f"Ngay_cap = NULL,"
            
        if hoten: 
            query += f"Ho_ten = N'{hoten}',"
        else:
            query += f"Ho_ten = NULL,"
            
        if ngaysinh: 
            query += f"Ngay_sinh = '{ngaysinh}',"
        else:
            query += f"Ngay_sinh = NULL,"
            
        if gioitinh: 
            query += f"Gioi_tinh = N'{gioitinh}',"
        else:
            query += f"Gioi_tinh = NULL,"
                
        if cmt: 
            query += f"CMT = N'{cmt}',"
        else:
            query += f"CMT = NULL,"   
            
        if quoctich: 
            query += f"Quoc_tich = N'{quoctich}',"
        else:
            query += f"Quoc_tich = NULL," 
            
        if dienthoai: 
            query += f"Sdt = N'{dienthoai}',"
        else:
            query += f"Sdt = NULL,"
            
        if thonxom: 
            query += f"Thon_xom = N'{thonxom}',"
        else:
            query += f"Thon_xom = NULL,"
            
        if phuongxa: 
            query += f"Phuong_xa = N'{phuongxa}',"
            
        if quanhuyen: 
            query += f"Quan_huyen = N'{quanhuyen}',"
        else:
            query += f"Quan_huyen = NULL,"
            
        if tinhthanhpho: 
            query += f"Tinh_TP = N'{tinhthanhpho}',"
        else:
            query += f"Tinh_TP = NULL,"
            
        if dantoc: 
            query += f"Dan_toc = N'{dantoc}',"
        else:
            query += f"Dan_toc = NULL,"
                
        if tongiao: 
            query += f"Ton_giao = N'{tongiao}',"
        else:
            query += f"Ton_giao = NULL,"
                
        if hocvan: 
            query += f"Trinh_do = N'{hocvan}',"
        else:
            query += f"Trinh_do = NULL,"
            
        if masothue: 
            query += f"Ma_so_thue = N'{masothue}',"
        else:
            query += f"Ma_so_thue = NULL,"   
            
        if nganhang: 
            query += f"Ngan_hang = N'{nganhang}',"
        else:
            query += f"Ngan_hang = NULL,"   
            
        if sotaikhoan: 
            query += f"So_tai_khoan = N'{sotaikhoan}',"
        else:
            query += f"So_tai_khoan = NULL,"
            
        if connho: 
            query += f"Con_nho = N'{connho}',"
        else:
            query += f"Con_nho = NULL,"

        if tenconnho1: 
            query += f"Ten_con_nho_1 = N'{tenconnho1}',"
        else:
            query += f"Ten_con_nho_1 = NULL,"
        
        if tenconnho2: 
            query += f"Ten_con_nho_2 = N'{tenconnho2}',"
        else:
            query += f"Ten_con_nho_2 = NULL,"
            
        if tenconnho3: 
            query += f"Ten_con_nho_3 = N'{tenconnho3}',"
        else:
            query += f"Ten_con_nho_3 = NULL,"
            
        if tenconnho4: 
            query += f"Ten_con_nho_4 = N'{tenconnho4}',"
        else:
            query += f"Ten_con_nho_4 = NULL,"
        
        if tenconnho5: 
            query += f"Ten_con_nho_5 = N'{tenconnho5}',"
        else:
            query += f"Ten_con_nho_5 = NULL,"
        
        if ngaysinhcon1: 
            query += f"Ngay_sinh_con_nho_1 = '{ngaysinhcon1}',"
        else:
            query += f"Ngay_sinh_con_nho_1 = NULL,"
        
        if ngaysinhcon2: 
            query += f"Ngay_sinh_con_nho_2 = '{ngaysinhcon2}',"
        else:
            query += f"Ngay_sinh_con_nho_2 = NULL,"
            
        if ngaysinhcon3: 
            query += f"Ngay_sinh_con_nho_3 = '{ngaysinhcon3}',"
        else:
            query += f"Ngay_sinh_con_nho_3 = NULL,"
            
        if ngaysinhcon4: 
            query += f"Ngay_sinh_con_nho_4 = '{ngaysinhcon4}',"
        else:
            query += f"Ngay_sinh_con_nho_4 = NULL,"
            
        if ngaysinhcon5: 
            query += f"Ngay_sinh_con_nho_5 = '{ngaysinhcon5}',"
        else:
            query += f"Ngay_sinh_con_nho_5 = NULL,"
        
        if jobtitlevn: 
            query += f"Job_title_VN = N'{jobtitlevn}',"
        else:
            query += f"Job_title_VN = NULL,"
            
        if jobtitleen: 
            query += f"Job_title_EN = '{jobtitleen}',"
        else:
            query += f"Job_title_EN = NULL,"
            
        if positioncode: 
            query += f"Position_code = '{positioncode}',"
        else:
            query += f"Position_code = NULL,"
        
        if positioncodedescription: 
            query += f"Position_code_description = '{positioncodedescription}',"
        else:
            query += f"Position_code_description = NULL,"  
        
        if chucvu: 
            query += f"Chuc_vu = '{chucvu}',"
        else:
            query += f"Chuc_vu = NULL," 
            
        if line: 
            query += f"Line = '{line}',"
        else:
            query += f"Line = NULL,"     
        
        if department: 
            query += f"Department = '{department}',"
        else:
            query += f"Department = NULL,"  
            
        if sectioncode: 
            query += f"Section_code = '{sectioncode}',"
        else:
            query += f"Section_code = NULL,"     
        
        if sectiondescription: 
            query += f"Section_description = '{sectiondescription}',"
        else:
            query += f"Section_description = NULL," 
            
        if hccategory: 
            query += f"Headcount_category = '{hccategory}',"
        else:
            query += f"Headcount_category = NULL," 
            
        if employeetype: 
            query += f"Emp_type = '{employeetype}',"
        else:
            query += f"Emp_type = NULL," 
            
        if gradecode: 
            query += f"Grade_code = '{gradecode}',"
        else:
            query += f"Grade_code = NULL," 
            
        if factory: 
            query += f"Factory = '{factory}',"
        else:
            query += f"Factory = NULL,"
            
        if kieuhopdong:
            query += f"Loai_hop_dong = N'{kieuhopdong}',"
        else:
            query += f"Loai_hop_dong = NULL,"
            
        if ngaybatdau:
            query += f"Ngay_ky_HD = '{ngaybatdau}',"
        else:
            query += f"Ngay_ky_HD = NULL,"
          
        if ngayketthuc:
            query += f"Ngay_het_han_HD = '{ngayketthuc}',"
        else:
            query += f"Ngay_het_han_HD = NULL,"
            
        if mucluong:
            query += f"Luong_co_ban = '{mucluong}',"
        else:
            query += f"Luong_co_ban = NULL,"
            
        if phucap:
            query += f"Phu_cap = N'{phucap}',"
        else:
            query += f"Phu_cap = NULL,"
            
        if tienphucap:
            query += f"Tong_phu_cap = '{tienphucap}',"
        else:
            query += f"Tong_phu_cap = NULL,"
        if trangthailamviec:
            query += f"Trang_thai_lam_viec = N'{trangthailamviec}',"
        else:
            query += f"trangthailamviec = NULL,"
        query = query[:-1] + f" WHERE MST = '{mst}' AND Factory='{current_user.macongty}'"
        conn = pyodbc.connect(used_db)
        cursor = conn.cursor()
        
        cursor.execute(query)
        conn.commit()
        return redirect("/muc3_2")
    
@app.route("/muc3_3", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def inhopdonglaodong():
    if request.method == "GET":
        return render_template("3_3.html", page="3.3 In hợp đồng lao động")
    elif request.method == "POST":
        bophan = request.form.get("bophan")
        kieuhopdong = request.form.get("kieuhopdong")
        mst = request.form.get("mst")
        ngaylamhopdong = request.form.get("ngayBatDau")[-2:]
        thanglamhopdong = request.form.get("ngayBatDau")[5:7]
        namlamhopdong = request.form.get("ngayBatDau")[:4]
        ngayketthuchopdong = request.form.get("ngayKetThuc")[-2:]
        thangketthuchopdong = request.form.get("ngayKetThuc")[5:7]
        namketthuchopdong = request.form.get("ngayKetThuc")[:4]
        tennhanvien = request.form.get("hoten")
        ngaysinh = datetime.strptime(request.form.get("ngaysinh"), "%Y-%m-%d").strftime("%d/%m/%Y")
        gioitinh = request.form.get("gioitinh")
        thuongtru = request.form.get("thuongtru")
        tamtru = request.form.get("tamtru")
        cccd = request.form.get("cccd")
        ngaycapcccd = datetime.strptime(request.form.get("ngaycapcccd"), "%Y-%m-%d").strftime("%d/%m/%Y")
        mucluong = request.form.get("luongcoban").replace(',','')
        chucvu = request.form.get("chucvu")
        capbac= request.form.get("capbac")
        phucap = request.form.get("phucap")
        tongphucap = request.form.get("tienphucap")
        songaythuviec = request.form.get("soNgay")
        try:
            file = inhopdongtheomau(kieuhopdong,
                                            mst,
                                            ngaylamhopdong,
                                            thanglamhopdong,
                                            namlamhopdong,
                                            ngayketthuchopdong,
                                            thangketthuchopdong,
                                            namketthuchopdong,
                                            tennhanvien,
                                            ngaysinh,
                                            gioitinh,
                                            thuongtru,
                                            tamtru,
                                            cccd,
                                            ngaycapcccd,
                                            mucluong,
                                            phucap,
                                            tongphucap,
                                            chucvu,
                                            bophan,
                                            capbac,
                                            songaythuviec)
            if file:
                return send_file(file, as_attachment=True, download_name="hopdonglaodong.xlsx")
            else:
                app.logger.info("NO FILE")
                return redirect("/muc3_3")
        except Exception as e:
            app.logger.info(e)
            return redirect("/muc3_3")   

@app.route("/muc3_4", methods=["GET","POST"])
@login_required
def danhsachsaphethanhopdong():
    if request.method == "GET":
        danhsach = laydanhsachsaphethanhopdong()
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("/3_4.html",
                               danhsach=paginated_rows,
                               pagination=pagination,
                               total = total)
    if request.method == "POST":
        danhsach = laydanhsachsaphethanhopdong()
        result = []
        for user in danhsach:
            result.append({
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
            })
        df = pd.DataFrame(result)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT, f"saphethan_{thoigian}.xlsx"), index=False)
        
        return send_file(os.path.join(FOLDER_XUAT, f"saphethan_{thoigian}.xlsx"), as_attachment=True)

@app.route("/muc5_1_1", methods=["GET","POST"])
@login_required
@roles_required('sa','gd','tbp')
def nhapkpi():
    danhsachdong = laydanhsachkpi(current_user.masothe,current_user.macongty)
    month = datetime.now().month-1
    year = datetime.now().year
    return render_template("5_1_1.html",page="Input Performance",danhsachdong=danhsachdong,month = month, year= year)

@app.route("/muc5_1_2", methods=["GET","POST"])
@login_required
@roles_required('sa','gd','tbp')
def danhgiakpi():
    return render_template("5_1_2.html",page="Dashboard Performance")
    
@app.route("/muc6_1", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def dieuchuyen():
    
    if request.method == "POST":
        mst = request.form["mst"]
        loaidieuchuyen = request.form["loaidieuchuyen"]
        ngaydieuchuyen = request.form.get("ngaydieuchuyen")
        ghichu = request.form.get("ghichu")
        
        vitricu = request.form.get("vitricu")
        vitrimoi = request.form.get("vitrimoi")
        
        vitriencu = request.form.get("vitriencu")
        vitrienmoi = request.form.get("vitrienmoi")
        
        chuyencu = request.form.get("chuyencu")
        chuyenmoi = request.form.get("chuyenmoi")
        
        gradecodecu = request.form.get("gradecodecu")
        gradecodemoi = request.form.get("gradecodemoi")
        
        sectioncodecu = request.form.get("sectioncodecu")
        sectioncodemoi = request.form.get("sectioncodemoi")
        
        hccategorycu = request.form.get("hccategorycu")
        hccategorymoi = request.form.get("hccategorymoi")
        
        departmentcu = request.form.get("departmentcu")
        departmentmoi = request.form.get("departmentmoi")
        
        sectiondescriptioncu = request.form.get("sectiondescriptioncu")
        sectiondescriptionmoi = request.form.get("sectiondescriptionmoi")
        
        employeetypecu = request.form.get("employeetypecu") 
        employeetypemoi = request.form.get("employeetypemoi")
        
        positioncodecu = request.form.get("positioncodecu") 
        positioncodemoi = request.form.get("positioncodemoi") 
        
        positioncodedescriptioncu = request.form.get("positioncodedescriptioncu") 
        positioncodedescriptionmoi = request.form.get("positioncodedescriptionmoi") 
        
        if loaidieuchuyen == "Chuyển vị trí":
            try:
                dieuchuyennhansu(mst,
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
                                )
            except Exception as e:
                app.logger.info(e)
                return redirect(f"/muc6_1")
            
        elif loaidieuchuyen == "Nghỉ việc":
            try:
                dichuyennghiviec(mst,
                    vitricu,
                    chuyencu,
                    ngaydieuchuyen,
                    ghichu
                            )
            except Exception as e:
                app.logger.info(e)
                return redirect(f"/muc6_1")
        else:
            try:
                dichuyennghithaisan(mst,
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
                            )
            except Exception as e:
                app.logger.info(e)
                return redirect(f"/muc6_1")
        return redirect(f"/muc6_1")
    else:  
        cacvitri= laycacvitri()
        return render_template("6_1.html",
                            cacvitri=cacvitri,
                            page="6.1 Điều chuyển chức vụ, bộ phận")
    
@app.route("/muc6_2", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def lichsucongtac():
    
    if request.method == "GET":
        mst = request.args.get("mst")
        hoten = request.args.get("hoten")
        ngay = request.args.get("ngay")
        kieudieuchuyen = request.args.get("kieudieuchuyen")
        rows = laylichsucongtac(mst,hoten,ngay,kieudieuchuyen)
        count = len(rows)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(rows)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = rows[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("6_2.html", page="6.2 Lịch sử công tác",
                               danhsach=paginated_rows, 
                               pagination=pagination,
                               mst=mst, 
                               count=count)

@app.route("/muc7_1_1", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def khaibaochamcong():
    if request.method == "GET":
        mst = request.args.get("mst")
        chuyen = request.args.get("chuyen") 
        phongban = request.args.get("phongban") 
        rows = laydanhsachcahientai(mst,chuyen,phongban)
        count = len(rows)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(rows)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = rows[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        cacca = laycacca()
        return render_template("7_1_1.html",
                                page="7.1.1 Đổi ca làm việc",
                                danhsach=paginated_rows,
                                pagination=pagination,
                                count=count,
                                cacca=cacca)

@app.route("/muc7_1_2", methods=["GET","POST"])
@login_required
def loichamcong():
    mst = request.args.get("mst")
    chuyen = request.args.get("chuyen")
    bophan = request.args.get("bophan")
    ngay = request.args.get("ngay")
    danhsach = laydanhsachloithe(mst,chuyen,bophan,ngay)
    count = len(danhsach)
    current_page = request.args.get(get_page_parameter(), type=int, default=1)
    per_page = 10
    total = len(danhsach)
    start = (current_page - 1) * per_page
    end = start + per_page
    paginated_rows = danhsach[start:end]
    pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
    return render_template("7_1_2.html",
                            page="7.1.2 Danh sách lỗi chấm công",
                            danhsach=paginated_rows, 
                            pagination=pagination,
                            count=count)

@app.route("/muc7_1_3", methods=["GET"])
@login_required
def diemdanhbu():
    
    mst = request.args.get("mst")
    hoten = request.args.get("hoten")
    chucvu = request.args.get("chucvu")
    chuyen = request.args.get("chuyen")
    bophan = request.args.get("bophan")
    loaidiemdanh = request.args.get("loaidiemdanh")
    ngay = request.args.get("ngay")
    lido = request.args.get("lido")
    trangthai = request.args.get("trangthai")
    danh_sach_chuyen = laydanhsachchuyen()
    danh_sach_bophan = laydanhsachbophan()
    danhsach = laydanhsachdiemdanhbu(mst,hoten,chucvu,chuyen,bophan,loaidiemdanh,ngay,lido,trangthai)
    count = len(danhsach)
    current_page = request.args.get(get_page_parameter(), type=int, default=1)
    per_page = 10
    total = len(danhsach)
    start = (current_page - 1) * per_page
    end = start + per_page
    paginated_rows = danhsach[start:end]
    pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
    return render_template("7_1_3.html",
                        page="7.1.3 Danh sách điểm danh bù",
                        danhsach=paginated_rows, 
                        pagination=pagination,
                        count=count,
                        danh_sach_chuyen=danh_sach_chuyen,
                        danh_sach_bophan=danh_sach_bophan)
 
@app.route("/muc7_1_4", methods=["GET"])
@login_required
def xinnghiphep():
    
        mst = request.args.get("mst")
        hoten = request.args.get("hoten")
        chucvu = request.args.get("chucvu")
        chuyen = request.args.get("chuyen")
        bophan = request.args.get("bophan")
        ngay = request.args.get("ngaynghi")
        lydo = request.args.get("lydo")
        trangthai = request.args.get("trangthai")
        danhsach = laydanhsachxinnghiphep(mst,hoten,chucvu,chuyen,bophan,ngay,lydo,trangthai)
        count = len(danhsach)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_4.html",
                            page="7.1.4 Danh sách xin nghỉ phép",
                            danhsach=paginated_rows, 
                            pagination=pagination,
                            count=count)

@app.route("/muc7_1_5", methods=["GET"])
@login_required
def xinnghikhongluong():
    if request.method == 'GET':
        mst = request.args.get("mst")
        hoten = request.args.get("hoten")
        chucvu = request.args.get("chucvu")
        chuyen = request.args.get("chuyen")
        bophan = request.args.get("bophan")
        ngay = request.args.get("ngaynghi")
        lydo = request.args.get("lydo")
        trangthai = request.args.get("trangthai")
        danhsach = laydanhsachxinnghikhongluong(mst,hoten,chucvu,chuyen,bophan,ngay,lydo,trangthai)
        count = len(danhsach)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_5.html",
                            page="7.1.5 Danh sách xin nghỉ không lương",
                            danhsach=paginated_rows,
                            pagination=pagination,
                            count=count)
    elif request.method == 'POST':
        mst = request.form.get("mst")
        hoten = request.form.get("hoten")
        chucvu = request.form.get("chucvu")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        ngay = request.form.get("ngaynghi")
        lydo = request.form.get("lydo")
        trangthai = request.form.get("trangthai")
        danhsach = laydanhsachxinnghikhongluong(mst,hoten,chucvu,chuyen,bophan,ngay,lydo,trangthai)
        data = []
        for row in danhsach:
            data.append({
                "Nhà máy": row[0],
                "Mã số thẻ": row[1],
                "Họ tên": row[2],
                "Chức danh": row[3],
                "Chuyền tổ": row[4], 
                "Phòng ban": row[5],
                "Ngày xin phép": row[6],
                "Tổng số phút": row[7],
                "Lý do": row[8],
                "Trạng thái": row[9]
            })
        df = pd.DataFrame(data)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT, f"xinnghikhongluong_{thoigian}.xlsx"), index=False)
    
        return send_file(os.path.join(FOLDER_XUAT, f"xinnghikhongluong_{thoigian}.xlsx"), as_attachment=True)
        
@app.route("/muc7_1_6", methods=["GET","POST"])
@login_required
def danhsachxinnghikhac():
    if request.method == "GET":
        mst = request.args.get("mst")
        ngaynghi = request.args.get("ngaynghi")
        loainghi = request.args.get("loainghi")
        danhsach = laydanhsachxinnghikhac(mst,ngaynghi,loainghi)
        count = len(danhsach)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_6.html", 
                               page="7.1.6 Danh sách xin nghỉ khác", 
                               danhsach=paginated_rows,
                                pagination=pagination,
                                count=count,)
    elif request.method == "POST":
        if 'file' not in request.files:
            return redirect("/muc7_1_6")
        file = request.files['file']
        if file.filename == '':
            return redirect("/muc7_1_6")
        if file:
            thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"xinnghikhac_{thoigian}.xlsx")
            file.save(filepath)
            data = pd.read_excel(filepath).to_dict(orient="records")
            for row in data:
                if row["Mã số thẻ"]!=np.nan:
                    try:
                        themxinnghikhac(
                            row["Mã công ty"],
                            int(row["Mã số thẻ"]),
                            row["Ngày nghỉ"],
                            int(row["Tổng số phút"]),
                            row["Loại nghỉ"]
                        )
                    except Exception as e:
                        app.logger.info(e)
                        break
        return redirect("/muc7_1_6")

@app.route("/muc7_1_7", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd','tk')
def dangkytangca():
    
    if request.method == "GET":
        mst = request.args.get("mst")
        phongban = request.args.get("phongban")
        chuyen = request.args.get("chuyen")
        ngay = request.args.get("ngay")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        danhsach = laydanhsachtangca(mst,phongban,chuyen,ngay,tungay,denngay)
        count = len(danhsach)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        messages = get_flashed_messages()
        return render_template("7_1_7.html", 
                               page="7.1.7 Đăng ký tăng ca",
                               danhsach=paginated_rows,
                               pagination=pagination,
                               count=count,
                               messages=messages
                               )
    elif request.method == "POST":
        data = []
        for row in danhsach:
            data.append({
                "Nhà máy": row[0],
                "MST": row[1],
                "Họ tên": row[2],
                "Chức vụ": row[3],
                "Chuyền tổ": row[4], 
                "Phòng ban": row[5],
                "Ngày đăng ký": row[6],
                "Giờ tăng ca": row[7],
                "Giờ tăng ca thực tế": row[8]
            })
        df = pd.DataFrame(data)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT,f"tangca_{thoigian}.xlsx"), index=False)

        return send_file(os.path.join(FOLDER_XUAT,f"tangca_{thoigian}.xlsx"), as_attachment=True)

@app.route("/muc7_1_8", methods=["GET","POST"])
@login_required
def chamcongtudong():
    
    mst = request.args.get("mst")
    phongban = request.args.get("phongban")
    tungay = request.args.get("tungay")
    denngay = request.args.get("denngay")
    phanloai = request.args.get("phanloai")
    rows = laydanhsachchamcong(mst,phongban,tungay,denngay,phanloai)
    count = len(rows)
    current_page = request.args.get(get_page_parameter(), type=int, default=1)
    per_page = 10
    total = len(rows)
    start = (current_page - 1) * per_page
    end = start + per_page
    paginated_rows = rows[start:end]
    pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
    return render_template("7_1_8.html", page="7.1.8 Bảng công 5 ngày gần nhất",
                           danhsach=paginated_rows, 
                           pagination=pagination,
                           count=count)
                
@app.route("/muc7_1_9", methods=["GET","POST"])
@login_required
def chamcongtudongchot():
    
    mst = request.args.get("mst")
    phongban = request.args.get("phongban")
    tungay = request.args.get("tungay")
    denngay = request.args.get("denngay")
    phanloai = request.args.get("phanloai")
    rows = laydanhsachchamcongchot(mst,phongban,tungay,denngay,phanloai)
    count = len(rows)
    current_page = request.args.get(get_page_parameter(), type=int, default=1)
    per_page = 10
    total = len(rows)
    start = (current_page - 1) * per_page
    end = start + per_page
    paginated_rows = rows[start:end]
    danhsachphongban = laycacphongban()
    pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
    return render_template("7_1_9.html", page="7.1.9 Bảng chấm công chốt",
                           danhsach=paginated_rows, 
                           pagination=pagination,
                           count=count,
                           danhsachphongban=danhsachphongban)

@app.route("/muc7_1_10", methods=["GET","POST"])
@login_required
def danhsachphepton():
    if request.method == "GET":
        mst = request.args.get("mst")
        danhsach = laydanhsachphepton(mst)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_10.html", page="7.1.10 Danh sách phép tồn",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    if request.method == "POST":
        mst = request.form.get("mst")
        danhsach = laydanhsachphepton(mst)
        result = []
        for row in danhsach:
            result.append({
                "Mã công ty": row[0],
                "Mã số thẻ": row[1],
                "Họ tên": row[2],
                "Chức danh": row[3],
                "Tháng": row[4],
                "Năm": row[5],
                "Số phút phép được dùng": row[6],
                "Số phút phép đã chốt": row[7],
                "Số phút phép chưa dùng": row[8],
                "Số phút phép cho dùng": row[9],
                "Số phút phép còn lại": row[10]
            })
        df = pd.DataFrame(result)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT, f"phepton_{thoigian}.xlsx"), index=False)
        
        return send_file(os.path.join(FOLDER_XUAT, f"phepton_{thoigian}.xlsx"), as_attachment=True)
    
@app.route("/muc7_1_10", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def capnhathulieuchamcong():
    if request.method == "GET":
        return render_template("7_1_10.html", page="7.1.10 Cập nhật dữ liệu chấm công")
    elif request.method == "POST":
        if (current_user.macongty == 'NT1' and current_user.masothe in [12579,2833] ) or (current_user.macongty == 'NT' and current_user.masothe in [12579,2833] ):
            thread = Thread(target=themdulieuchamcong2ngay)
            thread.start()    
            return render_template("7_1_10.html", 
                                page="7.1.10 Cập nhật dữ liệu chấm công", 
                                messages=["Đang cập nhật dữ liệu chấm công mới nhất, vui lòng đợi khoảng 10 phút ...","Trong lúc cập nhật các phần liên quan sẽ không hoạt động, vui lòng đợi đến khi cập nhật xong !!!"])
        else:
            return render_template("7_1_10.html", 
                                page="7.1.10 Cập nhật dữ liệu chấm công", 
                                messages=["Bạn không có quyền cập nhật, vui lòng liên hệ HR !!!"])
            
@app.route("/muc8_1", methods=["GET","POST"])
@login_required
def ykienkhieunai():

    return render_template("8_1.html", page="8.1 Danh sách ý kiến khiếu nại")

@app.route("/muc8_2", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def capnhatykienkhieunai():

    return render_template("8_2.html", page="8.2 Cập nhật ý kiến khiếu nại")
    
@app.route("/muc9_1", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def xulykiluat():
    
    if request.method == "GET":
        danhsach = laydanhsachkyluat()
        return render_template("9_1.html", page="9.1 Xử lý kỉ luật",danhsach=danhsach)
    else:
        mst = request.form.get("mst")
        hoten = request.form.get("hoten")
        chucvu = request.form.get("chucvu")
        bophan = request.form.get("bophan")
        chuyento = request.form.get("chuyento")
        ngayvao = request.form.get("ngayvao")
        ngayvipham = request.form.get("ngayvipham")
        diadiem = request.form.get("diadiem")
        ngaylapbienban = request.form.get("ngaylapbienban")
        noidung = request.form.get("noidung")
        bienphap = request.form.get("bienphap")
        try:
            themdanhsachkyluat(mst,hoten,chucvu,bophan,chuyento,ngayvao,ngayvipham,diadiem,ngaylapbienban,noidung,bienphap)
        except Exception as ex:
            app.logger.info(ex)
        return redirect("/muc9_1") 
    
@app.route("/muc10_1", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def phongvannghiviec():
        
    return render_template("10_1.html", page="10.1 Tổng hợp phỏng vấn nghỉ việc")

    
@app.route("/muc10_3", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def inchamduthopdong():
     
    if request.method == "GET":
        return render_template("10_3.html", page="10.3 In chấm dứt hợp đồng")
    elif request.method == "POST":
        mst = request.form.get("mst")
        ngaylamhopdong = datetime.strptime(request.form.get("ngaylamhd"),"%Y-%m-%d").strftime("%d")
        thanglamhopdong = datetime.strptime(request.form.get("ngaylamhd"),"%Y-%m-%d").strftime("%m")
        namlamhopdong = datetime.strptime(request.form.get("ngaylamhd"),"%Y-%m-%d").strftime("%Y")
        tennhanvien = request.form.get("hoten")
        chucvu = request.form.get("chucvu")
        ngaynghi = datetime.strptime(request.form.get("ngaynghi"),"%Y-%m-%d").strftime("%d/%m/%Y")
        ngaysinh = request.form.get("ngaysinh")
        diachi = request.form.get("diachi")
        bophan = request.form.get("bophan")
        lydo = request.form.get("lydo")
        try:
            file = inchamduthd(mst,
                ngaylamhopdong,
                thanglamhopdong,
                namlamhopdong,
                tennhanvien,
                chucvu,
                ngaynghi,
                ngaysinh,
                diachi,
                bophan,
                lydo)
            if file:
                return send_file(file, as_attachment=True, download_name="chamduthopdong.xlsx")
            else:
                app.logger.info("NO FILE")
                return redirect("/muc10_3")
        except Exception as e:
            app.logger.info(e)
            return redirect("/muc10_3")  

#############################################
#                "OTHER ENDPOINT"           #
#############################################

@app.route("/thaydoiphanquyen", methods=["POST"])
def thaydoiphanquyen():
    if request.method == "POST":
        userid= request.form["id"]
        newrole = request.form["newrole"]
        user = Users.query.filter_by(id=userid).first()
        if user:
            user.role = newrole
            db.session.commit()
        return redirect("/admin")
    
@app.route("/taimautangcanhom", methods=["POST"])
def taimautangcanhom():
    if request.method == "POST":        
        return send_file(FILE_MAU_DANGKY_TANGCA_NHOM, as_attachment=True)  
    
@app.route("/capnhattrangthaiungvien", methods=["POST"])
def capnhattrangthaiungvien():
    try:
        sdt = request.args.get("sdt")
        trangthai = request.args.get("trangthaimoi")
        luuhoso = request.args.get("luuhoso")
        if capnhattrangthaimoiungvien(sdt, trangthai, luuhoso):
            return {"status": "success"}, 200
        else:
            return {"status": "fail"}, 400
    except Exception as e:
        app.logger.info(e)
        return {"status": "fail"}, 400

@app.route("/laythongtincccd", methods=["POST"])
def laythongtincccd():
    
    conn = pyodbc.connect(used_db)
    cursor = conn.cursor()

    if request.method == "POST":
        cccd = request.args.get("cccd")  # lấy giá trị cccd từ form data
        if cccd:
            employee = cursor.execute("SELECT * FROM HR.dbo.Dang_ky_thong_tin WHERE CCCD = ?", cccd).fetchone()
            conn.close()
            if employee:
                tamtru = employee[10]
                data_tamtru = tamtru.split(",")
                
                # Chuyển đổi employee thành dict và trả về dạng JSON
                employee_dict = {
                    "Nhà máy": employee[0],
                    "Vị trí ứng tuyển": employee[1],
                    "Họ tên": employee[2],
                    "Số điện thoại": employee[3],
                    "Số CCCD": employee[4],
                    "Dân tộc": employee[5],
                    "Quốc tịch": employee[7],
                    "Tôn giáo": employee[6],
                    "Trình độ học vấn": employee[8],
                    "Nơi sinh" : employee[9],
                    "Tạm trú" : tamtru,
                    "Phường/Xã": data_tamtru[1] if len(data_tamtru) > 1 else "",
                    "Quận/huyện": data_tamtru[2] if len(data_tamtru) > 2 else "",
                    "Tỉnh/Thành phố": data_tamtru[3] if len(data_tamtru) > 3 else "",
                    "Số BHXH": employee[11],
                    "Mã số thuế": employee[12],
                    "Ngân hàng": employee[13],
                    "Số tài khoản": employee[14],
                    "Tên người thân": employee[15],
                    "SĐT người thân": employee[16],
                    "Kênh tuyển dụng": employee[17], 
                    "Kinh nghiệm": employee[18],
                    "Mức lương": employee[19], 
                    "Ngày có thể nhận việc": employee[20],
                    "Con nhỏ": employee[21],
                    "Tên con 1": employee[22],
                    "Ngày sinh con 1": employee[23],
                    "Tên con 2": employee[24],
                    "Ngày sinh con 2": employee[25],
                    "Tên con 3": employee[26],
                    "Ngày sinh con 3": employee[27],
                    "Tên con 4": employee[28],
                    "Ngày sinh con 4": employee[29],
                    "Tên con 5": employee[30],
                    "Ngày sinh con 5": employee[31],
                    "Ngày gửi": employee[32],
                    "Trạng thái": employee[33],
                    "Ngày cập nhật": employee[34],
                    "Ngày hẹn đi làm": employee[35],
                    "Hiệu suất": employee[36],
                    "Loại máy": employee[37],
                    "Ghi chú": employee[38]
                }
                return jsonify(employee_dict)
            else:
                return jsonify({"error": "Employee not found"}), 404
        else:
            return jsonify({"error": "CCCD is required"}), 400

@app.route("/kiemtrathongtinnld", methods=["POST"])
def kiemtrathongtinnld():

    if request.method == "POST":
        mst = request.args.get("masothe")
        if mst:
            users = laydanhsachtheomst(mst)
            if users:
                return jsonify(users[0]), 200
            else:
                return jsonify({"error": "User not found"}), 404
        else:
            return jsonify({"error": "MST is required"}), 400

@app.route("/dangkitangcacanhan", methods=["POST"])  
def dangkitangcacanhan():
    try:
        mst = request.form.get("mst")
        giotangca = request.form.get("giotangca")
        ngaytangca = request.form.get("ngaytangca")
        # app.logger.info(mst,giotangca,ngaytangca)
        user = laydanhsachtheomst(mst)
        # app.logger.info(user)
        if user:
            user = user[0]
            # app.logger.info(user)
            if kiemtrathuki(current_user.masothe,user['Line']):
                if insert_tangca(current_user.macongty,
                            mst,
                            user['Họ tên'],
                            user['Chức vụ'],
                            user['Line'],
                            user['Department'],
                            ngaytangca,
                            giotangca):
                    flash(f"{current_user.masothe} đã đăng ký tăng ca cho {mst} thành công", "success")
                else:
                    flash(f"{current_user.masothe} đã đăng ký tăng ca cho {mst} thất bại", "danger")
                return redirect(f"/muc7_1_6?ngay={ngaytangca}")
            else:
                flash(f"{current_user.masothe} không được phép đăng ký tăng ca cho {mst}", "danger")
                return redirect(f"/muc7_1_6")
        else:
            flash(f"Không tìm thấy nhân viên có {mst}", "danger")
            return redirect(f"/muc7_1_6")  
    except Exception as e:
        flash(f"Đăng ký tăng ca lỗi: {e}")
        return redirect(f"/muc7_1_6")
    
@app.route("/dangkitangcanhom", methods=["POST"])   
def dangkitangcanhom():
    
    if request.method == "POST":
        try:
            file = request.files['file']
            if file:
                ngaylam = datetime.now().strftime("%d%m%Y%H%M%S")
                filepath = os.path.join(FOLDER_NHAP, f"tangca_{current_user.phongban}_{ngaylam}.xlsx")
                file.save(filepath)
                data = pd.read_excel(filepath).to_dict(orient="records")
                for row in data:
                    kiemtra = kiemtrathuki(current_user.masothe,row["Chuyền tổ"])
                    if kiemtra:
                        app.logger.info(f"Thư ký {current_user.masothe} {row['Chuyền tổ']} dang ki tang ca cho {row['MST']} {row['Họ tên']} {row['Chức vụ']} {row['Phòng ban']} {row['Ngày đăng ký']} {row['Giờ tăng ca']}")
                        try:
                            if insert_tangca(current_user.macongty,row["MST"],row["Họ tên"],row["Chức vụ"],row["Chuyền tổ"],row["Phòng ban"],row["Ngày đăng ký"],row["Giờ tăng ca"]):
                                flash(f"{current_user.masothe} đã đăng ký tăng ca cho {row['MST']} thành công", "success")
                            else:
                                flash(f"{current_user.masothe} đã đăng ký tăng ca cho {row['MST']} thất bại", "danger")
                        except Exception as e:
                            app.logger.info(e)   
                    else:
                        flash(f"{current_user.masothe} không được đăng ký tăng ca cho {row['MST']}")            
            return redirect("/muc7_1_6")
        except Exception as e:
            flash(f"{current_user.masothe} không được đăng ký tăng ca cho {row['MST']} lỗi: {e}")
            return redirect("/muc7_1_6")

@app.route("/export_dstc", methods=["POST"])
def export_dstc():
    mst = request.form.get("mst")
    phongban = request.form.get("phongban")
    ngay = request.form.get("ngay")
    tungay = request.form.get("tungay")
    denngay = request.form.get("denngay")
    danhsach = laydanhsachtangca(mst,phongban,ngay,tungay,denngay)
    result = []
    for row in danhsach:
        result.append(
            {
                'Nhà máy': row[0],
                'MST': row[1],
                'Họ tên': row[2],
                'Chức danh': row[3],
                'Chuyền': row[4],
                'Phòng ban': row[5],
                'Ngày đăng ký': datetime.strptime(row[6], "%Y-%m-%d").strftime("%d/%m/%Y"),
                'Giờ tăng ca': row[7][:5] if row[7] else "",
            }
        )
    df = pd.DataFrame(result)
    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
    df.to_excel(os.path.join(FOLDER_XUAT, f"danhsach_{thoigian}.xlsx"), index=False)
    
    return send_file(os.path.join(FOLDER_XUAT, f"danhsach_{thoigian}.xlsx"), as_attachment=True)
    
       
@app.route("/export_dsnv", methods=["POST"])
def export_dsnv():
    
    mst = request.form.get("Mã số thẻ")
    hoten = request.form.get("Họ tên")
    sdt = request.form.get("Số điện thoại")
    cccd = request.form.get("Căn cước công dân")
    gioitinh = request.form.get("Giới tính")
    vaotungay = request.form.get("Vào từ ngày")
    vaodenngay = request.form.get("Vào đến ngày")
    nghitungay = request.form.get("Nghỉ từ ngày")
    nghidenngay = request.form.get("Nghỉ đến ngày")
    phongban = request.form.get("Phòng ban")
    chucvu = request.form.get("Chức danh")
    trangthai = request.form.get("Trạng thái")
    hccategory = request.form.get("Headcount Category")
        
    users = laydanhsachuser(mst, hoten, sdt, cccd, gioitinh, vaotungay, vaodenngay, nghitungay, nghidenngay, phongban, trangthai, hccategory, chucvu)      
    df = pd.DataFrame(users)
    df.to_excel(os.path.join(FOLDER_XUAT, "danhsach.xlsx"), index=False)
    
    return send_file(os.path.join(FOLDER_XUAT, "danhsach.xlsx"),  as_attachment=True)  

@app.route("/export_dslt", methods=["POST"])
def export_dslt():
    mst = request.form.get("mst")
    chuyen = request.form.get("chuyen")
    bophan = request.form.get("bophan")
    ngay = request.form.get("ngay")
    rows = laydanhsachloithe(mst, chuyen, bophan, ngay)
    df = pd.DataFrame(rows)
    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
    df.to_excel(os.path.join(FOLDER_XUAT, f"danhsachloithe_{thoigian}.xlsx"), index=False)
    
    return send_file(os.path.join(FOLDER_XUAT, f"danhsachloithe_{thoigian}.xlsx"), as_attachment=True) 

@app.route("/export_dsddb", methods=["POST"])
def export_dsddb():
    
    mst = request.form.get("mst")
    chuyen = request.form.get("chuyen")
    bophan = request.form.get("bophan")
    hoten = request.form.get("hoten")
    chucvu = request.form.get("chucvu")
    ngaydiemdanh = request.form.get("ngay")
    lydo = request.form.get("lydo")
    trangthai = request.form.get("trangthai")
    loaidiemdanh = request.form.get("loaidiemdanh")
    
    rows = laydanhsachdiemdanhbu(mst,hoten,chucvu,chuyen,bophan,loaidiemdanh,ngaydiemdanh,lydo,trangthai)
    result = []
    for row in rows:
        result.append({
            "Nhà máy": row[0],
            "MST": row[1],
            "Họ tên": row[2],
            "Chức vụ": row[3],
            "Chuyền tổ": row[4],
            "Bộ phận": row[5],
            "Loại điểm danh": row[6],
            "Ngày điểm danh": datetime.strptime(row[7], "%Y-%m-%d").strftime("%d/%m/%Y"),
            "Giờ điểm danh": row[8],
            "Lý do": row[9],
            "Trạng thái": row[10]
        })
    
    df = pd.DataFrame(result)
    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
    df.to_excel(os.path.join(FOLDER_XUAT, f"diemdanhbu_{thoigian}.xlsx"), index=False) # f"diemdanhbu_{thoigian}.xlsx", index=False)
    
    return send_file(os.path.join(FOLDER_XUAT, f"diemdanhbu_{thoigian}.xlsx"), as_attachment=True)  

@app.route("/export_dsxnp", methods=["POST"])
def export_dsxnp():
    
    mst = request.form.get("mst")
    hoten = request.form.get("hoten")
    chucvu = request.form.get("chucvu")
    chuyen = request.form.get("chuyen")
    bophan = request.form.get("bophan")
    ngay = request.form.get("ngaynghi")
    lydo = request.form.get("lydo")
    trangthai = request.form.get("trangthai")
    danhsach = laydanhsachxinnghiphep(mst,hoten,chucvu,chuyen,bophan,ngay,lydo,trangthai)
    result = []
    for row in danhsach:
        result.append({
            'Mã công ty': row[0],
            'Mã số thẻ': row[1],
            'Họ tên': row[2],
            'Chức vụ': row[3],
            'Chuyền tổ': row[4],
            'Phòng ban': row[5],
            'Ngày nghỉ phép': datetime.strptime(row[6], "%Y-%m-%d").strftime("%d/%m/%Y"),
            'Tổng số phút': row[7],
            'Lý do': row[8],
            'Trạng thái': row[9]
        })
    df = pd.DataFrame(result)
    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
    df.to_excel(os.path.join(FOLDER_XUAT, f"xinnghiphep_{thoigian}.xlsx"), index=False)
    
    return send_file(os.path.join(FOLDER_XUAT, f"xinnghiphep_{thoigian}.xlsx"), as_attachment=True) 

@app.route("/export_dsdktt", methods=["POST"])
def export_dsdktt():
    
    sdt = request.form.get("sdt")
    cccd = request.form.get("cccd")
    ngaygui = request.form.get("ngaygui")
    rows = laydanhsachdangkytuyendung(sdt, cccd, ngaygui)   
    df = pd.DataFrame(rows)
    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
    df.to_excel(os.path.join(FOLDER_XUAT, f"tuyendung_{thoigian}.xlsx"), index=False)
    
    return send_file(os.path.join(FOLDER_XUAT, f"tuyendung_{thoigian}.xlsx"), as_attachment=True)  
      
@app.route("/check_hcname", methods=["POST"])
def check_hcname():
    jobtitle = request.args.get("vitri")
    line = request.args.get("line")
    hcname = layhcname(jobtitle,line)
    if not hcname:
        return jsonify({
            "Line": "",
            "Detail_job_title_VN": "",
            "Detail_job_title_EN": "",
            "Employee_type": "",
            "Position_code": "",
            "Position_code_description": "",
            "Grade_code": "",
            "HC_category": "",
            "Factory": "",
            "Department": "",
            "Section_code": "",
            "Section_description": ""
        })
    return jsonify({
        "Line": hcname[0],
        "Detail_job_title_VN": hcname[1],
        "Detail_job_title_EN": hcname[2],
        "Employee_type": hcname[3],
        "Position_code": hcname[4],
        "Position_code_description": hcname[5],
        "Grade_code": hcname[6],
        "HC_category": hcname[7],
        "Factory": hcname[8],
        "Department": hcname[9],
        "Section_code": hcname[10],
        "Section_description": hcname[11]        
    })

@app.route("/check_line_from_detailjob", methods=["POST"])
def check_line_from_detailjob():
    vitri = request.args.get("vitrimoi")
    cacline = laydanhsachlinetheovitri(vitri)
    return jsonify(cacline)

@app.route("/xoanhanviencu", methods=["GET"])
def xoanhanviencu():
    mst = request.args.get("mst")
    try:
        app.logger.info(xoanhanvien(mst))
        return redirect(url_for('timdanhsachnhanvien', mst=mst))
    except Exception as e:
        app.logger.info(e)
        return redirect(url_for('timdanhsachnhanvien', mst=mst))

@app.route("/doicacanhan", methods=["POST"])
def doicacanhan():
    try:
        mst = request.form.get("mst")
        cacu = request.form.get("cacu")
        camoi = request.form.get("camoi")
        ngaybatdau = request.form.get("ngaybatdau")
        ngayketthuc = request.form.get("ngayketthuc")
        themdoicamoi(mst,cacu,camoi,ngaybatdau,ngayketthuc)
        flash(f"Đổi ca thành công cho MST {mst} thành {camoi}", "success")
        return redirect("/muc7_1_1")
    except Exception as e:
        app.logger.info(e)
        flash("Đổi ca bị lỗi, bạn vui lòng kiểm tra lại !!!")
        return redirect("/muc7_1_1")
    
@app.route("/doicanhom", methods=["POST"])
def doicanhom():
    try:
        cacongty = request.form.get("cacongty")
        if cacongty:
            danhsach = laydanhsachusercacongty(current_user.macongty)
        else:
            phongban = request.form.get("phongban")
            if phongban:
                danhsach = laydanhsachusertheophongban(phongban)
            else:
                chuyen = request.form.get("chuyento")   
                if chuyen: 
                    danhsach = laydanhsachusertheoline(chuyen)
                else:
                    file = request.files.get("file")
                    app.logger.info(file)
                    if file:
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
                        filepath = os.path.join(FOLDER_NHAP, f"doicanhom_{thoigian}.xlsx")
                        file.save(filepath)
                        data = pd.read_excel(filepath).to_dict(orient="records")
                        for row in data:
                            themdoicamoi(row['Mã số thẻ'],laycahientai(row['Mã số thẻ']),row['Ca mới'],row['Từ ngày'],row['Đến ngày'])
                    danhsach = None
        if danhsach:
            camoi = request.form.get("camoinhom")
            ngaybatdau = request.form.get("ngaybatdau")
            ngayketthuc = request.form.get("ngayketthuc")
            
            for user in danhsach:
                themdoicamoi(user['MST'],laycahientai(user['MST']),camoi,ngaybatdau,ngayketthuc)
            cacmst = [user['MST'] for user in danhsach]
            flash(f"Đổi ca thành công các MST {str(cacmst)} thành {camoi}", "success")
        return redirect("/muc7_1_1")
    except Exception as e:
        app.logger.info(e)
        flash("Đổi ca bị lỗi, bạn vui lòng kiểm tra lại !!!")
        return redirect("/muc7_1_1")
        
@app.route("/laycatheomst", methods=["POST"])
def laycatheomst():
    mst = request.args.get("mst")
    ca = laycahientai(mst)
    return jsonify({
        "Ca": ca
    })

@app.route("/taifilexinnghiphepkhacmau", methods=["POST"])
def taifilexinnghiphepkhacmau():
    file = FILE_MAU_DANGKY_XINNGHIKHAC
    return send_file(file, as_attachment=True)

@app.route("/taimaudoicanhom", methods=["POST"])
def taimaudoicanhom():
    file = FILE_MAU_DANGKY_DOICA_NHOM
    return send_file(file, as_attachment=True)

@app.route("/export_dscc", methods=["POST"])
def export_dscc():
    mst = request.form.get('mst')
    phongban = request.form.get('phongban')
    tungay = request.form.get("tungay")
    denngay = request.form.get("denngay")
    phanloai = request.form.get("phanloai")
    danhsach = laydanhsachchamcong(mst,phongban,tungay,denngay,phanloai)
    result = []
    for row in danhsach:
        result.append(
            {
                'Nhà máy': row[0],
                'MST': row[1],
                'Họ tên': row[2],
                'Chức danh': row[3],
                'Chuyền': row[4],
                'Phòng ban': row[5],
                'Cấp bậc': row[6],
                'Ngày': row[7],
                'Ca': row[8],
                'Số giờ làm việc': row[9],
                'Giờ vào': row[10],
                'Giờ ra': row[11],
                'Phút HC': row[12],
                'Phút nghỉ phép': row[13],
                'Phút tăng ca 100%': row[14],
                'Phút tăng ca 150%': row[15],
                'Phút tăng ca đêm': row[16],
                'Phút nghỉ không lương': row[17],
                'Phút nghỉ khác': row[18],
                'Loại nghỉ khác': row[19],
                'Phân loại': row[20]
            }
        )
    df = pd.DataFrame(result)
    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
    df.to_excel(os.path.join(FOLDER_XUAT, f"bangcong_{thoigian}.xlsx"), index=False)
    
    return send_file(os.path.join(FOLDER_XUAT, f"bangcong_{thoigian}.xlsx"), as_attachment=True)

@app.route("/export_dscctt", methods=["POST"])
def export_dscctt():
    mst = request.form.get('mst')
    phongban = request.form.get('phongban')
    tungay = request.form.get("tungay")
    denngay = request.form.get("denngay")
    phanloai = request.form.get("phanloai")
    danhsach = laydanhsachchamcongchot(mst,phongban,tungay,denngay,phanloai)
    result = []
    for row in danhsach:
        result.append(
            {
                'Nhà máy': row[0],
                'MST': row[1],
                'Họ tên': row[2],
                'Chức danh': row[3],
                'Chuyền': row[4],
                'Phòng ban': row[5],
                'Cấp bậc': row[6],
                'Ngày': row[7],
                'Ca': row[8],
                'Số giờ làm việc': row[9],
                'Giờ vào': row[10],
                'Giờ ra': row[11],
                'Phút HC': row[12],
                'Phút nghỉ phép': row[13],
                'Phút tăng ca 100%': row[14],
                'Phút tăng ca 150%': row[15],
                'Phút tăng ca đêm': row[16],
                'Phút nghỉ không lương': row[17],
                'Phút nghỉ khác': row[18],
                'Loại nghỉ khác': row[19],
                'Phân loại': row[20]
            }
        )
    df = pd.DataFrame(result)
    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
    df.to_excel(os.path.join(FOLDER_XUAT, f"bangcong_{thoigian}.xlsx"), index=False)
    
    return send_file(os.path.join(FOLDER_XUAT, f"bangcong_{thoigian}.xlsx"), as_attachment=True)

@app.route("/export_dsxnk", methods=["POST"])
def export_dsxnk():
    mst = request.form.get('mst')
    ngaynghi = request.form.get('ngaynghi')
    loainghi = request.form.get("loainghi")
    danhsach = laydanhsachxinnghikhac(mst,ngaynghi,loainghi)
    result = []
    for row in danhsach:
        result.append(
            {
                'Nhà máy': row[0],
                'MST': row[1],
                'Ngày nghỉ': datetime.strptime(row[2], '%Y-%m-%d').strftime('%d/%m/%Y'),
                'Tổng số phút': row[3],
                'Loại nghỉ': row[4],
            }
        )
    df = pd.DataFrame(result)
    thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
    df.to_excel(os.path.join(FOLDER_XUAT, f"xinnghikhac_{thoigian}.xlsx"), index=False)
    
    return send_file(os.path.join(FOLDER_XUAT, f"xinnghikhac_{thoigian}.xlsx"), as_attachment=True)

@app.route("/thuky_kiemtra_diemdanhbu", methods=["POST"])
def thukykiemtradiemdanhbu():
    if request.method == "POST":
        try:
            mst_filter = request.form["mst_filter"]
            hoten_filter = request.form["hoten_filter"]
            chucvu_filter = request.form["chucvu_filter"]
            chuyen_filter = request.form["chuyen_filter"]
            bophan_filter = request.form["bophan_filter"]
            loaidiemdanh_filter = request.form["loaidiemdanh_filter"]
            ngay_filter = request.form["ngay_filter"]
            lydo_filter = request.form["lydo_filter"]
            trangthai_filter = request.form["trangthai_filter"]
            chuyen = request.form["chuyen"]
            mstduyet = current_user.masothe
            kiemtra = request.form["kiemtra"]
            id = request.form["id"]
            mstdiemdanh = request.form["mst_diemdanh"]
            if mstdiemdanh==mstduyet:
                flash(f"Bạn không thể kiểm tra cho chính mình, vui lòng liên hệ thư ký !!!")
                return redirect(f"/muc7_1_3?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&loaidiemdanh={loaidiemdanh_filter}&ngay={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
            if thuky_duoc_phanquyen(mstduyet,chuyen):
                if kiemtra == "Kiểm tra":    
                    thuky_dakiemtra_diemdanhbu(id)
                    flash(f"Thư ký {current_user.hoten} đã kiểm tra phiếu điểm danh bù số {id} !!!")
                    return redirect(f"/muc7_1_3?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&loaidiemdanh={loaidiemdanh_filter}&ngay={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
                else:
                    thuky_tuchoi_diemdanhbu(id)
                    flash(f"Thư ký {current_user.hoten} đã từ chối điểm danh bù phiếu số {id}  !!!")
                    return redirect(f"/muc7_1_3?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&loaidiemdanh={loaidiemdanh_filter}&ngay={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
            else:
                flash(f"{current_user.hoten} không có quyền điểm danh chuyền {chuyen} !!!")
            return redirect(f"/muc7_1_3?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&loaidiemdanh={loaidiemdanh_filter}&ngay={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
        except Exception as e:
            flash(f"Lỗi thư ký điểm danh bù: {e}")
            return redirect(f"/muc7_1_3?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&loaidiemdanh={loaidiemdanh_filter}&ngay={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
        
@app.route("/quanly_pheduyet_diemdanhbu", methods=["POST"])
def quanlypheduyetdiemdanhbu():
    if request.method == "POST":
        try:
            mst_filter = request.form["mst_filter"]
            hoten_filter = request.form["hoten_filter"]
            chucvu_filter = request.form["chucvu_filter"]
            chuyen_filter = request.form["chuyen_filter"]
            bophan_filter = request.form["bophan_filter"]
            loaidiemdanh_filter = request.form["loaidiemdanh_filter"]
            ngay_filter = request.form["ngay_filter"]
            lydo_filter = request.form["lydo_filter"]
            trangthai_filter = request.form["trangthai_filter"]
            chuyen = request.form["chuyen"]
            mstduyet = current_user.masothe
            pheduyet = request.form["pheduyet"]
            id = request.form["id"]
            mstdiemdanh = request.form["mst_diemdanh"]
            if mstdiemdanh==mstduyet:
                flash(f"Bạn không thể phê duyệt cho chính mình, vui lòng liên hệ thư ký !!!")
                return redirect(f"/muc7_1_3?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&loaidiemdanh={loaidiemdanh_filter}&ngay={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
            if quanly_duoc_phanquyen(mstduyet,chuyen):
                if pheduyet == "Phê duyệt":    
                    quanly_pheduyet_diemdanhbu(id)
                    flash(f"Quản lý {current_user.hoten} đã phê duyệt điểm danh bù cho phiếu số {id} !!!")
                else:
                    quanly_tuchoi_diemdanhbu(id)
                    flash(f"Quản lý {current_user.hoten} đã từ chối điểm danh bù cho phiếu số {id}  !!!")
            else:
                flash(f"{current_user.hoten} không có quyền phê duyệt !!!")
            return redirect(f"/muc7_1_3?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&loaidiemdanh={loaidiemdanh_filter}&ngay={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
        except Exception as e:
            flash(f"Lỗi quản lý phê duyệt điểm danh bù: {e}")
            return redirect(f"/muc7_1_3?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&loaidiemdanh={loaidiemdanh_filter}&ngay={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")

@app.route("/thuky_kiemtra_xinnghiphep", methods=["POST"])
def thukykiemtraxinnghiphep():
    if request.method == "POST":
        try:
            mst_filter = request.form["mst_filter"]
            hoten_filter = request.form["hoten_filter"]
            chucvu_filter = request.form["chucvu_filter"]
            chuyen_filter = request.form["chuyen_filter"]
            bophan_filter = request.form["bophan_filter"]
            ngay_filter = request.form["ngay_filter"]
            trangthai_filter = request.form["trangthai_filter"]
            chuyen = request.form["chuyen"]
            mstduyet = current_user.masothe
            kiemtra = request.form["kiemtra"]
            id = request.form["id"]
            mstxinnghiphep = request.form["mst_xinnghiphep"]
            if mstxinnghiphep==mstduyet:
                flash(f"Bạn không thể kiểm tra cho chính mình, vui lòng liên hệ thư ký !!!")
                return redirect(f"/muc7_1_4?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&trangthai={trangthai_filter}")
            if thuky_duoc_phanquyen(mstduyet,chuyen):
                if kiemtra == "Kiểm tra":    
                    thuky_dakiemtra_xinnghiphep(id)
                    flash(f"Thư ký {current_user.hoten} đã kiểm tra phiếu xin nghỉ phép số {id} !!!")
                    return redirect(f"/muc7_1_4?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&trangthai={trangthai_filter}")
                else:
                    thuky_tuchoi_xinnghiphep(id)
                    flash(f"Thư ký {current_user.hoten} từ chối phiếu nghỉ phép số {id} !!!")
                    return redirect(f"/muc7_1_4?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&trangthai={trangthai_filter}")
            else:
                flash(f"{current_user.hoten} không có quyền kiểm tra !!!")
            return redirect(f"/muc7_1_4?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&trangthai={trangthai_filter}")
        except Exception as e:
            flash(f"Lỗi thư ký kiểm tra xin nghỉ phép: {e}")
            return redirect(f"/muc7_1_4?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&trangthai={trangthai_filter}")
        
@app.route("/quanly_pheduyet_xinnghiphep", methods=["POST"])
def quanlypheduyetxinnghiphep():
    if request.method == "POST":
        try:
            mst_filter = request.form["mst_filter"]
            hoten_filter = request.form["hoten_filter"]
            chucvu_filter = request.form["chucvu_filter"]
            chuyen_filter = request.form["chuyen_filter"]
            bophan_filter = request.form["bophan_filter"]
            ngay_filter = request.form["ngay_filter"]
            trangthai_filter = request.form["trangthai_filter"]
            chuyen = request.form["chuyen"]
            chuyen = request.form["chuyen"]
            mstduyet = current_user.masothe
            pheduyet = request.form["pheduyet"]
            id = request.form["id"]
            mstxinnghiphep = request.form["mst_xinnghiphep"]
            if mstxinnghiphep==mstduyet:
                flash(f"Bạn không thể phê duyệt cho chính mình, vui lòng liên hệ thư ký !!!")
                return redirect(f"/muc7_1_4?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&trangthai={trangthai_filter}")
            if quanly_duoc_phanquyen(mstduyet,chuyen):
                if pheduyet == "Phê duyệt":    
                    quanly_pheduyet_xinnghiphep(id)
                    flash(f"Quản lý {current_user.hoten} đã hê duyệt cho phiếu xin nghỉ phép số {id} !!!")
                    return redirect(f"/muc7_1_4?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&trangthai={trangthai_filter}")
                else:
                    quanly_tuchoi_xinnghiphep(id)
                    flash(f"Quản lý {current_user.hoten} từ chối hê duyệt phiếu xin nghỉ phép số {id}  !!!")
                    return redirect(f"/muc7_1_4?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&trangthai={trangthai_filter}")
            else:
                flash(f"{current_user.hoten} không có quyền phê duyệt !!!")
            return redirect(f"/muc7_1_4?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&trangthai={trangthai_filter}")
        except Exception as e:
            flash(f"Lỗi quản lý phê duyệt xin nghỉ phép: {e}")
            return redirect(f"/muc7_1_4?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&trangthai={trangthai_filter}")
        
@app.route("/thuky_kiemtra_xinnghikhongluong", methods=["POST"])
def thukykiemtraxinnghikhongluong():
    if request.method == "POST":
        try:
            mst_filter = request.form["mst_filter"]
            hoten_filter = request.form["hoten_filter"]
            chucvu_filter = request.form["chucvu_filter"]
            chuyen_filter = request.form["chuyen_filter"]
            bophan_filter = request.form["bophan_filter"]
            ngay_filter = request.form["ngay_filter"]
            lydo_filter = request.form["lydo_filter"]
            trangthai_filter = request.form["trangthai_filter"]
            chuyen = request.form["chuyen"]
            mstduyet = current_user.masothe
            kiemtra = request.form["kiemtra"]
            id = request.form["id"]
            mstxinnghikhongluong = request.form["mst_xinnghikhongluong"]
            if mstxinnghikhongluong==mstduyet:
                flash(f"Bạn không thể kiểm tra cho chính mình, vui lòng liên hệ thư ký !!!")
                return redirect(f"/muc7_1_5?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
            if thuky_duoc_phanquyen(mstduyet,chuyen):
                if kiemtra == "Kiểm tra":    
                    thuky_dakiemtra_xinnghikhongluong(id)
                    flash(f"Thư ký {current_user.hoten} đã kiểm tra cho phiếu xin nghỉ không lương số {id} !!!")
                    return redirect(f"/muc7_1_5?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
                else:
                    thuky_tuchoi_xinnghikhongluong(id)
                    flash(f"Thư ký {current_user.hoten} từ chối kiểm tra phiếu xin nghỉ không lương số {id}  !!!")
                    return redirect(f"/muc7_1_5?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
            else:
                flash(f"{current_user.hoten} không có quyền kiểm tra !!!")
            return redirect(f"/muc7_1_5?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
        except Exception as e:
            flash(f"Lỗi thư ký kiểm tra xin nghỉ không lương: {e}")
            return redirect(f"/muc7_1_5?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
        
@app.route("/quanly_pheduyet_xinnghikhongluong", methods=["POST"])
def quanlypheduyetnghikhongluong():
    if request.method == "POST":
        try:
            mst_filter = request.form["mst_filter"]
            hoten_filter = request.form["hoten_filter"]
            chucvu_filter = request.form["chucvu_filter"]
            chuyen_filter = request.form["chuyen_filter"]
            bophan_filter = request.form["bophan_filter"]
            ngay_filter = request.form["ngay_filter"]
            lydo_filter = request.form["lydo_filter"]
            trangthai_filter = request.form["trangthai_filter"]
            chuyen = request.form["chuyen"]
            mstduyet = current_user.masothe
            pheduyet = request.form["pheduyet"]
            id = request.form["id"]
            mstxinnghikhongluong = request.form["mst_xinnghikhongluong"]
            if mstxinnghikhongluong==mstduyet:
                flash(f"Bạn không thể phê duyệt cho chính mình, vui lòng liên hệ thư ký !!!")
                return redirect(f"/muc7_1_5?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
            if quanly_duoc_phanquyen(mstduyet,chuyen):
                if pheduyet == "Phê duyệt":    
                    quanly_pheduyet_xinnghikhongluong(id)
                    flash(f"Quản lý {current_user.hoten} đã phê duyệt cho phiếu xin nghỉ không lương số {id} !!!")
                    return redirect(f"/muc7_1_5?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
                else:
                    quanly_tuchoi_xinnghikhongluong(id)
                    flash(f"Quản lý {current_user.hoten} ttừ chối phê duyệt phiếu xin nghỉ không lương số {id}  !!!")
                    return redirect(f"/muc7_1_5?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
            else:
                flash(f"{current_user.hoten} không có quyền phê duyệt !!!")
            return redirect(f"/muc7_1_5?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
        except Exception as e:
            flash(f"Lỗi quản lý phê duyệt xin nghỉ không lương: {e}")
            return redirect(f"/muc7_1_5?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
  
@app.route("/themdulieukpi", methods=["POST"])
def themdulieukpi():
    try:
        json_data = request.get_json()
        macongty = json_data[0]
        masothe = json_data[1]
        hoten = json_data[2]
        for row in json_data[3:]:
            department = row['department']
            objectives = row['objectives']
            possibleMeasures = row['possibleMeasures']
            unitOfMeasurement = row['unitOfMeasurement']
            measurementSource = row['measurementSource']
            aTarget = row['aTarget']
            bTarget = row['bTarget']
            cTarget = row['cTarget']
            month = row['month']
            year = row['year']
            actualResult = row['actualResult']
            inputkpi(macongty,
                     masothe,
                     hoten,
                     department,
                     objectives,
                     possibleMeasures,
                     unitOfMeasurement,
                     measurementSource,
                     aTarget,
                     bTarget,
                     cTarget,
                     month,
                     year,
                     actualResult)
        
        return {"status": "success"}, 200
    except Exception as e:
        app.logger.info(e)
        return {"status":"failed"}, 400 
      
if __name__ == "__main__":
    app.logger.info("Khoi dong phan mem ...")
    serve(app, host='0.0.0.0', port=81, threads=32)
