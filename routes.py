# -*- encoding: utf-8 -*-

from app import *

##################################
#          MAIN ROUTES           #
##################################

@app.before_request
def run_before_every_request():
    try:
        if current_user.is_authenticated:
            f12 = trang_thai_function_12()
            g.notice={"f12":f12, "db":used_db ,"Tổng":0}
            conn = pyodbc.connect(used_db)
            cursor = conn.cursor()
            row = cursor.execute(f"select count(*) from Phan_quyen_thu_ky where MST_QL='{current_user.masothe}'").fetchone()
            if row[0]>0:
                quanly_soluong_diemdanhbu = cursor.execute(f"""
                    SELECT 
                        COUNT(*) as row_count 
                    FROM 
                        (select distinct Nha_may,Chuyen_to,MST_QL from phan_quyen_thu_ky) a
                    INNER JOIN 
                        Diem_danh_bu 
                    ON
                        Diem_danh_bu.Nha_may= a.Nha_may and Diem_danh_bu.Line=a.Chuyen_to
                    WHERE 
                        Diem_danh_bu.Trang_thai=N'Đã kiểm tra' and a.MST_QL='{current_user.masothe}'""").fetchone()[0]
                quanly_soluong_xinnghiphep = cursor.execute(f"""
                    SELECT 
                        COUNT(*) as row_count 
                    FROM 
                        (select distinct Nha_may,Chuyen_to,MST_QL from phan_quyen_thu_ky) a
                    INNER JOIN 
                        Xin_nghi_phep 
                    ON
                        Xin_nghi_phep.Nha_may= a.Nha_may and Xin_nghi_phep.Line=a.Chuyen_to
                    WHERE 
                        Xin_nghi_phep.Trang_thai=N'Đã kiểm tra' and a.MST_QL='{current_user.masothe}'""").fetchone()[0]
                quanly_soluong_xinnghikhongluong = cursor.execute(f"""
                    SELECT 
                        COUNT(*) as row_count 
                    FROM 
                        (select distinct Nha_may,Chuyen_to,MST_QL from phan_quyen_thu_ky) a
                    INNER JOIN 
                        Xin_nghi_khong_luong 
                    ON
                        Xin_nghi_khong_luong.Nha_may= a.Nha_may and Xin_nghi_khong_luong.Chuyen=a.Chuyen_to
                    WHERE 
                        Xin_nghi_khong_luong.Trang_thai=N'Đã kiểm tra' and a.MST_QL='{current_user.masothe}'""").fetchone()[0]
                
                g.notice["Quản lý"]={"Điểm danh bù":quanly_soluong_diemdanhbu,
                    "Xin nghỉ phép": quanly_soluong_xinnghiphep,
                    "Xin nghỉ không lương": quanly_soluong_xinnghikhongluong,
                    "Số thông báo": quanly_soluong_diemdanhbu + quanly_soluong_xinnghiphep + quanly_soluong_xinnghikhongluong
                    }
                g.notice["Tổng"] = g.notice["Tổng"] + quanly_soluong_diemdanhbu + quanly_soluong_xinnghiphep + quanly_soluong_xinnghikhongluong
            else:
                g.notice["Quản lý"]={}
            row = cursor.execute(f"select count(*) from Phan_quyen_thu_ky where MST='{current_user.masothe}'").fetchone()
            # print(f"Thuky: {row}")
            if row[0]>0:
                cac_chuyen_thuky_quanly = list(x[0] for x in cursor.execute(f"select distinct Chuyen_to from Phan_quyen_thu_ky where MST='{current_user.masothe}'").fetchall())
                query_kiemtra_loithe = f"""
                    SELECT 
                        COUNT(*) as row_count 
                    FROM 
                        (
                            SELECT DISTINCT Nha_may, Chuyen_to, MST
                            FROM Phan_quyen_thu_ky
                        ) as distinct_pqt 
                    INNER JOIN 
                        Danh_sach_loi_the_3
                    ON
                        Danh_sach_loi_the_3.Nha_may = distinct_pqt.Nha_may 
                        AND Danh_sach_loi_the_3.Chuyen_to = distinct_pqt.Chuyen_to
                    WHERE 
                        Danh_sach_loi_the_3.Trang_thai IS NULL 
                        AND distinct_pqt.MST = '{current_user.masothe}'"""
                # print(query_kiemtra_loithe)
                soluong_loithe = cursor.execute(query_kiemtra_loithe).fetchone()[0]    
                # print(f"Loi the: {soluong_loithe}")
                thuky_soluong_diemdanhbu = cursor.execute(f"""
                SELECT 
                    COUNT(*) as row_count 
                FROM 
                    (
                        SELECT DISTINCT Nha_may, Chuyen_to, MST
                        FROM Phan_quyen_thu_ky
                    ) as distinct_pqt 
                INNER JOIN 
                    Diem_danh_bu
                ON
                    Diem_danh_bu.Nha_may = distinct_pqt.Nha_may 
                    AND Diem_danh_bu.Line = distinct_pqt.Chuyen_to
                WHERE 
                    Diem_danh_bu.Trang_thai = N'Chờ kiểm tra' 
                    AND distinct_pqt.MST = '{current_user.masothe}'""").fetchone()[0]
                thuky_soluong_xinnghiphep = cursor.execute(f"""
                    SELECT 
                        COUNT(*) as row_count 
                    FROM 
                        (
                            SELECT DISTINCT Nha_may, Chuyen_to, MST
                            FROM Phan_quyen_thu_ky
                        ) as distinct_pqt 
                    INNER JOIN 
                        Xin_nghi_phep
                    ON
                        Xin_nghi_phep.Nha_may = distinct_pqt.Nha_may 
                        AND Xin_nghi_phep.Line = distinct_pqt.Chuyen_to
                    WHERE 
                        Xin_nghi_phep.Trang_thai = N'Chờ kiểm tra' 
                        AND distinct_pqt.MST = '{current_user.masothe}'""").fetchone()[0]
                thuky_soluong_xinnghikhongluong = cursor.execute(f"""
                    SELECT 
                        COUNT(*) as row_count 
                    FROM 
                        (
                            SELECT DISTINCT Nha_may, Chuyen_to, MST
                            FROM Phan_quyen_thu_ky
                        ) as distinct_pqt 
                    INNER JOIN 
                        Xin_nghi_khong_luong
                    ON
                        Xin_nghi_khong_luong.Nha_may = distinct_pqt.Nha_may 
                        AND Xin_nghi_khong_luong.Chuyen = distinct_pqt.Chuyen_to
                    WHERE 
                        Xin_nghi_khong_luong.Trang_thai = N'Chờ kiểm tra' 
                        AND distinct_pqt.MST = '{current_user.masothe}'""").fetchone()[0]
                g.notice["Thư ký"]={"Danh sách lỗi thẻ":soluong_loithe,
                                    "Điểm danh bù":thuky_soluong_diemdanhbu,
                                    "Xin nghỉ phép": thuky_soluong_xinnghiphep,
                                    "Xin nghỉ không lương": thuky_soluong_xinnghikhongluong,
                                    "Line":cac_chuyen_thuky_quanly[0] if len(cac_chuyen_thuky_quanly)==1 else "",
                                    "Số thông báo":soluong_loithe + thuky_soluong_diemdanhbu + thuky_soluong_xinnghiphep + thuky_soluong_xinnghikhongluong}
                g.notice["Tổng"] = g.notice["Tổng"] + soluong_loithe + thuky_soluong_diemdanhbu + thuky_soluong_xinnghiphep + thuky_soluong_xinnghikhongluong
            else:
                g.notice["Thư ký"]={}
            
            so_don_diemdanhbu_chuakiemtra = cursor.execute(f"""select count(*) from Diem_danh_bu 
                                                           where MST='{current_user.masothe}' and Trang_thai=N'Chờ kiểm tra' and Nha_may= '{current_user.macongty}'""").fetchone()[0]
            so_don_diemdanhbu_dakiemtra = cursor.execute(f"""select count(*) from Diem_danh_bu 
                                                           where MST='{current_user.masothe}' and Trang_thai=N'Đã kiểm tra' and Nha_may= '{current_user.macongty}'""").fetchone()[0]
            so_don_diemdanhbu_dapheduyet = cursor.execute(f"""select count(*) from Diem_danh_bu 
                                                           where MST='{current_user.masothe}' and Trang_thai=N'Đã phê duyệt' and Nha_may= '{current_user.macongty}'""").fetchone()[0]
            so_don_diemdanhbu_bituchoi = cursor.execute(f"""select count(*) from Diem_danh_bu 
                                                           where MST='{current_user.masothe}' and Trang_thai LIKE N'Bị từ chối%' and Nha_may= '{current_user.macongty}'""").fetchone()[0]           
            so_don_diemdanhbu = cursor.execute(f"""select count(*) from Diem_danh_bu 
                                                           where MST='{current_user.masothe}' and Nha_may= '{current_user.macongty}'""").fetchone()[0]

            so_don_xinnghiphep_chuakiemtra = cursor.execute(f"""select count(*) from Xin_nghi_phep 
                                                           where MST='{current_user.masothe}' and Trang_thai=N'Chờ kiểm tra' and Nha_may= '{current_user.macongty}'""").fetchone()[0]
            so_don_xinnghiphep_dakiemtra = cursor.execute(f"""select count(*) from Xin_nghi_phep 
                                                           where MST='{current_user.masothe}' and Trang_thai=N'Đã kiểm tra' and Nha_may= '{current_user.macongty}'""").fetchone()[0]
            so_don_xinnghiphep_dapheduyet = cursor.execute(f"""select count(*) from Xin_nghi_phep 
                                                           where MST='{current_user.masothe}' and Trang_thai=N'Đã phê duyệt' and Nha_may= '{current_user.macongty}'""").fetchone()[0]
            so_don_xinnghiphep_bituchoi = cursor.execute(f"""select count(*) from Xin_nghi_phep 
                                                           where MST='{current_user.masothe}' and Trang_thai LIKE N'Bị từ chối%' and Nha_may= '{current_user.macongty}'""").fetchone()[0]
            so_don_xinnghiphep = cursor.execute(f"""select count(*) from Xin_nghi_phep 
                                                           where MST='{current_user.masothe}' and Nha_may= '{current_user.macongty}'""").fetchone()[0]
            
            so_don_xinnghikhongluong_chuakiemtra = cursor.execute(f"""select count(*) from Xin_nghi_khong_luong 
                                                           where MST='{current_user.masothe}' and Trang_thai=N'Chờ kiểm tra' and Nha_may= '{current_user.macongty}'""").fetchone()[0]
            so_don_xinnghikhongluong_dakiemtra = cursor.execute(f"""select count(*) from Xin_nghi_khong_luong 
                                                           where MST='{current_user.masothe}' and Trang_thai=N'Đã kiểm tra' and Nha_may= '{current_user.macongty}'""").fetchone()[0]
            so_don_xinnghikhongluong_dapheduyet = cursor.execute(f"""select count(*) from Xin_nghi_khong_luong 
                                                           where MST='{current_user.masothe}' and Trang_thai=N'Đã phê duyệt' and Nha_may= '{current_user.macongty}'""").fetchone()[0]
            so_don_xinnghikhongluong_bituchoi = cursor.execute(f"""select count(*) from Xin_nghi_khong_luong 
                                                           where MST='{current_user.masothe}' and Trang_thai LIKE N'Bị từ chối%' and Nha_may= '{current_user.macongty}'""").fetchone()[0]
            so_don_xinnghikhongluong = cursor.execute(f"""select count(*) from Xin_nghi_khong_luong 
                                                           where MST='{current_user.masothe}' and Nha_may= '{current_user.macongty}'""").fetchone()[0]
            
            so_don = so_don_diemdanhbu + so_don_xinnghiphep + so_don_xinnghikhongluong
            so_lan_loi_cham_cong = cursor.execute(f"""select count(*) from Danh_sach_loi_the_3 
                                                           where MST='{current_user.masothe}' and Nha_may= '{current_user.macongty}'""").fetchone()[0]           
            
            g.notice["personal"]={"Điểm danh bù":{
                                                    "Chưa kiểm tra":so_don_diemdanhbu_chuakiemtra,
                                                    "Đã kiểm tra": so_don_diemdanhbu_dakiemtra,
                                                    "Đã phê duyệt": so_don_diemdanhbu_dapheduyet,
                                                    "Bị từ chối": so_don_diemdanhbu_bituchoi,
                                                    "Tổng": so_don_diemdanhbu
                                                },
                                  "Xin nghỉ phép":{
                                                    "Chưa kiểm tra":so_don_xinnghiphep_chuakiemtra,
                                                    "Đã kiểm tra": so_don_xinnghiphep_dakiemtra,
                                                    "Đã phê duyệt": so_don_xinnghiphep_dapheduyet,
                                                    "Tổng": so_don_xinnghiphep,
                                                    "Bị từ chối": so_don_xinnghiphep_bituchoi,
                                                },
                                  "Xin nghỉ không lương":{
                                                    "Chưa kiểm tra":so_don_xinnghikhongluong_chuakiemtra,
                                                    "Đã kiểm tra": so_don_xinnghikhongluong_dakiemtra,
                                                    "Đã phê duyệt": so_don_xinnghikhongluong_dapheduyet,
                                                    "Tổng": so_don_xinnghikhongluong,
                                                    "Bị từ chối": so_don_xinnghikhongluong_bituchoi,
                                                },
                                  "Xin nghỉ khác":{
                                                    "Chưa kiểm tra":0,
                                                    "Đã kiểm tra": 0,
                                                    "Đã phê duyệt": 0,
                                                    "Tổng": 0,
                                                    "Bị từ chối": 0,
                                                },
                                  "Tổng":so_don,
                                  "Lỗi chấm công": so_lan_loi_cham_cong
                                                  }
            conn.close()
    except Exception as e:  
        print(e)
        f12 = trang_thai_function_12()    
        g.notice={"f12":f12,"db":used_db }
    # print(g.notice)
    
@app.context_processor
def inject_notice():
    return dict(notice=getattr(g, 'notice', {}),personal = getattr(g, 'personal', {}))

@app.route('/unauthorized')
def unauthorized():
    return render_template_string("<h1>Bạn không thể vào mục này, vui lòng chọn mục khác!!!</h1><h3>Ấn vào <a href='/'>đây</a> để quay lại trang chủ</h3>")

@app.errorhandler(404)
def page_not_found(e):
    return render_template('blank.html'), 404
 
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        try:
            macongty = request.form['macongty']
            masothe = request.form['masothe']
            matkhau = request.form['matkhau']
            user = Nhanvien.query.filter_by(masothe=masothe, macongty=macongty).first()    
            if user and user.matkhau == matkhau:
                if login_user(user):    
                    print(f"Nguoi dung {masothe} o {macongty} vua  dang nhap thanh cong !!!")
                    return redirect(url_for('home'))
            return redirect(url_for("login"))
        except Exception as e:
            print(f'Nguoi dung {masothe} o {macongty} dang nhap that bai: {e} !!!')
            return redirect(url_for("login"))
    return render_template("login.html")

@app.route("/logout", methods=["POST"])
def logout():
    try:
        print(f"Nguoi dung {current_user.masothe} o {current_user.macongty} vua  dang xuat !!!")
        logout_user()
    except Exception as e:
        print(f'Không thế đăng xuất {e} !!!')
    return redirect("/")

@app.route("/doimatkhau", methods=['POST'])
def doimatkhau():
    macongty = request.form.get("macongty")
    masothe = request.form.get("masothe_doi")
    matkhaumoi = request.form.get("matkhaumoi")
    try:
        if doimatkhautaikhoan(macongty,masothe,matkhaumoi):
            print("Đổi mật khẩu thành công")
    except Exception as e:
        print(f"Đổi mật khẩu không thành công: {e}")
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
        chuyen = request.args.get("Chuyền")
        users = laydanhsachuser(mst, hoten, sdt, cccd, gioitinh, vaotungay, vaodenngay, nghitungay, nghidenngay, phongban, trangthai, hccategory, chucvu, ghichu, chuyen)   
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
        print(f"Xin chào {current_user.hoten} !!!")
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
        chuyen = request.form.get("Chuyền")
        users = laydanhsachuser(mst, hoten, sdt, cccd, gioitinh, vaotungay, vaodenngay, nghitungay, nghidenngay, phongban, trangthai, hccategory, chucvu, ghichu, chuyen)      
        df = pd.DataFrame(users)

        df["Ngày sinh"] = to_datetime(df['Ngày sinh'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày cấp CCCD"] = to_datetime(df['Ngày cấp CCCD'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày ký HĐ"] = to_datetime(df['Ngày ký HĐ'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày vào"] = to_datetime(df['Ngày vào'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày nghỉ"] = to_datetime(df['Ngày nghỉ'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày hết hạn"] = to_datetime(df['Ngày hết hạn'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày vào nối thâm niên"] = to_datetime(df['Ngày vào nối thâm niên'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày sinh con 1"] = to_datetime(df['Ngày sinh con 1'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày sinh con 2"] = to_datetime(df['Ngày sinh con 2'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày sinh con 3"] = to_datetime(df['Ngày sinh con 3'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày sinh con 4"] = to_datetime(df['Ngày sinh con 4'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày sinh con 5"] = to_datetime(df['Ngày sinh con 5'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày kí HĐ Thử việc"] = to_datetime(df['Ngày kí HĐ Thử việc'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày hết hạn HĐ Thử việc"] = to_datetime(df['Ngày hết hạn HĐ Thử việc'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày kí HĐ xác định thời hạn lần 1"] = to_datetime(df['Ngày kí HĐ xác định thời hạn lần 1'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày hết hạn HĐ xác định thời hạn lần 1"] = to_datetime(df['Ngày hết hạn HĐ xác định thời hạn lần 1'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày kí HĐ xác định thời hạn lần 2"] = to_datetime(df['Ngày kí HĐ xác định thời hạn lần 2'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày hết hạn HĐ xác định thời hạn lần 2"] = to_datetime(df['Ngày hết hạn HĐ xác định thời hạn lần 2'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        df["Ngày kí HĐ không thời hạn"] = to_datetime(df['Ngày kí HĐ không thời hạn'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Adjust column width and format the header row
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        # Style the header row
        header_fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")

        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font

        # Adjust column widths
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        # Save the modified workbook to the output BytesIO object
        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        
        # Generate the timestamp for the filename
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        
        # Return the file to the client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=danhsach_nhanvien_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response

@app.route("/muc2_1", methods=["GET","POST"])
@login_required
@roles_required('hr','tnc','sa','gd','td','tbp')
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
        cccd = request.form.get("cccd")
        if capnhatthongtinungvien(id,
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
            print("Cập nhật thông tin ứng viên thành công !!!")
        else:
            print("Cập nhật thông tin ứng viên thất bại !!!")
        return redirect(f"muc2_1?sdt={sdt}")

@app.route("/muc2_2", methods=["GET","POST"])
@login_required
@roles_required('tbp','gd','sa','td')
def dangkytuyendung():
    if request.method == "GET":
        maso = current_user.macongty[-1]
        danhsach = laydanhsachyeucautuyendung(maso)
        cactrangthaithuchien = ["Chưa tuyển","Đã đăng tuyển","Chờ phỏng vấn","Đã tuyển"]
        cacvitri = lay_cac_vitri_trong_phong(current_user.phongban)
        return render_template("2_2.html", 
                               page="2.2 Yêu cầu tuyển dụng",
                               danhsach=danhsach,
                               cacvitri = cacvitri,
                               cactrangthaithuchien=cactrangthaithuchien
                               )
    
    elif request.method == "POST":
        try:
            bophan = current_user.phongban
            vitri = request.form.get("vitri")
            vitrien = request.form.get("vitrien")
            capbac = request.form.get("capbac")
            bacluongtu = request.form.get("bacluongtu")
            bacluongden = request.form.get("bacluongden")
            bacluong = f"{bacluongtu.split(",")[0]}-{bacluongden.split(",")[0]}"
            soluong = request.form.get("soluong")
            mota = os.path.join(FOLDER_JD, f"{vitrien}.pdf")
            thoigiandukien = request.form.get("thoigiandukien")
            phanloai = request.form.get("phanloai")
            khoangluong = f"{bacluongtu.split(",")[1]}-{bacluongden.split(",")[1]}"
            if themyeucautuyendungmoi(bophan,vitri,soluong,mota,thoigiandukien,phanloai,khoangluong,capbac,bacluong):
                print("Thêm yêu cầu tuyển dụng mới thành công !!!")
            else:
                print("Thêm yêu cầu tuyển dụng mới thất bại !!!")
        except Exception as e:
            print(f"Thêm yêu cầu tuyển dụng mới thất bại ({e})!!!")
        return redirect("muc2_2")
    
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
        try:
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
            tencon1 = f"N'{request.form.get("tenconnho1")}'" if request.form.get("tenconnho1") else 'NULL'
            ngaysinhcon1 = f"'{request.form.get("ngaysinhcon1")}'" if request.form.get("ngaysinhcon1") else 'NULL'
            tencon2 = f"N'{request.form.get("tenconnho2")}'" if request.form.get("tenconnho2") else 'NULL'
            ngaysinhcon2 = f"'{request.form.get("ngaysinhcon2")}'" if request.form.get("ngaysinhcon2") else 'NULL'
            tencon3 = f"N'{request.form.get("tenconnho3")}'" if request.form.get("tenconnho3") else 'NULL'
            ngaysinhcon3 = f"'{request.form.get("ngaysinhcon3")}'" if request.form.get("ngaysinhcon3") else 'NULL'
            tencon4 = f"N'{request.form.get("tenconnho4")}'" if request.form.get("tenconnho4") else 'NULL'
            ngaysinhcon4 = f"'{request.form.get("ngaysinhcon4")}'" if request.form.get("ngaysinhcon4") else 'NULL'
            tencon5 = f"N'{request.form.get("tenconnho5")}'" if request.form.get("tenconnho5") else 'NULL'
            ngaysinhcon5 = f"'{request.form.get("ngaysinhcon5")}'" if request.form.get("ngaysinhcon5") else 'NULL'
            jobdetailvn = f"N'{request.form.get("vitri")}'"
            line = f"'{request.form.get("line")}'"
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
            sdtnguoithan = f"N'{request.form.get("sdtnguoithan")}'" if request.form.get("sdtnguoithan") else 'NULL'
            luongcoban = 'NULL'
            tongphucap = 'NULL'
            kieuhopdong = 'NULL'

            ngaybatdauthuviec = 'NULL'
            ngayvao = 'GETDATE()'
            ngayketthuc = 'NULL'
            ngayketthucthuviec = 'NULL'
            ngaybatdauhdcthl1 = "NULL"
            ngayketthuchdcthl1 = "NULL"
            ngaybatdauhdcthl2 = "NULL"
            ngayketthuchdcthl2 = "NULL"
            ngaybatdauhdvth = "NULL"

            nhanvienmoi = f"({masothe},{thechamcong},{hoten},{dienthoai},{ngaysinh},{gioitinh},{cccd},{ngaycapcccd},N'Cục cảnh sát',{cmt},{thuongtru},{thonxom},{phuongxa},{quanhuyen},{tinhthanhpho},{dantoc},{quoctich},{tongiao},{hocvan},{noisinh},{tamtru},{sobhxh},{masothue},{nganhang},{sotaikhoan},{connho},{tencon1},{ngaysinhcon1},{tencon2},{ngaysinhcon2},{tencon3},{ngaysinhcon3},{tencon4},{ngaysinhcon4},{tencon5},{ngaysinhcon5},{anh},{nguoithan}, {sdtnguoithan},{kieuhopdong},{ngayvao},{ngayketthuc},{jobdetailvn},{hccategory},{gradecode},{factory},{department},{chucvu},{sectioncode},{sectiondescription},{line},{employeetype},{jobdetailen},{positioncode},{positioncodedescription},{luongcoban},N'Không',{tongphucap},{ngayvao},NULL,N'Đang làm việc',{ngayvao},'1',{ngaybatdauthuviec},{ngayketthucthuviec},{ngaybatdauhdcthl1},{ngayketthuchdcthl1},{ngaybatdauhdcthl2},{ngayketthuchdcthl2},{ngaybatdauhdvth},'N', '', GETDATE())"             
            if themnhanvienmoi(nhanvienmoi):
                print("Thêm lao động mới thành công !!!")
                themdoicamoi(request.form.get("masothe"),laycatheochuyen(request.form.get("line")),laycatheochuyen(request.form.get("line")),datetime.now().date(),datetime(2054,12,31))
                flash("Tạo ca mặc định cho người mới thành công !!!")  
                themtaikhoanmoi(
                    int(request.form.get("masothe")),
                    request.form.get("hoten"),
                    request.form.get("department"),
                    request.form.get("gradecode")
                )              
            else:
                masothe = int(laymasothemoi())+1
                cacvitri= laycacvitri()
                cacto = laycacto()
                cacca = laycacca()
                print("Thêm lao động mới thất bại !!!")
        except Exception as e:
            print(f"Them lao dong moi that bai: {e} !!!")
        finally:
            return redirect("/muc3_1")
        
@app.route("/muc3_2", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def thaydoithongtinlaodong():
    
    if request.method == "GET":
        return render_template("3_2.html", page="3.2 Thay đổi thông tin người lao động")
    else:
        try:
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
            thuongtru = request.form.get("thuongtru")
            tamtru = request.form.get("tamtru")
            noisinh = request.form.get("noisinh")
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
            nguoithan = request.form.get("nguoithan")
            sdtnguoithan = request.form.get("sdtnguoithan")
            
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
            
            if thuongtru:
                query += f"Dia_chi_thuong_tru = N'{thuongtru}',"
            else:
                query += f"Dia_chi_thuong_tru = NULL,"
                
            if tamtru:
                query += f"Dia_chi_tam_tru = N'{tamtru}',"
            else:
                query += f"Dia_chi_tam_tru = NULL,"
                
            if noisinh:
                query += f"Noi_sinh = N'{noisinh}',"
            else:
                query += f"Noi_sinh = NULL,"
            
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
            
            if nguoithan:
                query += f"Nguoi_than = N'{nguoithan}',"
            else:
                query += f"Nguoi_than = NULL," 
                
            if sdtnguoithan:
                query += f"Sdt_Nguoithan = '{sdtnguoithan}',"
            else:
                query += f"Sdt_Nguoithan = NULL," 
            
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
                query += f"Chuc_vu = N'{chucvu}',"
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
                query += f"Trang_thai_lam_viec = NULL,"
            query = query[:-1] + f" WHERE MST = '{mst}' AND Factory='{current_user.macongty}'"
            conn = pyodbc.connect(used_db)
            cursor = conn.cursor()
            if current_user.macongty == "NT2":
                if not current_user.masothe == "4091":
                    print("Bạn không có quyền thay đổi thông tin người lao động !!!")
            cursor.execute(query)
            conn.commit()
            conn.close()
            print("Cập nhật thông tin người lao động thành công !!!")
        except Exception as e:
            print(e)
            print(f"Cập nhật thông tin người lao động thất bại: {e} !!!")
        return redirect("/muc3_2")
    
@app.route("/muc3_3", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def inhopdonglaodong():
    try:
        if request.method == "GET":
            mst = request.args.get("mst")
            if mst:
                danhsach = laydanhsach_hopdong_theomst(mst)
            else:
                danhsach = []
            return render_template("3_3.html", page="3.3 Quản lý hợp đồng lao động",danhsach=danhsach)
        elif request.method == "POST":
            nhamay = current_user.macongty
            mst = request.form.get("form_manhanvien")
            hoten = request.form.get("form_hovaten")
            gioitinh = request.form.get("form_gioitinh")
            ngaysinh =  request.form.get("form_ngaysinh")
            thuongtru = request.form.get("form_thuongtru")
            tamtru = request.form.get("form_tamtru")
            cccd = request.form.get("form_cccd")
            ngaycapcccd = request.form.get("form_ngaycapcccd")
            capbac =  request.form.get("gradecode")
            loaihopdong = request.form.get("form_loaihopdong")
            chucdanh = request.form.get("chucdanh")
            phongban = request.form.get("department")
            chuyen = request.form.get("chuyen")
            luongcoban = request.form.get("luongcoban")
            phucap = request.form.get("phucap")
            ngaybatdau = request.form.get("form_ngaykyhopdong")
            ngayketthuc = request.form.get("form_ngayhethanhopdong")
            vitrien = request.form.get("vitrien")
            employeetype = request.form.get("employeetype")
            positioncode = request.form.get("positioncode")
            postitioncodedescription = request.form.get("postitioncodedescription")
            hccategory = request.form.get("hccategory")
            sectioncode = request.form.get("sectioncode")
            sectiondescription = request.form.get("sectiondescription")
            if themhopdongmoi(nhamay,mst,hoten,gioitinh,ngaysinh,thuongtru,tamtru,cccd,ngaycapcccd,capbac,loaihopdong,chucdanh,phongban,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc):
                print("Thêm hợp đồng thành công !!!")
                capnhatthongtinhopdong(nhamay,mst,loaihopdong,chucdanh,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc,vitrien,employeetype,positioncode,postitioncodedescription,hccategory,sectioncode,sectiondescription)
            else:
                print("Thêm hợp đồng thất bại")
            return redirect("/muc3_3")
    except:
        return redirect("/muc3_3")
    
@app.route("/muc3_4", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
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
        print("Tải file thành công !!!")
        return send_file(os.path.join(FOLDER_XUAT, f"saphethan_{thoigian}.xlsx"), as_attachment=True)

@app.route("/muc5_1_1", methods=["GET","POST"])
@login_required
@roles_required('sa','tbp','gd')
def nhapkpi():
    if request.method == "GET":
        danhsach = laydanhsachkpichuaduyet(current_user.masothe,current_user.macongty)
        return render_template("5_1_1.html",page="Upload KPI",danhsach=danhsach)
    if request.method == "POST":
        try:
            file = request.files['file']
            if file:
                ngaylam = datetime.now().strftime("%d%m%Y%H%M%S")
                filepath = os.path.join(FOLDER_NHAP, f"kpi_{current_user.masothe}_{ngaylam}.xlsx")
                file.save(filepath)
                data = pd.read_excel(filepath).to_dict(orient="records")
                delete_kpidata(current_user.masothe,current_user.macongty)
                for row in data[1:]:
                    values=[]
                    for x in row.items():
                        values.append(str(x[1]).replace("'","") if x[1] else "")
                    if not insert_kpidata(current_user.masothe,current_user.macongty,values):
                        print("Upload new KPI failed: Cannot insert data !!!")
                        return redirect("/muc5_1_1")
                guimailthongbaodaguikpi(current_user.macongty,current_user.masothe,current_user.hoten)
                print("Upload new KPI successfully !!!")
            else:
                print("Upload new KPI failed: Cannot found data !!!")
        except Exception as e:
            print(f"Upload new KPI failed {e} !!!")
            print("Upload new KPI failed !!!")
        return redirect("/muc5_1_1")

@app.route("/muc5_1_2", methods=["GET","POST"])
@login_required
@roles_required('sa','gd')
def duyetkpi():
    if request.method == "GET":
        congty = request.args.get("company")
        pic = request.args.get("pic")
        mst = pic.split("_")[0] if pic else None
        danhsachquanly = laydanhsachquanly(congty)
        danhsach = laydanhsachkpichuaduyet(mst,congty)
        return render_template("5_1_2.html",page="Approve KPI",danhsach=danhsach,danhsachquanly=danhsachquanly)
    if request.method == "POST":
        congty = request.form.get("company")
        pic = request.args.get("pic")
        mst = pic.split("_")[0] if pic else None
        hoten = pic.split("_")[1] if pic else None
        email = layemailquanly(congty,mst)
        pheduyet = request.form.get("pheduyet")
        if pheduyet == "co":
            pheduyetkpi(mst,congty)
            guimailthongbaodapheduyetkpi(congty,mst,hoten,email)
        else:
            tuchoikpi(mst,congty)
            guimailthongbaodatuchoikpi(congty,mst,hoten,email)
        return redirect(f"/muc5_1_2?company={congty}&mst={mst}")

@app.route("/muc5_1_3_1", methods=["GET","POST"])
@login_required
@roles_required('sa','gd','tbp')
def baocaocanam():
    if request.method == "GET":
        congty = request.args.get("company")
        mst = request.args.get("mst")
        danhsachquanly = laydanhsachquanly(congty)
        danhsach = laydanhsachkpidaduyet(mst,congty)
        return render_template("5_1_3_1.html",page="Performance Report All Year",danhsach=danhsach,danhsachquanly=danhsachquanly)

@app.route("/muc5_1_3_2", methods=["GET","POST"])
@login_required
@roles_required('sa','gd','tbp')
def baocaoytd():
    congty = request.args.get("company")
    mst = request.args.get("mst")
    danhsachquanly = laydanhsachquanly(congty)
    danhsach = laydanhsachkpidaduyet(mst,congty)
    return render_template("5_1_3_2.html",page="Performance Report Year to date",danhsach=danhsach,danhsachquanly=danhsachquanly)


    
@app.route("/muc6_1", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def dieuchuyen():
    try:
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
                    print("Điều chuyển thành công !!!")
                except Exception as e:
                    print(e)
                    print("Điều chuyển thất bại !!!")
                    return redirect(f"/muc6_1")
                
            elif loaidieuchuyen == "Nghỉ việc":
                try:
                    dichuyennghiviec(mst,
                        vitricu,
                        chuyencu,
                        gradecodecu,
                        hccategorycu,
                        ngaydieuchuyen,
                        ghichu
                                )
                    print("Điều chuyển thành công !!!")
                except Exception as e:
                    print(e)
                    print("Điều chuyển thất bại !!!")
                    return redirect(f"/muc6_1")
            elif loaidieuchuyen=="Nghỉ thai sản":
                try:
                    dichuyennghithaisan(mst,
                                vitricu,
                                chuyencu,
                                gradecodecu,
                                hccategorycu,
                                ngaydieuchuyen
                                )
                    print("Điều chuyển thành công !!!")
                except Exception as e:
                    print(e)
                    print("Điều chuyển thất bại !!!")
                    return redirect(f"/muc6_1")
            elif loaidieuchuyen=="Thai sản đi làm lại":
                try:
                    dichuyenthaisandilamlai(mst,
                            vitricu,
                            chuyencu,
                            gradecodecu,
                            hccategorycu,
                            ngaydieuchuyen
                            )
                    print("Điều chuyển thành công !!!")
                except Exception as e:
                    print(e)
                    print("Điều chuyển thất bại !!!")
                    return redirect(f"/muc6_1")
            return redirect(f"/muc6_1")
        elif request.method == "GET":
            cacvitri= laycacvitri()
            return render_template("6_1.html",
                            cacvitri=cacvitri,
                            page="6.1 Điều chuyển chức vụ, bộ phận")
    except Exception as e:
        print(e)
        cacvitri= laycacvitri()
        return render_template("6_1.html",
                            cacvitri=cacvitri,
                            page="6.1 Điều chuyển chức vụ, bộ phận")
    
@app.route("/muc6_2", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def lichsudieuchuyen():
    
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
        return render_template("6_2.html", page="6.2 Lịch sử điều chuyển",
                               danhsach=paginated_rows, 
                               pagination=pagination,
                               mst=mst, 
                               count=count)
    if request.method == "POST":
        mst = request.args.get("mst")
        hoten = request.args.get("hoten")
        ngay = request.args.get("ngay")
        kieudieuchuyen = request.args.get("kieudieuchuyen")
        data = laylichsucongtac(mst,hoten,ngay,kieudieuchuyen)
        df = DataFrame(data)
        df["Ngày thực hiện"] = to_datetime(df['Ngày thực hiện'], errors='coerce', dayfirst=True)
        df["Ngày chính thức"] = to_datetime(df['Ngày chính thức'], errors='coerce', dayfirst=True)
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active
        # Create a date format for short date
        date_format = NamedStyle(name="short_date", number_format="DD/MM/YYYY")
        if "short_date" not in workbook.named_styles:
            workbook.add_named_style(date_format)
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    # Apply the date format to column L (assuming 'Ngày thực hiện' is in column 'L')
                    if cell.column_letter == 'C' and cell.value is not None:
                        cell.number_format = 'DD/MM/YYYY'
                    if cell.column_letter == 'K' and cell.value is not None:
                        cell.number_format = 'DD/MM/YYYY'
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width
        
        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=dieuchuyen_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response

@app.route("/muc6_3", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def lichsucongviec():
    if request.method == "GET":
        mst = request.args.get("mst")
        chuyen = request.args.get("chuyen")
        bophan = request.args.get("bophan")
        rows = laylichsucongviec(mst,chuyen,bophan)
        count = len(rows)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        total = len(rows)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = rows[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("6_3.html", page="6.3 Lịch sử công việc",
                               danhsach=paginated_rows, 
                               pagination=pagination,
                               count=count)
        
    elif request.method == "POST":
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        rows = laylichsucongviec(mst,chuyen,bophan)
        data = [{
            "Mã công ty": row[0],
            "Mã số thẻ": row[1],
            "Họ tên": row[2],
            "Chuyền": row[3],
            "Bộ phận": row[4],
            "Chức danh": row[5],
            "Cấp bậc": row[6],
            "HC category": row[11],
            "Trạng thái": row[7],
            "Ngày bắt đầu": row[8],
            "Ngày kết thúc": row[9]
        } for row in rows]
        df = DataFrame(data)
        df["Ngày bắt đầu"] = to_datetime(df['Ngày bắt đầu'], errors='coerce', dayfirst=True)
        df["Ngày kết thúc"] = to_datetime(df['Ngày kết thúc'], errors='coerce', dayfirst=True)
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active
        # Create a date format for short date
        date_format = NamedStyle(name="short_date", number_format="DD/MM/YYYY")
        if "short_date" not in workbook.named_styles:
            workbook.add_named_style(date_format)
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    # Apply the date format to column L (assuming 'Ngày thực hiện' is in column 'L')
                    if cell.column_letter == 'J' and cell.value is not None:
                        cell.number_format = 'DD/MM/YYYY'
                    if cell.column_letter == 'K' and cell.value is not None:
                        cell.number_format = 'DD/MM/YYYY'
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width
        
        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=lichsu_congviec_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response
    
@app.route("/muc7_1_1", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def khaibaochamcong():
    if request.method == "GET":
        try:
            mst = request.args.get("mst")
            chuyen = request.args.get("chuyen") 
            phongban = request.args.get("phongban") 
            rows = laydanhsachcahientai(mst,chuyen,phongban)
            count = len(rows)
            current_page = request.args.get(get_page_parameter(), type=int, default=1)
            per_page = 15
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
        except:
            return render_template("7_1_1.html",
                                    page="7.1.1 Đổi ca làm việc",
                                    danhsach=[])
    elif request.method == "POST":
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen") 
        phongban = request.form.get("phongban") 
        rows = laydanhsachcahientai(mst,chuyen,phongban)
        data =[]
        for row in rows:
            data.append({
                "Nhà máy": row[0],
                "Mã số thẻ": row[1],
                "Họ tên": row[2],
                "Chuyền tổ": row[3], 
                "Phòng ban": row[4],
                "Ca": row[5],
                "Đổi từ ngày": row[6],
                "Đổi đến ngày": row[7]
            })
        df = pd.DataFrame(data)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT, f"doica_{thoigian}.xlsx"), index=False)
        print("Tải file thành công !!!")
        return send_file(os.path.join(FOLDER_XUAT, f"doica_{thoigian}.xlsx"), as_attachment=True)
            
@app.route("/muc7_1_2", methods=["GET","POST"])
@login_required
def loichamcong():
    mstthuky = request.args.get("mstthuky")
    mst = request.args.get("mst")
    chuyen = request.args.get("chuyen")
    bophan = request.args.get("bophan")
    ngay = request.args.get("ngay")
    danhsach = laydanhsachloithe(mst,chuyen,bophan,ngay,mstthuky)
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
    mstthuky = request.args.get("mstthuky")
    mstquanly = request.args.get("mstquanly")
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
    danhsach = laydanhsachdiemdanhbu(mst,hoten,chucvu,chuyen,bophan,loaidiemdanh,ngay,lido,trangthai,mstquanly,mstthuky)
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
    mstthuky = request.args.get("mstthuky")
    mstquanly = request.args.get("mstquanly")
    mst = request.args.get("mst")
    hoten = request.args.get("hoten")
    chucvu = request.args.get("chucvu")
    chuyen = request.args.get("chuyen")
    bophan = request.args.get("bophan")
    ngay = request.args.get("ngaynghi")
    lydo = request.args.get("lydo")
    trangthai = request.args.get("trangthai")
    danhsach = laydanhsachxinnghiphep(mst,hoten,chucvu,chuyen,bophan,ngay,lydo,trangthai,mstquanly,mstthuky)
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

@app.route("/muc7_1_5", methods=["GET","POST"])
@login_required
def xinnghikhongluong():
    if request.method == 'GET':
        mstthuky = request.args.get("mstthuky")
        mstquanly = request.args.get("mstquanly")
        mst = request.args.get("mst")
        hoten = request.args.get("hoten")
        chucvu = request.args.get("chucvu")
        chuyen = request.args.get("chuyen")
        bophan = request.args.get("bophan")
        ngay = request.args.get("ngaynghi")
        lydo = request.args.get("lydo")
        trangthai = request.args.get("trangthai")
        danhsach = laydanhsachxinnghikhongluong(mst,hoten,chucvu,chuyen,bophan,ngay,lydo,trangthai,mstquanly,mstthuky)
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
        mstquanly = request.form.get("mstquanly")
        mst = request.form.get("mst")
        hoten = request.form.get("hoten")
        chucvu = request.form.get("chucvu")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        ngay = request.form.get("ngaynghi")
        lydo = request.form.get("lydo")
        trangthai = request.form.get("trangthai")
        danhsach = laydanhsachxinnghikhongluong(mst,hoten,chucvu,chuyen,bophan,ngay,lydo,trangthai,mstquanly)
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
        print("Tải file thành công !!!")
        return send_file(os.path.join(FOLDER_XUAT, f"xinnghikhongluong_{thoigian}.xlsx"), as_attachment=True)
        
@app.route("/muc7_1_6", methods=["GET","POST"])
@login_required
def danhsachxinnghikhac():
    if request.method == "GET":
        mst = request.args.get("mst")
        chuyen = request.args.get("chuyen")
        bophan = request.args.get("bophan")
        ngaynghi = request.args.get("ngaynghi")
        loainghi = request.args.get("loainghi")
        trangthai = request.args.get("trangthai")
        nhangiayto = request.args.get("nhangiayto")
        danhsach = laydanhsachxinnghikhac(mst,chuyen,bophan,ngaynghi,loainghi,trangthai,nhangiayto)
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
                                count=count)
    elif request.method == "POST":
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        ngaynghi = request.form.get("ngaynghi")
        loainghi = request.form.get("loainghi")
        danhsach = laydanhsachxinnghikhac(mst,chuyen,bophan,ngaynghi,loainghi)
        data = [{
            "Nhà máy": row[0],
            "Mã số thẻ": row[1],
            "Họ tên": row[8],
            "Bộ phận": row[10],
            "Chuyền": row[9],
            "Ngày nghỉ": row[2],
            "Tổng số phút": row[3],
            "Loại nghỉ": row[4],
            "Trạng thái": row[5],
            "Nhận giấy tờ": row[6],            
        } for row in danhsach] 
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='coerce')
        df["Ngày nghỉ"] = to_datetime(df['Ngày nghỉ'], errors='coerce')
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=xinnghikhac_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response
        
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
        return render_template("7_1_7.html", 
                               page="7.1.7 Đăng ký tăng ca",
                               danhsach=paginated_rows,
                               pagination=pagination,
                               count=count
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
        print("Tải file thành công !!!")
        return send_file(os.path.join(FOLDER_XUAT,f"tangca_{thoigian}.xlsx"), as_attachment=True)

@app.route("/muc7_1_8", methods=["GET","POST"])
@login_required
def chamcongtudong():
    
    mst = request.args.get("mst")
    chuyen = request.args.get("chuyen")
    phongban = request.args.get("phongban")
    tungay = request.args.get("tungay")
    denngay = request.args.get("denngay")
    phanloai = request.args.get("phanloai")
    rows = laydanhsachchamcong(mst,chuyen,phongban,tungay,denngay,phanloai)
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
        print("Tải file thành công !!!")
        return send_file(os.path.join(FOLDER_XUAT, f"phepton_{thoigian}.xlsx"), as_attachment=True)
            
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
            if themdanhsachkyluat(mst,hoten,chucvu,bophan,chuyento,ngayvao,ngayvipham,diadiem,ngaylapbienban,noidung,bienphap):
                print("Thêm biên bản kỷ luật thành công !!!")
        except Exception as ex:
            print("Thêm biên bản kỷ luật thất bại !!!")
            print(ex)
        return redirect("/muc9_1") 
    
@app.route("/muc10_1", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def phongvannghiviec():
        
    return render_template("10_1.html", page="10.1 Tổng hợp phỏng vấn nghỉ việc")

@app.route("/muc10_2", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def nhandonnghiviec():
    if request.method == "GET":
        mst = request.args.get("mst")
        hoten = request.args.get("hoten")
        chuyen = request.args.get("chuyen")
        phongban = request.args.get("phongban")
        ngaynopdon = request.args.get("ngaynopdon")
        ngaynghi = request.args.get("ngaynghi")
        sapdenhan = request.args.get("sapdenhan")
        danhsach = laydanhsach_chonghiviec(mst,hoten,chuyen,phongban,ngaynopdon,ngaynghi,sapdenhan)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("10_2.html", 
                            page="10.2 Tổng hợp đơn nghỉ việc",
                            danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        mst = request.form.get("form_manhanvien")
        hoten = request.form.get("form_hovaten")
        chucdanh = request.form.get("form_chucvu")
        chuyen = request.form.get("form_chuyento")
        phongban = request.form.get("form_bophan")
        ngaynopdon = request.form.get("form_ngaynopdon")
        ngaynghi = request.form.get("form_ngaydukiennghi")
        ghichu = request.form.get("form_ghichu")
        if themdonxinnghi(mst,hoten,chucdanh,chuyen,phongban,ngaynopdon,ngaynghi,ghichu):
            print("Thêm đơn xin nghỉ thành công !!!")
        else:
            print("Thêm đơn xin nghỉ thất bại !!!")
        return redirect(f"/muc10_2?mst={mst}")
    
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

                return redirect("/muc10_3")
        except Exception as e:
            print(e)
            return redirect("/muc10_3") 
        
@app.route("/muc12", methods=["GET","POST"])
@login_required
def khong_kiem_xuong():
    try:
        if request.method=="GET":
            return render_template("12.html")
        else:
            return "OK"
    except Exception as e:
        print(e)
        return "NOT OK"

#############################################
#                "OTHER ENDPOINT"           #
#############################################

# @app.route("/thaydoiphanquyen", methods=["POST"])
# def thaydoiphanquyen():
#     if request.method == "POST":
#         userid= request.form["id"]
#         newrole = request.form["newrole"]
#         user = Users.query.filter_by(id=userid).first()
#         if user:
#             user.role = newrole
#             db.session.commit()
#         return redirect("/admin")
    
@app.route("/admin", methods=["GET"])
@login_required
@roles_required('sa')
def admin_page():
    trangthai = trang_thai_function_12()
    return render_template("admin.html",trangthai=trangthai)

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
        print(e)
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

        user = laydanhsachtheomst(mst)

        if user:
            user = user[0]

            if kiemtrathuki(current_user.masothe,user['Line']):
                if insert_tangca(current_user.macongty,
                            mst,
                            user['Họ tên'],
                            user['Chức vụ'],
                            user['Line'],
                            user['Department'],
                            ngaytangca,
                            giotangca):
                    print(f"{current_user.masothe} đã đăng ký tăng ca cho {mst} thành công", "success")
                else:
                    print(f"{current_user.masothe} đã đăng ký tăng ca cho {mst} thất bại", "danger")
                return redirect(f"/muc7_1_6?ngay={ngaytangca}")
            else:
                print(f"{current_user.masothe} không được phép đăng ký tăng ca cho {mst}", "danger")
                return redirect(f"/muc7_1_6")
        else:
            print(f"Không tìm thấy nhân viên có {mst}", "danger")
            return redirect(f"/muc7_1_6")  
    except Exception as e:
        print(f"Đăng ký tăng ca lỗi: {e}")
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
                        print(f"Thư ký {current_user.masothe} {row['Chuyền tổ']} dang ki tang ca cho {row['MST']} {row['Họ tên']} {row['Chức vụ']} {row['Phòng ban']} {row['Ngày đăng ký']} {row['Giờ tăng ca']}")
                        try:
                            if insert_tangca(current_user.macongty,row["MST"],row["Họ tên"],row["Chức vụ"],row["Chuyền tổ"],row["Phòng ban"],row["Ngày đăng ký"],row["Giờ tăng ca"]):
                                print(f"{current_user.masothe} đã đăng ký tăng ca cho {row['MST']} thành công", "success")
                            else:
                                print(f"{current_user.masothe} đã đăng ký tăng ca cho {row['MST']} thất bại", "danger")
                        except Exception as e:
                            print(e)   
                    else:
                        print(f"{current_user.masothe} không được đăng ký tăng ca cho {row['MST']}")            
            return redirect("/muc7_1_6")
        except Exception as e:
            print(f"{current_user.masothe} không được đăng ký tăng ca cho {row['MST']} lỗi: {e}")
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
    mstquanly = request.form.get("mstquanly")
    mst = request.form.get("mst")
    chuyen = request.form.get("chuyen")
    bophan = request.form.get("bophan")
    hoten = request.form.get("hoten")
    chucvu = request.form.get("chucvu")
    ngaydiemdanh = request.form.get("ngay")
    lydo = request.form.get("lydo")
    trangthai = request.form.get("trangthai")
    loaidiemdanh = request.form.get("loaidiemdanh")
    
    rows = laydanhsachdiemdanhbu(mst,hoten,chucvu,chuyen,bophan,loaidiemdanh,ngaydiemdanh,lydo,trangthai,mstquanly)
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
    mstquanly = request.form.get("mstquanly")
    mst = request.form.get("mst")
    hoten = request.form.get("hoten")
    chucvu = request.form.get("chucvu")
    chuyen = request.form.get("chuyen")
    bophan = request.form.get("bophan")
    ngay = request.form.get("ngaynghi")
    lydo = request.form.get("lydo")
    trangthai = request.form.get("trangthai")
    danhsach = laydanhsachxinnghiphep(mst,hoten,chucvu,chuyen,bophan,ngay,lydo,trangthai,mstquanly)
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
    
    df["Ngày sinh con 1"] = to_datetime(df['Ngày sinh con 1'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
    df["Ngày sinh con 2"] = to_datetime(df['Ngày sinh con 2'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
    df["Ngày sinh con 3"] = to_datetime(df['Ngày sinh con 3'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
    df["Ngày sinh con 4"] = to_datetime(df['Ngày sinh con 4'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
    df["Ngày sinh con 5"] = to_datetime(df['Ngày sinh con 5'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
    df["Ngày gửi"] = to_datetime(df['Ngày gửi'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
    df["Ngày cập nhật"] = to_datetime(df['Ngày cập nhật'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
    df["Ngày hẹn đi làm"] = to_datetime(df['Ngày hẹn đi làm'], errors='coerce', dayfirst=True, format="%d/%m/%Y")
    df["Ngày nhận việc"] = to_datetime(df['Ngày nhận việc'], errors='coerce', dayfirst=True, format="%d/%m/%Y") 
      
    output = BytesIO()
    with ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    # Điều chỉnh độ rộng cột
    output.seek(0)
    workbook = openpyxl.load_workbook(output)
    sheet = workbook.active

    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column_letter].width = adjusted_width

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
    # Trả file về cho client
    response = make_response(output.read())
    response.headers['Content-Disposition'] = f'attachment; filename=danhsach_ungvien_{time_stamp}.xlsx'
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response 
      
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
        app.logger.error(e)
        flash(f"Đổi ca bị lỗi, {e} !!!")
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
                    print(file)
                    if file:
                        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
                        filepath = os.path.join(FOLDER_NHAP, f"doicanhom_{thoigian}.xlsx")
                        file.save(filepath)
                        data = pd.read_excel(filepath).to_dict(orient="records")
                        for row in data:
                            print(row)
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
        app.logger.error(e)
        flash(f"Đổi ca bị lỗi, {e} !!!")
        return redirect("/muc7_1_1")
        
@app.route("/laycatheomst", methods=["POST"])
def laycatheomst():
    mst = request.args.get("mst")
    ca = laycahientai(mst)
    return jsonify({
        "Ca": ca
    })
    
@app.route("/laycatheoline", methods=["POST"])
def laycatheoline():
    line = request.args.get("line")
    ca = laycatheochuyen(line)
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
    chuyen = request.form.get('chuyen')
    phongban = request.form.get('phongban')
    tungay = request.form.get("tungay")
    denngay = request.form.get("denngay")
    phanloai = request.form.get("phanloai")
    danhsach = laydanhsachchamcong(mst,chuyen,phongban,tungay,denngay,phanloai)
    result = [
            {'Nhà máy': row[0],
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
                'Phân loại': row[20]}
        for row in danhsach]
    df = DataFrame(result)
    output = BytesIO()
    with ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    # Điều chỉnh độ rộng cột
    output.seek(0)
    workbook = openpyxl.load_workbook(output)
    sheet = workbook.active

    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column_letter].width = adjusted_width

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
    # Trả file về cho client
    response = make_response(output.read())
    response.headers['Content-Disposition'] = f'attachment; filename=bang_chamcong_{time_stamp}.xlsx'
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response  

@app.route("/export_dscctt", methods=["POST"])
def export_dscctt():
    mst = request.form.get('mst')
    chuyen = request.form.get('chuyen')
    phongban = request.form.get('phongban')
    tungay = request.form.get("tungay")
    denngay = request.form.get("denngay")
    phanloai = request.form.get("phanloai")
    danhsach = laydanhsachchamcongchot(mst,chuyen,phongban,tungay,denngay,phanloai)
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
    df = DataFrame(result)
    output = BytesIO()
    with ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    # Điều chỉnh độ rộng cột
    output.seek(0)
    workbook = openpyxl.load_workbook(output)
    sheet = workbook.active

    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column_letter].width = adjusted_width

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
    # Trả file về cho client
    response = make_response(output.read())
    response.headers['Content-Disposition'] = f'attachment; filename=bang_chamcongchot_{time_stamp}.xlsx'
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response  

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
            mstquanly = request.form.get("mstquanly")
            mstthuky = request.form.get("mstthuky")
            # if mstdiemdanh==mstduyet:
            #     print(f"Bạn không thể kiểm tra cho chính mình, vui lòng liên hệ thư ký !!!")
            #     return redirect(f"/muc7_1_3?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&loaidiemdanh={loaidiemdanh_filter}&ngay={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
            if thuky_duoc_phanquyen(mstduyet,chuyen):
                if kiemtra == "Kiểm tra":    
                    thuky_dakiemtra_diemdanhbu(id)
                    flash(f"Thư ký {current_user.hoten} đã kiểm tra phiếu điểm danh bù số {id} !!!")
                else:
                    thuky_tuchoi_diemdanhbu(id)
                    flash(f"Thư ký {current_user.hoten} đã từ chối điểm danh bù phiếu số {id}  !!!")
            else:
                flash(f"{current_user.hoten} không có quyền điểm danh chuyền {chuyen} !!!")
        except Exception as e:
            flash(f"Lỗi thư ký điểm danh bù: {e}")
        if mstquanly:
            return redirect(f"/muc7_1_3?mstquanly={mstquanly}")
        else:
            if mstthuky:
                return redirect(f"/muc7_1_3?mstthuky={mstthuky}")
            else:
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
            mstquanly = request.form.get("mstquanly")
            mstthuky = request.form.get("mstthuky")
            if mstdiemdanh==mstduyet:
                print(f"Bạn không thể phê duyệt cho chính mình, vui lòng liên hệ quản lý !!!")
            if quanly_duoc_phanquyen(mstduyet,chuyen):
                if pheduyet == "Phê duyệt":    
                    quanly_pheduyet_diemdanhbu(id)
                    print(f"Quản lý {current_user.hoten} đã phê duyệt điểm danh bù cho phiếu số {id} !!!")
                else:
                    quanly_tuchoi_diemdanhbu(id)
                    print(f"Quản lý {current_user.hoten} đã từ chối điểm danh bù cho phiếu số {id}  !!!")
            else:
                print(f"{current_user.hoten} không có quyền phê duyệt !!!")
        except Exception as e:
            print(f"Lỗi quản lý phê duyệt điểm danh bù: {e}")
        if mstquanly:
            return redirect(f"/muc7_1_3?mstquanly={mstquanly}")
        else:
            if mstthuky:
                return redirect(f"/muc7_1_3?mstthuky={mstthuky}")
            else:  
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
            mstquanly = request.form.get("mstquanly")
            mstthuky = request.form.get("mstthuky")
            # if mstxinnghiphep==mstduyet:
            #     print(f"Bạn không thể kiểm tra cho chính mình, vui lòng liên hệ thư ký !!!")
            #     return redirect(f"/muc7_1_4?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&trangthai={trangthai_filter}")
            if thuky_duoc_phanquyen(mstduyet,chuyen):
                if kiemtra == "Kiểm tra":    
                    thuky_dakiemtra_xinnghiphep(id)
                    print(f"Thư ký {current_user.hoten} đã kiểm tra phiếu xin nghỉ phép số {id} !!!")
                else:
                    thuky_tuchoi_xinnghiphep(id)
                    print(f"Thư ký {current_user.hoten} từ chối phiếu nghỉ phép số {id} !!!")
            else:
                print(f"{current_user.hoten} không có quyền kiểm tra !!!")
        except Exception as e:
            print(f"Lỗi thư ký kiểm tra xin nghỉ phép: {e}")
            return redirect(f"/muc7_1_4?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&trangthai={trangthai_filter}")
        if mstquanly:
            return redirect(f"/muc7_1_4?mstquanly={mstquanly}")
        else:
            if mstthuky:
                return redirect(f"/muc7_1_4?mstthuky={mstthuky}")
            else:  
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
            mstquanly = request.form.get("mstquanly")
            mstthuky = request.form.get("mstthuky")
            if mstxinnghiphep==mstduyet:
                print(f"Bạn không thể phê duyệt cho chính mình, vui lòng liên hệ quản lý !!!")
            if quanly_duoc_phanquyen(mstduyet,chuyen):
                if pheduyet == "Phê duyệt":    
                    quanly_pheduyet_xinnghiphep(id)
                    print(f"Quản lý {current_user.hoten} đã hê duyệt cho phiếu xin nghỉ phép số {id} !!!")
                else:
                    quanly_tuchoi_xinnghiphep(id)
                    print(f"Quản lý {current_user.hoten} từ chối hê duyệt phiếu xin nghỉ phép số {id}  !!!")
            else:
                print(f"{current_user.hoten} không có quyền phê duyệt !!!")
        except Exception as e:
            print(f"Lỗi quản lý phê duyệt xin nghỉ phép: {e}")
        if mstquanly:
            return redirect(f"/muc7_1_3?mstquanly={mstquanly}")
        else:
            if mstthuky:
                return redirect(f"/muc7_1_4?mstthuky={mstthuky}")
            else:  
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
            mstquanly = request.form.get("mstquanly")
            mstthuky = request.form.get("mstthuky")
            # if mstxinnghikhongluong==mstduyet:
            #     print(f"Bạn không thể kiểm tra cho chính mình, vui lòng liên hệ thư ký !!!")
            #     return redirect(f"/muc7_1_5?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
            if thuky_duoc_phanquyen(mstduyet,chuyen):
                if kiemtra == "Kiểm tra":    
                    thuky_dakiemtra_xinnghikhongluong(id)
                    print(f"Thư ký {current_user.hoten} đã kiểm tra cho phiếu xin nghỉ không lương số {id} !!!")
                else:
                    thuky_tuchoi_xinnghikhongluong(id)
                    print(f"Thư ký {current_user.hoten} từ chối kiểm tra phiếu xin nghỉ không lương số {id}  !!!")
            else:
                print(f"{current_user.hoten} không có quyền kiểm tra !!!")
        except Exception as e:
            print(f"Lỗi thư ký kiểm tra xin nghỉ không lương: {e}")
        if mstquanly:
            return redirect(f"/muc7_1_5?mstquanly={mstquanly}")
        else:
            if mstthuky:
                return redirect(f"/muc7_1_5?mstthuky={mstthuky}")
            else:
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
            mstquanly = request.form.get("mstquanly")
            mstthuky = request.form.get("mstthuky")
            if mstxinnghikhongluong==mstduyet:
                print(f"Bạn không thể phê duyệt cho chính mình, vui lòng liên hệ quản lý !!!")
            if quanly_duoc_phanquyen(mstduyet,chuyen):
                if pheduyet == "Phê duyệt":    
                    quanly_pheduyet_xinnghikhongluong(id)
                    print(f"Quản lý {current_user.hoten} đã phê duyệt cho phiếu xin nghỉ không lương số {id} !!!")
                else:
                    quanly_tuchoi_xinnghikhongluong(id)
                    print(f"Quản lý {current_user.hoten} ttừ chối phê duyệt phiếu xin nghỉ không lương số {id}  !!!")
            else:
                print(f"{current_user.hoten} không có quyền phê duyệt !!!")
        except Exception as e:
            print(f"Lỗi quản lý phê duyệt xin nghỉ không lương: {e}")
        if mstquanly:
            return redirect(f"/muc7_1_5?mstquanly={mstquanly}")
        else:
            if mstthuky:
                return redirect(f"/muc7_1_5?mstthuky={mstthuky}")
            else:   
                return redirect(f"/muc7_1_5?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&ngaynghi={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")

@app.route("/thuky_kiemtra_xinnghikhac", methods=["POST"])
def thukykiemtraxinnghikhac():
    if request.method == "POST":
        try:
            mst = request.form.get("mst_xinnghikhac")
            mst_filter = request.form.get("mst_filter")
            chuyen = lay_chuyen_theo_mst(mst)
            mstduyet = current_user.masothe
            kiemtra = request.form.get("kiemtra")
            id = request.form["id"]
            # if mstdiemdanh==mstduyet:
            #     print(f"Bạn không thể kiểm tra cho chính mình, vui lòng liên hệ thư ký !!!")
            #     return redirect(f"/muc7_1_3?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&loaidiemdanh={loaidiemdanh_filter}&ngay={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
            if thuky_duoc_phanquyen(mstduyet,chuyen):
                if kiemtra == "Kiểm tra":    
                    thuky_dakiemtra_xinnghikhac(id)
                    flash(f"Thư ký {current_user.hoten} đã kiểm tra phiếu xin nghỉ khác số {id} !!!")
                else:
                    thuky_tuchoi_xinnghikhac(id)
                    flash(f"Thư ký {current_user.hoten} đã từ chối xin nghỉ khác phiếu số {id}  !!!")
            else:
                flash(f"{current_user.hoten} không có quyền điểm danh chuyền {chuyen} !!!")
        except Exception as e:
            flash(f"Lỗi thư ký xin nghỉ khác: {e}")
        return redirect(f"/muc7_1_6?mst={mst_filter}")
        
@app.route("/quanly_pheduyet_xinnghikhac", methods=["POST"])
def quanlypheduyetxinnghikhac():
    if request.method == "POST":
        try:
            mst = request.form.get("mst_xinnghikhac")
            mst_filter = request.form.get("mst_filter")
            chuyen = lay_chuyen_theo_mst(mst)
            mstduyet = current_user.masothe
            pheduyet = request.form.get("pheduyet")
            id = request.form["id"]
            if mst==mstduyet:
                print(f"Bạn không thể phê duyệt cho chính mình, vui lòng liên hệ quản lý !!!")
            if quanly_duoc_phanquyen(mstduyet,chuyen):
                if pheduyet == "Phê duyệt":    
                    quanly_pheduyet_xinnghikhac(id)
                    print(f"Quản lý {current_user.hoten} đã phê duyệt xin nghỉ khác cho phiếu số {id} !!!")
                else:
                    quanly_tuchoi_xinnghikhac(id)
                    print(f"Quản lý {current_user.hoten} đã từ chối xin nghỉ khác cho phiếu số {id}  !!!")
            else:
                print(f"{current_user.hoten} không có quyền phê duyệt !!!")
        except Exception as e:
            print(f"Lỗi quản lý phê duyệt xin nghỉ khác: {e}")
        return redirect(f"/muc7_1_6?mst={mst_filter}")

@app.route("/nhansu_nhangiayto_xinnghikhac", methods=["POST"])
def nhansunhangiaytoxinnghikhac():
    if request.method == "POST":
        try:
            mst_filter = request.form.get("mst_filter")
            id = request.form["id"]
            nhangiayto = request.form.get("nhangiayto")
            # if mstdiemdanh==mstduyet:
            #     print(f"Bạn không thể kiểm tra cho chính mình, vui lòng liên hệ thư ký !!!")
            #     return redirect(f"/muc7_1_3?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&loaidiemdanh={loaidiemdanh_filter}&ngay={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
            if (current_user.macongty=='NT1' and current_user.masothe==2833) or (current_user.macongty=='NT2' and current_user.masothe==2176) or (current_user.macongty=='NT2' and current_user.masothe==1369 ):
                if nhangiayto == "Có":    
                    nhansu_nhangiayto_xinnghikhac(id)
                    flash(f"Thư ký {current_user.hoten} đã nhận giấy tờ cho phiếu xin nghỉ khác số {id} !!!")
                else:
                    nhansu_khongnhangiayto_xinnghikhac(id)
                    flash(f"Thư ký {current_user.hoten} không nhận được giấy tờ cho phiếu số {id}  !!!")
            else:
                flash(f"{current_user.hoten} không có quyền nhận giấy tờ, vui lòng liên hệ HRD !!!")
        except Exception as e:
            flash(f"Lỗi thư ký xin nghỉ khác: {e}")
        return redirect(f"/muc7_1_6?mst={mst_filter}")

@app.route("/taifilemaukp", methods=["GET"])
def taifilemaukp():
    if request.method == "GET":
        try:
            file = FILE_MAU_DANGKY_KPI
            return send_file(file, as_attachment=True)
            
        except Exception as e:
            print(e)
            print("Download file error !!!")
            return redirect("/muc5_1_1")

@app.route("/rutdonxinnghiviec", methods=["POST"])
def rutdonxinnghiviec():
    if request.method == "POST":
        try:
            id = request.form.get("id")
            if rutdonnghiviec(id):
                print("Rút đơn nghỉ việc thành công !!!")
            else:
                print("Rút đơn nghỉ việc thất bại !!!")
            return redirect("/muc10_2")
        except Exception as e:
            print(e)
            print(f"Rút đơn bị lỗi ({e}) !!!")
            return redirect("/muc10_2")    
        
@app.route("/capnhatstk", methods=["POST"])
def capnhatstk():
    file = request.files.get("file")
    if file:
        try:
            thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
            filepath = os.path.join(FOLDER_NHAP, f"capnhatstk_{thoigian}.xlsx")
            file.save(filepath)
            print("Upload file success !!!")
            data = pd.read_excel(filepath, dtype={0: str,1: str}).to_dict(orient="records")
            for row in data:
                macongty = row['Mã công ty']
                mst= row['Mã số nhân viên']
                stk = row['Số tài khoản ngân hàng']
                if macongty == current_user.macongty:   
                    capnhat_stk(mst, stk, macongty)
        except Exception as e:
            print(f"Upload file error ({e}) !!!")
    else:
        print("Not found file !!!")
    return redirect("/muc3_2")

@app.route("/taifile_capnhatstk", methods=["POST"])
def taifile_capnhatstk():
    return send_file(FILE_MAU_CAPNHAT_STK, as_attachment=True)

@app.route("/inhopdong", methods=["POST"])
def inhopdong():
    if request.method=="POST":
        id = request.form.get("idhopdongin")
        hopdong = lay_thongtin_hopdong_theo_id(id)
        macongty = hopdong[1]
        masothe = hopdong[2]
        hoten = hopdong[3]
        gioitinh = hopdong[4]
        ngaysinh = datetime.strptime(hopdong[5], "%Y-%m-%d").strftime("%d/%m/%Y")
        thuongtru = hopdong[6]
        tamtru = hopdong[7]
        cccd = hopdong[8]
        ngaycapcccd = datetime.strptime(hopdong[9], "%Y-%m-%d").strftime("%d/%m/%Y")
        capbac = hopdong[10]
        loaihopdong = hopdong[11]
        chucdanh = hopdong[12]
        phongban = hopdong[13]
        chuyen = hopdong[14]
        luongcoban = hopdong[15]
        phucap = hopdong[16]
        ngaybatdau = datetime.strptime(hopdong[17], "%Y-%m-%d").strftime("%d/%m/%Y")
        ngayketthuc = datetime.strptime(hopdong[18], "%Y-%m-%d").strftime("%d/%m/%Y") if hopdong[18] else None
        file = inhopdongtheomau(macongty,masothe,hoten,gioitinh,ngaysinh,thuongtru,tamtru,cccd,ngaycapcccd,capbac,loaihopdong,chucdanh,phongban,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc)
        # print(file)
        if file:
            return send_file(file, as_attachment=True, download_name="hopdong.xlsx")
        else:
            return redirect("/muc3_3")

@app.route("/timcacchucdanh", methods=["POST"])
def timcacchucdanh():
    tutimkiem = request.args.get("tutimkiem")
    cacchucdanh = timkiemchucdanh(tutimkiem)
    return jsonify(cacchucdanh)

@app.route("/taifilethemhopdongmau", methods=["POST"])
def taifilethemhopdongmau():
    return send_file(FILE_MAU_THEM_HOPDONG, as_attachment=True, download_name="themhopdong.xlsx")
        
@app.route("/capnhathopdongtheofilemau", methods=["POST"])
def capnhathopdongtheofilemau():
    file = request.files.get("file")
    if file:
        try:
            thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
            filepath = os.path.join(FOLDER_NHAP, f"themhopdong_{thoigian}.xlsx")
            file.save(filepath)
            data = pd.read_excel(filepath).to_dict(orient="records")
            for row in data:
                nhamay = row['Mã công ty']
                mst = row['MST']
                hoten = row['Họ tên']
                gioitinh = row['Giới tính']
                ngaysinh = row['Ngày sinh']
                thuongtru = row['Địa chỉ thường trú']
                tamtru = row["Địa chỉ tạm trú"]
                cccd = row['CCCD']
                ngaycapcccd = row['Ngày cấp cccd']
                capbac = row['Cấp bậc']
                loaihopdong = row['Loại hợp đồng']
                luongcoban = row['Lương cơ bản']
                phucap = row['Phụ cấp']
                ngaybatdau = row['Ngày bắt đầu HĐ']
                ngayketthuc = row['Ngày kết thúc HĐ']
                chucdanh = row['Chức danh']
                phongban = row['Phòng ban']
                chuyen = row['Chuyền']
                hcname= layhcname(chucdanh,chuyen)
                if hcname:
                    vitrien = hcname[2]
                    employeetype = hcname[3]
                    posotioncode = hcname[4]
                    postitioncodedescription = hcname[5]
                    hccategory = hcname[7]
                    sectioncode = hcname[10]
                    sectiondescription = hcname[11]
                else:
                    vitrien = 'NULL'
                    employeetype = 'NULL'
                    posotioncode = 'NULL'
                    postitioncodedescription = 'NULL'
                    hccategory = 'NULL'
                    sectioncode = 'NULL'
                    sectiondescription = 'NULL'
                if themhopdongmoi(nhamay, mst, hoten, gioitinh, ngaysinh, thuongtru, tamtru, cccd, ngaycapcccd, capbac, loaihopdong, chucdanh, phongban, chuyen, luongcoban, phucap, ngaybatdau, ngayketthuc):
                    print("Them HD ok")
                if capnhatthongtinhopdong(nhamay,mst,loaihopdong,chucdanh,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc,vitrien,employeetype,posotioncode,postitioncodedescription,hccategory,sectioncode,sectiondescription):
                    print("Cap nhap HD ok")
            print("Cập nhật hợp đồng thành công !!!")
        except Exception as e:
            print(f"Cap nhat hop dong loi: {e}")
            print(f"Cập nhật hợp đồng lỗi: ({e}) !!!")
    else:
        print("Không tìm thấy dữ liệu hợp đồng !!!")
    return redirect("/muc3_3")

@app.route("/suahopdong", methods=["POST"])
def suahopdong():
    try:
        id = request.form.get('idhopdongsua')
        if id:
            hopdong = lay_thongtin_hopdong_theo_id(id)
            return render_template("suahopdong.html",hopdong=hopdong)
        return redirect("/muc3_3")
    except Exception as e:
        print(e)
        return redirect("/muc3_3")  
    
@app.route("/xoahopdong", methods=["POST"])
def xoahopdong():
    try:
        id = request.form.get('idhopdongxoa')
        print(id)
        if xoa_hopdong(id):
            flash(f"Xoá thành công hợp đồng có ID {id}")
        else:
            flash(f"Xoá hợp đồng có ID {id} không thành công !!!")
        return redirect("/muc3_3")
    except Exception as e:
        print(e)
        return redirect("/muc3_3")    
    
@app.route("/suahopdonglaodong", methods=["POST"])
def suahopdonglaodong():
    try:
        id = request.form.get('id_hopdong')
        
        masothe = request.form.get('masothe')
        hoten = request.form.get('hovaten')
        gioitinh = request.form.get('gioitinh')
        ngaysinh = request.form.get('ngaysinh')
        thuongtru = request.form.get('thuongtru')
        tamtru = request.form.get('tamtru')
        cccd = request.form.get('cccd')
        ngaycapcccd = request.form.get('ngaycapcccd')
        
        loaihopdong = request.form.get('loaihopdong')
        ngaybatdau = request.form.get('ngaykyhopdong')
        ngayketthuc = request.form.get('ngayhethanhopdong')
        
        chuyen = request.form.get('chuyen')
        capbac = request.form.get('gradecode')
        chucdanh = request.form.get('chucdanh')
        phongban = request.form.get('department')
        
        luongcoban = request.form.get('luongcoban')
        phucap = request.form.get('phucap')
        
        if thaydoithongtinhopdong(id,masothe,hoten,gioitinh,ngaysinh,thuongtru,tamtru,cccd,
                                  ngaycapcccd,loaihopdong,ngaybatdau,ngayketthuc,chuyen,capbac,
                                  chucdanh,phongban,luongcoban,phucap):
            print(f"Cập nhật hợp đồng số {id} thành công !!!")
        else:
            print(f"Cập nhật hợp đồng số {id} thất bại !!!")
    except Exception as e:
        print(e)
    return redirect("/muc3_3")  

@app.route("/qr_code", methods=["GET"])
def load_qr_code():
    kieu_qr = request.args.get("qr")
    if kieu_qr=="hp_diemdanhbu":
        qr_file = "hp_diemdanhbu.jpg"
    elif kieu_qr=="na_diemdanhbu":
        qr_file = "na_diemdanhbu.jpg"
    elif kieu_qr=="hp_xinnghiphep":
        qr_file = "hp_xinnghiphep.jpg"
    elif kieu_qr=="na_xinnghiphep":
        qr_file = "na_xinnghiphep.jpg"
    elif kieu_qr=="hp_xinnghikhongluong":
        qr_file = "hp_xinnghikhongluong.jpg"
    elif kieu_qr=="na_xinnghikhongluong":
        qr_file = "na_xinnghikhongluong.jpg"    
    else:
        qr_file = "linkphanmem.png"
    return render_template("qr_code.html", qr_file = qr_file, page="QR CODE")

@app.route("/diemdanhbu", methods=["POST"])
def diemdanhbu_web():
    try:
        if request.method == "POST":
            masothe = request.form.get("masothe_diemdanhbu")
            hoten = request.form.get("hoten_diemdanhbu")
            chuyen = request.form.get("chuyento_diemdanhbu")
            phongban = request.form.get("phongban_diemdanhbu")
            chucdanh = request.form.get("chucdanh_diemdanhbu")
            ngay = request.form.get("ngay_diemdanhbu")
            giovao = request.form.get("giovao_diemdanhbu")
            giora = request.form.get("giora_diemdanhbu")
            lydo = request.form.get("lydo_diemdanhbu")
            trangthai = "Chờ kiểm tra"
            if giovao:
                loaidiemdanh = "Điểm danh vào"
                if them_diemdanhbu(masothe,hoten,chucdanh,chuyen,phongban,loaidiemdanh,ngay,giovao,lydo,trangthai):
                    print(f"Thêm điểm danh vào cho {hoten} vào ngày {ngay} thành công !!!")
                else:
                    print(f"Thêm điểm danh vào cho {hoten} vào ngày {ngay} thất bại !!!")
            if giora:
                loaidiemdanh = "Điểm danh ra"
                if them_diemdanhbu(masothe,hoten,chucdanh,chuyen,phongban,loaidiemdanh,ngay,giora,lydo,trangthai):
                    print(f"Thêm điểm danh ra cho {hoten} vào ngày {ngay} thành công !!!") 
                else:
                    print(f"Thêm điểm danh vào cho {hoten} vào ngày {ngay}  thất bại !!!")
            return redirect(f"/muc7_1_2?chuyen={chuyen}")
    except Exception as e:
        print(f"Them diem danh bu loi {e}")
        print(f"Thêm điểm danh bù lỗi: {str(e)}")
        return redirect("/muc7_1_2")
    
@app.route("/xinnghiphep", methods=["POST"])
def xinnghiphep_web():
    try:
        masothe = request.form.get("masothe_xinnghiphep")
        hoten = request.form.get("hoten_xinnghiphep")
        chuyen = request.form.get("chuyento_xinnghiphep")
        phongban = request.form.get("phongban_xinnghiphep")
        chucdanh = request.form.get("chucdanh_xinnghiphep")
        ngay = request.form.get("ngay_xinnghiphep")
        sophut = request.form.get("sophut_xinnghiphep")
        trangthai = "Chờ kiểm tra"
        if them_xinnghiphep(masothe,hoten,chucdanh,chuyen,phongban,ngay,sophut,trangthai):
            print(f"Thêm xin nghỉ phép cho {hoten} vào ngày {ngay} thành công !!!")
        else:
            print(f"Thêm xin nghỉ phép cho {hoten} vào ngày {ngay} thất bại !!!")
        return redirect(f"/muc7_1_2?chuyen={chuyen}")
    except Exception as e:
        print(f"Them xin nghi phep loi {e}")
        print(f"Thêm xin nghỉ phép lỗi: {str(e)}")
        return redirect("/muc7_1_2")

@app.route("/xinnghikhongluong", methods=["POST"])
def xinnghikhongluong_web():
    try:
        masothe = request.form.get("masothe_xinnghikhongluong")
        hoten = request.form.get("hoten_xinnghikhongluong")
        chuyen = request.form.get("chuyento_xinnghikhongluong")
        phongban = request.form.get("phongban_xinnghikhongluong")
        chucdanh = request.form.get("chucdanh_xinnghikhongluong")
        ngay = request.form.get("ngay_xinnghikhongluong")
        sophut = request.form.get("sophut_xinnghikhongluong")
        lydo = request.form.get("lydo_xinnghikhongluong")
        trangthai = "Chờ kiểm tra"
        if them_xinnghikhongluong(masothe,hoten,chucdanh,chuyen,phongban,ngay,sophut,lydo,trangthai):
            print(f"Thêm xin nghỉ không lương cho {hoten} vào ngày {ngay} thành công !!!")
        else:
            print(f"Thêm xin nghỉ không lương cho {hoten} vào ngày {ngay} thất bại !!!")
        return redirect(f"/muc7_1_2?chuyen={chuyen}")
    except Exception as e:
        print(f"Them xin nghi khong luong loi {e}")
        print(f"Thêm xin nghỉ không lương lỗi: {str(e)}")
        return redirect("/muc7_1_2")

@app.route("/xinnghikhac", methods=["POST"])
def xinnghikhac_web():
    try:
        masothe = request.form.get("masothe_xinnghikhac")
        hoten = request.form.get("hoten_xinnghikhac")
        chuyen = request.form.get("chuyento_xinnghikhac")
        phongban = request.form.get("phongban_xinnghikhac")
        chucdanh = request.form.get("chucdanh_xinnghikhac")
        ngay = request.form.get("ngay_xinnghikhac")
        sophut = request.form.get("sophut_xinnghikhac")
        lydo = request.form.get("lydo_xinnghikhac")
        trangthai = "Chờ kiểm tra"
        nhangiayto = "Chưa"
        if them_xinnghikhac(masothe,ngay,sophut,lydo,trangthai,nhangiayto):
            print(f"Thêm xin nghỉ khác cho {hoten} vào ngày {ngay} thành công !!!")
        else:
            print(f"Thêm xin nghỉ khác cho {hoten} vào ngày {ngay} thất bại !!!")
        return redirect(f"/muc7_1_2?chuyen={chuyen}")
    except Exception as e:
        print(f"Them xin nghi khac loi {e}")
        print(f"Thêm xin nghỉ khác lỗi: {str(e)}")
        return redirect("/muc7_1_2")
    
@app.route("/taidanhsachdonxinnghiviec", methods=["POST"])
def taidanhsachdonxinnghiviec():
    mst = request.form.get("mst")
    hoten = request.form.get("hoten")
    chuyen = request.form.get("chuyen")
    phongban = request.form.get("phongban")
    ngaynopdon = request.form.get("ngaynopdon")
    ngaynghi = request.form.get("ngaynghi")
    sapdenhan = request.form.get("sapdenhan")
    danhsach = laydanhsach_chonghiviec(mst,hoten,chuyen,phongban,ngaynopdon,ngaynghi,sapdenhan)
    data = [{
        "Mã số thẻ": row[2],
        "Họ tên": row[3],
        "Chức danh": row[4],
        "Chuyền": row[5],
        "Phòng ban": row[6],
        "Ngày nộp đơn": row[7],
        "Ngày nghỉ dự kiến": row[8],
        "Ghi chú": row[9],
        "Trạng thái làm việc": row[10]
    } for row in danhsach]
    df = DataFrame(data)
    output = BytesIO()
    with ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    # Điều chỉnh độ rộng cột
    output.seek(0)
    workbook = openpyxl.load_workbook(output)
    sheet = workbook.active

    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column_letter].width = adjusted_width

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
    # Trả file về cho client
    response = make_response(output.read())
    response.headers['Content-Disposition'] = f'attachment; filename=danhsach_donxinghiviec_{time_stamp}.xlsx'
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response   

@app.route("/capnhat_lichsu_congtac", methods=["POST"])
def capnhat_lichsu_congtac():
    
    mst_filter = request.form.get("mst_filter")
    mst = request.form.get("mst")
    ngaythuchien = request.form.get("ngaythuchien")
    phanloai = request.form.get("phanloai")
    ghichumoi = request.form.get("ghichu")   
    try:
        if capnhat_ghichu_lichsu_congtac(mst,ngaythuchien,phanloai,ghichumoi):
            print(f"Cap nhat Lich su cong tac mst={mst} thanh cong")
        else:
            print(f"Cap nhat Lich su cong tac mst={mst} that bai")
    except Exception as e:
        print(f"Loi khi cap nhat lich su cong tac ({e})")
    return redirect(f"/muc6_2?mst={mst_filter}")

@app.route("/suadoi_dangky_ca", methods=["POST"])
def suadoi_dangky_ca():
    mst_filter = request.form.get("mst_filter")
    chuyen_filter = request.form.get("chuyen_filter")
    bophan_filter = request.form.get("bophan_filter")
    id = request.form.get("id")
    camoi = request.form.get("ca")
    try:
        if sua_dangky_ca(id,camoi):
            print(f"Sua dang ký ca id = {id} thanh cong")
        else:
            print(f"Sua dang ký ca id = {id} that bai")
    except Exception as e:
        print(f"Loi khi cap nhat lich su cong tac ({e})")
    return redirect(f"/muc7_1_1?mst={mst_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}")

@app.route("/suadoi_ngaybatdau_ca", methods=["POST"])
def suadoi_ngaybatdau_ca():
    mst_filter = request.form.get("mst_filter")
    chuyen_filter = request.form.get("chuyen_filter")
    bophan_filter = request.form.get("bophan_filter")
    id = request.form.get("id")
    ngaybatdau_camoi = request.form.get("ngaybatdau_ca")
    try:
        if suadoi_ngaybatdau_ca_dangky_ca(id,ngaybatdau_camoi):
            print(f"Sua dang ký ca id = {id} thanh cong")
        else:
            print(f"Sua dang ký ca id = {id} that bai")
    except Exception as e:
        print(f"Loi khi cap nhat lich su cong tac ({e})")
    return redirect(f"/muc7_1_1?mst={mst_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}")

@app.route("/suadoi_ngayketthuc_ca", methods=["POST"])
def suadoi_ngayketthuc_ca():
    mst_filter = request.form.get("mst_filter")
    chuyen_filter = request.form.get("chuyen_filter")
    bophan_filter = request.form.get("bophan_filter")
    id = request.form.get("id")
    ngayketthuc_camoi = request.form.get("ngayketthuc_ca")
    try:
        if suadoi_ngayketthuc_ca_dangky_ca(id,ngayketthuc_camoi):
            print(f"Sua dang ký ca id = {id} thanh cong")
        else:
            print(f"Sua dang ký ca id = {id} that bai")
    except Exception as e:
        print(f"Loi khi cap nhat lich su cong tac ({e})")
    return redirect(f"/muc7_1_1?mst={mst_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}")

@app.route("/bat_12", methods=["POST"])
def on_f12():
    try:
        if request.method == "POST":
            bat_function_12()
    except Exception as e:
        print(e)
    return redirect("/admin")

@app.route("/tat_12", methods=["POST"])
def off_f12():
    try:
        if request.method == "POST":
            tat_function_12()
    except Exception as e:
        print(e)
    return redirect("/admin")

@app.route("/dangki_tangca_web", methods=["GET","POST"])
def dangky_tangca_bangweb():
    if request.method=="GET":
        chuyen = request.args.getlist("chuyen")
        ngay = request.args.get("ngay")    
        cacchuyen = laychuyen_quanly(current_user.masothe,current_user.macongty)    
        danhsach = danhsach_tangca(chuyen,ngay)
        count = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 100
        total = len(danhsach)
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("dangky_tangca_web.html",
                               cacchuyen=cacchuyen,
                               danhsach=paginated_rows, 
                                pagination=pagination,
                                count=count)
        
@app.route("/capnhat_tangca", methods=["POST"])
def capnhat_tangca():
    if request.method=="POST":
        
        id = request.form.get("id")
        tangcasang = request.form.get("tangcasang")
        tangcasangthucte = request.form.get("tangcasangthucte")
        tangca = request.form.get("tangca")
        tangcathucte = request.form.get("tangcathucte")
        tangcadem = request.form.get("tangcadem")
        tangcademthucte = request.form.get("tangcademthucte")
        
        chuyen_filter = request.form.get("chuyen_filter")
        ngay_filter = request.form.get("ngay_filter")
        try:  
            if capnhat_tangca_thanhcong(id,tangcasang,tangcasangthucte,tangca,tangcathucte,tangcadem,tangcademthucte):
                flash(f"Cập nhật tăng ca id = {id} thành công")
            else:
                flash(f"Cập nhật tăng ca id = {id} thất bại")
        except Exception as e:   
            flash(f"Loi khi cap nhat tang ca ({e})")
        return redirect(f"/dangki_tangca_web?chuyen={chuyen_filter}&ngay={ngay_filter}")
    
@app.route("/bopheduyet_tangca", methods=["POST"])   
def bopheduyet_tangca():
    if request.method=="POST":
        chuyen_filter = request.form.get("chuyen_filter")
        ngay_filter = request.form.get("ngay_filter")
        id = request.form.get("id")
        try:
            if nhansu_bopheduyet_tangca(id):
                flash(f"Bỏ phê duyệt tăng ca ID = {id}")
            else:
                flash(f"Bỏ phê duyệt tăng ca ID = {id} không được")
        except Exception as e:   
            flash(f"Loi khi bo phe duyet tang ca tang ca ({e})")
        return redirect(f"/dangki_tangca_web?chuyen={chuyen_filter}&ngay={ngay_filter}")

@app.route("/pheduyet_tangca", methods=["POST"])   
def pheduyet_tangca():
    if request.method=="POST":
        chuyen_filter = request.form.get("chuyen_filter")
        ngay_filter = request.form.get("ngay_filter")
        id = request.form.get("id")
        try:
            if nhansu_pheduyet_tangca(id):
                flash(f"Phê duyệt tăng ca ID = {id}")
            else:
                flash(f"Phê duyệt tăng ca ID = {id} không được")
        except Exception as e:   
            flash(f"Loi khi bo phe duyet tang ca tang ca ({e})")
        return redirect(f"/dangki_tangca_web?chuyen={chuyen_filter}&ngay={ngay_filter}")
    
@app.route("/chamcong_sang_web", methods=["GET","POST"])
def chamcong_sang_web():
    if request.method=="GET":
        chuyen = request.args.get("chuyen")
        bophan = request.args.get("bophan")
        cochamcong = request.args.get("cochamcong") 
        ngay = datetime.now().date()
        danhsach = danhsach_chamcong_sang(chuyen,bophan,cochamcong)
        count = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 20
        total = len(danhsach)
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("chamcong_sang_web.html",
                               danhsach=paginated_rows, 
                                pagination=pagination,
                                count=count,
                                ngay=ngay)
    else:
        chuyen = request.args.get("chuyen")
        bophan = request.args.get("bophan")
        cochamcong = request.args.get("cochamcong")  
        danhsach = danhsach_chamcong_sang(chuyen,bophan,cochamcong)
        data = [{
        "Mã số thẻ": row[1],
        "Họ tên": row[2],
        "Chuyền": row[4],
        "Phòng ban": row[5],
        "Ngày": row[6],
        "Giờ vào": row[7]
        } for row in danhsach]
        df = DataFrame(data)
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=danhsach_chamcongsang_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response   
        
@app.route("/nhansu_themxinnghikhac", methods=["POST"])
def nhansu_them_xinnghikhac():
    if request.method=="POST":
        file = request.files.get("file")
        if file:
            try:
                thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
                filepath = os.path.join(FOLDER_NHAP, f"themxinnghikhac_{thoigian}.xlsx")
                file.save(filepath)
                data = pd.read_excel(filepath ).to_dict(orient="records")
                for row in data:
                    try:
                        masothe = int(row['Mã số thẻ'])
                        ngaynghi = str(row['Ngày nghỉ'])[:10]
                        sophut = int(row['Tổng số phút'])
                        loainghi = row['Loại nghỉ']
                        trangthai = "Đã phê duyệt"
                        nhangiayto = "Đã nhận"
                        print(masothe,ngaynghi,sophut,loainghi)
                        if them_xinnghikhac(masothe,ngaynghi,sophut,loainghi,trangthai,nhangiayto):
                            print("OK")
                        else:
                            print(data.index(row),"Failed")
                    except Exception as e:
                        print(f"Loi them dong xin nghi khac: {e}")
                        break
            except Exception as e:
                print(e)
        return redirect("/muc7_1_6")

@app.route("/tai_danhsach_tangca", methods=["POST"])
def tai_danhsach_tangca():
    if request.method=="POST":
        chuyen = request.form.get("chuyen").replace("[","").replace("]", "").replace("'", "").replace(" ", "").split(",")
        ngay = request.form.get("ngay")
        # print(chuyen, ngay)   
        danhsach = danhsach_tangca(chuyen,ngay)
        data = [x for x in danhsach if not x["HR phê duyệt"]]
        # print(data)
        ngay = datetime.now().date()     
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='coerce')
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=danhsach_tangca_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response   
    
@app.route("/tailen_danhsach_tangca", methods=["POST"])
def tailen_danhsach_tangca():
    if request.method=="POST":
        file = request.files.get("file")
        if file:
            try:
                thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
                filepath = os.path.join(FOLDER_NHAP, f"danhsach_tangca_{thoigian}.xlsx")
                file.save(filepath)
                data = pd.read_excel(filepath ).to_dict(orient="records")
                for row in data:
                    nhamay = row['Nhà máy']
                    mst = int(row["Mã số thẻ"])
                    hoten = row["Họ tên"]
                    chucdanh = row["Chức danh"]
                    chuyen = row["Chuyền"]
                    phongban = row["Phòng ban"]
                    ngay = row["Ngày"] 
                    giotangcasang = row["Tăng ca sáng"] if not pd.isna(row["Tăng ca sáng"]) else ""
                    giotangcasangthucte = row["Tăng ca sáng thực tế"] if not pd.isna(row["Tăng ca sáng thực tế"]) else ""
                    giotangca = row["Giờ tăng ca"] if not pd.isna(row["Giờ tăng ca"]) else ""
                    giotangcathucte = row["Giờ tăng ca thực tế"] if not pd.isna(row["Giờ tăng ca thực tế"]) else ""
                    giotangcadem = row["Tăng ca đêm"] if not pd.isna(row["Tăng ca đêm"]) else ""
                    giotangcademthucte = row["Tăng ca đêm thực tế"] if not pd.isna(row["Tăng ca đêm thực tế"]) else ""
                    ca = row["Ca"] if not pd.isna(row["Ca"]) else ""
                    giovao = row["Giờ vào"] if not pd.isna(row["Giờ vào"]) else ""
                    giora = row["Giờ ra"] if not pd.isna(row["Giờ ra"]) else ""
                    hrpheduyet = row["HR phê duyệt"] if not pd.isna(row["HR phê duyệt"]) else ""
                    if them_dangky_tangca(nhamay, mst, hoten, chucdanh, chuyen, phongban, ngay, giotangcasang, giotangcasangthucte, giotangca, giotangcathucte, giotangcadem, giotangcademthucte, ca, giovao, giora, hrpheduyet):
                        flash("Thêm đăng ký tăng ca thành công !!!")
                    else:
                        flash("Thêm đăng ký tăng ca thất bại !!!")        
            except Exception as e:
                print(e)
                    
        return redirect("/dangki_tangca_web")

@app.route("/taifilemaudieuchuyen", methods=["GET"])
def taifilemaudieuchuyen():
    if request.method=="GET":
        return send_file(FILE_MAU_DIEU_CHUYEN, as_attachment=True, download_name="dieuchuyen.xlsx")

@app.route("/capnhatdieuchuyentheofile", methods=["POST"])
def capnhatdieuchuyentheofile():
    if request.method=="POST":
        file = request.files.get("file")
        if file:
            try:
                thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
                filepath = os.path.join(FOLDER_NHAP, f"danhsach_tangca_{thoigian}.xlsx")
                file.save(filepath)
                data = pd.read_excel(filepath ).to_dict(orient="records")
                x = 1
                for row in data:
                    masothe = row["Mã số thẻ"]
                    chuyenmoi = row["Chuyền mới"]
                    chucdanhmoi = row["Chức danh mới"]
                    loaidieuchuyen = row["Loại điều chuyển"]
                    ngay = row["Ngày"]
                    ghichu = row["Ghi chú"]
                    hople = kiemtra_thongtin_dieuchuyen(x,masothe,chucdanhmoi,chuyenmoi,loaidieuchuyen)
                    print(hople)
                    if not hople["ketqua"]:
                        flash(f"Dòng {x} sai thông tin: {hople["lydo"]}")
                        return redirect("/muc6_2") 
                    else:
                        x += 1
                        
                for row in data:      
                    masothe = row["Mã số thẻ"]
                    chuyenmoi = row["Chuyền mới"]
                    chucdanhmoi = row["Chức danh mới"]
                    loaidieuchuyen = row["Loại điều chuyển"]
                    ngay = row["Ngày"]
                    ghichu = row["Ghi chú"] 
                    if loaidieuchuyen == "Chuyển vị trí":
                        
                        thongtin_laodong = laydanhsachtheomst(masothe)[0]
                        chucdanhcu = thongtin_laodong["Job title VN"]
                        chuyencu = thongtin_laodong["Line"]
                        capbaccu = thongtin_laodong["Gradecode"]
                        sectioncodecu = thongtin_laodong["Section code"]
                        hccategorycu = thongtin_laodong["HC category"]
                        phongbancu = thongtin_laodong["Department"]
                        sectiondescriptioncu = thongtin_laodong["Section description"]
                        employeetypecu = thongtin_laodong["Employee type"]
                        positioncodedescriptioncu = thongtin_laodong["Position description"]
                        positioncodecu = thongtin_laodong["Position code"]
                        chucdanhtacu = thongtin_laodong["Job title EN"]
                        
                        hc_name_moi = layhcname(chucdanhmoi,chuyenmoi)
                        capbacmoi = hc_name_moi[6]
                        sectioncodemoi = hc_name_moi[10]
                        hccategorymoi = hc_name_moi[7]
                        phongbanmoi = hc_name_moi[9]
                        sectiondescriptionmoi = hc_name_moi[11]
                        employeetypemoi = hc_name_moi[3]
                        positioncodemoi = hc_name_moi[4]
                        positioncodedescriptionmoi = hc_name_moi[5]
                        chucdanhtamoi = hc_name_moi[2]
                    
                    
                        
                        dieuchuyennhansu(masothe,loaidieuchuyen,chucdanhcu,chucdanhmoi,
                                         chuyencu, chuyenmoi,capbaccu,capbacmoi,
                                         sectioncodecu,sectioncodemoi,hccategorycu,hccategorymoi,
                                         phongbancu,phongbanmoi,sectiondescriptioncu,sectiondescriptionmoi,
                                         employeetypecu,employeetypemoi,positioncodedescriptioncu,positioncodedescriptionmoi,
                                         positioncodecu, positioncodemoi,chucdanhtacu,chucdanhtamoi,ngay,ghichu)
                        
                    elif loaidieuchuyen == "Nghỉ việc":
                        thongtin_laodong = laydanhsachtheomst(masothe)[0]
                        chucdanhcu = thongtin_laodong["Job title VN"]
                        chuyencu = thongtin_laodong["Line"]
                        capbaccu = thongtin_laodong["Gradecode"]
                        hccategorycu = thongtin_laodong["HC category"]
                        dichuyennghiviec(masothe,chucdanhcu,chuyencu,capbaccu,hccategorycu,ngay,ghichu)
                        
                    elif loaidieuchuyen == "Nghỉ thai sản":
                        thongtin_laodong = laydanhsachtheomst(masothe)[0]
                        chucdanhcu = thongtin_laodong["Job title VN"]
                        chuyencu = thongtin_laodong["Line"]
                        capbaccu = thongtin_laodong["Gradecode"]
                        hccategorycu = thongtin_laodong["HC category"]
                        dichuyennghithaisan(masothe,
                                            chucdanhcu,
                                            chuyencu,
                                            capbaccu,
                                            hccategorycu,
                                            ngay,
                                            ghichu)
                        
                    elif loaidieuchuyen == "Thai sản đi làm lại":
                        thongtin_laodong = laydanhsachtheomst(masothe)[0]
                        chucdanhcu = thongtin_laodong["Job title VN"]
                        chuyencu = thongtin_laodong["Line"]
                        capbaccu = thongtin_laodong["Gradecode"]
                        hccategorycu = thongtin_laodong["HC category"]
                        dichuyenthaisandilamlai(masothe,chucdanhcu,chuyencu,
                                                capbaccu,hccategorycu,ngay)
                        
                    flash("Cập nhật điều chuyển bằng file thành công !!!")
            except Exception as e:
                print(e)
                flash(f"Cập nhật điều chuyển bằng file thất bại {e} !!!")
    return redirect("/muc6_2")

@app.route("/bangcong_web", methods=["GET","POST"])
def bangcong_web():
    if request.method == "GET":
        thang = int(request.args.get("thang")) if request.args.get("thang") else 0
        nam = int(request.args.get("nam")) if request.args.get("nam") else 0
        mst = request.args.get("mst")
        bophan = request.args.get("bophan")
        chuyen = request.args.get("chuyen")
        danhsach = lay_bangcong_thucte(thang,nam,mst,bophan,chuyen)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("bangcong_web.html",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
        
    elif request.method == "POST":
        thang = request.form.get("thang")
        nam = request.form.get("nam")
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_bangcong_thucte(thang,nam,mst,bophan,chuyen)
        data = [{
            "Mã số thẻ": row[0],
            "Họ tên": row[1],
            "Bộ phận": row[2],
            "Chuyền": row[3],
            "Vị trí": row[4],
            "Chức danh": row[5],
            "Ngày vào": datetime.strptime(row[6],"%Y-%m-%d").strftime("%d/%m/%Y") if row[6] else "",
            "Ngày chính thức": datetime.strptime(row[7],"%Y-%m-%d").strftime("%d/%m/%Y") if row[7] else "",
            "Ca": row[8], 
            "01": row[9],
            "02": row[10],
            "03": row[11],
            "04": row[12],
            "05": row[13],
            "06": row[14],
            "07": row[15],
            "08": row[16],
            "09": row[17],
            "10": row[18],
            "11": row[19],
            "12": row[20],
            "13": row[21],
            "14": row[22],
            "15": row[23],
            "16": row[24],
            "17": row[25],
            "18": row[26],
            "19": row[27],
            "20": row[28],
            "21": row[29],
            "22": row[30],
            "23": row[31],
            "24": row[32],
            "25": row[33],
            "26": row[34],
            "27": row[35],
            "28": row[36],
            "29": row[37],
            "30": row[38],
            "31": row[39],
            "Thử việc": row[40],
            "Chính thức": row[41],
            "Tháng": row[42],
            "Năm": row[43],
            "Nhà máy": row[44]
        } for row in danhsach]  
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='coerce')
        df["Thử việc"] = to_numeric(df['Thử việc'], errors='coerce')
        df["Chính thức"] = to_numeric(df['Chính thức'], errors='coerce')
        df["Tháng"] = to_numeric(df['Tháng'], errors='coerce')
        df["Năm"] = to_numeric(df['Năm'], errors='coerce')
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=bangcongtong_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response   
    
@app.route("/tangcachedo_web", methods=["GET","POST"])
def tangcachedo_web():
    if request.method == "GET":
        thang = int(request.args.get("thang")) if request.args.get("thang") else 0
        nam = int(request.args.get("nam")) if request.args.get("nam") else 0
        mst = request.args.get("mst")
        bophan = request.args.get("bophan")
        chuyen = request.args.get("chuyen")
        danhsach = lay_tangcachedo_web(thang,nam,mst,bophan,chuyen)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("tangca_chedo.html",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        thang = request.form.get("thang")
        nam = request.form.get("nam")
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcachedo_web(thang,nam,mst,bophan,chuyen)
        data = [{
            "Mã số thẻ": row[0],
            "Họ tên": row[1],
            "Bộ phận": row[2],
            "Chuyền": row[3],
            "Vị trí": row[4],
            "Chức danh": row[5],
            "Ngày vào": datetime.strptime(row[6],"%Y-%m-%d").strftime("%d/%m/%Y") if row[6] else "",
            "Ngày chính thức": datetime.strptime(row[7],"%Y-%m-%d").strftime("%d/%m/%Y") if row[7] else "",
            "Ca": row[8], 
            "01": row[9],
            "02": row[10],
            "03": row[11],
            "04": row[12],
            "05": row[13],
            "06": row[14],
            "07": row[15],
            "08": row[16],
            "09": row[17],
            "10": row[18],
            "11": row[19],
            "12": row[20],
            "13": row[21],
            "14": row[22],
            "15": row[23],
            "16": row[24],
            "17": row[25],
            "18": row[26],
            "19": row[27],
            "20": row[28],
            "21": row[29],
            "22": row[30],
            "23": row[31],
            "24": row[32],
            "25": row[33],
            "26": row[34],
            "27": row[35],
            "28": row[36],
            "29": row[37],
            "30": row[38],
            "31": row[39],
            "Tổng": row[40],
            "Tháng": row[41],
            "Năm": row[42],
            "Nhà máy": row[43]
        } for row in danhsach]  
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='coerce')
        df["Tháng"] = to_numeric(df['Tháng'], errors='coerce')
        df["Năm"] = to_numeric(df['Năm'], errors='coerce')
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=bangtangcachedo_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response   
    
@app.route("/tangcangay_web", methods=["GET","POST"])
def tangcangay_web():
    if request.method == "GET":
        thang = int(request.args.get("thang")) if request.args.get("thang") else 0
        nam = int(request.args.get("nam")) if request.args.get("nam") else 0
        mst = request.args.get("mst")
        bophan = request.args.get("bophan")
        chuyen = request.args.get("chuyen")
        danhsach = lay_tangcangay_web(thang,nam,mst,bophan,chuyen)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("tangca_ngay.html",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        thang = request.form.get("thang")
        nam = request.form.get("nam")
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcangay_web(thang,nam,mst,bophan,chuyen)
        data = [{
            "Mã số thẻ": row[0],
            "Họ tên": row[1],
            "Bộ phận": row[2],
            "Chuyền": row[3],
            "Vị trí": row[4],
            "Chức danh": row[5],
            "Ngày vào": datetime.strptime(row[6],"%Y-%m-%d").strftime("%d/%m/%Y") if row[6] else "",
            "Ngày chính thức": datetime.strptime(row[7],"%Y-%m-%d").strftime("%d/%m/%Y") if row[7] else "",
            "Ca": row[8], 
            "01": row[9],
            "02": row[10],
            "03": row[11],
            "04": row[12],
            "05": row[13],
            "06": row[14],
            "07": row[15],
            "08": row[16],
            "09": row[17],
            "10": row[18],
            "11": row[19],
            "12": row[20],
            "13": row[21],
            "14": row[22],
            "15": row[23],
            "16": row[24],
            "17": row[25],
            "18": row[26],
            "19": row[27],
            "20": row[28],
            "21": row[29],
            "22": row[30],
            "23": row[31],
            "24": row[32],
            "25": row[33],
            "26": row[34],
            "27": row[35],
            "28": row[36],
            "29": row[37],
            "30": row[38],
            "31": row[39],
            "Tổng": row[40],
            "Tháng": row[41],
            "Năm": row[42],
            "Nhà máy": row[43]
        } for row in danhsach]  
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='coerce')
        df["Tháng"] = to_numeric(df['Tháng'], errors='coerce')
        df["Năm"] = to_numeric(df['Năm'], errors='coerce')
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=bangtangcangay_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response   

@app.route("/tangcadem_web", methods=["GET","POST"])
def tangcadem_web():
    if request.method == "GET":
        thang = int(request.args.get("thang")) if request.args.get("thang") else 0
        nam = int(request.args.get("nam")) if request.args.get("nam") else 0
        mst = request.args.get("mst")
        bophan = request.args.get("bophan")
        chuyen = request.args.get("chuyen")
        danhsach = lay_tangcadem_web(thang,nam,mst,bophan,chuyen)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("tangca_dem.html",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        thang = request.form.get("thang")
        nam = request.form.get("nam")
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcadem_web(thang,nam,mst,bophan,chuyen)
        data = [{
            "Mã số thẻ": row[0],
            "Họ tên": row[1],
            "Bộ phận": row[2],
            "Chuyền": row[3],
            "Vị trí": row[4],
            "Chức danh": row[5],
            "Ngày vào": datetime.strptime(row[6],"%Y-%m-%d").strftime("%d/%m/%Y") if row[6] else "",
            "Ngày chính thức": datetime.strptime(row[7],"%Y-%m-%d").strftime("%d/%m/%Y") if row[7] else "",
            "Ca": row[8], 
            "01": row[9],
            "02": row[10],
            "03": row[11],
            "04": row[12],
            "05": row[13],
            "06": row[14],
            "07": row[15],
            "08": row[16],
            "09": row[17],
            "10": row[18],
            "11": row[19],
            "12": row[20],
            "13": row[21],
            "14": row[22],
            "15": row[23],
            "16": row[24],
            "17": row[25],
            "18": row[26],
            "19": row[27],
            "20": row[28],
            "21": row[29],
            "22": row[30],
            "23": row[31],
            "24": row[32],
            "25": row[33],
            "26": row[34],
            "27": row[35],
            "28": row[36],
            "29": row[37],
            "30": row[38],
            "31": row[39],
            "Tổng": row[40],
            "Tháng": row[41],
            "Năm": row[42],
            "Nhà máy": row[43]
        } for row in danhsach]  
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='coerce')
        df["Tháng"] = to_numeric(df['Tháng'], errors='coerce')
        df["Năm"] = to_numeric(df['Năm'], errors='coerce')
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=bangtangcadem_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response
    
@app.route("/tangca_ngayle_web", methods=["GET","POST"])
def tangca_ngayle_web():
    if request.method == "GET":
        thang = int(request.args.get("thang")) if request.args.get("thang") else 0
        nam = int(request.args.get("nam")) if request.args.get("nam") else 0
        mst = request.args.get("mst")
        bophan = request.args.get("bophan")
        chuyen = request.args.get("chuyen")
        danhsach = lay_tangcangayle_web(thang,nam,mst,bophan,chuyen)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("tangca_ngayle.html",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        thang = request.form.get("thang")
        nam = request.form.get("nam")
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcangayle_web(thang,nam,mst,bophan,chuyen)
        data = [{
            "Mã số thẻ": row[0],
            "Họ tên": row[1],
            "Bộ phận": row[2],
            "Chuyền": row[3],
            "Vị trí": row[4],
            "Chức danh": row[5],
            "Ngày vào": datetime.strptime(row[6],"%Y-%m-%d").strftime("%d/%m/%Y") if row[6] else "",
            "Ngày chính thức": datetime.strptime(row[7],"%Y-%m-%d").strftime("%d/%m/%Y") if row[7] else "",
            "Ca": row[8], 
            "01": row[9],
            "02": row[10],
            "03": row[11],
            "04": row[12],
            "05": row[13],
            "06": row[14],
            "07": row[15],
            "08": row[16],
            "09": row[17],
            "10": row[18],
            "11": row[19],
            "12": row[20],
            "13": row[21],
            "14": row[22],
            "15": row[23],
            "16": row[24],
            "17": row[25],
            "18": row[26],
            "19": row[27],
            "20": row[28],
            "21": row[29],
            "22": row[30],
            "23": row[31],
            "24": row[32],
            "25": row[33],
            "26": row[34],
            "27": row[35],
            "28": row[36],
            "29": row[37],
            "30": row[38],
            "31": row[39],
            "Tổng": row[40],
            "Tháng": row[41],
            "Năm": row[42],
            "Nhà máy": row[43]
        } for row in danhsach]  
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='coerce')
        df["Tháng"] = to_numeric(df['Tháng'], errors='coerce')
        df["Năm"] = to_numeric(df['Năm'], errors='coerce')
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=bangtangcangayle_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response
    
@app.route("/tangca_chunhat_web", methods=["GET","POST"])
def tangca_chunhat_web():
    if request.method == "GET":
        thang = int(request.args.get("thang")) if request.args.get("thang") else 0
        nam = int(request.args.get("nam")) if request.args.get("nam") else 0
        mst = request.args.get("mst")
        bophan = request.args.get("bophan")
        chuyen = request.args.get("chuyen")
        danhsach = lay_tangcachunhat_web(thang,nam,mst,bophan,chuyen)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("tangca_chunhat.html",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        thang = request.form.get("thang")
        nam = request.form.get("nam")
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcachunhat_web(thang,nam,mst,bophan,chuyen)
        data = [{
            "Mã số thẻ": row[0],
            "Họ tên": row[1],
            "Bộ phận": row[2],
            "Chuyền": row[3],
            "Vị trí": row[4],
            "Chức danh": row[5],
            "Ngày vào": datetime.strptime(row[6],"%Y-%m-%d").strftime("%d/%m/%Y") if row[6] else "",
            "Ngày chính thức": datetime.strptime(row[7],"%Y-%m-%d").strftime("%d/%m/%Y") if row[7] else "",
            "Ca": row[8], 
            "01": row[9],
            "02": row[10],
            "03": row[11],
            "04": row[12],
            "05": row[13],
            "06": row[14],
            "07": row[15],
            "08": row[16],
            "09": row[17],
            "10": row[18],
            "11": row[19],
            "12": row[20],
            "13": row[21],
            "14": row[22],
            "15": row[23],
            "16": row[24],
            "17": row[25],
            "18": row[26],
            "19": row[27],
            "20": row[28],
            "21": row[29],
            "22": row[30],
            "23": row[31],
            "24": row[32],
            "25": row[33],
            "26": row[34],
            "27": row[35],
            "28": row[36],
            "29": row[37],
            "30": row[38],
            "31": row[39],
            "Tổng": row[40],
            "Tháng": row[41],
            "Năm": row[42],
            "Nhà máy": row[43]
        } for row in danhsach]  
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='coerce')
        df["Tháng"] = to_numeric(df['Tháng'], errors='coerce')
        df["Năm"] = to_numeric(df['Năm'], errors='coerce')
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=bangtangcachunhat_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response
    
@app.route("/chamcong_goc_web", methods=["GET","POST"])
def chamcong_goc_web():
    if request.method == "GET":
        mst = request.args.get("mst")
        chuyen = request.args.get("chuyen")
        bophan = request.args.get("bophan")
        ngay = request.args.get("ngay")
        danhsach = lay_dulieu_chamcong_web(mst,chuyen,bophan,ngay)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("dulieu_chamconggoc_web.html",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        ngay = request.form.get("ngay")
        danhsach = lay_dulieu_chamcong_web(mst,chuyen,bophan,ngay)
        data = [{
            "Mã số thẻ": row[0],
            "Họ tên": row[1],
            "Bộ phận": row[3],
            "Chuyền": row[2],
            "Chức danh": row[4],
            "Ngày vào": datetime.strptime(row[5],"%Y-%m-%d").strftime("%d/%m/%Y") if row[5] else "",
            "01": row[6],
            "02": row[7],
            "03": row[8],
            "04": row[9],
            "05": row[10],
            "06": row[11],
            "07": row[12],
            "08": row[13],
            "09": row[14],
            "10": row[15],
            "Nhà máy": row[16]
        } for row in danhsach] 
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='coerce')
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=dapthegoc_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response
        
@app.route("/bangcong5ngay_web", methods=["GET","POST"])
def bangcong5ngay_web():
    if request.method == "GET":
        masothe = request.args.get("mst")
        chuyen = request.args.get("chuyen")
        bophan = request.args.get("bophan")
        phanloai = request.args.get("phanloai")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        ngay = request.args.get("ngay")
        danhsach = lay_bangcong5ngay_web(masothe,chuyen,bophan,phanloai,ngay,tungay,denngay)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("bangcong5ngay_web.html",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        masothe = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        phanloai = request.form.get("phanloai")
        tungay = request.form.get("tungay")
        denngay = request.form.get("denngay")
        ngay = request.form.get("ngay")
        danhsach = lay_bangcong5ngay_web(masothe,chuyen,bophan,phanloai,ngay,tungay,denngay)
        data = [{
            "Nhà máy": row[0],
            "Mã số thẻ": row[1],
            "Họ tên": row[2],
            "Bộ phận": row[5],
            "Chuyền": row[4],
            "Chức danh": row[3],
            "Cấp bậc": row[6],
            "HC category": row[21],
            "Ngày": row[7],    
            "Ca": row[8],
            "Số phút ca": row[9],
            "Giờ vào": row[10],
            "Giờ ra": row[11],
            "Phút hành chính": row[12],
            "Phút nghỉ phép": row[13],
            "Phút tăng ca 100%": row[14],
            "Phút tăng ca 150%": row[15],
            "Phút tăng đêm": row[16],
            "Phút nghỉ không lương": row[17],
            "Phút nghỉ khác": row[18],
            "Loại nghỉ khác": row[19],
            "Phân loại": row[20]
        } for row in danhsach] 
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='coerce')
        df["Ngày"] = to_datetime(df['Ngày'], errors='coerce', dayfirst=True)
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.column_letter == 'I' and cell.value is not None:
                        cell.number_format = 'DD/MM/YYYY'
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=bangcong5ngay_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response
@app.route("/bangcongchot_web", methods=["GET","POST"])
def bangcongchot_web():
    if request.method == "GET":
        masothe = request.args.get("mst")
        chuyen = request.args.get("chuyen")
        bophan = request.args.get("bophan")
        phanloai = request.args.get("phanloai")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        ngay = request.args.get("ngay")
        danhsach = lay_bangcongchot_web(masothe,chuyen,bophan,phanloai,ngay,tungay,denngay)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("bangcongchot_web.html",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        masothe = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        phanloai = request.form.get("phanloai")
        tungay = request.form.get("tungay")
        denngay = request.form.get("denngay")
        ngay = request.form.get("ngay")
        danhsach = lay_bangcongchot_web(masothe,chuyen,bophan,phanloai,ngay,tungay,denngay)
        data = [{
            "Nhà máy": row[0],
            "Mã số thẻ": row[1],
            "Họ tên": row[2],
            "Bộ phận": row[5],
            "Chuyền": row[4],
            "Chức danh": row[3],
            "Cấp bậc": row[6],
            "HC category": row[20],
            "Ngày": row[7],    
            "Ca": row[8],
            "Số phút ca": row[9],
            "Giờ vào": row[10],
            "Giờ ra": row[11],
            "Phút hành chính": row[12],
            "Phút nghỉ phép": row[13],
            "Phút tăng ca 100%": row[14],
            "Phút tăng ca 150%": row[15],
            "Phút tăng đêm": row[16],
            "Phút nghỉ không lương": row[17],
            "Phút tăng nghỉ khác": row[18],
            "Phân loại": row[19]
        } for row in danhsach] 
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='coerce')
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=bangcongchot_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response
    
@app.route("/tailen_nhansu_pheduyet_tangca", methods=["POST"])
def tailen_nhansu_pheduyet_tangca():
    if request.method=="POST":
        file = request.files.get("file")
        if file:
            try:
                thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
                filepath = os.path.join(FOLDER_NHAP, f"danhsach_tangca_{thoigian}.xlsx")
                file.save(filepath)
                data = pd.read_excel(filepath ).to_dict(orient="records")
                for row in data:
                    id = row["ID"]
                    hrpheduyet = row["HR phê duyệt"] if not pd.isna(row["HR phê duyệt"]) else ""
                    if hr_pheduyet_tangca(id,hrpheduyet):
                        flash("Nhân sự phê duyệt tăng ca ID thành công !!!")
                    else:
                        flash("Nhân sự phê duyệt tăng ca ID thất bại !!!")        
            except Exception as e:
                print(e)
            finally:         
                return redirect("/dangki_tangca_web")
            
@app.route("/laybangcalamviec", methods=["POST"])
def laybangcalamviec():
    if request.method == "POST":
        bangcalamviec = lay_cacca_theobang()
        data = [{
            "Tên ca": row[0],
            "Giờ vào hành chính": row[1],
            "Giờ ra hành chính": row[2],
            "Giờ bắt đầu ăn trưa": row[3],
            "Giờ kết thúc ăn trưa": row[4],
            "Giờ tăng ca 100%": row[5]            
        } for row in bangcalamviec]
        return jsonify({"data": data})
    
@app.route("/thaydoi_ngaybatdau_lichsu_congviec", methods=["POST"])
def thaydoi_ngaybatdau_lichsu_congviec():
    if request.method == "POST":
        id = request.form.get("id")
        ngaybatdau = request.form.get("ngaybatdau")
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        if sua_ngaybatdau_lichsu_congviec(id,ngaybatdau):
            flash(f"Sửa ngày bắt đầu cho dòng lịch sử công việc số {id} sang {ngaybatdau} thành công")
        else:
            flash(f"Sửa ngày bắt đầu cho dòng lịch sử công việc số {id} sang {ngaybatdau} thất bại")
        return redirect(f"/muc6_3?mst={mst}&chuyen={chuyen}&bophan={bophan}")
    
@app.route("/thaydoi_ngayketthuc_lichsu_congviec", methods=["POST"])
def thaydoi_ngayketthuc_lichsu_congviec():
    if request.method == "POST":
        id = request.form.get("id")
        ngayketthuc = request.form.get("ngayketthuc")
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        if sua_ngayketthuc_lichsu_congviec(id,ngayketthuc):
            flash(f"Sửa ngày bắt đầu cho dòng lịch sử công việc số {id} sang {ngayketthuc} thành công")
        else:
            flash(f"Sửa ngày bắt đầu cho dòng lịch sử công việc số {id} sang {ngayketthuc} thất bại")
        return redirect(f"/muc6_3?mst={mst}&chuyen={chuyen}&bophan={bophan}")
    
@app.route("/lay_thongtin_vitri", methods=["POST"])
def lay_thongtin_vitri():
    try:
        vitri = request.args.get("vitri")
        return jsonify({"data": get_thongtin_vitri(vitri)})
    except Exception as e:
        print(e)
        flash(f"Lấy thông tin vị trí thất bại ({e})!!!")
        return jsonify({"data": []})


@app.route("/tailenjd", methods=["POST"])
def tailenjd():
    if request.method == "POST":
        try:
            file = request.files.get("file")
            if file:
                # print(file)
                vitri_en = request.form.get("jd_vitrien")
                # print(vitri_en)
                path = os.path.join(FOLDER_JD, f"{vitri_en}.pdf")
                # print(path)
                file.save(path)
            return redirect("/muc2_2")
        except Exception as e:
            print(e)
            return redirect("/muc2_2")

@app.route("/bangcong_tong_web", methods=["GET","POST"])
def bangcong_tong_web():
    if request.method == "GET":
        try:
            thang = int(request.args.get("thang")) if request.args.get("thang") else 0
            nam = int(request.args.get("nam")) if request.args.get("nam") else 0
            mst = request.args.get("mst")
            bophan = request.args.get("bophan")
            chuyen = request.args.get("chuyen")
            danhsach = lay_bangcongthang_web(mst,bophan,chuyen,thang,nam)
            total = len(danhsach)
            page = request.args.get(get_page_parameter(), type=int, default=1)
            per_page = 15
            start = (page - 1) * per_page
            end = start + per_page
            paginated_rows = danhsach[start:end]
            pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
            return render_template("bangcong_thang_web.html",
                                    danhsach=paginated_rows, 
                                    pagination=pagination,
                                    count=total)
        except Exception as e:
            print(e)
            return render_template("bangcong_thang_web.html",
                                    danhsach=[])
    else:
        thang = int(request.form.get("thang")) if request.args.get("thang") else 0
        nam = int(request.form.get("nam")) if request.args.get("nam") else 0
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_bangcongthang_web(mst,bophan,chuyen,thang,nam)
        data = [{
            "Mã số thẻ": row[0],
            "Họ tên": row[1],
            "Bộ phận": row[2],
            "Chuyền": row[3],
            "Vị trí": row[4],
            "Chức danh": row[5],
            "Ngày vào": datetime.strptime(row[6], "%Y-%m-%d").strftime("%d/%m/%Y") if row[6] else "",
            "Ngày chính thức": datetime.strptime(row[7], "%Y-%m-%d").strftime("%d/%m/%Y") if row[7] else "",
            "Ca": row[8],    
            "Công thử việc": row[9],
            "Công chính thức": row[10],
            "Tăng ca chế độ thử việc": row[11],
            "Tăng ca chế độ chính thức": row[12],
            "Tăng ca ngày thử việc": row[13],
            "Tăng ca ngày chính thức": row[14],
            "Tăng ca đêm thử việc": row[15],
            "Tăng ca đêm chính thức": row[16],
            "Tăng ca chủ nhật thử việc": row[17],
            "Tăng ca chủ nhật chính thức": row[18],
            "Tăng ca ngày lễ thử việc": row[19],
            "Tăng ca ngày lễ chính thức": row[20],
            "Tuân thủ nội quy": row[21],
            "Số lần nghỉ không lương": row[22],
            "Nghỉ tự do (UA)": row[23],
            "Số giờ UP": row[24],
            "Nghỉ không lương (UP)": row[25],
            "Nghỉ không lương không ảnh hưởng TTNQ (UP01,CL)": row[26],
            "Nghỉ phép(AL)": row[27],
            "Nghỉ phép không ảnh hưởng TTNQ(AL01)": row[28],
            "Nghỉ hưởng lương thử việc": row[29],
            "Nghỉ hưởng lương chính thức": row[30],
            "Nghỉ tai nạn lao động(OCL)": row[31],
            "Nghỉ ốm, con ốm(SL)": row[32],
            "Công tác(BL)": row[33],
            "Khám thai(ML03)": row[34],
            "Nghỉ vợ sinh(ML02)": row[35],
            "Nghỉ thai sản(LML)": row[36],
            "Nghỉ việc(OSL)": row[37],
            "Tổng cộng": row[38],
            "Số biên bản kỷ luật": row[39]           
        } for row in danhsach] 
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='coerce')
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = openpyxl.load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=bangcongthang_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response
    
@app.route("/gd_pheduyet_tuyendung", methods=["POST"])
def gd_pheduyet_tuyendung():
    if request.method == "POST":
        try:
            id = request.form.get("id")
            ketqua = capnhat_trangthai_yeucau_tuyendung(id,"Phê duyệt")
            if ketqua["ketqua"]:
                flash("Cập nhật trạng thái yêu cầu tuyển dụng thành công !!!")
            else:
                flash(f"Cập nhật trạng thái yêu cầu tuyển dụng thất bại ({ketqua["lido"]})!!!")
            return redirect("/muc2_2")
        except Exception as e:
            flash(f"Lỗi cập nhật trạng thái: {e}")
            return redirect("/muc2_2")
        
@app.route("/gd_tuchoi_tuyendung", methods=["POST"])
def gd_tuchoi_tuyendung():
    if request.method == "POST":
        try:
            id = request.form.get("id")
            ketqua = capnhat_trangthai_yeucau_tuyendung(id,"Từ chối")
            if ketqua["ketqua"]:
                flash("Cập nhật trạng thái yêu cầu tuyển dụng thành công !!!")
            else:
                flash(f"Cập nhật trạng thái yêu cầu tuyển dụng thất bại ({ketqua["lido"]})!!!")
            return redirect("/muc2_2")
        except Exception as e:
            flash(f"Lỗi cập nhật trạng thái: {e}")
            return redirect("/muc2_2")
            