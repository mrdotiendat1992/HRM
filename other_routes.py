from main_routes import *

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
    
    conn = pyodbc.connect(url_database_pyodbc)
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
    mstthuky = request.args.get("mstthuky")
    mst = request.form.get("mst")
    hoten = request.form.get("hoten")
    chucvu = request.form.get("chucvu")
    chuyen = request.form.get("chuyen")
    bophan = request.form.get("bophan")
    ngay = request.form.get("ngaynghi")
    lydo = request.form.get("lydo")
    trangthai = request.form.get("trangthai")
    danhsach = laydanhsachxinnghiphep(mst,hoten,chucvu,chuyen,bophan,ngay,lydo,trangthai,mstquanly,mstthuky)
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
    # Function to convert column letter to index
    def column_letter_to_index(letter):
        return openpyxl.utils.column_index_from_string(letter)
    sdt = request.form.get("sdt")
    cccd = request.form.get("cccd")
    ngaygui = request.form.get("ngaygui")
    rows = laydanhsachdangkytuyendung(sdt, cccd, ngaygui)   
    df = pd.DataFrame(rows)
    
    df["Ngày sinh con 1"] = to_datetime(df['Ngày sinh con 1'])
    df["Ngày sinh con 2"] = to_datetime(df['Ngày sinh con 2'])
    df["Ngày sinh con 3"] = to_datetime(df['Ngày sinh con 3'])
    df["Ngày sinh con 4"] = to_datetime(df['Ngày sinh con 4'])
    df["Ngày sinh con 5"] = to_datetime(df['Ngày sinh con 5'])
    df["Ngày gửi"] = to_datetime(df['Ngày gửi'])
    df["Ngày cập nhật"] = to_datetime(df['Ngày cập nhật'])
    df["Ngày hẹn đi làm"] = to_datetime(df['Ngày hẹn đi làm'])
    df["Ngày nhận việc"] = to_datetime(df['Ngày nhận việc']) 
    
    output = BytesIO()
    with ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    # Điều chỉnh độ rộng cột
    output.seek(0)
    workbook = openpyxl.load_workbook(output)
    date_format = NamedStyle()
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
                if cell.column_letter in ['V','Y','AA','AC','AE','AG','AH','AJ','AK'] and cell.value is not None:
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
            "Section_description": "",
            "Chuc_vu": ""
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
        "Section_description": hcname[11] ,
        "Chuc_vu": hcname[12]    
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
        thangdangkycalamviec(mst,cacu,camoi,ngaybatdau,ngayketthuc)
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
                            thangdangkycalamviec(row['Mã số thẻ'],laycahientai(row['Mã số thẻ']),row['Ca mới'],row['Từ ngày'],row['Đến ngày'])
                    danhsach = None
        if danhsach:
            camoi = request.form.get("camoinhom")
            ngaybatdau = request.form.get("ngaybatdau")
            ngayketthuc = request.form.get("ngayketthuc")
            
            for user in danhsach:
                thangdangkycalamviec(user['MST'],laycahientai(user['MST']),camoi,ngaybatdau,ngayketthuc)
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
            page = request.form.get("page")
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
        return redirect(f"/muc7_1_6?mst={mst_filter}&page={page}")
        
@app.route("/quanly_pheduyet_xinnghikhac", methods=["POST"])
def quanlypheduyetxinnghikhac():
    if request.method == "POST":
        try:
            mst = request.form.get("mst_xinnghikhac")
            mst_filter = request.form.get("mst_filter")
            chuyen = lay_chuyen_theo_mst(mst)
            mstduyet = current_user.masothe
            pheduyet = request.form.get("pheduyet")
            page = request.form.get("page")
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
        return redirect(f"/muc7_1_6?mst={mst_filter}&page={page}")

@app.route("/nhansu_nhangiayto_xinnghikhac", methods=["POST"])
def nhansunhangiaytoxinnghikhac():
    if request.method == "POST":
        try:
            mst_filter = request.form.get("mst_filter")
            page = request.form.get("page")
            id = request.form["id"]
            nhangiayto = request.form.get("nhangiayto")
            # if mstdiemdanh==mstduyet:
            #     print(f"Bạn không thể kiểm tra cho chính mình, vui lòng liên hệ thư ký !!!")
            #     return redirect(f"/muc7_1_3?mst={mst_filter}&hoten{hoten_filter}=&chucvu={chucvu_filter}&chuyen={chuyen_filter}&bophan={bophan_filter}&loaidiemdanh={loaidiemdanh_filter}&ngay={ngay_filter}&lydo={lydo_filter}&trangthai={trangthai_filter}")
            print(current_user.macongty,current_user.masothe)
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
        return redirect(f"/muc7_1_6?mst={mst_filter}&page={page}")

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
            x=1
            for row in data:
                try:
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
                    ketquathemhd = themhopdongmoi(nhamay, mst, hoten, gioitinh, ngaysinh, thuongtru, tamtru, cccd, ngaycapcccd, capbac, loaihopdong, chucdanh, phongban, chuyen, luongcoban, phucap, ngaybatdau, ngayketthuc)
                    if ketquathemhd["ketqua"]:
                        print(f"Them HD dòng số {x} ok")
                    else:
                        flash(f"Lỗi thêm HĐ dòng số {x}, lí do {ketquathemhd["lido"]}, query: {ketquathemhd["query"]}")
                    ketquacapnhathd =  capnhatthongtinhopdong(nhamay,mst,loaihopdong,chucdanh,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc,vitrien,employeetype,posotioncode,postitioncodedescription,hccategory,sectioncode,sectiondescription)
                    if ketquacapnhathd["ketqua"]:
                        print(f"Cap nhap HD dòng số {x} ok")
                    else:
                        flash(f"Lỗi thêm HĐ dòng số {x}, lí do {ketquacapnhathd["lido"]}, query: {ketquacapnhathd["query"]}")
                except Exception as e:
                    flash(f"Lỗi dòng số {x}, lí do: {e}")
                x += 1
            flash("Cập nhật hợp đồng thành công !!!")
        except Exception as e:
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
        qr_file = "hp_diemdanhbu.png"
    elif kieu_qr=="na_diemdanhbu":
        qr_file = "na_diemdanhbu.png"
    elif kieu_qr=="hp_xinnghiphep":
        qr_file = "hp_xinnghiphep.png"
    elif kieu_qr=="na_xinnghiphep":
        qr_file = "na_xinnghiphep.png"
    elif kieu_qr=="hp_xinnghikhongluong":
        qr_file = "hp_xinnghikhongluong.png"
    elif kieu_qr=="na_xinnghikhongluong":
        qr_file = "na_xinnghikhongluong.png"    
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
        nhangiayto = "Chưa nhận"
        if them_xinnghikhac(masothe,hoten,chuyen,phongban,chucdanh,ngay,sophut,lydo,trangthai,nhangiayto):
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
        mst = request.args.get("mst")
        chuyen = request.args.getlist("chuyen")
        ngay = request.args.get("ngay") 
        pheduyet = request.args.get("pheduyet")  
        cacchuyen = laychuyen_quanly(current_user.masothe,current_user.macongty)    
        danhsach = danhsach_tangca(mst,chuyen,ngay,pheduyet)
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
    if request.method=="POST":
        cacchuyen = request.form.getlist("chuyen")
        ngay = request.form.get("ngay")
        pheduyet = request.form.get("pheduyet")
        link = f"/dangki_tangca_web?ngay={ngay}&pheduyet={pheduyet}"
        for chuyen in cacchuyen:
            link += f"&chuyen={chuyen}"
        return redirect(link)
        
@app.route("/duieu_tangca_web", methods=["GET"])
def duieu_tangca_web():
    if request.method=="GET":
        mst = request.args.get("mst")
        chuyen = request.args.getlist("chuyen")
        ngay = request.args.get("ngay") 
        pheduyet = request.args.get("pheduyet")  
        cacchuyen = laychuyen_quanly(current_user.masothe,current_user.macongty)    
        danhsach = danhsach_tangca_quakhu(mst,chuyen,ngay,pheduyet)
        count = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 100
        total = len(danhsach)
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("duieu_tangca_web.html",
                               cacchuyen=cacchuyen,
                               danhsach=paginated_rows, 
                                pagination=pagination,
                                count=count)
           
@app.route("/capnhat_tangca", methods=["POST"])
def capnhat_tangca():
    data = request.json 
    id = data.get("id")
    tangcasang = data.get("tangcasang")
    tangcasangthucte = data.get("tangcasangthucte")
    tangca = data.get("tangca")
    tangcathucte = data.get("tangcathucte")
    tangcadem = data.get("tangcadem")
    tangcademthucte = data.get("tangcademthucte")
        
    try:  
        if capnhat_tangca_thanhcong(id,tangcasang,tangcasangthucte,tangca,tangcathucte,tangcadem,tangcademthucte):
            return jsonify({"status": "Success"})
        else:
            return jsonify({"status": "Error"}), 400
    except Exception as e:   
        return jsonify({"status": "Error"}), 400
    
@app.route("/pheduyet_tangca", methods=["POST"])   
def pheduyet_tangca():
    data = request.json
    id = data.get("id")
    type = data.get("type")
    try:
        if nhansu_pheduyet_tangca(id, type):
            return jsonify({"status": "Success"})
        else:
            return jsonify({"status": "Error"}), 400
    except Exception as e:   
        return jsonify({"status": "Error"}), 400

    
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
                x=1
                for row in data:
                    try:
                        masothe = int(row['Mã số thẻ'])
                        ngaynghi = str(row['Ngày nghỉ'])[:10]
                        sophut = int(row['Tổng số phút'])
                        loainghi = row['Loại nghỉ']
                        hoten = row["Họ tên"]
                        chucdanh = row["Chức danh"]
                        chuyen = row["Chuyền"]
                        bophan = row["Bộ phận"]
                        trangthai = "Đã phê duyệt"
                        nhangiayto = "Đã nhận"
                        if them_xinnghikhac(masothe,hoten,chuyen,bophan,chucdanh,ngaynghi,sophut,loainghi,trangthai,nhangiayto):
                            flash(f"Thêm xin nghỉ khác thành công, dòng {x}")
                        else:
                            flash(f"Thêm xin nghỉ khác thất bại, dòng {x}")
                        x+=1
                    except Exception as e:
                        print(f"Loi them xin nghi khac: {e}")
                        break
            except Exception as e:
                print(e)
        return redirect("/muc7_1_6")

@app.route("/tai_danhsach_tangca", methods=["POST"])
def tai_danhsach_tangca():
    if request.method=="POST":
        mst = request.args.get("mst")
        chuyen = request.form.getlist("chuyen")
        ngay = request.form.get("ngay")
        pheduyet =  request.form.get("pheduyet")
        danhsach = danhsach_tangca(mst,chuyen,ngay,pheduyet)
        data = [x for x in danhsach]
        ngay = datetime.now().date()     
        df = DataFrame(data)
        df["Ngày"] = to_datetime(df["Ngày"],errors="coerce")
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
                    if cell.column_letter in ['H'] and cell.value is not None:
                        cell.number_format = 'DD/MM/YYYY'
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
                data = pd.read_excel(filepath, engine='openpyxl').to_dict(orient="records")
                
                conn = pyodbc.connect(url_database_pyodbc)
                cursor = conn.cursor()

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
                    if them_dangky_tangca(cursor, conn, nhamay, mst, hoten, chucdanh, chuyen, phongban, ngay, giotangcasang, giotangcasangthucte, giotangca, giotangcathucte, giotangcadem, giotangcademthucte, ca, giovao, giora, hrpheduyet):
                        flash("Thêm đăng ký tăng ca thành công !!!")
                    else:
                        flash("Thêm đăng ký tăng ca thất bại !!!")

                conn.close()    
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
                filepath = os.path.join(FOLDER_NHAP, f"danhsach_dieuchuyen_{thoigian}.xlsx")
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
                        
                        thongtin_laodong = laydanhsachtheothechamcong(masothe)[0]
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

                        khongdoica= ""
                    
                        
                        dieuchuyennhansu(masothe,loaidieuchuyen,chucdanhcu,chucdanhmoi,
                                         chuyencu, chuyenmoi,capbaccu,capbacmoi,
                                         sectioncodecu,sectioncodemoi,hccategorycu,hccategorymoi,
                                         phongbancu,phongbanmoi,sectiondescriptioncu,sectiondescriptionmoi,
                                         employeetypecu,employeetypemoi,positioncodedescriptioncu,positioncodedescriptionmoi,
                                         positioncodecu, positioncodemoi,chucdanhtacu,chucdanhtamoi,ngay,ghichu,khongdoica)
                        
                    elif loaidieuchuyen == "Nghỉ việc":
                        thongtin_laodong = laydanhsachtheothechamcong(masothe)[0]
                        chucdanhcu = thongtin_laodong["Job title VN"]
                        chuyencu = thongtin_laodong["Line"]
                        capbaccu = thongtin_laodong["Gradecode"]
                        hccategorycu = thongtin_laodong["HC category"]
                        dichuyennghiviec(masothe,chucdanhcu,chuyencu,capbaccu,hccategorycu,ngay,ghichu)
                        
                    elif loaidieuchuyen == "Nghỉ thai sản":
                        thongtin_laodong = laydanhsachtheothechamcong(masothe)[0]
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
                        thongtin_laodong = laydanhsachtheothechamcong(masothe)[0]
                        chucdanhcu = thongtin_laodong["Job title VN"]
                        chuyencu = thongtin_laodong["Line"]
                        capbaccu = thongtin_laodong["Gradecode"]
                        hccategorycu = thongtin_laodong["HC category"]
                        dichuyenthaisandilamlai(masothe,chucdanhcu,chuyencu,
                                                capbaccu,hccategorycu,ngay)
                    flash("Cập nhật điều chuyển bằng file thành công !!!")
            except Exception as e:
                flash(f"Cập nhật điều chuyển bằng file thất bại {e} !!!")
    return redirect("/muc6_2")

@app.route("/bangcong_hanhchinh_web", methods=["GET","POST"])
def bangcong_hanhchinh_web():
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
        thang = request.form.get("thang") if request.form.get("thang") else datetime.now().month
        nam = request.form.get("nam") if request.form.get("nam") else datetime.now().year
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_bangcong_thucte(thang,nam,mst,bophan,chuyen)
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_HANHCHINH)

        sheet = workbook['BẢNG CHẤM CÔNG HÀNH CHÍNH']  # Thay 'Sheet1' bằng tên sheet của bạn
        image_path = HINHANH_LOGO
        # Tạo đối tượng hình ảnh
        img = Image(image_path)
        # Điều chỉnh kích thước hình ảnh xuống 70% so với kích thước gốc
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyển ảnh: anchor vào ô A2 và điều chỉnh tọa độ di chuyển
        img.anchor = 'A1'

        # Chèn hình ảnh vào sheet
        sheet.add_image(img)

        sheet['A2'] = f'Tháng {thang} năm {nam}'

        # Xóa hàng từ hàng 7 đến hàng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # number_style = NamedStyle(name="number_style", number_format="0.00")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này          

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_hanhchinh_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_hanhchinh_{timestamp}.xlsx"), as_attachment=True)
    
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
        thang = request.form.get("thang") if request.form.get("thang") else datetime.now().month
        nam = request.form.get("nam") if request.form.get("nam") else datetime.now().year
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcachedo_web(thang,nam,mst,bophan,chuyen)
        workbook = openpyxl.load_workbook(FILE_MAU_LAMTHEMGIO_CHEDO)

        sheet = workbook['Sheet1']  # Thay 'Sheet1' bằng tên sheet của bạn
        image_path = HINHANH_LOGO
        # Tạo đối tượng hình ảnh
        img = Image(image_path)
        # Điều chỉnh kích thước hình ảnh xuống 70% so với kích thước gốc
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyển ảnh: anchor vào ô A2 và điều chỉnh tọa độ di chuyển
        img.anchor = 'A1'

        # Chèn hình ảnh vào sheet
        sheet.add_image(img)
        sheet['A2'] = f'Tháng {thang} năm {nam}'
        # Xóa hàng từ hàng 7 đến hàng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # number_style = NamedStyle(name="number_style", number_format="0.00")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này          

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_chedo_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_chedo_{timestamp}.xlsx"), as_attachment=True)   
    
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
        thang = request.form.get("thang") if request.form.get("thang") else datetime.now().month
        nam = request.form.get("nam") if request.form.get("nam") else datetime.now().year
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcangay_web(thang,nam,mst,bophan,chuyen)
        workbook = openpyxl.load_workbook(FILE_MAU_LAMTHEMGIO_BANNGAY)

        sheet = workbook['Sheet1']  # Thay 'Sheet1' bằng tên sheet của bạn
        image_path = HINHANH_LOGO
        # Tạo đối tượng hình ảnh
        img = Image(image_path)
        # Điều chỉnh kích thước hình ảnh xuống 70% so với kích thước gốc
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyển ảnh: anchor vào ô A2 và điều chỉnh tọa độ di chuyển
        img.anchor = 'A1'

        # Chèn hình ảnh vào sheet
        sheet.add_image(img)
        sheet['A2'] = f'Tháng {thang} năm {nam}'
        # Xóa hàng từ hàng 7 đến hàng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # number_style = NamedStyle(name="number_style", number_format="0.00")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này          

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_banngay_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_banngay_{timestamp}.xlsx"), as_attachment=True)

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
        thang = request.form.get("thang") if request.form.get("thang") else datetime.now().month
        nam = request.form.get("nam") if request.form.get("nam") else datetime.now().year
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcadem_web(thang,nam,mst,bophan,chuyen)
        workbook = openpyxl.load_workbook(FILE_MAU_LAMTHEMGIO_BANDEM)

        sheet = workbook['Sheet1']  # Thay 'Sheet1' bằng tên sheet của bạn
        image_path = HINHANH_LOGO
        # Tạo đối tượng hình ảnh
        img = Image(image_path)
        # Điều chỉnh kích thước hình ảnh xuống 70% so với kích thước gốc
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyển ảnh: anchor vào ô A2 và điều chỉnh tọa độ di chuyển
        img.anchor = 'A1'

        # Chèn hình ảnh vào sheet
        sheet.add_image(img)
        sheet['A2'] = f'Tháng {thang} năm {nam}'
        # Xóa hàng từ hàng 7 đến hàng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # number_style = NamedStyle(name="number_style", number_format="0.00")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này          

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_bandem_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_bandem_{timestamp}.xlsx"), as_attachment=True)
    
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
        thang = request.form.get("thang") if request.form.get("thang") else datetime.now().month
        nam = request.form.get("nam") if request.form.get("nam") else datetime.now().year
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcangayle_web(thang,nam,mst,bophan,chuyen)
        workbook = openpyxl.load_workbook(FILE_MAU_LAMTHEMGIO_NGAYLE)

        sheet = workbook['Sheet1']  # Thay 'Sheet1' bằng tên sheet của bạn
        image_path = HINHANH_LOGO
        # Tạo đối tượng hình ảnh
        img = Image(image_path)
        # Điều chỉnh kích thước hình ảnh xuống 70% so với kích thước gốc
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyển ảnh: anchor vào ô A2 và điều chỉnh tọa độ di chuyển
        img.anchor = 'A1'

        # Chèn hình ảnh vào sheet
        sheet.add_image(img)
        sheet['A2'] = f'Tháng {thang} năm {nam}'
        # Xóa hàng từ hàng 7 đến hàng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # number_style = NamedStyle(name="number_style", number_format="0.00")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này          

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_ngayle_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_ngayle_{timestamp}.xlsx"), as_attachment=True)
    
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
        thang = request.form.get("thang") if request.form.get("thang") else datetime.now().month
        nam = request.form.get("nam") if request.form.get("nam") else datetime.now().year
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcachunhat_web(thang,nam,mst,bophan,chuyen)
        workbook = openpyxl.load_workbook(FILE_MAU_LAMTHEMGIO_CHUNHAT)

        sheet = workbook['Sheet1']  # Thay 'Sheet1' bằng tên sheet của bạn
        image_path = HINHANH_LOGO
        # Tạo đối tượng hình ảnh
        img = Image(image_path)
        # Điều chỉnh kích thước hình ảnh xuống 70% so với kích thước gốc
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyển ảnh: anchor vào ô A2 và điều chỉnh tọa độ di chuyển
        img.anchor = 'A1'

        # Chèn hình ảnh vào sheet
        sheet.add_image(img)
        sheet['A2'] = f'Tháng {thang} năm {nam}'
        # Xóa hàng từ hàng 7 đến hàng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # number_style = NamedStyle(name="number_style", number_format="0.00")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này          

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_chunhat_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_chunhat_{timestamp}.xlsx"), as_attachment=True)
    
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
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'])
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
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_CHUACHOT)

        sheet = workbook['Sheet1']  # Thay 'Sheet1' bằng tên sheet của bạn
        image_path = HINHANH_LOGO
        # Tạo đối tượng hình ảnh
        img = Image(image_path)
        # Điều chỉnh kích thước hình ảnh xuống 70% so với kích thước gốc
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyển ảnh: anchor vào ô A2 và điều chỉnh tọa độ di chuyển
        img.anchor = 'A1'
        
        # Chèn hình ảnh vào sheet
        sheet.add_image(img)

        # Xóa hàng từ hàng 7 đến hàng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        number_style = NamedStyle(name="number_style", number_format="0")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            # for col in ['J','M','N', 'O','P', 'Q','R', 'S','U']:
            #     cell = sheet[f"{col}{row}"]
            #     if cell.value and int(cell.value) > 0:
            #         try:
            #             cell.style = number_style
            #         except ValueError:
            #             pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chuachot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chuachot_{timestamp}.xlsx"), as_attachment=True)
    
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
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_CHOT)

        sheet = workbook['Sheet1']  # Thay 'Sheet1' bằng tên sheet của bạn
        image_path = HINHANH_LOGO
        # Tạo đối tượng hình ảnh
        img = Image(image_path)
        # Điều chỉnh kích thước hình ảnh xuống 70% so với kích thước gốc
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyển ảnh: anchor vào ô A2 và điều chỉnh tọa độ di chuyển
        img.anchor = 'A1'

        # Chèn hình ảnh vào sheet
        sheet.add_image(img)

        # Xóa hàng từ hàng 7 đến hàng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            data[7] = datetime.strptime(data[7],"%Y-%m-%d")
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        number_style = NamedStyle(name="number_style", number_format="0")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            # for col in ['J','M','N', 'O','P', 'Q','R', 'S','U']:
            #     cell = sheet[f"{col}{row}"]
            #     if cell.value and int(cell.value) > 0: 
            #         try:
            #             cell.style = number_style
            #         except ValueError:
            #             pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chot_{timestamp}.xlsx"), as_attachment=True)

@app.route("/bangcongchotquakhu_web", methods=["GET","POST"])
def bangcongchotquakhu_web():
    if request.method == "GET":
        masothe = request.args.get("mst")
        chuyen = request.args.get("chuyen")
        bophan = request.args.get("bophan")
        phanloai = request.args.get("phanloai")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        ngay = request.args.get("ngay")
        danhsach = lay_bangcongchotquakhu_web(masothe,chuyen,bophan,phanloai,ngay,tungay,denngay)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("bangcongchotquakhu_web.html",
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
        danhsach = lay_bangcongchotquakhu_web(masothe,chuyen,bophan,phanloai,ngay,tungay,denngay)
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_CHOT)

        sheet = workbook['Sheet1']  # Thay 'Sheet1' bằng tên sheet của bạn
        image_path = HINHANH_LOGO
        # Tạo đối tượng hình ảnh
        img = Image(image_path)
        # Điều chỉnh kích thước hình ảnh xuống 70% so với kích thước gốc
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyển ảnh: anchor vào ô A2 và điều chỉnh tọa độ di chuyển
        img.anchor = 'A1'

        # Chèn hình ảnh vào sheet
        sheet.add_image(img)

        # Xóa hàng từ hàng 7 đến hàng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            data[7] = datetime.strptime(data[7],"%Y-%m-%d")
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        number_style = NamedStyle(name="number_style", number_format="0")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            # for col in ['J','M','N', 'O','P', 'Q','R', 'S','U']:
            #     cell = sheet[f"{col}{row}"]
            #     if cell.value and int(cell.value) > 0:
            #         try:
            #             cell.style = number_style
            #         except ValueError:
            #             pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chot_{timestamp}.xlsx"), as_attachment=True)

@app.route("/tailen_nhansu_pheduyet_tangca", methods=["POST"])
@login_required
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
                        flash(f"Nhân sự phê duyệt tăng ca ID {id} thành công !!!")
                    else:
                        flash(f"Nhân sự phê duyệt tăng ca ID {id} thất bại !!!")        
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
@login_required
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
@login_required
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
@login_required
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

@app.route("/bangcong_thang_web", methods=["GET","POST"])
@login_required
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
        thang = int(request.form.get("thang")) if request.args.get("thang") else datetime.now().month
        nam = int(request.form.get("nam")) if request.args.get("nam") else datetime.now().year
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_bangcongthang_web(mst,bophan,chuyen,thang,nam)
        
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_TONGHOP)

        sheet = workbook['BẢNG CHẤM CÔNG TỔNG HỢP']  # Thay 'Sheet1' bằng tên sheet của bạn
        image_path = HINHANH_LOGO
        # Tạo đối tượng hình ảnh
        img = Image(image_path)
        # Điều chỉnh kích thước hình ảnh xuống 70% so với kích thước gốc
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyển ảnh: anchor vào ô A2 và điều chỉnh tọa độ di chuyển
        img.anchor = 'A1'

        # Chèn hình ảnh vào sheet
        sheet.add_image(img)

        sheet['A2'] = f'Tháng {thang} năm {nam}'

        # Xóa hàng từ hàng 7 đến hàng 10000
        sheet.delete_rows(6, 10000 - 6 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        number_style = NamedStyle(name="number_style", number_format="0.00")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(6, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            for col in ['J', 'K','L', 'M','N', 'O','P', 'Q','R', 'S','T', 'U', 'X','Y', 'Z','AA','AB', 'AC','AD', 'AE', 'AF','AG', 'AH','AI', 'AJ', 'AK','AL', 'AM', 'AN']:
                cell = sheet[f"{col}{row}"]
                if cell.value and int(cell.value) > 0:
                    try:
                        cell.style = number_style
                    except ValueError:
                        pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"), as_attachment=True)
    
@app.route("/bangcongtrangoai_web", methods=["GET","POST"])
@login_required
def bangcongtrangoai_web():
    if request.method == "GET":
        try:
            thang = int(request.args.get("thang")) if request.args.get("thang") else 0
            nam = int(request.args.get("nam")) if request.args.get("nam") else 0
            mst = request.args.get("mst")
            bophan = request.args.get("bophan")
            chuyen = request.args.get("chuyen")
            danhsach = lay_bangcongtrangoai_web(mst,chuyen,bophan,thang,nam)
            total = len(danhsach)
            page = request.args.get(get_page_parameter(), type=int, default=1)
            per_page = 15
            start = (page - 1) * per_page
            end = start + per_page
            paginated_rows = danhsach[start:end]
            pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
            return render_template("bangcongtrangoai_web.html",
                                    danhsach=paginated_rows, 
                                    pagination=pagination,
                                    count=total)
        except Exception as e:
            print(e)
            return render_template("bangcongtrangoai_web.html",
                                    danhsach=[])
    else:
        thang = int(request.form.get("thang")) if request.args.get("thang") else datetime.now().month
        nam = int(request.form.get("nam")) if request.args.get("nam") else datetime.now().year
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_bangcongtrangoai_web(mst,chuyen,bophan,thang,nam)
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_TRANGOAI)

        sheet = workbook['Sheet1']  # Thay 'Sheet1' bằng tên sheet của bạn
        image_path = HINHANH_LOGO
        # Tạo đối tượng hình ảnh
        img = Image(image_path)
        # Điều chỉnh kích thước hình ảnh xuống 70% so với kích thước gốc
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyển ảnh: anchor vào ô A2 và điều chỉnh tọa độ di chuyển
        img.anchor = 'A1'

        # Chèn hình ảnh vào sheet
        sheet.add_image(img)
        sheet['A2'] = f'Tháng {thang} năm {nam}'
        # Xóa hàng từ hàng 7 đến hàng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d")
            data[7] = datetime.strptime(data[7],"%Y-%m-%d")
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        number_style = NamedStyle(name="number_style", number_format="0")
        # Duyệt qua các ô trong khu vực G5:H10000
        for row in range(4, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['G','H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            for col in ['J','M','N', 'O','P', 'Q','R', 'S','T','U']:
                cell = sheet[f"{col}{row}"]
                if cell.value and int(cell.value) > 0:
                    try:
                        cell.style = number_style
                    except ValueError:
                        pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_trangoai_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_trangoai_{timestamp}.xlsx"), as_attachment=True)
    
@app.route("/gd_pheduyet_tuyendung", methods=["POST"])
@login_required
def gd_pheduyet_tuyendung():
    if request.method == "POST":
        try:
            id = request.form.get("id")
            mst_tbp = request.form.get("mst_tbp")
            ketqua = capnhat_trangthai_yeucau_tuyendung(id,"Phê duyệt")
            if ketqua["ketqua"]:
                flash("Cập nhật trạng thái yêu cầu tuyển dụng thành công !!!")
                if them_yeucau_tuyendung_duoc_pheduyet(id):
                    flash("Gửi email tuyển dụng được phê duyệt thành công !!!")
                else:
                    flash("Gửi email tuyển dụng được phê duyệt thất bại !!!")
            else:
                flash(f"Cập nhật trạng thái yêu cầu tuyển dụng thất bại ({ketqua["lido"]})!!!")
            return redirect("/muc2_2")
        except Exception as e:
            flash(f"Lỗi cập nhật trạng thái: {e}")
            return redirect("/muc2_2")
        
@app.route("/gd_tuchoi_tuyendung", methods=["POST"])
@login_required
def gd_tuchoi_tuyendung():
    if request.method == "POST":
        try:
            id = request.form.get("id")
            ketqua = capnhat_trangthai_yeucau_tuyendung(id,"Từ chối")
            if ketqua["ketqua"]:
                flash("Cập nhật trạng thái yêu cầu tuyển dụng thành công !!!")
                if them_yeucau_tuyendung_bi_tuchoi(id):
                    flash("Gửi email tuyển dụng bị từ chối thành công !!!")
                else:
                    flash("Gửi email tuyển dụng bị từ chối thất bại !!!")
            else:
                flash(f"Cập nhật trạng thái yêu cầu tuyển dụng thất bại ({ketqua["lido"]})!!!")
            return redirect("/muc2_2")
        except Exception as e:
            flash(f"Lỗi cập nhật trạng thái: {e}")
            return redirect("/muc2_2")
            
@app.route("/td_capnhat_tuyendung", methods=["POST"])
@login_required
def td_capnhat_tuyendung():
    if request.method == "POST":
        try:   
            id = request.form.get("id")
            trangthaimoi = request.form.get("trangthai")  
            ketqua = capnhat_trangthai_tuyendung(id,trangthaimoi)
            if ketqua["ketqua"]:
                flash("Cập nhật trạng thái thực hiện tuyển dụng thành công !!!")
            else:
                flash(f"Cập nhật trạng thái thực hiện tuyển dụng thất bại ({ketqua["lido"]})!!!")
            return redirect("/muc2_2")
        except Exception as e:
            flash(f"Lỗi cập nhật trạng thái: {e}")
            return redirect("/muc2_2")

@app.route("/td_capnhat_ghichu_tuyendung", methods=["POST"])
@login_required
def td_capnhat_ghichu_tuyendung():
    if request.method == "POST":
        try:   
            id = request.form.get("id")
            ghichu = request.form.get("ghichu")  
            ketqua = capnhat_ghichu_tuyendung(id,ghichu)
            if ketqua["ketqua"]:
                flash("Cập nhật trạng thái thực hiện tuyển dụng thành công !!!")
            else:
                flash(f"Cập nhật trạng thái thực hiện tuyển dụng thất bại ({ketqua["lido"]})!!!")
            return redirect("/muc2_2")
        except Exception as e:
            flash(f"Lỗi cập nhật trạng thái: {e}")
            return redirect("/muc2_2")

@app.route("/dangky_ngayle_web", methods=["GET","POST"])
@login_required
def dangky_ngayle_web():
    if request.method == "GET":
        try:
            mst = request.args.get("mst")
            chuyen = request.args.get("chuyen")
            bophan = request.args.get("bophan")
            ngay = request.args.get("ngay")
            danhsach = lay_danhsach_dangky_ngayle(mst,chuyen,bophan,ngay)
            total = len(danhsach)
            page = request.args.get(get_page_parameter(), type=int, default=1)
            per_page = 15
            start = (page - 1) * per_page
            end = start + per_page
            paginated_rows = danhsach[start:end]
            pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
            return render_template("dangky_ngayle_web.html",
                                    danhsach=paginated_rows, 
                                    pagination=pagination,
                                    count=total)
        except Exception as e:
            flash(f"Lỗi lấy bảng đăng ký làm ngày leex: ({e})")   
            return render_template("dangky_ngayle_web.html",
                                     danhsach=[])  
    elif request.method == "POST":
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        ngay = request.form.get("ngay")
        danhsach = lay_danhsach_dangky_ngayle(mst,chuyen,bophan,ngay)
        if "HR" not in current_user.phongban:
            if danhsach:
                data = [{
                    "Nhà máy": row[0],
                    "Mã số thẻ": row[1],
                    "Họ tên": row[2],
                    "Bộ phận": row[4],
                    "Chuyền": row[3],
                    "Vị trí": row[5],
                    "Ngày đăng ký": row[6] if row[6] else ""      
                } for row in danhsach] 
            else:
                data = [{
                    "Nhà máy": "",
                    "Mã số thẻ": "",
                    "Họ tên": "",
                    "Bộ phận": "",
                    "Chuyền": "",
                    "Vị trí": "",
                    "Ngày đăng ký": ""      
                }]
        else:
            if danhsach:
                data = [{
                    "Nhà máy": row[0],
                    "Mã số thẻ": row[1],
                    "Họ tên": row[2],
                    "Bộ phận": row[4],
                    "Chuyền": row[3],
                    "Vị trí": row[5],
                    "Ngày đăng ký": row[6] if row[6] else "",    
                    "HR phê duyệt": row[7],
                    "Công khai": row[8],
                    "ID":row[9]
                } for row in danhsach] 
            else:
                data = [{
                    "Nhà máy": "",
                    "Mã số thẻ": "",
                    "Họ tên": "",
                    "Bộ phận": "",
                    "Chuyền": "",
                    "Vị trí": "",
                    "Ngày đăng ký": "",
                    "HR phê duyệt": "",
                    "Công khai": "",
                    "ID":""
                }]
                
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'])
        df["Ngày đăng ký"] = to_datetime(df['Ngày đăng ký'])
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
                    if cell.column_letter in ['G'] and cell.value is not None:
                        cell.number_format = 'DD/MM/YYYY'
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
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=dangkylamngayle_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response   
            
@app.route("/dangky_chunhat_web", methods=["GET","POST"])
@login_required
def dangky_chunhat_web():
    if request.method == "GET":
        try:
            mst = request.args.get("mst")
            chuyen = request.args.get("chuyen")
            bophan = request.args.get("bophan")
            ngay = request.args.get("ngay")
            danhsach = lay_danhsach_dangky_chunhat(mst, chuyen, bophan, ngay)
            total = len(danhsach)
            page = request.args.get(get_page_parameter(), type=int, default=1)
            per_page = 15
            start = (page - 1) * per_page
            end = start + per_page
            paginated_rows = danhsach[start:end]
            pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
            return render_template("dangky_chunhat_web.html",
                                    danhsach=paginated_rows, 
                                    pagination=pagination,
                                    count=total)
        except Exception as e:
            flash(f"Lỗi lấy bảng đăng ký làm ngày lễ: ({e})")   
            return render_template("dangky_chunhat_web.html",
                                     danhsach=[]) 
    elif request.method == "POST":
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        ngay = request.form.get("ngay")
        danhsach = lay_danhsach_dangky_chunhat(mst, chuyen, bophan, ngay)
        if "HR" not in current_user.phongban:
            if danhsach:
                data = [{
                    "Nhà máy": row[0],
                    "Mã số thẻ": row[1],
                    "Họ tên": row[2],
                    "Bộ phận": row[4],
                    "Chuyền": row[3],
                    "Vị trí": row[5],
                    "Ngày đăng ký": row[6] if row[6] else ""      
                } for row in danhsach] 
            else:
                data = [{
                    "Nhà máy": "",
                    "Mã số thẻ": "",
                    "Họ tên": "",
                    "Bộ phận": "",
                    "Chuyền": "",
                    "Vị trí": "",
                    "Ngày đăng ký": ""      
                }]
        else:
            if danhsach:
                data = [{
                    "Nhà máy": row[0],
                    "Mã số thẻ": row[1],
                    "Họ tên": row[2],
                    "Bộ phận": row[4],
                    "Chuyền": row[3],
                    "Vị trí": row[5],
                    "Ngày đăng ký": row[6] if row[6] else "",    
                    "HR phê duyệt": row[7],
                    "Công khai": row[8],
                    "ID":row[9]
                } for row in danhsach] 
            else:
                data = [{
                    "Nhà máy": "",
                    "Mã số thẻ": "",
                    "Họ tên": "",
                    "Bộ phận": "",
                    "Chuyền": "",
                    "Vị trí": "",
                    "Ngày đăng ký": "",
                    "HR phê duyệt": "",
                    "Công khai": "",
                    "ID":""
                }]
                
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'])
        df["Ngày đăng ký"] = to_datetime(df['Ngày đăng ký'])
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
                    if cell.column_letter in ['G'] and cell.value is not None:
                        cell.number_format = 'DD/MM/YYYY'
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
        time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=dangkylamchunhat_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response  

@app.route("/dangky_dilam_ngayle", methods=["POST"])
@login_required
def dangky_dilam_ngayle():
    if request.method == "POST":
        file = request.files.get("file")
        if file:
            try:
                thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
                filepath = os.path.join(FOLDER_NHAP, f"dangki_dilam_ngayle_{thoigian}.xlsx")
                file.save(filepath)
                data = pd.read_excel(filepath ).to_dict(orient="records")
                if "HR" in current_user.phongban:
                    for row in data:
                        id = row["ID"]
                        hrpheduyet = row["HR phê duyệt"] if not pd.isna(row["HR phê duyệt"]) else ""
                        congkhai = row["Công khai"] if not pd.isna(row["Công khai"]) else ""
                        if hr_pheduyet_dilam_ngayle(id,hrpheduyet,congkhai):
                            flash(f"Nhân sự phê duyệt làm ngày lễ ID {id} thành công !!!")
                        else:
                            flash(f"Nhân sự phê duyệt làm ngày lễ ID {id} thất bại !!!")       
                else:
                    for row in data:
                        nhamay = current_user.macongty
                        mst = row["Mã số thẻ"]
                        hoten = row["Họ tên"]
                        chuyen = row["Chuyền"]
                        bophan = row["Bộ phận"]
                        vitri = row["Vị trí"]
                        ngay = row["Ngày đăng ký"]
                        if them_dangky_dilam_ngayle(nhamay,mst,hoten,chuyen,bophan,vitri,ngay):
                            flash(f"Thêm làm ngày lễ thành công !!!")
                        else:
                            flash(f"Thêm làm ngày lễ  thất bại !!!")
            except Exception as e:
                flash(e)
        return redirect("/dangky_ngayle_web")
            
@app.route("/dangky_dilam_chunhat", methods=["POST"])
@login_required
def dangky_dilam_chunhat():
    if request.method == "POST":
        file = request.files.get("file_tailen")
        if file:
            try:
                thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
                filepath = os.path.join(FOLDER_NHAP, f"dangki_dilam_chunhat_{thoigian}.xlsx")
                file.save(filepath)
                data = pd.read_excel(filepath ).to_dict(orient="records")
                for row in data:
                    nhamay = current_user.macongty
                    mst = row["Mã số thẻ"]
                    hoten = row["Họ tên"]
                    chuyen = row["Chuyền"]
                    bophan = row["Bộ phận"]
                    vitri = row["Vị trí"]
                    ngay = row["Ngày đăng ký"]
                    print(mst,hoten, chuyen, bophan,vitri,ngay)
                    if them_dangky_dilam_chunhat(nhamay,mst,hoten,chuyen,bophan,vitri,ngay):
                        flash(f"Thêm làm Chủ nhật thành công !!!")
                    else:
                        flash(f"Thêm làm Chủ nhật  thất bại !!!")       
            except Exception as e:
                print(e)
        return redirect("/dangky_chunhat_web")

@app.route("/hr_pheduyet_dangky_dilam_chunhat", methods=["POST"])
@login_required
def hr_pheduyet_dangky_dilam_chunhat():
    if request.method == "POST":
        file = request.files.get("file_pheduyet")
        if file:
            try:
                thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
                filepath = os.path.join(FOLDER_NHAP, f"dangki_dilam_chunhat_{thoigian}.xlsx")
                file.save(filepath)
                data = pd.read_excel(filepath ).to_dict(orient="records")
                for row in data:
                    id = row["ID"]
                    hrpheduyet = row["HR phê duyệt"] if not pd.isna(row["HR phê duyệt"]) else ""
                    congkhai = row["Công khai"] if not pd.isna(row["Công khai"]) else ""
                    if hr_pheduyet_dilam_chunhat(id,hrpheduyet,congkhai):
                        flash(f"Nhân sự phê duyệt làm Chủ nhật ID {id} thành công !!!")
                    else:
                        flash(f"Nhân sự phê duyệt làm Chủ nhật ID {id} thất bại !!!")       
            except Exception as e:
                print(e)
        return redirect("/dangky_chunhat_web")
    
@app.route('/download_JD',methods=["POST"])
@login_required
def download_file():
    try:
        filename = request.form.get("filename")
        print(os.path.exists(filename))
        return send_file(filename, as_attachment=True)
    except Exception as e:
        print(e)
        return redirect("/muc2_2")
 
@app.route('/duyet_hangloat_tangca',methods=["POST"])
@login_required
def duyet_hangloat_tangca():  
    try:
        mst = request.form.get("mst") 
        chuyen = request.form.getlist("chuyen")
        ngay = request.form.get("ngay") 
        pheduyet = ""  
        danhsach = danhsach_tangca(mst,chuyen,ngay,pheduyet)
        for x in danhsach:
            flash(x['ID'],hr_pheduyet_tangca(x['ID'],"OK") )   
    except Exception as e:
        flash(f"Lỗi phê duyệt hàng loạt: {e}")
    link = f"/dangki_tangca_web?ngay={ngay}"
    for ch in chuyen:
        link+=f"&chuyen={ch}"
    return redirect(link)  

@app.route('/boduyet_hangloat_tangca',methods=["POST"])
@login_required
def boduyet_hangloat_tangca():  
    try:
        mst = request.form.get("mst")
        chuyen = request.form.getlist("chuyen")
        ngay = request.form.get("ngay") 
        pheduyet = ""  
        danhsach = danhsach_tangca(mst,chuyen,ngay,pheduyet)
        for x in danhsach:
            flash(x['ID'],hr_pheduyet_tangca(x['ID'],"") )   
    except Exception as e:
        flash(f"Lỗi bỏ phê duyệt hàng loạt: {e}")
    link = f"/dangki_tangca_web?ngay={ngay}&mst={mst}"
    for ch in chuyen:
        link+=f"&chuyen={ch}"
    return redirect(link)  

@app.route("/thaydoi_chuyen_lichsu_congviec", methods=["POST"])
@login_required
def thaydoi_chuyen_lichsu_congviec():
    if request.method == "POST":
        id = request.form.get("id")
        chuyen_filter = request.form.get("chuyen_filter")
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        if sua_chuyen_lichsu_congviec(id,chuyen):
            flash(f"Sửa Chuyền cho dòng lịch sử công việc số {id} sang {chuyen} thành công")
        else:
            flash(f"Sửa Chuyền cho dòng lịch sử công việc số {id} sang {chuyen} thất bại")
        return redirect(f"/muc6_3?mst={mst}&chuyen={chuyen_filter}&bophan={bophan}")
    
@app.route("/thaydoi_bophan_lichsu_congviec", methods=["POST"])
@login_required
def thaydoi_bophan_lichsu_congviec():
    if request.method == "POST":
        id = request.form.get("id")
        chuyen = request.form.get("chuyen")
        mst = request.form.get("mst")
        bophan_filter = request.form.get("bophan_filter")
        bophan = request.form.get("bophan")
        if sua_bophan_lichsu_congviec(id,bophan):
            flash(f"Sửa bộ phận cho dòng lịch sử công việc số {id} sang {bophan} thành công")
        else:
            flash(f"Sửa bộ phận cho dòng lịch sử công việc số {id} sang {bophan} thất bại")
        return redirect(f"/muc6_3?mst={mst}&chuyen={chuyen}&bophan={bophan_filter}")
    
@app.route("/thaydoi_chucdanh_lichsu_congviec", methods=["POST"])
@login_required
def thaydoi_chucdanh_lichsu_congviec():
    if request.method == "POST":
        id = request.form.get("id")
        chucdanh = request.form.get("chucdanh")
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        if sua_chucdanh_lichsu_congviec(id,chucdanh):
            flash(f"Sửa chức danh cho dòng lịch sử công việc số {id} sang {chuyen} thành công")
        else:
            flash(f"Sửa chức danh cho dòng lịch sử công việc số {id} sang {chuyen} thất bại")
        return redirect(f"/muc6_3?mst={mst}&chuyen={chuyen}&bophan={bophan}")
    
@app.route("/thaydoi_capbac_lichsu_congviec", methods=["POST"])
@login_required
def thaydoi_capbac_lichsu_congviec():
    if request.method == "POST":
        id = request.form.get("id")
        capbac = request.form.get("capbac")
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        if sua_capbac_lichsu_congviec(id,capbac):
            flash(f"Sửa cấp bậc cho dòng lịch sử công việc số {id} sang {chuyen} thành công")
        else:
            flash(f"Sửa cấp bậc cho dòng lịch sử công việc số {id} sang {chuyen} thất bại")
        return redirect(f"/muc6_3?mst={mst}&chuyen={chuyen}&bophan={bophan}")
    
@app.route("/thaydoi_hccategory_lichsu_congviec", methods=["POST"])
@login_required
def thaydoi_hccategory_lichsu_congviec():
    if request.method == "POST":
        id = request.form.get("id")
        hccategory = request.form.get("hccategory")
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        if sua_hccategory_lichsu_congviec(id,hccategory):
            flash(f"Sửa HC category cho dòng lịch sử công việc số {id} sang {chuyen} thành công")
        else:
            flash(f"Sửa HC category cho dòng lịch sử công việc số {id} sang {chuyen} thất bại")
        return redirect(f"/muc6_3?mst={mst}&chuyen={chuyen}&bophan={bophan}")

@app.route("/xoa_lichsu_congviec", methods=["POST"])
@login_required
def xoa_lichsu_congviec():
    if request.method == "POST":
        id = request.form.get("id")
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        if xoabo_lichsu_congviec(id):
            flash(f"Xoá dòng lịch sử công việc số {id} sang {chuyen} thành công")
        else:
            flash(f"Xoá dòng lịch sử công việc số {id} sang {chuyen} thất bại")
        return redirect(f"/muc6_3?mst={mst}&chuyen={chuyen}&bophan={bophan}")
    
@app.route("/hosonhanvien", methods=["GET"])
@login_required
def hosonhanvien():
    if request.method == "GET":
        mst = request.args.get("mst")
        nhanvien = laydanhsachtheomst(mst)
        dulieucong = lay_dulieu_tongcong(mst)
        print(dulieucong)
        if not nhanvien:
            flash(f"Không tìm thấy nhân viên có mã số thẻ là {mst}")
            return redirect("/")
        return render_template("hosonhanvien.html",nhanvien=nhanvien[0],dulieucong=dulieucong)
    
@app.route("/lay_danhsach_userhientai", methods=["POST"])
@login_required
def lay_danhsach_userhientai():
    if request.method == "POST":
        users = laydanhsachuserhientai()
        for user in users:
            user["Ngày sinh"] = datetime.strptime(user["Ngày sinh"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["Ngày sinh"]!="" else ""
            user["Ngày cấp CCCD"] = datetime.strptime(user["Ngày cấp CCCD"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["Ngày cấp CCCD"]!="" else ""
            user["Ngày ký HĐ"] = datetime.strptime(user["Ngày ký HĐ"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["Ngày ký HĐ"]!="" else ""
            user["Ngày vào"] = datetime.strptime(user["Ngày vào"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["Ngày vào"]!="" else ""
            user["Ngày nghỉ"] = datetime.strptime(user["Ngày nghỉ"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["Ngày nghỉ"]!="" else ""
            user["Ngày hết hạn"] = datetime.strptime(user["Ngày hết hạn"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["Ngày hết hạn"]!="" else ""
            user["Ngày vào nối thâm niên"] = datetime.strptime(user["Ngày vào nối thâm niên"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["Ngày vào nối thâm niên"]!="" else ""
            user["Ngày kí HĐ Thử việc"] = datetime.strptime(user["Ngày kí HĐ Thử việc"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["Ngày kí HĐ Thử việc"]!="" else ""
            user["Ngày hết hạn HĐ Thử việc"] = datetime.strptime(user["Ngày hết hạn HĐ Thử việc"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["Ngày hết hạn HĐ Thử việc"]!="" else ""
            user["Ngày hết hạn HĐ xác định thời hạn lần 1"] = datetime.strptime(user["Ngày hết hạn HĐ xác định thời hạn lần 1"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["Ngày hết hạn HĐ xác định thời hạn lần 1"]!="" else ""
            user["Ngày kí HĐ xác định thời hạn lần 1"] = datetime.strptime(user["Ngày kí HĐ xác định thời hạn lần 1"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["Ngày kí HĐ xác định thời hạn lần 1"]!="" else ""
            user["Ngày kí HĐ không thời hạn"] = datetime.strptime(user["Ngày kí HĐ không thời hạn"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["Ngày kí HĐ không thời hạn"]!="" else ""
            

        df = pd.DataFrame(users)

        df["Ngày sinh"] = to_datetime(df['Ngày sinh'])
        df["Ngày cấp CCCD"] = to_datetime(df['Ngày cấp CCCD'])
        df["Ngày ký HĐ"] = to_datetime(df['Ngày ký HĐ'])
        df["Ngày vào"] = to_datetime(df['Ngày vào'])
        df["Ngày nghỉ"] = to_datetime(df['Ngày nghỉ'])
        df["Ngày hết hạn"] = to_datetime(df['Ngày hết hạn'])
        df["Ngày vào nối thâm niên"] = to_datetime(df['Ngày vào nối thâm niên'])
        df["Ngày sinh con 1"] = to_datetime(df['Ngày sinh con 1'])
        df["Ngày sinh con 2"] = to_datetime(df['Ngày sinh con 2'])
        df["Ngày sinh con 3"] = to_datetime(df['Ngày sinh con 3'])
        df["Ngày sinh con 4"] = to_datetime(df['Ngày sinh con 4'])
        df["Ngày sinh con 5"] = to_datetime(df['Ngày sinh con 5'])
        df["Ngày kí HĐ Thử việc"] = to_datetime(df['Ngày kí HĐ Thử việc'])
        df["Ngày hết hạn HĐ Thử việc"] = to_datetime(df['Ngày hết hạn HĐ Thử việc'])
        df["Ngày kí HĐ xác định thời hạn lần 1"] = to_datetime(df['Ngày kí HĐ xác định thời hạn lần 1'])
        df["Ngày hết hạn HĐ xác định thời hạn lần 1"] = to_datetime(df['Ngày hết hạn HĐ xác định thời hạn lần 1'])
        df["Ngày kí HĐ không thời hạn"] = to_datetime(df['Ngày kí HĐ không thời hạn'])
        
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
                    if cell.column_letter in ['E','H','AB','AD','AF','AF','AJ','AO','AP','BG','BH','BJ','BL','BM','BM','BO','BP','BQ','BR'] and cell.value is not None:
                        cell.number_format = 'DD/MM/YYYY'
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
    
@app.route("/capnhat_chuyenmoi_lichsu_congtac", methods=["POST"])
@login_required
def capnhat_chuyenmoi_lichsu_congtac():
    if request.method == "POST":
        id = request.form.get("id")
        chuyenmoi = request.form.get("chuyenmoi")
        mst_filter = request.form.get("mst_filter")
        ketqua = thaydoi_chuyen_lichsu_congtac(id,chuyenmoi)
        if ketqua["ketqua"]:
            flash(f"Thay đổi lịch sử công tác dòng {id} chuyền thành {chuyenmoi} thành công")
        else:
            flash(f"Thay đổi lịch sử công tác dòng {id} chuyền thành {chuyenmoi} thất bại !!!\nLí do: {ketqua["lido"]}\nQuery: {ketqua["query"]}")
        return redirect(f"/muc6_2?mst={mst_filter}")
    
@app.route("/capnhat_vitrimoi_lichsu_congtac", methods=["POST"])
@login_required
def capnhat_vitrimoi_lichsu_congtac():
    if request.method == "POST":
        id = request.form.get("id")
        vitrimoi = request.form.get("vitrimoi")
        mst_filter = request.form.get("mst_filter")
        ketqua = thaydoi_vitri_lichsu_congtac(id,vitrimoi)
        if ketqua["ketqua"]:
            flash(f"Thay đổi lịch sử công tác dòng {id} vị trí thành {vitrimoi} thành công")
        else:
            flash(f"Thay đổi lịch sử công tác dòng {id} vị trí thành {vitrimoi} thất bại !!!\nLí do: {ketqua["lido"]}\nQuery: {ketqua["query"]}")
        return redirect(f"/muc6_2?mst={mst_filter}")
    
@app.route("/capnhat_phanloaimoi_lichsu_congtac", methods=["POST"])
@login_required
def capnhat_phanloaimoi_lichsu_congtac():
    if request.method == "POST":
        id = request.form.get("id")
        phanloaimoi = request.form.get("phanloaimoi")
        mst_filter = request.form.get("mst_filter")
        ketqua = thaydoi_phanloai_lichsu_congtac(id,phanloaimoi)
        if ketqua["ketqua"]:
            flash(f"Thay đổi lịch sử công tác dòng {id} phân loại thành {phanloaimoi} thành công")
        else:
            flash(f"Thay đổi lịch sử công tác dòng {id} phân loại thành {phanloaimoi} thất bại !!!\nLí do: {ketqua["lido"]}\nQuery: {ketqua["query"]}")
        return redirect(f"/muc6_2?mst={mst_filter}")
    
@app.route("/capnhat_ngaythuchienmoi_lichsu_congtac", methods=["POST"])
@login_required
def capnhat_ngaythuchienmoi_lichsu_congtac():
    if request.method == "POST":
        id = request.form.get("id")
        ngaythuchienmoi = request.form.get("ngaythuchienmoi")
        mst_filter = request.form.get("mst_filter")
        ketqua = thaydoi_ngaythuchien_lichsu_congtac(id,ngaythuchienmoi)
        if ketqua["ketqua"]:
            flash(f"Thay đổi lịch sử công tác dòng {id} ngày thực hiện thành {ngaythuchienmoi} thành công")
        else:
            flash(f"Thay đổi lịch sử công tác dòng {id} ngày thực hiện thành {ngaythuchienmoi} thất bại !!!\nLí do: {ketqua["lido"]}\nQuery: {ketqua["query"]}")
        return redirect(f"/muc6_2?mst={mst_filter}")
    
@app.route("/capnhat_ghichumoi_lichsu_congtac", methods=["POST"])
@login_required
def capnhat_ghichumoi_lichsu_congtac():
    if request.method == "POST":
        id = request.form.get("id")
        ghichumoi = request.form.get("ghichumoi")
        mst_filter = request.form.get("mst_filter")
        ketqua = thaydoi_ghichu_lichsu_congtac(id,ghichumoi)
        if ketqua["ketqua"]:
            flash(f"Thay đổi lịch sử công tác dòng {id} ghi chú thành {ghichumoi} thành công")
        else:
            flash(f"Thay đổi lịch sử công tác dòng {id} ghi chú thành {ghichumoi} thất bại !!!\nLí do: {ketqua["lido"]}\nQuery: {ketqua["query"]}")
        return redirect(f"/muc6_2?mst={mst_filter}")
    
@app.route("/xoa_lichsu_congtac", methods=["POST"])
@login_required
def xoa_lichsu_congtac():
    if request.method == "POST":
        id = request.form.get("id")
        mst_filter = request.form.get("mst_filter")
        ketqua = xoabo_lichsu_congtac(id)
        if ketqua["ketqua"]:
            flash(f"Xoá lịch sử công tác dòng {id} thành công")
        else:
            flash(f"Xoá lịch sử công tác dòng {id} thất bại !!!\nLí do: {ketqua["lido"]}\nQuery: {ketqua["query"]}")
        return redirect(f"/muc6_2?mst={mst_filter}")

@app.route("/hr_pheduyet_hangloat_xinnghikhac", methods=["POST"])
@login_required
def hr_pheduyet_hangloat_xinnghikhac():
    if request.method == "POST":
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        ngaynghi = request.form.get("ngaynghi")
        loainghi = request.form.get("loainghi")
        trangthai = request.form.get("trangthai")
        nhangiayto = request.form.get("nhangiayto")
        danhsach = laydanhsachxinnghikhac(mst,chuyen,bophan,ngaynghi,loainghi,trangthai,nhangiayto)
        for dong in danhsach:
            if dong[5]=="Đã phê duyệt" or dong[5]=="Đã phê duyệt":
                nhansu_nhangiayto_xinnghikhac(dong[7])
            else:
                flash(f"{dong[7]} chưa phê duyệt")
        return redirect(f"/muc7_1_6?mst={mst}&bophan={bophan}&chuyen={chuyen}&ngaynghi={ngaynghi}&loainghi={loainghi}&trangthai={trangthai}&nhangiayto={nhangiayto}")

@app.route("/check_phanquyen", methods=["POST"])
@login_required
def check_phanquyen():
    if request.method == "POST":
        masothe = request.args.get("masothe")
        macongty= request.args.get("macongty")
        phanquyen = lay_phanquyen_hientai(macongty,masothe)
        return jsonify({"phanquyen":phanquyen})
    
@app.route("/capnhat_phanquyen", methods=["POST"])
@login_required
def capnhat_phanquyen():
    if request.method == "POST":
        masothe = request.form.get("masothe")
        macongty= request.form.get("macongty")
        phanquyen = request.form.get("phanquyenmoi")
        suadoi_phanquyen(macongty,masothe,phanquyen)
        return redirect("/admin")

@app.route("/phanquyenthuky", methods=["GET"])
@login_required
def phanquyen_thuky():
    try:
        if (current_user.macongty=='NT1' and current_user.masothe==2833) or (current_user.macongty=='NT2' and current_user.masothe==2176) or (current_user.macongty=='NT2' and current_user.masothe==1369):
            action = request.args.get("action")
            if action == "Xóa tìm kiếm":
                return redirect("/phanquyenthuky")

            mst = request.args.get("mst")
            chuyen = request.args.get("chuyen")
            mst_ql = request.args.get("mst_ql")

            filters = {
                "mst": mst,
                "chuyen_to": chuyen,
                "mst_ql": mst_ql
            }
            danhsach = laydanhsach_phanquyenthuky(filters)
            page = request.args.get(get_page_parameter(), type=int, default=1)
            per_page = 20
            total = len(danhsach)
            start = (page - 1) * per_page
            end = start + per_page
            paginated_rows = danhsach[start:end]
            pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')

            return render_template("phanquyenthuky.html", danhsach=paginated_rows, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("phanquyenthuky.html", danhsach=[])

@app.route("/add_phanquyenthuky", methods=["POST"])
@login_required
def add_phanquyenthuky():
    try:
        if (current_user.macongty=='NT1' and current_user.masothe==2833) or (current_user.macongty=='NT2' and current_user.masothe==2176) or (current_user.macongty=='NT2' and current_user.masothe==1369):
            data = request.json
            conn = pyodbc.connect(url_database_pyodbc)
            cur = conn.cursor()
            query = f"INSERT INTO Phan_quyen_thu_ky VALUES ('{current_user.macongty}', {data.get("mst", "")}, '{data.get("chuyen", "")}', {data.get("mst_ql", "")})"
            print(query)
            cur.execute(query)
            cur.commit()
            conn.close()

        return {"message": "Thêm thành công"}
    except Exception as e:
        print(e)
        return {"message": "Thêm thất bại"}

@app.route("/update_phanquyenthuky", methods=["POST"])
@login_required
def update_phanquyenthuky():
    try:
        if (current_user.macongty=='NT1' and current_user.masothe==2833) or (current_user.macongty=='NT2' and current_user.masothe==2176) or (current_user.macongty=='NT2' and current_user.masothe==1369):
            data = request.json
            conn = pyodbc.connect(url_database_pyodbc)
            cur = conn.cursor()
            query = f"UPDATE Phan_quyen_thu_ky SET MST = '{data.get("mst", "")}', Chuyen_to = '{data.get("chuyen", "")}', MST_QL = '{data.get("mst_ql", "")}' WHERE ID = {data.get("id", "")}"
            print(query)
            cur.execute(query)
            cur.commit()
            conn.close()

        return {"message": "Sửa thành công"}
    except Exception as e:
        print(e)
        return {"message": "Sửa thất bại"}

@app.route("/delete_phanquyenthuky", methods=["GET"])
@login_required
def delete_phanquyenthuky():
    try:
        if (current_user.macongty=='NT1' and current_user.masothe==2833) or (current_user.macongty=='NT2' and current_user.masothe==2176) or (current_user.macongty=='NT2' and current_user.masothe==1369):
            id = request.args.get("id")
            conn = pyodbc.connect(url_database_pyodbc)
            cur = conn.cursor()
            query = f"DELETE FROM Phan_quyen_thu_ky WHERE ID = {id}"
            cur.execute(query)
            cur.commit()
            conn.close()

        return {"message": "Xóa thành công"}
    except Exception as e:
        print(e)
        return {"message": "Xóa thất bại"}

@app.route("/thaydoi_ngaypv1_yctd", methods=["POST"])
@login_required
def thaydoi_ngaypv1_yctd():
    try:
        id = request.form.get("id")
        ngay = request.form.get("ngaypv1")
        id_yctd = request.form.get("id_yctd")
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung_chi_tiet SET Ngay_PV_lan_1='{ngay}' WHERE ID = {id}"
        cur.execute(query)
        cur.commit()
        conn.close()
        return redirect(f"/muc2_2_1?id={id_yctd}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2_1?id={id_yctd}")

@app.route("/thaydoi_ketquapv1_yctd", methods=["POST"])
@login_required
def thaydoi_ketquapv1_yctd():
    try:
        id = request.form.get("id")
        ketqua = request.form.get("ketquapv1")
        id_yctd = request.form.get("id_yctd")
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung_chi_tiet SET Ket_qua_PV_lan_1='{ketqua}' WHERE ID = {id}"
        cur.execute(query)
        cur.commit()
        conn.close()
        return redirect(f"/muc2_2_1?id={id_yctd}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2_1?id={id_yctd}")

@app.route("/upload_danhgia_uvtd_pv1", methods=["POST"])
@login_required
def upload_danhgia_uvtd_pv1():
    try:
        id = request.form.get("id")
        id_yctd = request.form.get("id_yctd")
        file = request.files.get("file")
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        save_path = os.path.join(FOLDER_DGPV,f"dgpvl1_{timestamp}.pdf")
        file.save(save_path)
        link = save_path.replace("\\","/").split("HRM/")[1]
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung_chi_tiet SET File_danh_gia_RV_lan_1=N'{link}' WHERE ID = {id}"
        cur.execute(query)
        cur.commit()
        conn.close()
        return redirect(f"/muc2_2_1?id={id_yctd}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2_1?id={id_yctd}")
    
@app.route("/thaydoi_ngaypv2_yctd", methods=["POST"])
@login_required
def thaydoi_ngaypv2_yctd():
    try:
        id = request.form.get("id")
        ngay = request.form.get("ngaypv2")
        id_yctd = request.form.get("id_yctd")
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung_chi_tiet SET Ngay_PV_lan_2='{ngay}' WHERE ID = {id}"
        cur.execute(query)
        cur.commit()
        conn.close()
        return redirect(f"/muc2_2_1?id={id_yctd}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2_1?id={id_yctd}")

@app.route("/thaydoi_ketquapv2_yctd", methods=["POST"])
@login_required
def thaydoi_ketquapv2_yctd():
    try:
        id = request.form.get("id")
        ketqua = request.form.get("ketquapv2")
        id_yctd = request.form.get("id_yctd")
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung_chi_tiet SET Ket_qua_PV_lan_2='{ketqua}' WHERE ID = {id}"
        cur.execute(query)
        cur.commit()
        conn.close()
        return redirect(f"/muc2_2_1?id={id_yctd}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2_1?id={id_yctd}")

@app.route("/upload_danhgia_uvtd_pv2", methods=["POST"])
@login_required
def upload_danhgia_uvtd_pv2():
    try:
        id = request.form.get("id")
        id_yctd = request.form.get("id_yctd")
        file = request.files.get("file")
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        save_path = os.path.join(FOLDER_DGPV,f"dgpvl2_{timestamp}.pdf")
        file.save(save_path)
        link = save_path.replace("\\","/").split("HRM/")[1]
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung_chi_tiet SET File_danh_gia_RV_lan_2=N'{link}' WHERE ID = {id}"
        cur.execute(query)
        cur.commit()
        conn.close()
        return redirect(f"/muc2_2_1?id={id_yctd}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2_1?id={id_yctd}")
    
@app.route("/thaydoi_ngaypv3_yctd", methods=["POST"])
@login_required
def thaydoi_ngaypv3_yctd():
    try:
        id = request.form.get("id")
        ngay = request.form.get("ngaypv2")
        print(ngay)
        id_yctd = request.form.get("id_yctd")
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung_chi_tiet SET Ngay_PV_lan_3='{ngay}' WHERE ID = {id}"
        cur.execute(query)
        cur.commit()
        conn.close()
        return redirect(f"/muc2_2_1?id={id_yctd}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2_1?id={id_yctd}")

@app.route("/thaydoi_ketquapv3_yctd", methods=["POST"])
@login_required
def thaydoi_ketquapv3_yctd():
    try:
        id = request.form.get("id")
        ketqua = request.form.get("ketquapv3")
        id_yctd = request.form.get("id_yctd")
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung_chi_tiet SET Ket_qua_PV_lan_3='{ketqua}' WHERE ID = {id}"
        cur.execute(query)
        cur.commit()
        conn.close()
        return redirect(f"/muc2_2_1?id={id_yctd}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2_1?id={id_yctd}")

@app.route("/upload_danhgia_uvtd_pv3", methods=["POST"])
@login_required
def upload_danhgia_uvtd_pv3():
    try:
        id = request.form.get("id")
        id_yctd = request.form.get("id_yctd")
        file = request.files.get("file")
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        save_path = os.path.join(FOLDER_DGPV,f"dgpvl3_{timestamp}.pdf")
        file.save(save_path)
        link = save_path.replace("\\","/").split("HRM/")[1]
        print(link)
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung_chi_tiet SET File_danh_gia_RV_lan_3=N'{link}' WHERE ID = {id}"
        cur.execute(query)
        cur.commit()
        conn.close()
        return redirect(f"/muc2_2_1?id={id_yctd}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2_1?id={id_yctd}")

@app.route("/thaydoi_trangthai_uvtd", methods=["POST"])
@login_required
def thaydoi_trangthai_uvtd():
    try:
        id = request.form.get("id")
        trangthai = request.form.get("trangthai")
        id_yctd = request.form.get("id_yctd")
        capnhat_trangthai_ungvien_chitiet(id,trangthai,id_yctd)
        return redirect(f"/muc2_2_1?id={id_yctd}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2_1?id={id_yctd}")


@app.route("/thaydoi_ghichu_uvtd", methods=["POST"])
@login_required
def thaydoi_ghichu_uvtd():
    try:
        id = request.form.get("id")
        ghichu = request.form.get("ghichu")
        id_yctd = request.form.get("id_yctd")
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung_chi_tiet SET Ghi_chu=N'{ghichu}' WHERE ID = {id}"
        cur.execute(query)
        cur.commit()
        conn.close()
        return redirect(f"/muc2_2_1?id={id_yctd}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2_1?id={id_yctd}")

@app.route("/xoa_uvtd", methods=["POST"])
@login_required
def xoa_uvtd():
    try:
        id = request.form.get("id")
        id_yctd = request.form.get("id_yctd")
        if xoa_tuyendung_chitiet(id,id_yctd):
            flash(f"Xóa ứng viên tuyển dụng chi tiết thành công")
        else:
            flash(f"Xóa ứng viên tuyển dụng chi tiết thất bại")
        return redirect(f"/muc2_2_1?id={id_yctd}")
    except Exception as e:
        flash(f"Xóa ứng viên tuyển dụng chi tiết thất bại {e}")
        return redirect(f"/muc2_2_1?id={id_yctd}")

@app.route("/xoa_tuyendung", methods=["POST"])
@login_required
def xoa_tuyendung():
    try:
        id = request.form.get("id")
        mst= request.form.get("mst")
        phongban = request.form.get("phongban")
        trangthaiyeucau= request.form.get("trangthaiyeucau")
        trangthaithuchien = request.form.get("trangthaithuchien")
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"DELETE Yeu_cau_tuyen_dung WHERE ID = {id}"
        cur.execute(query)
        cur.commit()
        conn.close()
        return redirect(f"/muc2_2?phongban={phongban}&trangthaiyeucau={trangthaiyeucau}&trangthaithuchien={trangthaithuchien}&mst={mst}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2?phongban={phongban}&trangthaiyeucau={trangthaiyeucau}&trangthaithuchien={trangthaithuchien}&mst={mst}")

@app.route("/mo_yeucautuyendung", methods=["POST"])
@login_required
def mo_yeucautuyendung():
    try:
        id = request.form.get("id")
        mst= request.form.get("mst")
        phongban = request.form.get("phongban")
        trangthaiyeucau= request.form.get("trangthaiyeucau")
        trangthaithuchien = request.form.get("trangthaithuchien")
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung SET Ngay_dong_yeu_cau = NULL WHERE ID = {id}"
        cur.execute(query)
        cur.commit()
        conn.close()
        return redirect(f"/muc2_2?phongban={phongban}&trangthaiyeucau={trangthaiyeucau}&trangthaithuchien={trangthaithuchien}&mst={mst}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2?phongban={phongban}&trangthaiyeucau={trangthaiyeucau}&trangthaithuchien={trangthaithuchien}&mst={mst}")

@app.route("/dong_yeucautuyendung", methods=["POST"])
@login_required
def dong_yeucautuyendung():
    try:
        id = request.form.get("id")
        mst= request.form.get("mst")
        phongban = request.form.get("phongban")
        trangthaiyeucau= request.form.get("trangthaiyeucau")
        trangthaithuchien = request.form.get("trangthaithuchien")
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"UPDATE Yeu_cau_tuyen_dung SET Ngay_dong_yeu_cau = GETDATE() WHERE ID = {id}"
        print(query)
        cur.execute(query)
        cur.commit()
        conn.close()
        return redirect(f"/muc2_2?phongban={phongban}&trangthaiyeucau={trangthaiyeucau}&trangthaithuchien={trangthaithuchien}&mst={mst}")
    except Exception as e:
        flash(e)
        return redirect(f"/muc2_2?phongban={phongban}&trangthaiyeucau={trangthaiyeucau}&trangthaithuchien={trangthaithuchien}&mst={mst}")

@app.route("/nghibukiemxuong", methods=["GET"])
@login_required
def nghibukiemxuong():
    try:
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()

        mst = request.args.get("mst")
        bo_phan = request.args.get("bo_phan")
        ngay = request.args.get("ngay")

        filters = {
            "mst": mst,
            "bo_phan": bo_phan,
            "ngay": ngay
        }

        query = f"select * from NGHI_BU_KIEM_XUONG where nha_may='{current_user.macongty}'"
        query_condition  = " and ".join([f"{key} LIKE '%{value}%'" for key,value in filters.items() if value])
        if query_condition:
            query += f" and {query_condition}"
        
        danhsach = cur.execute(query).fetchall()
        cur.commit()
        conn.close()

        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 20
        total = len(danhsach)
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]

        formatted_rows = []
        for row in paginated_rows:
            formatted_row = list(row)
            original_date = row[4]
            if original_date:
                formatted_row[4] = datetime.strptime(original_date, '%Y-%m-%d').strftime('%d/%m/%Y')
            formatted_rows.append(tuple(formatted_row))

        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')

        return render_template("nghibukiemxuong.html", danhsach=formatted_rows, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("nghibukiemxuong.html", danhsach=[])

@app.route("/update_nghibukiemxuong", methods=["POST"])
@login_required
def update_nghibukiemxuong():
    try:
        data = request.json
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"UPDATE NGHI_BU_KIEM_XUONG SET MST = '{data.get("mst", "")}', HO_TEN = N'{data.get("ho_ten", "")}', BO_PHAN = '{data.get("bo_phan", "")}', NGAY = '{data.get("ngay", "")}' WHERE ID = {data.get("id", "")}"
        cur.execute(query)
        cur.commit()
        conn.close()

        return {"message": "Sửa thành công"}
    except Exception as e:
        print(e)
        return {"message": "Sửa thất bại"}

@app.route("/delete_nghibukiemxuong", methods=["DELETE"])
@login_required
def delete_nghibukiemxuong():
    try:
        id = request.args.get("id")
        conn = pyodbc.connect(url_database_pyodbc)
        cur = conn.cursor()
        query = f"DELETE FROM NGHI_BU_KIEM_XUONG WHERE ID = {id}"
        cur.execute(query)
        cur.commit()
        conn.close()

        return {"message": "Xóa thành công"}
    except Exception as e:
        print(e)
        return {"message": "Xóa thất bại"}

@app.route("/tai_sample_nghibukiemxuong", methods=["POST"])
def tai_sample_nghibukiemxuong():
    headers = ["MST", "HO_TEN", "BO_PHAN", "NGAY"]
    
    df = pd.DataFrame(columns=headers)
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    output.seek(0)
    workbook = openpyxl.load_workbook(output)
    sheet = workbook.active

    header_fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    time_stamp = datetime.now().strftime("%d%m%Y%H%M%S")
    
    response = make_response(output.read())
    response.headers['Content-Disposition'] = f'attachment; filename=bukiemxuong_{time_stamp}.xlsx'
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response  

@app.route("/tailen_nghibukiemxuong", methods=["POST"])
def tailen_nghibukiemxuong():
    file = request.files.get("file")
    if file:
        try:
            conn = pyodbc.connect(url_database_pyodbc)
            cursor = conn.cursor()
            
            df = pd.read_excel(file)
            df["NHA_MAY"] = current_user.macongty
            
            insert_query = """
                INSERT INTO NGHI_BU_KIEM_XUONG (NHA_MAY, MST, HO_TEN, BO_PHAN, NGAY)
                VALUES (?, ?, ?, ?, ?)
            """
            data_to_insert = df[["NHA_MAY", "MST", "HO_TEN", "BO_PHAN", "NGAY"]].values.tolist()
            print(data_to_insert)
            cursor.executemany(insert_query, data_to_insert)

            conn.commit() 
            conn.close()    
        except Exception as e:
            print(e)
                
    return redirect("/nghibukiemxuong")

@app.route("/capnhat_dulieu_chamcong", methods=["POST"])
def capnhat_dulieu_chamcong():
    try:
        conn = pyodbc.connect(url_database_pyodbc)
        cursor = conn.cursor()
        cursor.execute("Exec Dong_bo_CheckInOut")
        cursor.commit()
        conn.close()
        return redirect("/chamcong_sang_web")
    except Exception as e:
        flash(f"Lỗi cập nhật dữ liệu chấm công {e}") 
        return redirect("/chamcong_sang_web")

@app.route("/them_congnhan_vao_yctd", methods=["POST"])
def them_congnhan_vao_yctd():
    try:
        id = request.form.get("id")
        id_yctd = request.form.get("id_yctd")
    
        conn = pyodbc.connect(url_database_pyodbc)
        cursor = conn.cursor()
        query = f"select Ho_ten,Kenh_tuyen_dung from Dang_ky_thong_tin where ID = '{id}'"
        print(query)
        data = cursor.execute(query).fetchone()
        print(data)
        query1 = f"select Bo_phan from Yeu_cau_tuyen_dung where ID = '{id_yctd}'"
        print(query1)
        phongban = cursor.execute(query1).fetchone()[0]
        query2 = f"""insert into Yeu_cau_tuyen_dung_chi_tiet (Ho_ten,Kenh_tuyen_dung,Trang_thai,ID_YCTD, Phong_ban)
                    values (N'{data[0]}',N'{data[1]}',N'Chưa phỏng vấn','{id_yctd}','{phongban}')
                """
        print(query2)
        cursor.execute(query2)
        cursor.commit()
        conn.close()
        return redirect(f"/muc2_2_1?id={id_yctd}")
    except Exception as e:
        flash(f"Lỗi thêm công nhân vào yêu cầu tuyển dụng chi tiết {e}") 
        return redirect(f"/muc2_2_1?id={id_yctd}")

