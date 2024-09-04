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
    # Function to convert column letter to index
    def column_letter_to_index(letter):
        return openpyxl.utils.column_index_from_string(letter)
    sdt = request.form.get("sdt")
    cccd = request.form.get("cccd")
    ngaygui = request.form.get("ngaygui")
    rows = laydanhsachdangkytuyendung(sdt, cccd, ngaygui)   
    df = pd.DataFrame(rows)
    
    df["Ngày sinh con 1"] = to_datetime(df['Ngày sinh con 1'], errors='ignore').dt.date
    df["Ngày sinh con 2"] = to_datetime(df['Ngày sinh con 2'], errors='ignore').dt.date
    df["Ngày sinh con 3"] = to_datetime(df['Ngày sinh con 3'], errors='ignore').dt.date
    df["Ngày sinh con 4"] = to_datetime(df['Ngày sinh con 4'], errors='ignore').dt.date
    df["Ngày sinh con 5"] = to_datetime(df['Ngày sinh con 5'], errors='ignore').dt.date
    df["Ngày gửi"] = to_datetime(df['Ngày gửi'], errors='ignore')
    df["Ngày cập nhật"] = to_datetime(df['Ngày cập nhật'], errors='ignore')
    df["Ngày hẹn đi làm"] = to_datetime(df['Ngày hẹn đi làm'], errors='ignore')
    df["Ngày nhận việc"] = to_datetime(df['Ngày nhận việc'], errors='ignore') 
    
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
            flash("Cập nhật hợp đồng thành công !!!")
        except Exception as e:
            print(f"Cap nhat hop dong loi: {e}")
            flash(f"Cập nhật hợp đồng lỗi: ({e}) !!!")
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
        pheduyet = request.args.get("pheduyet")  
        cacchuyen = laychuyen_quanly(current_user.masothe,current_user.macongty)    
        danhsach = danhsach_tangca(chuyen,ngay,pheduyet)
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
        chuyen_filter = request.form.getlist("chuyen")
        ngay_filter = request.form.get("ngay_filter")
        pheduyet = request.form.get("pheduyet")
        id = request.form.get("id")
        try:
            if nhansu_bopheduyet_tangca(id):
                flash(f"Bỏ phê duyệt tăng ca ID = {id}")
            else:
                flash(f"Bỏ phê duyệt tăng ca ID = {id} không được")
        except Exception as e:   
            flash(f"Loi khi bo phe duyet tang ca tang ca ({e})")
        link = f"/dangki_tangca_web?ngay={ngay_filter}&pheduyet={pheduyet}"
        for chuyen in chuyen_filter:
            link += f"&chuyen={chuyen}"
        return redirect(link)

@app.route("/pheduyet_tangca", methods=["POST"])   
def pheduyet_tangca():
    if request.method=="POST":
        chuyen_filter = request.form.getlist("chuyen")
        ngay_filter = request.form.get("ngay_filter")
        pheduyet = request.form.get("pheduyet")
        id = request.form.get("id")
        try:
            if nhansu_pheduyet_tangca(id):
                flash(f"Phê duyệt tăng ca ID = {id}")
            else:
                flash(f"Phê duyệt tăng ca ID = {id} không được")
        except Exception as e:   
            flash(f"Loi khi bo phe duyet tang ca tang ca ({e})")
        link = f"/dangki_tangca_web?ngay={ngay_filter}&pheduyet={pheduyet}"
        for chuyen in chuyen_filter:
            link += f"&chuyen={chuyen}"
        return redirect(link)
    
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
                        if them_xinnghikhac(masothe,ngaynghi,sophut,loainghi,trangthai,nhangiayto):
                            flash(f"Thêm xin nghỉ khác thành công")
                        else:
                            flash(f"Thêm xin nghỉ khác thất bại")
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
        pheduyet =  request.form.get("pheduyet")
        danhsach = danhsach_tangca(chuyen,ngay,pheduyet)
        data = [x for x in danhsach]
        # print(data)
        ngay = datetime.now().date()     
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
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
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
        df["Thử việc"] = to_numeric(df['Thử việc'], errors='ignore')
        df["Chính thức"] = to_numeric(df['Chính thức'], errors='ignore')
        df["Tháng"] = to_numeric(df['Tháng'], errors='ignore')
        df["Năm"] = to_numeric(df['Năm'], errors='ignore')
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
            "Ngày vào": row[6] if row[6] else "",
            "Ngày chính thức": row[7] if row[7] else "",
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
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
        df["Tháng"] = to_numeric(df['Tháng'], errors='ignore')
        df["Năm"] = to_numeric(df['Năm'], errors='ignore')
        df["01"] = to_numeric(df['01'], errors='ignore')
        df["02"] = to_numeric(df['02'], errors='ignore')
        df["03"] = to_numeric(df['03'], errors='ignore')
        df["04"] = to_numeric(df['04'], errors='ignore')
        df["05"] = to_numeric(df['05'], errors='ignore')
        df["06"] = to_numeric(df['06'], errors='ignore')
        df["07"] = to_numeric(df['07'], errors='ignore')
        df["08"] = to_numeric(df['08'], errors='ignore')
        df["09"] = to_numeric(df['09'], errors='ignore')
        df["10"] = to_numeric(df['10'], errors='ignore')
        df["11"] = to_numeric(df['11'], errors='ignore')
        df["12"] = to_numeric(df['12'], errors='ignore')
        df["13"] = to_numeric(df['13'], errors='ignore')
        df["14"] = to_numeric(df['14'], errors='ignore')
        df["15"] = to_numeric(df['15'], errors='ignore')
        df["16"] = to_numeric(df['16'], errors='ignore')
        df["17"] = to_numeric(df['17'], errors='ignore')
        df["18"] = to_numeric(df['18'], errors='ignore')
        df["19"] = to_numeric(df['19'], errors='ignore')
        df["20"] = to_numeric(df['20'], errors='ignore')
        df["21"] = to_numeric(df['21'], errors='ignore')
        df["22"] = to_numeric(df['22'], errors='ignore')
        df["23"] = to_numeric(df['23'], errors='ignore')
        df["24"] = to_numeric(df['24'], errors='ignore')
        df["25"] = to_numeric(df['25'], errors='ignore')
        df["26"] = to_numeric(df['26'], errors='ignore')
        df["27"] = to_numeric(df['27'], errors='ignore')
        df["28"] = to_numeric(df['28'], errors='ignore')
        df["29"] = to_numeric(df['29'], errors='ignore')
        df["30"] = to_numeric(df['30'], errors='ignore')
        df["31"] = to_numeric(df['31'], errors='ignore')
        df["Tổng"] = to_numeric(df['Tổng'], errors='ignore')
        df["Ngày vào"] = to_datetime(df['Ngày vào'], errors='ignore')
        df["Ngày chính thức"] = to_datetime(df['Ngày chính thức'], errors='ignore')
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
                    if cell.column_letter in ['G','H'] and cell.value is not None:
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
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
        df["Tháng"] = to_numeric(df['Tháng'], errors='ignore')
        df["Năm"] = to_numeric(df['Năm'], errors='ignore')
        df["01"] = to_numeric(df['01'], errors='ignore')
        df["02"] = to_numeric(df['02'], errors='ignore')
        df["03"] = to_numeric(df['03'], errors='ignore')
        df["04"] = to_numeric(df['04'], errors='ignore')
        df["05"] = to_numeric(df['05'], errors='ignore')
        df["06"] = to_numeric(df['06'], errors='ignore')
        df["07"] = to_numeric(df['07'], errors='ignore')
        df["08"] = to_numeric(df['08'], errors='ignore')
        df["09"] = to_numeric(df['09'], errors='ignore')
        df["10"] = to_numeric(df['10'], errors='ignore')
        df["11"] = to_numeric(df['11'], errors='ignore')
        df["12"] = to_numeric(df['12'], errors='ignore')
        df["13"] = to_numeric(df['13'], errors='ignore')
        df["14"] = to_numeric(df['14'], errors='ignore')
        df["15"] = to_numeric(df['15'], errors='ignore')
        df["16"] = to_numeric(df['16'], errors='ignore')
        df["17"] = to_numeric(df['17'], errors='ignore')
        df["18"] = to_numeric(df['18'], errors='ignore')
        df["19"] = to_numeric(df['19'], errors='ignore')
        df["20"] = to_numeric(df['20'], errors='ignore')
        df["21"] = to_numeric(df['21'], errors='ignore')
        df["22"] = to_numeric(df['22'], errors='ignore')
        df["23"] = to_numeric(df['23'], errors='ignore')
        df["24"] = to_numeric(df['24'], errors='ignore')
        df["25"] = to_numeric(df['25'], errors='ignore')
        df["26"] = to_numeric(df['26'], errors='ignore')
        df["27"] = to_numeric(df['27'], errors='ignore')
        df["28"] = to_numeric(df['28'], errors='ignore')
        df["29"] = to_numeric(df['29'], errors='ignore')
        df["30"] = to_numeric(df['30'], errors='ignore')
        df["31"] = to_numeric(df['31'], errors='ignore')
        df["Tổng"] = to_numeric(df['Tổng'], errors='ignore')
        df["Ngày vào"] = to_datetime(df['Ngày vào'], errors='ignore')
        df["Ngày chính thức"] = to_datetime(df['Ngày chính thức'], errors='ignore')
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
                    if cell.column_letter in ['G','H'] and cell.value is not None:
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
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
        df["Tháng"] = to_numeric(df['Tháng'], errors='ignore')
        df["Năm"] = to_numeric(df['Năm'], errors='ignore')
        df["01"] = to_numeric(df['01'], errors='ignore')
        df["02"] = to_numeric(df['02'], errors='ignore')
        df["03"] = to_numeric(df['03'], errors='ignore')
        df["04"] = to_numeric(df['04'], errors='ignore')
        df["05"] = to_numeric(df['05'], errors='ignore')
        df["06"] = to_numeric(df['06'], errors='ignore')
        df["07"] = to_numeric(df['07'], errors='ignore')
        df["08"] = to_numeric(df['08'], errors='ignore')
        df["09"] = to_numeric(df['09'], errors='ignore')
        df["10"] = to_numeric(df['10'], errors='ignore')
        df["11"] = to_numeric(df['11'], errors='ignore')
        df["12"] = to_numeric(df['12'], errors='ignore')
        df["13"] = to_numeric(df['13'], errors='ignore')
        df["14"] = to_numeric(df['14'], errors='ignore')
        df["15"] = to_numeric(df['15'], errors='ignore')
        df["16"] = to_numeric(df['16'], errors='ignore')
        df["17"] = to_numeric(df['17'], errors='ignore')
        df["18"] = to_numeric(df['18'], errors='ignore')
        df["19"] = to_numeric(df['19'], errors='ignore')
        df["20"] = to_numeric(df['20'], errors='ignore')
        df["21"] = to_numeric(df['21'], errors='ignore')
        df["22"] = to_numeric(df['22'], errors='ignore')
        df["23"] = to_numeric(df['23'], errors='ignore')
        df["24"] = to_numeric(df['24'], errors='ignore')
        df["25"] = to_numeric(df['25'], errors='ignore')
        df["26"] = to_numeric(df['26'], errors='ignore')
        df["27"] = to_numeric(df['27'], errors='ignore')
        df["28"] = to_numeric(df['28'], errors='ignore')
        df["29"] = to_numeric(df['29'], errors='ignore')
        df["30"] = to_numeric(df['30'], errors='ignore')
        df["31"] = to_numeric(df['31'], errors='ignore')
        df["Tổng"] = to_numeric(df['Tổng'], errors='ignore')
        df["Ngày vào"] = to_datetime(df['Ngày vào'], errors='ignore')
        df["Ngày chính thức"] = to_datetime(df['Ngày chính thức'], errors='ignore')
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
                    if cell.column_letter in ['G','H'] and cell.value is not None:
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
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
        df["Tháng"] = to_numeric(df['Tháng'], errors='ignore')
        df["Năm"] = to_numeric(df['Năm'], errors='ignore')
        df["01"] = to_numeric(df['01'], errors='ignore')
        df["02"] = to_numeric(df['02'], errors='ignore')
        df["03"] = to_numeric(df['03'], errors='ignore')
        df["04"] = to_numeric(df['04'], errors='ignore')
        df["05"] = to_numeric(df['05'], errors='ignore')
        df["06"] = to_numeric(df['06'], errors='ignore')
        df["07"] = to_numeric(df['07'], errors='ignore')
        df["08"] = to_numeric(df['08'], errors='ignore')
        df["09"] = to_numeric(df['09'], errors='ignore')
        df["10"] = to_numeric(df['10'], errors='ignore')
        df["11"] = to_numeric(df['11'], errors='ignore')
        df["12"] = to_numeric(df['12'], errors='ignore')
        df["13"] = to_numeric(df['13'], errors='ignore')
        df["14"] = to_numeric(df['14'], errors='ignore')
        df["15"] = to_numeric(df['15'], errors='ignore')
        df["16"] = to_numeric(df['16'], errors='ignore')
        df["17"] = to_numeric(df['17'], errors='ignore')
        df["18"] = to_numeric(df['18'], errors='ignore')
        df["19"] = to_numeric(df['19'], errors='ignore')
        df["20"] = to_numeric(df['20'], errors='ignore')
        df["21"] = to_numeric(df['21'], errors='ignore')
        df["22"] = to_numeric(df['22'], errors='ignore')
        df["23"] = to_numeric(df['23'], errors='ignore')
        df["24"] = to_numeric(df['24'], errors='ignore')
        df["25"] = to_numeric(df['25'], errors='ignore')
        df["26"] = to_numeric(df['26'], errors='ignore')
        df["27"] = to_numeric(df['27'], errors='ignore')
        df["28"] = to_numeric(df['28'], errors='ignore')
        df["29"] = to_numeric(df['29'], errors='ignore')
        df["30"] = to_numeric(df['30'], errors='ignore')
        df["31"] = to_numeric(df['31'], errors='ignore')
        df["Tổng"] = to_numeric(df['Tổng'], errors='ignore')
        df["Ngày vào"] = to_datetime(df['Ngày vào'], errors='ignore')
        df["Ngày chính thức"] = to_datetime(df['Ngày chính thức'], errors='ignore')
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
                    if cell.column_letter in ['G','H'] and cell.value is not None:
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
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
        df["Tháng"] = to_numeric(df['Tháng'], errors='ignore')
        df["Năm"] = to_numeric(df['Năm'], errors='ignore')
        df["01"] = to_numeric(df['01'], errors='ignore')
        df["02"] = to_numeric(df['02'], errors='ignore')
        df["03"] = to_numeric(df['03'], errors='ignore')
        df["04"] = to_numeric(df['04'], errors='ignore')
        df["05"] = to_numeric(df['05'], errors='ignore')
        df["06"] = to_numeric(df['06'], errors='ignore')
        df["07"] = to_numeric(df['07'], errors='ignore')
        df["08"] = to_numeric(df['08'], errors='ignore')
        df["09"] = to_numeric(df['09'], errors='ignore')
        df["10"] = to_numeric(df['10'], errors='ignore')
        df["11"] = to_numeric(df['11'], errors='ignore')
        df["12"] = to_numeric(df['12'], errors='ignore')
        df["13"] = to_numeric(df['13'], errors='ignore')
        df["14"] = to_numeric(df['14'], errors='ignore')
        df["15"] = to_numeric(df['15'], errors='ignore')
        df["16"] = to_numeric(df['16'], errors='ignore')
        df["17"] = to_numeric(df['17'], errors='ignore')
        df["18"] = to_numeric(df['18'], errors='ignore')
        df["19"] = to_numeric(df['19'], errors='ignore')
        df["20"] = to_numeric(df['20'], errors='ignore')
        df["21"] = to_numeric(df['21'], errors='ignore')
        df["22"] = to_numeric(df['22'], errors='ignore')
        df["23"] = to_numeric(df['23'], errors='ignore')
        df["24"] = to_numeric(df['24'], errors='ignore')
        df["25"] = to_numeric(df['25'], errors='ignore')
        df["26"] = to_numeric(df['26'], errors='ignore')
        df["27"] = to_numeric(df['27'], errors='ignore')
        df["28"] = to_numeric(df['28'], errors='ignore')
        df["29"] = to_numeric(df['29'], errors='ignore')
        df["30"] = to_numeric(df['30'], errors='ignore')
        df["31"] = to_numeric(df['31'], errors='ignore')
        df["Tổng"] = to_numeric(df['Tổng'], errors='ignore')
        df["Ngày vào"] = to_datetime(df['Ngày vào'], errors='ignore')
        df["Ngày chính thức"] = to_datetime(df['Ngày chính thức'], errors='ignore')
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
                    if cell.column_letter in ['G','H'] and cell.value is not None:
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
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
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
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
        df["Ngày"] = to_datetime(df['Ngày'], errors='ignore', dayfirst=True)
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
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
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

@app.route("/bangcong_thang_web", methods=["GET","POST"])
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
            "Ngày vào": row[6] if row[6] else "",
            "Ngày chính thức": row[7] if row[7] else "",
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
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
        df["Công thử việc"] = to_numeric(df['Công thử việc'], errors='ignore')
        df["Công chính thức"] = to_numeric(df['Công chính thức'], errors='ignore')
        df["Tăng ca chế độ thử việc"] = to_numeric(df['Tăng ca chế độ thử việc'], errors='ignore')
        df["Tăng ca chế độ chính thức"] = to_numeric(df['Tăng ca chế độ chính thức'], errors='ignore')
        df["Tăng ca ngày thử việc"] = to_numeric(df['Tăng ca ngày thử việc'], errors='ignore')
        df["Tăng ca ngày chính thức"] = to_numeric(df['Tăng ca ngày chính thức'], errors='ignore')
        df["Tăng ca đêm thử việc"] = to_numeric(df['Tăng ca đêm thử việc'], errors='ignore')
        df["Tăng ca đêm chính thức"] = to_numeric(df['Tăng ca đêm chính thức'], errors='ignore')
        df["Tăng ca chủ nhật thử việc"] = to_numeric(df['Tăng ca chủ nhật thử việc'], errors='ignore')
        df["Tăng ca chủ nhật chính thức"] = to_numeric(df['Tăng ca chủ nhật chính thức'], errors='ignore')
        df["Tăng ca ngày lễ thử việc"] = to_numeric(df['Tăng ca ngày lễ thử việc'], errors='ignore')
        df["Tăng ca ngày lễ chính thức"] = to_numeric(df['Tăng ca ngày lễ chính thức'], errors='ignore')
        df["Số lần nghỉ không lương"] = to_numeric(df['Số lần nghỉ không lương'], errors='ignore')
        df["Nghỉ tự do (UA)"] = to_numeric(df['Nghỉ tự do (UA)'], errors='ignore')
        df["Số giờ UP"] = to_numeric(df['Số giờ UP'], errors='ignore')
        df["Nghỉ không lương (UP)"] = to_numeric(df['Nghỉ không lương (UP)'], errors='ignore')
        df["Nghỉ không lương không ảnh hưởng TTNQ (UP01,CL)"] = to_numeric(df['Nghỉ không lương không ảnh hưởng TTNQ (UP01,CL)'], errors='ignore')
        df["Nghỉ phép(AL)"] = to_numeric(df['Nghỉ phép(AL)'], errors='ignore')
        df["Nghỉ phép không ảnh hưởng TTNQ(AL01)"] = to_numeric(df['Nghỉ phép không ảnh hưởng TTNQ(AL01)'], errors='ignore')
        df["Nghỉ hưởng lương thử việc"] = to_numeric(df['Nghỉ hưởng lương thử việc'], errors='ignore')
        df["Nghỉ hưởng lương chính thức"] = to_numeric(df['Nghỉ hưởng lương chính thức'], errors='ignore')
        df["Nghỉ tai nạn lao động(OCL)"] = to_numeric(df['Nghỉ tai nạn lao động(OCL)'], errors='ignore')
        df["Nghỉ ốm, con ốm(SL)"] = to_numeric(df['Nghỉ ốm, con ốm(SL)'], errors='ignore')
        df["Công tác(BL)"] = to_numeric(df['Công tác(BL)'], errors='ignore')
        df["Khám thai(ML03)"] = to_numeric(df['Khám thai(ML03)'], errors='ignore')
        df["Nghỉ vợ sinh(ML02)"] = to_numeric(df['Nghỉ vợ sinh(ML02)'], errors='ignore')
        df["Nghỉ thai sản(LML)"] = to_numeric(df['Nghỉ thai sản(LML)'], errors='ignore')
        df["Nghỉ việc(OSL)"] = to_numeric(df['Nghỉ việc(OSL)'], errors='ignore')
        df["Tổng cộng"] = to_numeric(df['Tổng cộng'], errors='ignore')
        df["Số biên bản kỷ luật"] = to_numeric(df['Số biên bản kỷ luật'], errors='ignore')
        df["Ngày vào"] = to_datetime(df['Ngày vào'], errors='ignore',yearfirst=True)
        df["Ngày chính thức"] = to_datetime(df['Ngày chính thức'], errors='ignore',yearfirst=True)
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
                list_col = ['G','H']
                try:
                    # Apply the date format to column L (assuming 'Ngày thực hiện' is in column 'L')
                    if cell.column_letter in list_col and cell.value is not None:
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
        response.headers['Content-Disposition'] = f'attachment; filename=bangcongthang_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response

@app.route("/bangcongtrangoai_web", methods=["GET","POST"])
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
            "Ngày vào": row[6] if row[6] else "",
            "Ngày chính thức": row[7] if row[7] else "",
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
            "Tăng ca ngày lễ chính thức": row[20]          
        } for row in danhsach] 
        df = DataFrame(data)
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
        df["Công thử việc"] = to_numeric(df['Công thử việc'], errors='ignore')
        df["Công chính thức"] = to_numeric(df['Công chính thức'], errors='ignore')
        df["Tăng ca chế độ thử việc"] = to_numeric(df['Tăng ca chế độ thử việc'], errors='ignore')
        df["Tăng ca chế độ chính thức"] = to_numeric(df['Tăng ca chế độ chính thức'], errors='ignore')
        df["Tăng ca ngày thử việc"] = to_numeric(df['Tăng ca ngày thử việc'], errors='ignore')
        df["Tăng ca ngày chính thức"] = to_numeric(df['Tăng ca ngày chính thức'], errors='ignore')
        df["Tăng ca đêm thử việc"] = to_numeric(df['Tăng ca đêm thử việc'], errors='ignore')
        df["Tăng ca đêm chính thức"] = to_numeric(df['Tăng ca đêm chính thức'], errors='ignore')
        df["Tăng ca chủ nhật thử việc"] = to_numeric(df['Tăng ca chủ nhật thử việc'], errors='ignore')
        df["Tăng ca chủ nhật chính thức"] = to_numeric(df['Tăng ca chủ nhật chính thức'], errors='ignore')
        df["Tăng ca ngày lễ thử việc"] = to_numeric(df['Tăng ca ngày lễ thử việc'], errors='ignore')
        df["Tăng ca ngày lễ chính thức"] = to_numeric(df['Tăng ca ngày lễ chính thức'], errors='ignore')
        df["Ngày vào"] = to_datetime(df['Ngày vào'], errors='ignore')
        df["Ngày chính thức"] = to_datetime(df['Ngày chính thức'], errors='ignore')
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
                list_col = ['G','H']
                try:
                    # Apply the date format to column L (assuming 'Ngày thực hiện' is in column 'L')
                    if cell.column_letter in list_col and cell.value is not None:
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
        response.headers['Content-Disposition'] = f'attachment; filename=bangcongtrangoai_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response
    
@app.route("/gd_pheduyet_tuyendung", methods=["POST"])
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
                    "Ngày đăng ký": datetime.strptime(row[6], "%Y-%m-%d").strftime("%d/%m/%Y") if row[6] else ""      
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
                    "Ngày đăng ký": datetime.strptime(row[6], "%Y-%m-%d").strftime("%d/%m/%Y") if row[6] else "",    
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
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
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
        response.headers['Content-Disposition'] = f'attachment; filename=dangkylamngayle_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response   
            
@app.route("/dangky_chunhat_web", methods=["GET","POST"])
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
            flash(f"Lỗi lấy bảng đăng ký làm ngày leex: ({e})")   
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
                    "Ngày đăng ký": datetime.strptime(row[6], "%Y-%m-%d").strftime("%d/%m/%Y") if row[6] else ""      
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
                    "Ngày đăng ký": datetime.strptime(row[6], "%Y-%m-%d").strftime("%d/%m/%Y") if row[6] else "",    
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
        df["Mã số thẻ"] = to_numeric(df['Mã số thẻ'], errors='ignore')
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
        response.headers['Content-Disposition'] = f'attachment; filename=dangkylamchunhat_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response  

@app.route("/dangky_dilam_ngayle", methods=["POST"])
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
def dangky_dilam_chunhat():
    if request.method == "POST":
        file = request.files.get("file")
        if file:
            try:
                thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
                filepath = os.path.join(FOLDER_NHAP, f"dangki_dilam_chunhat_{thoigian}.xlsx")
                file.save(filepath)
                data = pd.read_excel(filepath ).to_dict(orient="records")
                if "HR" in current_user.phongban:
                    for row in data:
                        id = row["ID"]
                        hrpheduyet = row["HR phê duyệt"] if not pd.isna(row["HR phê duyệt"]) else ""
                        congkhai = row["Công khai"] if not pd.isna(row["Công khai"]) else ""
                        if hr_pheduyet_dilam_chunhat(id,hrpheduyet,congkhai):
                            flash(f"Nhân sự phê duyệt làm Chủ nhật ID {id} thành công !!!")
                        else:
                            flash(f"Nhân sự phê duyệt làm Chủ nhật ID {id} thất bại !!!")       
                else:
                    for row in data:
                        nhamay = current_user.macongty
                        mst = row["Mã số thẻ"]
                        hoten = row["Họ tên"]
                        chuyen = row["Chuyền"]
                        bophan = row["Bộ phận"]
                        vitri = row["Vị trí"]
                        ngay = row["Ngày đăng ký"]
                        if them_dangky_dilam_chunhat(nhamay,mst,hoten,chuyen,bophan,vitri,ngay):
                            flash(f"Thêm làm Chủ nhật thành công !!!")
                        else:
                            flash(f"Thêm làm Chủ nhật  thất bại !!!")       
            except Exception as e:
                print(e)
            finally:         
                return redirect("/dangky_chunhat_web")

@app.route('/download_JD',methods=["POST"])
def download_file():
    try:
        filename = request.form.get("filename")
        print(os.path.exists(filename))
        return send_file(filename, as_attachment=True)
    except Exception as e:
        print(e)
        return redirect("/muc2_2")
 
@app.route('/duyet_hangloat_tangca',methods=["POST"])
def duyet_hangloat_tangca():  
    try:
        chuyen = request.form.getlist("chuyen")
        ngay = request.form.get("ngay") 
        pheduyet = ""  
        danhsach = danhsach_tangca(chuyen,ngay,pheduyet)
        print(len(danhsach))
        for x in danhsach:
            print(x['ID'],hr_pheduyet_tangca(x['ID'],"OK") )   
    except Exception as e:
        flash(f"Lỗi phê duyệt hàng loạt")
    return redirect("/dangki_tangca_web")   

@app.route("/thaydoi_chuyen_lichsu_congviec", methods=["POST"])
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
            flash(f"Sửa Chuyền bắt đầu cho dòng lịch sử công việc số {id} sang {chuyen} thất bại")
        return redirect(f"/muc6_3?mst={mst}&chuyen={chuyen_filter}&bophan={bophan}")
    
@app.route("/thaydoi_bophan_lichsu_congviec", methods=["POST"])
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
            flash(f"Sửa bộ phận bắt đầu cho dòng lịch sử công việc số {id} sang {bophan} thất bại")
        return redirect(f"/muc6_3?mst={mst}&chuyen={chuyen}&bophan={bophan_filter}")
    
@app.route("/thaydoi_chucdanh_lichsu_congviec", methods=["POST"])
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
            flash(f"Sửa chức danh bắt đầu cho dòng lịch sử công việc số {id} sang {chuyen} thất bại")
        return redirect(f"/muc6_3?mst={mst}&chuyen={chuyen}&bophan={bophan}")
    
@app.route("/thaydoi_capbac_lichsu_congviec", methods=["POST"])
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
            flash(f"Sửa cấp bậc bắt đầu cho dòng lịch sử công việc số {id} sang {chuyen} thất bại")
        return redirect(f"/muc6_3?mst={mst}&chuyen={chuyen}&bophan={bophan}")
    
@app.route("/thaydoi_hccategory_lichsu_congviec", methods=["POST"])
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
            flash(f"Sửa HC category bắt đầu cho dòng lịch sử công việc số {id} sang {chuyen} thất bại")
        return redirect(f"/muc6_3?mst={mst}&chuyen={chuyen}&bophan={bophan}")