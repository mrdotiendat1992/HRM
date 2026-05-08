# -*- encoding: utf-8 -*-

from app import *

##################################
#          MAIN ROUTES           #
##################################

# from functools import wraps
from flask import g, flash, request, render_template
from flask_login import current_user
from jinja2 import TemplateNotFound

# ── Helpers ───────────────────────────────────────────────────────────────────

def _sum_don(counts: dict) -> dict:
    """Thêm key 'Tổng' vào dict đếm đơn."""
    counts["Tổng"] = sum(counts.values())
    return counts


def _lay_don_ca_nhan(macongty, masothe) -> dict:
    """
    Gom tất cả thông tin cá nhân vào 1 lần — lý tưởng nhất nên
    gộp thành 1 stored-procedure / query trả về nhiều result-set.
    Hiện tại vẫn gọi hàm cũ nhưng đã tách riêng để dễ tối ưu sau.
    """
    def _nhom(fn_chua, fn_da, fn_duyet, fn_tuchoi):
        c = fn_chua(macongty, masothe)
        d = fn_da(macongty, masothe)
        p = fn_duyet(macongty, masothe)
        r = fn_tuchoi(macongty, masothe)
        return {"Chưa kiểm tra": c, "Đã kiểm tra": d,
                "Đã phê duyệt": p, "Bị từ chối": r, "Tổng": c + d + p + r}

    ddb  = _nhom(lay_soluong_diemdanhbu_chuakiemtra,   lay_soluong_diemdanhbu_dakiemtra,
                 lay_soluong_diemdanhbu_dapheduyet,    lay_soluong_diemdanhbu_bituchoi)
    nphep = _nhom(lay_soluong_xinnghiphep_chuakiemtra,  lay_soluong_xinnghiphep_dakiemtra,
                  lay_soluong_xinnghiphep_dapheduyet,   lay_soluong_xinnghiphep_bituchoi)
    nkl   = _nhom(lay_soluong_xinnghikhongluong_chuakiemtra, lay_soluong_xinnghikhongluong_dakiemtra,
                  lay_soluong_xinnghikhongluong_dapheduyet,  lay_soluong_xinnghikhongluong_bituchoi)
    nkhac = _nhom(lay_soluong_xinnghikhac_chuakiemtra, lay_soluong_xinnghikhac_dakiemtra,
                  lay_soluong_xinnghikhac_dapheduyet,  lay_soluong_xinnghikhac_bituchoi)

    return {
        "Điểm danh bù":       ddb,
        "Xin nghỉ phép":      nphep,
        "Xin nghỉ không lương": nkl,
        "Xin nghỉ khác":      nkhac,
        "Tổng":               ddb["Tổng"] + nphep["Tổng"] + nkl["Tổng"] + nkhac["Tổng"],
        "Lỗi chấm công":      lay_soluong_loichamcong(macongty, masothe),
    }


# Map phanquyen → phòng ban cần truyền (None = toàn công ty)
_TUYEN_DUNG_PHONGBAN = {
    "gd":  None,
    "td":  None,
    "sa":  None,
}

def _lay_tuyen_dung(macongty, phanquyen, phongban) -> dict:
    """
    Trả về dict thông báo tuyển dụng theo quyền.
    gd  → chỉ cần đếm 'chờ phê duyệt'
    tbp / thư ký → theo phòng ban
    td / sa → toàn công ty
    """
    if phanquyen == "gd":
        return {"Tuyển dụng chờ phê duyệt": lay_soluong_yeucautuyendung_chopheduyet(macongty, None)}

    # Xác định scope phòng ban
    if phanquyen in ("td", "sa"):
        pb = None
    elif phanquyen == "tbp" or kiemtra_danhsach_thuki():
        pb = phongban
    else:
        return {}

    return {
        "Tuyển dụng chờ kiểm tra":  lay_soluong_yeucautuyendung_chokiemtra(macongty, pb),
        "Tuyển dụng chờ phê duyệt": lay_soluong_yeucautuyendung_chopheduyet(macongty, pb),
        "Tuyển dụng được duyệt":    lay_soluong_yeucautuyendung_dapheduyet(macongty, pb),
        "Tuyển dụng bị từ chối":    lay_soluong_yeucautuyendung_bituchoi(macongty, pb),
    }


# ── before_request ─────────────────────────────────────────────────────────────

@app.before_request
def run_before_every_request():
    """Kiểm tra đăng nhập, gom thông báo vào g.notice."""
    if not current_user.is_authenticated:
        return

    f12  = trang_thai_function_12()
    mact = current_user.macongty
    mast = current_user.masothe

    notice = {"f12": f12, "db": url_database_pyodbc, "Tổng": 0}

    try:
        # ── Quản lý ──────────────────────────────────────────────────────────
        if la_quanly(mact, mast):
            ql = {
                "Điểm danh bù":       lay_soluong_diemdanhbu_quanly_canduyet(mact, mast),
                "Xin nghỉ phép":      lay_soluong_xinnghiphep_quanly_canduyet(mact, mast),
                "Xin nghỉ không lương": lay_soluong_xinnghikhongluong_quanly_canduyet(mact, mast),
                "Xin nghỉ khác":      lay_soluong_xinnghikhac_quanly_canduyet(mact, mast),
            }
            ql["Số thông báo"] = sum(ql.values())
            notice["Quản lý"]  = ql
            notice["Tổng"]    += ql["Số thông báo"]
        else:
            notice["Quản lý"] = {}

        # ── Thư ký ───────────────────────────────────────────────────────────
        if la_thuky(mact, mast):
            chuyen = lay_danhsach_chuyen_thuky_quanly(mact, mast)
            tk = {
                "Danh sách lỗi thẻ":    lay_soluong_loithe_thuky_canxuly(mact, mast),
                "Điểm danh bù":         lay_soluong_diemdanhbu_thuky_cankiemtra(mact, mast),
                "Xin nghỉ phép":        lay_soluong_xinnghiphep_thuky_cankiemtra(mact, mast),
                "Xin nghỉ không lương": lay_soluong_xinnghikhongluong_thuky_cankiemtra(mact, mast),
                "Xin nghỉ khác":        lay_soluong_xinnghikhac_thuky_cankiemtra(mact, mast),
                "Line":                 chuyen[0] if len(chuyen) == 1 else "",
            }
            tk["Số thông báo"] = (tk["Danh sách lỗi thẻ"] + tk["Điểm danh bù"]
                                  + tk["Xin nghỉ phép"] + tk["Xin nghỉ không lương"])
            notice["Thư ký"]   = tk
            notice["Tổng"]    += tk["Số thông báo"] + tk["Xin nghỉ khác"]
        else:
            notice["Thư ký"] = {}

        # ── Cá nhân ──────────────────────────────────────────────────────────
        notice["personal"] = _lay_don_ca_nhan(mact, mast)

        # ── Tuyển dụng ───────────────────────────────────────────────────────
        td = _lay_tuyen_dung(mact, current_user.phanquyen, current_user.phongban)
        for k, v in td.items():
            if v > 0:
                notice[k]       = v
                notice["Tổng"] += v
            else:
                notice.setdefault(k, 0)

    except Exception as e:
        flash(f"Lỗi cập nhật thông tin chuông: {e}")
        notice = {"f12": f12, "db": url_database_pyodbc}

    g.notice = notice


@app.context_processor
def inject_notice():
    return dict(notice=getattr(g, "notice", {}),
                personal=getattr(g, "personal", {}))

def _is_mobile() -> bool:
    """Phát hiện truy cập từ thiết bị di động dựa vào User-Agent (đơn giản)."""
    ua = (request.user_agent.string or "").lower()
    return any(x in ua for x in ("iphone", "android", "ipad"))


def _render_with_mobile_fallback(default_template: str, **context):
    """Thử render template mobile/..., nếu không có thì dùng template mặc định.

    Ví dụ: default_template="home.html" → ưu tiên "mobile/home.html".
    """
    if _is_mobile():
        mobile_name = f"mobile/{default_template}"
        try:
            return render_template(mobile_name, **context)
        except TemplateNotFound:
            pass
    return render_template(default_template, **context)


@app.route('/unauthorized')
def unauthorized():
    return render_template_string("<h1>Bạn không thể vào mục này, vui lòng chọn mục khác!!!</h1><h3>Ấn vào <a href='/'>đây</a> để quay lại trang chủ</h3>")

@app.errorhandler(404)
def page_not_found(e):
    return render_template('blank.html'), 404

@login_manager.unauthorized_handler
def unauthorized():
    return redirect(url_for('login'))

@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("home"))

    if request.method == "POST":
        macongty = request.form.get("macongty", "").strip()
        masothe  = request.form.get("masothe",  "").strip()
        matkhau  = request.form.get("matkhau",  "")

        user = Nhanvien.query.filter_by(masothe=masothe, macongty=macongty).first()

        if user and user.matkhau == matkhau and login_user(user):
            app.logger.info(f"[LOGIN OK] {masothe}@{macongty}")
            next_page = request.form.get("next") or ""
            return redirect(next_page if next_page and next_page != "/login"
                            else url_for("home"))

        app.logger.warning(f"[LOGIN FAIL] {masothe}@{macongty}")
        flash("Sai thông tin đăng nhập.", "danger")
        return redirect(url_for("login"))

    return _render_with_mobile_fallback("login.html")

@app.route("/logout", methods=["POST"])
def logout():
    try:
        app.logger.info(f"Nguoi dung {current_user.masothe} o {current_user.macongty} vua  dang xuat !!!")
        logout_user()
    except Exception as e:
        app.logger.error(f'Không thế đăng xuất {e} !!!')
    return redirect("/")

@app.route("/doimatkhau", methods=['POST'])
def doimatkhau():
    macongty = request.form.get("macongty")
    masothe = request.form.get("masothe_doi")
    matkhaumoi = request.form.get("matkhaumoi")
    try:
        if doimatkhautaikhoan(macongty,masothe,matkhaumoi):
            flash("Đổi mật khẩu thành công")
    except Exception as e:
        app.logger.error(f"{masothe} o {macongty} doi mat khau thanh {matkhaumoi} thanh cong !!!")
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
        hccategory = request.args.get("HC Category")
        ghichu = request.args.get("Ghi chú")
        chuyen = request.args.get("Chuyền")
        users = laydanhsachuser(mst, hoten, sdt, cccd, gioitinh, vaotungay, vaodenngay, nghitungay, nghidenngay, phongban, trangthai, hccategory, chucvu, ghichu, chuyen)   
        count = len(users)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(users)
        start = (page - 1) * per_page
        end = start + per_page
        paginated_users = users[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        songuoi_danglamviec = lay_soluong_danglamviec()
        songuoi_dangnghithaisan = lay_soluong_dangnghithaisan()
        flash(f"Xin chào {current_user.hoten} !!!")
        return _render_with_mobile_fallback(
            "home.html",
            users=paginated_users,
            page="Trang chủ",
            pagination=pagination,
            count=count,
            songuoi_danglamviec=songuoi_danglamviec,
            songuoi_dangnghithaisan=songuoi_dangnghithaisan,
        )
    else:
        try:
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
            
            # Chuyển thông tin ngày về định dạng YYYY-MM-DD
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

            df["Ngày sinh"] = to_datetime(df['Ngày sinh'],errors='coerce')
            df["Ngày cấp CCCD"] = to_datetime(df['Ngày cấp CCCD'],errors='coerce')
            df["Ngày ký HĐ"] = to_datetime(df['Ngày ký HĐ'],errors='coerce')
            df["Ngày vào"] = to_datetime(df['Ngày vào'],errors='coerce')
            df["Ngày nghỉ"] = to_datetime(df['Ngày nghỉ'],errors='coerce')
            df["Ngày hết hạn"] = to_datetime(df['Ngày hết hạn'],errors='coerce')
            df["Ngày vào nối thâm niên"] = to_datetime(df['Ngày vào nối thâm niên'],errors='coerce')
            df["Ngày sinh con 1"] = to_datetime(df['Ngày sinh con 1'],errors='coerce')
            df["Ngày sinh con 2"] = to_datetime(df['Ngày sinh con 2'],errors='coerce')
            df["Ngày sinh con 3"] = to_datetime(df['Ngày sinh con 3'],errors='coerce')
            df["Ngày sinh con 4"] = to_datetime(df['Ngày sinh con 4'],errors='coerce')
            df["Ngày sinh con 5"] = to_datetime(df['Ngày sinh con 5'],errors='coerce')
            df["Ngày kí HĐ Thử việc"] = to_datetime(df['Ngày kí HĐ Thử việc'],errors='coerce')
            df["Ngày hết hạn HĐ Thử việc"] = to_datetime(df['Ngày hết hạn HĐ Thử việc'],errors='coerce')
            df["Ngày kí HĐ xác định thời hạn lần 1"] = to_datetime(df['Ngày kí HĐ xác định thời hạn lần 1'],errors='coerce')
            df["Ngày hết hạn HĐ xác định thời hạn lần 1"] = to_datetime(df['Ngày hết hạn HĐ xác định thời hạn lần 1'],errors='coerce')
            df["Ngày kí HĐ không thời hạn"] = to_datetime(df['Ngày kí HĐ không thời hạn'],errors='coerce')
            
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
        except Exception as e:
            flash(f"Lỗi kết xuất danh sách nhân viên ({e})")
            app.logger.error(f"Lỗi kết xuất danh sách nhân viên ({e})")
            return redirect(url_for("home"))

@app.route("/dashboard", methods=["GET"])
def dashboard():
    data = get_dashboard_data()
    return render_template("dashboard.html", page="Dashboard", data=data)

@app.route("/muc2_1", methods=["GET", "POST"])
@login_required
@roles_required('hr', 'tnc', 'sa', 'gd', 'td', 'tbp')
def danhsachdangkytuyendung():

    if request.method == "GET":
        try:
            hoten   = request.args.get("hoten")
            vitri   = request.args.get("vitri")
            sdt     = request.args.get("sdt")
            cccd    = request.args.get("cccd")
            ngaygui = request.args.get("ngaygui")

            rows  = laydanhsachdangkytuyendung(sdt, cccd, ngaygui, hoten, vitri)
            count = len(rows)

            current_page = request.args.get(get_page_parameter(), type=int, default=1)
            per_page     = 10
            start        = (current_page - 1) * per_page
            pagination   = Pagination(page=current_page, per_page=per_page,
                                      total=count, css_framework='bootstrap4')

            return render_template("2_1.html",
                                   page="2.1 Danh sách ứng viên",
                                   danhsach=rows[start: start + per_page],
                                   pagination=pagination,
                                   count=count)
        except Exception as e:
            flash(f"Lỗi lấy danh sách ứng viên: {e}")
            app.logger.error(f"muc2_1 GET: {e}")
            return redirect(url_for("home"))

    # POST
    try:
        ketqua = capnhatthongtinungvien(
            form_data = request.form,
            id        = request.form.get("id", "").strip(),
            macongty  = current_user.macongty,
        )
        if ketqua["ketqua"]:
            flash("Cập nhật thông tin ứng viên thành công !!!")
        else:
            flash(f"Cập nhật thất bại — {ketqua.get('lido')}")
            app.logger.error(f"muc2_1 POST: {ketqua.get('lido')}")
    except Exception as e:
        flash(f"Lỗi: {e}")
        app.logger.error(f"muc2_1 POST exception: {e}")

    return redirect("/muc2_1")

@app.route("/muc2_2", methods=["GET","POST"])
@login_required
def dangkytuyendung():
    if request.method == "GET":
        try:
            lathuki = kiemtra_danhsach_thuki()
            if (current_user.phanquyen not in ['tbp','gd','sa','td']) and not lathuki:
                return redirect("/unauthorized")
            phongban = request.args.get("phongban")
            danhsach = laydanhsachyeucautuyendung(phongban)
            danhsach_vitri_cacongty = lay_danhsach_vitri_theo_hcname(current_user.macongty)
            # flash(danhsach_vitri_cacongty)
            return render_template("2_2.html", 
                                page= "2.2 Yêu cầu tuyển dụng",
                                danhsach = danhsach,
                                lathuki = lathuki,
                                danhsach_vitri_cacongty=danhsach_vitri_cacongty
                                )
        except Exception as e:
            flash(f"Lỗi lấy danh sách yêu cầu tuyển dụng ({e})")
            app.logger.error(f"Lỗi lấy danh sách yêu cầu tuyển dụng ({e})")
            return redirect(url_for("home"))
        
    elif request.method == "POST":
        try:
            bophan = current_user.phongban
            vitri = request.form.get("vitri")
            if "công nhân" in vitri.lower():
                kieulaodong = "Công nhân"
            else:
                kieulaodong = "Nhân viên"
            vitrien = request.form.get("vitrien")
            capbac = request.form.get("capbac")
            soluong = request.form.get("soluong")
            mota = os.path.join(FOLDER_JD, f"{vitrien}.pdf")
            thoigiandukien = request.form.get("thoigiandukien")
            phanloai = request.form.get("phanloai")
            budget = request.form.get("trong_budget")
            trongbudget = "Trong" if budget else"Ngoài"
            if themyeucautuyendungmoi(bophan,vitri,soluong,mota,thoigiandukien,phanloai,capbac,kieulaodong,trongbudget):
                flash("Thêm yêu cầu tuyển dụng mới thành công !!!")
                flash(them_thongbao_co_yeucautuyendung(vitri,soluong,trongbudget))
            else:
                flash("Thêm yêu cầu tuyển dụng mới thất bại !!!")
        except Exception as e:
            flash(f"Thêm yêu cầu tuyển dụng mới thất bại ({e})!!!")
        return redirect("muc2_2")

@app.route("/muc2_2_1", methods=["GET","POST"])
@login_required
@roles_required('tbp','gd','sa','td')
def tuyendungchitiet():
    if request.method == "GET":
        try:
            id_yeucautuyendung = request.args.get("id")
            thongtin_tuyendung = lay_thongtin_yeucautuyendung(id_yeucautuyendung)
            vitri_tuyendung = thongtin_tuyendung[0]
            phongban = thongtin_tuyendung[1]
            danhsach = lay_danhsach_ungvien(id_yeucautuyendung)
            danhsach_ungvien_tiemnang = lay_danhsach_ungvien_tiemnang(vitri_tuyendung)
            danhsach_ungvien_2_1 = lay_danhsach_ungvien_2_1()
            so_ungvien_tong = len(danhsach)
            so_ungvien_chophongvan = 0
            so_ungvien_dangphongvan = 0
            so_ungvien_quaphongvan = 0
            so_ungvien_danhanviec = 0
            so_ungvien_khongnhanviec = 0
            for ungvien in danhsach:
                if ungvien[16] == "Chưa phỏng vấn":
                    so_ungvien_chophongvan += 1
                elif ungvien[16] == "Đang phỏng vấn":
                    so_ungvien_dangphongvan += 1
                elif ungvien[16] == "Qua phỏng vấn":
                    so_ungvien_quaphongvan += 1
                elif ungvien[16] == "Đã nhận việc":
                    so_ungvien_danhanviec += 1
                elif ungvien[16] == "Không nhận việc":
                    so_ungvien_khongnhanviec += 1
            phongban = lay_phongban_theo_idyctd(id_yeucautuyendung)
            return render_template("2_2_1.html", 
                                page="2.2.1 Danh sách ứng viên tuyển dụng",
                                vitri_tuyendung=vitri_tuyendung,
                                danhsach=danhsach,
                                phongban=phongban,
                                so_ungvien_tong=so_ungvien_tong,
                                so_ungvien_chophongvan=so_ungvien_chophongvan,
                                so_ungvien_dangphongvan=so_ungvien_dangphongvan,
                                so_ungvien_quaphongvan=so_ungvien_quaphongvan,
                                so_ungvien_danhanviec=so_ungvien_danhanviec,
                                so_ungvien_khongnhanviec=so_ungvien_khongnhanviec,
                                danhsach_ungvien_tiemnang=danhsach_ungvien_tiemnang,
                                danhsach_congnhan_ungtuyen=danhsach_ungvien_2_1
                                ) 
        except Exception as e:
            flash(f"Lỗi lấy danh sách ứng viên ({e})")
            app.logger.error(f"Lỗi lấy danh sách ứng viên ({e})")
            return redirect(url_for("home"))
    else:
        try:
            id_yeucautuyendung = request.form.get("id")
            phongban = request.form.get("phongban")
            hoten = request.form.get("hoten")
            gioitinh = request.form.get("gioitinh")
            tuoi = request.form.get("tuoi")
            namkinhnghiem = request.form.get("namkinhnghiem")
            linkcv = request.files.get("linkcv")
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            save_path = os.path.join(FOLDER_CV,f"cv_{timestamp}.pdf")
            linkcv.save(save_path)
            kenhtuyendung = request.form.get("kenhtuyendung")
            if them_ungvientuyendung(id_yeucautuyendung,phongban,hoten,gioitinh,tuoi,namkinhnghiem,save_path,kenhtuyendung):
                flash("Thêm ứng viên thành công")
            return redirect(f"muc2_2_1?id={id_yeucautuyendung}")
        except Exception as e:
            flash(f"Lỗi thêm ứng viên ({e})")
            app.logger.error(f"Lỗi thêm ứng viên ({e})")
            return redirect(url_for("home"))
        
@app.route("/muc3_1", methods=["GET", "POST"])
@login_required
@roles_required('hr', 'sa', 'gd')
def nhapthongtinlaodongmoi():

    if request.method == "GET":
        try:
            masothe  = checkformatmst(int(laymasothemoi()) + 1)
            cacvitri = laycacvitri()
            cacto    = laycacto()
            cacca    = laycacca()
            return render_template(
                "3_1.html",
                page       = "3.1 Nhập thông tin lao động mới",
                qrcccd     = request.args.get("scan-qrcode"),
                masothe    = masothe,
                ngaybatdau = datetime.now(),
                cacvitri   = cacvitri,
                cacto      = cacto,
                cacca      = cacca,
                macongty   = current_user.macongty,
            )
        except Exception as e:
            flash(f"Lỗi lấy thông tin lao động mới: {e}")
            app.logger.error(f"muc3_1 GET error: {e}")
            return redirect(url_for("home"))

    # ── POST ──────────────────────────────────────────────────────────────────
    def _s(key):
        v = request.form.get(key, "").strip()
        return v if v else None

    def _sql_nstr(key):
        v = _s(key)
        return f"N'{v}'" if v else "NULL"

    def _sql_str(key):
        v = _s(key)
        return f"'{v}'" if v else "NULL"

    def _sql_date(key):
        v = _s(key)
        if not v:
            return "NULL"
        try:
            datetime.strptime(v, "%Y-%m-%d")
            return f"'{v}'"
        except ValueError:
            app.logger.warning(f"_sql_date: key='{key}' value='{v}' không hợp lệ → NULL")
            return "NULL"

    try:
        # ── Ảnh ──────────────────────────────────────────────────────────────
        anh  = "NULL"
        file = request.files.get("anh")
        if file and file.filename:
            file_path = os.path.join(FOLDER_AVATAR, _s("masothe") + ".jpg")
            file.save(file_path)
            anh = f"'{file_path}'"

        # ── Định danh ────────────────────────────────────────────────────────
        masothe     = f"'{_s('masothe')}'"
        thechamcong = str(int(_s('masothe')))  # int, không quotes

        # ── Cá nhân ──────────────────────────────────────────────────────────
        hoten        = _sql_nstr("hoten")
        ngaysinh     = _sql_date("ngaysinh")
        gioitinh     = _sql_nstr("gioitinh")
        cmt          = _sql_str("cmt")
        cccd         = _sql_str("cccd")
        ngaycapcccd  = _sql_date("ngaycap")
        noicapcccd   = _sql_nstr("noicap")
        thuongtru    = _sql_nstr("thuongtru")
        noisinh      = _sql_nstr("noisinh")
        tamtru       = _sql_nstr("tamtru")
        quoctich     = _sql_nstr("quoctich")
        dantoc       = _sql_nstr("dantoc")
        tongiao      = _sql_nstr("tongiao")
        hocvan       = _sql_nstr("hocvan")
        thonxom      = _sql_nstr("thonxom")
        phuongxa     = _sql_nstr("phuongxa")
        quanhuyen    = _sql_nstr("quanhuyen")
        tinhthanhpho = _sql_nstr("tinhthanhpho")
        nguoithan    = _sql_nstr("nguoithan")
        sdtnguoithan = _sql_nstr("sdtnguoithan")

        # ── Tài chính / liên hệ ──────────────────────────────────────────────
        nganhang   = _sql_nstr("nganhang")
        sotaikhoan = _sql_str("sotaikhoan")
        dienthoai  = _sql_str("dienthoai")
        sobhxh     = _sql_str("sobhxh")
        masothue   = _sql_str("masothue")

        # ── Con nhỏ ──────────────────────────────────────────────────────────
        connho      = _sql_nstr("connho")
        tencon      = [_sql_nstr(f"tenconnho{i}")  for i in range(1, 6)]
        ngaysinhcon = [_sql_date(f"ngaysinhcon{i}") for i in range(1, 6)]

        # ── Vị trí ───────────────────────────────────────────────────────────
        jobdetailvn             = _sql_nstr("vitri")
        line                    = _sql_str("line")
        factory                 = f"'{current_user.macongty}'"
        hccategory              = _sql_nstr("hccategory")
        gradecode               = _sql_nstr("gradecode")
        department              = _sql_nstr("phongban")
        chucvu                  = _sql_nstr("chucvu")
        employeetype            = _sql_nstr("loailaodong")
        sectioncode             = _sql_nstr("mabophan")
        sectiondescription      = _sql_nstr("bophan")
        jobdetailen             = _sql_nstr("vitrien")
        positioncode            = _sql_nstr("mavitri")
        positioncodedescription = _sql_nstr("tenvitri")

        # ── COST_ID ──────────────────────────────────────────────────────────
        cost_id = _sql_str("ntid")

        # ── Cố định ──────────────────────────────────────────────────────────
        luongcoban  = "NULL"
        tongphucap  = "NULL"
        kieuhopdong = "NULL"
        diachimoi   = "NULL"
        nd          = "NULL"  # null date

        # ── INSERT VALUES (75 cột đúng thứ tự) ───────────────────────────────
        nhanvienmoi = (
            # 1-2: Định danh
            f"({masothe},{thechamcong},"
            # 3-6: Cá nhân cơ bản
            f"{hoten},{dienthoai},{ngaysinh},{gioitinh},"
            # 7-11: CCCD + địa chỉ thường trú
            f"{cccd},{ngaycapcccd},{noicapcccd},{cmt},{thuongtru},"
            # 12-15: Địa chỉ chi tiết
            f"{thonxom},{phuongxa},{quanhuyen},{tinhthanhpho},"
            # 16-21: Cá nhân khác
            f"{dantoc},{quoctich},{tongiao},{hocvan},{noisinh},{tamtru},"
            # 22-25: Tài chính
            f"{sobhxh},{masothue},{nganhang},{sotaikhoan},"
            # 26: Con nhỏ
            f"{connho},"
            # 27-36: Con nhỏ 1-5
            f"{tencon[0]},{ngaysinhcon[0]},"
            f"{tencon[1]},{ngaysinhcon[1]},"
            f"{tencon[2]},{ngaysinhcon[2]},"
            f"{tencon[3]},{ngaysinhcon[3]},"
            f"{tencon[4]},{ngaysinhcon[4]},"
            # 37-39: Ảnh, người thân
            f"{anh},{nguoithan},{sdtnguoithan},"
            # 40-42: Hợp đồng
            f"{kieuhopdong},GETDATE(),{nd},"
            # 43-55: Vị trí
            f"{jobdetailvn},{hccategory},{gradecode},{factory},"
            f"{department},{chucvu},{sectioncode},{sectiondescription},"
            f"{line},{employeetype},{jobdetailen},"
            f"{positioncode},{positioncodedescription},"
            # 56-58: Lương (Luong_co_ban, Phu_cap, Tong_phu_cap)
            f"{luongcoban},{nd},{tongphucap},"
            # 59-63: Ngày tháng hành chính
            # Ngay_vao, Ngay_nghi, Trang_thai_lam_viec,
            # Ngay_vao_noi_tham_nien, Mat_khau
            f"GETDATE(),NULL,N'Đang làm việc',GETDATE(),'1',"
            # 64-65: HDTV
            f"{nd},{nd},"
            # 66-67: HDXDTH Lần 1
            f"{nd},{nd},"
            # 68-69: HDXDTH Lần 2
            f"{nd},{nd},"
            # 70: HDKXDTH
            f"{nd},"
            # 71-74: Truong_BP, Ghi_chu, Time_Stamp, Dia_chi_moi
            f"'N','',GETDATE(),{diachimoi},"
            # 75: COST_ID
            f"{cost_id})"
        )

        # app.logger.debug(f"muc3_1 INSERT values: {nhanvienmoi}")

        ketqua = themnhanvienmoi(nhanvienmoi)

        if ketqua["ketqua"]:
            flash("Thêm lao động mới thành công !!!")
            ca = laycatheochuyen(request.form.get("line"))
            thangdangkycalamviec(
                request.form.get("masothe"), ca, ca,
                datetime.now().date().strftime("%Y-%m-%d"),
                datetime(2054, 12, 31).strftime("%Y-%m-%d"),
            )
            themtaikhoanmoi(
                int(request.form.get("masothe")),
                request.form.get("hoten"),
                request.form.get("phongban"),
                request.form.get("gradecode"),
            )
        else:
            flash(f"Thêm lao động mới thất bại: {ketqua['lido']}")
            app.logger.error(f"muc3_1 INSERT failed: {ketqua['lido']}")

    except Exception as e:
        flash(f"Thêm lao động mới thất bại: {e}")
        app.logger.error(f"muc3_1 POST error: {e}")

    finally:
        return redirect("/muc3_1")
        
@app.route("/muc3_2", methods=["GET", "POST"])
@login_required
@roles_required('hr', 'sa', 'gd')
def thaydoithongtinlaodong():

    if request.method == "GET":
        return render_template("3_2.html", page="3.2 Thay đổi thông tin người lao động")

    # ── POST ──────────────────────────────────────────────────────────────────
    try:
        mst = request.form.get("mst", "").strip()

        # ── Ảnh ──
        anh = None
        file = request.files.get("anh")
        if file and file.filename:
            file_path = os.path.join(FOLDER_AVATAR, mst + ".jpg")
            if os.path.exists(file_path):
                os.remove(file_path)
            file.save(file_path)
            anh = file_path

        def _v(key):
            """Trả về giá trị string, None nếu rỗng."""
            v = request.form.get(key, "").strip()
            return v if v else None

        def _num(key):
            v = _v(key)
            return v.replace(",", "") if v else None

        # ── Map: form_key → (col_name, is_nvarchar) ──
        FIELD_MAP = [
            # Cá nhân
            ("cccd",                    "CCCD",                       False),
            ("ngaycapcccd",             "Ngay_cap",                   False),
            ("noicapcccd",              "Noi_cap",                    True),
            ("hoten",                   "Ho_ten",                     True),
            ("ngaysinh",                "Ngay_sinh",                  False),
            ("gioitinh",                "Gioi_tinh",                  True),
            ("cmt",                     "CMT",                        True),
            ("quoctich",                "Quoc_tich",                  True),
            ("dienthoai",               "Sdt",                        False),
            ("thonxom",                 "Thon_xom",                   True),
            ("phuongxa",                "Phuong_xa",                  True),
            ("quanhuyen",               "Quan_huyen",                 True),
            ("tinhthanhpho",            "Tinh_TP",                    True),
            ("thuongtru",               "Dia_chi_thuong_tru",         True),
            ("tamtru",                  "Dia_chi_tam_tru",            True),
            ("noisinh",                 "Noi_sinh",                   True),
            ("dantoc",                  "Dan_toc",                    True),
            ("tongiao",                 "Ton_giao",                   True),
            ("hocvan",                  "Trinh_do",                   True),
            ("masothue",                "Ma_so_thue",                 False),
            ("nganhang",                "Ngan_hang",                  True),
            ("sotaikhoan",              "So_tai_khoan",               False),
            ("diachimoi",               "Dia_chi_moi",                True),
            ("connho",                  "Con_nho",                    True),
            # Con nhỏ
            ("tenconnho1",              "Ten_con_nho_1",              True),
            ("tenconnho2",              "Ten_con_nho_2",              True),
            ("tenconnho3",              "Ten_con_nho_3",              True),
            ("tenconnho4",              "Ten_con_nho_4",              True),
            ("tenconnho5",              "Ten_con_nho_5",              True),
            ("ngaysinhcon1",            "Ngay_sinh_con_nho_1",        False),
            ("ngaysinhcon2",            "Ngay_sinh_con_nho_2",        False),
            ("ngaysinhcon3",            "Ngay_sinh_con_nho_3",        False),
            ("ngaysinhcon4",            "Ngay_sinh_con_nho_4",        False),
            ("ngaysinhcon5",            "Ngay_sinh_con_nho_5",        False),
            ("nguoithan",               "Nguoi_than",                 True),
            ("sdtnguoithan",            "Sdt_Nguoithan",              False),
            # Vị trí
            ("jobtitlevn",              "Job_title_VN",               True),
            ("jobtitleen",              "Job_title_EN",               False),
            ("positioncode",            "Position_code",              False),
            ("positioncodedescription", "Position_code_description",  False),
            ("chucvu",                  "Chuc_vu",                    True),
            ("line",                    "Line",                       False),
            ("department",              "Department",                 False),
            ("sectioncode",             "Section_code",               False),
            ("sectiondescription",      "Section_description",        False),
            ("hccategory",              "Headcount_category",         False),
            ("employeetype",            "Emp_type",                   False),
            ("gradecode",               "Grade_code",                 False),
            ("factory",                 "Factory",                    False),
            # Hợp đồng
            ("kieuhopdong",             "Loai_hop_dong",              True),
            ("ngaybatdau",              "Ngay_ky_HD",                 False),
            ("ngayketthuc",             "Ngay_het_han_HD",            False),
            ("phucap",                  "Phu_cap",                    True),
            # Ngày tháng / trạng thái
            ("trangthai",               "Trang_thai_lam_viec",        True),
            ("ngayvao",                 "Ngay_vao",                   False),
            ("ngaynghi",                "Ngay_nghi",                  False),
            ("ngaykyhdtv",              "Ngay_ky_HDTV",               False),
            ("ngayhethanhdtv",          "Ngay_het_han_HDTV",          False),
            ("COST_ID",                 "COST_ID",                    False),
        ]

        # ── Xây dựng SET clause ──
        set_parts = []

        # Các trường thông thường từ FIELD_MAP
        for form_key, col, is_nv in FIELD_MAP:
            val = _v(form_key)
            if val:
                prefix = "N" if is_nv else ""
                set_parts.append(f"{col} = {prefix}'{val}'")
            else:
                set_parts.append(f"{col} = NULL")

        # The_cham_cong: ép kiểu int
        try:
            tcc = int(mst) if mst else None
        except (ValueError, TypeError):
            tcc = None
        set_parts.append(f"The_cham_cong = {tcc}" if tcc is not None else "The_cham_cong = NULL")

        # Các cột số (cần strip dấu phẩy)
        for form_key, col in [("mucluong", "Luong_co_ban"), ("tongphucap", "Tong_phu_cap")]:
            val = _num(form_key)
            set_parts.append(f"{col} = '{val}'" if val else f"{col} = NULL")

        # Ảnh
        set_parts.append(f"Anh_chan_dung = '{anh}'" if anh else "Anh_chan_dung = NULL")

        query = (
            "UPDATE Danh_sach_CBCNV SET "
            + ", ".join(set_parts)
            + f" WHERE MST = '{mst}' AND Factory = '{current_user.macongty}'"
        )

        conn = pyodbc.connect(url_database_pyodbc)
        cursor = conn.cursor()
        cursor.execute(query)
        conn.commit()
        conn.close()
        flash("Cập nhật thông tin người lao động thành công !!!")

    except Exception as e:
        flash(f"Cập nhật thông tin người lao động thất bại: {e}")
        app.logger.error(f"muc3_2 POST error: {e}")

    return redirect("/muc3_2")
    
@app.route("/muc3_3", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def quanlyhopdong():
    try:
        if request.method == "GET":
            mst = request.args.get("mst")
            if not mst:
                mst = current_user.masothe
            danhsach = laydanhsach_hopdong_theomst(mst)
            return _render_with_mobile_fallback(
                "3_3.html",
                page="3.3 Quản lý hợp đồng lao động",
                danhsach=danhsach,
            )
        elif request.method == "POST":
            nhamay = current_user.macongty
            mst = request.form.get("form_manhanvien")
            hoten = request.form.get("form_hovaten")
            gioitinh = request.form.get("form_gioitinh")
            ngaysinh =  request.form.get("form_ngaysinh")
            thuongtru = request.form.get("form_thuongtru")
            tamtru = request.form.get("form_tamtru")
            cccd = request.form.get("form_cccd")
            noicapcccd = request.form.get("form_noicap_cccd")
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

            if themhopdongmoi(nhamay,mst,hoten,gioitinh,ngaysinh,thuongtru,tamtru,cccd,noicapcccd,ngaycapcccd,capbac,loaihopdong,chucdanh,phongban,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc):
                flash("Thêm hợp đồng thành công !!!")
                # capnhatthongtinhopdong(nhamay,mst,loaihopdong,chucdanh,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc,vitrien,employeetype,positioncode,postitioncodedescription,hccategory,sectioncode,sectiondescription)
            else:
                flash("Thêm hợp đồng thất bại")
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
        flash("Tải file thành công !!!")
        return send_file(os.path.join(FOLDER_XUAT, f"saphethan_{thoigian}.xlsx"), as_attachment=True)

@app.route("/muc3_5", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def danhsachsapnghihuu():
    if request.method == "GET":
        danhsach = laydanhsachsapnghihuu()
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("/3_5.html",
                               danhsach=paginated_rows,
                               pagination=pagination,
                               total = total)

    elif request.method == "POST":
        try:
            # tải danh sách sắp nghỉ hưu xuống excel
            # MST, Ho_ten, Chuc_danh, Gioi_tinh, Chuyen, Bo_phan, Ngay_sinh, Ngay_nghi_huu, So_thang_con_lai
            danhsach = laydanhsachsapnghihuu()
            df = pd.DataFrame(danhsach)
            df.columns = ["MST", "Ho_ten", "Chuc_danh", "Gioi_tinh", "Chuyen", "Bo_phan", "Ngay_sinh", "Ngay_nghi_huu", "So_thang_con_lai"]
            thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
            df.to_excel(os.path.join(FOLDER_XUAT, f"sapnghihuu_{thoigian}.xlsx"), index=False)
            flash("Tải file thành công !!!")
            return send_file(os.path.join(FOLDER_XUAT, f"sapnghihuu_{thoigian}.xlsx"), as_attachment=True)
        except Exception as e:
            flash(str(e))
            return redirect("/muc3_5")

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
                        flash("Upload new KPI failed: Cannot insert data !!!")
                        return redirect("/muc5_1_1")
                guimailthongbaodaguikpi(current_user.macongty,current_user.masothe,current_user.hoten)
                flash("Upload new KPI successfully !!!")
            else:
                flash("Upload new KPI failed: Cannot found data !!!")
        except Exception as e:
            flash(f"Upload new KPI failed {e} !!!")
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
            
            khongdoica = request.form.get("khongdoica") 
            
            if loaidieuchuyen == "Chuyển vị trí":
                try:
                    ketqua = dieuchuyennhansu(mst,
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
                                    )
                    if ketqua["ketqua"]:
                        flash("Điều chuyển thành công !!!")
                    else:
                        flash(f"Điều chuyển thất bại, lí do: {ketqua['lido']}, query: {ketqua['query']} !!!")
                except Exception as e:
                    flash(f"Điều chuyển thất bại, lí do: {e}")
                return redirect(f"/muc6_1")
                
            elif loaidieuchuyen == "Nghỉ việc":
                try:
                    ketqua = dichuyennghiviec(mst,
                        vitricu,
                        chuyencu,
                        gradecodecu,
                        hccategorycu,
                        ngaydieuchuyen,
                        ghichu)
                    if ketqua["ketqua"]:
                        flash("Điều chuyển thành công !!!")
                    else:
                        flash(f"Điều chuyển thất bại, lí do: {ketqua['lido']}, query: {ketqua['query']} !!!")
                except Exception as e:
                    flash(f"Điều chuyển thất bại, lí do: {e}")
                return redirect(f"/muc6_1")
            elif loaidieuchuyen=="Nghỉ thai sản":
                try:
                    ketqua = dichuyennghi(mst,
                                vitricu,
                                chuyencu,
                                gradecodecu,
                                hccategorycu,
                                ngaydieuchuyen,
                                'Nghỉ thai sản'
                                )
                    if ketqua["ketqua"]:
                        flash("Điều chuyển thành công !!!")
                    else:
                        flash(f"Điều chuyển thất bại, lí do: {ketqua['lido']}, query: {ketqua['query']} !!!")
                except Exception as e:
                    flash(f"Điều chuyển thất bại, lí do: {e}")
                return redirect(f"/muc6_1")
            elif loaidieuchuyen=="Thai sản đi làm lại":
                try:
                    ketqua = dichuyendilamlai(mst,
                                    vitricu,
                                    vitrimoi,
                                    chuyencu,
                                    chuyenmoi,
                                    gradecodecu,
                                    gradecodemoi,
                                    hccategorycu,
                                    hccategorymoi,
                                    ngaydieuchuyen,
                                    'Thai sản đi làm lại'
                            )
                    if ketqua["ketqua"]:
                        flash("Điều chuyển thành công !!!")
                    else:
                        flash(f"Điều chuyển thất bại, lí do: {ketqua['lido']}, query: {ketqua['query']} !!!")
                except Exception as e:
                    flash(f"Điều chuyển thất bại, lí do: {e}")
                return redirect(f"/muc6_1")
            elif loaidieuchuyen=="Tạm hoãn hợp đồng":
                try:
                    ketqua = dichuyennghi(mst,
                                vitricu,
                                chuyencu,
                                gradecodecu,
                                hccategorycu,
                                ngaydieuchuyen,
                                'Tạm hoãn hợp đồng'
                                )
                    if ketqua["ketqua"]:
                        flash("Điều chuyển thành công !!!")
                    else:
                        flash(f"Điều chuyển thất bại, lí do: {ketqua['lido']}, query: {ketqua['query']} !!!")
                except Exception as e:
                    flash(f"Điều chuyển thất bại, lí do: {e}")
                return redirect(f"/muc6_1")
            elif loaidieuchuyen=="Đi làm lại":
                try:
                    ketqua = dichuyendilamlai(mst,
                                vitricu,
                                vitrimoi,
                                chuyencu,
                                chuyenmoi,
                                gradecodecu,
                                gradecodemoi,
                                hccategorycu,
                                hccategorymoi,
                                ngaydieuchuyen,
                                'Đi làm lại'
                            )
                    if ketqua["ketqua"]:
                        flash("Điều chuyển thành công !!!")
                    else:
                        flash(f"Điều chuyển thất bại, lí do: {ketqua['lido']}, query: {ketqua['query']} !!!")
                except Exception as e:
                    flash(f"Điều chuyển thất bại, lí do: {e}")
                return redirect(f"/muc6_1")
            return redirect(f"/muc6_1")
        elif request.method == "GET":
            cacvitri= laycacvitri()
            return render_template("6_1.html",
                            cacvitri=cacvitri,
                            page="6.1 Điều chuyển chức vụ, bộ phận")
    except Exception as e:
        flash(e)
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
        df["Ngày thực hiện"] = to_datetime(df['Ngày thực hiện'])
        df["Ngày chính thức"] = to_datetime(df['Ngày chính thức'])
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
        df["Ngày bắt đầu"] = to_datetime(df['Ngày bắt đầu'])
        df["Ngày kết thúc"] = to_datetime(df['Ngày kết thúc'])
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
    
@app.route("/muc7_1_1", methods=["GET","POST"]) # Đổi ca làm việc
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
        flash("Tải file thành công !!!")
        return send_file(os.path.join(FOLDER_XUAT, f"doica_{thoigian}.xlsx"), as_attachment=True)
            
@app.route("/muc7_1_2", methods=["GET","POST"]) # Danh sách lỗi chấm công
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
    return _render_with_mobile_fallback(
        "7_1_2.html",
        page="Lỗi chấm công",
        danhsach=paginated_rows,
        pagination=pagination,
        count=count,
    )


@app.route("/muc7_1_3", methods=["GET","POST"]) # Danh sách điểm danh bù
@login_required
def diemdanhbu():
    if request.method == "GET":
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
        danhsach = laydanhsachdiemdanhbu(mst,hoten,chucvu,chuyen,bophan,loaidiemdanh,ngay,lido,trangthai,mstquanly,mstthuky)
        count = len(danhsach)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 20
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return _render_with_mobile_fallback(
            "7_1_3.html",
            page="Lỗi chấm công",
            danhsach=paginated_rows,
            pagination=pagination,
            count=count,
        )
    elif request.method == "POST":
        mstthuky = request.form.get("mstthuky")
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
        
        rows = laydanhsachdiemdanhbu(mst,hoten,chucvu,chuyen,bophan,loaidiemdanh,ngaydiemdanh,lydo,trangthai,mstquanly,mstthuky)
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
                "Trạng thái": row[10],
                "ID":row[11],
                "Thời gian tạo": row[12],
                "Thời gian duyệt": row[13]
            })
        
        df = pd.DataFrame(result)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT, f"diemdanhbu_{thoigian}.xlsx"), index=False) # f"diemdanhbu_{thoigian}.xlsx", index=False)
        
        return send_file(os.path.join(FOLDER_XUAT, f"diemdanhbu_{thoigian}.xlsx"), as_attachment=True)

@app.route("/muc7_1_3/kiemtra", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def kiemtradiemdanhbu():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = thuky_dakiemtra_diemdanhbu(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_3/tuchoi_kiemtra", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def tuchoi_kiemtradiemdanhbu():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = thuky_tuchoi_diemdanhbu(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_3/pheduyet", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def pheduyetdiemdanhbu():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = quanly_pheduyet_diemdanhbu(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_3/tuchoi_pheduyet", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def tuchoi_pheduyetdiemdanhbu():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = quanly_tuchoi_diemdanhbu(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_4", methods=["GET","POST"]) # Danh sách xin nghỉ phép 
@login_required
def xinnghiphep():
    if request.method == "GET":
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
        per_page = 20
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return _render_with_mobile_fallback(
            "7_1_4.html",
            page="Lỗi chấm công",
            danhsach=paginated_rows,
            pagination=pagination,
            count=count,
        )
    elif request.method == "POST":
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
                'Trạng thái': row[9],
                'ID': row[10],
                'Thời gian tạo': row[11],
                'Thời gian duyệt': row[12]
            })
        df = pd.DataFrame(result)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT, f"xinnghiphep_{thoigian}.xlsx"), index=False)
        
        return send_file(os.path.join(FOLDER_XUAT, f"xinnghiphep_{thoigian}.xlsx"), as_attachment=True)

@app.route("/muc7_1_4/kiemtra", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def kiemtraxinnghiphep():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = thuky_dakiemtra_xinnghiphep(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_4/tuchoi_kiemtra", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def tuchoi_kiemtraxinnghiphep():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = thuky_tuchoi_xinnghiphep(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_4/pheduyet", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def pheduyetxinnghiphep():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = quanly_pheduyet_xinnghiphep(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_4/tuchoi_pheduyet", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def tuchoi_pheduyetxinnghiphep():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = quanly_tuchoi_xinnghiphep(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})


@app.route("/muc7_1_5", methods=["GET","POST"]) # Danh sách xin nghỉ không lương
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
        per_page = 20
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return _render_with_mobile_fallback(
            "7_1_5.html",
            page="Lỗi chấm công",
            danhsach=paginated_rows,
            pagination=pagination,
            count=count,
        )
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
        mstthuky = request.form.get("mstthuky")
        danhsach = laydanhsachxinnghikhongluong(mst,hoten,chucvu,chuyen,bophan,ngay,lydo,trangthai,mstquanly,mstthuky)
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
                "Loại nghỉ": row[8],
                "Trạng thái": row[9],
                "ID": row[10],
                "Thời gian tạo": row[11],
                "Thời gian duyệt": row[12]
            })
        df = pd.DataFrame(data)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT, f"xinnghikhongluong_{thoigian}.xlsx"), index=False)
        flash("Tải file thành công !!!")
        return send_file(os.path.join(FOLDER_XUAT, f"xinnghikhongluong_{thoigian}.xlsx"), as_attachment=True)

@app.route("/muc7_1_5/kiemtra", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def kiemtraxinnghikhongluong():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = thuky_dakiemtra_xinnghikhongluong(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_5/tuchoi_kiemtra", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def tuchoi_kiemtraxinnghikhongluong():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = thuky_tuchoi_xinnghikhongluong(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_5/pheduyet", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def pheduyetxinnghikhongluong():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = quanly_pheduyet_xinnghikhongluong(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_5/tuchoi_pheduyet", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def tuchoi_pheduyetxinnghikhongluong():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = quanly_tuchoi_xinnghikhongluong(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})
        
@app.route("/muc7_1_6", methods=["GET","POST"]) # Danh sách xin nghỉ khác
@login_required
def danhsachxinnghikhac():
    if request.method == "GET":
        mstthuky = request.args.get("mstthuky")
        mstquanly = request.args.get("mstquanly")
        mst = request.args.get("mst")
        chuyen = request.args.get("chuyen")
        bophan = request.args.get("bophan")
        ngaynghi = request.args.get("ngaynghi")
        loainghi = request.args.get("loainghi")
        trangthai = request.args.get("trangthai")
        nhangiayto = request.args.get("nhangiayto")
        danhsach = laydanhsachxinnghikhac(mst,chuyen,bophan,ngaynghi,loainghi,trangthai,nhangiayto,mstthuky,mstquanly)
        count = len(danhsach)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 20
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return _render_with_mobile_fallback(
            "7_1_6.html",
            page="Lỗi chấm công",
            danhsach=paginated_rows,
            pagination=pagination,
            count=count,
        )
    elif request.method == "POST":
        mstthuky = request.form.get("mstthuky")
        mstquanly = request.form.get("mstquanly")
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        ngaynghi = request.form.get("ngaynghi")
        loainghi = request.form.get("loainghi")
        trangthai = request.form.get("trangthai")
        nhangiayto = request.form.get("nhangiayto")
        danhsach = laydanhsachxinnghikhac(mst,chuyen,bophan,ngaynghi,loainghi,trangthai,nhangiayto,mstthuky,mstquanly)
        data = [{
            "Nhà máy": row[0],
            "Mã số thẻ": row[1],
            "Họ tên": row[2],
            "Chức danh": row[3],
            "Chuyền": row[4],
            "Bộ phận": row[5],
            "Ngày nghỉ": row[6],
            "Tổng số phút": row[7],
            "Loại nghỉ": row[8],
            "Trạng thái": row[9],
            "Nhận giấy tờ": row[10],  
            "ID": row[11],
            "Thời gian tạo": row[12],
            "Thời gian duyệt": row[13]          
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

@app.route("/muc7_1_6/kiemtra", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def kiemtraxinnghikhac():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = thuky_dakiemtra_xinnghikhac(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_6/tuchoi_kiemtra", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def tuchoi_kiemtraxinnghikhac():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = thuky_tuchoi_xinnghikhac(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_6/pheduyet", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def pheduyetxinnghikhac():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = quanly_pheduyet_xinnghikhac(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_6/tuchoi_pheduyet", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def tuchoi_pheduyetxinnghikhac():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = quanly_tuchoi_xinnghikhac(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_6/nhan_giayto", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def nhan_giaytoxinnghikhac():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = nhansu_nhangiayto_xinnghikhac(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_6/khongnhan_giayto", methods=["POST"]) # Danh sách điểm danh bù
@login_required
def khongnhan_giaytoxinnghikhac():
    if request.method == "POST":
        try:
            data = request.form.getlist("selected_ids[]")
            result = []
            x=0
            for item in data:
                x = nhansu_khongnhangiayto_xinnghikhac(item)
                result.append({
                    "id": item,
                    "status": x
                })
            result = {"success": True, "result": result}
            return jsonify(result)
        except Exception as e:
            print(f"Error: {e}")
            return jsonify({"success": False, "message": str(e)})
    return jsonify({"success": False, "message": "Invalid request method"})

@app.route("/muc7_1_7", methods=["GET","POST"]) # Danh sách phép tồn
@login_required
def muc7_1_7():
    if request.method == "GET":
        mst = request.args.get("mst")
        thang = request.args.get("thang")
        nam = request.args.get("nam")
        danhsach = laydanhsachphepton(mst,thang,nam)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_7.html", page="Lỗi chấm công",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    if request.method == "POST":
        mst = request.form.get("mst")
        thang = request.form.get("thang")
        nam = request.form.get("nam")
        danhsach = laydanhsachphepton(mst,thang,nam)
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
        response.headers['Content-Disposition'] = f'attachment; filename=danhsachphepton_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response

@app.route("/muc7_1_7mobile", methods=["GET","POST"]) # Danh sách phép tồn
@login_required
def muc7_1_7mobile():
    nhamay = current_user.macongty
    mst = request.args.get("mst") or current_user.masothe
    today = datetime.today()
    thang = request.args.get("thang", type=int) or today.month
    nam = request.args.get("nam", type=int) or today.year

    # hiện tại đang reuse laydanhsachphepton(mst) -> trả list tất cả tháng/năm
    danhsach = laydanhsachphepton(mst,thang,nam)

    return _render_with_mobile_fallback(
    "mobile/7_1_7.html",
    mst=mst,
    thang=thang,
    nam=nam,
    danhsach=danhsach,
    page="Phép tồn"
    )   

@app.route("/muc7_1_8", methods=["GET","POST"]) # Đăng ký làm thêm giờ
@login_required
def muc7_1_8():
    
    if request.method == "GET":
        mst = request.args.get("mst")
        phongban = request.args.get("phongban")
        chuyen = request.args.get("chuyen")
        ngay = request.args.get("ngay")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        backday = True
        danhsach = laydanhsachtangca(mst,phongban,chuyen,ngay,tungay,denngay,backday)
        count = len(danhsach)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(danhsach)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_8.html", 
                               page="Làm thêm giờ",
                               danhsach=paginated_rows,
                               pagination=pagination,
                               count=count
                               )
    elif request.method == "POST":
        data = []
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        phongban = request.form.get("phongban")
        ngay = request.form.get("ngay")
        tungay = request.form.get("tungay")
        denngay = request.form.get("denngay")
        danhsach = laydanhsachtangca(mst,phongban,chuyen,ngay,tungay,denngay)
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
        response.headers['Content-Disposition'] = f'attachment; filename=dangkytangca_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response

@app.route("/muc7_1_9", methods=["GET","POST"]) # Bảng làm thêm giờ chế độ
@login_required
def muc7_1_9():
    if request.method == "GET":
        thang = int(request.args.get("thang")) if request.args.get("thang") else 0
        nam = int(request.args.get("nam")) if request.args.get("nam") else 0
        mst = request.args.get("mst")
        bophan = request.args.get("bophan")
        chuyen = request.args.get("chuyen")
        danhsach = lay_tangcachedo(thang,nam,mst,bophan,chuyen)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_9.html", page="Làm thêm giờ",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        thang = request.form.get("thang") if request.form.get("thang") else datetime.now().month
        nam = request.form.get("nam") if request.form.get("nam") else datetime.now().year
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcachedo(thang,nam,mst,bophan,chuyen)
        workbook = openpyxl.load_workbook(FILE_MAU_LAMTHEMGIO_CHEDO_KX)

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

@app.route("/muc7_1_10", methods=["GET","POST"]) # Danh sách làm thêm giờ ban ngày
@login_required
def muc7_1_10():
    if request.method == "GET":
        thang = int(request.args.get("thang")) if request.args.get("thang") else 0
        nam = int(request.args.get("nam")) if request.args.get("nam") else 0
        mst = request.args.get("mst")
        bophan = request.args.get("bophan")
        chuyen = request.args.get("chuyen")
        danhsach = lay_tangcangay(thang,nam,mst,bophan,chuyen)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_10.html", page="Làm thêm giờ",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        thang = request.form.get("thang") if request.form.get("thang") else datetime.now().month
        nam = request.form.get("nam") if request.form.get("nam") else datetime.now().year
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcangay(thang,nam,mst,bophan,chuyen)
        workbook = openpyxl.load_workbook(FILE_MAU_LAMTHEMGIO_BANNGAY_KX)

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

@app.route("/muc7_1_11", methods=["GET","POST"]) # Danh sách làm thêm giờ ban đêm
@login_required
def muc7_1_11():
    if request.method == "GET":
        thang = request.args.get("thang")
        nam = request.args.get("nam")
        mst = request.args.get("mst")
        bophan = request.args.get("bophan")
        chuyen = request.args.get("chuyen")
        danhsach = lay_tangcadem(thang,nam,mst,bophan,chuyen)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_11.html", page="Làm thêm giờ",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        thang = request.form.get("thang")
        nam = request.form.get("nam")
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcadem(thang,nam,mst,bophan,chuyen)
        workbook = openpyxl.load_workbook(FILE_MAU_LAMTHEMGIO_BANDEM_KX)

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

@app.route("/muc7_1_12", methods=["GET","POST"]) # Danh sách làm thêm giờ Chủ nhật
@login_required
def muc7_1_12():
    if request.method == "GET":
        thang = request.args.get("thang") if request.args.get("thang") else datetime.now().month
        nam = request.args.get("nam") if request.args.get("nam") else datetime.now().year
        mst = request.args.get("mst")
        bophan = request.args.get("bophan")
        chuyen = request.args.get("chuyen")
        danhsach = lay_tangcachunhat(thang,nam,mst,bophan,chuyen)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_12.html", page="Làm thêm giờ",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        thang = request.form.get("thang")
        nam = request.form.get("nam")
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcachunhat(thang,nam,mst,bophan,chuyen)
        workbook = openpyxl.load_workbook(FILE_MAU_LAMTHEMGIO_CHUNHAT_KX)

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
        
@app.route("/muc7_1_13", methods=["GET","POST"]) # Danh sách làm thêm giờ ngày lễ
@login_required
def muc7_1_13():
    if request.method == "GET":
        thang = request.args.get("thang") if request.args.get("thang") else datetime.now().month
        nam = request.args.get("nam") if request.args.get("nam") else datetime.now().year
        mst = request.args.get("mst")
        bophan = request.args.get("bophan")
        chuyen = request.args.get("chuyen")
        danhsach = lay_tangcangayle(thang,nam,mst,bophan,chuyen)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_13.html", page="Làm thêm giờ",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
    elif request.method == "POST":
        thang = request.form.get("thang")
        nam = request.form.get("nam")
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_tangcangayle(thang,nam,mst,bophan,chuyen)
        workbook = openpyxl.load_workbook(FILE_MAU_LAMTHEMGIO_NGAYLE_KX)

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

@app.route("/muc7_1_14", methods=["GET","POST"]) # Bảng chấm công chi tiết chưa chốt
@login_required
def muc7_1_14():
    if request.method=="GET":
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
        return _render_with_mobile_fallback(
            "7_1_14.html",
            page="Bảng chấm công",
            danhsach=paginated_rows,
            pagination=pagination,
            count=count,
        )
    elif request.method=="POST":
        mst = request.form.get('mst')
        chuyen = request.form.get('chuyen')
        phongban = request.form.get('phongban')
        tungay = request.form.get("tungay")
        denngay = request.form.get("denngay")
        phanloai = request.form.get("phanloai")
        danhsach = laydanhsachchamcong(mst,chuyen,phongban,tungay,denngay,phanloai)
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_CHUACHOT_KX)

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
        # Duyệt qua các ô trong khu vực G4:H10000
        for row in range(4, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này            

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chuachot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chuachot_{timestamp}.xlsx"), as_attachment=True)
                
@app.route("/muc7_1_15", methods=["GET","POST"]) # Bảng chấm công chi tiết chốt
@login_required
def muc7_1_15():
    if request.method=="GET":
        mst = request.args.get("mst")
        chuyen = request.args.get('chuyen')
        phongban = request.args.get("phongban")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        phanloai = request.args.get("phanloai")
        rows = laydanhsachchamcongchot(mst,chuyen,phongban,tungay,denngay,phanloai)
        count = len(rows)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(rows)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = rows[start:end]
        danhsachphongban = laycacphongban()
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_15.html", page="Bảng chấm công",
                            danhsach=paginated_rows, 
                            pagination=pagination,
                            count=count,
                            danhsachphongban=danhsachphongban)
    elif request.method=="POST":
        mst = request.form.get('mst')
        chuyen = request.form.get('chuyen')
        phongban = request.form.get('phongban')
        tungay = request.form.get("tungay")
        denngay = request.form.get("denngay")
        phanloai = request.form.get("phanloai")
        danhsach = laydanhsachchamcongchot(mst,chuyen,phongban,tungay,denngay,phanloai)
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_CHOT_KX)

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
        sheet.delete_rows(4, 50000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            data[7] = datetime.strptime(data[7],"%Y-%m-%d")
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        number_style = NamedStyle(name="number_style", number_format="0")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 50001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chot_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc7_1_16", methods=["GET","POST"]) # Bảng chấm công hành chính
@login_required
def muc7_1_16():
    
    if request.method == "GET":
        thang = int(request.args.get("thang")) if request.args.get("thang") else 0
        nam = int(request.args.get("nam")) if request.args.get("nam") else 0
        mst = request.args.get("mst")
        bophan = request.args.get("bophan")
        chuyen = request.args.get("chuyen")
        danhsach = lay_bangcong_kx(thang,nam,mst,bophan,chuyen)
        total = len(danhsach)
        page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 15
        start = (page - 1) * per_page
        end = start + per_page
        paginated_rows = danhsach[start:end]
        pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_16.html", page="Bảng chấm công",
                                danhsach=paginated_rows, 
                                pagination=pagination,
                                count=total)
        
    elif request.method == "POST":
        thang = request.form.get("thang")
        nam = request.form.get("nam")
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        danhsach = lay_bangcong_kx(thang,nam,mst,bophan,chuyen)
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_HANHCHINH_KX)

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
    
@app.route("/muc7_1_17", methods=["GET","POST"]) # Bảng chấm công tổng hợp
def muc7_1_17():
    if request.method == "GET":
        try:
            thang = int(request.args.get("thang")) if request.args.get("thang") else 0
            nam = int(request.args.get("nam")) if request.args.get("nam") else 0
            mst = request.args.get("mst")
            bophan = request.args.get("bophan")
            chuyen = request.args.get("chuyen")
            # if (nam > 2025 or (nam == 2025 and thang > 6)):
            #     danhsach = lay_bangcongthang_kx_sau_072025(mst,bophan,chuyen,thang,nam)
            # else:
            #     danhsach = lay_bangcongthang_kx(mst,bophan,chuyen,thang,nam)
            danhsach = lay_bangcongthang_kx_sau_072025(mst,bophan,chuyen,thang,nam)
            count = len(danhsach)
            page = request.args.get(get_page_parameter(), type=int, default=1)
            per_page = 15
            start = (page - 1) * per_page
            end = start + per_page
            paginated_rows = danhsach[start:end]
            pagination = Pagination(page=page, per_page=per_page, total=count, css_framework='bootstrap4')
            
            # if (nam > 2025 or (nam == 2025 and thang > 6)) or (nam == 0 and thang == 0):
            #     return render_template("7_1_17_sau_072025.html", page="Bảng chấm công",
            #                         danhsach=paginated_rows, 
            #                         pagination=pagination,
            #                         count=count)
            # else:
            #     return render_template("7_1_17.html", page="Bảng chấm công",
            #                         danhsach=paginated_rows, 
            #                         pagination=pagination,
            #                         count=count)

            return render_template("7_1_17_sau_072025.html", page="Bảng chấm công",
                                    danhsach=paginated_rows, 
                                    pagination=pagination,
                                    count=count)
        except Exception as e:
            flash(f"Lỗi tải trang: {e}")
            return render_template("7_1_17_sau_072025.html",
                                    danhsach=[])
    else:
        thang = int(request.form.get("thang")) if request.args.get("thang") else 0
        nam = int(request.form.get("nam")) if request.args.get("nam") else 0
        mst = request.form.get("mst")
        bophan = request.form.get("bophan")
        chuyen = request.form.get("chuyen")
        # if (nam > 2025 or (nam == 2025 and thang > 6) or (nam == 0 and thang == 0)):
        #     danhsach = lay_bangcongthang_kx_sau_072025(mst,bophan,chuyen,thang,nam)
        #     workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_TONGHOP_KX_SAU_072025)
        # else:
        #     danhsach = lay_bangcongthang_kx(mst,bophan,chuyen,thang,nam)
        #     workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_TONGHOP_KX)

        danhsach = lay_bangcongthang_kx_sau_072025(mst,bophan,chuyen,thang,nam)
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_TONGHOP_KX_SAU_072025)

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
            # if (nam < 2025 or (nam == 2025 and thang > 6)):
            #     data = [y for y in row]
            # else:
            #     # Chỉ lấy các cột cần thiết và sắp xếp lại thứ tự
            #     data = [y for y in row[:-7]] + [row[-1]] + [y for y in row[-7:-4]] 
            
            data = [y for y in row]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            # data[-2] = round(data[-2]) if data[-2] else 0
            sheet.append(data)


        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        number_style = NamedStyle(name="number_style", number_format="0.00")
        # if (nam > 2025 or (nam == 2025 and thang > 6)):
        #     # Duyệt qua các ô trong khu vực G7:H10000
        #     for row in range(6, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
        #         for col in ['G', 'H']:
        #             cell = sheet[f"{col}{row}"]
                    
        #             try:
        #                 cell.style = date_style
        #             except ValueError:
        #                 pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
        #         for col in ['J', 'K','L', 'M','N', 'O','P', 'Q','R', 'S','T', 'U', 'X','Y', 'Z','AA','AB', 'AC','AD', 'AE', 'AF','AG', 'AH','AI', 'AJ']:
        #             cell = sheet[f"{col}{row}"]
        #             if cell.value and int(cell.value) > 0:
        #                 try:
        #                     cell.style = number_style
        #                 except ValueError:
        #                     pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            

        #     timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        #     workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"))
        #     return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"), as_attachment=True)
        # else:
        #     # Duyệt qua các ô trong khu vực G7:H10000
        #     for row in range(6, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
        #         for col in ['G', 'H']:
        #             cell = sheet[f"{col}{row}"]
                    
        #             try:
        #                 cell.style = date_style
        #             except ValueError:
        #                 pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
        #         for col in ['J', 'K','L', 'M','N', 'O','P', 'Q','R', 'S','T', 'U', 'W', 'X','Y', 'Z','AA','AB', 'AC','AD', 'AE', 'AF','AG', 'AH','AI', 'AJ', 'AK']:
        #             cell = sheet[f"{col}{row}"]
        #             if cell.value and int(cell.value) > 0:
        #                 try:
        #                     cell.style = number_style
        #                 except ValueError:
        #                     pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            

        #     timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        #     workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"))
        #     return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"), as_attachment=True)
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(6, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            for col in ['J', 'K','L', 'M','N', 'O','P', 'Q','R', 'S','T', 'U', 'X','Y', 'Z','AA','AB', 'AC','AD', 'AE', 'AF','AG', 'AH','AI', 'AJ']:
                cell = sheet[f"{col}{row}"]
                if cell.value and int(cell.value) > 0:
                    try:
                        cell.style = number_style
                    except ValueError:
                        pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"), as_attachment=True)
    
@app.route("/muc7_1_18", methods=["GET","POST"]) # Bảng chấm công chi tiết chốt quá khứ
@login_required
def muc7_1_18():
    if request.method=="GET":
        mst = request.args.get("mst")
        chuyen = request.args.get('chuyen')
        phongban = request.args.get("phongban")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        phanloai = request.args.get("phanloai")
        rows = laydanhsachchamcongchotquakhu(mst,chuyen,phongban,tungay,denngay,phanloai)
        count = len(rows)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(rows)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = rows[start:end]
        danhsachphongban = laycacphongban()
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_18.html", page="Bảng chấm công",
                            danhsach=paginated_rows, 
                            pagination=pagination,
                            count=count,
                            danhsachphongban=danhsachphongban)
    elif request.method=="POST":
        mst = request.form.get('mst')
        chuyen = request.form.get('chuyen')
        phongban = request.form.get('phongban')
        tungay = request.form.get("tungay")
        denngay = request.form.get("denngay")
        phanloai = request.form.get("phanloai")
        danhsach = laydanhsachchamcongchotquakhu(mst,chuyen,phongban,tungay,denngay,phanloai)
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_CHOT_KX)

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
        sheet.delete_rows(4, 50000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            data[7] = datetime.strptime(data[7],"%Y-%m-%d")
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 50001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chot_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc7_1_19", methods=["GET","POST"]) # Bảng chấm công chi tiết chưa chốt
@login_required
def muc7_1_19():
    if request.method=="GET":
        mst = request.args.get("mst")
        chuyen = request.args.get("chuyen")
        phongban = request.args.get("phongban")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        phanloai = request.args.get("phanloai")
        rows = laydanhsachchamcongchunhatchuachot(mst, chuyen, phongban, tungay, denngay, phanloai)
        count = len(rows)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(rows)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = rows[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_19.html", page="Bảng chấm công Chủ Nhật chi tiết chưa chốt",
                            danhsach=paginated_rows, 
                            pagination=pagination,
                            count=count)
    elif request.method=="POST":
        try:
            mst = request.form.get('mst')
            chuyen = request.form.get('chuyen')
            phongban = request.form.get('phongban')
            tungay = request.form.get("tungay")
            denngay = request.form.get("denngay")
            phanloai = request.form.get("phanloai")
            danhsach = laydanhsachchamcongchunhatchuachot(mst,chuyen,phongban,tungay,denngay,phanloai)
            workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_CHUNHAT_CHUACHOT_KX)

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
                data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
                sheet.append(data)

            # Tạo kiểu định dạng ngày
            date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
            # Duyệt qua các ô trong khu vực G4:H10000
            for row in range(4, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
                for col in ['G']:
                    cell = sheet[f"{col}{row}"]
                    
                    try:
                        cell.style = date_style
                    except ValueError:
                        pass  # Nếu giá trị không phải là ngày, bỏ qua ô này            

            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chunhat_chuachot_{timestamp}.xlsx"))
            return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chunhat_chuachot_{timestamp}.xlsx"), as_attachment=True)
        except Exception as e:
            flash(f"Lỗi tải trang: {e}")
            return render_template("7_1_19.html",
                                    danhsach=[])
@app.route("/muc7_1_20", methods=["GET","POST"]) # Bảng chấm công chi tiết chốt
@login_required
def muc7_1_20():
    if request.method=="GET":
        mst = request.args.get("mst")
        chuyen = request.args.get('chuyen')
        phongban = request.args.get("phongban")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        phanloai = request.args.get("phanloai")
        rows = laydanhsachchamcongchunhatchot(mst, chuyen, phongban, tungay, denngay, phanloai)
        count = len(rows)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(rows)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = rows[start:end]
        danhsachphongban = laycacphongban()
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_20.html", page="Bảng chấm công",
                            danhsach=paginated_rows, 
                            pagination=pagination,
                            count=count,
                            danhsachphongban=danhsachphongban)
    elif request.method=="POST":
        mst = request.form.get('mst')
        chuyen = request.form.get('chuyen')
        phongban = request.form.get('phongban')
        tungay = request.form.get("tungay")
        denngay = request.form.get("denngay")
        phanloai = request.form.get("phanloai")
        danhsach = laydanhsachchamcongchunhatchot(mst,chuyen,phongban,tungay,denngay,phanloai)
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_CHUNHAT_CHOT_KX)

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
        sheet.delete_rows(4, 50000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d")
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        number_style = NamedStyle(name="number_style", number_format="0")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 50001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['F']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chunhat_chitiet_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chunhat_chitiet_chot_{timestamp}.xlsx"), as_attachment=True)


@app.route("/muc7_1_21", methods=["GET","POST"]) # Bảng chấm công chi tiết chốt quá khứ
@login_required
def muc7_1_21():
    if request.method=="GET":
        mst = request.args.get("mst")
        chuyen = request.args.get('chuyen')
        phongban = request.args.get("phongban")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        phanloai = request.args.get("phanloai")
        rows = laydanhsachchamcongchunhatchotquakhu(mst, chuyen, phongban, tungay, denngay, phanloai)
        count = len(rows)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(rows)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = rows[start:end]
        danhsachphongban = laycacphongban()
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_21.html", page="Bảng chấm công",
                            danhsach=paginated_rows, 
                            pagination=pagination,
                            count=count)
    elif request.method=="POST":
        mst = request.form.get('mst')
        chuyen = request.form.get('chuyen')
        phongban = request.form.get('phongban')
        tungay = request.form.get("tungay")
        denngay = request.form.get("denngay")   
        phanloai = request.form.get("phanloai")
        danhsach = laydanhsachchamcongchunhatchotquakhu(mst,chuyen,phongban,tungay,denngay,phanloai)
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_CHUNHAT_CHOT_KX)

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
        sheet.delete_rows(4, 50000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            # data[7] = datetime.strptime(data[7],"%Y-%m-%d")
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 50001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chunhat_chitiet_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chunhat_chitiet_chot_{timestamp}.xlsx"), as_attachment=True)
    
@app.route("/muc7_1_22", methods=["GET","POST"]) # Bảng chấm công chi tiết ngày lễ chưa chốt
@login_required
def muc7_1_22():
    if request.method=="GET":
        mst = request.args.get("mst")
        chuyen = request.args.get("chuyen")
        phongban = request.args.get("phongban")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        phanloai = request.args.get("phanloai")
        rows = laydanhsachchamcongngaylechuachot(mst, chuyen, phongban, tungay, denngay, phanloai)
        count = len(rows)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(rows)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = rows[start:end]
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_22.html", page="Bảng chấm công ngày lễ chi tiết chưa chốt",
                            danhsach=paginated_rows, 
                            pagination=pagination,
                            count=count)
    elif request.method=="POST":
        try:
            mst = request.form.get('mst')
            chuyen = request.form.get('chuyen')
            phongban = request.form.get('phongban')
            tungay = request.form.get("tungay")
            denngay = request.form.get("denngay")
            phanloai = request.form.get("phanloai")
            danhsach = laydanhsachchamcongngaylechuachot(mst,chuyen,phongban,tungay,denngay,phanloai)
            workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_NGAYLE_CHUACHOT_KX)

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
                data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
                sheet.append(data)

            # Tạo kiểu định dạng ngày
            date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
            # Duyệt qua các ô trong khu vực G4:H10000
            for row in range(4, 10001):  # Bắt đầu từ dòng 7 đến dòng 10000
                for col in ['G']:
                    cell = sheet[f"{col}{row}"]
                    
                    try:
                        cell.style = date_style
                    except ValueError:
                        pass  # Nếu giá trị không phải là ngày, bỏ qua ô này            

            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_ngayle_chuachot_{timestamp}.xlsx"))
            return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_ngayle_chuachot_{timestamp}.xlsx"), as_attachment=True)
        except Exception as e:
            flash(f"Lỗi tải trang: {e}")
            return render_template("7_1_22.html",
                                    danhsach=[])
@app.route("/muc7_1_23", methods=["GET","POST"]) # Bảng chấm công chi tiết ngày lễ chốt
@login_required
def muc7_1_23():
    if request.method=="GET":
        mst = request.args.get("mst")
        chuyen = request.args.get('chuyen')
        phongban = request.args.get("phongban")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        phanloai = request.args.get("phanloai")
        rows = laydanhsachchamcongngaylechot(mst, chuyen, phongban, tungay, denngay, phanloai)
        count = len(rows)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(rows)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = rows[start:end]
        danhsachphongban = laycacphongban()
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_23.html", page="Bảng chấm công chi tiết ngày lễ chốt",
                            danhsach=paginated_rows, 
                            pagination=pagination,
                            count=count,
                            danhsachphongban=danhsachphongban)
    elif request.method=="POST":
        mst = request.form.get('mst')
        chuyen = request.form.get('chuyen')
        phongban = request.form.get('phongban')
        tungay = request.form.get("tungay")
        denngay = request.form.get("denngay")
        phanloai = request.form.get("phanloai")
        danhsach = laydanhsachchamcongngaylechot(mst,chuyen,phongban,tungay,denngay,phanloai)
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_NGAYLE_CHOT_KX)

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
        sheet.delete_rows(4, 50000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d")
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        number_style = NamedStyle(name="number_style", number_format="0")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 50001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['F']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_ngayle_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_ngayle_chot_{timestamp}.xlsx"), as_attachment=True)


@app.route("/muc7_1_24", methods=["GET","POST"]) # Bảng chấm công chi tiết ngày lễ quá khứ
@login_required
def muc7_1_24():
    if request.method=="GET":
        mst = request.args.get("mst")
        chuyen = request.args.get('chuyen')
        phongban = request.args.get("phongban")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        phanloai = request.args.get("phanloai")
        rows = laydanhsachchamcongngaylechotquakhu(mst, chuyen, phongban, tungay, denngay, phanloai)
        count = len(rows)
        current_page = request.args.get(get_page_parameter(), type=int, default=1)
        per_page = 10
        total = len(rows)
        start = (current_page - 1) * per_page
        end = start + per_page
        paginated_rows = rows[start:end]
        danhsachphongban = laycacphongban()
        pagination = Pagination(page=current_page, per_page=per_page, total=total, css_framework='bootstrap4')
        return render_template("7_1_24.html", page="Bảng chấm công chi tiết ngày lễ quá khứ",
                            danhsach=paginated_rows, 
                            pagination=pagination,
                            count=count,
                            danhsachphongban=danhsachphongban)
    elif request.method=="POST":
        mst = request.form.get('mst')
        chuyen = request.form.get('chuyen')
        phongban = request.form.get('phongban')
        tungay = request.form.get("tungay")
        denngay = request.form.get("denngay")   
        phanloai = request.form.get("phanloai")
        danhsach = laydanhsachchamcongngaylechotquakhu(mst,chuyen,phongban,tungay,denngay,phanloai)
        workbook = openpyxl.load_workbook(FILE_MAU_BANGCONG_NGAYLE_CHOT_KX)

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
        sheet.delete_rows(4, 50000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            # data[7] = datetime.strptime(data[7],"%Y-%m-%d")
            sheet.append(data)

        # Tạo kiểu định dạng ngày
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # Duyệt qua các ô trong khu vực G7:H10000
        for row in range(4, 50001):  # Bắt đầu từ dòng 7 đến dòng 10000
            for col in ['H']:
                cell = sheet[f"{col}{row}"]
                
                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Nếu giá trị không phải là ngày, bỏ qua ô này
            
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_ngayle_chitiet_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_ngayle_chitiet_chot_{timestamp}.xlsx"), as_attachment=True)

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
        try:
            mst = request.form.get("mst")
            if not mst:
                flash("Chưa có thông tin người vi phạm")
                return redirect("/muc9_1") 
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
            cacanhvipham = request.files.getlist("file_anh")
            bienbankiluat = request.files.get("file_bienban") 
            os.makedirs(os.path.join(FOLDER_BIENBAN,f"{mst}_{ngayvipham}"),exist_ok=True)
            
            for anh in cacanhvipham:
                anh.save(os.path.join(FOLDER_BIENBAN,f"{mst}_{ngayvipham}",f"{cacanhvipham.index(anh,start=1)}.jpg"))  
            bienbankiluat.save(os.path.join(FOLDER_BIENBAN,f"{mst}_{ngayvipham}"),"bienban.pdf")
            if themdanhsachkyluat(mst,hoten,chucvu,bophan,chuyento,ngayvao,ngayvipham,diadiem,ngaylapbienban,noidung,bienphap):
                flash("Thêm biên bản kỷ luật thành công !!!")
            else:
                flash("Thêm biên bản kỷ luật thất bại !!!")
        except Exception as e:
            flash(f"Thêm biên bản kỷ luật thất bại {e}!!!")
        return redirect("/muc9_1") 
    
@app.route("/muc10_1", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def phongvannghiviec():
        
    return render_template("10_1.html", page="10.1 Tổng hợp phỏng vấn nghỉ việc")

@app.route("/muc10_2", methods=["GET","POST"])
@login_required
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
            flash("Thêm đơn xin nghỉ thành công !!!")
        else:
            flash("Thêm đơn xin nghỉ thất bại !!!")
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
            flash(f"Lỗi tải trang: {e}")
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

@app.route("/admin", methods=["GET"])
@login_required
@roles_required('sa')
def admin_page():
    trangthai = trang_thai_function_12()
    rows = lay_lich_su_dong_bo_cham_cong()
    return render_template("admin.html",trangthai=trangthai, rows=rows)