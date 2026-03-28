# -*- encoding: utf-8 -*-

from app import *

##################################
#          MAIN ROUTES           #
##################################

# from functools import wraps
from flask import g, flash, request, render_template
from flask_login import current_user
from jinja2 import TemplateNotFound

# â"?â"? Helpers â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?

def _sum_don(counts: dict) -> dict:
    """ThAªm key 'Tá»ng' vA o dict Ä`áº¿m Ä`Æ¡n."""
    counts["Tá»ng"] = sum(counts.values())
    return counts


def _lay_don_ca_nhan(macongty, masothe) -> dict:
    """
    Gom táº¥t cáº£ thA'ng tin cA¡ nhA¢n vA o 1 láºn â?" lA½ tÆ°á»Yng nháº¥t nAªn
    gá»Tp thA nh 1 stored-procedure / query tráº£ vá»? nhiá»?u result-set.
    Hiá»╪n táº¡i váº«n gá»?i hA m cÅc nhÆ°ng Ä`A£ tA¡ch riAªng Ä`á»ƒ dá». tá»`i Æ°u sau.
    """
    def _nhom(fn_chua, fn_da, fn_duyet, fn_tuchoi):
        c = fn_chua(macongty, masothe)
        d = fn_da(macongty, masothe)
        p = fn_duyet(macongty, masothe)
        r = fn_tuchoi(macongty, masothe)
        return {"ChÆ°a kiá»ƒm tra": c, "Ä?A£ kiá»ƒm tra": d,
                "Ä?A£ phAª duyá»╪t": p, "Bá»< tá»« chá»`i": r, "Tá»ng": c + d + p + r}

    ddb  = _nhom(lay_soluong_diemdanhbu_chuakiemtra,   lay_soluong_diemdanhbu_dakiemtra,
                 lay_soluong_diemdanhbu_dapheduyet,    lay_soluong_diemdanhbu_bituchoi)
    nphep = _nhom(lay_soluong_xinnghiphep_chuakiemtra,  lay_soluong_xinnghiphep_dakiemtra,
                  lay_soluong_xinnghiphep_dapheduyet,   lay_soluong_xinnghiphep_bituchoi)
    nkl   = _nhom(lay_soluong_xinnghikhongluong_chuakiemtra, lay_soluong_xinnghikhongluong_dakiemtra,
                  lay_soluong_xinnghikhongluong_dapheduyet,  lay_soluong_xinnghikhongluong_bituchoi)
    nkhac = _nhom(lay_soluong_xinnghikhac_chuakiemtra, lay_soluong_xinnghikhac_dakiemtra,
                  lay_soluong_xinnghikhac_dapheduyet,  lay_soluong_xinnghikhac_bituchoi)

    return {
        "Ä?iá»ƒm danh bA1":       ddb,
        "Xin nghá»% phAcp":      nphep,
        "Xin nghá»% khA'ng lÆ°Æ¡ng": nkl,
        "Xin nghá»% khA¡c":      nkhac,
        "Tá»ng":               ddb["Tá»ng"] + nphep["Tá»ng"] + nkl["Tá»ng"] + nkhac["Tá»ng"],
        "Lá»-i cháº¥m cA'ng":      lay_soluong_loichamcong(macongty, masothe),
    }


# Map phanquyen â+' phA²ng ban cáºn truyá»?n (None = toA n cA'ng ty)
_TUYEN_DUNG_PHONGBAN = {
    "gd":  None,
    "td":  None,
    "sa":  None,
}

def _lay_tuyen_dung(macongty, phanquyen, phongban) -> dict:
    """
    Tráº£ vá»? dict thA'ng bA¡o tuyá»ƒn dá»¥ng theo quyá»?n.
    gd  â+' chá»% cáºn Ä`áº¿m 'chá»? phAª duyá»╪t'
    tbp / thÆ° kA½ â+' theo phA²ng ban
    td / sa â+' toA n cA'ng ty
    """
    if phanquyen == "gd":
        return {"Tuyá»ƒn dá»¥ng chá»? phAª duyá»╪t": lay_soluong_yeucautuyendung_chopheduyet(macongty, None)}

    # XA¡c Ä`á»<nh scope phA²ng ban
    if phanquyen in ("td", "sa"):
        pb = None
    elif phanquyen == "tbp" or kiemtra_danhsach_thuki():
        pb = phongban
    else:
        return {}

    return {
        "Tuyá»ƒn dá»¥ng chá»? kiá»ƒm tra":  lay_soluong_yeucautuyendung_chokiemtra(macongty, pb),
        "Tuyá»ƒn dá»¥ng chá»? phAª duyá»╪t": lay_soluong_yeucautuyendung_chopheduyet(macongty, pb),
        "Tuyá»ƒn dá»¥ng Ä`Æ°á»£c duyá»╪t":    lay_soluong_yeucautuyendung_dapheduyet(macongty, pb),
        "Tuyá»ƒn dá»¥ng bá»< tá»« chá»`i":    lay_soluong_yeucautuyendung_bituchoi(macongty, pb),
    }


# â"?â"? before_request â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?

@app.before_request
def run_before_every_request():
    """Kiá»ƒm tra Ä`Äƒng nháº-p, gom thA'ng bA¡o vA o g.notice."""
    if not current_user.is_authenticated:
        return

    f12  = trang_thai_function_12()
    mact = current_user.macongty
    mast = current_user.masothe

    notice = {"f12": f12, "db": url_database_pyodbc, "Tá»ng": 0}

    try:
        # â"?â"? Quáº£n lA½ â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
        if la_quanly(mact, mast):
            ql = {
                "Ä?iá»ƒm danh bA1":       lay_soluong_diemdanhbu_quanly_canduyet(mact, mast),
                "Xin nghá»% phAcp":      lay_soluong_xinnghiphep_quanly_canduyet(mact, mast),
                "Xin nghá»% khA'ng lÆ°Æ¡ng": lay_soluong_xinnghikhongluong_quanly_canduyet(mact, mast),
                "Xin nghá»% khA¡c":      lay_soluong_xinnghikhac_quanly_canduyet(mact, mast),
            }
            ql["Sá»` thA'ng bA¡o"] = sum(ql.values())
            notice["Quáº£n lA½"]  = ql
            notice["Tá»ng"]    += ql["Sá»` thA'ng bA¡o"]
        else:
            notice["Quáº£n lA½"] = {}

        # â"?â"? ThÆ° kA½ â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
        if la_thuky(mact, mast):
            chuyen = lay_danhsach_chuyen_thuky_quanly(mact, mast)
            tk = {
                "Danh sA¡ch lá»-i tháº»":    lay_soluong_loithe_thuky_canxuly(mact, mast),
                "Ä?iá»ƒm danh bA1":         lay_soluong_diemdanhbu_thuky_cankiemtra(mact, mast),
                "Xin nghá»% phAcp":        lay_soluong_xinnghiphep_thuky_cankiemtra(mact, mast),
                "Xin nghá»% khA'ng lÆ°Æ¡ng": lay_soluong_xinnghikhongluong_thuky_cankiemtra(mact, mast),
                "Xin nghá»% khA¡c":        lay_soluong_xinnghikhac_thuky_cankiemtra(mact, mast),
                "Line":                 chuyen[0] if len(chuyen) == 1 else "",
            }
            tk["Sá»` thA'ng bA¡o"] = (tk["Danh sA¡ch lá»-i tháº»"] + tk["Ä?iá»ƒm danh bA1"]
                                  + tk["Xin nghá»% phAcp"] + tk["Xin nghá»% khA'ng lÆ°Æ¡ng"])
            notice["ThÆ° kA½"]   = tk
            notice["Tá»ng"]    += tk["Sá»` thA'ng bA¡o"] + tk["Xin nghá»% khA¡c"]
        else:
            notice["ThÆ° kA½"] = {}

        # â"?â"? CA¡ nhA¢n â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
        notice["personal"] = _lay_don_ca_nhan(mact, mast)

        # â"?â"? Tuyá»ƒn dá»¥ng â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
        td = _lay_tuyen_dung(mact, current_user.phanquyen, current_user.phongban)
        for k, v in td.items():
            if v > 0:
                notice[k]       = v
                notice["Tá»ng"] += v
            else:
                notice.setdefault(k, 0)

    except Exception as e:
        flash(f"Lá»-i cáº-p nháº-t thA'ng tin chuA'ng: {e}")
        notice = {"f12": f12, "db": url_database_pyodbc}

    g.notice = notice


@app.context_processor
def inject_notice():
    return dict(notice=getattr(g, "notice", {}),
                personal=getattr(g, "personal", {}))

def _is_mobile() -> bool:
    """PhA¡t hiá»╪n truy cáº-p tá»« thiáº¿t bá»< di Ä`á»Tng dá»±a vA o User-Agent (Ä`Æ¡n giáº£n)."""
    ua = (request.user_agent.string or "").lower()
    return any(x in ua for x in ("iphone", "android", "ipad"))


def _render_with_mobile_fallback(default_template: str, **context):
    """Thá»- render template mobile/..., náº¿u khA'ng cA3 thA¬ dA1ng template máº·c Ä`á»<nh.

    VA- dá»¥: default_template="home.html" â+' Æ°u tiAªn "mobile/home.html".
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
    return render_template_string("<h1>Báº¡n khA'ng thá»ƒ vA o má»¥c nA y, vui lA²ng chá»?n má»¥c khA¡c!!!</h1><h3>áºn vA o <a href='/'>Ä`A¢y</a> Ä`á»ƒ quay láº¡i trang chá»</h3>")

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
        flash("Sai thA'ng tin Ä`Äƒng nháº-p.", "danger")
        return redirect(url_for("login"))

    return _render_with_mobile_fallback("login.html")

@app.route("/logout", methods=["POST"])
def logout():
    try:
        app.logger.info(f"Nguoi dung {current_user.masothe} o {current_user.macongty} vua  dang xuat !!!")
        logout_user()
    except Exception as e:
        app.logger.error(f"Khong the dang xuat {e} !!!")
    return redirect("/")

@app.route("/doimatkhau", methods=['POST'])
def doimatkhau():
    macongty = request.form.get("macongty")
    masothe = request.form.get("masothe_doi")
    matkhaumoi = request.form.get("matkhaumoi")
    try:
        if doimatkhautaikhoan(macongty,masothe,matkhaumoi):
            flash("Ä?á»i máº-t kháºcu thA nh cA'ng")
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
        mst = request.args.get("MA£ sá»` tháº»")
        hoten = request.args.get("Há»? tAªn")
        sdt = request.args.get("Sá»` Ä`iá»╪n thoáº¡i")
        cccd = request.args.get("CÄƒn cÆ°á»>c cA'ng dA¢n")
        gioitinh = request.args.get("Giá»>i tA-nh")
        vaotungay = request.args.get("VA o tá»« ngA y")
        vaodenngay = request.args.get("VA o Ä`áº¿n ngA y")
        nghitungay = request.args.get("Nghá»% tá»« ngA y")
        nghidenngay = request.args.get("Nghá»% Ä`áº¿n ngA y")
        phongban = request.args.get("PhA²ng ban")
        chucvu = request.args.get("Chá»cc danh")
        trangthai = request.args.get("Tráº¡ng thA¡i")
        hccategory = request.args.get("HC Category")
        ghichu = request.args.get("Ghi chAº")
        chuyen = request.args.get("Chuyá»?n")
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
        flash(f"Xin chA o {current_user.hoten} !!!")
        return _render_with_mobile_fallback(
            "home.html",
            users=paginated_users,
            page="Trang chá»",
            pagination=pagination,
            count=count,
            songuoi_danglamviec=songuoi_danglamviec,
            songuoi_dangnghithaisan=songuoi_dangnghithaisan,
        )
    else:
        try:
            mst = request.form.get("MA£ sá»` tháº»")
            hoten = request.form.get("Há»? tAªn")
            sdt = request.form.get("Sá»` Ä`iá»╪n thoáº¡i")
            cccd = request.form.get("CÄƒn cÆ°á»>c cA'ng dA¢n")
            gioitinh = request.form.get("Giá»>i tA-nh")
            vaotungay = request.form.get("VA o tá»« ngA y")
            vaodenngay = request.form.get("VA o Ä`áº¿n ngA y")
            nghitungay = request.form.get("Nghá»% tá»« ngA y")
            nghidenngay = request.form.get("Nghá»% Ä`áº¿n ngA y")
            phongban = request.form.get("PhA²ng ban")
            chucvu = request.form.get("Chá»cc danh")
            trangthai = request.form.get("Tráº¡ng thA¡i")
            hccategory = request.form.get("Headcount Category")
            ghichu = request.form.get("Ghi chAº")
            chuyen = request.form.get("Chuyá»?n")
            users = laydanhsachuser(mst, hoten, sdt, cccd, gioitinh, vaotungay, vaodenngay, nghitungay, nghidenngay, phongban, trangthai, hccategory, chucvu, ghichu, chuyen)

            # Chuyá»ƒn thA'ng tin ngA y vá»? Ä`á»<nh dáº¡ng YYYY-MM-DD
            for user in users:
                user["NgA y sinh"] = datetime.strptime(user["NgA y sinh"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["NgA y sinh"]!="" else ""
                user["NgA y cáº¥p CCCD"] = datetime.strptime(user["NgA y cáº¥p CCCD"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["NgA y cáº¥p CCCD"]!="" else ""
                user["NgA y kA½ HÄ?"] = datetime.strptime(user["NgA y kA½ HÄ?"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["NgA y kA½ HÄ?"]!="" else ""
                user["NgA y vA o"] = datetime.strptime(user["NgA y vA o"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["NgA y vA o"]!="" else ""
                user["NgA y nghá»%"] = datetime.strptime(user["NgA y nghá»%"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["NgA y nghá»%"]!="" else ""
                user["NgA y háº¿t háº¡n"] = datetime.strptime(user["NgA y háº¿t háº¡n"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["NgA y háº¿t háº¡n"]!="" else ""
                user["NgA y vA o ná»`i thA¢m niAªn"] = datetime.strptime(user["NgA y vA o ná»`i thA¢m niAªn"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["NgA y vA o ná»`i thA¢m niAªn"]!="" else ""
                user["NgA y kA- HÄ? Thá»- viá»╪c"] = datetime.strptime(user["NgA y kA- HÄ? Thá»- viá»╪c"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["NgA y kA- HÄ? Thá»- viá»╪c"]!="" else ""
                user["NgA y háº¿t háº¡n HÄ? Thá»- viá»╪c"] = datetime.strptime(user["NgA y háº¿t háº¡n HÄ? Thá»- viá»╪c"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["NgA y háº¿t háº¡n HÄ? Thá»- viá»╪c"]!="" else ""
                user["NgA y háº¿t háº¡n HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 1"] = datetime.strptime(user["NgA y háº¿t háº¡n HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 1"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["NgA y háº¿t háº¡n HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 1"]!="" else ""
                user["NgA y kA- HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 1"] = datetime.strptime(user["NgA y kA- HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 1"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["NgA y kA- HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 1"]!="" else ""
                user["NgA y kA- HÄ? khA'ng thá»?i háº¡n"] = datetime.strptime(user["NgA y kA- HÄ? khA'ng thá»?i háº¡n"],"%d/%m/%Y").strftime("%Y-%m-%d") if user["NgA y kA- HÄ? khA'ng thá»?i háº¡n"]!="" else ""


            df = pd.DataFrame(users)

            df["NgA y sinh"] = to_datetime(df['NgA y sinh'],errors='coerce')
            df["NgA y cáº¥p CCCD"] = to_datetime(df['NgA y cáº¥p CCCD'],errors='coerce')
            df["NgA y kA½ HÄ?"] = to_datetime(df['NgA y kA½ HÄ?'],errors='coerce')
            df["NgA y vA o"] = to_datetime(df['NgA y vA o'],errors='coerce')
            df["NgA y nghá»%"] = to_datetime(df['NgA y nghá»%'],errors='coerce')
            df["NgA y háº¿t háº¡n"] = to_datetime(df['NgA y háº¿t háº¡n'],errors='coerce')
            df["NgA y vA o ná»`i thA¢m niAªn"] = to_datetime(df['NgA y vA o ná»`i thA¢m niAªn'],errors='coerce')
            df["NgA y sinh con 1"] = to_datetime(df['NgA y sinh con 1'],errors='coerce')
            df["NgA y sinh con 2"] = to_datetime(df['NgA y sinh con 2'],errors='coerce')
            df["NgA y sinh con 3"] = to_datetime(df['NgA y sinh con 3'],errors='coerce')
            df["NgA y sinh con 4"] = to_datetime(df['NgA y sinh con 4'],errors='coerce')
            df["NgA y sinh con 5"] = to_datetime(df['NgA y sinh con 5'],errors='coerce')
            df["NgA y kA- HÄ? Thá»- viá»╪c"] = to_datetime(df['NgA y kA- HÄ? Thá»- viá»╪c'],errors='coerce')
            df["NgA y háº¿t háº¡n HÄ? Thá»- viá»╪c"] = to_datetime(df['NgA y háº¿t háº¡n HÄ? Thá»- viá»╪c'],errors='coerce')
            df["NgA y kA- HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 1"] = to_datetime(df['NgA y kA- HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 1'],errors='coerce')
            df["NgA y háº¿t háº¡n HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 1"] = to_datetime(df['NgA y háº¿t háº¡n HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 1'],errors='coerce')
            # Dòng dưới bị lỗi mã hóa Unicode, tạm thời bỏ qua chuyển đổi cột này để tránh SyntaxError
            # df["Ngay ky HD khong thoi han"] = to_datetime(df['Ngay ky HD khong thoi han'], errors='coerce')
            
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
                        # Apply the date format to column L (assuming 'NgA y thá»±c hiá»╪n' is in column 'L')
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
            flash(f"Lá»-i káº¿t xuáº¥t danh sA¡ch nhA¢n viAªn ({e})")
            app.logger.error(f"Lá»-i káº¿t xuáº¥t danh sA¡ch nhA¢n viAªn ({e})")
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
                                   page="2.1 Danh sA¡ch á»cng viAªn",
                                   danhsach=rows[start: start + per_page],
                                   pagination=pagination,
                                   count=count)
        except Exception as e:
            flash(f"Lá»-i láº¥y danh sA¡ch á»cng viAªn: {e}")
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
            flash("Cap nhat thong tin ung vien thanh cong !!!")
        else:
            flash(f"Cap nhat that bai - {ketqua.get('lido')}")
            app.logger.error(f"muc2_1 POST: {ketqua.get('lido')}")
    except Exception as e:
        flash(f"Lá»-i: {e}")
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
                                page= "2.2 YAªu cáºu tuyá»ƒn dá»¥ng",
                                danhsach = danhsach,
                                lathuki = lathuki,
                                danhsach_vitri_cacongty=danhsach_vitri_cacongty
                                )
        except Exception as e:
            flash(f"Lá»-i láº¥y danh sA¡ch yAªu cáºu tuyá»ƒn dá»¥ng ({e})")
            app.logger.error(f"Lá»-i láº¥y danh sA¡ch yAªu cáºu tuyá»ƒn dá»¥ng ({e})")
            return redirect(url_for("home"))

    elif request.method == "POST":
        try:
            bophan = current_user.phongban
            vitri = request.form.get("vitri")
            if "cA'ng nhA¢n" in vitri.lower():
                kieulaodong = "CA'ng nhA¢n"
            else:
                kieulaodong = "NhA¢n viAªn"
            vitrien = request.form.get("vitrien")
            capbac = request.form.get("capbac")
            soluong = request.form.get("soluong")
            mota = os.path.join(FOLDER_JD, f"{vitrien}.pdf")
            thoigiandukien = request.form.get("thoigiandukien")
            phanloai = request.form.get("phanloai")
            budget = request.form.get("trong_budget")
            trongbudget = "Trong" if budget else"NgoA i"
            if themyeucautuyendungmoi(bophan,vitri,soluong,mota,thoigiandukien,phanloai,capbac,kieulaodong,trongbudget):
                flash("ThAªm yAªu cáºu tuyá»ƒn dá»¥ng má»>i thA nh cA'ng !!!")
                flash(them_thongbao_co_yeucautuyendung(vitri,soluong,trongbudget))
            else:
                flash("ThAªm yAªu cáºu tuyá»ƒn dá»¥ng má»>i tháº¥t báº¡i !!!")
        except Exception as e:
            flash(f"ThAªm yAªu cáºu tuyá»ƒn dá»¥ng má»>i tháº¥t báº¡i ({e})!!!")
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
                if ungvien[16] == "ChÆ°a phá»?ng váº¥n":
                    so_ungvien_chophongvan += 1
                elif ungvien[16] == "Ä?ang phá»?ng váº¥n":
                    so_ungvien_dangphongvan += 1
                elif ungvien[16] == "Qua phá»?ng váº¥n":
                    so_ungvien_quaphongvan += 1
                elif ungvien[16] == "Ä?A£ nháº-n viá»╪c":
                    so_ungvien_danhanviec += 1
                elif ungvien[16] == "KhA'ng nháº-n viá»╪c":
                    so_ungvien_khongnhanviec += 1
            phongban = lay_phongban_theo_idyctd(id_yeucautuyendung)
            return render_template("2_2_1.html",
                                page="2.2.1 Danh sA¡ch á»cng viAªn tuyá»ƒn dá»¥ng",
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
            flash(f"Lá»-i láº¥y danh sA¡ch á»cng viAªn ({e})")
            app.logger.error(f"Lá»-i láº¥y danh sA¡ch á»cng viAªn ({e})")
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
                flash("ThAªm á»cng viAªn thA nh cA'ng")
            return redirect(f"muc2_2_1?id={id_yeucautuyendung}")
        except Exception as e:
            flash(f"Lá»-i thAªm á»cng viAªn ({e})")
            app.logger.error(f"Lá»-i thAªm á»cng viAªn ({e})")
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
                page       = "3.1 Nháº-p thA'ng tin lao Ä`á»Tng má»>i",
                qrcccd     = request.args.get("scan-qrcode"),
                masothe    = masothe,
                ngaybatdau = datetime.now(),
                cacvitri   = cacvitri,
                cacto      = cacto,
                cacca      = cacca,
                macongty   = current_user.macongty,
            )
        except Exception as e:
            flash(f"Lá»-i láº¥y thA'ng tin lao Ä`á»Tng má»>i: {e}")
            app.logger.error(f"muc3_1 GET error: {e}")
            return redirect(url_for("home"))

    # â"?â"? POST â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
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
            app.logger.warning(f"_sql_date: key='{key}' value='{v}' khA'ng há»£p lá»╪ â+' NULL")
            return "NULL"

    try:
        # â"?â"? áº¢nh â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
        anh  = "NULL"
        file = request.files.get("anh")
        if file and file.filename:
            file_path = os.path.join(FOLDER_AVATAR, _s("masothe") + ".jpg")
            file.save(file_path)
            anh = f"'{file_path}'"

        # â"?â"? Ä?á»<nh danh â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
        masothe     = f"'{_s('masothe')}'"
        thechamcong = str(int(_s('masothe')))  # int, khA'ng quotes

        # â"?â"? CA¡ nhA¢n â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
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

        # â"?â"? TA i chA-nh / liAªn há»╪ â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
        nganhang   = _sql_nstr("nganhang")
        sotaikhoan = _sql_str("sotaikhoan")
        dienthoai  = _sql_str("dienthoai")
        sobhxh     = _sql_str("sobhxh")
        masothue   = _sql_str("masothue")

        # â"?â"? Con nhá»? â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
        connho      = _sql_nstr("connho")
        tencon      = [_sql_nstr(f"tenconnho{i}")  for i in range(1, 6)]
        ngaysinhcon = [_sql_date(f"ngaysinhcon{i}") for i in range(1, 6)]

        # â"?â"? Vá»< trA- â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
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

        # â"?â"? COST_ID â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
        cost_id = _sql_str("ntid")

        # â"?â"? Cá»` Ä`á»<nh â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
        luongcoban  = "NULL"
        tongphucap  = "NULL"
        kieuhopdong = "NULL"
        diachimoi   = "NULL"
        nd          = "NULL"  # null date

        # â"?â"? INSERT VALUES (75 cá»Tt Ä`Aºng thá»c tá»±) â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
        nhanvienmoi = (
            # 1-2: Ä?á»<nh danh
            f"({masothe},{thechamcong},"
            # 3-6: CA¡ nhA¢n cÆ¡ báº£n
            f"{hoten},{dienthoai},{ngaysinh},{gioitinh},"
            # 7-11: CCCD + Ä`á»<a chá»% thÆ°á»?ng trAº
            f"{cccd},{ngaycapcccd},{noicapcccd},{cmt},{thuongtru},"
            # 12-15: Ä?á»<a chá»% chi tiáº¿t
            f"{thonxom},{phuongxa},{quanhuyen},{tinhthanhpho},"
            # 16-21: CA¡ nhA¢n khA¡c
            f"{dantoc},{quoctich},{tongiao},{hocvan},{noisinh},{tamtru},"
            # 22-25: TA i chA-nh
            f"{sobhxh},{masothue},{nganhang},{sotaikhoan},"
            # 26: Con nhá»?
            f"{connho},"
            # 27-36: Con nhá»? 1-5
            f"{tencon[0]},{ngaysinhcon[0]},"
            f"{tencon[1]},{ngaysinhcon[1]},"
            f"{tencon[2]},{ngaysinhcon[2]},"
            f"{tencon[3]},{ngaysinhcon[3]},"
            f"{tencon[4]},{ngaysinhcon[4]},"
            # 37-39: áº¢nh, ngÆ°á»?i thA¢n
            f"{anh},{nguoithan},{sdtnguoithan},"
            # 40-42: Há»£p Ä`á»"ng
            f"{kieuhopdong},GETDATE(),{nd},"
            # 43-55: Vá»< trA-
            f"{jobdetailvn},{hccategory},{gradecode},{factory},"
            f"{department},{chucvu},{sectioncode},{sectiondescription},"
            f"{line},{employeetype},{jobdetailen},"
            f"{positioncode},{positioncodedescription},"
            # 56-58: LÆ°Æ¡ng (Luong_co_ban, Phu_cap, Tong_phu_cap)
            f"{luongcoban},{nd},{tongphucap},"
            # 59-63: NgA y thA¡ng hA nh chA-nh
            # Ngay_vao, Ngay_nghi, Trang_thai_lam_viec,
            # Ngay_vao_noi_tham_nien, Mat_khau
            f"GETDATE(),NULL,N'Ä?ang lA m viá»╪c',GETDATE(),'1',"
            # 64-65: HDTV
            f"{nd},{nd},"
            # 66-67: HDXDTH Láºn 1
            f"{nd},{nd},"
            # 68-69: HDXDTH Láºn 2
            f"{nd},{nd},"
            # 70: HDKXDTH
            f"{nd},"
            # 71-74: Truong_BP, Ghi_chu, Time_Stamp, Dia_chi_moi
            f"'N','',GETDATE(),{diachimoi},"
            # 75: COST_ID
            f"{cost_id})"
        )

        app.logger.debug(f"muc3_1 INSERT values: {nhanvienmoi}")

        ketqua = themnhanvienmoi(nhanvienmoi)

        if ketqua["ketqua"]:
            flash("ThAªm lao Ä`á»Tng má»>i thA nh cA'ng !!!")
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
            flash(f"ThAªm lao Ä`á»Tng má»>i tháº¥t báº¡i: {ketqua['lido']}")
            app.logger.error(f"muc3_1 INSERT failed: {ketqua['lido']}")

    except Exception as e:
        flash(f"ThAªm lao Ä`á»Tng má»>i tháº¥t báº¡i: {e}")
        app.logger.error(f"muc3_1 POST error: {e}")

    finally:
        return redirect("/muc3_1")

@app.route("/muc3_2", methods=["GET", "POST"])
@login_required
@roles_required('hr', 'sa', 'gd')
def thaydoithongtinlaodong():

    if request.method == "GET":
        return render_template("3_2.html", page="3.2 Thay Ä`á»i thA'ng tin ngÆ°á»?i lao Ä`á»Tng")

    # â"?â"? POST â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?â"?
    try:
        mst = request.form.get("mst", "").strip()

        # â"?â"? áº¢nh â"?â"?
        anh = None
        file = request.files.get("anh")
        if file and file.filename:
            file_path = os.path.join(FOLDER_AVATAR, mst + ".jpg")
            if os.path.exists(file_path):
                os.remove(file_path)
            file.save(file_path)
            anh = file_path

        def _v(key):
            """Tráº£ vá»? giA¡ trá»< string, None náº¿u rá»-ng."""
            v = request.form.get(key, "").strip()
            return v if v else None

        def _num(key):
            v = _v(key)
            return v.replace(",", "") if v else None

        # â"?â"? Map: form_key â+' (col_name, is_nvarchar) â"?â"?
        FIELD_MAP = [
            # CA¡ nhA¢n
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
            # Con nhá»?
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
            # Vá»< trA-
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
            # Há»£p Ä`á»"ng
            ("kieuhopdong",             "Loai_hop_dong",              True),
            ("ngaybatdau",              "Ngay_ky_HD",                 False),
            ("ngayketthuc",             "Ngay_het_han_HD",            False),
            ("phucap",                  "Phu_cap",                    True),
            # NgA y thA¡ng / tráº¡ng thA¡i
            ("trangthai",               "Trang_thai_lam_viec",        True),
            ("ngayvao",                 "Ngay_vao",                   False),
            ("ngaynghi",                "Ngay_nghi",                  False),
            ("ngaykyhdtv",              "Ngay_ky_HDTV",               False),
            ("ngayhethanhdtv",          "Ngay_het_han_HDTV",          False),
            ("COST_ID",                 "COST_ID",                    False),
        ]

        # â"?â"? XA¢y dá»±ng SET clause â"?â"?
        set_parts = []

        # CA¡c trÆ°á»?ng thA'ng thÆ°á»?ng tá»« FIELD_MAP
        for form_key, col, is_nv in FIELD_MAP:
            val = _v(form_key)
            if val:
                prefix = "N" if is_nv else ""
                set_parts.append(f"{col} = {prefix}'{val}'")
            else:
                set_parts.append(f"{col} = NULL")

        # The_cham_cong: Acp kiá»ƒu int
        try:
            tcc = int(mst) if mst else None
        except (ValueError, TypeError):
            tcc = None
        set_parts.append(f"The_cham_cong = {tcc}" if tcc is not None else "The_cham_cong = NULL")

        # CA¡c cá»Tt sá»` (cáºn strip dáº¥u pháºcy)
        for form_key, col in [("mucluong", "Luong_co_ban"), ("tongphucap", "Tong_phu_cap")]:
            val = _num(form_key)
            set_parts.append(f"{col} = '{val}'" if val else f"{col} = NULL")

        # áº¢nh
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
        flash("Cáº-p nháº-t thA'ng tin ngÆ°á»?i lao Ä`á»Tng thA nh cA'ng !!!")

    except Exception as e:
        flash(f"Cáº-p nháº-t thA'ng tin ngÆ°á»?i lao Ä`á»Tng tháº¥t báº¡i: {e}")
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
                page="3.3 Quáº£n lA½ há»£p Ä`á»"ng lao Ä`á»Tng",
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
                flash("ThAªm há»£p Ä`á»"ng thA nh cA'ng !!!")
                # capnhatthongtinhopdong(nhamay,mst,loaihopdong,chucdanh,chuyen,luongcoban,phucap,ngaybatdau,ngayketthuc,vitrien,employeetype,positioncode,postitioncodedescription,hccategory,sectioncode,sectiondescription)
            else:
                flash("ThAªm há»£p Ä`á»"ng tháº¥t báº¡i")
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
            "Tháº» cháº¥m cA'ng": user[1],
            "Há»? tAªn": user[2],
            "Sá»` Ä`iá»╪n thoáº¡i": user[3],
            "NgA y sinh": datetime.strptime(user[4], '%Y-%m-%d').strftime("%d/%m/%Y") if user[4] else None,
            "Giá»>i tA-nh": user[5],
            "CCCD": user[6],
            "NgA y cáº¥p CCCD": datetime.strptime(user[7], '%Y-%m-%d').strftime("%d/%m/%Y") if user[7] else None ,
            "NÆ¡i cáº¥p": user[8],
            "CMT": user[9],
            "ThÆ°á»?ng trAº": user[10],
            "ThA'n xA3m": user[11],
            "PhÆ°á»?ng xA£": user[12],
            "Quáº-n huyá»╪n": user[13],
            "Tá»%nh thA nh phá»`": user[14],
            "DA¢n tá»Tc": user[15],
            "Quá»`c tá»<ch": user[16],
            "TA'n giA¡o": user[17],
            "Há»?c váº¥n": user[18],
            "NÆ¡i sinh": user[19],
            "Táº¡m trAº": user[20],
            "Sá»` BHXH": user[21],
            "MA£ sá»` thuáº¿": user[22],
            "NgA¢n hA ng": user[23],
            "Sá»` tA i khoáº£n": user[24],
            "Con nhá»?": user[25],
            "TAªn con 1": user[26],
            "NgA y sinh con 1": user[27],
            "TAªn con 2": user[28],
            "NgA y sinh con 2": user[29],
            "TAªn con 3": user[30],
            "NgA y sinh con 3": user[31],
            "TAªn con 4": user[32],
            "NgA y sinh con 4": user[33],
            "TAªn con 5": user[34],
            "NgA y sinh con 5": user[35],
            "áº¢nh chA¢n dung": user[36],
            "NgÆ°á»?i thA¢n": user[37],
            "SÄ?T liAªn há»╪": user[38],
            "Loáº¡i há»£p Ä`á»"ng": user[39],
            "NgA y kA½ HÄ?": datetime.strptime(user[40], '%Y-%m-%d').strftime("%d/%m/%Y") if user[40] else None,
            "NgA y háº¿t háº¡n": datetime.strptime(user[41], '%Y-%m-%d').strftime("%d/%m/%Y") if user[41] else None,
            "Job title VN": user[42],
            "HC category": user[43],
            "Gradecode": user[44],
            "Factory": user[45],
            "Department": user[46],
            "Chá»cc vá»¥": user[47],
            "Section code": user[48],
            "Section description": user[49],
            "Line": user[50],
            "Employee type": user[51],
            "Job title EN": user[52],
            "Position code": user[53],
            "Position description": user[54],
            "LÆ°Æ¡ng cÆ¡ báº£n": user[55],
            "Phá»¥ cáº¥p": user[56],
            "Tiá»?n phá»¥ cáº¥p": user[57],
            "NgA y vA o": datetime.strptime(user[58], '%Y-%m-%d').strftime("%d/%m/%Y"),
            "NgA y nghá»%": datetime.strptime(user[59], '%Y-%m-%d').strftime("%d/%m/%Y") if user[59] else None,
            "Tráº¡ng thA¡i": user[60],
            "NgA y vA o ná»`i thA¢m niAªn": datetime.strptime(user[61], '%Y-%m-%d').strftime("%d/%m/%Y") if user[61] else None,
            "Máº-t kháºcu": user[62],
            "NgA y kA- HÄ? Thá»- viá»╪c": datetime.strptime(user[63], '%Y-%m-%d').strftime("%d/%m/%Y") if user[63] else None,
            "NgA y háº¿t háº¡n HÄ? Thá»- viá»╪c": datetime.strptime(user[64], '%Y-%m-%d').strftime("%d/%m/%Y") if user[64] else None,
            "NgA y kA- HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 1": datetime.strptime(user[65], '%Y-%m-%d').strftime("%d/%m/%Y") if user[65] else None,
            "NgA y háº¿t háº¡n HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 1": datetime.strptime(user[66], '%Y-%m-%d').strftime("%d/%m/%Y") if user[66] else None,
            "NgA y kA- HÄ? HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 2": datetime.strptime(user[67], '%Y-%m-%d').strftime("%d/%m/%Y") if user[67] else None,
            "NgA y háº¿t háº¡n HÄ? xA¡c Ä`á»<nh thá»?i háº¡n láºn 2": datetime.strptime(user[68], '%Y-%m-%d').strftime("%d/%m/%Y") if user[68] else None,
            "NgA y kA- HÄ? khA'ng thá»?i háº¡n": datetime.strptime(user[69], '%Y-%m-%d').strftime("%d/%m/%Y") if user[69] else None,
            "Ghi chAº": user[71] if user[71] else None
            })
        df = pd.DataFrame(result)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT, f"saphethan_{thoigian}.xlsx"), index=False)
        flash("Táº£i file thA nh cA'ng !!!")
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
            # táº£i danh sA¡ch sáº_p nghá»% hÆ°u xuá»`ng excel
            # MST, Ho_ten, Chuc_danh, Gioi_tinh, Chuyen, Bo_phan, Ngay_sinh, Ngay_nghi_huu, So_thang_con_lai
            danhsach = laydanhsachsapnghihuu()
            df = pd.DataFrame(danhsach)
            df.columns = ["MST", "Ho_ten", "Chuc_danh", "Gioi_tinh", "Chuyen", "Bo_phan", "Ngay_sinh", "Ngay_nghi_huu", "So_thang_con_lai"]
            thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
            df.to_excel(os.path.join(FOLDER_XUAT, f"sapnghihuu_{thoigian}.xlsx"), index=False)
            flash("Táº£i file thA nh cA'ng !!!")
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

            if loaidieuchuyen == "Chuyá»ƒn vá»< trA-":
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
                        flash("Ä?iá»?u chuyá»ƒn thA nh cA'ng !!!")
                    else:
                        flash(f"Ä?iá»?u chuyá»ƒn tháº¥t báº¡i, lA- do: {ketqua['lido']}, query: {ketqua['query']} !!!")
                except Exception as e:
                    flash(f"Ä?iá»?u chuyá»ƒn tháº¥t báº¡i, lA- do: {e}")
                return redirect(f"/muc6_1")

            elif loaidieuchuyen == "Nghá»% viá»╪c":
                try:
                    ketqua = dichuyennghiviec(mst,
                        vitricu,
                        chuyencu,
                        gradecodecu,
                        hccategorycu,
                        ngaydieuchuyen,
                        ghichu)
                    if ketqua["ketqua"]:
                        flash("Ä?iá»?u chuyá»ƒn thA nh cA'ng !!!")
                    else:
                        flash(f"Ä?iá»?u chuyá»ƒn tháº¥t báº¡i, lA- do: {ketqua['lido']}, query: {ketqua['query']} !!!")
                except Exception as e:
                    flash(f"Ä?iá»?u chuyá»ƒn tháº¥t báº¡i, lA- do: {e}")
                return redirect(f"/muc6_1")
            elif loaidieuchuyen=="Nghá»% thai sáº£n":
                try:
                    ketqua = dichuyennghi(mst,
                                vitricu,
                                chuyencu,
                                gradecodecu,
                                hccategorycu,
                                ngaydieuchuyen,
                                'Nghá»% thai sáº£n'
                                )
                    if ketqua["ketqua"]:
                        flash("Ä?iá»?u chuyá»ƒn thA nh cA'ng !!!")
                    else:
                        flash(f"Ä?iá»?u chuyá»ƒn tháº¥t báº¡i, lA- do: {ketqua['lido']}, query: {ketqua['query']} !!!")
                except Exception as e:
                    flash(f"Ä?iá»?u chuyá»ƒn tháº¥t báº¡i, lA- do: {e}")
                return redirect(f"/muc6_1")
            elif loaidieuchuyen=="Thai sáº£n Ä`i lA m láº¡i":
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
                                    'Thai sáº£n Ä`i lA m láº¡i'
                            )
                    if ketqua["ketqua"]:
                        flash("Ä?iá»?u chuyá»ƒn thA nh cA'ng !!!")
                    else:
                        flash(f"Ä?iá»?u chuyá»ƒn tháº¥t báº¡i, lA- do: {ketqua['lido']}, query: {ketqua['query']} !!!")
                except Exception as e:
                    flash(f"Ä?iá»?u chuyá»ƒn tháº¥t báº¡i, lA- do: {e}")
                return redirect(f"/muc6_1")
            elif loaidieuchuyen=="Táº¡m hoA£n há»£p Ä`á»"ng":
                try:
                    ketqua = dichuyennghi(mst,
                                vitricu,
                                chuyencu,
                                gradecodecu,
                                hccategorycu,
                                ngaydieuchuyen,
                                'Táº¡m hoA£n há»£p Ä`á»"ng'
                                )
                    if ketqua["ketqua"]:
                        flash("Ä?iá»?u chuyá»ƒn thA nh cA'ng !!!")
                    else:
                        flash(f"Ä?iá»?u chuyá»ƒn tháº¥t báº¡i, lA- do: {ketqua['lido']}, query: {ketqua['query']} !!!")
                except Exception as e:
                    flash(f"Ä?iá»?u chuyá»ƒn tháº¥t báº¡i, lA- do: {e}")
                return redirect(f"/muc6_1")
            elif loaidieuchuyen=="Ä?i lA m láº¡i":
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
                                'Ä?i lA m láº¡i'
                            )
                    if ketqua["ketqua"]:
                        flash("Ä?iá»?u chuyá»ƒn thA nh cA'ng !!!")
                    else:
                        flash(f"Ä?iá»?u chuyá»ƒn tháº¥t báº¡i, lA- do: {ketqua['lido']}, query: {ketqua['query']} !!!")
                except Exception as e:
                    flash(f"Ä?iá»?u chuyá»ƒn tháº¥t báº¡i, lA- do: {e}")
                return redirect(f"/muc6_1")
            return redirect(f"/muc6_1")
        elif request.method == "GET":
            cacvitri= laycacvitri()
            return render_template("6_1.html",
                            cacvitri=cacvitri,
                            page="6.1 Ä?iá»?u chuyá»ƒn chá»cc vá»¥, bá»T pháº-n")
    except Exception as e:
        flash(e)
        cacvitri= laycacvitri()
        return render_template("6_1.html",
                            cacvitri=cacvitri,
                            page="6.1 Ä?iá»?u chuyá»ƒn chá»cc vá»¥, bá»T pháº-n")

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
        return render_template("6_2.html", page="6.2 Lá»<ch sá»- Ä`iá»?u chuyá»ƒn",
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
        df["NgA y thá»±c hiá»╪n"] = to_datetime(df['NgA y thá»±c hiá»╪n'])
        df["NgA y chA-nh thá»cc"] = to_datetime(df['NgA y chA-nh thá»cc'])
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Ä?iá»?u chá»%nh Ä`á»T rá»Tng cá»Tt
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
                    # Apply the date format to column L (assuming 'NgA y thá»±c hiá»╪n' is in column 'L')
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
        # Tráº£ file vá»? cho client
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
        return render_template("6_3.html", page="6.3 Lá»<ch sá»- cA'ng viá»╪c",
                               danhsach=paginated_rows,
                               pagination=pagination,
                               count=count)

    elif request.method == "POST":
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        bophan = request.form.get("bophan")
        rows = laylichsucongviec(mst,chuyen,bophan)
        data = [{
            "MA£ cA'ng ty": row[0],
            "MA£ sá»` tháº»": row[1],
            "Há»? tAªn": row[2],
            "Chuyá»?n": row[3],
            "Bá»T pháº-n": row[4],
            "Chá»cc danh": row[5],
            "Cáº¥p báº-c": row[6],
            "HC category": row[11],
            "Tráº¡ng thA¡i": row[7],
            "NgA y báº_t Ä`áºu": row[8],
            "NgA y káº¿t thAºc": row[9]
        } for row in rows]
        df = DataFrame(data)
        df["NgA y báº_t Ä`áºu"] = to_datetime(df['NgA y báº_t Ä`áºu'])
        df["NgA y káº¿t thAºc"] = to_datetime(df['NgA y káº¿t thAºc'])
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Ä?iá»?u chá»%nh Ä`á»T rá»Tng cá»Tt
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
                    # Apply the date format to column L (assuming 'NgA y thá»±c hiá»╪n' is in column 'L')
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
        # Tráº£ file vá»? cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=lichsu_congviec_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response

@app.route("/muc7_1_1", methods=["GET","POST"]) # Ä?á»i ca lA m viá»╪c
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
                                    page="7.1.1 Ä?á»i ca lA m viá»╪c",
                                    danhsach=paginated_rows,
                                    pagination=pagination,
                                    count=count,
                                    cacca=cacca)
        except:
            return render_template("7_1_1.html",
                                    page="7.1.1 Ä?á»i ca lA m viá»╪c",
                                    danhsach=[])
    elif request.method == "POST":
        mst = request.form.get("mst")
        chuyen = request.form.get("chuyen")
        phongban = request.form.get("phongban")
        rows = laydanhsachcahientai(mst,chuyen,phongban)
        data =[]
        for row in rows:
            data.append({
                "NhA  mA¡y": row[0],
                "MA£ sá»` tháº»": row[1],
                "Há»? tAªn": row[2],
                "Chuyá»?n tá»": row[3],
                "PhA²ng ban": row[4],
                "Ca": row[5],
                "Ä?á»i tá»« ngA y": row[6],
                "Ä?á»i Ä`áº¿n ngA y": row[7]
            })
        df = pd.DataFrame(data)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT, f"doica_{thoigian}.xlsx"), index=False)
        flash("Táº£i file thA nh cA'ng !!!")
        return send_file(os.path.join(FOLDER_XUAT, f"doica_{thoigian}.xlsx"), as_attachment=True)

@app.route("/muc7_1_2", methods=["GET","POST"]) # Danh sA¡ch lá»-i cháº¥m cA'ng
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
        page="Lá»-i cháº¥m cA'ng",
        danhsach=paginated_rows,
        pagination=pagination,
        count=count,
    )


@app.route("/muc7_1_3", methods=["GET","POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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
            page="Lá»-i cháº¥m cA'ng",
            danhsach=paginated_rows,
            pagination=pagination,
            count=count,
        )
    elif request.method == "POST":
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
                "NhA  mA¡y": row[0],
                "MST": row[1],
                "Há»? tAªn": row[2],
                "Chá»cc vá»¥": row[3],
                "Chuyá»?n tá»": row[4],
                "Bá»T pháº-n": row[5],
                "Loáº¡i Ä`iá»ƒm danh": row[6],
                "NgA y Ä`iá»ƒm danh": datetime.strptime(row[7], "%Y-%m-%d").strftime("%d/%m/%Y"),
                "Giá»? Ä`iá»ƒm danh": row[8],
                "LA½ do": row[9],
                "Tráº¡ng thA¡i": row[10],
                "ID":row[11],
                "Thá»?i gian táº¡o": row[12],
                "Thá»?i gian duyá»╪t": row[13]
            })

        df = pd.DataFrame(result)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT, f"diemdanhbu_{thoigian}.xlsx"), index=False) # f"diemdanhbu_{thoigian}.xlsx", index=False)

        return send_file(os.path.join(FOLDER_XUAT, f"diemdanhbu_{thoigian}.xlsx"), as_attachment=True)

@app.route("/muc7_1_3/kiemtra", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_3/tuchoi_kiemtra", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_3/pheduyet", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_3/tuchoi_pheduyet", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_4", methods=["GET","POST"]) # Danh sA¡ch xin nghá»% phAcp
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
            page="Lá»-i cháº¥m cA'ng",
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
                'MA£ cA'ng ty': row[0],
                'MA£ sá»` tháº»': row[1],
                'Há»? tAªn': row[2],
                'Chá»cc vá»¥': row[3],
                'Chuyá»?n tá»': row[4],
                'PhA²ng ban': row[5],
                'NgaI?y nghiI% phAcp': datetime.strptime(row[6], "%Y-%m-%d").strftime("%d/%m/%Y"),
                'Tá»ng sá»` phAºt': row[7],
                'LA½ do': row[8],
                'Tráº¡ng thA¡i': row[9],
                'ID': row[10],
                'Thá»?i gian táº¡o': row[11],
                'Thá»?i gian duyá»╪t': row[12]
            })
        df = pd.DataFrame(result)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT, f"xinnghiphep_{thoigian}.xlsx"), index=False)

        return send_file(os.path.join(FOLDER_XUAT, f"xinnghiphep_{thoigian}.xlsx"), as_attachment=True)

@app.route("/muc7_1_4/kiemtra", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_4/tuchoi_kiemtra", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_4/pheduyet", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_4/tuchoi_pheduyet", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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


@app.route("/muc7_1_5", methods=["GET","POST"]) # Danh sA¡ch xin nghá»% khA'ng lÆ°Æ¡ng
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
            page="Lá»-i cháº¥m cA'ng",
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
                "NhA  mA¡y": row[0],
                "MA£ sá»` tháº»": row[1],
                "Há»? tAªn": row[2],
                "Chá»cc danh": row[3],
                "Chuyá»?n tá»": row[4],
                "PhA²ng ban": row[5],
                "NgA y xin phAcp": row[6],
                "Tá»ng sá»` phAºt": row[7],
                "Loáº¡i nghá»%": row[8],
                "Tráº¡ng thA¡i": row[9],
                "ID": row[10],
                "Thá»?i gian táº¡o": row[11],
                "Thá»?i gian duyá»╪t": row[12]
            })
        df = pd.DataFrame(data)
        thoigian = datetime.now().strftime("%d%m%Y%H%M%S")
        df.to_excel(os.path.join(FOLDER_XUAT, f"xinnghikhongluong_{thoigian}.xlsx"), index=False)
        flash("Táº£i file thA nh cA'ng !!!")
        return send_file(os.path.join(FOLDER_XUAT, f"xinnghikhongluong_{thoigian}.xlsx"), as_attachment=True)

@app.route("/muc7_1_5/kiemtra", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_5/tuchoi_kiemtra", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_5/pheduyet", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_5/tuchoi_pheduyet", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_6", methods=["GET","POST"]) # Danh sA¡ch xin nghá»% khA¡c
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
            page="Lá»-i cháº¥m cA'ng",
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
            "NhA  mA¡y": row[0],
            "MA£ sá»` tháº»": row[1],
            "Há»? tAªn": row[2],
            "Chá»cc danh": row[3],
            "Chuyá»?n": row[4],
            "Bá»T pháº-n": row[5],
            "NgA y nghá»%": row[6],
            "Tá»ng sá»` phAºt": row[7],
            "Loáº¡i nghá»%": row[8],
            "Tráº¡ng thA¡i": row[9],
            "Nháº-n giáº¥y tá»?": row[10],
            "ID": row[11],
            "Thá»?i gian táº¡o": row[12],
            "Thá»?i gian duyá»╪t": row[13]
        } for row in danhsach]
        df = DataFrame(data)
        df["MA£ sá»` tháº»"] = to_numeric(df['MA£ sá»` tháº»'], errors='coerce')
        df["NgA y nghá»%"] = to_datetime(df['NgA y nghá»%'], errors='coerce')
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Ä?iá»?u chá»%nh Ä`á»T rá»Tng cá»Tt
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
        # Tráº£ file vá»? cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=xinnghikhac_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response

@app.route("/muc7_1_6/kiemtra", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_6/tuchoi_kiemtra", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_6/pheduyet", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_6/tuchoi_pheduyet", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_6/nhan_giayto", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_6/khongnhan_giayto", methods=["POST"]) # Danh sA¡ch Ä`iá»ƒm danh bA1
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

@app.route("/muc7_1_7", methods=["GET","POST"]) # Danh sA¡ch phAcp tá»"n
@login_required
def muc7_1_7():\n    # existing desktop view
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
        return render_template("7_1_7.html", page="Lá»-i cháº¥m cA'ng",
                                danhsach=paginated_rows,
                                pagination=pagination,
                                count=total)
    if request.method == "POST":
        mst = request.form.get("mst")
        danhsach = laydanhsachphepton(mst)
        result = []
        for row in danhsach:
            result.append({
                "MA£ cA'ng ty": row[0],
                "MA£ sá»` tháº»": row[1],
                "Há»? tAªn": row[2],
                "Chá»cc danh": row[3],
                "ThA¡ng": row[4],
                "NÄƒm": row[5],
                "Sá»` phAºt phAcp Ä`Æ°á»£c dA1ng": row[6],
                "Sá»` phAºt phAcp Ä`A£ chá»`t": row[7],
                "Sá»` phAºt phAcp chÆ°a dA1ng": row[8],
                "Sá»` phAºt phAcp cho dA1ng": row[9],
                "Sá»` phAºt phAcp cA²n láº¡i": row[10]
            })
        df = pd.DataFrame(result)
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Ä?iá»?u chá»%nh Ä`á»T rá»Tng cá»Tt
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
        # Tráº£ file vá»? cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=danhsachphepton_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response

@app.route("/muc7_1_8", methods=["GET","POST"]) # Ä?Äƒng kA½ lA m thAªm giá»?
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
                               page="LA m thAªm giá»?",
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
                "NhA  mA¡y": row[0],
                "MST": row[1],
                "Há»? tAªn": row[2],
                "Chá»cc vá»¥": row[3],
                "Chuyá»?n tá»": row[4],
                "PhA²ng ban": row[5],
                "NgA y Ä`Äƒng kA½": row[6],
                "Giá»? tÄƒng ca": row[7],
                "Giá»? tÄƒng ca thá»±c táº¿": row[8]
            })
        df = pd.DataFrame(data)
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Ä?iá»?u chá»%nh Ä`á»T rá»Tng cá»Tt
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
        # Tráº£ file vá»? cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=dangkytangca_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response

@app.route("/muc7_1_9", methods=["GET","POST"]) # Báº£ng lA m thAªm giá»? cháº¿ Ä`á»T
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
        return render_template("7_1_9.html", page="LA m thAªm giá»?",
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

        sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)
        sheet['A2'] = f'ThA¡ng {thang} nÄƒm {nam}'
        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # number_style = NamedStyle(name="number_style", number_format="0.00")
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        for row in range(4, 10001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_chedo_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_chedo_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc7_1_10", methods=["GET","POST"]) # Danh sA¡ch lA m thAªm giá»? ban ngA y
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
        return render_template("7_1_10.html", page="LA m thAªm giá»?",
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

        sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)
        sheet['A2'] = f'ThA¡ng {thang} nÄƒm {nam}'
        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # number_style = NamedStyle(name="number_style", number_format="0.00")
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        for row in range(4, 10001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_banngay_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_banngay_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc7_1_11", methods=["GET","POST"]) # Danh sA¡ch lA m thAªm giá»? ban Ä`Aªm
@login_required
def muc7_1_11():
    if request.method == "GET":
        thang = request.form.get("thang") if request.form.get("thang") else datetime.now().month
        nam = request.form.get("nam") if request.form.get("nam") else datetime.now().year
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
        return render_template("7_1_11.html", page="LA m thAªm giá»?",
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

        sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)
        sheet['A2'] = f'ThA¡ng {thang} nÄƒm {nam}'
        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # number_style = NamedStyle(name="number_style", number_format="0.00")
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        for row in range(4, 10001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_bandem_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_bandem_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc7_1_12", methods=["GET","POST"]) # Danh sA¡ch lA m thAªm giá»? Chá» nháº-t
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
        return render_template("7_1_12.html", page="LA m thAªm giá»?",
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

        sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)

        sheet['A2'] = f'ThA¡ng {thang} nÄƒm {nam}'

        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # number_style = NamedStyle(name="number_style", number_format="0.00")
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        for row in range(4, 10001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_chunhat_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_chunhat_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc7_1_13", methods=["GET","POST"]) # Danh sA¡ch lA m thAªm giá»? ngA y lá».
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
        return render_template("7_1_13.html", page="LA m thAªm giá»?",
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

        sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)

        sheet['A2'] = f'ThA¡ng {thang} nÄƒm {nam}'

        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # number_style = NamedStyle(name="number_style", number_format="0.00")
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        for row in range(4, 10001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_ngayle_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bang_lamthemgio_ngayle_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc7_1_14", methods=["GET","POST"]) # Báº£ng cháº¥m cA'ng chi tiáº¿t chÆ°a chá»`t
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
            page="Báº£ng cháº¥m cA'ng",
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

        sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)

        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G4:H10000
        for row in range(4, 10001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['H']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chuachot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chuachot_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc7_1_15", methods=["GET","POST"]) # Báº£ng cháº¥m cA'ng chi tiáº¿t chá»`t
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
        return render_template("7_1_15.html", page="Báº£ng cháº¥m cA'ng",
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

        sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)

        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(4, 50000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            data[7] = datetime.strptime(data[7],"%Y-%m-%d")
            sheet.append(data)

        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        number_style = NamedStyle(name="number_style", number_format="0")
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        for row in range(4, 50001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['H']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chot_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc7_1_16", methods=["GET","POST"]) # Báº£ng cháº¥m cA'ng hA nh chA-nh
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
        return render_template("7_1_16.html", page="Báº£ng cháº¥m cA'ng",
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

        sheet = workbook['Báº¢NG CHáºM CA"NG HA?NH CHA?NH']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)

        sheet['A2'] = f'ThA¡ng {thang} nÄƒm {nam}'

        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(4, 10000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-3]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            sheet.append(data)

        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # number_style = NamedStyle(name="number_style", number_format="0.00")
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        for row in range(4, 10001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_hanhchinh_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_hanhchinh_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc7_1_17", methods=["GET","POST"]) # Báº£ng cháº¥m cA'ng tá»ng há»£p
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
            #     return render_template("7_1_17_sau_072025.html", page="Báº£ng cháº¥m cA'ng",
            #                         danhsach=paginated_rows,
            #                         pagination=pagination,
            #                         count=count)
            # else:
            #     return render_template("7_1_17.html", page="Báº£ng cháº¥m cA'ng",
            #                         danhsach=paginated_rows,
            #                         pagination=pagination,
            #                         count=count)

            return render_template("7_1_17_sau_072025.html", page="Báº£ng cháº¥m cA'ng",
                                    danhsach=paginated_rows,
                                    pagination=pagination,
                                    count=count)
        except Exception as e:
            flash(f"Lá»-i táº£i trang: {e}")
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

        sheet = workbook['Báº¢NG CHáºM CA"NG Tá»"NG Há»¢P']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)

        sheet['A2'] = f'ThA¡ng {thang} nÄƒm {nam}'

        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(6, 10000 - 6 + 1)

        for row in danhsach:
            # if (nam < 2025 or (nam == 2025 and thang > 6)):
            #     data = [y for y in row]
            # else:
            #     # Chá»% láº¥y cA¡c cá»Tt cáºn thiáº¿t vA  sáº_p xáº¿p láº¡i thá»c tá»±
            #     data = [y for y in row[:-7]] + [row[-1]] + [y for y in row[-7:-4]]

            data = [y for y in row]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
            data[7] = datetime.strptime(data[7],"%Y-%m-%d") if data[7] else ""
            data[-2] = round(data[-2]) if data[-2] else 0
            sheet.append(data)


        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        number_style = NamedStyle(name="number_style", number_format="0.00")
        # if (nam > 2025 or (nam == 2025 and thang > 6)):
        #     # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        #     for row in range(6, 10001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
        #         for col in ['G', 'H']:
        #             cell = sheet[f"{col}{row}"]

        #             try:
        #                 cell.style = date_style
        #             except ValueError:
        #                 pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y
        #         for col in ['J', 'K','L', 'M','N', 'O','P', 'Q','R', 'S','T', 'U', 'X','Y', 'Z','AA','AB', 'AC','AD', 'AE', 'AF','AG', 'AH','AI', 'AJ']:
        #             cell = sheet[f"{col}{row}"]
        #             if cell.value and int(cell.value) > 0:
        #                 try:
        #                     cell.style = number_style
        #                 except ValueError:
        #                     pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y


        #     timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        #     workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"))
        #     return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"), as_attachment=True)
        # else:
        #     # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        #     for row in range(6, 10001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
        #         for col in ['G', 'H']:
        #             cell = sheet[f"{col}{row}"]

        #             try:
        #                 cell.style = date_style
        #             except ValueError:
        #                 pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y
        #         for col in ['J', 'K','L', 'M','N', 'O','P', 'Q','R', 'S','T', 'U', 'W', 'X','Y', 'Z','AA','AB', 'AC','AD', 'AE', 'AF','AG', 'AH','AI', 'AJ', 'AK']:
        #             cell = sheet[f"{col}{row}"]
        #             if cell.value and int(cell.value) > 0:
        #                 try:
        #                     cell.style = number_style
        #                 except ValueError:
        #                     pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y


        #     timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        #     workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"))
        #     return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"), as_attachment=True)
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        for row in range(6, 10001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['G', 'H']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y
            for col in ['J', 'K','L', 'M','N', 'O','P', 'Q','R', 'S','T', 'U', 'X','Y', 'Z','AA','AB', 'AC','AD', 'AE', 'AF','AG', 'AH','AI', 'AJ']:
                cell = sheet[f"{col}{row}"]
                if cell.value and int(cell.value) > 0:
                    try:
                        cell.style = number_style
                    except ValueError:
                        pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y


        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_tonghop_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc7_1_18", methods=["GET","POST"]) # Báº£ng cháº¥m cA'ng chi tiáº¿t chá»`t quA¡ khá»c
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
        return render_template("7_1_18.html", page="Báº£ng cháº¥m cA'ng",
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

        sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)

        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(4, 50000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            data[7] = datetime.strptime(data[7],"%Y-%m-%d")
            sheet.append(data)

        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        for row in range(4, 50001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['H']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chot_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc7_1_19", methods=["GET","POST"]) # Báº£ng cháº¥m cA'ng chi tiáº¿t chÆ°a chá»`t
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
        return render_template("7_1_19.html", page="Báº£ng cháº¥m cA'ng Chá» Nháº-t chi tiáº¿t chÆ°a chá»`t",
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

            sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
            image_path = HINHANH_LOGO
            # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
            img = Image(image_path)
            # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
            img.width = img.width * 0.25
            img.height = img.height * 0.25

            # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
            img.anchor = 'A1'

            # ChA"n hA¬nh áº£nh vA o sheet
            sheet.add_image(img)

            # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
            sheet.delete_rows(4, 10000 - 4 + 1)

            for row in danhsach:
                data = [y for y in row[:-1]]
                data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
                sheet.append(data)

            # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
            date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
            # Duyá»╪t qua cA¡c A' trong khu vá»±c G4:H10000
            for row in range(4, 10001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
                for col in ['G']:
                    cell = sheet[f"{col}{row}"]

                    try:
                        cell.style = date_style
                    except ValueError:
                        pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y

            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chunhat_chuachot_{timestamp}.xlsx"))
            return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_chunhat_chuachot_{timestamp}.xlsx"), as_attachment=True)
        except Exception as e:
            flash(f"Lá»-i táº£i trang: {e}")
            return render_template("7_1_19.html",
                                    danhsach=[])
@app.route("/muc7_1_20", methods=["GET","POST"]) # Báº£ng cháº¥m cA'ng chi tiáº¿t chá»`t
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
        return render_template("7_1_20.html", page="Báº£ng cháº¥m cA'ng",
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

        sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)

        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(4, 50000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d")
            sheet.append(data)

        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        number_style = NamedStyle(name="number_style", number_format="0")
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        for row in range(4, 50001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['F']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chunhat_chitiet_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chunhat_chitiet_chot_{timestamp}.xlsx"), as_attachment=True)


@app.route("/muc7_1_21", methods=["GET","POST"]) # Báº£ng cháº¥m cA'ng chi tiáº¿t chá»`t quA¡ khá»c
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
        return render_template("7_1_21.html", page="Báº£ng cháº¥m cA'ng",
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

        sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)

        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(4, 50000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            # data[7] = datetime.strptime(data[7],"%Y-%m-%d")
            sheet.append(data)

        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        for row in range(4, 50001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['H']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chunhat_chitiet_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chunhat_chitiet_chot_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc7_1_22", methods=["GET","POST"]) # Báº£ng cháº¥m cA'ng chi tiáº¿t ngA y lá». chÆ°a chá»`t
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
        return render_template("7_1_22.html", page="Báº£ng cháº¥m cA'ng ngA y lá». chi tiáº¿t chÆ°a chá»`t",
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

            sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
            image_path = HINHANH_LOGO
            # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
            img = Image(image_path)
            # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
            img.width = img.width * 0.25
            img.height = img.height * 0.25

            # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
            img.anchor = 'A1'

            # ChA"n hA¬nh áº£nh vA o sheet
            sheet.add_image(img)

            # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
            sheet.delete_rows(4, 10000 - 4 + 1)

            for row in danhsach:
                data = [y for y in row[:-1]]
                data[6] = datetime.strptime(data[6],"%Y-%m-%d") if data[6] else ""
                sheet.append(data)

            # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
            date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
            # Duyá»╪t qua cA¡c A' trong khu vá»±c G4:H10000
            for row in range(4, 10001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
                for col in ['G']:
                    cell = sheet[f"{col}{row}"]

                    try:
                        cell.style = date_style
                    except ValueError:
                        pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y

            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_ngayle_chuachot_{timestamp}.xlsx"))
            return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_ngayle_chuachot_{timestamp}.xlsx"), as_attachment=True)
        except Exception as e:
            flash(f"Lá»-i táº£i trang: {e}")
            return render_template("7_1_22.html",
                                    danhsach=[])
@app.route("/muc7_1_23", methods=["GET","POST"]) # Báº£ng cháº¥m cA'ng chi tiáº¿t ngA y lá». chá»`t
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
        return render_template("7_1_23.html", page="Báº£ng cháº¥m cA'ng chi tiáº¿t ngA y lá». chá»`t",
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

        sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)

        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(4, 50000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            data[6] = datetime.strptime(data[6],"%Y-%m-%d")
            sheet.append(data)

        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        number_style = NamedStyle(name="number_style", number_format="0")
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        for row in range(4, 50001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['F']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_ngayle_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_chitiet_ngayle_chot_{timestamp}.xlsx"), as_attachment=True)


@app.route("/muc7_1_24", methods=["GET","POST"]) # Báº£ng cháº¥m cA'ng chi tiáº¿t ngA y lá». quA¡ khá»c
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
        return render_template("7_1_24.html", page="Báº£ng cháº¥m cA'ng chi tiáº¿t ngA y lá». quA¡ khá»c",
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

        sheet = workbook['Sheet1']  # Thay 'Sheet1' báº±ng tAªn sheet cá»a báº¡n
        image_path = HINHANH_LOGO
        # Táº¡o Ä`á»`i tÆ°á»£ng hA¬nh áº£nh
        img = Image(image_path)
        # Ä?iá»?u chá»%nh kA-ch thÆ°á»>c hA¬nh áº£nh xuá»`ng 70% so vá»>i kA-ch thÆ°á»>c gá»`c
        img.width = img.width * 0.25
        img.height = img.height * 0.25

        # Di chuyá»ƒn áº£nh: anchor vA o A' A2 vA  Ä`iá»?u chá»%nh tá»?a Ä`á»T di chuyá»ƒn
        img.anchor = 'A1'

        # ChA"n hA¬nh áº£nh vA o sheet
        sheet.add_image(img)

        # XA3a hA ng tá»« hA ng 7 Ä`áº¿n hA ng 10000
        sheet.delete_rows(4, 50000 - 4 + 1)

        for row in danhsach:
            data = [y for y in row[:-1]]
            # data[7] = datetime.strptime(data[7],"%Y-%m-%d")
            sheet.append(data)

        # Táº¡o kiá»ƒu Ä`á»<nh dáº¡ng ngA y
        date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")
        # Duyá»╪t qua cA¡c A' trong khu vá»±c G7:H10000
        for row in range(4, 50001):  # Báº_t Ä`áºu tá»« dA²ng 7 Ä`áº¿n dA²ng 10000
            for col in ['H']:
                cell = sheet[f"{col}{row}"]

                try:
                    cell.style = date_style
                except ValueError:
                    pass  # Náº¿u giA¡ trá»< khA'ng pháº£i lA  ngA y, bá»? qua A' nA y

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        workbook.save(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_ngayle_chitiet_chot_{timestamp}.xlsx"))
        return send_file(os.path.join(os.path.dirname(__file__),f"nhapxuat/xuat/bangchamcong_ngayle_chitiet_chot_{timestamp}.xlsx"), as_attachment=True)

@app.route("/muc8_1", methods=["GET","POST"])
@login_required
def ykienkhieunai():

    return render_template("8_1.html", page="8.1 Danh sA¡ch A½ kiáº¿n khiáº¿u náº¡i")

@app.route("/muc8_2", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def capnhatykienkhieunai():

    return render_template("8_2.html", page="8.2 Cáº-p nháº-t A½ kiáº¿n khiáº¿u náº¡i")

@app.route("/muc9_1", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def xulykiluat():

    if request.method == "GET":
        danhsach = laydanhsachkyluat()
        return render_template("9_1.html", page="9.1 Xá»- lA½ ká»% luáº-t",danhsach=danhsach)
    else:
        try:
            mst = request.form.get("mst")
            if not mst:
                flash("ChÆ°a cA3 thA'ng tin ngÆ°á»?i vi pháº¡m")
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
                flash("ThAªm biAªn báº£n ká»· luáº-t thA nh cA'ng !!!")
            else:
                flash("ThAªm biAªn báº£n ká»· luáº-t tháº¥t báº¡i !!!")
        except Exception as e:
            flash(f"ThAªm biAªn báº£n ká»· luáº-t tháº¥t báº¡i {e}!!!")
        return redirect("/muc9_1")

@app.route("/muc10_1", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def phongvannghiviec():

    return render_template("10_1.html", page="10.1 Tá»ng há»£p phá»?ng váº¥n nghá»% viá»╪c")

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
                            page="10.2 Tá»ng há»£p Ä`Æ¡n nghá»% viá»╪c",
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
            flash("ThAªm Ä`Æ¡n xin nghá»% thA nh cA'ng !!!")
        else:
            flash("ThAªm Ä`Æ¡n xin nghá»% tháº¥t báº¡i !!!")
        return redirect(f"/muc10_2?mst={mst}")

@app.route("/muc10_3", methods=["GET","POST"])
@login_required
@roles_required('hr','sa','gd')
def inchamduthopdong():

    if request.method == "GET":
        return render_template("10_3.html", page="10.3 In cháº¥m dá»ct há»£p Ä`á»"ng")
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
            flash(f"Lá»-i táº£i trang: {e}")
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
\n\n@app.route( /muc7_1_7mobile, methods=[GET]) # Phép tồn mobile\n@login_required\ndef muc7_1_7mobile():\n    nhamay = current_user.macongty\n    mst = request.args.get(mst) or current_user.masothe\n    today = datetime.today()\n    thang = request.args.get(thang, [int]) or today.month\n    nam = request.args.get(nam, [int]) or today.year\n    danhsach = laydanhsachphepton(mst)\n    return render_template(mobile/7_1_7mobile.html, mst=mst, thang=thang, nam=nam, danhsach=danhsach, page=Phép tồn)\n
\n\n@app.route( /muc7_1_7mobile, methods=[GET]) # Phép tồn mobile\n@login_required\ndef muc7_1_7mobile():\n    nhamay = current_user.macongty\n    mst = request.args.get(mst) or current_user.masothe\n    today = datetime.today()\n    thang = request.args.get(thang, [int]) or today.month\n    nam = request.args.get(nam, [int]) or today.year\n    danhsach = laydanhsachphepton(mst)\n    return render_template(mobile/7_1_7mobile.html, mst=mst, thang=thang, nam=nam, danhsach=danhsach, page=Phép tồn)\n
