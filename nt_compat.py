# -*- coding: utf-8 -*-
"""
============================================================================
 nt_compat.py - Lop tuong thich khi chuyen tu Windows sang Linux/Docker
============================================================================

 VAN DE
 ------
 Tren Windows, cac cho dung `DRIVER={SQL Server}` thuc chat chay bang
 SQLSRV32.DLL - driver doi 2000, KHONG biet cac kieu du lieu SQL Server
 2008+ la DATE / TIME / DATETIME2, nen no tra ve CHUOI:

     row[7]  ->  '2024-01-05'        (str)
     row[3]  ->  '08:30:00'          (str)

 Tren Linux khong co driver do; Dockerfile tro alias {SQL Server} sang
 ODBC Driver 17. Driver 17 hieu day du cac kieu nay nen tra ve OBJECT:

     row[7]  ->  datetime.date(2024, 1, 5)
     row[3]  ->  datetime.time(8, 30)

 Code hien tai duoc viet theo hanh vi cu nen sinh 2 loi:
     TypeError: strptime() argument 1 must be str, not datetime.date
     TypeError: Object of type time is not JSON serializable

 CACH XU LY
 ----------
 File nay KHONG sua logic nghiep vu. No chi khoi phuc hanh vi cu o dung
 hai cho:

   1. install_pyodbc_legacy_dates()
        Dang ky output converter cua pyodbc: cot DATE va TIME duoc tra ve
        thanh chuoi 'YYYY-MM-DD' / 'HH:MM:SS' y het driver cu.
        Cot DATETIME/DATETIME2 GIU NGUYEN kieu datetime - vi driver cu
        cung tra ve datetime cho kieu do, khong doi.

   2. install_json_provider(app)
        Cho Flask jsonify() tu doi date/time/datetime/Decimal thanh chuoi
        thay vi nem TypeError.

 CACH DUNG - them 3 dong vao dau app.py, NGAY SAU dong `app = Flask(...)`:

     from nt_compat import install_pyodbc_legacy_dates, install_json_provider
     install_pyodbc_legacy_dates()
     install_json_provider(app)

 KIEM TRA:  python nt_compat.py        (chay self-test, khong can DB)

 LUU Y: day la cau noi de chay duoc ngay, khong phai giai phap vinh vien.
 Ve lau dai nen sua code de lam viec truc tiep voi date/time object.
============================================================================
"""

import datetime
import struct
from decimal import Decimal

# --- Ma kieu SQL (ODBC) ----------------------------------------------------
SQL_TYPE_DATE = 91
SQL_TYPE_TIME = 92
SQL_SS_TIME2 = -154          # kieu TIME(n) that su cua SQL Server


# ---------------------------------------------------------------------------
# 1. Tra DATE / TIME ve dang chuoi nhu driver cu
# ---------------------------------------------------------------------------
def _date_to_str(raw):
    """SQL_DATE_STRUCT: SQLSMALLINT year, SQLUSMALLINT month, SQLUSMALLINT day."""
    try:
        if len(raw) >= 6:
            y, m, d = struct.unpack("<hHH", raw[:6])
            if 1 <= m <= 12 and 1 <= d <= 31 and 1 <= y <= 9999:
                return "%04d-%02d-%02d" % (y, m, d)
    except Exception:
        pass
    return _fallback_text(raw)


def _time_to_str(raw):
    """SQL_TIME_STRUCT (6 byte) hoac SQL_SS_TIME2_STRUCT (12 byte)."""
    try:
        if len(raw) >= 6:
            h, mi, s = struct.unpack("<HHH", raw[:6])
            if h <= 23 and mi <= 59 and s <= 61:
                return "%02d:%02d:%02d" % (h, mi, s)
    except Exception:
        pass
    return _fallback_text(raw)


def _fallback_text(raw):
    """Neu layout khong khop du doan thi doc nhu chuoi, khong lam hong du lieu."""
    for enc in ("utf-16-le", "utf-8", "latin-1"):
        try:
            txt = raw.decode(enc).rstrip("\x00").strip()
            if txt:
                return txt
        except Exception:
            continue
    return raw


def install_pyodbc_legacy_dates():
    """Bat che do 'DATE/TIME tra ve chuoi' cho MOI connection pyodbc moi tao."""
    import pyodbc

    if getattr(pyodbc, "_nt_compat_installed", False):
        return
    _real_connect = pyodbc.connect

    def _connect(*args, **kwargs):
        conn = _real_connect(*args, **kwargs)
        conn.add_output_converter(SQL_TYPE_DATE, _date_to_str)
        conn.add_output_converter(SQL_TYPE_TIME, _time_to_str)
        conn.add_output_converter(SQL_SS_TIME2, _time_to_str)
        return conn

    pyodbc.connect = _connect
    pyodbc._nt_compat_installed = True


# ---------------------------------------------------------------------------
# 2. Flask jsonify() khong con nem TypeError voi date/time/Decimal
# ---------------------------------------------------------------------------
def install_json_provider(app):
    from flask.json.provider import DefaultJSONProvider

    class NTJSONProvider(DefaultJSONProvider):
        @staticmethod
        def default(o):
            if isinstance(o, datetime.datetime):
                return o.strftime("%Y-%m-%d %H:%M:%S")
            if isinstance(o, datetime.date):
                return o.strftime("%Y-%m-%d")
            if isinstance(o, datetime.time):
                return o.strftime("%H:%M:%S")
            if isinstance(o, datetime.timedelta):
                return str(o)
            if isinstance(o, Decimal):
                return float(o)
            if isinstance(o, (bytes, bytearray)):
                return o.decode("utf-8", "replace")
            return DefaultJSONProvider.default(o)

    app.json = NTJSONProvider(app)
    return app


# ---------------------------------------------------------------------------
# Self-test - chay duoc ma khong can SQL Server
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    ok = True

    def check(ten, thuc_te, mong_doi):
        global ok
        dat = thuc_te == mong_doi
        ok = ok and dat
        print(("  OK  " if dat else "  SAI ") + "%-34s %r" % (ten, thuc_te))

    print("1. Doc byte tho cua ODBC -> chuoi")
    check("DATE 2024-01-05", _date_to_str(struct.pack("<hHH", 2024, 1, 5)), "2024-01-05")
    check("DATE 1999-12-31", _date_to_str(struct.pack("<hHH", 1999, 12, 31)), "1999-12-31")
    check("TIME 08:30:00", _time_to_str(struct.pack("<HHH", 8, 30, 0)), "08:30:00")
    check("TIME2 23:59:59", _time_to_str(struct.pack("<HHHI", 23, 59, 59, 0)), "23:59:59")
    check("byte la -> doc thanh chuoi",
          _date_to_str("2024-01-05".encode("utf-16-le")), "2024-01-05")

    print("\n2. strptime tren chuoi tra ve (code hien tai dang lam vay)")
    check("strptime DATE",
          datetime.datetime.strptime(_date_to_str(struct.pack("<hHH", 2024, 1, 5)),
                                     "%Y-%m-%d").strftime("%d/%m/%Y"), "05/01/2024")

    print("\n3. Flask jsonify")
    try:
        from flask import Flask, jsonify
        app = install_json_provider(Flask(__name__))
        with app.test_request_context():
            body = jsonify({"data": [{
                "gio": datetime.time(8, 30),
                "ngay": datetime.date(2024, 1, 5),
                "luc": datetime.datetime(2024, 1, 5, 8, 30),
                "tien": Decimal("1234.50"),
            }]}).get_data(as_text=True)
        check("jsonify time",     '"gio":"08:30:00"' in body.replace(" ", ""), True)
        check("jsonify date",     '"ngay":"2024-01-05"' in body.replace(" ", ""), True)
        check("jsonify datetime", '"luc":"2024-01-0508:30:00"' in body.replace(" ", ""), True)
        check("jsonify Decimal",  '"tien":1234.5' in body.replace(" ", ""), True)
    except ImportError:
        print("  --  bo qua (may nay chua cai flask)")

    print("\n" + ("TAT CA DEU DAT" if ok else "CO TRUONG HOP SAI - xem lai"))
