from lib import *

used_db = r"Driver={SQL Server};Server=172.16.60.100;Database=HR;UID=huynguyen;PWD=Namthuan@123;"

# used_db = r"Driver={SQL Server};Server=Admin;Database=HR;Trusted_Connection=yes;"

# mcc_db = r"Driver={SQL Server};Server=10.0.0.252\SQLEXPRESS;Database=MITACOSQL;UID=sa;PWD=Namthuan1"

FOLDER_NHAP = os.path.join(os.path.dirname(__file__), r'nhapxuat/nhap')
FOLDER_XUAT = os.path.join(os.path.dirname(__file__), r'nhapxuat/xuat')

FOLDER_JD = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/jd')

FILE_MAU_HDTV_NT1_O2_TROLEN = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/hopdong/nt1/o2_trolen/hdtv.xlsx')
FILE_MAU_HDTV_NT1_DUOI_O2 = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/hopdong/nt1/duoi_o2/hdtv.xlsx')
FILE_MAU_HDCTH_NT1_O2_TROLEN = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/hopdong/nt1/o2_trolen/hdcth.xlsx')
FILE_MAU_HDCTH_NT1_DUOI_O2 = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/hopdong/nt1/duoi_o2/hdcth.xlsx')
FILE_MAU_HDVTH_NT1_O2_TROLEN = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/hopdong/nt1/o2_trolen/hdvth.xlsx')
FILE_MAU_HDVTH_NT1_DUOI_O2 = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/hopdong/nt1/duoi_o2/hdvth.xlsx')
FILE_MAU_HDNH_NT1_O2_TROLEN = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/hopdong/nt1/o2_trolen/hdnh.xlsx')
FILE_MAU_HDNH_NT1_DUOI_O2 = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/hopdong/nt1/duoi_o2/hdnh.xlsx')
FILE_MAU_HDTV_NT2 = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/hopdong/nt2/hdtv.xlsx')
FILE_MAU_HDCTH_NT2 = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/hopdong/nt2/hdcth.xlsx')
FILE_MAU_HDVTH_NT2 = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/hopdong/nt2/hdvth.xlsx')
FILE_MAU_HDNH_NT2 = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/hopdong/nt2/hdnh.xlsx')

FILE_MAU_CDHD_NT1 = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/chamduthopdong/nt1/cdhd.xlsx')
FILE_MAU_CDHD_NT2 = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/chamduthopdong/nt2/cdhd.xlsx')

FILE_MAU_DANGKY_TANGCA_NHOM = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/tangca/tangcanhom.xlsx') 

FILE_MAU_DANGKY_DOICA_NHOM = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/doica/doicanhom.xlsx') 

FILE_MAU_DANGKY_XINNGHIKHAC = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/xinnghikhac/xinnghikhac.xlsx') 

FILE_MAU_DANGKY_KPI = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/kpi/Kpi report sample.xlsx') 

FILE_MAU_CAPNHAT_STK = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/capnhatstk/capnhatstk.xlsx')

FILE_MAU_THEM_HOPDONG = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/nhaphophong/taohopdong.xlsx')

FILE_MAU_DIEU_CHUYEN = os.path.join(os.path.dirname(__file__), r'static/uploads/mau/dieuchuyen/dieuchuyen.xlsx')