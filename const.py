from lib import *

used_db = r"Driver={SQL Server};Server=172.16.60.100;Database=HR;UID=huynguyen;PWD=Namthuan@123;"
mccdb = r"Driver={SQL Server}; Server=10.0.0.252\SQLEXPRESS; Database=MITACOSQL; UID=sa;PWD=Namthuan1;"

FOLDER_NHAP = os.path.join(os.path.dirname(__file__), r'nhapxuat\nhap')
FOLDER_XUAT = os.path.join(os.path.dirname(__file__), r'nhapxuat\xuat')

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

CA_THEO_CHUYEN = {'0IED00': 'A1-01', '0MGT00': 'A1-02', '0PPC00': 'A1-02', '0SUS00': 'A1-01', '11C00': 'A1-03', '11C01': 'A1-03', '11F00': 'A1-03', '11F01': 'A1-03', '11S00': 'A1-02', '11S01': 'A1-02', '11S03': 'A1-02', '11S05': 'A1-02', '11S07': 'A1-02', '11S09': 'A1-02', '11S11': 'A1-02', '11S13': 'A1-02', '12C02': 'A1-03', '12F02': 'A1-03', '12S00': 'A1-02', '12S01': 'A1-02', '12S03': 'A1-02', '12S05': 'A1-02', '12S07': 'A1-02', '12S09': 'A1-02', '12S11': 'A1-02', '12S13': 'A1-02', '1ACC00': 'A1-03', '1ACT00': 'A1-01', '1ADM00': 'A1-03', '1CMD00': 'A1-03', '1FAB00': 'A1-03', '1FQC00': 'A1-03', '1FSG00': 'A1-03', '1HRD00': 'A1-01', '1IED00': 'A1-01', '1ISD00': 'A1-01', '1LOG00': 'A1-01', '1MEC00': 'A1-02', '1MMD00': 'A1-03', '1MTN00': 'A1-02', '1NDC00': 'A1-03', '1PAY00': 'A1-01', '1PDN00': 'A1-02', '1PPC00': 'A1-01', '1PUR00': 
'A1-01', '1QAD00': 'A1-02', '1SPL00': 'A1-02', '1SUS00': 'A1-01', '1TEC00': 'A1-02', '1TNC00': 'A1-02', '1WHS00': 'A1-03', '11QC101': 'A1-02', '12QC101': 'A1-02', '11QC201': 'A1-03', '12QC201': 'A1-03', '11I01': 'A1-03', '11I02': 'A1-03', '11I03': 'A1-03', '12I01': 'A1-03', '12I02': 'A1-03', '12I03': 'A1-03'}