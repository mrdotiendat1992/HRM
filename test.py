from datetime import datetime

tuoi_nghi_huu_nam = 61
thang_nghi_huu_nam = 0

Ngay_sinh = '09/11/1992'

ngay_sinh = datetime.strptime(Ngay_sinh, '%d/%m/%Y')

# Tinh toán đến ngày hiện tại ngjười này bao nhiêu tuổi, bao nhiêu tháng

tuoi_hien_tai = datetime.now().year - ngay_sinh.year
thang_hien_tai = datetime.now().month - ngay_sinh.month

#  Nếu tháng hiện tại nhỏ hơn tháng sinh thì trừ 1 tuổi
if thang_hien_tai < 0:
    tuoi_hien_tai -= 1
    thang_hien_tai += 12

# Tính số tháng hiện tại
so_thang_hien_tai = tuoi_hien_tai * 12 + thang_hien_tai

# Tính số tháng nghỉ hưu
so_thang_nghi_huu = tuoi_nghi_huu_nam * 12 + thang_nghi_huu_nam

# Tính số tháng còn lại
so_thang_con_lai = so_thang_nghi_huu - so_thang_hien_tai
so_nam_con_lai = so_thang_con_lai // 12
so_thang_con_lai = so_thang_con_lai % 12

print("Tuổi hiện tại: {} tuổi {} tháng".format(tuoi_hien_tai, thang_hien_tai))

# NẾu nhỏ hơn 1 năm thì cảnh báo là sắp đến tuối nghỉ hưu
if so_nam_con_lai < 1:
    print("Sắp đến tuối nghỉ hưu")

else:
    print("Còn {} năm {} tháng để nghỉ hưu".format(so_nam_con_lai, so_thang_con_lai))