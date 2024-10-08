# HRM
Phần mềm quản lý nhân sự

```
Webapp quản lý bằng sqlserver trên nền tảng flask python
```

## Webapp Flask
```
Địa chỉ HP: http://10.0.0.252:81
Địa chỉ NA: http://172.16.60.98:81

```

## Main Endpoint
```
    "/"        :  Trang chủ - Danh sách nhân viên 
    "/muc2.1"  :  Danh sách ứng viên đã quét mã QR đăng ký thông tin
    "/muc2.2.1":  Thêm yêu cầu tuyển dụng, danh sách yêu cầu tuyển dụng theo phòng ban của người đăng nhập
    "/muc2.2.2":  Phê duyệt yêu cầu tuyển dụng (chỉ người được phê duyệt mới được vào)
    "/muc3.1"  :  Thêm lao động mới lên hệ thống
    "/muc3.2"  :  Thay đổi thông tin lao động
    "/muc3.3"  :  Quản lý hợp đồng lao động
    "/muc3.4"  :  Danh sách nhân viên sắp hết hạn Hợp đồng
    "/muc6.1"  :  Điều chuyển nhân sự
    "/muc6.2"  :  Lịch sử điều chuyển
    "/muc6.3"  :  Lịch sử công việc
    "/muc7.1.1":  Đổi ca làm việc
    "/muc7.1.2":  Danh sách lỗi chấm công
    "/muc7.1.3":  Danh sách điểm danh bù bằng QR
    "/muc7.1.4":  Danh sách xin nghỉ phép bằng QR
    "/muc7.1.5":  Danh sách xin nghỉ không lương
    "/muc7.1.6":  Danh sách xin nghỉ khác
    "/muc7.1.7":  Đăng ký tăng ca
    "/muc7.1.8":  Bảng chấm công 5 ngày gần nhất
    "/muc7.1.9":  Bảng chấm công chốt
    "/muc7.1.10": Bảng phép tồn
    "/muc8.1"  :  Ý kiến khiếu nại
    "/muc8.2"  :  Cập nhật khiếu nại
    "/muc9.1"  :  Xử lý kỷ luật
    "/muc10.1" :  Phỏng vấn nghỉ việc
    "/muc10.2" :  Nhận đơn nghỉ việc
    "/muc10.3" :  In chấm dứt hợp đồng
    "/muc10.4" :  Bàn giao
    "/muc12"   :  Không kiểm xưởng
```

## Phân quyền
```
    "/"         :   ALL
    "/muc2.1"   :   HRD, TNC, TD, GD, SA
    "/muc2.2.1" :   TBP, GD, SA
    "/muc2.2.2" :   HRD,TBP, GD, SA
    "/muc3.1"   :   HRD, GD, SA
    "/muc3.2"   :   HRD, GD, SA
    "/muc3.3"   :   HRD, GD, SA
    "/muc3.4"   :   HRD, GD, SA
    "/muc5.1.1" :   TBP, GD, SA
    "/muc5.1.2" :   GD, SA
    "/muc5.1.3.1" : TBP, GD, SA
    "/muc5.1.3.2" : TBP, GD, SA
    "/muc6.1"   :   HRD, GD, SA
    "/muc6.2"   :   HRD, GD, SA
    "/muc6.3"   :   HRD, GD, SA
    "/muc7.1.1" :   HRD, GD, SA
    "/muc7.1.2" :   ALL
    "/muc7.1.3" :   ALL
    "/muc7.1.4" :   ALL
    "/muc7.1.5" :   ALL
    "/muc7.1.6" :   ALL
    "/muc7.1.7" :   HRD, GD, SA, TK
    "/muc7.1.8" :   ALL
    "/muc7.1.9" :   ALL
    "/muc7.1.10":   ALL    
    "/muc8.1"   :   HRD, GD, SA 
    "/muc8.2"   :   HRD, GD, SA 
    "/muc9.1"   :   HRD, GD, SA
    "/muc10.2"  :   HRD, GD, SA 
    "/muc10.3"  :   HRD, GD, SA 
```

## MST ảo
### NT1:
MST |  Tên                      | Phòng ban 
2   | Lưu Thị Hằng              | IED
3   | Nguyễn Thị Kim Anh        | IED
5   | Nguyễn Hoàng Xuân Phúc    | MGT
6   | Trần Lê Đại Dương         | MGT
9   | Vũ Thị Hương Lan          | QAD
10  | Tiến                      | ACT

### NT1:
MST |  Tên                      | Phòng ban 
19  | Nguyễn Thị Thắng          | PPC
23  | Vũ Thị Hương Lan          | QAD
25  | Nguyễn Hoàng Xuân Phúc    | MGT
28  | Trần Lê Đại Dương         | MGT