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
    "/muc6.2"  :  Lịch sử công tác
    "/muc7.1.1":  Đổi ca làm việc
    "/muc7.1.2":  Danh sách lỗi chấm công
    "/muc7.1.3":  Danh sách điểm danh bù bằng QR
    "/muc7.1.4":  Danh sách xin nghỉ phép bằng QR
    "/muc7.1.5":  Danh sách xin nghỉ khác
    "/muc7.1.6":  Đăng ký tăng ca
    "/muc7.1.7":  Bảng chấm công 5 ngày gần nhất
    "/muc7.1.8":  Bảng chấm công chốt
    "/muc7.1.9":  Bảng phép tồn
    "/muc7.1.8":  Cập nhật dữ liệu chấm công
    "/muc8.1"  :  Ý kiến khiếu nại
    "/8.2"     :  Cập nhật khiếu nại
    "/muc9.1"  :  Xử lý kỷ luật
    "/muc10.3" :  In chấm dứt hợp đồng
```

## Phân quyền
```
    "/"         :  ALL
    "/muc2.1"   :  HRD00, TNC00
    "/muc2.2.1" :  TBP
    "/muc2.2.2" :  HRD
    "/muc3.1"   :  HRD
    "/muc3.2"   :  HRD
    "/muc3.3"   :  HRD
    "/muc3.4"   :  ALL
    "/muc6.1"   :  HRD
    "/muc6.2"   :  ALL
    "/muc7.1.1" :  HRD
    "/muc7.1.2" :  ALL
    "/muc7.1.3" :  ALL
    "/muc7.1.4" :  ALL
    "/muc7.1.5" :  ALL
    "/muc7.1.6" :  HRD, THUKI
    "/muc7.1.7" :  ALL
    "/muc7.1.8" :  ALL
    "/muc7.1.9" :  ALL
    "/muc7.1.10":  HRD (2MST)
    "/muc8.1"   :  ALL 
    "/muc8.2"   :  HRD 
    "/muc9.1"   :  HRD
    "/muc10.3"  :  HRD 
```