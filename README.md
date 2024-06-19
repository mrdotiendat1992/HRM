# HRM
Phần mềm quản lý nhân sự

```
Phần mềm có 2 process:
1. Tự động xử lý các dữ liệu trên sql server, cào dữ liệu từ sharepoint lists về
2. Webapp tương tác với sqlserver trên nền tảng flask python
```

## Process 1: Script auto
```
Có 4 process:
1. Lúc 00:00 hàng ngày, cào dữ liệu trong List "Dang_ky_thong_tin" để đẩy về bảng "Dang_ky_thong_tin" trong SQLSERVER có địa chỉ "172.16.60.100"
2. Lúc 07:45 hàng ngày, xoá dữ liệu ngày hiện tại trên máy SQLSERVER có địa chỉ "172.16.60.100", lấy dữ liệu 2 ngày gần nhất trên SQLSERVER máy MITA có địa chỉ "10.0.0.252", và đẩy lên SQLSERVER có địa chỉ "172.16.60.100"
3. Lúc 09:00 hàng ngày, cào dữ liệu trong List "Form đăng ký điểm danh bù" để đẩy về bảng "Diem_danh_bu" trong SQLSERVER có địa chỉ "172.16.60.100"
4. Lúc 09:00 hàng ngày, cào dữ liệu trong List "Form xin nghỉ phép" để đẩy về bảng "Xin_nghi_phep" trong SQLSERVER có địa chỉ "172.16.60.100" 
```

## Process 1: Webapp Flask
```
Địa chỉ: http://10.0.0.252:81

```