# HRM
Phần mềm quản lý nhân sự

# INPUT
```
I - Dữ liệu chấm công: Bảng Check_in_out
    1 - Dữ liệu chấm công ở NT1: Database Mitacosql, bảng checkinout
    2 - Dữ liệu chấm công ở NT2: Database CheckInOut

II - Lịch sử làm việc
    1 - Dữ liệu Lịch sử làm việc: Bảng LICH_SU_CONG_VIEC
    2 - Dữ liệu Ca làm việc: Bảng Dang_ky_ca_lam_viec
    3 - Dữ liệu Danh sach ca: Bảng Ca_lam_viec
    4 - Headcount Name: Bảng HC_name

III - Các form xử lý lỗi chấm công
    1 - Form đăng ký tăng ca: Bảng Dang_ky_tang_ca
    2 - Form Điểm danh bù : Bảng Diem_danh_bu
    3 - Form Xin nghỉ phép: Bảng Xin_nghi_phep
    4 - Form Xin nghỉ khác: Bảng Xin_nghi_khac
    3 - Form Xin nghỉ không: Bảng Xin_nghi_khong_luong
```

# SQL SERVER PROCEDURE
```
    - Cập nhật dữ liệu chấm công: Dong_bo_CheckInOut
    - Cập nhật dữ liệu giờ ra vào hôm trước vào form tăng ca: CAP_NHAT_GIO_VAO_RA_SANG_DK_TANG_CA_TU_DONG
    - Cập nhật bảng công chi tiết trong đợt chốt công: Cap_nhat_cham_cong_chi_tiet_chua_chot
    - Cập nhật bảng công chi tiết trong cả tháng: Cap_nhat_cham_cong_chi_tiet_ca_thang
    - Cập nhật bảng công hành chính thực tế đợt chốt công: Cap_nhat_bang_cong_hanh_chinh_thuc_te
    - Cập nhật bảng công hành chính kiểm xưởng đợt chốt công: Cap_nhat_bang_cong_hanh_chinh_kiem_xuong
    - Cập nhật bảng công tổng hợp thực tế đợt chốt công: Cap_nhat_bang_cong_tong_hop_thuc_te
    - Cập nhật bảng công tổng hợp kiểm xưởng đợt chốt công: Cap_nhat_bang_cong_tong_hop_kiem_xuong
    - Đưa các bảng chấm công qua đợt chấm công về bảng công đã chốt hoặc bảng công quá khứ: Cham_cong_tu_dong

```