# Giao diện gười dùng:

```
    - Chọn nhà máy muốn đăng nhập
    - Tài khoản chính là mã số thẻ của cán bộ công nhân viên thuộc nhà máy
    - Điền mật khẩu đã đăng ký ( Mặc định là "1")
    - Nếu điền đúng mật khẩu sẽ chuyển sang giao diện trang chủ
```



![Giao diện đăng nhập](<imgs/Giao diện đăng nhập.png>)

![Giao diện trang chủ](<imgs/Giao diện trang chủ.png>)

# Kĩ thuật:

```
    - Sử dụng thư viện flask_login để quản lý session đằng nhập
    - Query từ các cột "masothe", "macongty", "matkhau" ở bảng "Nhanvien" để kiểm tra người dùng tồn tại hay 
    - Sử dụng function "login"
```
![login](imgs/Function/login.png)

```
    - Nếu đăng nhập thành công, sẽ lấy các thông tin người dùng từ DB để hiển thị 
    - Hiển thị thông tin cá nhân (Họ tên, mã số thẻ, phòng ban trong biến current_user)
    - Lấy các form điểm danh bù, xin nghỉ phép, xin nghỉ không lương, xin nghỉ khác có mã số thẻ và nhà máy trùng với người đăng nhập để hiển thị ở phần Lỗi chấm công cá nhân (trong biến g.notice)
    - Lấy các form điểm danh bù, xin nghỉ phép, xin nghỉ không lương, xin nghỉ khác, các yêu cầu tuyển dụng chưa được xử lý và được quản lý bởi nhân viên có mã số thẻ và nhà máy trùng với người đăng nhập để hiển thị ở phần Quản lý các lỗi chấm công của bộ phận được phân công (trong biến g.notice)
    - Lấy trạng thái của biến f12 để bật/tắt function 12 trên giao diện (ở file f12.ini, nếu on=1 thì mở, on=0 là tắt)
    - Ở các trang, sẽ chỉ lấy thông tin của nhà máy trùng với nhà máy mà người dùng đã đăng nhập vào (trong biến current_user)
```