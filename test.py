from other_routes import *

if __name__ == "__main__":
    while True:
        try:
            print("TEST")
            serve(app, host="0.0.0.0", port=96, _quiet=True, threads=8)
        except subprocess.CalledProcessError as e:
            print(f"Flask gặp lỗi: {e}")
            print("Đang khởi động flask...")
            time.sleep(1)  # Đợi một khoảng thời gian trước khi khởi động lại
        except Exception as e:
            print(f"Lỗi không xác định: {e}")
            print("Đang khởi động lại flask...")
            time.sleep(1)  # Đợi một khoảng thời gian trước khi khởi động lại