import subprocess
import time

while True:
    try:
        # Chạy ứng dụng Flask cùng với waitress
        subprocess.run(["python", "-m", "waitress", "--port=81", "routes:app"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Flask gặp lỗi: {e}")
        print("Đang khởi động lại Flask...")
        time.sleep(1)  # Đợi một khoảng thời gian trước khi khởi động lại
    except Exception as e:
        print(f"Lỗi không xác định: {e}")
        print("Đang khởi động lại Flask...")
        time.sleep(1)