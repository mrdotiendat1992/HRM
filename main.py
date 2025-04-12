from other_routes import *

if __name__ == "__main__":
    while True:
        try:
            flash("PRODUCT")
            serve(app, host="0.0.0.0", port=81, _quiet=True, threads=100)
        except subprocess.CalledProcessError as e:
            flash(f"Flask gặp lỗi: {e}")
            flash("Đang khởi động flask...")
            time.sleep(1)  # Đợi một khoảng thời gian trước khi khởi động lại
        except Exception as e:
            flash(f"Lỗi không xác định: {e}")
            flash("Đang khởi động lại flask...")
            time.sleep(1)  # Đợi một khoảng thời gian trước khi khởi động lại



    
            