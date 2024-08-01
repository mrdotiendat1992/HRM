from routes import *

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python main.py <type_run>")
        print("<type_run>:")
        print("  1 - Chạy phần mềm với các thông số sản phẩm")
        print("  2 - Chạy phần mềm với các thư viện")
        sys.exit(1)

    try:
        type_run = sys.argv[1]
        if type_run == "1":  # Chạy phần mềm với các thông số sản phẩm
            while True:
                try:
                    serve(app, host="0.0.0.0", port=81, _quiet=True, threads=8)
                except subprocess.CalledProcessError as e:
                    print(f"Flask gặp lỗi: {e}")
                    print("Đang khởi động flask...")
                    time.sleep(1)  # Đợi một khoảng thời gian trước khi khởi động lại
                except Exception as e:
                    print(f"Lỗi không xác định: {e}")
                    print("Đang khởi động lại flask...")
                    time.sleep(1)  # Đợi một khoảng thời gian trước khi khởi động lại
                else:
                    break
        elif type_run == "2":  # Chạy phần mềm với các thư viện
            serve(app, host="0.0.0.0", port=81, _quiet=True, threads=8)
        else:
            print("Invalid type_run. Use 1 or 2.")
            sys.exit(1)
    except Exception as e:
        print(f"Lỗi: {e}")
        sys.exit(1)

    
            