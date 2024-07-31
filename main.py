from routes import *

if __name__ == "__main__":
    try:
        type_run = sys.argv[1]
        if type_run == "1": # Chạy phần mềm với các thông số sản phẩm
            while True:
                try:
                    serve(app, host="0.0.0.0", port=81, _quiet=True, threads=8)
                except subprocess.CalledProcessError as e:
                    flash(f"Flask gap loi: {e}")
                    flash("Đang khoi dong flask...")
                    time.sleep(1)  # Đợi một khoảng thời gian trước khi khởi động lại
                except Exception as e:
                    flash(f"Loi khong xac dinh: {e}")
                    flash("Đang khoi dong lai flask ...")
                    time.sleep(1)
                    
        elif type_run == "2": # Chạy phần mềm với các thông số developer
            app.run(host="0.0.0.0", port=81, debug=True)
        
    except:
        print("Vui long chon 1 hoac 2")
        time.sleep(3)
        sys.exit()

            