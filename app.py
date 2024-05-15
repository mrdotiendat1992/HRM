from flask import Flask, request, render_template, render_template_string
import json

database = {
    'users': {
        "NT1": [
            {
                "mst": 0,
                "matkhau": "1"
            }
        ],
        "NT2": [
            {
                "mst": 0,
                "matkhau": "1"
            }
        ],
        "NT0": [
            {
                "mst": 0,
                "matkhau": "1"
            }
        ]   
    }
}

app = Flask("hrm")

@app.route("/")
def index():
    return render_template("login.html")

@app.route("/home")
def home():
    return render_template("home.html")

@app.route("/login", methods=["POST"])
def login():
    congty= request.args.get("congty")
    mst = request.args.get("mst")
    matkhau = request.args.get("matkhau")
    print(congty, mst, matkhau)
    if request.method == "POST":
        if mst == "12579" and matkhau == "1":
            return json.dumps({'status':'success', 'message':'Đăng nhập thành công'})
        else:
            return json.dumps({'status':'failed', 'message':'Đăng nhập thất bại'})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=80)