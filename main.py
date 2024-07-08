from routes import *

if __name__ == "__main__":
    app.logger.info("Khoi dong phan mem ...")
    serve(app, host='0.0.0.0', port=81, threads=16)