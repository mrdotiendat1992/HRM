from routes import app
from waitress import serve

if __name__ == "__main__":
    serve(app, host='0.0.0.0', port=81, threads=16)