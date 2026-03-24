from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
import logging
from logging.handlers import RotatingFileHandler


db = SQLAlchemy()
login_manager = LoginManager()


def init_logging(app):
    """Khởi tạo cấu hình logging dùng chung cho app."""
    handler = RotatingFileHandler('app.log', backupCount=1, encoding='utf-8')
    handler.setLevel(logging.DEBUG)
    formatter = logging.Formatter(
        '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
    )
    handler.setFormatter(formatter)
    app.logger.addHandler(handler)
    app.logger.setLevel(logging.DEBUG)
