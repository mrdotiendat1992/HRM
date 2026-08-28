from flask_login import UserMixin

from extensions import db, login_manager


class Nhanvien(UserMixin, db.Model):
    __tablename__ = 'Nhanvien'

    id = db.Column(db.Integer, primary_key=True)
    macongty = db.Column(db.String(10), nullable=False)
    masothe = db.Column(db.Integer, nullable=False)
    hoten = db.Column(db.Unicode(50), nullable=False)
    phongban = db.Column(db.String(10), nullable=False)
    capbac = db.Column(db.String(10), nullable=False)
    phanquyen = db.Column(db.String(10), nullable=False)
    matkhau = db.Column(db.String(10), nullable=False)

    def __repr__(self) -> str:
        return f"<User {self.hoten}>"


@login_manager.user_loader
def load_user(user_id: str):
    """Hàm load user cho flask-login."""
    return Nhanvien.query.get(int(user_id))
