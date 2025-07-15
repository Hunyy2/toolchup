import os
from flask import Flask, request, jsonify
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
import jwt
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv

load_dotenv()
# --- KHỞI TẠO ỨNG DỤNG ---
app = Flask(__name__)

# --- CẤU HÌNH ---
# Lấy các biến môi trường từ Render hoặc dùng giá trị mặc định để test local
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY")
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("DATABASE_URL")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)


# --- TẠO MODEL DATABASE (BẢNG NGƯỜI DÙNG) ---
# --- TẠO MODEL DATABASE (BẢNG NGƯỜI DÙNG) ---
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    is_active = db.Column(db.Boolean, default=True, nullable=False)  # Thêm trường này

    def __repr__(self):
        return f"<User {self.username}>"


# --- CÁC ROUTE (ENDPOINT) CỦA API ---
@app.route("/register", methods=["POST"])
def register():
    data = request.get_json()
    if not data or not "username" in data or not "password" in data:
        return jsonify({"message": "Thiếu username hoặc password"}), 400

    username = data["username"]
    password = data["password"]

    if User.query.filter_by(username=username).first():
        return jsonify({"message": "Tên đăng nhập đã tồn tại"}), 409

    # Băm mật khẩu để bảo mật
    hashed_password = generate_password_hash(password, method="pbkdf2:sha256")
    new_user = User(username=username, password_hash=hashed_password)
    db.session.add(new_user)
    db.session.commit()

    return jsonify({"message": "Đăng ký thành công!"}), 201


@app.route("/login", methods=["POST"])
def login():
    data = request.get_json()
    if not data or not "username" in data or not "password" in data:
        return jsonify({"message": "Thiếu username hoặc password"}), 400

    user = User.query.filter_by(username=data["username"]).first()

    # Kiểm tra người dùng và mật khẩu đã băm
    if not user or not check_password_hash(user.password_hash, data["password"]):
        return jsonify({"message": "Sai tên đăng nhập hoặc mật khẩu!"}), 401

    # Tạo token JWT có thời hạn 8 giờ
    token = jwt.encode(
        {
            "user_id": user.id,
            "exp": datetime.now(timezone.utc) + timedelta(hours=2),
        },
        app.config["SECRET_KEY"],
        algorithm="HS256",
    )

    return jsonify({"token": token})


# Route để kiểm tra token có hợp lệ không (tùy chọn)
@app.route("/validate", methods=["POST"])
def validate_token():
    token = request.headers.get("Authorization")
    if not token:
        return jsonify({"message": "Thiếu token"}), 401

    try:
        # Bỏ qua 'Bearer ' nếu có
        if " " in token:
            token = token.split(" ")[1]

        jwt.decode(token, app.config["SECRET_KEY"], algorithms=["HS256"])
        return jsonify({"message": "Token hợp lệ"}), 200
    except jwt.ExpiredSignatureError:
        return jsonify({"message": "Token đã hết hạn"}), 401
    except jwt.InvalidTokenError:
        return jsonify({"message": "Token không hợp lệ"}), 401


# --- KHỞI CHẠY APP ---
if __name__ == "__main__":
    # Tạo bảng trong database nếu chưa có
    with app.app_context():
        db.create_all()
    app.run(debug=True)
