import os
import secrets
from functools import wraps
from pathlib import Path

from flask import Flask, flash, jsonify, redirect, render_template, request, session, url_for

from extractor import load_data_json


def create_app(test_config=None):
    app = Flask(__name__)
    app.config.update(
        UPLOAD_PASSWORD=os.environ.get("UPLOAD_PASSWORD", "changeme"),
        SECRET_KEY=os.environ.get("FLASK_SECRET_KEY", "dev-secret-key-change-me"),
        UPLOAD_DIR=Path("uploads"),
        CONFIG_PATH=Path("config.yaml"),
        DATA_PATH=Path("data.json"),
    )
    if test_config:
        app.config.update(test_config)

    @app.route("/")
    def index():
        return redirect(url_for("dashboard"))

    @app.route("/dashboard")
    def dashboard():
        return render_template("dashboard.html")

    @app.route("/api/data")
    def api_data():
        data_path = Path(app.config["DATA_PATH"])
        if not data_path.exists():
            return jsonify({"kpis": [], "charts": [], "tables": []})
        return jsonify(load_data_json(data_path))

    def login_required(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if not session.get("logged_in"):
                return redirect(url_for("login"))
            return view(*args, **kwargs)
        return wrapped

    @app.route("/login", methods=["GET", "POST"])
    def login():
        error = None
        if request.method == "POST":
            password = request.form.get("password", "")
            if secrets.compare_digest(password, app.config["UPLOAD_PASSWORD"]):
                session["logged_in"] = True
                return redirect(url_for("upload"))
            error = "密码错误"
        return render_template("login.html", error=error)

    @app.route("/upload", methods=["GET"])
    @login_required
    def upload():
        return render_template("upload.html")

    return app


if __name__ == "__main__":
    if os.environ.get("UPLOAD_PASSWORD") is None:
        print("警告: 未设置环境变量 UPLOAD_PASSWORD，使用默认密码 'changeme'，请在生产环境中修改。")
    flask_app = create_app()
    flask_app.run(host="0.0.0.0", port=5000, debug=False)
