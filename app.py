import os
from pathlib import Path

from flask import Flask, jsonify, redirect, render_template, url_for

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

    return app


if __name__ == "__main__":
    if os.environ.get("UPLOAD_PASSWORD") is None:
        print("警告: 未设置环境变量 UPLOAD_PASSWORD，使用默认密码 'changeme'，请在生产环境中修改。")
    flask_app = create_app()
    flask_app.run(host="0.0.0.0", port=5000, debug=False)
