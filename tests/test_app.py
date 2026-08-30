import io

import pytest

from app import create_app


@pytest.fixture
def client(tmp_path):
    uploads_dir = tmp_path / "uploads"
    uploads_dir.mkdir()
    config_path = tmp_path / "config.yaml"
    config_path.write_text("kpis: []\ncharts: []\ntables: []\n", encoding="utf-8")
    app = create_app(
        {
            "TESTING": True,
            "UPLOAD_PASSWORD": "test-pass",
            "SECRET_KEY": "test-secret",
            "UPLOAD_DIR": uploads_dir,
            "CONFIG_PATH": config_path,
            "DATA_PATH": tmp_path / "data.json",
        }
    )
    return app.test_client()


def test_index_redirects_to_dashboard(client):
    response = client.get("/")
    assert response.status_code == 302
    assert response.headers["Location"].endswith("/dashboard")


def test_dashboard_is_publicly_accessible(client):
    response = client.get("/dashboard")
    assert response.status_code == 200


def test_api_data_returns_empty_state_when_no_data_json(client):
    response = client.get("/api/data")
    assert response.status_code == 200
    body = response.get_json()
    assert body["kpis"] == []
    assert body["charts"] == []
    assert body["tables"] == []


def test_api_data_returns_saved_data(client, tmp_path):
    from extractor import save_data_json

    save_data_json(
        {"kpis": [{"key": "k"}], "charts": [], "tables": []},
        tmp_path / "data.json",
    )

    response = client.get("/api/data")

    assert response.get_json()["kpis"] == [{"key": "k"}]


def test_login_page_loads(client):
    response = client.get("/login")
    assert response.status_code == 200


def test_login_with_wrong_password_shows_error(client):
    response = client.post("/login", data={"password": "wrong"})
    assert response.status_code == 200
    assert "密码错误".encode() in response.data


def test_login_with_correct_password_redirects_to_upload(client):
    response = client.post("/login", data={"password": "test-pass"})
    assert response.status_code == 302
    assert response.headers["Location"].endswith("/upload")


def test_upload_page_requires_login(client):
    response = client.get("/upload")
    assert response.status_code == 302
    assert "/login" in response.headers["Location"]


def test_upload_page_accessible_after_login(client):
    client.post("/login", data={"password": "test-pass"})
    response = client.get("/upload")
    assert response.status_code == 200


def test_upload_saves_file_and_generates_data_json(client, tmp_path):
    import openpyxl

    client.post("/login", data={"password": "test-pass"})

    (tmp_path / "config.yaml").write_text(
        "kpis:\n"
        "  - key: total_revenue\n"
        "    label: \"总营收\"\n"
        "    source_file: \"经营数据.xlsx\"\n"
        "    sheet: \"汇总\"\n"
        "    mode: fixed_range\n"
        "    range: \"B2\"\n"
        "charts: []\n"
        "tables: []\n",
        encoding="utf-8",
    )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "汇总"
    ws.append(["指标", "数值"])
    ws.append(["总营收", 500])
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    response = client.post(
        "/upload",
        data={"files": (buffer, "经营数据.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )

    assert response.status_code == 200
    assert (tmp_path / "uploads" / "经营数据.xlsx").exists()

    from extractor import load_data_json

    data = load_data_json(tmp_path / "data.json")
    assert data["kpis"][0]["value"] == 500


def test_upload_overwrites_existing_file_with_same_name(client, tmp_path):
    (tmp_path / "uploads" / "经营数据.xlsx").write_bytes(b"old content")
    client.post("/login", data={"password": "test-pass"})

    response = client.post(
        "/upload",
        data={"files": (io.BytesIO(b"new content"), "经营数据.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )

    assert response.status_code == 200
    assert (tmp_path / "uploads" / "经营数据.xlsx").read_bytes() == b"new content"


def test_upload_accepts_multiple_files_in_one_request(client, tmp_path):
    client.post("/login", data={"password": "test-pass"})

    response = client.post(
        "/upload",
        data={
            "files": [
                (io.BytesIO(b"a"), "经营数据.xlsx"),
                (io.BytesIO(b"b"), "销售明细.xlsx"),
            ]
        },
        content_type="multipart/form-data",
        follow_redirects=True,
    )

    assert response.status_code == 200
    assert (tmp_path / "uploads" / "经营数据.xlsx").exists()
    assert (tmp_path / "uploads" / "销售明细.xlsx").exists()


def test_upload_with_corrupted_excel_file_does_not_crash(client, tmp_path):
    client.post("/login", data={"password": "test-pass"})

    (tmp_path / "config.yaml").write_text(
        "kpis:\n"
        "  - key: total_revenue\n"
        "    label: \"总营收\"\n"
        "    source_file: \"损坏文件.xlsx\"\n"
        "    sheet: \"汇总\"\n"
        "    mode: fixed_range\n"
        "    range: \"B2\"\n"
        "charts: []\n"
        "tables: []\n",
        encoding="utf-8",
    )

    response = client.post(
        "/upload",
        data={"files": (io.BytesIO(b"this is not a real excel file, just garbage bytes"), "损坏文件.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )

    assert response.status_code == 200

    from extractor import load_data_json

    data = load_data_json(tmp_path / "data.json")
    assert data["kpis"][0]["value"] is None
    assert data["kpis"][0]["error"]

    # The pipeline must stay usable afterwards: a second, valid upload should
    # still succeed instead of the corrupted file bricking every future /upload.
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "汇总"
    ws.append(["指标", "数值"])
    ws.append(["总营收", 42])
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    (tmp_path / "config.yaml").write_text(
        "kpis:\n"
        "  - key: total_revenue\n"
        "    label: \"总营收\"\n"
        "    source_file: \"经营数据.xlsx\"\n"
        "    sheet: \"汇总\"\n"
        "    mode: fixed_range\n"
        "    range: \"B2\"\n"
        "charts: []\n"
        "tables: []\n",
        encoding="utf-8",
    )

    response2 = client.post(
        "/upload",
        data={"files": (buffer, "经营数据.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )

    assert response2.status_code == 200
    data2 = load_data_json(tmp_path / "data.json")
    assert data2["kpis"][0]["value"] == 42


def test_upload_rejects_path_traversal_filename(client, tmp_path):
    client.post("/login", data={"password": "test-pass"})

    response = client.post(
        "/upload",
        data={"files": (io.BytesIO(b"data"), "../evil.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )

    assert response.status_code == 200
    assert not (tmp_path / "evil.xlsx").exists()
    assert not (tmp_path / "uploads" / "../evil.xlsx").resolve().exists()


def test_upload_rejects_non_excel_extension(client, tmp_path):
    client.post("/login", data={"password": "test-pass"})

    response = client.post(
        "/upload",
        data={"files": (io.BytesIO(b"data"), "malware.exe")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )

    assert response.status_code == 200
    assert not (tmp_path / "uploads" / "malware.exe").exists()


def test_dashboard_page_includes_section_containers(client):
    response = client.get("/dashboard")
    html = response.data.decode("utf-8")

    assert 'id="kpi-section"' in html
    assert 'id="chart-section"' in html
    assert 'id="table-section"' in html
    assert "dashboard.css" in html
    assert "dashboard.js" in html


def test_static_dashboard_css_is_served(client):
    response = client.get("/static/css/dashboard.css")
    assert response.status_code == 200


def test_static_dashboard_js_is_served(client):
    response = client.get("/static/js/dashboard.js")
    assert response.status_code == 200


def test_dashboard_page_includes_echarts_cdn(client):
    response = client.get("/dashboard")
    html = response.data.decode("utf-8")
    assert "echarts" in html.lower()


def test_dashboard_js_includes_table_rendering(client):
    response = client.get("/static/js/dashboard.js")
    body = response.data.decode("utf-8")
    assert "renderTables" in body


def test_dashboard_js_includes_table_view_group_tabs(client):
    response = client.get("/static/js/dashboard.js")
    body = response.data.decode("utf-8")
    assert "renderTableGroup" in body
    assert "view_group" in body


def test_dashboard_js_fetches_api_data_with_relative_path(client):
    response = client.get("/static/js/dashboard.js")
    body = response.data.decode("utf-8")
    assert 'fetch("api/data")' in body


def test_index_redirect_honors_forwarded_prefix(client):
    response = client.get("/", headers={"X-Forwarded-Prefix": "/kanban"})
    assert response.status_code == 302
    assert response.headers["Location"] == "/kanban/dashboard"


def test_dashboard_static_links_honor_forwarded_prefix(client):
    response = client.get("/dashboard", headers={"X-Forwarded-Prefix": "/kanban"})
    html = response.data.decode("utf-8")
    assert "/kanban/static/css/dashboard.css" in html
    assert "/kanban/static/js/dashboard.js" in html
