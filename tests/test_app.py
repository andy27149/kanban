import pytest

from app import create_app


@pytest.fixture
def client(tmp_path):
    uploads_dir = tmp_path / "uploads"
    uploads_dir.mkdir()
    app = create_app(
        {
            "TESTING": True,
            "UPLOAD_PASSWORD": "test-pass",
            "SECRET_KEY": "test-secret",
            "UPLOAD_DIR": uploads_dir,
            "CONFIG_PATH": tmp_path / "config.yaml",
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
