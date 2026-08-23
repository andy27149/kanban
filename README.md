# 老板看板（Boss Dashboard）

数据维护人通过密码保护的上传页面提交 Excel 文件，系统按 `config.yaml` 配置解析生成看板数据，供老板通过公开链接查看 KPI、图表与明细表格。

## 环境要求

- Python 3.9+

## 安装

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## 配置

1. 设置数据维护人登录密码（默认是 `changeme`，务必修改）：

   ```bash
   export UPLOAD_PASSWORD="你的密码"
   ```

2. 设置 Flask session 密钥（默认是开发用的固定字符串，生产环境务必修改）：

   ```bash
   export FLASK_SECRET_KEY="一个随机字符串"
   ```

3. 编辑 `config.yaml`，为每个 KPI/图表/表格指定数据来源文件名（`source_file`，需与上传时的 Excel 文件名完全一致，支持中文）、sheet 名，以及取数方式：
   - `mode: fixed_range` — 指定固定单元格范围（如 `B2`），适合位置稳定的单一数值。
   - `mode: header_match` — 指定表头文字（如 `header: "利润"`），程序自动定位对应列，能容忍列顺序变化；用在 KPI 上时会对该列所有数值求和。

## 运行

```bash
source venv/bin/activate
python app.py
```

默认监听 `http://localhost:5000`。

## 使用流程

1. 数据维护人打开 `http://localhost:5000/upload`，输入 `UPLOAD_PASSWORD` 登录。
2. 选择一个或多个本地 Excel 文件上传（文件名需与 `config.yaml` 中的 `source_file` 一致）。上传后会覆盖 `uploads/` 目录下的同名文件，并立即基于**整个** `uploads/` 目录重新生成 `data.json`。
3. 页面会提示成功或具体的失败原因（如"找不到表头 '销售额'"），失败的单个指标不影响其余指标正常生成。
4. 老板或任何人打开 `http://localhost:5000/dashboard` 即可查看最新数据，无需密码。

## 运行测试

```bash
source venv/bin/activate
pytest -v
```

## 部署说明

当前架构本地运行与服务器部署无需改动核心代码——只需将 `python app.py` 换成生产级 WSGI 服务器（如 `gunicorn app:create_app()`），并确保 `UPLOAD_PASSWORD`/`FLASK_SECRET_KEY` 通过环境变量正确设置。
