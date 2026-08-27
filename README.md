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
   - `mode: computed`（仅用于 KPI）— 不直接读 Excel，而是用其他已解析出的 KPI 计算得到。通过 `operation` 指定运算方式（目前仅支持 `subtract`，即相减），`from`/`minus` 分别填另外两个 KPI 的 `key`。例如"库存账实差异" = 账面库存 KPI 减去实际盘点库存 KPI：`operation: subtract`、`from: book_stock`、`minus: actual_stock`。被引用的 KPI 必须在同一份 `config.yaml` 中存在且能成功取值，否则该计算型 KPI 会显示取值失败。
   - `mode: group_by_sum`（仅用于表格）— 按 `group_by_header` 指定的列分组，对 `sum_header` 指定的列求和，并按合计值从大到小排序，适合做排名类视图（如"按客户汇总销售额"）。
   - `view_group`/`view_label`（仅用于表格，可选）— 多个表格若填了相同的 `view_group` 值，会被合并展示成一张卡片，通过标签页切换查看；`view_label` 就是每个标签页按钮上显示的文字。例如把"销售明细"（逐条明细）和"按客户汇总"（分组汇总）设成同一个 `view_group`，看板上就会显示一张带"明细"/"按客户汇总"两个标签的卡片。

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

生产环境部署（Docker + 宿主机 nginx 反向代理）见 [DEPLOY.md](DEPLOY.md)。
