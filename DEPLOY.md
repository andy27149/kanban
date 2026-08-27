# 部署手册（Docker + 宿主机 nginx）

本项目通过 Docker 打包运行，容器只监听 `127.0.0.1:5000`，由宿主机上已有的 nginx 做反向代理对外提供服务。Docker 镜像本身不包含 nginx。

## 前置条件

- 服务器已安装 Docker 与 Docker Compose（`docker --version` / `docker compose version` 能正常输出）。
- 服务器已安装并运行 nginx（Docker 之外）。

## 首次部署

1. 拉取代码到服务器，例如 `/opt/kanban`，进入目录。

2. 创建环境变量文件：

   ```bash
   cp .env.example .env
   ```

   编辑 `.env`，把 `UPLOAD_PASSWORD`（数据维护人登录密码）和 `FLASK_SECRET_KEY`（session 密钥，可用 `python3 -c "import secrets; print(secrets.token_hex(32))"` 生成）改成真实值。`.env` 不会被提交到 git。

3. 创建持久化数据目录，并预先建好 `data.json`（必须写入合法的空结构，而不是空文件——Docker 会把不存在的文件错误地当成目录挂载，而空文件会导致应用在第一次上传数据前访问看板报 500）：

   ```bash
   mkdir -p data/uploads
   echo '{"kpis": [], "charts": [], "tables": []}' > data/data.json
   ```

4. 确认 `config.yaml` 已按需配置（参见 `README.md` 的"配置"一节）。

5. 一键构建并启动：

   ```bash
   docker compose up -d --build
   ```

6. 验证容器已在本机 5000 端口响应：

   ```bash
   curl -I http://127.0.0.1:5000/dashboard
   ```

   看到 `200 OK`（或跳转正常）即表示应用已启动。

## 配置 nginx 反向代理

在 nginx 里新增一个 server block（示例，按实际域名/证书调整）：

```nginx
server {
    listen 80;
    server_name kanban.example.com;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

改完后 `nginx -t` 测试配置，再 `systemctl reload nginx`（或对应发行版的 reload 命令）。如需 HTTPS，用 certbot 等工具在这个 server block 上申请证书即可，容器内部无需感知。

## 日常操作

- **上传新 Excel 数据**：直接通过浏览器访问 `https://你的域名/upload`，用 `UPLOAD_PASSWORD` 登录后上传，无需重启容器——上传后应用会立即重新生成 `data.json`。
- **修改 `config.yaml`**：`data.json` 只在每次 `/upload` 上传时重新生成（详见 `app.py` 里 `build_dashboard_data` 的唯一调用点），单纯 `docker compose restart` **不会**让新配置生效，看板会继续显示旧配置生成的数据。正确做法：编辑完 `config.yaml` 后，到 `/upload` 页面重新上传一次 Excel（哪怕文件内容没变，重新上传同一份文件也会用最新的 `config.yaml` 重新计算并覆盖 `data.json`）。因为 `config.yaml` 是每次上传请求时才从磁盘读取的，所以既不需要 restart，也不需要重新 build 镜像。
- **更新代码后重新部署**：

  ```bash
  git pull
  docker compose up -d --build
  ```

- **查看日志**：

  ```bash
  docker compose logs -f kanban
  ```

- **备份数据**：直接备份宿主机上的 `data/` 目录（`uploads/` 原始 Excel + `data.json` 解析结果）和 `config.yaml`，例如 `rsync`/`cp` 到异地即可，不依赖容器内部状态。

- **停止服务**：

  ```bash
  docker compose down
  ```

## 排查

- 容器起不来：`docker compose logs kanban` 看报错，常见原因是 `.env` 里密钥留空，或 `data/data.json` 被 Docker 误建成了目录（删掉后按 Step 3 的方法重建成合法的空 JSON 文件，不是空文件）。
- nginx 502：确认容器确实在监听 `127.0.0.1:5000`（`docker compose ps`、`curl http://127.0.0.1:5000/dashboard`），以及 nginx 配置里的 `proxy_pass` 端口一致。
