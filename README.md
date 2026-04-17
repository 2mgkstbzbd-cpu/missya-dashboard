# 数据分析大屏

这是一个基于 Flask 的数据分析网页，支持上传 Excel/CSV 数据后在服务端完成解析，并在前端大屏中展示汇总指标、趋势图和明细表。

## 本地运行

1. 安装依赖：

```bash
pip install -r requirements.txt
```

2. 启动服务：

```bash
python screen_app.py
```

3. 打开浏览器访问：

```text
http://127.0.0.1:5050
```

默认端口为 `5050`，也可以通过环境变量 `PORT` 覆盖。

## 部署到 Render

仓库已经包含 Render 所需的配置文件 [`render.yaml`](/Users/lkcat/Desktop/codex1/render.yaml)，直接连接 GitHub 仓库即可部署。

### 部署步骤

1. 把项目推送到 GitHub。
2. 登录 [Render](https://render.com/)。
3. 点击 `New` -> `Web Service`。
4. 选择你的 GitHub 仓库。
5. Render 会自动读取 `render.yaml` 并生成部署配置。
6. 确认实例类型为 `Free` 后开始部署。

### 当前 Render 配置

- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn --bind 0.0.0.0:$PORT screen_app:app`
- Runtime: `Python`

## 重要说明

- 这是一个有后端的 Flask 应用，不适合直接部署到 GitHub Pages 这类纯静态托管平台。
- 当前上传的数据保存在服务进程内存中，实例重启或休眠后需要重新上传文件。
- Render 免费实例在长时间无访问后会休眠，首次唤醒会有冷启动等待时间。
- `.gitignore` 默认忽略 Excel/CSV 数据文件和日志文件，避免将本地数据直接上传到 GitHub。

## 建议上传到 GitHub 的文件

- `screen_app.py`
- `templates/`
- `static/`
- `requirements.txt`
- `render.yaml`
- `.gitignore`
- `README.md`
