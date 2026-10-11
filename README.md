# LessonLens 教案点评系统

一个前后端分离的多模态 LLM 教案点评系统。用户上传 `.pptx` / `.ppt` 后，后端提取每页文本与结构信息；本机安装 LibreOffice 时会额外生成每页 PNG 预览；配置 LLM 后会向 OpenAI 兼容接口请求 JSON 点评，否则使用本地规则点评，方便先体验完整交互。

## 启动

需要 Python 3.11+。在项目目录执行：

```powershell
Copy-Item .env.example .env
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r backend\requirements.txt
# 可选：在当前 PowerShell 会话中设置 LLM_API_KEY / LLM_BASE_URL / LLM_MODEL
.\.venv\Scripts\python.exe -m uvicorn backend.main:app --host 127.0.0.1 --port 8878
```

浏览器打开 `http://127.0.0.1:8878`。

Windows 上如果安装了 LibreOffice 并且 `soffice` 在 PATH 中，系统会把旧版 `.ppt` 转成 `.pptx`，再生成页面图片预览。没有 LibreOffice 时，`.pptx` 会使用 Pillow 按幻灯片尺寸绘制结构化页面预览，包含背景、图片、形状和文字；旧版 `.ppt` 仍建议安装 LibreOffice 后使用。

## LLM 配置

服务兼容 OpenAI Chat Completions 格式：

```powershell
$env:LLM_API_KEY="你的 API Key"
$env:LLM_BASE_URL="https://api.openai.com/v1"
$env:LLM_MODEL="gpt-4o-mini"
```

LLM 被要求只返回 JSON，服务会直接保存结果。每份点评保存在 `data/jobs.json`，上传文件位于 `data/uploads`，渲染图片位于 `data/renders`。

## 主要接口

- `GET /api/health`：检查服务和 LLM 配置。
- `GET /api/jobs`：点评库列表。
- `GET /api/jobs/{id}`：单份点评详情。
- `POST /api/reviews`：multipart 上传 PPT 文件并返回点评结果。
- `GET /api/jobs/{id}/slides/{file}`：页面 PNG 预览。
