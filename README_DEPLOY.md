# TikTok爆款视频解析&复盘专用｜平铺模块化部署版

当前线上版本采用 **一个 Streamlit 网站 + 根目录平铺多文件** 的结构。

- `爆款拆解`：SOP1，独立业务文件与独立 Gemini 调用文件。
- `爆款对比`：SOP2，独立使用 `gemini-3.8-flash`，不回退到 SOP1 模型。
- `数据复盘`：独立页面与独立 Gemini 调用文件。
- `历史记录`：统一入口，按主账号/分账号权限查看；已支持 Supabase 持久化兼容。
- `app.py` 只负责页面配置、登录与顶部导航。

这样以后修改 SOP1 或 SOP2 时，优先只改各自文件，不需要反复重写整个 `app.py`。

---

## 当前 GitHub 根目录结构

```text
app.py
config.py
requirements.txt

auth_service.py
common.py
history_service.py
export_service.py

gemini_base.py
gemini_sop1.py
gemini_sop2.py
gemini_review.py

sop1_breakdown.py
sop1_schema.py
sop2_compare.py
sop2_schema.py
review_page.py
review_schema.py
review_utils.py
history_page.py

README_DEPLOY.md
.devcontainer/
```

> 当前版本 **不使用** `pages/`、`services/`、`schemas/` 三层文件夹。所有 Python 模块都直接放在仓库根目录，避免 GitHub 网页上传时目录错位导致 Streamlit 找不到模块。

---

## 分支策略

- `backup-v1-stable`：保留旧稳定版，不动。
- `main`：当前线上部署分支。
- `modular-sop2`：模块化开发/历史分支，可保留作为对照。

Streamlit Cloud 的 Main file path 保持：

```text
app.py
```

不需要新建第二个 Streamlit 网站。

---

## Secrets

继续使用 Streamlit Secrets，不要把 `secrets.toml`、API Key 或数据库密钥上传到 GitHub。

需要的配置项按实际启用功能填写，例如：

```toml
GEMINI_API_KEY = "..."
STAFF_PASSWORD = "..."
ADMIN_PASSWORD = "..."
SUPABASE_URL = "..."
SUPABASE_SECRET_KEY = "..."
```

---

## Python 依赖

`requirements.txt` 当前固定：

```text
streamlit==1.62.0
google-genai==2.22.0
pandas==3.0.5
openpyxl==3.1.5
```

---

## SOP1｜爆款拆解

主要维护文件：

```text
sop1_breakdown.py
gemini_sop1.py
sop1_schema.py
```

公共依赖主要来自：

```text
config.py
common.py
gemini_base.py
history_service.py
export_service.py
```

修改 SOP1 时，除非确实涉及共享能力，否则不要修改 SOP2 文件。

---

## SOP2｜爆款对比

主要维护文件：

```text
sop2_compare.py
gemini_sop2.py
sop2_schema.py
```

当前流程：

```text
产品信息
↓
上传 1-3 条爆款
↓
上传 1-3 条我的作品
↓
Gemini 3.8 Flash / LOW：快速预分析
↓
AI 推荐比较组合（只推荐，不代替人选择）
↓
使用人选择 1 条爆款 + 1 条自己的作品
↓
Gemini 3.8 Flash / MEDIUM：直接看两条原视频深度对比
↓
10维差距
↓
优势 / 劣势
↓
重剪价值：高 / 中 / 低
↓
保留 / 删除 / 前移 / 补拍
↓
导出 Excel + ChatGPT JSON
```

SOP2 单次上限：

- 爆款视频最多 3 条
- 我的作品最多 3 条
- 合计最多 6 条
- 小文件自动走 Inline
- 较大文件自动切换 Gemini Files API

SOP2 当前模型：

```text
gemini-3.8-flash
```

预分析使用 LOW thinking，深度对比使用 MEDIUM thinking。

---

## 数据复盘

主要维护文件：

```text
review_page.py
gemini_review.py
review_schema.py
review_utils.py
```

---

## 历史记录

主要维护文件：

```text
history_page.py
history_service.py
auth_service.py
```

历史记录优先写入 Supabase；若 Supabase 暂时不可用，保留本地兼容/待同步机制，避免数据库瞬时故障中断 AI 主流程。

---

## 公共层｜尽量少动

```text
app.py
config.py
common.py
auth_service.py
history_service.py
gemini_base.py
export_service.py
```

公共层只有在多个模块都需要同一能力时才修改。单独调整 SOP1 或 SOP2 的提示词、Schema、页面逻辑时，应优先在对应模块文件内完成。

---

## 上线后核验清单

1. 登录正常。
2. 顶部显示：`爆款拆解 / 爆款对比 / 数据复盘 / 历史记录`。
3. SOP1 能上传多视频并解析。
4. SOP1 切换主参考视频时，卖点与分析跟随正确视频变化。
5. SOP1 能生成 3 个方向与中文执行脚本。
6. SOP2 页面可以正常打开，不再出现 `No module named 'gemini_sop2'`。
7. SOP2 能上传 `1-3 条爆款 + 1-3 条我的作品` 并完成预分析。
8. SOP2 由使用人亲自选择比较对象后，能完成深度对比。
9. SOP2 能导出 Excel 和 ChatGPT JSON。
10. 历史记录能够正常读取、筛选和保存。

---

## 2026-09-07 修复记录

已将错误文件名：

```text
gemini_sop2 .py
```

修正为：

```text
gemini_sop2.py
```

该空格会导致 `sop2_compare.py` 执行 `from gemini_sop2 import ...` 时出现：

```text
ModuleNotFoundError: No module named 'gemini_sop2'
```

当前 `main` 已使用正确文件名。
