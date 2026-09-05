# TikTok Tools Flat Modular Deployment

这是与原模块化版本功能一致的“平铺多文件”部署版，专门用于 GitHub 网页一次性上传。

- 一个网站
- SOP1 / SOP2 / 数据复盘 / 历史记录分别独立文件
- 不需要 pages/services/schemas 文件夹
- app.py 仅负责登录和顶部导航

上传时把本目录内全部文件直接上传到 modular-sop2 分支根目录即可。

# TikTok 爆款视频解析&复盘专用｜模块化部署版

## 这版做了什么

这是从你当前线上稳定版拆出来的模块化版本：

- `爆款拆解`：SOP1，继续使用原有 Gemini 3.5 模型链与原有业务流程。
- `爆款对比`：SOP2，独立使用 `gemini-3.8-flash`，不回退到 SOP1 模型。
- `数据复盘`：从原 `app.py` 独立出来。
- `历史记录`：统一入口，仍按主账号/分账号权限查看。

页面、Gemini 服务、Schema 已分开。以后改 SOP1 主要只动 SOP1 文件；改 SOP2 主要只动 SOP2 文件。

---

## GitHub 部署步骤

### 1. 先备份现在线上稳定版

在 GitHub 当前仓库创建一个分支，例如：

`backup-v1-stable`

这个分支先不要改。

### 2. 把本压缩包解压后的“文件内容”上传到现有仓库根目录

根目录最终应该看到：

```text
app.py
config.py
requirements.txt
pages/
services/
schemas/
.streamlit/
```

注意：不要把最外层 `tiktok_tools_modular/` 再套一层上传，否则 Streamlit Cloud 会找不到根目录的 `app.py`。

### 3. 覆盖原文件

- 用新 `app.py` 覆盖旧 `app.py`
- 用新 `requirements.txt` 覆盖旧文件
- 新增 `config.py`
- 新增 `pages/`、`services/`、`schemas/`、`.streamlit/`

### 4. Streamlit Cloud 设置

Main file path 继续保持：

`app.py`

不需要新建第二个 Streamlit 网站。

### 5. Secrets 不变

继续使用现有 Secrets：

```toml
GEMINI_API_KEY = "你的Key"
STAFF_PASSWORD = "你的员工密码"
ADMIN_PASSWORD = "你的管理员密码"
```

不要把 `secrets.toml` 上传到 GitHub。

### 6. 等待自动重启

`requirements.txt` 已更新为：

```text
streamlit==1.62.0
google-genai==2.22.0
pandas==3.0.5
openpyxl==3.1.5
```

Streamlit Cloud 会重新安装依赖并启动。

---

## 上线后只检查这 8 项

1. 登录正常。
2. 顶部显示：`爆款拆解 / 爆款对比 / 数据复盘 / 历史记录`。
3. SOP1 能正常上传多视频并解析。
4. SOP1 选不同主参考视频时，卖点跟着对应视频变化。
5. SOP1 填写真实卖点后，仍由使用人选择 `爆款为主 / 我的为主 / 融合`。
6. SOP1 能生成 3 个方向与中文执行脚本。
7. SOP2 能上传 `1-3 条爆款 + 1-3 条我的作品`，先预分析，再由使用人亲自选择比较对象，最后深度对比。
8. SOP2 能导出 Excel 和 ChatGPT JSON。

---

## SOP2 当前逻辑

```text
产品信息
↓
上传 1-3 条爆款
↓
上传 1-3 条我的作品
↓
Gemini 3.8 Flash / LOW：快速预分析
↓
AI推荐比较组合（只推荐，不代替人选择）
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

SOP2 单次上限暂定 6 条：3 条爆款 + 3 条自己的作品。总文件较小时自动走 Inline，较大时自动切换 Files API。

---

## 后续怎么维护

### 只改 SOP1

优先修改：

```text
pages/sop1_breakdown.py
services/gemini_sop1.py
schemas/sop1_schema.py
```

### 只改 SOP2

优先修改：

```text
pages/sop2_compare.py
services/gemini_sop2.py
schemas/sop2_schema.py
```

### 公共层

以下文件尽量少动，因为它们是共享基础设施：

```text
app.py
config.py
services/auth_service.py
services/history_service.py
services/gemini_base.py
services/export_service.py
```

