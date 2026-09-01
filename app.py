import io
import json
import os
import tempfile
import threading
import time
import uuid
import hashlib
from datetime import datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# 1. APP 配置
# ============================================================

APP_TITLE = "TikTok爆款视频解析&复盘专用"

# 低延迟 / 高吞吐模型
MODEL_NAME = "gemini-3.5-flash-lite"

# 常见 20-40 秒视频优先 Inline Data
INLINE_VIDEO_MAX_MB = 18

# 大文件回退 Files API
FILE_POLL_INTERVAL_SEC = 2
FILE_PROCESS_TIMEOUT_SEC = 180

# 默认登录密码
# 正式部署推荐在 Streamlit Secrets 中覆盖
DEFAULT_STAFF_PASSWORD = "8888"
DEFAULT_ADMIN_PASSWORD = "8888-admin"

# 历史记录
HISTORY_FILE = Path("history_log.csv")
HISTORY_LOCK = threading.Lock()


# ============================================================
# 2. 民宿实景场景库
# ============================================================

SCENE_LIBRARY = {
    "民宿客厅·沙发/茶几休闲区": {
        "prompt": (
            "固定在美国民宿风格客厅，利用沙发、茶几、边几完成真实生活流。"
            "产品必须自然进入日常动作，不使用棚拍式僵硬展示。"
        ),
        "guide": (
            "推荐机位：手持 POV、茶几 45°俯拍、沙发侧前方中近景、产品微距。\n\n"
            "执行重点：第一帧优先展示痛点、反差、结果或最强产品动作；"
            "茶几仅保留必要道具。"
        ),
    },

    "民宿厨房·多功能岛台": {
        "prompt": (
            "固定在民宿厨房岛台或操作台，以连续手部动作、问题展示、"
            "产品解决过程和结果证明为核心。"
        ),
        "guide": (
            "推荐机位：胸口 POV、岛台 45°俯拍、手部超近距、侧前方固定机位。\n\n"
            "执行重点：非常适合厨房垃圾桶、厨房工具、收纳和清洁产品；"
            "台面只保留必要道具。"
        ),
    },

    "民宿卧室·床头/梳妆台": {
        "prompt": (
            "固定在民宿卧室，重点使用床头柜、床面或梳妆台，"
            "呈现真实拿取、使用和放回动作。"
        ),
        "guide": (
            "推荐机位：床头近景、梳妆台 45°、第一人称拿取、床面俯拍。\n\n"
            "执行重点：适合耳机、便携用品、阅读用品和个人护理产品；"
            "让产品自然成为生活流程的一部分。"
        ),
    },

    "民宿卫生间/浴室·镜前特写": {
        "prompt": (
            "固定在卫生间、浴室、洗手台或镜柜区域。"
            "先放大用户真实痛点，再展示产品解决动作。"
        ),
        "guide": (
            "推荐机位：镜前中近景、洗手台 45°、手部微距、镜柜第一人称。\n\n"
            "执行重点：避免镜中穿帮、严重反光以及杂乱洗护用品。"
        ),
    },

    "民宿阳台/落地窗·自然光": {
        "prompt": (
            "固定在阳台、落地窗或窗边自然光区域，"
            "利用自然光展示产品材质、便携性和真实生活方式。"
        ),
        "guide": (
            "推荐机位：窗边侧光、手持 POV、产品近景、中景生活流。\n\n"
            "执行重点：适合耳机、便携用品和家居用品；"
            "避免强逆光导致产品主体过暗。"
        ),
    },

    "纯桌面特写·无真人露脸": {
        "prompt": (
            "整个视频只允许桌面、产品、双手和必要道具出现，"
            "重点展示功能动作、产品细节、Before/After 和结果证明。"
        ),
        "guide": (
            "推荐机位：正上方俯拍、45°俯拍、超近距功能特写、前后对比。\n\n"
            "执行重点：特别适合隐私保护印章等功能型产品；"
            "第一帧直接进入最强功能动作。"
        ),
    },
}


# ============================================================
# 3. 产品品类 / 流量类型
# ============================================================

PRODUCT_CATEGORIES = [
    "隐私保护印章 / Privacy Roller Stamp",
    "挂式厨房垃圾桶 / Hanging Kitchen Trash Can",
    "无线蓝牙耳机 / Wireless Bluetooth Earbuds",
    "厨房工具 / Kitchen Gadgets",
    "家居收纳 / Home Organization",
    "个人护理 / Personal Care",
    "便携小工具 / Portable Gadgets",
    "其他",
]

TRAFFIC_TYPES = [
    "自然流 Organic",
    "Custom Mode Video Shopping Ads",
    "自然流 + 付费混合",
]


# ============================================================
# 4. Excel 字段
# ============================================================

EXCEL_COLUMNS = [
    "分镜序号",
    "景别/机位",
    "画面描述(道具/动作)",
    "英文口播文案/字幕",
    "音效/节奏提示",
    "设计目的(底层逻辑)",
]


# ============================================================
# 5. 历史记录字段
# ============================================================

HISTORY_COLUMNS = [
    "record_id",
    "created_at_utc",
    "created_at_cn",
    "record_type",
    "role",
    "operator",
    "tiktok_account",
    "product_category",
    "product_name",
    "selling_points",
    "selling_point_version",
    "scene",
    "traffic_type",
    "video_name",
    "video_size_mb",
    "analysis_mode",
    "analysis_seconds",
    "hook_summary",
    "conversion_logic",
    "priority_issue",
    "diagnosis_summary",
    "account_diagnosis",
    "metrics_json",
    "account_baseline_json",
    "metric_assessment_json",
    "full_output_json",
    "script_markdown",
    "original_script",
]


# ============================================================
# 6. 内部复盘工作区间
# 不是 TikTok 官方 benchmark
# 有账号基准时优先按账号自身基准判断
# ============================================================

SOP_BANDS = {
    "retention_3s_pct": {
        "low": 55.0,
        "high": 70.0,
    },
    "completion_rate_pct": {
        "low": 15.0,
        "high": 25.0,
    },
    "product_ctr_pct": {
        "low": 1.5,
        "high": 3.0,
    },
    "order_conversion_pct": {
        "low": 2.0,
        "high": 5.0,
    },
    "engagement_rate_pct": {
        "low": 1.5,
        "high": 3.0,
    },
}


# ============================================================
# 7. JSON Schema
# ============================================================

SELLING_POINTS_SCHEMA = {
    "type": "object",
    "properties": {
        "video_product_insight": {
            "type": "string",
        },
        "hook_summary": {
            "type": "string",
        },
        "conversion_logic": {
            "type": "string",
        },
        "options": {
            "type": "array",
            "minItems": 3,
            "maxItems": 3,
            "items": {
                "type": "object",
                "properties": {
                    "name": {
                        "type": "string",
                    },
                    "angle": {
                        "type": "string",
                    },
                    "selling_points": {
                        "type": "string",
                    },
                    "reason": {
                        "type": "string",
                    },
                },
                "required": [
                    "name",
                    "angle",
                    "selling_points",
                    "reason",
                ],
            },
        },
    },
    "required": [
        "video_product_insight",
        "hook_summary",
        "conversion_logic",
        "options",
    ],
}


STORYBOARD_SCHEMA = {
    "type": "object",
    "properties": {
        "storyboard": {
            "type": "array",
            "minItems": 5,
            "maxItems": 10,
            "items": {
                "type": "object",
                "properties": {
                    "sequence": {
                        "type": "string",
                    },
                    "shot": {
                        "type": "string",
                    },
                    "visual": {
                        "type": "string",
                    },
                    "copy_en": {
                        "type": "string",
                    },
                    "audio": {
                        "type": "string",
                    },
                    "rationale": {
                        "type": "string",
                    },
                },
                "required": [
                    "sequence",
                    "shot",
                    "visual",
                    "copy_en",
                    "audio",
                    "rationale",
                ],
            },
        },
    },
    "required": [
        "storyboard",
    ],
}


REVIEW_SCHEMA = {
    "type": "object",
    "properties": {
        "priority_issue": {
            "type": "string",
        },
        "diagnosis_summary": {
            "type": "string",
        },
        "account_diagnosis": {
            "type": "string",
        },
        "metric_diagnosis": {
            "type": "array",
            "minItems": 5,
            "maxItems": 8,
            "items": {
                "type": "object",
                "properties": {
                    "metric": {
                        "type": "string",
                    },
                    "status": {
                        "type": "string",
                    },
                    "meaning": {
                        "type": "string",
                    },
                    "action": {
                        "type": "string",
                    },
                },
                "required": [
                    "metric",
                    "status",
                    "meaning",
                    "action",
                ],
            },
        },
        "priority_actions": {
            "type": "array",
            "minItems": 3,
            "maxItems": 6,
            "items": {
                "type": "string",
            },
        },
        "storyboard": {
            "type": "array",
            "minItems": 5,
            "maxItems": 10,
            "items": {
                "type": "object",
                "properties": {
                    "sequence": {
                        "type": "string",
                    },
                    "shot": {
                        "type": "string",
                    },
                    "visual": {
                        "type": "string",
                    },
                    "copy_en": {
                        "type": "string",
                    },
                    "audio": {
                        "type": "string",
                    },
                    "rationale": {
                        "type": "string",
                    },
                },
                "required": [
                    "sequence",
                    "shot",
                    "visual",
                    "copy_en",
                    "audio",
                    "rationale",
                ],
            },
        },
    },
    "required": [
        "priority_issue",
        "diagnosis_summary",
        "account_diagnosis",
        "metric_diagnosis",
        "priority_actions",
        "storyboard",
    ],
}


# ============================================================
# 8. Streamlit 页面
# ============================================================

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>

    .block-container {
        max-width: 1180px;
        padding-top: 1.3rem;
        padding-bottom: 3rem;
    }

    h1 {
        font-size: 1.8rem !important;
        margin-bottom: 1rem !important;
    }

    h2 {
        font-size: 1.45rem !important;
    }

    h3 {
        font-size: 1.12rem !important;
        margin-top: 1.1rem !important;
        margin-bottom: .7rem !important;
    }

    div[data-testid="stButton"] button,
    div[data-testid="stDownloadButton"] button,
    div[data-testid="stFormSubmitButton"] button {
        border-radius: 8px;
        font-weight: 600;
        min-height: 42px;
    }

    div[data-testid="stForm"] {
        border: none;
        padding: 0;
    }

    .summary-card {
        padding: .85rem 1rem;
        border-radius: 10px;
        background: rgba(100,116,139,.06);
        border: 1px solid rgba(100,116,139,.12);
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 9. Secrets / Client
# ============================================================

def get_secret(name, default=""):
    try:
        value = st.secrets[name]
    except Exception:
        return default

    if value is None:
        return default

    return str(value).strip()


def get_api_key():
    return get_secret(
        "GEMINI_API_KEY",
        "",
    )


def create_gemini_client(api_key):
    return genai.Client(
        api_key=api_key
    )


# ============================================================
# 10. Session
# ============================================================

def initialize_session_state():
    defaults = {
        "authenticated": False,
        "role": "",
        "operator": "",

        "selling_point_analysis": None,
        "selling_point_options": [],
        "selling_point_metadata": {},

        "final_selling_points": "",
        "selected_selling_point_version": "",

        "review_original_script": "",
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def logout():
    st.session_state["authenticated"] = False
    st.session_state["role"] = ""
    st.session_state["operator"] = ""
    st.rerun()


initialize_session_state()


# ============================================================
# 11. 通用工具
# ============================================================

def clean_text(value):
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    return str(value).strip()


def compact_dict(data):
    return {
        key: value
        for key, value in data.items()
        if value is not None and value != ""
    }


def json_dumps(data):
    return json.dumps(
        data,
        ensure_ascii=False,
        separators=(",", ":"),
    )


def parse_json_output(raw_text):
    if not raw_text:
        raise ValueError(
            "AI未返回有效结果，请重新执行。"
        )

    try:
        return json.loads(raw_text)

    except json.JSONDecodeError as exc:
        raise ValueError(
            "AI返回格式异常，请重新执行。"
        ) from exc


def get_video_signature(uploaded_video):
    if uploaded_video is None:
        return ""

    return hashlib.sha256(
        uploaded_video.getvalue()
    ).hexdigest()[:20]


def thinking_config():
    return types.ThinkingConfig(
        thinking_level="minimal"
    )


# ============================================================
# 12. History
# ============================================================

def empty_history_dataframe():
    return pd.DataFrame(
        columns=HISTORY_COLUMNS
    )


def normalize_history_dataframe(dataframe):
    if dataframe is None:
        return empty_history_dataframe()

    dataframe = dataframe.copy()

    for column in HISTORY_COLUMNS:
        if column not in dataframe.columns:
            dataframe[column] = ""

    return dataframe.reindex(
        columns=HISTORY_COLUMNS
    )


def load_history():
    if not HISTORY_FILE.exists():
        return empty_history_dataframe()

    try:
        dataframe = pd.read_csv(
            HISTORY_FILE,
            dtype=str,
            keep_default_na=False,
            encoding="utf-8-sig",
        )

        return normalize_history_dataframe(
            dataframe
        )

    except Exception:
        return empty_history_dataframe()


def write_history(dataframe):
    dataframe = normalize_history_dataframe(
        dataframe
    )

    temporary_file = HISTORY_FILE.with_name(
        "history_log.tmp.csv"
    )

    dataframe.to_csv(
        temporary_file,
        index=False,
        encoding="utf-8-sig",
    )

    os.replace(
        temporary_file,
        HISTORY_FILE,
    )


def append_history(record):
    row = {
        column: clean_text(
            record.get(column, "")
        )
        for column in HISTORY_COLUMNS
    }

    if not row["record_id"]:
        row["record_id"] = uuid.uuid4().hex[:12]

    now = datetime.now(
        timezone.utc
    )

    row["created_at_utc"] = now.isoformat(
        timespec="seconds"
    )

    row["created_at_cn"] = (
        now
        .astimezone(
            ZoneInfo("Asia/Shanghai")
        )
        .strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    )

    with HISTORY_LOCK:
        current = load_history()

        updated = pd.concat(
            [
                current,
                pd.DataFrame([row]),
            ],
            ignore_index=True,
        )

        write_history(
            updated
        )


def scoped_history():
    dataframe = load_history()

    if st.session_state["role"] == "主账号(Admin)":
        return dataframe

    return dataframe[
        dataframe["operator"]
        == st.session_state["operator"]
    ].copy()


# ============================================================
# 13. 登录
# ============================================================

def render_login_sidebar():
    staff_password = get_secret(
        "STAFF_PASSWORD",
        DEFAULT_STAFF_PASSWORD,
    )

    admin_password = get_secret(
        "ADMIN_PASSWORD",
        DEFAULT_ADMIN_PASSWORD,
    )

    with st.sidebar:
        st.markdown("### 团队登录")

        if st.session_state["authenticated"]:
            st.success(
                f'{st.session_state["role"]} · '
                f'{st.session_state["operator"]}'
            )

            if st.button(
                "退出",
                use_container_width=True,
            ):
                logout()

            return

        role = st.selectbox(
            "身份",
            [
                "分账号(运营专员)",
                "主账号(Admin)",
            ],
        )

        operator = st.text_input(
            "操作人",
            placeholder="例如：小凡",
        )

        password = st.text_input(
            "密码",
            type="password",
        )

        if st.button(
            "登录",
            type="primary",
            use_container_width=True,
        ):
            expected = (
                admin_password
                if role == "主账号(Admin)"
                else staff_password
            )

            if not operator.strip():
                st.error(
                    "请输入操作人。"
                )

            elif password != expected:
                st.error(
                    "密码错误。"
                )

            else:
                st.session_state["authenticated"] = True
                st.session_state["role"] = role
                st.session_state["operator"] = operator.strip()

                st.rerun()


# ============================================================
# 14. 侧栏历史
# ============================================================

def render_sidebar_history():
    if not st.session_state["authenticated"]:
        return

    history = scoped_history()

    with st.sidebar:
        st.divider()
        st.markdown("### 历史")

        if history.empty:
            st.caption(
                "暂无记录"
            )
            return

        filtered = history.copy()

        if st.session_state["role"] == "主账号(Admin)":
            operators = ["全部"] + sorted(
                [
                    x
                    for x in filtered["operator"].unique()
                    if x
                ]
            )

            selected_operator = st.selectbox(
                "操作人",
                operators,
                key="sidebar_operator_filter",
            )

            if selected_operator != "全部":
                filtered = filtered[
                    filtered["operator"]
                    == selected_operator
                ]

        record_types = ["全部"] + sorted(
            [
                x
                for x in filtered["record_type"].unique()
                if x
            ]
        )

        selected_type = st.selectbox(
            "类型",
            record_types,
            key="sidebar_record_filter",
        )

        if selected_type != "全部":
            filtered = filtered[
                filtered["record_type"]
                == selected_type
            ]

        st.caption(
            f"{len(filtered)} 条记录"
        )

        csv_bytes = filtered.to_csv(
            index=False,
            encoding="utf-8-sig",
        ).encode(
            "utf-8-sig"
        )

        st.download_button(
            "下载历史 CSV",
            data=csv_bytes,
            file_name=(
                "history_"
                + datetime.now().strftime(
                    "%Y%m%d_%H%M"
                )
                + ".csv"
            ),
            mime="text/csv",
            use_container_width=True,
        )


# ============================================================
# 15. Excel 读取
# ============================================================

def excel_script_to_text(uploaded_file):
    try:
        dataframe = pd.read_excel(
            uploaded_file,
            engine="openpyxl",
        )

    except Exception as exc:
        raise ValueError(
            f"Excel读取失败：{exc}"
        )

    if dataframe.empty:
        raise ValueError(
            "Excel中没有可读取脚本。"
        )

    missing = [
        column
        for column in EXCEL_COLUMNS
        if column not in dataframe.columns
    ]

    if missing:
        raise ValueError(
            "Excel缺少字段："
            + "、".join(missing)
        )

    blocks = []

    for _, row in dataframe.iterrows():
        blocks.append(
            f"""
分镜 {clean_text(row.get("分镜序号"))}
景别/机位：{clean_text(row.get("景别/机位"))}
画面描述：{clean_text(row.get("画面描述(道具/动作)"))}
英文口播/字幕：{clean_text(row.get("英文口播文案/字幕"))}
音效/节奏：{clean_text(row.get("音效/节奏提示"))}
设计目的：{clean_text(row.get("设计目的(底层逻辑)"))}
""".strip()
        )

    return "\n\n".join(
        blocks
    )


# ============================================================
# 16. 分镜 / Excel 输出
# ============================================================

def storyboard_to_dataframe(rows):
    result = []

    for index, row in enumerate(
        rows or [],
        start=1,
    ):
        result.append(
            {
                "分镜序号":
                    clean_text(
                        row.get("sequence")
                    )
                    or str(index),

                "景别/机位":
                    clean_text(
                        row.get("shot")
                    ),

                "画面描述(道具/动作)":
                    clean_text(
                        row.get("visual")
                    ),

                "英文口播文案/字幕":
                    clean_text(
                        row.get("copy_en")
                    ),

                "音效/节奏提示":
                    clean_text(
                        row.get("audio")
                    ),

                "设计目的(底层逻辑)":
                    clean_text(
                        row.get("rationale")
                    ),
            }
        )

    return pd.DataFrame(
        result,
        columns=EXCEL_COLUMNS,
    )


def markdown_escape(value):
    return (
        clean_text(value)
        .replace("\\", "\\\\")
        .replace("|", "\\|")
        .replace("\r\n", "<br>")
        .replace("\n", "<br>")
    )


def dataframe_to_markdown(dataframe):
    header = (
        "| "
        + " | ".join(EXCEL_COLUMNS)
        + " |"
    )

    divider = (
        "| "
        + " | ".join(
            ["---"] * len(EXCEL_COLUMNS)
        )
        + " |"
    )

    rows = []

    for _, record in dataframe.iterrows():
        rows.append(
            "| "
            + " | ".join(
                markdown_escape(
                    record[column]
                )
                for column in EXCEL_COLUMNS
            )
            + " |"
        )

    return "\n".join(
        [
            header,
            divider,
            *rows,
        ]
    )


def dataframe_to_excel_bytes(
    dataframe,
    sheet_name,
):
    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        dataframe.to_excel(
            writer,
            index=False,
            sheet_name=sheet_name,
        )

        worksheet = writer.book[
            sheet_name
        ]

        fill = PatternFill(
            "solid",
            fgColor="1F2937",
        )

        font = Font(
            color="FFFFFF",
            bold=True,
        )

        for cell in worksheet[1]:
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        widths = {
            1: 11,
            2: 18,
            3: 46,
            4: 44,
            5: 25,
            6: 42,
        }

        for index, width in widths.items():
            worksheet.column_dimensions[
                get_column_letter(index)
            ].width = width

        for row in worksheet.iter_rows(
            min_row=2
        ):
            for cell in row:
                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True,
                )

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

    output.seek(0)

    return output.getvalue()


# ============================================================
# 17. Files API 大文件回退
# ============================================================

def wait_until_file_active(
    client,
    uploaded_file,
):
    started = time.monotonic()

    current = uploaded_file

    while True:
        state = getattr(
            current,
            "state",
            None,
        )

        state_name = getattr(
            state,
            "name",
            "",
        )

        if state_name == "ACTIVE":
            return current

        if state_name in {
            "FAILED",
            "ERROR",
        }:
            raise RuntimeError(
                f"视频处理失败：{state_name}"
            )

        if (
            time.monotonic()
            - started
            > FILE_PROCESS_TIMEOUT_SEC
        ):
            raise TimeoutError(
                "视频处理超时，请压缩视频后重新上传。"
            )

        time.sleep(
            FILE_POLL_INTERVAL_SEC
        )

        current = client.files.get(
            name=current.name
        )


# ============================================================
# 18. 爆款视频解析
# ============================================================

def build_selling_point_prompt(
    category,
    product_name,
):
    return f"""
你负责分析美国 TikTok Shop 20-40秒爆款商品视频。

产品品类：
{category}

产品：
{product_name}

分析上传的视频画面、字幕、口播、音效、Demo和剪辑节奏。

输出：
1. 用户真正被什么产品价值打动。
2. 前0-3秒Hook为什么能留人。
3. 视频从Hook到购买的转化逻辑。
4. 三种不同卖点角度：

A 痛点驱动
用户为什么现在需要它。

B 功能证明
什么Demo最能证明产品价值。

C 场景利益
进入日常生活以后具体哪里更方便。

每套3-5条核心卖点。

selling_points必须使用自然简洁的美式英语，
使用分号分隔。

不得复制品牌、完整原句或独特剧情。
不得虚构功能、认证、折扣、销量或医疗安全声明。

严格输出JSON Schema。
""".strip()


def analyze_video_selling_points(
    client,
    uploaded_video,
    category,
    product_name,
):
    started = time.perf_counter()

    video_bytes = uploaded_video.getvalue()

    size_mb = (
        len(video_bytes)
        / 1024
        / 1024
    )

    mime_type = (
        uploaded_video.type
        or "video/mp4"
    )

    prompt = build_selling_point_prompt(
        category,
        product_name,
    )

    remote_file = None
    temp_path = None

    try:
        if size_mb <= INLINE_VIDEO_MAX_MB:
            analysis_mode = "快速解析"

            video_part = types.Part.from_bytes(
                data=video_bytes,
                mime_type=mime_type,
            )

            response = client.models.generate_content(
                model=MODEL_NAME,

                contents=[
                    video_part,
                    prompt,
                ],

                config=types.GenerateContentConfig(
                    thinking_config=thinking_config(),
                    max_output_tokens=2200,
                    response_mime_type="application/json",
                    response_json_schema=SELLING_POINTS_SCHEMA,
                ),
            )

        else:
            analysis_mode = "大文件解析"

            with tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".mp4",
            ) as temp:
                temp.write(
                    video_bytes
                )

                temp_path = temp.name

            remote_file = client.files.upload(
                file=temp_path
            )

            remote_file = wait_until_file_active(
                client,
                remote_file,
            )

            response = client.models.generate_content(
                model=MODEL_NAME,

                contents=[
                    remote_file,
                    prompt,
                ],

                config=types.GenerateContentConfig(
                    thinking_config=thinking_config(),
                    max_output_tokens=2200,
                    response_mime_type="application/json",
                    response_json_schema=SELLING_POINTS_SCHEMA,
                ),
            )

        result = parse_json_output(
            response.text
        )

        return result, {
            "video_size_mb": round(
                size_mb,
                2,
            ),
            "analysis_mode": analysis_mode,
            "analysis_seconds": round(
                time.perf_counter()
                - started,
                1,
            ),
        }

    finally:
        if remote_file is not None:
            try:
                client.files.delete(
                    name=remote_file.name
                )
            except Exception:
                pass

        if (
            temp_path
            and os.path.exists(temp_path)
        ):
            try:
                os.remove(
                    temp_path
                )
            except OSError:
                pass


# ============================================================
# 19. 新脚本生成
# ============================================================

def build_storyboard_prompt(
    category,
    product_name,
    selling_points,
    scene,
    video_analysis,
):
    scene_info = SCENE_LIBRARY[
        scene
    ]

    return f"""
你负责美国 TikTok Shop 15-40秒短视频拍摄SOP。

产品品类：
{category}

产品：
{product_name}

最终确认卖点：
{selling_points}

爆款视频前3秒机制：
{video_analysis.get("hook_summary", "")}

爆款视频转化机制：
{video_analysis.get("conversion_logic", "")}

爆款产品洞察：
{video_analysis.get("video_product_insight", "")}

固定拍摄场景：
{scene}

场景要求：
{scene_info["prompt"]}

生成一套全新的5-10个分镜。

要求：
- 只学习底层机制，不复制原视频。
- 第一帧直接进入痛点、结果、反差或最强产品动作。
- 前3秒必须具体写出画面、动作、字幕/口播。
- 产品Demo尽早出现。
- 口播使用自然美式英语和短句。
- 全部镜头必须能在当前民宿真实落地。
- 不虚构产品能力。
- 每个分镜明确解决停留、理解、信任、欲望、点击或成交。
- 严格输出JSON Schema。
""".strip()


def generate_storyboard(
    client,
    category,
    product_name,
    selling_points,
    scene,
    video_analysis,
):
    started = time.perf_counter()

    prompt = build_storyboard_prompt(
        category,
        product_name,
        selling_points,
        scene,
        video_analysis,
    )

    response = client.models.generate_content(
        model=MODEL_NAME,

        contents=prompt,

        config=types.GenerateContentConfig(
            thinking_config=thinking_config(),
            max_output_tokens=2600,
            response_mime_type="application/json",
            response_json_schema=STORYBOARD_SCHEMA,
        ),
    )

    result = parse_json_output(
        response.text
    )

    return result, round(
        time.perf_counter()
        - started,
        1,
    )


# ============================================================
# 20. 数据复盘
# ============================================================

def grade_metric(
    key,
    value,
):
    if value is None:
        return "未填写"

    band = SOP_BANDS[key]

    if value < band["low"]:
        return "偏低"

    if value >= band["high"]:
        return "较强"

    return "中间区间"


def compare_metric(
    key,
    value,
    baseline,
):
    if value is None:
        return {
            "status": "未填写"
        }

    if (
        baseline is not None
        and baseline > 0
    ):
        ratio = (
            value
            / baseline
        )

        if ratio < 0.8:
            status = "明显低于账号基准"

        elif ratio > 1.2:
            status = "明显高于账号基准"

        else:
            status = "接近账号基准"

        return {
            "value": value,
            "account_baseline": baseline,
            "ratio": round(
                ratio,
                3,
            ),
            "status": status,
            "basis": "账号近7天同类视频",
        }

    return {
        "value": value,
        "status": grade_metric(
            key,
            value,
        ),
        "basis": "内部SOP工作区间",
    }


def build_metric_assessment(
    metrics,
    baseline,
):
    assessment = {
        "前3秒留存率": compare_metric(
            "retention_3s_pct",
            metrics.get(
                "retention_3s_pct"
            ),
            baseline.get(
                "retention_3s_pct"
            ),
        ),

        "平均完播率": compare_metric(
            "completion_rate_pct",
            metrics.get(
                "completion_rate_pct"
            ),
            baseline.get(
                "completion_rate_pct"
            ),
        ),

        "商品CTR": compare_metric(
            "product_ctr_pct",
            metrics.get(
                "product_ctr_pct"
            ),
            baseline.get(
                "product_ctr_pct"
            ),
        ),

        "订单转化率": compare_metric(
            "order_conversion_pct",
            metrics.get(
                "order_conversion_pct"
            ),
            baseline.get(
                "order_conversion_pct"
            ),
        ),

        "互动率": compare_metric(
            "engagement_rate_pct",
            metrics.get(
                "engagement_rate_pct"
            ),
            baseline.get(
                "engagement_rate_pct"
            ),
        ),
    }

    actual_roas = metrics.get(
        "actual_roas"
    )

    target_roas = metrics.get(
        "target_roas"
    )

    if (
        actual_roas is not None
        and target_roas is not None
        and target_roas > 0
    ):
        assessment["ROAS"] = {
            "actual": actual_roas,
            "target": target_roas,
            "status": (
                "达标"
                if actual_roas >= target_roas
                else "未达标"
            ),
        }

    actual_cpc = metrics.get(
        "actual_cpc"
    )

    target_cpc = metrics.get(
        "target_cpc"
    )

    if (
        actual_cpc is not None
        and target_cpc is not None
        and target_cpc > 0
    ):
        assessment["CPC"] = {
            "actual": actual_cpc,
            "target": target_cpc,
            "status": (
                "达标"
                if actual_cpc <= target_cpc
                else "高于目标"
            ),
        }

    return assessment


def build_review_prompt(
    category,
    product_name,
    selling_points,
    scene,
    traffic_type,
    original_script,
    metrics,
    baseline,
    assessment,
):
    scene_info = SCENE_LIBRARY[
        scene
    ]

    return f"""
你是美国 TikTok Shop 短视频数据复盘负责人。

产品：
{category}
{product_name}

卖点：
{selling_points}

流量：
{traffic_type}

下一版场景：
{scene}

场景要求：
{scene_info["prompt"]}

原始脚本：
{original_script}

本条数据：
{json.dumps(metrics, ensure_ascii=False)}

账号基准：
{json.dumps(baseline, ensure_ascii=False)}

系统判断：
{json.dumps(assessment, ensure_ascii=False)}

严格按漏斗判断：

前3秒留存 → Hook
完播率 → 中段节奏
互动率 → 美国受众共鸣
商品CTR → 商品点击兴趣
订单转化率 → 成交能力
ROAS/CPC → 付费流量质量

规则：
- 必须引用具体数值。
- 有账号基准时优先对比账号自身。
- 没账号基准时才能使用内部SOP区间。
- 内部区间不是TikTok官方标准。
- 找出最优先修复的1个环节。
- 下一版优先只修复这个环节。
- 生成5-10个民宿可执行分镜。
- 英文必须自然美式口语。
- 不虚构产品功能。
- 严格输出JSON Schema。
""".strip()


def review_script(
    client,
    category,
    product_name,
    selling_points,
    scene,
    traffic_type,
    original_script,
    metrics,
    baseline,
    assessment,
):
    started = time.perf_counter()

    response = client.models.generate_content(
        model=MODEL_NAME,

        contents=build_review_prompt(
            category,
            product_name,
            selling_points,
            scene,
            traffic_type,
            original_script,
            metrics,
            baseline,
            assessment,
        ),

        config=types.GenerateContentConfig(
            thinking_config=thinking_config(),
            max_output_tokens=2800,
            response_mime_type="application/json",
            response_json_schema=REVIEW_SCHEMA,
        ),
    )

    result = parse_json_output(
        response.text
    )

    return result, round(
        time.perf_counter()
        - started,
        1,
    )


# ============================================================
# 21. Number Input
# ============================================================

def optional_number(
    label,
    key,
    max_value=None,
    step=0.1,
):
    return st.number_input(
        label,
        min_value=0.0,
        max_value=max_value,
        value=None,
        step=step,
        key=key,
        placeholder="可留空",
    )


# ============================================================
# 22. 卖点自动同步
# ============================================================

def sync_selling_point_choice():
    options = st.session_state.get(
        "selling_point_options",
        [],
    )

    selected_index = st.session_state.get(
        "selling_point_choice",
        0,
    )

    if (
        options
        and 0 <= selected_index < len(options)
    ):
        option = options[
            selected_index
        ]

        st.session_state[
            "final_selling_points"
        ] = option.get(
            "selling_points",
            "",
        )

        st.session_state[
            "selected_selling_point_version"
        ] = option.get(
            "name",
            "",
        )


# ============================================================
# 23. 初始化
# ============================================================

render_login_sidebar()

st.title(
    APP_TITLE
)

if not st.session_state["authenticated"]:
    st.info(
        "请先在左侧登录。"
    )
    st.stop()

render_sidebar_history()

api_key = get_api_key()

if not api_key:
    st.error(
        "系统未配置 Gemini API Key，请联系管理员。"
    )
    client = None

else:
    client = create_gemini_client(
        api_key
    )


# ============================================================
# 24. Tabs
# ============================================================

tab_generate, tab_review, tab_history = st.tabs(
    [
        "爆款拆解",
        "数据复盘",
        "历史记录",
    ]
)


# ============================================================
# TAB 1：爆款拆解
# ============================================================

with tab_generate:

    # --------------------------------------------------------
    # 产品
    # --------------------------------------------------------

    st.markdown("### ① 产品")

    c1, c2, c3 = st.columns(3)

    with c1:
        generate_account = st.text_input(
            "TikTok账号",
            key="generate_account",
            placeholder="用于归档，例如：官号",
            help="当前只用于历史归档和后续账号基准分析，不会在线抓取TikTok账号数据。",
        )

    with c2:
        generate_category = st.selectbox(
            "产品品类",
            PRODUCT_CATEGORIES,
            key="generate_category",
        )

    with c3:
        generate_product_name = st.text_input(
            "产品名称 / SKU",
            key="generate_product_name",
            placeholder="例如：隐私印章",
        )

    # --------------------------------------------------------
    # 视频
    # --------------------------------------------------------

    st.markdown("### ② 爆款视频")

    uploaded_video = st.file_uploader(
        "上传 .mp4",
        type=["mp4"],
        accept_multiple_files=False,
        key="generate_video",
        label_visibility="collapsed",
    )

    if uploaded_video is not None:
        size_mb = (
            len(uploaded_video.getvalue())
            / 1024
            / 1024
        )

        signature = get_video_signature(
            uploaded_video
        )

        if (
            st.session_state.get(
                "active_video_signature"
            )
            != signature
        ):
            st.session_state[
                "active_video_signature"
            ] = signature

            st.session_state[
                "selling_point_analysis"
            ] = None

            st.session_state[
                "selling_point_options"
            ] = []

            st.session_state[
                "final_selling_points"
            ] = ""

            st.session_state[
                "selected_selling_point_version"
            ] = ""

            st.session_state.pop(
                "generated_storyboard",
                None,
            )

        st.caption(
            f"{size_mb:.2f} MB"
        )

    if st.button(
        "AI解析爆款视频",
        type="primary",
        use_container_width=True,
        disabled=(
            client is None
            or uploaded_video is None
        ),
    ):
        try:
            with st.spinner(
                "正在提取 Hook、转化逻辑和核心卖点…"
            ):
                result, metadata = (
                    analyze_video_selling_points(
                        client,
                        uploaded_video,
                        generate_category,
                        generate_product_name,
                    )
                )

            options = result.get(
                "options",
                [],
            )

            st.session_state[
                "selling_point_analysis"
            ] = result

            st.session_state[
                "selling_point_options"
            ] = options

            st.session_state[
                "selling_point_metadata"
            ] = metadata

            if options:
                st.session_state[
                    "selling_point_choice"
                ] = 0

                st.session_state[
                    "final_selling_points"
                ] = options[0].get(
                    "selling_points",
                    "",
                )

                st.session_state[
                    "selected_selling_point_version"
                ] = options[0].get(
                    "name",
                    "",
                )

            append_history(
                {
                    "record_type": "爆款解析",
                    "role": st.session_state["role"],
                    "operator": st.session_state["operator"],
                    "tiktok_account": generate_account,
                    "product_category": generate_category,
                    "product_name": generate_product_name,
                    "video_name": uploaded_video.name,
                    "video_size_mb": metadata["video_size_mb"],
                    "analysis_mode": metadata["analysis_mode"],
                    "analysis_seconds": metadata["analysis_seconds"],
                    "hook_summary": result.get(
                        "hook_summary",
                        "",
                    ),
                    "conversion_logic": result.get(
                        "conversion_logic",
                        "",
                    ),
                    "full_output_json": json_dumps(
                        result
                    ),
                }
            )

            st.success(
                "解析完成"
            )

        except Exception as exc:
            st.error(
                f"解析失败：{exc}"
            )

    # --------------------------------------------------------
    # 分析结果
    # --------------------------------------------------------

    selling_analysis = st.session_state.get(
        "selling_point_analysis"
    )

    selling_options = st.session_state.get(
        "selling_point_options",
        [],
    )

    if selling_analysis:

        with st.expander(
            "查看爆款拆解详情",
            expanded=False,
        ):
            st.markdown(
                "**用户真正被什么打动**"
            )

            st.write(
                selling_analysis.get(
                    "video_product_insight",
                    "",
                )
            )

            d1, d2 = st.columns(2)

            with d1:
                st.markdown(
                    "**前3秒 Hook**"
                )

                st.write(
                    selling_analysis.get(
                        "hook_summary",
                        "",
                    )
                )

            with d2:
                st.markdown(
                    "**转化逻辑**"
                )

                st.write(
                    selling_analysis.get(
                        "conversion_logic",
                        "",
                    )
                )

            metadata = st.session_state.get(
                "selling_point_metadata",
                {},
            )

            st.caption(
                "解析耗时："
                f'{metadata.get("analysis_seconds", "-")} 秒'
            )

        # ----------------------------------------------------
        # 卖点选择
        # ----------------------------------------------------

        st.markdown("### ③ 核心卖点")

        if selling_options:

            def option_label(index):
                option = selling_options[
                    index
                ]

                return (
                    f'{option.get("name", f"方案{index + 1}")}'
                    f'｜{option.get("angle", "")}'
                )

            st.radio(
                "卖点方向",
                options=list(
                    range(
                        len(selling_options)
                    )
                ),
                format_func=option_label,
                key="selling_point_choice",
                on_change=sync_selling_point_choice,
                label_visibility="collapsed",
            )

            selected_option = selling_options[
                st.session_state.get(
                    "selling_point_choice",
                    0,
                )
            ]

            st.caption(
                selected_option.get(
                    "reason",
                    "",
                )
            )

        final_selling_points = st.text_area(
            "最终核心卖点（可编辑）",
            height=105,
            key="final_selling_points",
        )

        # ----------------------------------------------------
        # 场景
        # ----------------------------------------------------

        st.markdown("### ④ 拍摄场景")

        generate_scene = st.selectbox(
            "拍摄场景",
            list(
                SCENE_LIBRARY.keys()
            ),
            key="generate_scene",
            label_visibility="collapsed",
        )

        with st.expander(
            "查看拍摄建议",
            expanded=False,
        ):
            st.write(
                SCENE_LIBRARY[
                    generate_scene
                ]["guide"]
            )

        # ----------------------------------------------------
        # 脚本
        # ----------------------------------------------------

        if st.button(
            "生成拍摄脚本",
            type="primary",
            use_container_width=True,
            disabled=(
                client is None
                or not final_selling_points.strip()
            ),
        ):
            try:
                with st.spinner(
                    "正在生成脚本…"
                ):
                    result, seconds = (
                        generate_storyboard(
                            client,
                            generate_category,
                            generate_product_name,
                            final_selling_points.strip(),
                            generate_scene,
                            selling_analysis,
                        )
                    )

                dataframe = storyboard_to_dataframe(
                    result.get(
                        "storyboard",
                        [],
                    )
                )

                markdown = dataframe_to_markdown(
                    dataframe
                )

                st.session_state[
                    "generated_storyboard"
                ] = result

                st.session_state[
                    "generated_storyboard_dataframe"
                ] = dataframe

                st.session_state[
                    "generated_storyboard_seconds"
                ] = seconds

                append_history(
                    {
                        "record_type": "脚本生成",
                        "role": st.session_state["role"],
                        "operator": st.session_state["operator"],
                        "tiktok_account": generate_account,
                        "product_category": generate_category,
                        "product_name": generate_product_name,
                        "selling_points": final_selling_points,
                        "selling_point_version": st.session_state.get(
                            "selected_selling_point_version",
                            "",
                        ),
                        "scene": generate_scene,
                        "analysis_mode": "脚本生成",
                        "analysis_seconds": seconds,
                        "hook_summary": selling_analysis.get(
                            "hook_summary",
                            "",
                        ),
                        "conversion_logic": selling_analysis.get(
                            "conversion_logic",
                            "",
                        ),
                        "full_output_json": json_dumps(
                            result
                        ),
                        "script_markdown": markdown,
                    }
                )

                st.success(
                    "脚本已生成"
                )

            except Exception as exc:
                st.error(
                    f"生成失败：{exc}"
                )

    if st.session_state.get(
        "generated_storyboard"
    ):
        dataframe = st.session_state[
            "generated_storyboard_dataframe"
        ]

        st.divider()

        st.markdown("### 拍摄脚本")

        st.markdown(
            dataframe_to_markdown(
                dataframe
            )
        )

        st.download_button(
            "下载 Excel",
            data=dataframe_to_excel_bytes(
                dataframe,
                "新脚本",
            ),
            file_name=(
                "TikTok新脚本_"
                + datetime.now().strftime(
                    "%Y%m%d_%H%M"
                )
                + ".xlsx"
            ),
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )


# ============================================================
# TAB 2：数据复盘
# ============================================================

with tab_review:

    st.markdown("### ① 原始脚本")

    uploaded_excel = st.file_uploader(
        "上传脚本 Excel",
        type=["xlsx"],
        key="review_excel",
    )

    if uploaded_excel is not None:
        signature = (
            uploaded_excel.name,
            uploaded_excel.size,
        )

        if (
            st.session_state.get(
                "last_excel_signature"
            )
            != signature
        ):
            try:
                parsed = excel_script_to_text(
                    uploaded_excel
                )

                st.session_state[
                    "review_original_script"
                ] = parsed

                st.session_state[
                    "last_excel_signature"
                ] = signature

                st.success(
                    "Excel读取成功"
                )

            except Exception as exc:
                st.error(
                    f"Excel解析失败：{exc}"
                )

    review_account = st.text_input(
        "TikTok账号",
        key="review_account",
    )

    meta1, meta2, meta3 = st.columns(3)

    with meta1:
        review_category = st.selectbox(
            "产品品类",
            PRODUCT_CATEGORIES,
            key="review_category",
        )

    with meta2:
        review_product_name = st.text_input(
            "产品名称 / SKU",
            key="review_product_name",
        )

    with meta3:
        review_traffic_type = st.selectbox(
            "流量类型",
            TRAFFIC_TYPES,
            key="review_traffic",
        )

    review_selling_points = st.text_area(
        "产品核心卖点",
        height=85,
        key="review_selling_points",
    )

    original_script = st.text_area(
        "原始脚本（可编辑）",
        height=250,
        key="review_original_script",
    )

    # --------------------------------------------------------
    # 核心指标
    # --------------------------------------------------------

    st.markdown("### ② 核心数据")

    m1, m2, m3 = st.columns(3)

    with m1:
        retention = optional_number(
            "前3秒留存率 (%)",
            "retention",
            100.0,
        )

        completion = optional_number(
            "平均完播率 (%)",
            "completion",
            100.0,
        )

    with m2:
        ctr = optional_number(
            "商品CTR (%)",
            "ctr",
            100.0,
            0.01,
        )

        conversion = optional_number(
            "订单转化率 (%)",
            "conversion",
            100.0,
            0.01,
        )

    with m3:
        engagement = optional_number(
            "互动率 (%)",
            "engagement",
            100.0,
            0.01,
        )

    # --------------------------------------------------------
    # 广告
    # --------------------------------------------------------

    with st.expander(
        "广告数据（选填）",
        expanded=False,
    ):
        a1, a2 = st.columns(2)

        with a1:
            actual_roas = optional_number(
                "实际 ROAS",
                "actual_roas",
            )

            target_roas = optional_number(
                "目标 ROAS",
                "target_roas",
            )

        with a2:
            actual_cpc = optional_number(
                "实际 CPC ($)",
                "actual_cpc",
                step=0.01,
            )

            target_cpc = optional_number(
                "目标 CPC ($)",
                "target_cpc",
                step=0.01,
            )

    # --------------------------------------------------------
    # 账号基准
    # --------------------------------------------------------

    with st.expander(
        "账号近7天基准（建议填写）",
        expanded=False,
    ):
        b1, b2, b3 = st.columns(3)

        with b1:
            base_retention = optional_number(
                "平均3秒留存 (%)",
                "base_retention",
                100.0,
            )

            base_completion = optional_number(
                "平均完播率 (%)",
                "base_completion",
                100.0,
            )

        with b2:
            base_ctr = optional_number(
                "平均商品CTR (%)",
                "base_ctr",
                100.0,
                0.01,
            )

            base_conversion = optional_number(
                "平均转化率 (%)",
                "base_conversion",
                100.0,
                0.01,
            )

        with b3:
            base_engagement = optional_number(
                "平均互动率 (%)",
                "base_engagement",
                100.0,
                0.01,
            )

    # --------------------------------------------------------
    # 下一版场景
    # --------------------------------------------------------

    st.markdown("### ③ 下一版场景")

    review_scene = st.selectbox(
        "下一版拍摄场景",
        list(
            SCENE_LIBRARY.keys()
        ),
        key="review_scene",
        label_visibility="collapsed",
    )

    metrics = compact_dict(
        {
            "retention_3s_pct": retention,
            "completion_rate_pct": completion,
            "product_ctr_pct": ctr,
            "order_conversion_pct": conversion,
            "engagement_rate_pct": engagement,
            "actual_roas": actual_roas,
            "target_roas": target_roas,
            "actual_cpc": actual_cpc,
            "target_cpc": target_cpc,
        }
    )

    baseline = compact_dict(
        {
            "retention_3s_pct": base_retention,
            "completion_rate_pct": base_completion,
            "product_ctr_pct": base_ctr,
            "order_conversion_pct": base_conversion,
            "engagement_rate_pct": base_engagement,
        }
    )

    assessment = build_metric_assessment(
        metrics,
        baseline,
    )

    if st.button(
        "开始复盘",
        type="primary",
        use_container_width=True,
        disabled=(
            client is None
        ),
    ):
        if not original_script.strip():
            st.error(
                "请上传或填写原始脚本。"
            )

        elif not review_selling_points.strip():
            st.error(
                "请填写产品核心卖点。"
            )

        elif not metrics:
            st.error(
                "请至少填写一项核心数据。"
            )

        else:
            try:
                with st.spinner(
                    "正在诊断漏斗…"
                ):
                    result, seconds = review_script(
                        client,
                        review_category,
                        review_product_name,
                        review_selling_points,
                        review_scene,
                        review_traffic_type,
                        original_script,
                        metrics,
                        baseline,
                        assessment,
                    )

                dataframe = storyboard_to_dataframe(
                    result.get(
                        "storyboard",
                        [],
                    )
                )

                st.session_state[
                    "review_result"
                ] = result

                st.session_state[
                    "review_dataframe"
                ] = dataframe

                st.session_state[
                    "review_seconds"
                ] = seconds

                append_history(
                    {
                        "record_type": "数据复盘",
                        "role": st.session_state["role"],
                        "operator": st.session_state["operator"],
                        "tiktok_account": review_account,
                        "product_category": review_category,
                        "product_name": review_product_name,
                        "selling_points": review_selling_points,
                        "scene": review_scene,
                        "traffic_type": review_traffic_type,
                        "analysis_seconds": seconds,
                        "priority_issue": result.get(
                            "priority_issue",
                            "",
                        ),
                        "diagnosis_summary": result.get(
                            "diagnosis_summary",
                            "",
                        ),
                        "account_diagnosis": result.get(
                            "account_diagnosis",
                            "",
                        ),
                        "metrics_json": json_dumps(
                            metrics
                        ),
                        "account_baseline_json": json_dumps(
                            baseline
                        ),
                        "metric_assessment_json": json_dumps(
                            assessment
                        ),
                        "full_output_json": json_dumps(
                            result
                        ),
                        "script_markdown": dataframe_to_markdown(
                            dataframe
                        ),
                        "original_script": original_script,
                    }
                )

                st.success(
                    "复盘完成"
                )

            except Exception as exc:
                st.error(
                    f"复盘失败：{exc}"
                )

    if st.session_state.get(
        "review_result"
    ):
        result = st.session_state[
            "review_result"
        ]

        dataframe = st.session_state[
            "review_dataframe"
        ]

        st.divider()

        st.markdown("### 复盘结论")

        st.info(
            result.get(
                "diagnosis_summary",
                "",
            )
        )

        st.markdown(
            f'**优先修复：** '
            f'{result.get("priority_issue", "-")}'
        )

        with st.expander(
            "查看逐项指标诊断",
            expanded=False,
        ):
            diagnosis = result.get(
                "metric_diagnosis",
                [],
            )

            if diagnosis:
                diagnosis_df = pd.DataFrame(
                    [
                        {
                            "指标": item.get(
                                "metric",
                                "",
                            ),
                            "状态": item.get(
                                "status",
                                "",
                            ),
                            "问题": item.get(
                                "meaning",
                                "",
                            ),
                            "调整": item.get(
                                "action",
                                "",
                            ),
                        }
                        for item in diagnosis
                    ]
                )

                st.dataframe(
                    diagnosis_df,
                    hide_index=True,
                    use_container_width=True,
                )

        st.markdown("### 优化脚本")

        st.markdown(
            dataframe_to_markdown(
                dataframe
            )
        )

        st.download_button(
            "下载优化版 Excel",
            data=dataframe_to_excel_bytes(
                dataframe,
                "优化版脚本",
            ),
            file_name=(
                "TikTok优化版_"
                + datetime.now().strftime(
                    "%Y%m%d_%H%M"
                )
                + ".xlsx"
            ),
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )


# ============================================================
# TAB 3：历史
# ============================================================

with tab_history:

    history = scoped_history()

    if history.empty:
        st.info(
            "暂无历史记录。"
        )

    else:
        filtered = history.copy()

        f1, f2, f3 = st.columns(3)

        with f1:
            if st.session_state["role"] == "主账号(Admin)":
                operators = ["全部"] + sorted(
                    [
                        x
                        for x in filtered["operator"].unique()
                        if x
                    ]
                )

                selected_operator = st.selectbox(
                    "操作人",
                    operators,
                    key="history_operator",
                )

                if selected_operator != "全部":
                    filtered = filtered[
                        filtered["operator"]
                        == selected_operator
                    ]

        with f2:
            accounts = ["全部"] + sorted(
                [
                    x
                    for x in filtered["tiktok_account"].unique()
                    if x
                ]
            )

            selected_account = st.selectbox(
                "TikTok账号",
                accounts,
                key="history_account",
            )

            if selected_account != "全部":
                filtered = filtered[
                    filtered["tiktok_account"]
                    == selected_account
                ]

        with f3:
            record_types = ["全部"] + sorted(
                [
                    x
                    for x in filtered["record_type"].unique()
                    if x
                ]
            )

            selected_type = st.selectbox(
                "类型",
                record_types,
                key="history_type",
            )

            if selected_type != "全部":
                filtered = filtered[
                    filtered["record_type"]
                    == selected_type
                ]

        display_columns = [
            "created_at_cn",
            "record_type",
            "operator",
            "tiktok_account",
            "product_name",
            "selling_point_version",
            "scene",
            "priority_issue",
        ]

        st.dataframe(
            filtered[
                display_columns
            ].iloc[::-1],
            hide_index=True,
            use_container_width=True,
        )

        csv_bytes = filtered.to_csv(
            index=False,
            encoding="utf-8-sig",
        ).encode(
            "utf-8-sig"
        )

        st.download_button(
            "下载历史 CSV",
            data=csv_bytes,
            file_name=(
                "TikTok历史_"
                + datetime.now().strftime(
                    "%Y%m%d_%H%M"
                )
                + ".csv"
            ),
            mime="text/csv",
            use_container_width=True,
        )

        record_ids = (
            filtered
            .iloc[::-1][
                "record_id"
            ]
            .tolist()
        )

        if record_ids:
            selected_id = st.selectbox(
                "查看完整记录",
                record_ids,
                key="history_detail",
            )

            row = filtered[
                filtered["record_id"]
                == selected_id
            ].iloc[0]

            if row["selling_points"]:
                st.markdown(
                    "**核心卖点**"
                )
                st.write(
                    row["selling_points"]
                )

            if row["diagnosis_summary"]:
                st.markdown(
                    "**复盘结论**"
                )
                st.write(
                    row["diagnosis_summary"]
                )

            if row["script_markdown"]:
                with st.expander(
                    "查看脚本",
                    expanded=False,
                ):
                    st.markdown(
                        row["script_markdown"]
                    )
