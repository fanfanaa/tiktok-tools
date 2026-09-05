import hashlib
import io
import json
import os
import random
import tempfile
import threading
import time
import uuid
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
# 1. APP CONFIG
# ============================================================

APP_TITLE = "TikTok爆款视频解析&复盘专用"

PRIMARY_MODEL = "gemini-3.5-flash-lite"

FALLBACK_MODELS = [
    "gemini-3.1-flash-lite",
    "gemini-3.6-flash",
]

MODEL_CHAIN = [
    PRIMARY_MODEL,
    *FALLBACK_MODELS,
]

MAX_ATTEMPTS_PER_MODEL = 2

# 最多一次比较 5 条视频
MAX_COMPARE_VIDEOS = 5

# 多视频总大小 <= 18MB 时优先 Inline
INLINE_BATCH_MAX_MB = 18

FILE_POLL_INTERVAL_SEC = 2
FILE_PROCESS_TIMEOUT_SEC = 180

DEFAULT_STAFF_PASSWORD = "8888"
DEFAULT_ADMIN_PASSWORD = "8888-admin"

HISTORY_FILE = Path("history_log.csv")
HISTORY_LOCK = threading.Lock()


# ============================================================
# 2. 民宿小场景库
# ============================================================

SCENE_LIBRARY = {
    "客厅·茶几快递拆包区": (
        "茶几上放快递盒、信封、产品。"
        "适合隐私印章、剪刀、小工具、开箱类产品。"
    ),

    "客厅·沙发边桌": (
        "坐在沙发旁，只拍手和桌面。"
        "适合耳机、便携用品、家居小工具。"
    ),

    "客厅·电视柜/玄关柜": (
        "站立第一视角完成拿取、使用、收纳。"
        "适合日用品、隐私用品、收纳产品。"
    ),

    "厨房·岛台正面": (
        "岛台作为主要操作区，第一视角手部连续操作。"
        "适合厨房工具、垃圾桶、清洁用品。"
    ),

    "厨房·切菜区": (
        "切菜、清理厨余、产品介入形成明显前后反差。"
        "适合厨房垃圾桶和效率型工具。"
    ),

    "厨房·水槽旁": (
        "利用水槽、抹布、清洁动作形成真实生活场景。"
        "适合清洁、收纳和厨房用品。"
    ),

    "卧室·床头柜": (
        "睡前或起床后的第一视角使用。"
        "适合耳机、阅读用品、便携用品。"
    ),

    "卧室·梳妆台": (
        "利用镜前桌面、抽屉、随手拿取动作。"
        "适合护理、收纳、小工具。"
    ),

    "卧室·床面": (
        "俯拍床面或第一视角展示产品。"
        "适合开箱、耳机、便携产品。"
    ),

    "卫生间·洗手台": (
        "镜前但不露脸，以手部、台面、产品为主。"
        "适合个人护理、清洁用品。"
    ),

    "卫生间·镜柜": (
        "第一视角开镜柜、取产品、使用、放回。"
        "适合收纳和护理类产品。"
    ),

    "阳台·落地窗边桌": (
        "自然光环境，适合展示产品材质、便携性和生活方式。"
    ),

    "纯桌面·白桌": (
        "完全不露脸，只使用产品、双手和必要道具。"
        "最适合功能型产品和高频脚本测试。"
    ),

    "纯桌面·快递箱/文件场景": (
        "桌面放快递标签、信封、文件、账单。"
        "特别适合隐私保护印章。"
    ),
}


# ============================================================
# 3. 产品与流量
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
# 4. Excel Columns
# ============================================================

VIDEO_ANALYSIS_COLUMNS = [
    "视频编号",
    "文件名",
    "一句话核心",
    "爆款脚本路线",
    "人群画像",
    "年龄预估",
    "前3秒Hook",
    "画面与节奏",
    "最值得吸收的3点",
]


DIRECTION_EXCEL_COLUMNS = [
    "方向",
    "核心思路",
    "目标人群",
    "小场景",
    "不露脸设计",
    "第一视角UGC",
    "强手部动作设计",
    "分镜序号",
    "时间段",
    "景别/机位",
    "画面描述(道具/动作)",
    "手部动作",
    "英文口播/字幕",
    "音效/节奏提示",
    "可吸收点",
    "差异化点",
    "设计目的(底层逻辑)",
]


OLD_EXCEL_COLUMNS = [
    "分镜序号",
    "景别/机位",
    "画面描述(道具/动作)",
    "英文口播文案/字幕",
    "音效/节奏提示",
    "设计目的(底层逻辑)",
]


# ============================================================
# 5. History
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
    "video_names",
    "video_count",
    "model_used",
    "fallback_used",
    "retry_count",
    "analysis_seconds",
    "priority_issue",
    "diagnosis_summary",
    "metrics_json",
    "account_baseline_json",
    "full_output_json",
]


# ============================================================
# 6. 内部复盘区间
# 非 TikTok 官方 Benchmark
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
# 7. 多视频爆款解析 Schema
# ============================================================

VIDEO_ANALYSIS_SCHEMA = {
    "type": "object",

    "properties": {
        "comparison_summary": {
            "type": "object",

            "properties": {
                "one_sentence_core": {
                    "type": "string",
                },

                "common_script_route": {
                    "type": "string",
                },

                "common_audience": {
                    "type": "string",
                },

                "age_estimate": {
                    "type": "string",
                },

                "common_hook_pattern": {
                    "type": "string",
                },

                "visual_rhythm": {
                    "type": "string",
                },

                "top_absorb_points": {
                    "type": "array",
                    "minItems": 3,
                    "maxItems": 3,

                    "items": {
                        "type": "string",
                    },
                },

                "key_differences": {
                    "type": "string",
                },
            },

            "required": [
                "one_sentence_core",
                "common_script_route",
                "common_audience",
                "age_estimate",
                "common_hook_pattern",
                "visual_rhythm",
                "top_absorb_points",
                "key_differences",
            ],
        },

        "videos": {
            "type": "array",
            "minItems": 1,
            "maxItems": MAX_COMPARE_VIDEOS,

            "items": {
                "type": "object",

                "properties": {
                    "video_index": {
                        "type": "integer",
                    },

                    "filename": {
                        "type": "string",
                    },

                    "one_sentence_core": {
                        "type": "string",
                    },

                    "script_route": {
                        "type": "string",
                    },

                    "audience_profile": {
                        "type": "string",
                    },

                    "age_estimate": {
                        "type": "string",
                    },

                    "first_3s_hook": {
                        "type": "string",
                    },

                    "visual_rhythm": {
                        "type": "string",
                    },

                    "top_absorb_points": {
                        "type": "array",
                        "minItems": 3,
                        "maxItems": 3,

                        "items": {
                            "type": "string",
                        },
                    },
                },

                "required": [
                    "video_index",
                    "filename",
                    "one_sentence_core",
                    "script_route",
                    "audience_profile",
                    "age_estimate",
                    "first_3s_hook",
                    "visual_rhythm",
                    "top_absorb_points",
                ],
            },
        },
    },

    "required": [
        "comparison_summary",
        "videos",
    ],
}


# ============================================================
# 8. 3个拍摄方向 Schema
# ============================================================

DIRECTIONS_SCHEMA = {
    "type": "object",

    "properties": {
        "directions": {
            "type": "array",
            "minItems": 3,
            "maxItems": 3,

            "items": {
                "type": "object",

                "properties": {
                    "direction_name": {
                        "type": "string",
                    },

                    "core_idea": {
                        "type": "string",
                    },

                    "target_audience": {
                        "type": "string",
                    },

                    "estimated_duration": {
                        "type": "string",
                    },

                    "small_scene": {
                        "type": "string",
                    },

                    "faceless_design": {
                        "type": "string",
                    },

                    "pov_ugc": {
                        "type": "string",
                    },

                    "hand_action_design": {
                        "type": "string",
                    },

                    "absorb_points": {
                        "type": "array",
                        "minItems": 2,
                        "maxItems": 4,

                        "items": {
                            "type": "string",
                        },
                    },

                    "differentiation_points": {
                        "type": "array",
                        "minItems": 2,
                        "maxItems": 4,

                        "items": {
                            "type": "string",
                        },
                    },

                    "storyboard": {
                        "type": "array",
                        "minItems": 5,
                        "maxItems": 12,

                        "items": {
                            "type": "object",

                            "properties": {
                                "sequence": {
                                    "type": "string",
                                },

                                "time_range": {
                                    "type": "string",
                                },

                                "shot": {
                                    "type": "string",
                                },

                                "visual": {
                                    "type": "string",
                                },

                                "hand_action": {
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
                                "time_range",
                                "shot",
                                "visual",
                                "hand_action",
                                "copy_en",
                                "audio",
                                "rationale",
                            ],
                        },
                    },
                },

                "required": [
                    "direction_name",
                    "core_idea",
                    "target_audience",
                    "estimated_duration",
                    "small_scene",
                    "faceless_design",
                    "pov_ugc",
                    "hand_action_design",
                    "absorb_points",
                    "differentiation_points",
                    "storyboard",
                ],
            },
        },
    },

    "required": [
        "directions",
    ],
}


# ============================================================
# 9. 数据复盘 Schema
# ============================================================

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

        "optimized_script": {
            "type": "array",
            "minItems": 5,
            "maxItems": 12,

            "items": {
                "type": "object",

                "properties": {
                    "sequence": {
                        "type": "string",
                    },

                    "time_range": {
                        "type": "string",
                    },

                    "shot": {
                        "type": "string",
                    },

                    "visual": {
                        "type": "string",
                    },

                    "hand_action": {
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
                    "time_range",
                    "shot",
                    "visual",
                    "hand_action",
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
        "optimized_script",
    ],
}


# ============================================================
# 10. 页面
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
        padding-top: 1.2rem;
        padding-bottom: 3rem;
    }

    h1 {
        font-size: 1.8rem !important;
        margin-bottom: .8rem !important;
    }

    h2 {
        font-size: 1.35rem !important;
    }

    h3 {
        font-size: 1.10rem !important;
        margin-top: 1.1rem !important;
        margin-bottom: .6rem !important;
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

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 11. Secrets
# ============================================================

def get_secret(
    name,
    default="",
):
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


def create_client(
    api_key,
):
    return genai.Client(
        api_key=api_key
    )


# ============================================================
# 12. Session
# ============================================================

def init_session():
    defaults = {
        "authenticated": False,
        "role": "",
        "operator": "",

        "video_analysis": None,
        "video_analysis_meta": {},

        "directions_result": None,
        "directions_meta": {},

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


init_session()


# ============================================================
# 13. Helpers
# ============================================================

def clean_text(
    value,
):
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""

    except Exception:
        pass

    return str(value).strip()


def json_dumps(
    data,
):
    return json.dumps(
        data,
        ensure_ascii=False,
        separators=(",", ":"),
    )


def compact_dict(
    data,
):
    return {
        key: value
        for key, value in data.items()
        if value is not None
        and value != ""
    }


def parse_json_output(
    raw_text,
):
    if not raw_text:
        raise ValueError(
            "AI没有返回有效结果。"
        )

    try:
        return json.loads(
            raw_text
        )

    except json.JSONDecodeError as exc:
        raise ValueError(
            "AI返回格式异常，请重新执行。"
        ) from exc


def thinking_config():
    return types.ThinkingConfig(
        thinking_level="minimal"
    )


def video_batch_signature(
    uploaded_videos,
    category,
    product_name,
):
    hasher = hashlib.sha256()

    for video in uploaded_videos or []:
        hasher.update(
            video.getvalue()
        )

    hasher.update(
        clean_text(
            category
        ).encode("utf-8")
    )

    hasher.update(
        clean_text(
            product_name
        ).encode("utf-8")
    )

    return hasher.hexdigest()[:24]


# ============================================================
# 14. 429 / 503 自动重试
# ============================================================

TRANSIENT_MARKERS = [
    "429",
    "500",
    "502",
    "503",
    "504",
    "RESOURCE_EXHAUSTED",
    "TOO_MANY_REQUESTS",
    "RATE_LIMIT",
    "INTERNAL",
    "UNAVAILABLE",
    "SERVICE_UNAVAILABLE",
    "DEADLINE_EXCEEDED",
    "HIGH DEMAND",
    "TEMPORARILY",
    "OVERLOADED",
]


MODEL_ERROR_MARKERS = [
    "MODEL_NOT_FOUND",
    "MODEL NOT FOUND",
    "NOT_FOUND",
]


def is_transient_error(
    exc,
):
    text = str(
        exc
    ).upper()

    return any(
        marker in text
        for marker in TRANSIENT_MARKERS
    )


def is_model_error(
    exc,
):
    text = str(
        exc
    ).upper()

    return any(
        marker in text
        for marker in MODEL_ERROR_MARKERS
    )


def friendly_error(
    exc,
):
    text = str(
        exc
    ).upper()

    if (
        "401" in text
        or "UNAUTHENTICATED" in text
    ):
        return (
            "AI密钥认证失败，请联系管理员检查系统配置。"
        )

    if (
        "400" in text
        or "INVALID_ARGUMENT" in text
    ):
        return (
            "AI请求参数异常，请联系管理员检查程序配置。"
        )

    if (
        "429" in text
        or "RESOURCE_EXHAUSTED" in text
    ):
        return (
            "当前AI请求较多，系统已经自动重试并尝试备用线路，请稍后再试。"
        )

    if (
        "503" in text
        or "UNAVAILABLE" in text
        or "HIGH DEMAND" in text
    ):
        return (
            "当前AI服务繁忙，系统已自动重试并尝试备用线路，请稍后再试。"
        )

    return (
        "AI暂时未完成本次任务，请稍后重新执行。"
    )


def generate_resilient(
    client,
    contents,
    config,
):
    last_exception = None
    total_attempts = 0

    for model_index, model_name in enumerate(
        MODEL_CHAIN
    ):
        for attempt_index in range(
            MAX_ATTEMPTS_PER_MODEL
        ):
            total_attempts += 1

            try:
                response = (
                    client.models.generate_content(
                        model=model_name,
                        contents=contents,
                        config=config,
                    )
                )

                meta = {
                    "model_used": model_name,

                    "fallback_used": (
                        model_index > 0
                    ),

                    "retry_count": max(
                        total_attempts - 1,
                        0,
                    ),
                }

                return (
                    response,
                    meta,
                )

            except Exception as exc:
                last_exception = exc

                if is_model_error(
                    exc
                ):
                    break

                if not is_transient_error(
                    exc
                ):
                    raise

                if (
                    attempt_index
                    < MAX_ATTEMPTS_PER_MODEL - 1
                ):
                    delay = (
                        1.3
                        * (
                            2 ** attempt_index
                        )
                        + random.uniform(
                            0.2,
                            0.8,
                        )
                    )

                    time.sleep(
                        delay
                    )

                    continue

                break

    raise RuntimeError(
        friendly_error(
            last_exception
        )
    ) from last_exception


# ============================================================
# 15. History
# ============================================================

def empty_history():
    return pd.DataFrame(
        columns=HISTORY_COLUMNS
    )


def normalize_history(
    dataframe,
):
    if dataframe is None:
        return empty_history()

    dataframe = dataframe.copy()

    for column in HISTORY_COLUMNS:
        if column not in dataframe.columns:
            dataframe[column] = ""

    return dataframe.reindex(
        columns=HISTORY_COLUMNS
    )


def load_history():
    if not HISTORY_FILE.exists():
        return empty_history()

    try:
        dataframe = pd.read_csv(
            HISTORY_FILE,
            dtype=str,
            keep_default_na=False,
            encoding="utf-8-sig",
        )

        return normalize_history(
            dataframe
        )

    except Exception:
        return empty_history()


def write_history(
    dataframe,
):
    dataframe = normalize_history(
        dataframe
    )

    HISTORY_FILE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temp_file = HISTORY_FILE.with_name(
        "history_log.tmp.csv"
    )

    dataframe.to_csv(
        temp_file,
        index=False,
        encoding="utf-8-sig",
    )

    os.replace(
        temp_file,
        HISTORY_FILE,
    )


def append_history(
    record,
):
    row = {
        column: clean_text(
            record.get(
                column,
                "",
            )
        )
        for column in HISTORY_COLUMNS
    }

    if not row["record_id"]:
        row["record_id"] = (
            uuid.uuid4().hex[:12]
        )

    now = datetime.now(
        timezone.utc
    )

    row["created_at_utc"] = (
        now.isoformat(
            timespec="seconds"
        )
    )

    row["created_at_cn"] = (
        now
        .astimezone(
            ZoneInfo(
                "Asia/Shanghai"
            )
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
                pd.DataFrame(
                    [row]
                ),
            ],
            ignore_index=True,
        )

        write_history(
            updated
        )


def scoped_history():
    dataframe = load_history()

    if (
        st.session_state["role"]
        == "主账号(Admin)"
    ):
        return dataframe

    return dataframe[
        dataframe["operator"]
        == st.session_state["operator"]
    ].copy()


# ============================================================
# 16. Login
# ============================================================

def render_login():
    staff_password = get_secret(
        "STAFF_PASSWORD",
        DEFAULT_STAFF_PASSWORD,
    )

    admin_password = get_secret(
        "ADMIN_PASSWORD",
        DEFAULT_ADMIN_PASSWORD,
    )

    with st.sidebar:
        st.markdown(
            "### 团队登录"
        )

        if st.session_state[
            "authenticated"
        ]:
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
                st.session_state[
                    "authenticated"
                ] = True

                st.session_state[
                    "role"
                ] = role

                st.session_state[
                    "operator"
                ] = operator.strip()

                st.rerun()


# ============================================================
# 17. Sidebar History
# ============================================================

def render_sidebar_history():
    if not st.session_state[
        "authenticated"
    ]:
        return

    history = scoped_history()

    with st.sidebar:
        st.divider()

        st.markdown(
            "### 历史"
        )

        if history.empty:
            st.caption(
                "暂无记录"
            )
            return

        filtered = history.copy()

        if (
            st.session_state["role"]
            == "主账号(Admin)"
        ):
            operators = (
                ["全部"]
                + sorted(
                    [
                        x
                        for x
                        in filtered[
                            "operator"
                        ].unique()
                        if x
                    ]
                )
            )

            selected_operator = (
                st.selectbox(
                    "操作人",
                    operators,
                    key="sidebar_operator",
                )
            )

            if (
                selected_operator
                != "全部"
            ):
                filtered = filtered[
                    filtered["operator"]
                    == selected_operator
                ]

        st.caption(
            f"{len(filtered)} 条记录"
        )

        csv_bytes = (
            filtered.to_csv(
                index=False,
                encoding="utf-8-sig",
            )
            .encode(
                "utf-8-sig"
            )
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
# 18. Video File API
# ============================================================

def wait_until_active(
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
                "视频预处理超时。"
            )

        time.sleep(
            FILE_POLL_INTERVAL_SEC
        )

        current = client.files.get(
            name=current.name
        )


# ============================================================
# 19. 多视频解析 Prompt
# ============================================================

def build_video_analysis_prompt(
    category,
    product_name,
    filenames,
):
    file_lines = "\n".join(
        [
            f"视频{i + 1}：{name}"
            for i, name
            in enumerate(
                filenames
            )
        ]
    )

    return f"""
你是美国 TikTok Shop 爆款短视频拆解负责人。

本次上传：
{file_lines}

产品品类：
{category}

产品名称：
{product_name}

请逐条分析所有视频。

所有分析内容必须用中文。
英文字幕或口播原意可以保留必要英文片段。

每一条视频固定输出：

1. 一句话核心
用一句话说明这条视频真正为什么能吸引用户。

2. 爆款脚本路线
例如：
冲突Hook → 痛点放大 → 产品出现 → Demo → 结果证明 → CTA。

不要只写标签，要结合这条视频实际发生的内容。

3. 人群画像
根据视频画面、语言、产品使用场景、痛点和内容风格，
推测最可能被吸引的美国受众。

4. 年龄预估
给出核心年龄段和次级年龄段。
这属于内容推断，不是平台后台真实人口数据，
必须明确使用“预估”。

5. 前3秒Hook
拆解：
第一帧、第一动作、字幕/口播、声音、冲突或视觉信息。

6. 画面与节奏
分析镜头长度、切换节奏、手部动作、产品出现时间、
Demo密度和结果出现时机。

7. 最值得吸收的3点
只能吸收底层机制，
不能直接复制原句、品牌或独特剧情。

如果上传2条及以上视频：

还必须额外做横向对比：

- 一句话共同核心
- 共同爆款脚本路线
- 共同人群
- 年龄预估
- 共同Hook模式
- 共同画面与节奏
- 最值得共同吸收的3点
- 各条视频最关键的差异

不要根据我的产品卖点反向美化爆款视频。
这一阶段只负责客观拆解视频本身。

严格按照JSON Schema输出。
""".strip()


# ============================================================
# 20. 多视频一次解析
# ============================================================

def analyze_videos(
    client,
    uploaded_videos,
    category,
    product_name,
):
    if not uploaded_videos:
        raise ValueError(
            "请至少上传1条视频。"
        )

    if (
        len(uploaded_videos)
        > MAX_COMPARE_VIDEOS
    ):
        raise ValueError(
            f"单次最多上传 {MAX_COMPARE_VIDEOS} 条视频。"
        )

    started = (
        time.perf_counter()
    )

    total_bytes = sum(
        len(
            video.getvalue()
        )
        for video
        in uploaded_videos
    )

    total_mb = (
        total_bytes
        / 1024
        / 1024
    )

    filenames = [
        video.name
        for video
        in uploaded_videos
    ]

    prompt = (
        build_video_analysis_prompt(
            category,
            product_name,
            filenames,
        )
    )

    remote_files = []
    temp_paths = []

    try:
        parts = []

        # ----------------------------------------------------
        # 总体较小时，一次 Inline 分析全部视频
        # ----------------------------------------------------

        if (
            total_mb
            <= INLINE_BATCH_MAX_MB
        ):
            analysis_mode = (
                "多视频快速解析"
            )

            for index, video in enumerate(
                uploaded_videos,
                start=1,
            ):
                parts.append(
                    types.Part.from_text(
                        text=(
                            f"【视频{index}】"
                            f"文件名：{video.name}"
                        )
                    )
                )

                parts.append(
                    types.Part.from_bytes(
                        data=video.getvalue(),
                        mime_type=(
                            video.type
                            or "video/mp4"
                        ),
                    )
                )

        # ----------------------------------------------------
        # 总体较大时，全部走 Files API
        # ----------------------------------------------------

        else:
            analysis_mode = (
                "多视频大文件解析"
            )

            for index, video in enumerate(
                uploaded_videos,
                start=1,
            ):
                suffix = (
                    Path(
                        video.name
                    ).suffix
                    or ".mp4"
                )

                with (
                    tempfile.NamedTemporaryFile(
                        delete=False,
                        suffix=suffix,
                    )
                ) as temp:
                    temp.write(
                        video.getvalue()
                    )

                    temp_path = (
                        temp.name
                    )

                temp_paths.append(
                    temp_path
                )

                remote_file = (
                    client.files.upload(
                        file=temp_path
                    )
                )

                remote_file = (
                    wait_until_active(
                        client,
                        remote_file,
                    )
                )

                remote_files.append(
                    remote_file
                )

                parts.append(
                    types.Part.from_text(
                        text=(
                            f"【视频{index}】"
                            f"文件名：{video.name}"
                        )
                    )
                )

                parts.append(
                    types.Part.from_uri(
                        file_uri=(
                            remote_file.uri
                        ),
                        mime_type=(
                            remote_file.mime_type
                            or "video/mp4"
                        ),
                    )
                )

        parts.append(
            types.Part.from_text(
                text=prompt
            )
        )

        content = types.Content(
            role="user",
            parts=parts,
        )

        config = (
            types.GenerateContentConfig(
                thinking_config=(
                    thinking_config()
                ),

                max_output_tokens=5200,

                response_mime_type=(
                    "application/json"
                ),

                response_json_schema=(
                    VIDEO_ANALYSIS_SCHEMA
                ),
            )
        )

        response, call_meta = (
            generate_resilient(
                client=client,
                contents=content,
                config=config,
            )
        )

        result = parse_json_output(
            response.text
        )

        metadata = {
            "video_count":
                len(uploaded_videos),

            "total_size_mb":
                round(
                    total_mb,
                    2,
                ),

            "analysis_mode":
                analysis_mode,

            "analysis_seconds":
                round(
                    time.perf_counter()
                    - started,
                    1,
                ),

            **call_meta,
        }

        return (
            result,
            metadata,
        )

    finally:
        for remote_file in remote_files:
            try:
                client.files.delete(
                    name=remote_file.name
                )

            except Exception:
                pass

        for temp_path in temp_paths:
            try:
                if os.path.exists(
                    temp_path
                ):
                    os.remove(
                        temp_path
                    )

            except OSError:
                pass


# ============================================================
# 21. 3方向生成 Prompt
# ============================================================

def build_directions_prompt(
    category,
    product_name,
    product_selling_points,
    video_analysis,
    allowed_scenes,
):
    scene_text = "\n".join(
        [
            f"- {scene}：{SCENE_LIBRARY[scene]}"
            for scene
            in allowed_scenes
        ]
    )

    return f"""
你是美国 TikTok Shop 短视频拍摄SOP负责人。

【我们的产品】

品类：
{category}

产品：
{product_name}

真实产品卖点：
{product_selling_points}

【已经完成的爆款视频拆解】

{json.dumps(
    video_analysis,
    ensure_ascii=False,
    indent=2
)}

【可使用的民宿小场景】

{scene_text}

现在不是复制任何一条爆款视频。

你的任务是：

吸收多个爆款的共同底层逻辑，
结合我们自己的真实产品卖点，
重新生成3个明显不同的拍摄方向。

三个方向不能只是换一句文案。

必须在以下至少3项上形成明显差异：

- 前3秒Hook机制
- 小场景
- 手部动作
- 产品Demo顺序
- 用户痛点
- 叙事角度
- CTA路径

每个方向必须满足：

1. 小场景
必须从给出的民宿可用场景中选择一个具体小场景。
不能只写“厨房”“客厅”这种大场景。

2. 不露脸
不能设计正脸出镜。
允许：
手部、身体局部、背影极少量出现。

3. 第一视角 UGC
以真实用户本人操作产品的感觉进行拍摄，
不要生成传统广告大片。

4. 强手部动作
前3秒尤其需要有清晰动作：
拿、滚、撕、推、挂、按、打开、关闭、倒、切、戴等。

5. 逐秒拍摄脚本
每个分镜必须有明确时间段，例如：
0-1.5s
1.5-3s
3-5s

整条视频控制在15-40秒。

6. 英文字幕 / 口播
必须是自然美国TikTok表达。
短句、口语化，不像广告说明书。

7. 可吸收点
明确说明这个方向吸收了爆款视频的什么底层机制。

8. 差异化点
明确说明它与原爆款哪里不同，
避免同质化复制。

9. 产品真实性
不虚构：
功能、销量、认证、折扣、医疗、安全效果。

严格输出3个方向。
严格按照JSON Schema输出。
""".strip()


def generate_directions(
    client,
    category,
    product_name,
    product_selling_points,
    video_analysis,
    allowed_scenes,
):
    started = (
        time.perf_counter()
    )

    prompt = build_directions_prompt(
        category,
        product_name,
        product_selling_points,
        video_analysis,
        allowed_scenes,
    )

    config = (
        types.GenerateContentConfig(
            thinking_config=(
                thinking_config()
            ),

            max_output_tokens=7600,

            response_mime_type=(
                "application/json"
            ),

            response_json_schema=(
                DIRECTIONS_SCHEMA
            ),
        )
    )

    response, call_meta = (
        generate_resilient(
            client=client,
            contents=prompt,
            config=config,
        )
    )

    result = parse_json_output(
        response.text
    )

    metadata = {
        "analysis_seconds":
            round(
                time.perf_counter()
                - started,
                1,
            ),

        **call_meta,
    }

    return (
        result,
        metadata,
    )


# ============================================================
# 22. 分析 DataFrame
# ============================================================

def analysis_to_dataframe(
    result,
):
    rows = []

    for video in result.get(
        "videos",
        [],
    ):
        absorb = "\n".join(
            [
                f"{i + 1}. {item}"
                for i, item
                in enumerate(
                    video.get(
                        "top_absorb_points",
                        [],
                    )
                )
            ]
        )

        rows.append(
            {
                "视频编号":
                    video.get(
                        "video_index",
                        "",
                    ),

                "文件名":
                    video.get(
                        "filename",
                        "",
                    ),

                "一句话核心":
                    video.get(
                        "one_sentence_core",
                        "",
                    ),

                "爆款脚本路线":
                    video.get(
                        "script_route",
                        "",
                    ),

                "人群画像":
                    video.get(
                        "audience_profile",
                        "",
                    ),

                "年龄预估":
                    video.get(
                        "age_estimate",
                        "",
                    ),

                "前3秒Hook":
                    video.get(
                        "first_3s_hook",
                        "",
                    ),

                "画面与节奏":
                    video.get(
                        "visual_rhythm",
                        "",
                    ),

                "最值得吸收的3点":
                    absorb,
            }
        )

    return pd.DataFrame(
        rows,
        columns=VIDEO_ANALYSIS_COLUMNS,
    )


# ============================================================
# 23. 方向 DataFrame
# ============================================================

def direction_to_dataframe(
    direction,
):
    rows = []

    absorb = "；".join(
        direction.get(
            "absorb_points",
            [],
        )
    )

    difference = "；".join(
        direction.get(
            "differentiation_points",
            [],
        )
    )

    for shot in direction.get(
        "storyboard",
        [],
    ):
        rows.append(
            {
                "方向":
                    direction.get(
                        "direction_name",
                        "",
                    ),

                "核心思路":
                    direction.get(
                        "core_idea",
                        "",
                    ),

                "目标人群":
                    direction.get(
                        "target_audience",
                        "",
                    ),

                "小场景":
                    direction.get(
                        "small_scene",
                        "",
                    ),

                "不露脸设计":
                    direction.get(
                        "faceless_design",
                        "",
                    ),

                "第一视角UGC":
                    direction.get(
                        "pov_ugc",
                        "",
                    ),

                "强手部动作设计":
                    direction.get(
                        "hand_action_design",
                        "",
                    ),

                "分镜序号":
                    shot.get(
                        "sequence",
                        "",
                    ),

                "时间段":
                    shot.get(
                        "time_range",
                        "",
                    ),

                "景别/机位":
                    shot.get(
                        "shot",
                        "",
                    ),

                "画面描述(道具/动作)":
                    shot.get(
                        "visual",
                        "",
                    ),

                "手部动作":
                    shot.get(
                        "hand_action",
                        "",
                    ),

                "英文口播/字幕":
                    shot.get(
                        "copy_en",
                        "",
                    ),

                "音效/节奏提示":
                    shot.get(
                        "audio",
                        "",
                    ),

                "可吸收点":
                    absorb,

                "差异化点":
                    difference,

                "设计目的(底层逻辑)":
                    shot.get(
                        "rationale",
                        "",
                    ),
            }
        )

    return pd.DataFrame(
        rows,
        columns=DIRECTION_EXCEL_COLUMNS,
    )


# ============================================================
# 24. Excel Formatting
# ============================================================

def format_sheet(
    worksheet,
):
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

    for row in worksheet.iter_rows(
        min_row=2
    ):
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

    for index in range(
        1,
        worksheet.max_column + 1,
    ):
        worksheet.column_dimensions[
            get_column_letter(
                index
            )
        ].width = 22

    if worksheet.max_column >= 10:
        worksheet.column_dimensions[
            "K"
        ].width = 42

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = (
        worksheet.dimensions
    )


# ============================================================
# 25. 一键生成完整 Excel
# ============================================================

def build_full_excel(
    video_analysis,
    directions_result,
):
    output = io.BytesIO()

    summary = video_analysis.get(
        "comparison_summary",
        {},
    )

    summary_df = pd.DataFrame(
        [
            {
                "项目": "一句话共同核心",
                "内容": summary.get(
                    "one_sentence_core",
                    "",
                ),
            },

            {
                "项目": "共同爆款脚本路线",
                "内容": summary.get(
                    "common_script_route",
                    "",
                ),
            },

            {
                "项目": "共同人群",
                "内容": summary.get(
                    "common_audience",
                    "",
                ),
            },

            {
                "项目": "年龄预估",
                "内容": summary.get(
                    "age_estimate",
                    "",
                ),
            },

            {
                "项目": "共同前3秒Hook",
                "内容": summary.get(
                    "common_hook_pattern",
                    "",
                ),
            },

            {
                "项目": "共同画面与节奏",
                "内容": summary.get(
                    "visual_rhythm",
                    "",
                ),
            },

            {
                "项目": "最值得吸收的3点",
                "内容": "\n".join(
                    [
                        f"{i + 1}. {item}"
                        for i, item
                        in enumerate(
                            summary.get(
                                "top_absorb_points",
                                [],
                            )
                        )
                    ]
                ),
            },

            {
                "项目": "关键差异",
                "内容": summary.get(
                    "key_differences",
                    "",
                ),
            },
        ]
    )

    video_df = (
        analysis_to_dataframe(
            video_analysis
        )
    )

    all_direction_frames = []

    directions = (
        directions_result.get(
            "directions",
            [],
        )
    )

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        summary_df.to_excel(
            writer,
            index=False,
            sheet_name="爆款对比总结",
        )

        video_df.to_excel(
            writer,
            index=False,
            sheet_name="逐条视频拆解",
        )

        format_sheet(
            writer.book[
                "爆款对比总结"
            ]
        )

        format_sheet(
            writer.book[
                "逐条视频拆解"
            ]
        )

        for index, direction in enumerate(
            directions,
            start=1,
        ):
            direction_df = (
                direction_to_dataframe(
                    direction
                )
            )

            all_direction_frames.append(
                direction_df
            )

            sheet_name = (
                f"方向{index}"
            )

            direction_df.to_excel(
                writer,
                index=False,
                sheet_name=sheet_name,
            )

            format_sheet(
                writer.book[
                    sheet_name
                ]
            )

        if all_direction_frames:
            all_df = pd.concat(
                all_direction_frames,
                ignore_index=True,
            )

            all_df.to_excel(
                writer,
                index=False,
                sheet_name="拍摄执行总表",
            )

            format_sheet(
                writer.book[
                    "拍摄执行总表"
                ]
            )

    output.seek(0)

    return output.getvalue()


# ============================================================
# 26. Excel Script Import
# ============================================================

def get_script_sheets(
    uploaded_file,
):
    data = uploaded_file.getvalue()

    excel_file = pd.ExcelFile(
        io.BytesIO(data),
        engine="openpyxl",
    )

    result = []

    for sheet in excel_file.sheet_names:
        try:
            df = pd.read_excel(
                io.BytesIO(data),
                sheet_name=sheet,
                nrows=2,
                engine="openpyxl",
            )

            columns = set(
                df.columns.tolist()
            )

            new_match = {
                "分镜序号",
                "时间段",
                "画面描述(道具/动作)",
                "英文口播/字幕",
            }.issubset(
                columns
            )

            old_match = set(
                OLD_EXCEL_COLUMNS
            ).issubset(
                columns
            )

            if (
                new_match
                or old_match
            ):
                result.append(
                    sheet
                )

        except Exception:
            continue

    return result


def excel_sheet_to_script_text(
    uploaded_file,
    sheet_name,
):
    df = pd.read_excel(
        io.BytesIO(
            uploaded_file.getvalue()
        ),
        sheet_name=sheet_name,
        engine="openpyxl",
    )

    blocks = []

    for _, row in df.iterrows():

        if (
            "时间段"
            in df.columns
        ):
            block = f"""
分镜：{clean_text(row.get("分镜序号"))}
时间：{clean_text(row.get("时间段"))}
小场景：{clean_text(row.get("小场景"))}
景别/机位：{clean_text(row.get("景别/机位"))}
画面：{clean_text(row.get("画面描述(道具/动作)"))}
手部动作：{clean_text(row.get("手部动作"))}
英文口播/字幕：{clean_text(row.get("英文口播/字幕"))}
音效/节奏：{clean_text(row.get("音效/节奏提示"))}
设计目的：{clean_text(row.get("设计目的(底层逻辑)"))}
""".strip()

        else:
            block = f"""
分镜：{clean_text(row.get("分镜序号"))}
景别/机位：{clean_text(row.get("景别/机位"))}
画面：{clean_text(row.get("画面描述(道具/动作)"))}
英文口播/字幕：{clean_text(row.get("英文口播文案/字幕"))}
音效/节奏：{clean_text(row.get("音效/节奏提示"))}
设计目的：{clean_text(row.get("设计目的(底层逻辑)"))}
""".strip()

        blocks.append(
            block
        )

    return "\n\n".join(
        blocks
    )


# ============================================================
# 27. Review Metrics
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
            status = (
                "明显低于账号基准"
            )

        elif ratio > 1.2:
            status = (
                "明显高于账号基准"
            )

        else:
            status = (
                "接近账号基准"
            )

        return {
            "value": value,
            "baseline": baseline,
            "ratio": round(
                ratio,
                3,
            ),
            "status": status,
        }

    band = SOP_BANDS[
        key
    ]

    if value < band["low"]:
        status = "偏低"

    elif value >= band["high"]:
        status = "较强"

    else:
        status = "中间区间"

    return {
        "value": value,
        "status": status,
        "basis": "内部SOP工作区间",
    }


def build_metric_assessment(
    metrics,
    baseline,
):
    result = {
        "前3秒留存": compare_metric(
            "retention_3s_pct",
            metrics.get(
                "retention_3s_pct"
            ),
            baseline.get(
                "retention_3s_pct"
            ),
        ),

        "完播率": compare_metric(
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

    return result


# ============================================================
# 28. Review Prompt
# ============================================================

def build_review_prompt(
    category,
    product_name,
    selling_points,
    original_script,
    metrics,
    baseline,
    assessment,
    traffic_type,
):
    return f"""
你是美国 TikTok Shop 短视频数据复盘负责人。

产品：
{category}
{product_name}

真实产品卖点：
{selling_points}

流量类型：
{traffic_type}

原始拍摄脚本：
{original_script}

本条视频数据：
{json.dumps(metrics, ensure_ascii=False)}

账号近7天同类视频基准：
{json.dumps(baseline, ensure_ascii=False)}

系统初步判断：
{json.dumps(assessment, ensure_ascii=False)}

严格按照以下漏斗判断：

1. 前3秒留存率 → Hook问题
2. 平均完播率 → 中段节奏问题
3. 互动率 → 美国受众共鸣问题
4. 商品CTR → 商品点击兴趣
5. 订单转化率 → 成交问题
6. ROAS/CPC → 付费流量问题

规则：

- 必须引用真实输入的数据。
- 有账号基准时优先和账号自身比较。
- 没有账号基准时才能使用内部SOP工作区间。
- 内部区间不是TikTok官方标准。
- 最终只选一个最优先修复环节。
- 不要一次把全部变量都改掉。
- 优化脚本继续保持：
  不露脸、第一视角UGC、强手部动作、小场景。
- 输出逐秒脚本。
- 英文必须自然美式口语。
- 不虚构产品功能。
- 严格按照JSON Schema。
""".strip()


def review_script(
    client,
    category,
    product_name,
    selling_points,
    original_script,
    metrics,
    baseline,
    assessment,
    traffic_type,
):
    started = (
        time.perf_counter()
    )

    config = (
        types.GenerateContentConfig(
            thinking_config=(
                thinking_config()
            ),

            max_output_tokens=4200,

            response_mime_type=(
                "application/json"
            ),

            response_json_schema=(
                REVIEW_SCHEMA
            ),
        )
    )

    response, call_meta = (
        generate_resilient(
            client=client,

            contents=(
                build_review_prompt(
                    category,
                    product_name,
                    selling_points,
                    original_script,
                    metrics,
                    baseline,
                    assessment,
                    traffic_type,
                )
            ),

            config=config,
        )
    )

    result = parse_json_output(
        response.text
    )

    metadata = {
        "analysis_seconds":
            round(
                time.perf_counter()
                - started,
                1,
            ),

        **call_meta,
    }

    return (
        result,
        metadata,
    )


# ============================================================
# 29. Review Script DataFrame
# ============================================================

def review_script_dataframe(
    rows,
):
    output = []

    for row in rows or []:
        output.append(
            {
                "分镜序号":
                    row.get(
                        "sequence",
                        "",
                    ),

                "时间段":
                    row.get(
                        "time_range",
                        "",
                    ),

                "景别/机位":
                    row.get(
                        "shot",
                        "",
                    ),

                "画面描述(道具/动作)":
                    row.get(
                        "visual",
                        "",
                    ),

                "手部动作":
                    row.get(
                        "hand_action",
                        "",
                    ),

                "英文口播/字幕":
                    row.get(
                        "copy_en",
                        "",
                    ),

                "音效/节奏提示":
                    row.get(
                        "audio",
                        "",
                    ),

                "设计目的(底层逻辑)":
                    row.get(
                        "rationale",
                        "",
                    ),
            }
        )

    return pd.DataFrame(
        output
    )


# ============================================================
# 30. Init
# ============================================================

render_login()

st.title(
    APP_TITLE
)

if not st.session_state[
    "authenticated"
]:
    st.info(
        "请先在左侧登录。"
    )

    st.stop()


render_sidebar_history()


api_key = get_api_key()

if not api_key:
    st.error(
        "系统未配置AI密钥，请联系管理员。"
    )

    client = None

else:
    client = create_client(
        api_key
    )


# ============================================================
# 31. Tabs
# ============================================================

tab_analysis, tab_review, tab_history = (
    st.tabs(
        [
            "爆款拆解",
            "数据复盘",
            "历史记录",
        ]
    )
)


# ============================================================
# TAB 1 - 爆款拆解
# ============================================================

with tab_analysis:

    # --------------------------------------------------------
    # ① 产品
    # --------------------------------------------------------

    st.markdown(
        "### ① 产品信息"
    )

    c1, c2, c3 = st.columns(
        3
    )

    with c1:
        tiktok_account = (
            st.text_input(
                "TikTok账号",
                placeholder="用于历史归档",
                key="analysis_account",
            )
        )

    with c2:
        category = (
            st.selectbox(
                "产品品类",
                PRODUCT_CATEGORIES,
                key="analysis_category",
            )
        )

    with c3:
        product_name = (
            st.text_input(
                "产品名称 / SKU",
                placeholder="例如：隐私印章",
                key="analysis_product",
            )
        )

    product_selling_points = (
        st.text_area(
            "我们的真实产品卖点",
            height=90,
            key="analysis_selling_points",
            placeholder=(
                "填写真实卖点即可。"
                "AI会结合爆款底层逻辑重新生成方向。"
            ),
        )
    )

    # --------------------------------------------------------
    # ② 上传爆款
    # --------------------------------------------------------

    st.markdown(
        "### ② 上传爆款视频"
    )

    uploaded_videos = (
        st.file_uploader(
            "支持同时上传1-5条 .mp4 视频",
            type=["mp4"],
            accept_multiple_files=True,
            key="comparison_videos",
        )
    )

    if uploaded_videos:
        count = len(
            uploaded_videos
        )

        total_mb = (
            sum(
                len(
                    video.getvalue()
                )
                for video
                in uploaded_videos
            )
            / 1024
            / 1024
        )

        st.caption(
            f"已上传 {count} 条 · 总大小 {total_mb:.2f} MB"
        )

        if (
            count
            > MAX_COMPARE_VIDEOS
        ):
            st.warning(
                f"单次最多分析 {MAX_COMPARE_VIDEOS} 条，"
                "请删除多余视频。"
            )

        signature = (
            video_batch_signature(
                uploaded_videos,
                category,
                product_name,
            )
        )

        if (
            st.session_state.get(
                "video_batch_signature"
            )
            != signature
        ):
            st.session_state[
                "video_batch_signature"
            ] = signature

            st.session_state[
                "video_analysis"
            ] = None

            st.session_state[
                "directions_result"
            ] = None

    analyze_button = (
        st.button(
            "解析爆款视频",
            type="primary",
            use_container_width=True,
            disabled=(
                client is None
                or not uploaded_videos
                or len(
                    uploaded_videos
                )
                > MAX_COMPARE_VIDEOS
            ),
            key="analyze_multi_video",
        )
    )

    if analyze_button:
        try:
            with st.spinner(
                "正在拆解视频并进行横向对比…"
            ):
                (
                    analysis_result,
                    analysis_meta,
                ) = analyze_videos(
                    client,
                    uploaded_videos,
                    category,
                    product_name,
                )

            st.session_state[
                "video_analysis"
            ] = analysis_result

            st.session_state[
                "video_analysis_meta"
            ] = analysis_meta

            append_history(
                {
                    "record_type":
                        "爆款对比解析",

                    "role":
                        st.session_state[
                            "role"
                        ],

                    "operator":
                        st.session_state[
                            "operator"
                        ],

                    "tiktok_account":
                        tiktok_account,

                    "product_category":
                        category,

                    "product_name":
                        product_name,

                    "selling_points":
                        product_selling_points,

                    "video_names":
                        " | ".join(
                            [
                                video.name
                                for video
                                in uploaded_videos
                            ]
                        ),

                    "video_count":
                        len(
                            uploaded_videos
                        ),

                    "model_used":
                        analysis_meta.get(
                            "model_used",
                            "",
                        ),

                    "fallback_used":
                        analysis_meta.get(
                            "fallback_used",
                            "",
                        ),

                    "retry_count":
                        analysis_meta.get(
                            "retry_count",
                            "",
                        ),

                    "analysis_seconds":
                        analysis_meta.get(
                            "analysis_seconds",
                            "",
                        ),

                    "full_output_json":
                        json_dumps(
                            analysis_result
                        ),
                }
            )

            st.success(
                "爆款拆解完成"
            )

        except Exception as exc:
            st.error(
                friendly_error(
                    exc
                )
            )

    # --------------------------------------------------------
    # 中文拆解
    # --------------------------------------------------------

    analysis_result = (
        st.session_state.get(
            "video_analysis"
        )
    )

    if analysis_result:

        summary = (
            analysis_result.get(
                "comparison_summary",
                {},
            )
        )

        st.markdown(
            "### ③ 中文爆款拆解"
        )

        st.info(
            summary.get(
                "one_sentence_core",
                "",
            )
        )

        with st.expander(
            "查看完整爆款拆解",
            expanded=False,
        ):

            st.markdown(
                "**爆款脚本路线**"
            )

            st.write(
                summary.get(
                    "common_script_route",
                    "",
                )
            )

            p1, p2 = st.columns(
                2
            )

            with p1:
                st.markdown(
                    "**人群画像**"
                )

                st.write(
                    summary.get(
                        "common_audience",
                        "",
                    )
                )

            with p2:
                st.markdown(
                    "**年龄预估**"
                )

                st.write(
                    summary.get(
                        "age_estimate",
                        "",
                    )
                )

            st.markdown(
                "**前3秒 Hook**"
            )

            st.write(
                summary.get(
                    "common_hook_pattern",
                    "",
                )
            )

            st.markdown(
                "**画面与节奏**"
            )

            st.write(
                summary.get(
                    "visual_rhythm",
                    "",
                )
            )

            st.markdown(
                "**最值得吸收的3点**"
            )

            for item in summary.get(
                "top_absorb_points",
                [],
            ):
                st.markdown(
                    f"- {item}"
                )

            if (
                len(
                    analysis_result.get(
                        "videos",
                        [],
                    )
                )
                > 1
            ):
                st.markdown(
                    "**多视频关键差异**"
                )

                st.write(
                    summary.get(
                        "key_differences",
                        "",
                    )
                )

            st.divider()

            st.markdown(
                "**逐条视频拆解**"
            )

            videos = (
                analysis_result.get(
                    "videos",
                    [],
                )
            )

            tabs = st.tabs(
                [
                    f"视频{i + 1}"
                    for i
                    in range(
                        len(videos)
                    )
                ]
            )

            for tab, video in zip(
                tabs,
                videos,
            ):
                with tab:
                    st.markdown(
                        f'**一句话核心：** '
                        f'{video.get("one_sentence_core", "")}'
                    )

                    st.markdown(
                        "**爆款脚本路线**"
                    )

                    st.write(
                        video.get(
                            "script_route",
                            "",
                        )
                    )

                    st.markdown(
                        "**人群 / 年龄预估**"
                    )

                    st.write(
                        video.get(
                            "audience_profile",
                            "",
                        )
                    )

                    st.write(
                        video.get(
                            "age_estimate",
                            "",
                        )
                    )

                    st.markdown(
                        "**前3秒 Hook**"
                    )

                    st.write(
                        video.get(
                            "first_3s_hook",
                            "",
                        )
                    )

                    st.markdown(
                        "**画面与节奏**"
                    )

                    st.write(
                        video.get(
                            "visual_rhythm",
                            "",
                        )
                    )

                    st.markdown(
                        "**最值得吸收的3点**"
                    )

                    for point in video.get(
                        "top_absorb_points",
                        [],
                    ):
                        st.markdown(
                            f"- {point}"
                        )

        # ----------------------------------------------------
        # ④ 3个参考方向
        # ----------------------------------------------------

        st.markdown(
            "### ④ 生成3个参考方向"
        )

        allowed_scenes = (
            st.multiselect(
                "可使用的小场景",
                options=list(
                    SCENE_LIBRARY.keys()
                ),
                default=list(
                    SCENE_LIBRARY.keys()
                ),
                key="allowed_scenes",
            )
        )

        generate_button = (
            st.button(
                "生成3个拍摄方向",
                type="primary",
                use_container_width=True,
                disabled=(
                    client is None
                    or not product_selling_points.strip()
                    or not allowed_scenes
                ),
                key="generate_three_directions",
            )
        )

        if generate_button:
            try:
                with st.spinner(
                    "正在生成不露脸第一视角UGC脚本…"
                ):
                    (
                        directions_result,
                        directions_meta,
                    ) = generate_directions(
                        client,
                        category,
                        product_name,
                        product_selling_points,
                        analysis_result,
                        allowed_scenes,
                    )

                st.session_state[
                    "directions_result"
                ] = directions_result

                st.session_state[
                    "directions_meta"
                ] = directions_meta

                append_history(
                    {
                        "record_type":
                            "3方向拍摄脚本",

                        "role":
                            st.session_state[
                                "role"
                            ],

                        "operator":
                            st.session_state[
                                "operator"
                            ],

                        "tiktok_account":
                            tiktok_account,

                        "product_category":
                            category,

                        "product_name":
                            product_name,

                        "selling_points":
                            product_selling_points,

                        "video_names":
                            " | ".join(
                                [
                                    video.name
                                    for video
                                    in uploaded_videos
                                ]
                            ),

                        "video_count":
                            len(
                                uploaded_videos
                            ),

                        "model_used":
                            directions_meta.get(
                                "model_used",
                                "",
                            ),

                        "fallback_used":
                            directions_meta.get(
                                "fallback_used",
                                "",
                            ),

                        "retry_count":
                            directions_meta.get(
                                "retry_count",
                                "",
                            ),

                        "analysis_seconds":
                            directions_meta.get(
                                "analysis_seconds",
                                "",
                            ),

                        "full_output_json":
                            json_dumps(
                                directions_result
                            ),
                    }
                )

                st.success(
                    "3个拍摄方向已生成"
                )

            except Exception as exc:
                st.error(
                    friendly_error(
                        exc
                    )
                )

    # --------------------------------------------------------
    # 方向输出
    # --------------------------------------------------------

    directions_result = (
        st.session_state.get(
            "directions_result"
        )
    )

    if (
        directions_result
        and analysis_result
    ):
        st.markdown(
            "### ⑤ 拍摄执行"
        )

        directions = (
            directions_result.get(
                "directions",
                [],
            )
        )

        direction_tabs = st.tabs(
            [
                f"方向 {i + 1}"
                for i
                in range(
                    len(directions)
                )
            ]
        )

        for tab, direction in zip(
            direction_tabs,
            directions,
        ):
            with tab:
                st.markdown(
                    f'### {direction.get("direction_name", "")}'
                )

                st.write(
                    direction.get(
                        "core_idea",
                        "",
                    )
                )

                info1, info2, info3 = (
                    st.columns(
                        3
                    )
                )

                info1.metric(
                    "小场景",
                    direction.get(
                        "small_scene",
                        "-",
                    ),
                )

                info2.metric(
                    "预计时长",
                    direction.get(
                        "estimated_duration",
                        "-",
                    ),
                )

                info3.metric(
                    "目标人群",
                    direction.get(
                        "target_audience",
                        "-",
                    ),
                )

                with st.expander(
                    "拍摄设计",
                    expanded=False,
                ):
                    st.markdown(
                        "**不露脸设计**"
                    )

                    st.write(
                        direction.get(
                            "faceless_design",
                            "",
                        )
                    )

                    st.markdown(
                        "**第一视角 UGC**"
                    )

                    st.write(
                        direction.get(
                            "pov_ugc",
                            "",
                        )
                    )

                    st.markdown(
                        "**强手部动作**"
                    )

                    st.write(
                        direction.get(
                            "hand_action_design",
                            "",
                        )
                    )

                    st.markdown(
                        "**可吸收点**"
                    )

                    for item in direction.get(
                        "absorb_points",
                        [],
                    ):
                        st.markdown(
                            f"- {item}"
                        )

                    st.markdown(
                        "**差异化点**"
                    )

                    for item in direction.get(
                        "differentiation_points",
                        [],
                    ):
                        st.markdown(
                            f"- {item}"
                        )

                direction_df = (
                    direction_to_dataframe(
                        direction
                    )
                )

                display_columns = [
                    "分镜序号",
                    "时间段",
                    "景别/机位",
                    "画面描述(道具/动作)",
                    "手部动作",
                    "英文口播/字幕",
                    "音效/节奏提示",
                    "设计目的(底层逻辑)",
                ]

                st.dataframe(
                    direction_df[
                        display_columns
                    ],
                    hide_index=True,
                    use_container_width=True,
                )

        excel_bytes = (
            build_full_excel(
                analysis_result,
                directions_result,
            )
        )

        st.download_button(
            "一键导出完整 Excel",
            data=excel_bytes,
            file_name=(
                "TikTok爆款拆解_拍摄执行_"
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
            type="primary",
            use_container_width=True,
        )


# ============================================================
# TAB 2 - 数据复盘
# ============================================================

with tab_review:

    st.markdown(
        "### ① 上传拍摄脚本"
    )

    review_excel = (
        st.file_uploader(
            "上传之前导出的 Excel",
            type=["xlsx"],
            key="review_excel",
        )
    )

    selected_sheet = None

    if review_excel:
        try:
            script_sheets = (
                get_script_sheets(
                    review_excel
                )
            )

            if script_sheets:
                selected_sheet = (
                    st.selectbox(
                        "选择本次实际发布的脚本",
                        script_sheets,
                        key="review_sheet",
                    )
                )

                if st.button(
                    "读取该脚本",
                    use_container_width=True,
                    key="load_review_script",
                ):
                    st.session_state[
                        "review_original_script"
                    ] = (
                        excel_sheet_to_script_text(
                            review_excel,
                            selected_sheet,
                        )
                    )

                    st.success(
                        "脚本已读取"
                    )

            else:
                st.warning(
                    "未识别到可复盘的脚本Sheet。"
                )

        except Exception as exc:
            st.error(
                f"Excel读取失败：{exc}"
            )

    r1, r2, r3 = st.columns(
        3
    )

    with r1:
        review_account = (
            st.text_input(
                "TikTok账号",
                key="review_account",
            )
        )

    with r2:
        review_category = (
            st.selectbox(
                "产品品类",
                PRODUCT_CATEGORIES,
                key="review_category",
            )
        )

    with r3:
        review_product = (
            st.text_input(
                "产品名称 / SKU",
                key="review_product",
            )
        )

    review_selling_points = (
        st.text_area(
            "产品核心卖点",
            height=80,
            key="review_selling_points",
        )
    )

    review_traffic = (
        st.selectbox(
            "流量类型",
            TRAFFIC_TYPES,
            key="review_traffic",
        )
    )

    original_script = (
        st.text_area(
            "原始脚本（可编辑）",
            height=260,
            key="review_original_script",
        )
    )

    # --------------------------------------------------------
    # 数据
    # --------------------------------------------------------

    st.markdown(
        "### ② 视频核心数据"
    )

    m1, m2, m3 = (
        st.columns(
            3
        )
    )

    with m1:
        retention = optional_number(
            "前3秒留存率 (%)",
            "review_retention",
            100.0,
        )

        completion = optional_number(
            "平均完播率 (%)",
            "review_completion",
            100.0,
        )

    with m2:
        ctr = optional_number(
            "商品锚点CTR (%)",
            "review_ctr",
            100.0,
            0.01,
        )

        conversion = optional_number(
            "订单转化率 (%)",
            "review_conversion",
            100.0,
            0.01,
        )

    with m3:
        engagement = optional_number(
            "互动率 (%)",
            "review_engagement",
            100.0,
            0.01,
        )

    with st.expander(
        "广告数据（选填）",
        expanded=False,
    ):
        a1, a2 = st.columns(
            2
        )

        with a1:
            actual_roas = (
                optional_number(
                    "实际ROAS",
                    "review_roas",
                )
            )

            target_roas = (
                optional_number(
                    "目标ROAS",
                    "review_target_roas",
                )
            )

        with a2:
            actual_cpc = (
                optional_number(
                    "实际CPC ($)",
                    "review_cpc",
                    step=0.01,
                )
            )

            target_cpc = (
                optional_number(
                    "目标CPC ($)",
                    "review_target_cpc",
                    step=0.01,
                )
            )

    with st.expander(
        "账号近7天基准（建议填写）",
        expanded=False,
    ):
        b1, b2, b3 = st.columns(
            3
        )

        with b1:
            base_retention = (
                optional_number(
                    "账号平均3秒留存 (%)",
                    "base_retention",
                    100.0,
                )
            )

            base_completion = (
                optional_number(
                    "账号平均完播率 (%)",
                    "base_completion",
                    100.0,
                )
            )

        with b2:
            base_ctr = (
                optional_number(
                    "账号平均CTR (%)",
                    "base_ctr",
                    100.0,
                    0.01,
                )
            )

            base_conversion = (
                optional_number(
                    "账号平均转化率 (%)",
                    "base_conversion",
                    100.0,
                    0.01,
                )
            )

        with b3:
            base_engagement = (
                optional_number(
                    "账号平均互动率 (%)",
                    "base_engagement",
                    100.0,
                    0.01,
                )
            )

    metrics = compact_dict(
        {
            "retention_3s_pct":
                retention,

            "completion_rate_pct":
                completion,

            "product_ctr_pct":
                ctr,

            "order_conversion_pct":
                conversion,

            "engagement_rate_pct":
                engagement,

            "actual_roas":
                actual_roas,

            "target_roas":
                target_roas,

            "actual_cpc":
                actual_cpc,

            "target_cpc":
                target_cpc,
        }
    )

    baseline = compact_dict(
        {
            "retention_3s_pct":
                base_retention,

            "completion_rate_pct":
                base_completion,

            "product_ctr_pct":
                base_ctr,

            "order_conversion_pct":
                base_conversion,

            "engagement_rate_pct":
                base_engagement,
        }
    )

    assessment = (
        build_metric_assessment(
            metrics,
            baseline,
        )
    )

    review_button = (
        st.button(
            "开始数据复盘",
            type="primary",
            use_container_width=True,
            disabled=(
                client is None
            ),
            key="run_review",
        )
    )

    if review_button:

        if not original_script.strip():
            st.error(
                "请先上传或填写原始脚本。"
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
                    "正在诊断视频跑偏环节…"
                ):
                    (
                        review_result,
                        review_meta,
                    ) = review_script(
                        client,
                        review_category,
                        review_product,
                        review_selling_points,
                        original_script,
                        metrics,
                        baseline,
                        assessment,
                        review_traffic,
                    )

                st.session_state[
                    "review_result"
                ] = review_result

                st.session_state[
                    "review_meta"
                ] = review_meta

                append_history(
                    {
                        "record_type":
                            "数据复盘",

                        "role":
                            st.session_state[
                                "role"
                            ],

                        "operator":
                            st.session_state[
                                "operator"
                            ],

                        "tiktok_account":
                            review_account,

                        "product_category":
                            review_category,

                        "product_name":
                            review_product,

                        "selling_points":
                            review_selling_points,

                        "model_used":
                            review_meta.get(
                                "model_used",
                                "",
                            ),

                        "fallback_used":
                            review_meta.get(
                                "fallback_used",
                                "",
                            ),

                        "retry_count":
                            review_meta.get(
                                "retry_count",
                                "",
                            ),

                        "analysis_seconds":
                            review_meta.get(
                                "analysis_seconds",
                                "",
                            ),

                        "priority_issue":
                            review_result.get(
                                "priority_issue",
                                "",
                            ),

                        "diagnosis_summary":
                            review_result.get(
                                "diagnosis_summary",
                                "",
                            ),

                        "metrics_json":
                            json_dumps(
                                metrics
                            ),

                        "account_baseline_json":
                            json_dumps(
                                baseline
                            ),

                        "full_output_json":
                            json_dumps(
                                review_result
                            ),
                    }
                )

                st.success(
                    "复盘完成"
                )

            except Exception as exc:
                st.error(
                    friendly_error(
                        exc
                    )
                )

    if st.session_state.get(
        "review_result"
    ):
        result = (
            st.session_state[
                "review_result"
            ]
        )

        st.divider()

        st.markdown(
            "### 复盘结论"
        )

        st.info(
            result.get(
                "diagnosis_summary",
                "",
            )
        )

        st.markdown(
            "**最优先修复：** "
            + result.get(
                "priority_issue",
                "",
            )
        )

        with st.expander(
            "查看逐项诊断",
            expanded=False,
        ):
            diagnosis_df = (
                pd.DataFrame(
                    result.get(
                        "metric_diagnosis",
                        [],
                    )
                )
            )

            if not diagnosis_df.empty:
                st.dataframe(
                    diagnosis_df,
                    hide_index=True,
                    use_container_width=True,
                )

        optimized_df = (
            review_script_dataframe(
                result.get(
                    "optimized_script",
                    [],
                )
            )
        )

        st.markdown(
            "### 优化后脚本"
        )

        st.dataframe(
            optimized_df,
            hide_index=True,
            use_container_width=True,
        )


# ============================================================
# TAB 3 - History
# ============================================================

with tab_history:

    history = scoped_history()

    if history.empty:
        st.info(
            "暂无历史记录。"
        )

    else:
        filtered = history.copy()

        f1, f2, f3 = (
            st.columns(
                3
            )
        )

        with f1:
            if (
                st.session_state[
                    "role"
                ]
                == "主账号(Admin)"
            ):
                operators = (
                    ["全部"]
                    + sorted(
                        [
                            x
                            for x
                            in filtered[
                                "operator"
                            ].unique()
                            if x
                        ]
                    )
                )

                operator_filter = (
                    st.selectbox(
                        "操作人",
                        operators,
                        key="history_operator",
                    )
                )

                if (
                    operator_filter
                    != "全部"
                ):
                    filtered = filtered[
                        filtered[
                            "operator"
                        ]
                        == operator_filter
                    ]

        with f2:
            types_list = (
                ["全部"]
                + sorted(
                    [
                        x
                        for x
                        in filtered[
                            "record_type"
                        ].unique()
                        if x
                    ]
                )
            )

            type_filter = (
                st.selectbox(
                    "类型",
                    types_list,
                    key="history_type",
                )
            )

            if type_filter != "全部":
                filtered = filtered[
                    filtered[
                        "record_type"
                    ]
                    == type_filter
                ]

        with f3:
            accounts = (
                ["全部"]
                + sorted(
                    [
                        x
                        for x
                        in filtered[
                            "tiktok_account"
                        ].unique()
                        if x
                    ]
                )
            )

            account_filter = (
                st.selectbox(
                    "TikTok账号",
                    accounts,
                    key="history_account",
                )
            )

            if (
                account_filter
                != "全部"
            ):
                filtered = filtered[
                    filtered[
                        "tiktok_account"
                    ]
                    == account_filter
                ]

        display_columns = [
            "created_at_cn",
            "record_type",
            "operator",
            "tiktok_account",
            "product_name",
            "video_count",
            "priority_issue",
        ]

        st.dataframe(
            filtered[
                display_columns
            ].iloc[::-1],
            hide_index=True,
            use_container_width=True,
        )

        history_csv = (
            filtered.to_csv(
                index=False,
                encoding="utf-8-sig",
            )
            .encode(
                "utf-8-sig"
            )
        )

        st.download_button(
            "下载历史 CSV",
            data=history_csv,
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
