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

MAX_COMPARE_VIDEOS = 5

INLINE_BATCH_MAX_MB = 18

FILE_POLL_INTERVAL_SEC = 2
FILE_PROCESS_TIMEOUT_SEC = 180

DEFAULT_STAFF_PASSWORD = "8888"
DEFAULT_ADMIN_PASSWORD = "8888-admin"

HISTORY_FILE = Path("history_log.csv")
HISTORY_LOCK = threading.Lock()

CN_TZ = ZoneInfo("Asia/Shanghai")


# ============================================================
# 2. 场景库
# ============================================================

SCENE_LIBRARY = {

    "客厅·茶几快递拆包区":
        "茶几上放快递盒、信封、快递袋、标签纸。"
        "适合隐私印章、开箱、文件处理类产品。",

    "客厅·沙发边桌":
        "只拍手和桌面，不拍正脸。"
        "适合耳机、便携用品、收纳类产品。",

    "客厅·电视柜/玄关柜":
        "第一视角拿取、使用、放回。"
        "适合家居小工具和日常用品。",

    "厨房·岛台正面":
        "岛台作为主要操作区。"
        "适合厨房垃圾桶、厨房工具、清洁类产品。",

    "厨房·切菜区":
        "切菜、处理厨余、动作明显。"
        "适合挂式厨房垃圾桶和厨房用品。",

    "厨房·水槽旁":
        "洗、擦、收纳、清理的真实生活动作。"
        "适合清洁和厨房效率类产品。",

    "卧室·床头柜":
        "睡前或起床后的真实生活场景。"
        "适合耳机、阅读用品、个人小工具。",

    "卧室·梳妆台":
        "镜前但不露脸，以手和台面为主体。"
        "适合个人护理、收纳用品。",

    "卧室·床面":
        "俯拍床面或第一人称展示。"
        "适合开箱和便携用品。",

    "卫生间·洗手台":
        "不拍正脸，以洗手台、产品、手部为主体。"
        "适合个护和清洁用品。",

    "卫生间·镜柜":
        "第一人称开镜柜、取产品、使用、放回。"
        "适合收纳和个人护理用品。",

    "阳台·落地窗边桌":
        "自然光场景。"
        "适合展示产品材质、外观和生活方式。",

    "纯桌面·白桌":
        "完全不露脸，只出现手、产品和必要道具。"
        "适合功能型产品和高频脚本测试。",

    "纯桌面·快递箱/文件场景":
        "桌面放快递标签、信封、文件、账单。"
        "特别适合隐私保护印章。",
}


PERSPECTIVE_OPTIONS = [
    "第一人称 POV",
    "第三人称手部/局部视角",
]


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
# 3. EXCEL COLUMNS
# ============================================================

ANALYSIS_VIDEO_COLUMNS = [
    "视频编号",
    "文件名",
    "一句话核心",
    "该视频推理卖点",
    "爆款脚本路线",
    "人群画像",
    "年龄预估",
    "前3秒Hook",
    "画面与节奏",
    "最值得吸收的3点",
    "参考价值判断",
    "推荐指数",
]


DIRECTION_SUMMARY_COLUMNS = [
    "方向",
    "核心思路",
    "目标人群",
    "前3秒Hook",
    "产品切入方式",
    "推荐视角",
    "推荐小场景",
    "可吸收点",
    "差异化点",
]


FINAL_SCRIPT_COLUMNS = [
    "分镜序号",
    "时间段",
    "机位/视角",
    "画面描述(道具/动作)",
    "手部动作",
    "中文口播/字幕参考",
    "英文口播/字幕",
    "音效/节奏提示",
    "爆款吸收点",
    "差异化处理",
    "设计目的(底层逻辑)",
]


OLD_SCRIPT_COLUMNS = [
    "分镜序号",
    "景别/机位",
    "画面描述(道具/动作)",
    "英文口播文案/字幕",
    "音效/节奏提示",
    "设计目的(底层逻辑)",
]


# ============================================================
# 4. HISTORY
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

    "input_selling_points",
    "inferred_selling_points",
    "effective_selling_points",
    "selling_point_mode",

    "reference_video_index",
    "reference_video_name",

    "direction_name",

    "selected_scene",
    "selected_perspective",

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
# 5. INTERNAL SOP
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
# 6. JSON SCHEMAS
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

        "common_inferred_selling_points": {

            "type": "array",
            "minItems": 3,
            "maxItems": 5,

            "items": {
                "type": "string",
            },
        },

        "recommended_reference_video_index": {
            "type": "integer",
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

                    "inferred_selling_points": {

                        "type": "array",
                        "minItems": 3,
                        "maxItems": 5,

                        "items": {
                            "type": "string",
                        },
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

                    "fit_reason": {
                        "type": "string",
                    },

                    "recommend_score": {
                        "type": "integer",
                    },

                    "selling_point_relation": {
                        "type": "string",
                    },

                    "selling_point_relation_reason": {
                        "type": "string",
                    },

                    "blended_selling_points": {
                        "type": "string",
                    },

                    "suggested_mode": {
                        "type": "string",
                    },
                },

                "required": [
                    "video_index",
                    "filename",
                    "one_sentence_core",
                    "inferred_selling_points",
                    "script_route",
                    "audience_profile",
                    "age_estimate",
                    "first_3s_hook",
                    "visual_rhythm",
                    "top_absorb_points",
                    "fit_reason",
                    "recommend_score",
                    "selling_point_relation",
                    "selling_point_relation_reason",
                    "blended_selling_points",
                    "suggested_mode",
                ],
            },
        },
    },

    "required": [
        "comparison_summary",
        "common_inferred_selling_points",
        "recommended_reference_video_index",
        "videos",
    ],
}


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

                    "hook": {
                        "type": "string",
                    },

                    "product_entry": {
                        "type": "string",
                    },

                    "recommended_perspective": {
                        "type": "string",
                    },

                    "recommended_scene": {
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
                },

                "required": [
                    "direction_name",
                    "core_idea",
                    "target_audience",
                    "hook",
                    "product_entry",
                    "recommended_perspective",
                    "recommended_scene",
                    "absorb_points",
                    "differentiation_points",
                ],
            },
        },
    },

    "required": [
        "directions",
    ],
}


FINAL_SCRIPT_SCHEMA = {

    "type": "object",

    "properties": {

        "shooting_notes": {
            "type": "string",
        },

        "storyboard": {

            "type": "array",
            "minItems": 6,
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

                    "copy_cn": {
                        "type": "string",
                    },

                    "copy_en": {
                        "type": "string",
                    },

                    "audio": {
                        "type": "string",
                    },

                    "absorb_point": {
                        "type": "string",
                    },

                    "difference_point": {
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
                    "copy_cn",
                    "copy_en",
                    "audio",
                    "absorb_point",
                    "difference_point",
                    "rationale",
                ],
            },
        },
    },

    "required": [
        "shooting_notes",
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

        "optimized_script": {

            "type": "array",
            "minItems": 6,
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

                    "copy_cn": {
                        "type": "string",
                    },

                    "copy_en": {
                        "type": "string",
                    },

                    "audio": {
                        "type": "string",
                    },

                    "absorb_point": {
                        "type": "string",
                    },

                    "difference_point": {
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
                    "copy_cn",
                    "copy_en",
                    "audio",
                    "absorb_point",
                    "difference_point",
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
# 7. PAGE CONFIG
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
        font-size: 1.9rem !important;
        margin-bottom: .8rem !important;
    }

    h2 {
        font-size: 1.35rem !important;
    }

    h3 {
        font-size: 1.10rem !important;
        margin-top: 1rem !important;
        margin-bottom: .55rem !important;
    }

    div[data-testid="stButton"] button,
    div[data-testid="stDownloadButton"] button {
        border-radius: 8px;
        font-weight: 600;
        min-height: 42px;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 8. HELPERS
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


def json_dumps(
    data,
):

    return json.dumps(
        data,
        ensure_ascii=False,
        separators=(",", ":"),
    )


def parse_json_output(
    raw_text,
):

    if not raw_text:

        raise ValueError(
            "AI未返回有效结果。"
        )

    try:

        return json.loads(
            raw_text
        )

    except json.JSONDecodeError as exc:

        raise ValueError(
            "AI返回格式异常，请重新执行。"
        ) from exc


def compact_dict(
    data,
):

    return {

        key: value

        for key, value
        in data.items()

        if value is not None
        and value != ""
    }


def thinking_config():

    return types.ThinkingConfig(
        thinking_level="minimal"
    )


def list_to_joined(
    items,
):

    if not items:
        return ""

    return "; ".join(
        [
            clean_text(item)
            for item
            in items
            if clean_text(item)
        ]
    )


def safe_int(
    value,
    default=0,
):

    try:

        return int(
            value
        )

    except Exception:

        return default


def parse_optional_float(
    value,
):

    value = clean_text(
        value
    )

    if not value:
        return None

    try:

        return float(
            value
        )

    except ValueError:

        return None


def make_signature(
    *values,
):

    serialized = []

    for value in values:

        if isinstance(
            value,
            (
                dict,
                list,
            )
        ):

            serialized.append(
                json.dumps(
                    value,
                    ensure_ascii=False,
                    sort_keys=True,
                )
            )

        else:

            serialized.append(
                clean_text(
                    value
                )
            )

    raw = "||".join(
        serialized
    )

    return hashlib.sha256(
        raw.encode(
            "utf-8"
        )
    ).hexdigest()[:24]


def video_batch_signature(
    uploaded_videos,
    category,
    product_name,
    input_selling_points,
):

    hasher = hashlib.sha256()

    for video in uploaded_videos or []:

        hasher.update(
            video.getvalue()
        )

        hasher.update(
            clean_text(
                video.name
            ).encode(
                "utf-8"
            )
        )

    hasher.update(
        clean_text(
            category
        ).encode(
            "utf-8"
        )
    )

    hasher.update(
        clean_text(
            product_name
        ).encode(
            "utf-8"
        )
    )

    hasher.update(
        clean_text(
            input_selling_points
        ).encode(
            "utf-8"
        )
    )

    return hasher.hexdigest()[:24]


# ============================================================
# 9. SESSION
# ============================================================

def init_session():

    defaults = {

        "authenticated": False,

        "role": "",

        "operator": "",

        "video_analysis_result": None,

        "video_analysis_meta": {},

        "video_batch_signature": "",

        "selected_reference_video_index": None,

        "directions_result": None,

        "directions_meta": {},

        "directions_context_signature": "",

        "selected_direction_index": 0,

        "last_direction_control_signature": "",

        "final_script_result": None,

        "final_script_meta": {},

        "final_script_context_signature": "",

        "review_result": None,

        "review_meta": {},

        "review_original_script": "",
    }

    for key, value in defaults.items():

        if key not in st.session_state:

            st.session_state[
                key
            ] = value


def logout():

    st.session_state[
        "authenticated"
    ] = False

    st.session_state[
        "role"
    ] = ""

    st.session_state[
        "operator"
    ] = ""

    st.rerun()


init_session()


# ============================================================
# 10. RETRY
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
        for marker
        in TRANSIENT_MARKERS
    )


def is_model_error(
    exc,
):

    text = str(
        exc
    ).upper()

    return any(
        marker in text
        for marker
        in MODEL_ERROR_MARKERS
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
            "Gemini API Key 认证失败，"
            "请联系管理员检查 Streamlit Secrets。"
        )

    if (
        "400" in text
        or "INVALID_ARGUMENT" in text
    ):

        return (
            "AI请求参数异常，"
            "请联系管理员检查模型调用或结构化输出配置。"
        )

    if (
        "429" in text
        or "RESOURCE_EXHAUSTED" in text
    ):

        return (
            "当前 AI 请求较多，"
            "系统已自动重试并尝试备用线路，"
            "请稍后再次执行。"
        )

    if (
        "503" in text
        or "UNAVAILABLE" in text
        or "HIGH DEMAND" in text
    ):

        return (
            "当前 AI 服务繁忙，"
            "系统已自动重试并尝试备用线路，"
            "请稍后再次执行。"
        )

    return (
        "AI暂时未完成本次任务，"
        "请稍后重新执行。"
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

                metadata = {

                    "model_used":
                        model_name,

                    "fallback_used":
                        model_index > 0,

                    "retry_count":
                        max(
                            total_attempts - 1,
                            0,
                        ),
                }

                return (
                    response,
                    metadata,
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
# 11. HISTORY
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

    dataframe = (
        dataframe.copy()
    )

    for column in HISTORY_COLUMNS:

        if column not in dataframe.columns:

            dataframe[
                column
            ] = ""

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

    temp_file = (
        HISTORY_FILE.with_name(
            "history_log.tmp.csv"
        )
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

        column:
            clean_text(
                record.get(
                    column,
                    "",
                )
            )

        for column
        in HISTORY_COLUMNS
    }

    if not row[
        "record_id"
    ]:

        row[
            "record_id"
        ] = uuid.uuid4().hex[:12]

    now_utc = (
        datetime.now(
            timezone.utc
        )
    )

    row[
        "created_at_utc"
    ] = now_utc.isoformat(
        timespec="seconds"
    )

    row[
        "created_at_cn"
    ] = (
        now_utc
        .astimezone(
            CN_TZ
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

    dataframe = (
        load_history()
    )

    if (
        st.session_state[
            "role"
        ]
        == "主账号(Admin)"
    ):

        return dataframe

    return dataframe[
        dataframe[
            "operator"
        ]
        == st.session_state[
            "operator"
        ]
    ].copy()


# ============================================================
# 12. LOGIN
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
                if role
                == "主账号(Admin)"
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


def render_sidebar_history():

    if not st.session_state[
        "authenticated"
    ]:

        return

    history = (
        scoped_history()
    )

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

        filtered = (
            history.copy()
        )

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
                        value
                        for value
                        in filtered[
                            "operator"
                        ].unique()
                        if value
                    ]
                )
            )

            operator_filter = (
                st.selectbox(
                    "操作人",
                    operators,
                    key="sidebar_operator_filter",
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

        st.caption(
            f"{len(filtered)} 条记录"
        )

        csv_data = (
            filtered
            .to_csv(
                index=False,
                encoding="utf-8-sig",
            )
            .encode(
                "utf-8-sig"
            )
        )

        st.download_button(
            "下载历史 CSV",
            data=csv_data,
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
# 13. EXCEL
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
        ].width = 23

    worksheet.freeze_panes = "A2"

    worksheet.auto_filter.ref = (
        worksheet.dimensions
    )


def analysis_videos_to_df(
    result,
):

    rows = []

    for video in result.get(
        "videos",
        [],
    ):

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

                "该视频推理卖点":
                    "\n".join(
                        [
                            f"{i + 1}. {item}"
                            for i, item
                            in enumerate(
                                video.get(
                                    "inferred_selling_points",
                                    [],
                                )
                            )
                        ]
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
                    "\n".join(
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
                    ),

                "参考价值判断":
                    video.get(
                        "fit_reason",
                        "",
                    ),

                "推荐指数":
                    video.get(
                        "recommend_score",
                        "",
                    ),
            }
        )

    return pd.DataFrame(
        rows,
        columns=ANALYSIS_VIDEO_COLUMNS,
    )


def directions_summary_to_df(
    result,
):

    rows = []

    for direction in result.get(
        "directions",
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

                "前3秒Hook":
                    direction.get(
                        "hook",
                        "",
                    ),

                "产品切入方式":
                    direction.get(
                        "product_entry",
                        "",
                    ),

                "推荐视角":
                    direction.get(
                        "recommended_perspective",
                        "",
                    ),

                "推荐小场景":
                    direction.get(
                        "recommended_scene",
                        "",
                    ),

                "可吸收点":
                    "\n".join(
                        [
                            f"- {item}"
                            for item
                            in direction.get(
                                "absorb_points",
                                [],
                            )
                        ]
                    ),

                "差异化点":
                    "\n".join(
                        [
                            f"- {item}"
                            for item
                            in direction.get(
                                "differentiation_points",
                                [],
                            )
                        ]
                    ),
            }
        )

    return pd.DataFrame(
        rows,
        columns=DIRECTION_SUMMARY_COLUMNS,
    )


def final_script_to_df(
    result,
):

    rows = []

    for shot in result.get(
        "storyboard",
        [],
    ):

        rows.append(
            {

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

                "机位/视角":
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

                "中文口播/字幕参考":
                    shot.get(
                        "copy_cn",
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

                "爆款吸收点":
                    shot.get(
                        "absorb_point",
                        "",
                    ),

                "差异化处理":
                    shot.get(
                        "difference_point",
                        "",
                    ),

                "设计目的(底层逻辑)":
                    shot.get(
                        "rationale",
                        "",
                    ),
            }
        )

    return pd.DataFrame(
        rows,
        columns=FINAL_SCRIPT_COLUMNS,
    )


def review_script_to_df(
    result,
):

    rows = []

    for shot in result.get(
        "optimized_script",
        [],
    ):

        rows.append(
            {

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

                "机位/视角":
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

                "中文口播/字幕参考":
                    shot.get(
                        "copy_cn",
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

                "爆款吸收点":
                    shot.get(
                        "absorb_point",
                        "",
                    ),

                "差异化处理":
                    shot.get(
                        "difference_point",
                        "",
                    ),

                "设计目的(底层逻辑)":
                    shot.get(
                        "rationale",
                        "",
                    ),
            }
        )

    return pd.DataFrame(
        rows,
        columns=FINAL_SCRIPT_COLUMNS,
    )


def build_analysis_export_excel(
    analysis_result,
    directions_result=None,
    final_script_result=None,
):

    output = io.BytesIO()

    summary = (
        analysis_result.get(
            "comparison_summary",
            {},
        )
    )

    summary_df = pd.DataFrame(
        [

            {
                "项目":
                    "一句话共同核心",

                "内容":
                    summary.get(
                        "one_sentence_core",
                        "",
                    ),
            },

            {
                "项目":
                    "共同爆款脚本路线",

                "内容":
                    summary.get(
                        "common_script_route",
                        "",
                    ),
            },

            {
                "项目":
                    "共同人群",

                "内容":
                    summary.get(
                        "common_audience",
                        "",
                    ),
            },

            {
                "项目":
                    "年龄预估",

                "内容":
                    summary.get(
                        "age_estimate",
                        "",
                    ),
            },

            {
                "项目":
                    "共同前3秒Hook",

                "内容":
                    summary.get(
                        "common_hook_pattern",
                        "",
                    ),
            },

            {
                "项目":
                    "共同画面与节奏",

                "内容":
                    summary.get(
                        "visual_rhythm",
                        "",
                    ),
            },

            {
                "项目":
                    "多视频关键差异",

                "内容":
                    summary.get(
                        "key_differences",
                        "",
                    ),
            },
        ]
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

        analysis_videos_to_df(
            analysis_result
        ).to_excel(
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

        if directions_result:

            directions_summary_to_df(
                directions_result
            ).to_excel(
                writer,
                index=False,
                sheet_name="3个方向概览",
            )

            format_sheet(
                writer.book[
                    "3个方向概览"
                ]
            )

        if final_script_result:

            final_script_to_df(
                final_script_result
            ).to_excel(
                writer,
                index=False,
                sheet_name="最终拍摄脚本",
            )

            format_sheet(
                writer.book[
                    "最终拍摄脚本"
                ]
            )

    output.seek(
        0
    )

    return output.getvalue()


def build_review_export_excel(
    review_result,
):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        review_script_to_df(
            review_result
        ).to_excel(
            writer,
            index=False,
            sheet_name="优化版脚本",
        )

        format_sheet(
            writer.book[
                "优化版脚本"
            ]
        )

    output.seek(
        0
    )

    return output.getvalue()


# ============================================================
# 14. FILE API
# ============================================================

def wait_until_active(
    client,
    uploaded_file,
):

    started = (
        time.monotonic()
    )

    current = (
        uploaded_file
    )

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
                "视频预处理超时，请压缩视频后重新上传。"
            )

        time.sleep(
            FILE_POLL_INTERVAL_SEC
        )

        current = (
            client.files.get(
                name=current.name
            )
        )


# ============================================================
# 15. PROMPTS
# ============================================================

def build_video_analysis_prompt(
    category,
    product_name,
    filenames,
    input_selling_points,
):

    file_lines = "\n".join(
        [
            f"视频{i + 1}：{filename}"
            for i, filename
            in enumerate(
                filenames
            )
        ]
    )

    return f"""
你是美国 TikTok Shop 爆款短视频分析负责人。

非常重要：

除文件名、必要英文原句之外，
所有分析结果必须使用简体中文。

禁止用英文回答：
- 卖点
- 脚本路线
- 人群
- 年龄
- Hook分析
- 节奏
- 推荐原因
- 差异分析
- 参考价值

如果这些字段输出英文，视为任务失败。

==================================================

本次上传：

{file_lines}

产品品类：

{category}

产品名称：

{product_name}

用户填写的真实产品卖点：

{input_selling_points if clean_text(input_selling_points) else "未填写"}

==================================================

请逐条单独分析视频。

不同视频可能主打完全不同的卖点。

所以每一条视频必须独立输出：

1. 一句话核心

2. inferred_selling_points

单独总结该视频真正主推的3-5条核心卖点。

必须中文。

不能混入其他视频卖点。

3. 爆款脚本路线

4. 人群画像

5. 年龄预估

6. 前3秒Hook

7. 画面与节奏

8. 最值得吸收的3点

9. 参考价值判断

10. 推荐指数 0-100

==================================================

如果用户填写了真实卖点：

请比较：

该视频推理卖点

VS

用户真实卖点

输出：

selling_point_relation

只能是：

similar

different

selling_point_relation_reason：

必须中文。

blended_selling_points：

如果融合，输出中文卖点表达。

suggested_mode：

只能是：

viral_first

user_first

blend

注意：

suggested_mode只是你的推荐。

最终决定权属于使用人。

你不能替使用人做最终选择。

==================================================

如果用户没有填写卖点：

selling_point_relation：

no_user_input

selling_point_relation_reason：

中文说明。

blended_selling_points：

用中文整理该视频推理卖点。

suggested_mode：

viral_first

==================================================

所有视频还要输出：

comparison_summary

包括：

- 一句话共同核心
- 共同爆款脚本路线
- 共同人群
- 年龄预估
- 共同Hook
- 共同画面节奏
- 最值得共同吸收3点
- 多视频关键差异

全部中文。

最后：

recommended_reference_video_index

推荐一条最值得作为主参考的视频。

但这只是AI推荐。

最终由使用人选择。

==================================================

禁止：

- 虚构产品功能
- 虚构TikTok后台数据
- 虚构销量
- 虚构认证
- 虚构医疗效果
- 复制品牌
- 复制原视频完整台词

严格按JSON Schema。
""".strip()


def build_directions_prompt(
    category,
    product_name,
    comparison_summary,
    selected_video,
    input_selling_points,
    effective_selling_points,
    selling_point_mode,
):

    scene_names = "\n".join(
        [
            f"- {scene}"
            for scene
            in SCENE_LIBRARY.keys()
        ]
    )

    return f"""
你是美国 TikTok Shop 中文短视频拍摄策划负责人。

除必要英文产品名之外，
本次输出全部使用简体中文。

禁止用英文输出：
- 方向名
- 核心思路
- 目标人群
- Hook分析
- 产品切入方式
- 推荐视角
- 推荐场景
- 可吸收点
- 差异化点

==================================================

产品：

{category}

{product_name}

用户真实卖点：

{input_selling_points if clean_text(input_selling_points) else "未填写"}

最终由使用人确认的有效卖点：

{effective_selling_points}

用户选择的卖点策略：

{selling_point_mode}

==================================================

主参考视频：

{json.dumps(
    selected_video,
    ensure_ascii=False,
    indent=2
)}

其他视频共同规律：

{json.dumps(
    comparison_summary,
    ensure_ascii=False,
    indent=2
)}

==================================================

硬性拍摄限制：

真人不露脸。

禁止正脸。

禁止主播面对镜头讲话。

允许：

- 手
- 手臂
- 少量身体局部
- 少量背影

拍摄视角只允许：

第一人称 POV

第三人称手部/局部视角

默认优先第一人称。

==================================================

可使用场景：

{scene_names}

recommended_scene必须从上面场景名称选择。

==================================================

生成3个明显不同的创意方向。

不能只是换一句文案。

至少在3个维度形成差异：

- 前3秒Hook
- 痛点
- 手部动作
- Demo顺序
- 产品切入
- 场景
- 情绪
- CTA

每个方向输出：

direction_name
core_idea
target_audience
hook
product_entry
recommended_perspective
recommended_scene
absorb_points
differentiation_points

全部中文。

本阶段不要输出逐秒脚本。

严格按JSON Schema。
""".strip()


def build_final_script_prompt(
    category,
    product_name,
    selected_video,
    effective_selling_points,
    chosen_direction,
    selected_scene,
    selected_perspective,
):

    return f"""
你是美国 TikTok Shop 中文拍摄SOP导演。

这是给中国拍摄团队直接执行的脚本。

非常重要：

除 copy_en 字段之外，
所有字段必须使用简体中文。

包括：

shooting_notes
shot
visual
hand_action
copy_cn
audio
absorb_point
difference_point
rationale

全部必须中文。

只有：

copy_en

使用自然美国英语。

==================================================

产品：

{category}

{product_name}

最终确认卖点：

{effective_selling_points}

主参考视频：

{json.dumps(
    selected_video,
    ensure_ascii=False,
    indent=2
)}

最终选择方向：

{json.dumps(
    chosen_direction,
    ensure_ascii=False,
    indent=2
)}

实际场景：

{selected_scene}

场景说明：

{SCENE_LIBRARY[selected_scene]}

实际视角：

{selected_perspective}

==================================================

真人限制：

不露脸。

不允许正脸。

不允许主播对镜头讲话。

允许：

- 手
- 手臂
- 身体局部
- 少量背影

==================================================

视频时长：

15-40秒。

要求：

- 0-3秒必须强Hook
- 前3秒必须有明显动作
- 产品尽快出现
- 必须真实UGC
- 必须可以在民宿直接拍
- 不需要专业摄影设备
- 每个动作必须具体
- 不要写抽象脚本

例如：

错误：

“展示印章效果”

正确：

“手机固定俯拍，左手压住快递标签右侧，
右手从画面右边拿起印章，
从姓名和地址区域由左向右滚一次，
镜头不断，让观众完整看到文字被遮盖。”

==================================================

每个分镜输出：

sequence

time_range

shot
中文

visual
中文

hand_action
中文

copy_cn
中文口播/字幕参考

copy_en
真正用于美国TikTok的英文字幕/口播

audio
中文说明

absorb_point
中文

difference_point
中文

rationale
中文

==================================================

英文 copy_en 要求：

- 美国口语
- 短句
- TikTok风格
- 不像广告说明书

严格按JSON Schema。
""".strip()


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
你是美国 TikTok Shop 中文视频复盘负责人。

除了 optimized_script 中的 copy_en，
其余所有内容必须使用简体中文。

产品：

{category}

{product_name}

卖点：

{selling_points}

流量：

{traffic_type}

原始脚本：

{original_script}

实际数据：

{json.dumps(
    metrics,
    ensure_ascii=False
)}

账号基准：

{json.dumps(
    baseline,
    ensure_ascii=False
)}

系统初判：

{json.dumps(
    assessment,
    ensure_ascii=False
)}

==================================================

漏斗：

前3秒留存
→ Hook

平均完播率
→ 中段节奏

互动率
→ 共鸣

商品CTR
→ 商品兴趣

订单转化
→ 成交

ROAS/CPC
→ 付费流量

==================================================

只找最优先修复的一个问题。

有账号基准优先账号基准。

没有再参考内部SOP。

优化版：

- 不露脸
- 第一人称优先
- 强手部动作
- 可直接拍摄
- 中文执行脚本
- copy_en保留美国英语

严格按JSON Schema。
""".strip()


# ============================================================
# 16. AI CALLS
# ============================================================

def analyze_videos(
    client,
    uploaded_videos,
    category,
    product_name,
    input_selling_points,
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
            f"单次最多上传{MAX_COMPARE_VIDEOS}条视频。"
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
            input_selling_points,
        )
    )

    remote_files = []
    temp_paths = []

    try:

        parts = []

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

                with tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=suffix,
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
                        file_uri=remote_file.uri,
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

        config = types.GenerateContentConfig(

            system_instruction=(
                "你必须严格遵守语言规则："
                "本次爆款拆解所有分析字段必须使用简体中文。"
                "不得因为视频是美国TikTok内容就改用英文分析。"
            ),

            thinking_config=(
                thinking_config()
            ),

            max_output_tokens=6200,

            response_mime_type="application/json",

            response_json_schema=(
                VIDEO_ANALYSIS_SCHEMA
            ),
        )

        response, meta = (
            generate_resilient(
                client,
                content,
                config,
            )
        )

        result = (
            parse_json_output(
                response.text
            )
        )

        return (
            result,
            {

                "analysis_mode":
                    analysis_mode,

                "total_size_mb":
                    round(
                        total_mb,
                        2,
                    ),

                "video_count":
                    len(
                        uploaded_videos
                    ),

                "analysis_seconds":
                    round(
                        time.perf_counter()
                        - started,
                        1,
                    ),

                **meta,
            },
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


def generate_directions(
    client,
    category,
    product_name,
    comparison_summary,
    selected_video,
    input_selling_points,
    effective_selling_points,
    selling_point_mode,
):

    started = (
        time.perf_counter()
    )

    prompt = (
        build_directions_prompt(
            category,
            product_name,
            comparison_summary,
            selected_video,
            input_selling_points,
            effective_selling_points,
            selling_point_mode,
        )
    )

    config = types.GenerateContentConfig(

        system_instruction=(
            "除必要英文产品名之外，"
            "三个创意方向全部必须使用简体中文。"
        ),

        thinking_config=(
            thinking_config()
        ),

        max_output_tokens=3200,

        response_mime_type="application/json",

        response_json_schema=(
            DIRECTIONS_SCHEMA
        ),
    )

    response, meta = (
        generate_resilient(
            client,
            prompt,
            config,
        )
    )

    result = (
        parse_json_output(
            response.text
        )
    )

    return (
        result,
        {

            "analysis_seconds":
                round(
                    time.perf_counter()
                    - started,
                    1,
                ),

            **meta,
        },
    )


def generate_final_script(
    client,
    category,
    product_name,
    selected_video,
    effective_selling_points,
    chosen_direction,
    selected_scene,
    selected_perspective,
):

    started = (
        time.perf_counter()
    )

    prompt = (
        build_final_script_prompt(
            category,
            product_name,
            selected_video,
            effective_selling_points,
            chosen_direction,
            selected_scene,
            selected_perspective,
        )
    )

    config = types.GenerateContentConfig(

        system_instruction=(
            "这是给中国团队执行的拍摄脚本。"
            "除copy_en字段之外，"
            "所有字段必须使用简体中文。"
            "copy_en才使用自然美国英语。"
        ),

        thinking_config=(
            thinking_config()
        ),

        max_output_tokens=4800,

        response_mime_type="application/json",

        response_json_schema=(
            FINAL_SCRIPT_SCHEMA
        ),
    )

    response, meta = (
        generate_resilient(
            client,
            prompt,
            config,
        )
    )

    result = (
        parse_json_output(
            response.text
        )
    )

    return (
        result,
        {

            "analysis_seconds":
                round(
                    time.perf_counter()
                    - started,
                    1,
                ),

            **meta,
        },
    )


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

    prompt = (
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
    )

    config = types.GenerateContentConfig(

        system_instruction=(
            "这是给中国运营团队看的复盘。"
            "除copy_en外全部使用简体中文。"
        ),

        thinking_config=(
            thinking_config()
        ),

        max_output_tokens=4400,

        response_mime_type="application/json",

        response_json_schema=(
            REVIEW_SCHEMA
        ),
    )

    response, meta = (
        generate_resilient(
            client,
            prompt,
            config,
        )
    )

    result = (
        parse_json_output(
            response.text
        )
    )

    return (
        result,
        {

            "analysis_seconds":
                round(
                    time.perf_counter()
                    - started,
                    1,
                ),

            **meta,
        },
    )


# ============================================================
# 17. REVIEW HELPERS
# ============================================================

def compare_metric(
    key,
    value,
    baseline,
):

    if value is None:

        return {
            "status":
                "未填写"
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

            "value":
                value,

            "baseline":
                baseline,

            "ratio":
                round(
                    ratio,
                    3,
                ),

            "status":
                status,
        }

    band = (
        SOP_BANDS[
            key
        ]
    )

    if value < band["low"]:

        status = "偏低"

    elif value >= band["high"]:

        status = "较强"

    else:

        status = "中间区间"

    return {

        "value":
            value,

        "status":
            status,

        "basis":
            "内部SOP工作区间",
    }


def build_metric_assessment(
    metrics,
    baseline,
):

    return {

        "前3秒留存":
            compare_metric(
                "retention_3s_pct",
                metrics.get(
                    "retention_3s_pct"
                ),
                baseline.get(
                    "retention_3s_pct"
                ),
            ),

        "完播率":
            compare_metric(
                "completion_rate_pct",
                metrics.get(
                    "completion_rate_pct"
                ),
                baseline.get(
                    "completion_rate_pct"
                ),
            ),

        "商品CTR":
            compare_metric(
                "product_ctr_pct",
                metrics.get(
                    "product_ctr_pct"
                ),
                baseline.get(
                    "product_ctr_pct"
                ),
            ),

        "订单转化率":
            compare_metric(
                "order_conversion_pct",
                metrics.get(
                    "order_conversion_pct"
                ),
                baseline.get(
                    "order_conversion_pct"
                ),
            ),

        "互动率":
            compare_metric(
                "engagement_rate_pct",
                metrics.get(
                    "engagement_rate_pct"
                ),
                baseline.get(
                    "engagement_rate_pct"
                ),
            ),
    }


def render_optional_float_input(
    label,
    key,
):

    raw = (
        st.text_input(
            label,
            key=key,
            placeholder="可留空",
        )
    )

    value = (
        parse_optional_float(
            raw
        )
    )

    if (
        clean_text(
            raw
        )
        and value is None
    ):

        st.caption(
            "请输入数字，例如：12.5"
        )

    return value


def get_script_sheets(
    uploaded_file,
):

    data = (
        uploaded_file.getvalue()
    )

    excel = pd.ExcelFile(
        io.BytesIO(
            data
        ),
        engine="openpyxl",
    )

    result = []

    for sheet in excel.sheet_names:

        try:

            dataframe = pd.read_excel(
                io.BytesIO(
                    data
                ),
                sheet_name=sheet,
                nrows=3,
                engine="openpyxl",
            )

            columns = set(
                dataframe.columns.tolist()
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
                OLD_SCRIPT_COLUMNS
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

    dataframe = pd.read_excel(
        io.BytesIO(
            uploaded_file.getvalue()
        ),
        sheet_name=sheet_name,
        engine="openpyxl",
    )

    blocks = []

    for _, row in dataframe.iterrows():

        block = f"""
分镜：{clean_text(row.get("分镜序号"))}
时间：{clean_text(row.get("时间段"))}
机位/视角：{clean_text(row.get("机位/视角", row.get("景别/机位", "")))}
画面：{clean_text(row.get("画面描述(道具/动作)"))}
手部动作：{clean_text(row.get("手部动作"))}
中文口播/字幕参考：{clean_text(row.get("中文口播/字幕参考"))}
英文口播/字幕：{clean_text(row.get("英文口播/字幕", row.get("英文口播文案/字幕", "")))}
音效/节奏：{clean_text(row.get("音效/节奏提示"))}
爆款吸收点：{clean_text(row.get("爆款吸收点"))}
差异化处理：{clean_text(row.get("差异化处理"))}
设计目的：{clean_text(row.get("设计目的(底层逻辑)"))}
""".strip()

        blocks.append(
            block
        )

    return "\n\n".join(
        blocks
    )


# ============================================================
# 18. INIT
# ============================================================

render_login_sidebar()

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


api_key = (
    get_api_key()
)


if not api_key:

    st.error(
        "系统未配置 Gemini API Key，请联系管理员。"
    )

    client = None

else:

    client = (
        create_client(
            api_key
        )
    )


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
# 19. 爆款拆解
# ============================================================

with tab_analysis:

    # --------------------------------------------------------
    # ① 产品
    # --------------------------------------------------------

    st.markdown(
        "### ① 产品信息"
    )

    c1, c2, c3 = (
        st.columns(
            3
        )
    )

    with c1:

        tiktok_account = (
            st.text_input(
                "TikTok账号",
                key="analysis_account",
                placeholder="用于历史归档",
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
                key="analysis_product_name",
                placeholder="例如：隐私印章",
            )
        )

    input_selling_points = (
        st.text_area(
            "我们的真实产品卖点（选填）",
            key="analysis_user_selling_points",
            height=90,
            placeholder=(
                "可不填。"
                "填写后AI只负责比较和推荐，"
                "最终使用哪套卖点由你自己选择。"
            ),
        )
    )

    # --------------------------------------------------------
    # ② 视频
    # --------------------------------------------------------

    st.markdown(
        "### ② 上传爆款视频"
    )

    uploaded_videos = (
        st.file_uploader(
            "支持同时上传 1-5 条 .mp4 视频",
            type=["mp4"],
            accept_multiple_files=True,
            key="analysis_videos",
        )
    )

    if uploaded_videos:

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
            f"已上传 {len(uploaded_videos)} 条 "
            f"· 总大小 {total_mb:.2f} MB"
        )

        signature = (
            video_batch_signature(
                uploaded_videos,
                category,
                product_name,
                input_selling_points,
            )
        )

        if (
            st.session_state[
                "video_batch_signature"
            ]
            != signature
        ):

            st.session_state[
                "video_batch_signature"
            ] = signature

            st.session_state[
                "video_analysis_result"
            ] = None

            st.session_state[
                "directions_result"
            ] = None

            st.session_state[
                "directions_context_signature"
            ] = ""

            st.session_state[
                "final_script_result"
            ] = None

            st.session_state[
                "final_script_context_signature"
            ] = ""

            st.session_state[
                "selected_reference_video_index"
            ] = None

            st.session_state[
                "selected_direction_index"
            ] = 0

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

            key="run_video_analysis",
        )
    )

    if analyze_button:

        try:

            with st.spinner(
                "正在逐条拆解视频并提取独立卖点…"
            ):

                (
                    result,
                    metadata,
                ) = analyze_videos(
                    client,
                    uploaded_videos,
                    category,
                    product_name,
                    input_selling_points,
                )

            st.session_state[
                "video_analysis_result"
            ] = result

            st.session_state[
                "video_analysis_meta"
            ] = metadata

            recommended = safe_int(
                result.get(
                    "recommended_reference_video_index",
                    1,
                ),
                1,
            )

            st.session_state[
                "selected_reference_video_index"
            ] = recommended

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

                    "input_selling_points":
                        input_selling_points,

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
                        metadata.get(
                            "model_used",
                            "",
                        ),

                    "fallback_used":
                        metadata.get(
                            "fallback_used",
                            "",
                        ),

                    "retry_count":
                        metadata.get(
                            "retry_count",
                            "",
                        ),

                    "analysis_seconds":
                        metadata.get(
                            "analysis_seconds",
                            "",
                        ),

                    "full_output_json":
                        json_dumps(
                            result
                        ),
                }
            )

            st.success(
                "爆款视频解析完成。"
            )

        except Exception as exc:

            st.error(
                friendly_error(
                    exc
                )
            )

    analysis_result = (
        st.session_state.get(
            "video_analysis_result"
        )
    )

    # --------------------------------------------------------
    # ③ 拆解
    # --------------------------------------------------------

    if analysis_result:

        summary = (
            analysis_result.get(
                "comparison_summary",
                {},
            )
        )

        videos = (
            analysis_result.get(
                "videos",
                [],
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
                "**共同爆款脚本路线**"
            )

            st.write(
                summary.get(
                    "common_script_route",
                    "",
                )
            )

            detail1, detail2 = (
                st.columns(
                    2
                )
            )

            with detail1:

                st.markdown(
                    "**共同人群画像**"
                )

                st.write(
                    summary.get(
                        "common_audience",
                        "",
                    )
                )

            with detail2:

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
                "**共同前3秒 Hook**"
            )

            st.write(
                summary.get(
                    "common_hook_pattern",
                    "",
                )
            )

            st.markdown(
                "**共同画面与节奏**"
            )

            st.write(
                summary.get(
                    "visual_rhythm",
                    "",
                )
            )

            st.markdown(
                "**最值得共同吸收的3点**"
            )

            for item in summary.get(
                "top_absorb_points",
                [],
            ):

                st.markdown(
                    f"- {item}"
                )

            if len(videos) > 1:

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

            for video in videos:

                with st.expander(
                    (
                        f'视频 {video.get("video_index", "")}'
                        f'｜{video.get("filename", "")}'
                    ),
                    expanded=False,
                ):

                    st.markdown(
                        f'**一句话核心：** '
                        f'{video.get("one_sentence_core", "")}'
                    )

                    st.markdown(
                        "**该视频独立卖点**"
                    )

                    for item in video.get(
                        "inferred_selling_points",
                        [],
                    ):

                        st.markdown(
                            f"- {item}"
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
                        "**人群画像**"
                    )

                    st.write(
                        video.get(
                            "audience_profile",
                            "",
                        )
                    )

                    st.markdown(
                        "**年龄预估**"
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

                    for item in video.get(
                        "top_absorb_points",
                        [],
                    ):

                        st.markdown(
                            f"- {item}"
                        )

                    st.markdown(
                        "**参考价值判断**"
                    )

                    st.write(
                        video.get(
                            "fit_reason",
                            "",
                        )
                    )

                    st.markdown(
                        f'**推荐指数：** '
                        f'{video.get("recommend_score", "")}'
                    )

        # ----------------------------------------------------
        # ④ 主参考视频
        # ----------------------------------------------------

        st.markdown(
            "### ④ 选择主参考视频"
        )

        recommended_index = (
            safe_int(
                analysis_result.get(
                    "recommended_reference_video_index",
                    1,
                ),
                1,
            )
        )

        video_indices = [

            safe_int(
                video.get(
                    "video_index"
                ),
                index + 1,
            )

            for index, video
            in enumerate(
                videos
            )
        ]

        if (
            st.session_state[
                "selected_reference_video_index"
            ]
            not in video_indices
        ):

            st.session_state[
                "selected_reference_video_index"
            ] = recommended_index

        selected_ref_index = (
            st.radio(
                "请选择本次主要参考的视频",

                options=video_indices,

                format_func=lambda value: (
                    f'视频{value}｜'
                    f'{next((v.get("filename", "") for v in videos if safe_int(v.get("video_index"), 0) == value), "")}'
                    + (
                        "（AI推荐）"
                        if value
                        == recommended_index
                        else ""
                    )
                ),

                key="selected_reference_video_index",
            )
        )

        chosen_ref_video = next(
            (
                video
                for video
                in videos
                if safe_int(
                    video.get(
                        "video_index"
                    ),
                    0,
                )
                == selected_ref_index
            ),
            videos[0],
        )

        st.caption(
            "AI推荐依据："
        )

        st.write(
            chosen_ref_video.get(
                "fit_reason",
                "",
            )
        )

        # ----------------------------------------------------
        # ⑤ 卖点选择
        # ----------------------------------------------------

        st.markdown(
            "### ⑤ 卖点参考逻辑"
        )

        selected_video_points = (
            chosen_ref_video.get(
                "inferred_selling_points",
                [],
            )
        )

        if not selected_video_points:

            selected_video_points = (
                analysis_result.get(
                    "common_inferred_selling_points",
                    [],
                )
            )

        viral_points_text = (
            list_to_joined(
                selected_video_points
            )
        )

        st.markdown(
            f'**当前参考：视频{selected_ref_index}'
            f'｜{chosen_ref_video.get("filename", "")}**'
        )

        st.markdown(
            "**该视频推理出的核心卖点：**"
        )

        for item in selected_video_points:

            st.markdown(
                f"- {item}"
            )

        # ----------------------------------------------------
        # 没有人工卖点
        # ----------------------------------------------------

        if not clean_text(
            input_selling_points
        ):

            st.info(
                "你没有填写自己的产品卖点，"
                "因此当前使用该爆款视频推理卖点。"
            )

            selling_point_mode = (
                "viral_first"
            )

            default_effective_points = (
                viral_points_text
            )

        # ----------------------------------------------------
        # 有人工卖点：永远保留选择权
        # ----------------------------------------------------

        else:

            relation = (
                clean_text(
                    chosen_ref_video.get(
                        "selling_point_relation",
                        "",
                    )
                )
            )

            relation_reason = (
                clean_text(
                    chosen_ref_video.get(
                        "selling_point_relation_reason",
                        "",
                    )
                )
            )

            suggested_mode = (
                clean_text(
                    chosen_ref_video.get(
                        "suggested_mode",
                        "blend",
                    )
                )
            )

            if suggested_mode not in {
                "viral_first",
                "user_first",
                "blend",
            }:

                suggested_mode = (
                    "blend"
                )

            recommendation_label = {

                "viral_first":
                    "以当前爆款视频卖点为主",

                "user_first":
                    "以我的真实产品卖点为主",

                "blend":
                    "融合两者",

            }[
                suggested_mode
            ]

            st.info(
                f"AI推荐：{recommendation_label}"
            )

            if relation_reason:

                st.caption(
                    f"推荐原因：{relation_reason}"
                )

            # =================================================
            # 最重要：
            # AI不给你自动做主
            # 必须人工选择
            # =================================================

            selling_choice_key = (
                f"manual_selling_choice_"
                f"{selected_ref_index}"
            )

            selling_point_mode = (
                st.radio(
                    "请由使用人决定本次采用哪套卖点",

                    options=[
                        "viral_first",
                        "user_first",
                        "blend",
                    ],

                    index=None,

                    format_func=lambda mode: {

                        "viral_first":
                            "以当前爆款视频卖点为主",

                        "user_first":
                            "以我的真实产品卖点为主",

                        "blend":
                            "融合两者",

                    }[mode],

                    key=selling_choice_key,
                )
            )

            if selling_point_mode is None:

                default_effective_points = ""

            elif (
                selling_point_mode
                == "viral_first"
            ):

                default_effective_points = (
                    viral_points_text
                )

            elif (
                selling_point_mode
                == "user_first"
            ):

                default_effective_points = (
                    clean_text(
                        input_selling_points
                    )
                )

            else:

                default_effective_points = (
                    clean_text(
                        chosen_ref_video.get(
                            "blended_selling_points",
                            "",
                        )
                    )
                )

                if not default_effective_points:

                    default_effective_points = (
                        f"{viral_points_text}; "
                        f"{clean_text(input_selling_points)}"
                    )

        # ----------------------------------------------------
        # 最终卖点人工可编辑
        # ----------------------------------------------------

        if selling_point_mode is not None:

            effective_edit_key = (
                f"effective_points_edit_"
                f"{selected_ref_index}_"
                f"{selling_point_mode}"
            )

            if (
                effective_edit_key
                not in st.session_state
            ):

                st.session_state[
                    effective_edit_key
                ] = default_effective_points

            effective_selling_points = (
                st.text_area(
                    "最终用于生成脚本的卖点（可编辑）",
                    height=110,
                    key=effective_edit_key,
                    help=(
                        "这里是最终送给AI生成三个方向的卖点。"
                        "你可以继续人工修改。"
                    ),
                )
            )

        else:

            effective_selling_points = ""

            st.warning(
                "请先选择上面的卖点参考方式，"
                "再生成3个拍摄方向。"
            )

        current_direction_context = (
            make_signature(
                selected_ref_index,
                chosen_ref_video,
                selling_point_mode,
                effective_selling_points,
            )
        )

        old_direction_context = (
            st.session_state.get(
                "directions_context_signature",
                "",
            )
        )

        if (
            old_direction_context
            and old_direction_context
            != current_direction_context
        ):

            st.session_state[
                "directions_result"
            ] = None

            st.session_state[
                "directions_context_signature"
            ] = ""

            st.session_state[
                "final_script_result"
            ] = None

            st.session_state[
                "final_script_context_signature"
            ] = ""

            st.session_state[
                "selected_direction_index"
            ] = 0

        # ----------------------------------------------------
        # ⑥ 方向
        # ----------------------------------------------------

        st.markdown(
            "### ⑥ 生成 3 个参考方向"
        )

        generate_direction_button = (
            st.button(
                "生成3个拍摄方向",
                type="primary",
                use_container_width=True,

                disabled=(
                    client is None
                    or selling_point_mode is None
                    or not clean_text(
                        effective_selling_points
                    )
                ),

                key="generate_directions_button",
            )
        )

        if generate_direction_button:

            try:

                with st.spinner(
                    "正在生成3个不同拍摄方向…"
                ):

                    (
                        directions_result,
                        directions_meta,
                    ) = generate_directions(
                        client,
                        category,
                        product_name,
                        summary,
                        chosen_ref_video,
                        input_selling_points,
                        effective_selling_points,
                        selling_point_mode,
                    )

                st.session_state[
                    "directions_result"
                ] = directions_result

                st.session_state[
                    "directions_meta"
                ] = directions_meta

                st.session_state[
                    "directions_context_signature"
                ] = current_direction_context

                st.session_state[
                    "selected_direction_index"
                ] = 0

                st.session_state[
                    "final_script_result"
                ] = None

                append_history(
                    {

                        "record_type":
                            "3方向生成",

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

                        "input_selling_points":
                            input_selling_points,

                        "inferred_selling_points":
                            viral_points_text,

                        "effective_selling_points":
                            effective_selling_points,

                        "selling_point_mode":
                            selling_point_mode,

                        "reference_video_index":
                            selected_ref_index,

                        "reference_video_name":
                            chosen_ref_video.get(
                                "filename",
                                "",
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
                    "3个拍摄方向已生成。"
                )

            except Exception as exc:

                st.error(
                    friendly_error(
                        exc
                    )
                )

        directions_result = (
            st.session_state.get(
                "directions_result"
            )
        )

        # ----------------------------------------------------
        # ⑦ 方向选择
        # ----------------------------------------------------

        if directions_result:

            directions = (
                directions_result.get(
                    "directions",
                    [],
                )
            )

            if directions:

                st.markdown(
                    "### ⑦ 选择 1 个方向"
                )

                direction_indices = (
                    list(
                        range(
                            len(
                                directions
                            )
                        )
                    )
                )

                if (
                    st.session_state[
                        "selected_direction_index"
                    ]
                    not in direction_indices
                ):

                    st.session_state[
                        "selected_direction_index"
                    ] = 0

                selected_direction_index = (
                    st.radio(
                        "请选择你要继续生成最终脚本的方向",

                        options=direction_indices,

                        format_func=lambda index: (
                            f'方向{index + 1}｜'
                            f'{directions[index].get("direction_name", "")}'
                        ),

                        key="selected_direction_index",
                    )
                )

                chosen_direction = (
                    directions[
                        selected_direction_index
                    ]
                )

                # 不再使用 tabs
                # 选哪个只显示哪个

                st.divider()

                st.markdown(
                    f'### 方向{selected_direction_index + 1}'
                    f'｜{chosen_direction.get("direction_name", "")}'
                )

                st.markdown(
                    f'**核心思路：** '
                    f'{chosen_direction.get("core_idea", "")}'
                )

                d1, d2 = (
                    st.columns(
                        2
                    )
                )

                with d1:

                    st.markdown(
                        f'**目标人群：** '
                        f'{chosen_direction.get("target_audience", "")}'
                    )

                    st.markdown(
                        f'**前3秒 Hook：** '
                        f'{chosen_direction.get("hook", "")}'
                    )

                    st.markdown(
                        f'**产品切入方式：** '
                        f'{chosen_direction.get("product_entry", "")}'
                    )

                with d2:

                    st.markdown(
                        f'**推荐视角：** '
                        f'{chosen_direction.get("recommended_perspective", "")}'
                    )

                    st.markdown(
                        f'**推荐小场景：** '
                        f'{chosen_direction.get("recommended_scene", "")}'
                    )

                st.markdown(
                    "**可吸收点**"
                )

                for item in chosen_direction.get(
                    "absorb_points",
                    [],
                ):

                    st.markdown(
                        f"- {item}"
                    )

                st.markdown(
                    "**差异化点**"
                )

                for item in chosen_direction.get(
                    "differentiation_points",
                    [],
                ):

                    st.markdown(
                        f"- {item}"
                    )

                # ------------------------------------------------
                # 方向改变后刷新场景/视角
                # ------------------------------------------------

                control_signature = (
                    make_signature(
                        current_direction_context,
                        selected_direction_index,
                        chosen_direction,
                    )
                )

                if (
                    st.session_state.get(
                        "last_direction_control_signature",
                        "",
                    )
                    != control_signature
                ):

                    recommended_perspective = (
                        clean_text(
                            chosen_direction.get(
                                "recommended_perspective",
                                "",
                            )
                        )
                    )

                    if (
                        "第三"
                        in recommended_perspective
                    ):

                        st.session_state[
                            "final_selected_perspective"
                        ] = PERSPECTIVE_OPTIONS[1]

                    else:

                        st.session_state[
                            "final_selected_perspective"
                        ] = PERSPECTIVE_OPTIONS[0]

                    recommended_scene = (
                        clean_text(
                            chosen_direction.get(
                                "recommended_scene",
                                "",
                            )
                        )
                    )

                    if (
                        recommended_scene
                        in SCENE_LIBRARY
                    ):

                        st.session_state[
                            "final_selected_scene"
                        ] = recommended_scene

                    else:

                        st.session_state[
                            "final_selected_scene"
                        ] = list(
                            SCENE_LIBRARY.keys()
                        )[0]

                    st.session_state[
                        "last_direction_control_signature"
                    ] = control_signature

                    st.session_state[
                        "final_script_result"
                    ] = None

                # ------------------------------------------------
                # ⑧ 最终脚本
                # ------------------------------------------------

                st.markdown(
                    "### ⑧ 生成最终拍摄脚本"
                )

                script1, script2 = (
                    st.columns(
                        2
                    )
                )

                with script1:

                    selected_perspective = (
                        st.radio(
                            "拍摄视角",
                            PERSPECTIVE_OPTIONS,
                            key="final_selected_perspective",
                        )
                    )

                with script2:

                    selected_scene = (
                        st.selectbox(
                            "实际拍摄小场景",
                            list(
                                SCENE_LIBRARY.keys()
                            ),
                            key="final_selected_scene",
                        )
                    )

                st.caption(
                    "固定限制：真人不露脸 / 不正面出镜；"
                    "允许手、手臂和少量身体局部。"
                )

                current_final_context = (
                    make_signature(
                        current_direction_context,
                        selected_direction_index,
                        selected_perspective,
                        selected_scene,
                    )
                )

                old_final_context = (
                    st.session_state.get(
                        "final_script_context_signature",
                        "",
                    )
                )

                if (
                    old_final_context
                    and old_final_context
                    != current_final_context
                ):

                    st.session_state[
                        "final_script_result"
                    ] = None

                    st.session_state[
                        "final_script_context_signature"
                    ] = ""

                generate_script_button = (
                    st.button(
                        "生成最终拍摄脚本",
                        type="primary",
                        use_container_width=True,
                        disabled=(
                            client is None
                        ),
                        key="generate_final_script_button",
                    )
                )

                if generate_script_button:

                    try:

                        with st.spinner(
                            "正在生成中文可执行拍摄脚本…"
                        ):

                            (
                                final_result,
                                final_meta,
                            ) = generate_final_script(
                                client,
                                category,
                                product_name,
                                chosen_ref_video,
                                effective_selling_points,
                                chosen_direction,
                                selected_scene,
                                selected_perspective,
                            )

                        st.session_state[
                            "final_script_result"
                        ] = final_result

                        st.session_state[
                            "final_script_meta"
                        ] = final_meta

                        st.session_state[
                            "final_script_context_signature"
                        ] = current_final_context

                        append_history(
                            {

                                "record_type":
                                    "最终拍摄脚本",

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

                                "input_selling_points":
                                    input_selling_points,

                                "inferred_selling_points":
                                    viral_points_text,

                                "effective_selling_points":
                                    effective_selling_points,

                                "selling_point_mode":
                                    selling_point_mode,

                                "reference_video_index":
                                    selected_ref_index,

                                "reference_video_name":
                                    chosen_ref_video.get(
                                        "filename",
                                        "",
                                    ),

                                "direction_name":
                                    chosen_direction.get(
                                        "direction_name",
                                        "",
                                    ),

                                "selected_scene":
                                    selected_scene,

                                "selected_perspective":
                                    selected_perspective,

                                "model_used":
                                    final_meta.get(
                                        "model_used",
                                        "",
                                    ),

                                "fallback_used":
                                    final_meta.get(
                                        "fallback_used",
                                        "",
                                    ),

                                "retry_count":
                                    final_meta.get(
                                        "retry_count",
                                        "",
                                    ),

                                "analysis_seconds":
                                    final_meta.get(
                                        "analysis_seconds",
                                        "",
                                    ),

                                "full_output_json":
                                    json_dumps(
                                        final_result
                                    ),
                            }
                        )

                        st.success(
                            "最终拍摄脚本已生成。"
                        )

                    except Exception as exc:

                        st.error(
                            friendly_error(
                                exc
                            )
                        )

        # ----------------------------------------------------
        # ⑨ 最终脚本
        # ----------------------------------------------------

        final_script_result = (
            st.session_state.get(
                "final_script_result"
            )
        )

        if final_script_result:

            st.markdown(
                "### ⑨ 最终拍摄脚本"
            )

            shooting_notes = (
                clean_text(
                    final_script_result.get(
                        "shooting_notes",
                        "",
                    )
                )
            )

            if shooting_notes:

                st.info(
                    shooting_notes
                )

            final_dataframe = (
                final_script_to_df(
                    final_script_result
                )
            )

            st.dataframe(
                final_dataframe,
                hide_index=True,
                use_container_width=True,
            )

            excel_data = (
                build_analysis_export_excel(
                    analysis_result,
                    st.session_state.get(
                        "directions_result"
                    ),
                    final_script_result,
                )
            )

            st.download_button(
                "一键导出 Excel",

                data=excel_data,

                file_name=(
                    "TikTok爆款解析_拍摄脚本_"
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
# 20. 数据复盘
# ============================================================

with tab_review:

    st.markdown(
        "### ① 上传原始脚本"
    )

    review_excel = (
        st.file_uploader(
            "上传之前导出的 Excel",
            type=["xlsx"],
            key="review_excel",
        )
    )

    if review_excel:

        try:

            sheets = (
                get_script_sheets(
                    review_excel
                )
            )

            if sheets:

                selected_sheet = (
                    st.selectbox(
                        "选择本次实际发布的脚本",
                        sheets,
                        key="review_selected_sheet",
                    )
                )

                if st.button(
                    "读取该脚本",
                    use_container_width=True,
                    key="read_review_script",
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
                        "脚本已读取。"
                    )

            else:

                st.warning(
                    "未识别到可复盘脚本。"
                )

        except Exception as exc:

            st.error(
                f"Excel读取失败：{exc}"
            )

    meta1, meta2, meta3 = (
        st.columns(
            3
        )
    )

    with meta1:

        review_account = (
            st.text_input(
                "TikTok账号",
                key="review_account",
            )
        )

    with meta2:

        review_category = (
            st.selectbox(
                "产品品类",
                PRODUCT_CATEGORIES,
                key="review_category",
            )
        )

    with meta3:

        review_product_name = (
            st.text_input(
                "产品名称 / SKU",
                key="review_product_name",
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

    st.markdown(
        "### ② 核心数据"
    )

    metric1, metric2, metric3 = (
        st.columns(
            3
        )
    )

    with metric1:

        retention = (
            render_optional_float_input(
                "前3秒留存率 (%)",
                "review_retention",
            )
        )

        completion = (
            render_optional_float_input(
                "平均完播率 (%)",
                "review_completion",
            )
        )

    with metric2:

        ctr = (
            render_optional_float_input(
                "商品锚点 CTR (%)",
                "review_ctr",
            )
        )

        conversion = (
            render_optional_float_input(
                "订单转化率 (%)",
                "review_conversion",
            )
        )

    with metric3:

        engagement = (
            render_optional_float_input(
                "互动率 (%)",
                "review_engagement",
            )
        )

    with st.expander(
        "广告数据（选填）",
        expanded=False,
    ):

        ad1, ad2 = (
            st.columns(
                2
            )
        )

        with ad1:

            actual_roas = (
                render_optional_float_input(
                    "实际 ROAS",
                    "review_actual_roas",
                )
            )

            target_roas = (
                render_optional_float_input(
                    "目标 ROAS",
                    "review_target_roas",
                )
            )

        with ad2:

            actual_cpc = (
                render_optional_float_input(
                    "实际 CPC ($)",
                    "review_actual_cpc",
                )
            )

            target_cpc = (
                render_optional_float_input(
                    "目标 CPC ($)",
                    "review_target_cpc",
                )
            )

    with st.expander(
        "账号近7天基准（建议填写）",
        expanded=False,
    ):

        base1, base2, base3 = (
            st.columns(
                3
            )
        )

        with base1:

            base_retention = (
                render_optional_float_input(
                    "账号平均3秒留存 (%)",
                    "base_retention",
                )
            )

            base_completion = (
                render_optional_float_input(
                    "账号平均完播率 (%)",
                    "base_completion",
                )
            )

        with base2:

            base_ctr = (
                render_optional_float_input(
                    "账号平均CTR (%)",
                    "base_ctr",
                )
            )

            base_conversion = (
                render_optional_float_input(
                    "账号平均转化率 (%)",
                    "base_conversion",
                )
            )

        with base3:

            base_engagement = (
                render_optional_float_input(
                    "账号平均互动率 (%)",
                    "base_engagement",
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

        if not clean_text(
            original_script
        ):

            st.error(
                "请先上传或填写原始脚本。"
            )

        elif not clean_text(
            review_selling_points
        ):

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
                    "正在进行数据复盘…"
                ):

                    (
                        review_result,
                        review_meta,
                    ) = review_script(
                        client,
                        review_category,
                        review_product_name,
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
                            review_product_name,

                        "effective_selling_points":
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
                    "复盘完成。"
                )

            except Exception as exc:

                st.error(
                    friendly_error(
                        exc
                    )
                )

    review_result = (
        st.session_state.get(
            "review_result"
        )
    )

    if review_result:

        st.markdown(
            "### ③ 复盘结果"
        )

        st.info(
            review_result.get(
                "diagnosis_summary",
                "",
            )
        )

        st.markdown(
            f'**最优先修复：** '
            f'{review_result.get("priority_issue", "")}'
        )

        with st.expander(
            "查看逐项指标诊断",
            expanded=False,
        ):

            diagnosis_dataframe = (
                pd.DataFrame(
                    review_result.get(
                        "metric_diagnosis",
                        [],
                    )
                )
            )

            if not diagnosis_dataframe.empty:

                st.dataframe(
                    diagnosis_dataframe,
                    hide_index=True,
                    use_container_width=True,
                )

        st.markdown(
            "### ④ 优化版脚本"
        )

        review_dataframe = (
            review_script_to_df(
                review_result
            )
        )

        st.dataframe(
            review_dataframe,
            hide_index=True,
            use_container_width=True,
        )

        st.download_button(
            "下载优化版 Excel",

            data=(
                build_review_export_excel(
                    review_result
                )
            ),

            file_name=(
                "TikTok优化版脚本_"
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
# 21. HISTORY
# ============================================================

with tab_history:

    history = (
        scoped_history()
    )

    if history.empty:

        st.info(
            "暂无历史记录。"
        )

    else:

        filtered = (
            history.copy()
        )

        h1, h2, h3 = (
            st.columns(
                3
            )
        )

        with h1:

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
                            value
                            for value
                            in filtered[
                                "operator"
                            ].unique()
                            if value
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

        with h2:

            record_types = (
                ["全部"]
                + sorted(
                    [
                        value
                        for value
                        in filtered[
                            "record_type"
                        ].unique()
                        if value
                    ]
                )
            )

            type_filter = (
                st.selectbox(
                    "类型",
                    record_types,
                    key="history_type",
                )
            )

            if (
                type_filter
                != "全部"
            ):

                filtered = filtered[
                    filtered[
                        "record_type"
                    ]
                    == type_filter
                ]

        with h3:

            account_options = (
                ["全部"]
                + sorted(
                    [
                        value
                        for value
                        in filtered[
                            "tiktok_account"
                        ].unique()
                        if value
                    ]
                )
            )

            account_filter = (
                st.selectbox(
                    "TikTok账号",
                    account_options,
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
            "reference_video_name",
            "direction_name",
            "selected_scene",
            "selected_perspective",
            "priority_issue",
        ]

        available_columns = [
            column
            for column
            in display_columns
            if column
            in filtered.columns
        ]

        st.dataframe(
            filtered[
                available_columns
            ].iloc[::-1],
            hide_index=True,
            use_container_width=True,
        )

        st.download_button(
            "下载历史 CSV",

            data=(
                filtered
                .to_csv(
                    index=False,
                    encoding="utf-8-sig",
                )
                .encode(
                    "utf-8-sig"
                )
            ),

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
