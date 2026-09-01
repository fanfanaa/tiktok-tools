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

# ------------------------------------------------------------
# Gemini 模型策略
# ------------------------------------------------------------
# 主模型：最低延迟、高吞吐
PRIMARY_MODEL = "gemini-3.5-flash-lite"

# 主模型持续 503 时自动切换
FALLBACK_MODELS = [
    "gemini-3.1-flash-lite",
    "gemini-3.6-flash",
]

MODEL_CHAIN = [
    PRIMARY_MODEL,
    *FALLBACK_MODELS,
]

# 每个模型最多尝试次数
MAX_ATTEMPTS_PER_MODEL = 2

# 20-40 秒常规视频优先 Inline
INLINE_VIDEO_MAX_MB = 18

# Files API 大文件回退
FILE_POLL_INTERVAL_SEC = 2
FILE_PROCESS_TIMEOUT_SEC = 180

# 登录默认密码
DEFAULT_STAFF_PASSWORD = "8888"
DEFAULT_ADMIN_PASSWORD = "8888-admin"

# 历史记录
HISTORY_FILE = Path("history_log.csv")
HISTORY_LOCK = threading.Lock()


# ============================================================
# 2. 民宿拍摄场景
# ============================================================

SCENE_LIBRARY = {
    "民宿客厅·沙发/茶几休闲区": {
        "prompt": (
            "固定在美国民宿风格客厅。利用沙发、茶几、边几完成真实生活流。"
            "产品自然进入使用动作，不做僵硬棚拍展示。"
        ),
        "guide": (
            "推荐机位：手持 POV、茶几 45°俯拍、沙发侧前方中近景、产品微距。\n\n"
            "执行重点：第一帧优先展示痛点、反差、结果或最强产品动作；"
            "茶几只保留必要道具。"
        ),
    },

    "民宿厨房·多功能岛台": {
        "prompt": (
            "固定在民宿厨房岛台或操作台。"
            "核心是连续手部动作、真实痛点、产品解决过程和结果证明。"
        ),
        "guide": (
            "推荐机位：胸口 POV、岛台 45°俯拍、手部超近距、侧前方固定机位。\n\n"
            "适合：挂式厨房垃圾桶、厨房工具、清洁、收纳类产品。\n"
            "执行重点：台面只保留必要道具。"
        ),
    },

    "民宿卧室·床头/梳妆台": {
        "prompt": (
            "固定在民宿卧室。利用床头柜、床面或梳妆台，"
            "呈现真实拿取、使用、体验和放回动作。"
        ),
        "guide": (
            "推荐机位：床头近景、梳妆台 45°、第一人称拿取、床面俯拍。\n\n"
            "适合：耳机、阅读用品、便携用品、个人护理类产品。"
        ),
    },

    "民宿卫生间/浴室·镜前特写": {
        "prompt": (
            "固定在卫生间、浴室、洗手台或镜柜区域。"
            "先放大真实痛点，再进入产品解决动作。"
        ),
        "guide": (
            "推荐机位：镜前中近景、洗手台 45°、手部微距、镜柜第一人称。\n\n"
            "执行重点：避免镜中穿帮、严重反光和杂乱洗护用品。"
        ),
    },

    "民宿阳台/落地窗·自然光": {
        "prompt": (
            "固定在阳台、落地窗或窗边自然光区域。"
            "利用自然光展示产品外观、材质、便携性和真实生活方式。"
        ),
        "guide": (
            "推荐机位：窗边侧光、手持 POV、产品近景、中景生活流。\n\n"
            "执行重点：避免强逆光导致产品主体过暗。"
        ),
    },

    "纯桌面特写·无真人露脸": {
        "prompt": (
            "视频只允许出现桌面、产品、双手和必要道具。"
            "重点展示功能动作、细节、Before/After 和结果证明。"
        ),
        "guide": (
            "推荐机位：正上方俯拍、45°俯拍、超近距功能特写、前后对比。\n\n"
            "特别适合隐私保护印章等功能型产品。"
            "第一帧直接进入最强产品动作。"
        ),
    },
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
# 4. Excel
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
# 5. 历史字段
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

    "model_used",
    "fallback_used",
    "retry_count",

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
# ============================================================
# 非 TikTok 官方 Benchmark
# 有账号真实基准时优先使用账号数据
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
# 7. JSON SCHEMA
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
# 8. PAGE
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
        font-size: 1.4rem !important;
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

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 9. SECRETS
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


def create_gemini_client(
    api_key,
):
    return genai.Client(
        api_key=api_key
    )


# ============================================================
# 10. SESSION
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
# 11. GENERAL HELPERS
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


def compact_dict(
    data,
):
    return {
        key: value
        for key, value in data.items()
        if value is not None
        and value != ""
    }


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
            "AI未返回有效结果，请重新执行。"
        )

    try:
        return json.loads(
            raw_text
        )

    except json.JSONDecodeError as exc:
        raise ValueError(
            "AI返回格式异常，请重新执行。"
        ) from exc


def get_video_signature(
    uploaded_video,
):
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
# 12. Gemini 503 / 429 自动重试 + 模型切换
# ============================================================

TRANSIENT_ERROR_MARKERS = [
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
        in TRANSIENT_ERROR_MARKERS
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


def friendly_ai_error(
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
            "AI请求参数异常。"
            "请联系管理员检查模型或结构化输出配置。"
        )

    if (
        "429" in text
        or "RESOURCE_EXHAUSTED" in text
    ):
        return (
            "当前 Gemini 请求额度或并发较高。"
            "系统已经自动重试并切换备用模型，"
            "请稍后再次执行。"
        )

    if (
        "503" in text
        or "UNAVAILABLE" in text
        or "HIGH DEMAND" in text
    ):
        return (
            "当前 AI 模型服务繁忙。"
            "系统已自动重试并尝试备用模型，"
            "请稍后再次执行。"
        )

    return (
        "AI服务暂时未完成本次任务，"
        "请稍后重新执行。"
    )


def generate_content_resilient(
    client,
    contents,
    config,
):
    """
    稳定调用策略：

    3.5 Flash-Lite
        ↓
    短暂错误自动重试
        ↓
    3.1 Flash-Lite
        ↓
    3.6 Flash

    400/401 等明确配置错误不会浪费时间重试。
    """

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

                call_metadata = {
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
                    call_metadata,
                )

            except Exception as exc:
                last_exception = exc

                # 模型不存在：
                # 不继续重试当前模型，直接下一模型
                if is_model_error(
                    exc
                ):
                    break

                # 400 / 401 等真正配置问题：
                # 不进行无意义重试
                if not is_transient_error(
                    exc
                ):
                    raise

                # 当前模型还有重试机会
                if (
                    attempt_index
                    < MAX_ATTEMPTS_PER_MODEL - 1
                ):
                    delay_seconds = (
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
                        delay_seconds
                    )

                    continue

                # 当前模型重试完成
                # 自动进入备用模型
                break

    raise RuntimeError(
        friendly_ai_error(
            last_exception
        )
    ) from last_exception


# ============================================================
# 13. HISTORY
# ============================================================

def empty_history_dataframe():
    return pd.DataFrame(
        columns=HISTORY_COLUMNS
    )


def normalize_history_dataframe(
    dataframe,
):
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


def write_history(
    dataframe,
):
    dataframe = normalize_history_dataframe(
        dataframe
    )

    HISTORY_FILE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temporary_file = (
        HISTORY_FILE.with_name(
            "history_log.tmp.csv"
        )
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
        for column
        in HISTORY_COLUMNS
    }

    if not row[
        "record_id"
    ]:
        row[
            "record_id"
        ] = uuid.uuid4().hex[:12]

    now = datetime.now(
        timezone.utc
    )

    row[
        "created_at_utc"
    ] = now.isoformat(
        timespec="seconds"
    )

    row[
        "created_at_cn"
    ] = (
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
# 14. LOGIN
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
                if role == "主账号(Admin)"
                else staff_password
            )

            if not operator.strip():
                st.error(
                    "请输入操作人。"
                )

            elif (
                password
                != expected
            ):
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
# 15. SIDEBAR HISTORY
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

            selected_operator = (
                st.selectbox(
                    "操作人",
                    operators,
                    key=(
                        "sidebar_"
                        "operator_filter"
                    ),
                )
            )

            if (
                selected_operator
                != "全部"
            ):
                filtered = filtered[
                    filtered[
                        "operator"
                    ]
                    == selected_operator
                ]

        record_types = (
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

        selected_type = (
            st.selectbox(
                "类型",
                record_types,
                key=(
                    "sidebar_"
                    "record_filter"
                ),
            )
        )

        if (
            selected_type
            != "全部"
        ):
            filtered = filtered[
                filtered[
                    "record_type"
                ]
                == selected_type
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
# 16. EXCEL IMPORT
# ============================================================

def excel_script_to_text(
    uploaded_file,
):
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
        for column
        in EXCEL_COLUMNS
        if column
        not in dataframe.columns
    ]

    if missing:
        raise ValueError(
            "Excel缺少字段："
            + "、".join(
                missing
            )
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
# 17. STORYBOARD FORMAT
# ============================================================

def storyboard_to_dataframe(
    rows,
):
    output = []

    for index, row in enumerate(
        rows or [],
        start=1,
    ):
        output.append(
            {
                "分镜序号":
                    clean_text(
                        row.get(
                            "sequence"
                        )
                    )
                    or str(index),

                "景别/机位":
                    clean_text(
                        row.get(
                            "shot"
                        )
                    ),

                "画面描述(道具/动作)":
                    clean_text(
                        row.get(
                            "visual"
                        )
                    ),

                "英文口播文案/字幕":
                    clean_text(
                        row.get(
                            "copy_en"
                        )
                    ),

                "音效/节奏提示":
                    clean
