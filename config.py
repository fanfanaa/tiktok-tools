from pathlib import Path
from zoneinfo import ZoneInfo

APP_TITLE = "TikTok爆款视频解析&复盘专用"

# SOP1: 保持线上稳定模型链，不受 SOP2 影响
SOP1_PRIMARY_MODEL = "gemini-3.5-flash-lite"
SOP1_FALLBACK_MODELS = ["gemini-3.1-flash-lite", "gemini-3.6-flash"]
SOP1_MODEL_CHAIN = [SOP1_PRIMARY_MODEL, *SOP1_FALLBACK_MODELS]

# 数据复盘继续沿用原稳定链
REVIEW_MODEL_CHAIN = list(SOP1_MODEL_CHAIN)

# SOP2: 与 SOP1 完全平行，绝不回退到 SOP1 模型
SOP2_MODEL = "gemini-3.8-flash"
SOP2_MODEL_CHAIN = [SOP2_MODEL]

MAX_ATTEMPTS_PER_MODEL = 2
MAX_COMPARE_VIDEOS = 5
SOP2_MAX_VIRAL_VIDEOS = 3
SOP2_MAX_OWN_VIDEOS = 3
SOP2_MAX_TOTAL_VIDEOS = 6
INLINE_BATCH_MAX_MB = 18
FILE_POLL_INTERVAL_SEC = 2
FILE_PROCESS_TIMEOUT_SEC = 180

DEFAULT_STAFF_PASSWORD = "8888"
DEFAULT_ADMIN_PASSWORD = "8888-admin"

HISTORY_FILE = Path("history_log.csv")
CN_TZ = ZoneInfo("Asia/Shanghai")

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
    "其他 / 自定义",
]

TRAFFIC_TYPES = [
    "自然流 Organic",
    "Custom Mode Video Shopping Ads",
    "自然流 + 付费混合",
]

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

HISTORY_COLUMNS = [
    "record_id", "created_at_utc", "created_at_cn",
    "module", "record_type", "role", "operator",
    "tiktok_account", "product_category", "product_name",
    "input_selling_points", "inferred_selling_points",
    "effective_selling_points", "selling_point_mode",
    "reference_video_index", "reference_video_name",
    "viral_video_name", "own_video_name",
    "direction_name", "selected_scene", "selected_perspective",
    "video_names", "video_count",
    "model_used", "fallback_used", "retry_count", "analysis_seconds",
    "priority_issue", "diagnosis_summary", "reedit_value",
    "metrics_json", "account_baseline_json", "full_output_json",
]
