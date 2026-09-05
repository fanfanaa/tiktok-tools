import json
import time
from google.genai import types

from config import REVIEW_MODEL_CHAIN
from review_schema import REVIEW_SCHEMA
from common import parse_json_output
from gemini_base import generate_resilient as _base_generate_resilient


def thinking_config():
    return types.ThinkingConfig(thinking_level="minimal")

def generate_resilient(client, contents, config):
    return _base_generate_resilient(client, contents, config, model_chain=REVIEW_MODEL_CHAIN)

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
