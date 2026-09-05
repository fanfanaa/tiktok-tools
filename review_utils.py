import io
import pandas as pd
import streamlit as st

from config import SOP_BANDS, OLD_SCRIPT_COLUMNS
from common import clean_text, parse_optional_float

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

