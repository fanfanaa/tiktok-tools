from datetime import datetime

import pandas as pd
import streamlit as st

from config import PRODUCT_CATEGORIES, TRAFFIC_TYPES
from common import clean_text, compact_dict, get_api_key, json_dumps
from gemini_base import create_client, friendly_error
from gemini_review import review_script
from history_service import append_history
from export_service import review_script_to_df, build_review_export_excel
from review_utils import build_metric_assessment, render_optional_float_input, get_script_sheets, excel_sheet_to_script_text

st.session_state.setdefault("review_result", None)
st.session_state.setdefault("review_meta", {})
st.session_state.setdefault("review_original_script", "")
api_key = get_api_key()
client = create_client(api_key) if api_key else None
if not api_key:
    st.error("系统未配置 Gemini API Key，请联系管理员。")

st.caption("数据复盘｜发布后根据真实指标定位优先问题并生成优化脚本")

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

