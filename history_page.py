from datetime import datetime

import streamlit as st

from config import HISTORY_COLUMNS
from common import clean_text
from history_service import scoped_history

st.caption("历史记录｜主账号查看全员，分账号仅查看本人")

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
