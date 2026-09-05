from datetime import datetime
import streamlit as st

from config import DEFAULT_STAFF_PASSWORD, DEFAULT_ADMIN_PASSWORD
from common import get_secret
from history_service import scoped_history


def init_auth_session():
    st.session_state.setdefault("authenticated", False)
    st.session_state.setdefault("role", "")
    st.session_state.setdefault("operator", "")


def logout():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

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
