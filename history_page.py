from datetime import datetime
import json

import pandas as pd
import streamlit as st

from common import clean_text
from history_service import import_history_dataframe, scoped_history
from workspace_service import restore_workspace_record

st.caption("历史记录｜工作记录可继续打开｜操作历史永久保存｜Supabase 持久化")

history = scoped_history()
work_history = history[history["record_type"] == "工作记录"].copy() if not history.empty else history.copy()
operation_history = history[history["record_type"] != "工作记录"].copy() if not history.empty else history.copy()

work_tab, operation_tab = st.tabs(["📌 工作记录（可继续）", "🧾 操作历史"])

with work_tab:
    st.caption("这里保存的是正在做/已经做到一半的页面状态、人工选择和AI结果。点“继续/查看”会回到对应页面。")

    filtered_work = work_history.copy()
    f1, f2, f3 = st.columns(3)

    with f1:
        if st.session_state["role"] == "主账号(Admin)":
            operator_options = ["全部"] + sorted([x for x in filtered_work.get("operator", pd.Series(dtype=str)).unique() if x])
            selected_operator = st.selectbox("操作人", operator_options, key="work_history_operator")
            if selected_operator != "全部":
                filtered_work = filtered_work[filtered_work["operator"] == selected_operator]

    with f2:
        module_options = ["全部"] + sorted([x for x in filtered_work.get("module", pd.Series(dtype=str)).unique() if x])
        selected_module = st.selectbox("模块", module_options, key="work_history_module")
        if selected_module != "全部":
            filtered_work = filtered_work[filtered_work["module"] == selected_module]

    with f3:
        account_options = ["全部"] + sorted([x for x in filtered_work.get("tiktok_account", pd.Series(dtype=str)).unique() if x])
        selected_account = st.selectbox("TikTok账号", account_options, key="work_history_account")
        if selected_account != "全部":
            filtered_work = filtered_work[filtered_work["tiktok_account"] == selected_account]

    if filtered_work.empty:
        st.info("暂无工作记录。开始上传视频、填写复盘或生成分析后，会自动保存在这里。")
    else:
        if "created_at_utc" in filtered_work.columns:
            filtered_work = filtered_work.sort_values("created_at_utc", ascending=False)

        for _, row in filtered_work.iterrows():
            module = clean_text(row.get("module", "")) or "工作"
            product_name = clean_text(row.get("product_name", "")) or "未填写产品名"
            step = clean_text(row.get("diagnosis_summary", "")) or "处理中"
            saved_at = clean_text(row.get("created_at_cn", ""))
            operator = clean_text(row.get("operator", ""))
            account = clean_text(row.get("tiktok_account", ""))
            videos = clean_text(row.get("video_names", ""))
            record_id = clean_text(row.get("record_id", ""))

            with st.container(border=True):
                top1, top2 = st.columns([4, 1])
                with top1:
                    st.markdown(f"**{module}｜{product_name}**")
                    st.caption(" · ".join([x for x in [saved_at, operator, account] if x]))
                    st.write(f"当前进度：{step}")
                    if videos:
                        st.caption(f"素材：{videos}")
                with top2:
                    if st.button("继续 / 查看", key=f"resume_workspace_{record_id}", use_container_width=True, type="primary"):
                        try:
                            payload = json.loads(clean_text(row.get("full_output_json", "")) or "{}")
                            restore_workspace_record(payload)
                            target_page = clean_text(payload.get("page_path", ""))
                            if not target_page:
                                raise ValueError("这条记录没有可恢复的页面路径。")
                            st.switch_page(target_page)
                        except Exception as exc:
                            st.error(f"恢复工作记录失败：{exc}")

with operation_tab:
    st.caption("这里是原来的操作历史：记录每次爆款拆解、深度对比、数据复盘等已执行动作。")

    if st.session_state["role"] == "主账号(Admin)":
        with st.expander("恢复 / 导入旧历史 CSV", expanded=False):
            st.caption("仅用于恢复以前下载的操作历史 CSV；如果确认没有旧记录，这个入口后续可以删除。")
            old_history_file = st.file_uploader("选择历史 CSV", type=["csv"], key="history_restore_csv")
            if st.button(
                "导入到永久历史库",
                type="primary",
                use_container_width=True,
                disabled=old_history_file is None,
                key="history_restore_btn",
            ):
                try:
                    old_history = pd.read_csv(old_history_file, dtype=str, keep_default_na=False, encoding="utf-8-sig")
                    imported_count = import_history_dataframe(old_history)
                    st.success(f"已导入 {imported_count} 条历史记录。")
                    st.rerun()
                except Exception as exc:
                    st.error(f"历史导入失败：{exc}")

    if operation_history.empty:
        st.info("暂无操作历史。")
    else:
        filtered = operation_history.copy()
        h1, h2, h3 = st.columns(3)

        with h1:
            if st.session_state["role"] == "主账号(Admin)":
                operators = ["全部"] + sorted([value for value in filtered["operator"].unique() if value])
                operator_filter = st.selectbox("操作人", operators, key="history_operator")
                if operator_filter != "全部":
                    filtered = filtered[filtered["operator"] == operator_filter]

        with h2:
            record_types = ["全部"] + sorted([value for value in filtered["record_type"].unique() if value])
            type_filter = st.selectbox("类型", record_types, key="history_type")
            if type_filter != "全部":
                filtered = filtered[filtered["record_type"] == type_filter]

        with h3:
            account_options = ["全部"] + sorted([value for value in filtered["tiktok_account"].unique() if value])
            account_filter = st.selectbox("TikTok账号", account_options, key="history_account")
            if account_filter != "全部":
                filtered = filtered[filtered["tiktok_account"] == account_filter]

        display_columns = [
            "created_at_cn", "record_type", "operator", "tiktok_account", "product_name",
            "reference_video_name", "direction_name", "selected_scene", "selected_perspective", "priority_issue",
        ]
        available_columns = [column for column in display_columns if column in filtered.columns]
        st.dataframe(filtered[available_columns].iloc[::-1], hide_index=True, use_container_width=True)

        st.download_button(
            "下载操作历史 CSV",
            data=filtered.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
            file_name="TikTok操作历史_" + datetime.now().strftime("%Y%m%d_%H%M") + ".csv",
            mime="text/csv",
            use_container_width=True,
        )
