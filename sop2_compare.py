from datetime import datetime
import json

import pandas as pd
import streamlit as st

from config import PRODUCT_CATEGORIES, SOP2_MAX_VIRAL_VIDEOS, SOP2_MAX_OWN_VIDEOS
from common import clean_text, get_api_key, json_dumps, make_signature, safe_int
from gemini_base import create_client, friendly_error
from gemini_sop2 import pre_analyze, deep_compare
from history_service import append_history
from export_service import build_sop2_chatgpt_payload, build_sop2_export_excel
from background_service import (
    render_background_task_status,
    submit_background_task,
    take_background_task,
    task_key,
)
from workspace_service import (
    clear_retained_uploads,
    get_retained_uploads,
    new_workspace_id,
    remember_uploads,
    render_resume_notice,
    restore_page_memory,
    save_workspace_record,
    snapshot_page_memory,
)


SOP2_PAGE_ID = "sop2"
SOP2_WIDGET_KEYS = [
    "sop2_account", "sop2_category_mode", "sop2_product_name",
    "sop2_custom_category", "sop2_category_select", "sop2_user_points",
    "sop2_selected_viral", "sop2_selected_own",
]
SOP2_STATE_KEYS = SOP2_WIDGET_KEYS + [
    "sop2_pre_result", "sop2_pre_meta", "sop2_pre_completed_at",
    "sop2_upload_signature", "sop2_deep_result", "sop2_deep_meta",
    "sop2_deep_context",
]


def _init_state():
    defaults = {
        "sop2_pre_result": None,
        "sop2_pre_meta": {},
        "sop2_pre_completed_at": "",
        "sop2_upload_signature": "",
        "sop2_selected_viral": None,
        "sop2_selected_own": None,
        "sop2_deep_result": None,
        "sop2_deep_meta": {},
        "sop2_deep_context": "",
    }
    for k,v in defaults.items():
        st.session_state.setdefault(k,v)


def _uploads_signature(viral, own, category, product_name, points):
    chunks = [category, product_name, points]
    for group in [viral or [], own or []]:
        for f in group:
            chunks.extend([f.name, len(f.getvalue()), f.getvalue()[:128]])
    return make_signature(*[x.hex() if isinstance(x, bytes) else x for x in chunks])


_init_state()
restore_page_memory(SOP2_PAGE_ID, SOP2_WIDGET_KEYS)
render_resume_notice(SOP2_PAGE_ID)
api_key = get_api_key()
client = create_client(api_key) if api_key else None
SOP2_PRE_TASK_KEY = task_key("sop2_pre_analysis")
SOP2_DEEP_TASK_KEY = task_key("sop2_deep_compare")
if not api_key:
    st.error("系统未配置 Gemini API Key，请联系管理员。")

st.caption("SOP2｜爆款 VS 我的作品 → 找差距 → 重剪 / 补拍 → 导出 Excel + ChatGPT JSON")

st.markdown("### ① 产品信息")
c1,c2,c3 = st.columns(3)
with c1:
    tiktok_account = st.text_input("TikTok账号", key="sop2_account", placeholder="用于历史归档")
with c2:
    category_mode = st.radio(
        "产品品类",
        ["选择现有品类", "自定义输入"],
        horizontal=True,
        key="sop2_category_mode",
    )
with c3:
    product_name = st.text_input("产品名称 / SKU", key="sop2_product_name")

normal_categories = [x for x in PRODUCT_CATEGORIES if x != "其他 / 自定义"]
if category_mode == "自定义输入":
    custom_category = st.text_input(
        "自定义产品品类",
        key="sop2_custom_category",
        placeholder="直接输入，例如：LED阅读灯 / Flat Panel Book Light",
    )
    category = clean_text(custom_category) or "自定义品类（未填写）"
else:
    category = st.selectbox(
        "选择现有产品品类",
        normal_categories,
        key="sop2_category_select",
    )

user_points = st.text_area("我们的真实产品卖点（选填）", key="sop2_user_points", height=85,
                           placeholder="用于约束对比结论，避免AI把爆款里不存在于我们产品的功能当成优化方向。")

st.markdown("### ② 上传爆款视频")
new_viral_videos = st.file_uploader(
    f"爆款视频 1-{SOP2_MAX_VIRAL_VIDEOS} 条",
    type=["mp4"], accept_multiple_files=True, key="sop2_viral_upload"
)
if new_viral_videos:
    remember_uploads(SOP2_PAGE_ID, "viral", new_viral_videos)
retained_viral_videos = get_retained_uploads(SOP2_PAGE_ID, "viral")
viral_videos = new_viral_videos or retained_viral_videos
if retained_viral_videos and not new_viral_videos:
    st.caption("✅ 页面切换前的爆款视频仍已保留：" + "、".join(v.name for v in retained_viral_videos))

st.markdown("### ③ 上传我的拍摄作品")
new_own_videos = st.file_uploader(
    f"我的作品 1-{SOP2_MAX_OWN_VIDEOS} 条",
    type=["mp4"], accept_multiple_files=True, key="sop2_own_upload"
)
if new_own_videos:
    remember_uploads(SOP2_PAGE_ID, "own", new_own_videos)
retained_own_videos = get_retained_uploads(SOP2_PAGE_ID, "own")
own_videos = new_own_videos or retained_own_videos
if retained_own_videos and not new_own_videos:
    st.caption("✅ 页面切换前的我的作品仍已保留：" + "、".join(v.name for v in retained_own_videos))

if (retained_viral_videos or retained_own_videos) and not (new_viral_videos or new_own_videos):
    if st.button("清除已保留视频", key="sop2_clear_retained_uploads"):
        clear_retained_uploads(SOP2_PAGE_ID)
        st.session_state["sop2_upload_signature"] = ""
        st.session_state["sop2_pre_result"] = None
        st.session_state["sop2_pre_completed_at"] = ""
        st.session_state["sop2_selected_viral"] = None
        st.session_state["sop2_selected_own"] = None
        st.session_state["sop2_deep_result"] = None
        st.session_state["sop2_deep_context"] = ""
        new_workspace_id("SOP2")
        st.rerun()

if viral_videos or own_videos:
    total_count = len(viral_videos or []) + len(own_videos or [])
    total_mb = sum(len(x.getvalue()) for x in (viral_videos or []) + (own_videos or [])) / 1024 / 1024
    st.caption(f"当前共 {total_count} 条 · 总大小 {total_mb:.2f} MB；系统会自动选择 Inline 或 Files API。")
    sig = _uploads_signature(viral_videos, own_videos, category, product_name, user_points)
    if st.session_state["sop2_upload_signature"] != sig:
        previous_signature = st.session_state.get("sop2_upload_signature", "")
        if previous_signature:
            new_workspace_id("SOP2")
        st.session_state["sop2_upload_signature"] = sig
        st.session_state["sop2_pre_result"] = None
        st.session_state["sop2_pre_completed_at"] = ""
        st.session_state["sop2_selected_viral"] = None
        st.session_state["sop2_selected_own"] = None
        st.session_state["sop2_deep_result"] = None
        st.session_state["sop2_deep_context"] = ""

invalid = (
    not viral_videos or not own_videos or
    len(viral_videos or []) > SOP2_MAX_VIRAL_VIDEOS or
    len(own_videos or []) > SOP2_MAX_OWN_VIDEOS
)

pre_done = st.session_state.get("sop2_pre_result") is not None
pre_button_label = "重新分析" if pre_done else "快速对比预分析"
pre_button_type = "secondary" if pre_done else "primary"

if pre_done:
    completed_at = st.session_state.get("sop2_pre_completed_at", "")
    st.caption(f"✅ 预分析已完成{(' · ' + completed_at) if completed_at else ''}。如需重新跑一次，可点击下方“重新分析”。")
    pre_meta = st.session_state.get("sop2_pre_meta", {}) or {}
    if clean_text(pre_meta.get("fallback_used")):
        st.info("Gemini 3.8 当前繁忙，本次已自动切换 Gemini 3.5 完成预分析。")

# 预分析改为后台任务：换页不会中断。
pre_task = take_background_task(SOP2_PRE_TASK_KEY)
if pre_task:
    task_context = pre_task.get("context", {})
    if pre_task.get("status") == "done":
        result, meta = pre_task.get("result")
        # 只把结果写回发起任务时的同一批输入，避免用户中途换视频后旧结果覆盖新状态。
        if task_context.get("input_signature") == st.session_state.get("sop2_upload_signature"):
            st.session_state["sop2_pre_result"] = result
            st.session_state["sop2_pre_meta"] = meta
            st.session_state["sop2_pre_completed_at"] = datetime.now().strftime("%H:%M")
            st.session_state["sop2_selected_viral"] = None
            st.session_state["sop2_selected_own"] = None
            st.session_state["sop2_deep_result"] = None
            append_history({
                "module":"SOP2", "record_type":"爆款对比预分析",
                "role":task_context.get("role", ""), "operator":task_context.get("operator", ""),
                "tiktok_account":task_context.get("tiktok_account", ""),
                "product_category":task_context.get("category", ""),
                "product_name":task_context.get("product_name", ""),
                "input_selling_points":task_context.get("user_points", ""),
                "video_names":task_context.get("video_names", ""),
                "video_count":task_context.get("video_count", ""),
                "model_used":meta.get("model_used",""),
                "fallback_used":meta.get("fallback_used",""),
                "retry_count":meta.get("retry_count",""),
                "analysis_seconds":meta.get("analysis_seconds",""),
                "full_output_json":json_dumps(result),
            })
            st.success("爆款对比预分析已在后台完成。")
        else:
            st.warning("刚才的后台预分析已完成，但当前上传内容已经变化，因此没有覆盖当前页面。")
    elif pre_task.get("status") == "error":
        st.error(friendly_error(pre_task.get("error")))

if st.button(pre_button_label, type=pre_button_type, use_container_width=True, disabled=(client is None or invalid), key="sop2_pre_btn"):
    task_viral = get_retained_uploads(SOP2_PAGE_ID, "viral") or viral_videos
    task_own = get_retained_uploads(SOP2_PAGE_ID, "own") or own_videos
    submitted = submit_background_task(
        SOP2_PRE_TASK_KEY,
        pre_analyze,
        client, task_viral, task_own, category, product_name, user_points,
        context={
            "input_signature": st.session_state.get("sop2_upload_signature", ""),
            "role": st.session_state.get("role", ""),
            "operator": st.session_state.get("operator", ""),
            "tiktok_account": tiktok_account,
            "category": category,
            "product_name": product_name,
            "user_points": user_points,
            "video_names": "爆款: " + " | ".join(v.name for v in task_viral) + "；我的: " + " | ".join(v.name for v in task_own),
            "video_count": len(task_viral) + len(task_own),
        },
    )
    if submitted:
        st.rerun()
    else:
        st.info("这项预分析已经在后台运行，无需重复提交。")

render_background_task_status(
    SOP2_PRE_TASK_KEY,
    "Gemini 3.8 正在后台做爆款对比预分析；如遇高峰会自动切换 Gemini 3.5。",
)

pre = st.session_state.get("sop2_pre_result")
if pre:
    st.markdown("### ④ 查看预分析，并选择比较对象")
    st.info("AI推荐组合：爆款视频{} VS 我的作品{}\n\n{}".format(
        pre.get("recommended_viral_index",""), pre.get("recommended_own_index",""), pre.get("recommendation_reason","")
    ))

    quick = pre.get("quick_comparison", {}) or {}
    if quick:
        st.markdown("#### AI推荐组合｜快速优缺点对比")
        q1, q2 = st.columns(2)
        with q1:
            st.markdown("**爆款优势**")
            for x in quick.get("viral_advantages", []): st.markdown(f"- {x}")
            st.markdown("**爆款短板**")
            for x in quick.get("viral_weaknesses", []): st.markdown(f"- {x}")
        with q2:
            st.markdown("**我的优势**")
            for x in quick.get("own_advantages", []): st.markdown(f"- {x}")
            st.markdown("**我的短板**")
            for x in quick.get("own_weaknesses", []): st.markdown(f"- {x}")
        st.markdown(f'**最大差距：** {quick.get("biggest_gap", "")}')
        st.markdown("**最值得吸收：**")
        for x in quick.get("most_worth_learning", []): st.markdown(f"- {x}")

    with st.expander("查看全部视频详细预分析", expanded=False):
        st.markdown("**爆款视频**")
        for v in pre.get("viral_videos",[]):
            st.markdown(f'**爆款{v.get("video_index")}｜{v.get("filename","")}｜推荐指数 {v.get("recommend_score","")}**')
            st.write(v.get("one_sentence_core",""))
            st.markdown("**脚本路线**")
            st.write(v.get("script_route",""))
            st.markdown("**前3秒 Hook**")
            st.write(v.get("first_3s_hook",""))
            st.markdown("**画面与节奏**")
            st.write(v.get("visual_rhythm",""))
            va, vb = st.columns(2)
            with va:
                st.markdown("**优点**")
                for x in v.get("strengths",[]): st.markdown(f"- {x}")
            with vb:
                st.markdown("**短板 / 风险**")
                for x in v.get("weaknesses",[]): st.markdown(f"- {x}")
            st.markdown(f'**为什么有效：** {v.get("why_it_works", "")}')
            st.markdown(f'**最优先改进：** {v.get("improvement_priority", "")}')
            st.markdown(f'**对比价值：** {v.get("compare_value", "")}')
            st.divider()
        st.markdown("**我的作品**")
        for v in pre.get("own_videos",[]):
            st.markdown(f'**作品{v.get("video_index")}｜{v.get("filename","")}｜推荐指数 {v.get("recommend_score","")}**')
            st.write(v.get("one_sentence_core",""))
            st.markdown("**脚本路线**")
            st.write(v.get("script_route",""))
            st.markdown("**前3秒 Hook**")
            st.write(v.get("first_3s_hook",""))
            st.markdown("**画面与节奏**")
            st.write(v.get("visual_rhythm",""))
            oa, ob = st.columns(2)
            with oa:
                st.markdown("**优点**")
                for x in v.get("strengths",[]): st.markdown(f"- {x}")
            with ob:
                st.markdown("**短板 / 风险**")
                for x in v.get("weaknesses",[]): st.markdown(f"- {x}")
            st.markdown(f'**为什么有效：** {v.get("why_it_works", "")}')
            st.markdown(f'**最优先改进：** {v.get("improvement_priority", "")}')
            st.markdown(f'**对比价值：** {v.get("compare_value", "")}')
            st.divider()

    viral_options = [safe_int(v.get("video_index"),i+1) for i,v in enumerate(pre.get("viral_videos",[]))]
    own_options = [safe_int(v.get("video_index"),i+1) for i,v in enumerate(pre.get("own_videos",[]))]
    s1,s2 = st.columns(2)
    with s1:
        selected_viral = st.radio(
            "请选择主要参考爆款",
            options=viral_options,
            index=None,
            format_func=lambda x: f'爆款{x}｜{next((v.get("filename","") for v in pre.get("viral_videos",[]) if safe_int(v.get("video_index"),0)==x),"")}' + ("（AI推荐）" if x==safe_int(pre.get("recommended_viral_index"),0) else ""),
            key="sop2_selected_viral",
        )
    with s2:
        selected_own = st.radio(
            "请选择我的比较作品",
            options=own_options,
            index=None,
            format_func=lambda x: f'作品{x}｜{next((v.get("filename","") for v in pre.get("own_videos",[]) if safe_int(v.get("video_index"),0)==x),"")}' + ("（AI推荐）" if x==safe_int(pre.get("recommended_own_index"),0) else ""),
            key="sop2_selected_own",
        )

    if selected_viral is None or selected_own is None:
        st.warning("AI只做推荐。请你亲自选择一条爆款和一条自己的作品，才能开始深度对比。")
    else:
        viral_summary = next(v for v in pre.get("viral_videos",[]) if safe_int(v.get("video_index"),0)==selected_viral)
        own_summary = next(v for v in pre.get("own_videos",[]) if safe_int(v.get("video_index"),0)==selected_own)
        deep_context = make_signature(selected_viral, selected_own, viral_summary, own_summary, user_points)
        if st.session_state.get("sop2_deep_context") and st.session_state["sop2_deep_context"] != deep_context:
            st.session_state["sop2_deep_result"] = None
            st.session_state["sop2_deep_context"] = ""

        deep_task = take_background_task(SOP2_DEEP_TASK_KEY)
        if deep_task:
            task_context = deep_task.get("context", {})
            if deep_task.get("status") == "done":
                result, meta = deep_task.get("result")
                if task_context.get("deep_context") == deep_context:
                    st.session_state["sop2_deep_result"] = result
                    st.session_state["sop2_deep_meta"] = meta
                    st.session_state["sop2_deep_context"] = deep_context
                    append_history({
                        "module":"SOP2", "record_type":"爆款深度对比",
                        "role":task_context.get("role", ""), "operator":task_context.get("operator", ""),
                        "tiktok_account":task_context.get("tiktok_account", ""),
                        "product_category":task_context.get("category", ""),
                        "product_name":task_context.get("product_name", ""),
                        "input_selling_points":task_context.get("user_points", ""),
                        "viral_video_name":task_context.get("viral_video_name", ""),
                        "own_video_name":task_context.get("own_video_name", ""),
                        "reference_video_index":task_context.get("selected_viral", ""),
                        "reference_video_name":task_context.get("viral_video_name", ""),
                        "model_used":meta.get("model_used",""),
                        "fallback_used":meta.get("fallback_used",""),
                        "retry_count":meta.get("retry_count",""),
                        "analysis_seconds":meta.get("analysis_seconds",""),
                        "diagnosis_summary":result.get("one_sentence_conclusion",""),
                        "reedit_value":result.get("reedit_value",""),
                        "full_output_json":json_dumps(result),
                    })
                    st.success("深度对比已在后台完成。")
                else:
                    st.warning("刚才的后台深度对比已完成，但当前比较对象已经变化，因此没有覆盖当前结果。")
            elif deep_task.get("status") == "error":
                st.error(friendly_error(deep_task.get("error")))

        if st.button("开始深度对比", type="primary", use_container_width=True, key="sop2_deep_btn"):
            task_viral_files = get_retained_uploads(SOP2_PAGE_ID, "viral") or viral_videos
            task_own_files = get_retained_uploads(SOP2_PAGE_ID, "own") or own_videos
            try:
                viral_file = task_viral_files[selected_viral - 1]
                own_file = task_own_files[selected_own - 1]
            except Exception:
                st.error("原视频文件已不在当前页面，请重新上传后再进行深度对比。")
            else:
                submitted = submit_background_task(
                    SOP2_DEEP_TASK_KEY,
                    deep_compare,
                    client, viral_file, own_file, category, product_name, user_points, viral_summary, own_summary,
                    context={
                        "deep_context": deep_context,
                        "role": st.session_state.get("role", ""),
                        "operator": st.session_state.get("operator", ""),
                        "tiktok_account": tiktok_account,
                        "category": category,
                        "product_name": product_name,
                        "user_points": user_points,
                        "viral_video_name": viral_file.name,
                        "own_video_name": own_file.name,
                        "selected_viral": selected_viral,
                    },
                )
                if submitted:
                    st.rerun()
                else:
                    st.info("这项深度对比已经在后台运行，无需重复提交。")

        render_background_task_status(
            SOP2_DEEP_TASK_KEY,
            "Gemini 3.8 正在后台做深度对比；如遇高峰会自动切换 Gemini 3.5。",
        )

        deep = st.session_state.get("sop2_deep_result")
        if deep:
            deep_meta = st.session_state.get("sop2_deep_meta", {}) or {}
            if clean_text(deep_meta.get("fallback_used")):
                st.info("Gemini 3.8 当前繁忙，本次深度对比已自动切换 Gemini 3.5 完成。")
            st.markdown("### ⑤ 核心差距结论")
            st.info(deep.get("one_sentence_conclusion",""))
            st.markdown(f'**核心差距：** {deep.get("core_gap","")}')

            st.markdown("### ⑥ 脚本路线对比")
            r1,r2 = st.columns(2)
            with r1:
                st.markdown("**爆款脚本路线**")
                st.write(deep.get("viral_script_route",""))
            with r2:
                st.markdown("**我的脚本路线**")
                st.write(deep.get("own_script_route",""))

            st.markdown("### ⑦ 10维差距表")
            dims = pd.DataFrame(deep.get("comparison_dimensions",[])).rename(columns={"dimension":"对比项","viral":"爆款视频","own":"我的作品","gap":"核心差距","suggestion":"建议"})
            st.dataframe(dims, hide_index=True, use_container_width=True)

            st.markdown("### ⑧ 双方优缺点对比")
            a,b = st.columns(2)
            with a:
                st.markdown("**爆款优势**")
                for x in deep.get("viral_strengths",[]): st.markdown(f"- {x}")
                st.markdown("**爆款短板 / 不建议照搬**")
                for x in deep.get("viral_weaknesses",[]): st.markdown(f"- {x}")
            with b:
                st.markdown("**我的优势**")
                for x in deep.get("own_strengths",[]): st.markdown(f"- {x}")
                st.markdown("**我的短板 / 优先修正**")
                for x in deep.get("own_weaknesses",[]): st.markdown(f"- {x}")

            st.markdown("### ⑨ 重剪与补拍判断")
            value = deep.get("reedit_value","")
            st.info(f"重剪价值：{value}\n\n{deep.get('reedit_reason','')}")
            e1,e2 = st.columns(2)
            with e1:
                st.markdown("**可以保留**")
                for x in deep.get("keep_segments",[]): st.markdown(f'- `{x.get("time_range","")}` {x.get("content","")} — {x.get("reason","")}')
                st.markdown("**建议删除**")
                for x in deep.get("delete_segments",[]): st.markdown(f'- `{x.get("time_range","")}` {x.get("content","")} — {x.get("reason","")}')
            with e2:
                st.markdown("**建议前移 / 调整**")
                for x in deep.get("move_segments",[]): st.markdown(f'- `{x.get("source_time","")} → {x.get("target_time","")}` {x.get("content","")} — {x.get("reason","")}')
                st.markdown("**必须补拍**")
                for x in deep.get("reshoot_segments",[]): st.markdown(f'- {x.get("shot","")}｜{x.get("action","")} — {x.get("purpose","")}')
            st.markdown("**整体重剪计划**")
            st.write(deep.get("editing_plan",""))
            st.markdown("**下次拍摄优化方向**")
            st.write(deep.get("optimization_plan",""))

            viral_name = viral_summary.get("filename","")
            own_name = own_summary.get("filename","")
            product_info = {"category":category,"product_name":product_name,"real_selling_points":user_points}
            payload = build_sop2_chatgpt_payload(deep, product_info, viral_name, own_name)

            st.markdown("### ⑩ 导出")
            with st.expander("查看给 ChatGPT 的 JSON", expanded=False):
                st.code(json.dumps(payload, ensure_ascii=False, indent=2), language="json")
            d1,d2 = st.columns(2)
            with d1:
                st.download_button(
                    "导出 Excel", data=build_sop2_export_excel(deep,payload),
                    file_name="SOP2_爆款对比_"+datetime.now().strftime("%Y%m%d_%H%M")+".xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary"
                )
            with d2:
                st.download_button(
                    "导出 ChatGPT JSON", data=json.dumps(payload,ensure_ascii=False,indent=2).encode("utf-8"),
                    file_name="SOP2_ChatGPT_"+datetime.now().strftime("%Y%m%d_%H%M")+".json",
                    mime="application/json", use_container_width=True
                )

# 页面状态与“工作记录”自动保存
snapshot_page_memory(SOP2_PAGE_ID, SOP2_WIDGET_KEYS)
_workspace_pre = st.session_state.get("sop2_pre_result") or {}
_workspace_deep = st.session_state.get("sop2_deep_result") or {}
_workspace_selected_viral = st.session_state.get("sop2_selected_viral")
_workspace_selected_own = st.session_state.get("sop2_selected_own")
_workspace_ref_name = ""
_workspace_own_name = ""
if isinstance(_workspace_pre, dict):
    for _item in _workspace_pre.get("viral_videos", []):
        if safe_int(_item.get("video_index"), 0) == safe_int(_workspace_selected_viral, 0):
            _workspace_ref_name = clean_text(_item.get("filename", ""))
            break
    for _item in _workspace_pre.get("own_videos", []):
        if safe_int(_item.get("video_index"), 0) == safe_int(_workspace_selected_own, 0):
            _workspace_own_name = clean_text(_item.get("filename", ""))
            break

if _workspace_deep:
    _workspace_step = "深度对比已完成"
elif _workspace_pre and _workspace_selected_viral and _workspace_selected_own:
    _workspace_step = f"预分析已完成 · 已选：{_workspace_ref_name or '爆款'} VS {_workspace_own_name or '我的作品'}"
elif _workspace_pre:
    _workspace_step = "预分析已完成，待选择比较对象"
elif viral_videos or own_videos:
    _workspace_step = "视频已上传，待预分析"
else:
    _workspace_step = "产品信息填写中"

_workspace_video_names = (
    "爆款: " + " | ".join(v.name for v in (viral_videos or []))
    + "；我的: " + " | ".join(v.name for v in (own_videos or []))
)
_workspace_meaningful = bool(
    viral_videos or own_videos or clean_text(tiktok_account) or clean_text(product_name)
    or clean_text(user_points) or _workspace_pre or _workspace_deep
)

save_workspace_record(
    module="SOP2",
    page_id=SOP2_PAGE_ID,
    page_path="sop2_compare.py",
    step=_workspace_step,
    state_keys=SOP2_STATE_KEYS,
    tiktok_account=tiktok_account,
    product_category=category,
    product_name=product_name,
    input_selling_points=user_points,
    video_names=_workspace_video_names,
    reference_video_name=_workspace_ref_name,
    meaningful=_workspace_meaningful,
)

