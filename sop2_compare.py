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
api_key = get_api_key()
client = create_client(api_key) if api_key else None
if not api_key:
    st.error("系统未配置 Gemini API Key，请联系管理员。")

st.caption("SOP2｜爆款 VS 我的作品 → 找差距 → 重剪 / 补拍 → 导出 Excel + ChatGPT JSON")

st.markdown("### ① 产品信息")
c1,c2,c3 = st.columns(3)
with c1:
    tiktok_account = st.text_input("TikTok账号", key="sop2_account", placeholder="用于历史归档")
with c2:
    category = st.selectbox("产品品类", PRODUCT_CATEGORIES, key="sop2_category")
with c3:
    product_name = st.text_input("产品名称 / SKU", key="sop2_product_name")
user_points = st.text_area("我们的真实产品卖点（选填）", key="sop2_user_points", height=85,
                           placeholder="用于约束对比结论，避免AI把爆款里不存在于我们产品的功能当成优化方向。")

st.markdown("### ② 上传爆款视频")
viral_videos = st.file_uploader(f"爆款视频 1-{SOP2_MAX_VIRAL_VIDEOS} 条", type=["mp4"], accept_multiple_files=True, key="sop2_viral_upload")

st.markdown("### ③ 上传我的拍摄作品")
own_videos = st.file_uploader(f"我的作品 1-{SOP2_MAX_OWN_VIDEOS} 条", type=["mp4"], accept_multiple_files=True, key="sop2_own_upload")

if viral_videos or own_videos:
    total_count = len(viral_videos or []) + len(own_videos or [])
    total_mb = sum(len(x.getvalue()) for x in (viral_videos or []) + (own_videos or [])) / 1024 / 1024
    st.caption(f"当前共 {total_count} 条 · 总大小 {total_mb:.2f} MB；系统会自动选择 Inline 或 Files API。")
    sig = _uploads_signature(viral_videos, own_videos, category, product_name, user_points)
    if st.session_state["sop2_upload_signature"] != sig:
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

if st.button(pre_button_label, type=pre_button_type, use_container_width=True, disabled=(client is None or invalid), key="sop2_pre_btn"):
    try:
        with st.spinner("Gemini 3.8 Flash 正在逐条预分析，并推荐最值得比较的组合…"):
            result, meta = pre_analyze(client, viral_videos, own_videos, category, product_name, user_points)
        st.session_state["sop2_pre_result"] = result
        st.session_state["sop2_pre_meta"] = meta
        st.session_state["sop2_pre_completed_at"] = datetime.now().strftime("%H:%M")
        st.session_state["sop2_selected_viral"] = None
        st.session_state["sop2_selected_own"] = None
        st.session_state["sop2_deep_result"] = None
        append_history({
            "module":"SOP2", "record_type":"爆款对比预分析", "role":st.session_state["role"], "operator":st.session_state["operator"],
            "tiktok_account":tiktok_account, "product_category":category, "product_name":product_name,
            "input_selling_points":user_points, "video_names":"爆款: " + " | ".join(v.name for v in viral_videos) + "；我的: " + " | ".join(v.name for v in own_videos),
            "video_count":len(viral_videos)+len(own_videos), "model_used":meta.get("model_used",""),
            "fallback_used":meta.get("fallback_used",""), "retry_count":meta.get("retry_count",""), "analysis_seconds":meta.get("analysis_seconds",""),
            "full_output_json":json_dumps(result),
        })
        st.rerun()
    except Exception as exc:
        st.error(friendly_error(exc))

pre = st.session_state.get("sop2_pre_result")
if pre:
    st.markdown("### ④ 查看预分析，并选择比较对象")
    st.info("AI推荐组合：爆款视频{} VS 我的作品{}\n\n{}".format(
        pre.get("recommended_viral_index",""), pre.get("recommended_own_index",""), pre.get("recommendation_reason","")
    ))

    with st.expander("查看全部视频预分析", expanded=False):
        st.markdown("**爆款视频**")
        for v in pre.get("viral_videos",[]):
            st.markdown(f'**爆款{v.get("video_index")}｜{v.get("filename","")}｜推荐指数 {v.get("recommend_score","")}**')
            st.write(v.get("one_sentence_core",""))
            st.caption("脚本路线：" + clean_text(v.get("script_route","")))
            st.caption("前3秒：" + clean_text(v.get("first_3s_hook","")))
        st.divider()
        st.markdown("**我的作品**")
        for v in pre.get("own_videos",[]):
            st.markdown(f'**作品{v.get("video_index")}｜{v.get("filename","")}｜推荐指数 {v.get("recommend_score","")}**')
            st.write(v.get("one_sentence_core",""))
            st.caption("脚本路线：" + clean_text(v.get("script_route","")))
            st.caption("前3秒：" + clean_text(v.get("first_3s_hook","")))

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

        if st.button("开始深度对比", type="primary", use_container_width=True, key="sop2_deep_btn"):
            # 使用人选的是预分析编号，对应当前 uploader 列表的 1-based 顺序
            try:
                viral_file = viral_videos[selected_viral - 1]
                own_file = own_videos[selected_own - 1]
            except Exception:
                st.error("原视频文件已不在当前页面，请重新上传后再进行深度对比。")
            else:
                try:
                    with st.spinner("Gemini 3.8 Flash 正在直接观看两条原视频，进行深度差距、重剪和补拍诊断…"):
                        result, meta = deep_compare(client, viral_file, own_file, category, product_name, user_points, viral_summary, own_summary)
                    st.session_state["sop2_deep_result"] = result
                    st.session_state["sop2_deep_meta"] = meta
                    st.session_state["sop2_deep_context"] = deep_context
                    append_history({
                        "module":"SOP2", "record_type":"爆款深度对比", "role":st.session_state["role"], "operator":st.session_state["operator"],
                        "tiktok_account":tiktok_account, "product_category":category, "product_name":product_name,
                        "input_selling_points":user_points, "viral_video_name":viral_file.name, "own_video_name":own_file.name,
                        "reference_video_index":selected_viral, "reference_video_name":viral_file.name,
                        "model_used":meta.get("model_used",""), "retry_count":meta.get("retry_count",""), "analysis_seconds":meta.get("analysis_seconds",""),
                        "diagnosis_summary":result.get("one_sentence_conclusion",""), "reedit_value":result.get("reedit_value",""),
                        "full_output_json":json_dumps(result),
                    })
                    st.success("深度对比完成。")
                except Exception as exc:
                    st.error(friendly_error(exc))

        deep = st.session_state.get("sop2_deep_result")
        if deep:
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

            st.markdown("### ⑧ 我的优势 / 我的劣势")
            a,b = st.columns(2)
            with a:
                st.markdown("**我的优势**")
                for x in deep.get("own_strengths",[]): st.markdown(f"- {x}")
            with b:
                st.markdown("**我的劣势**")
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
