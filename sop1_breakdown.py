from datetime import datetime
import html
import io

import streamlit as st

from config import PRODUCT_CATEGORIES, PERSPECTIVE_OPTIONS, SCENE_LIBRARY, MAX_COMPARE_VIDEOS
from common import clean_text, get_api_key, json_dumps, list_to_joined, make_signature, safe_int, video_batch_signature
from gemini_base import create_client, friendly_error
from gemini_sop1 import analyze_videos, generate_directions, generate_final_script
from history_service import append_history
from export_service import final_script_to_df, build_analysis_export_excel
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


SOP1_PAGE_ID = "sop1"
SOP1_WIDGET_KEYS = [
    "analysis_account", "analysis_category_mode", "analysis_product_name",
    "analysis_custom_category", "analysis_category_select",
    "analysis_user_selling_points", "selected_reference_video_index",
    "selected_direction_index", "final_selected_perspective", "final_selected_scene",
]
SOP1_STATE_KEYS = SOP1_WIDGET_KEYS + [
    "video_analysis_result", "video_analysis_meta", "video_batch_signature",
    "directions_result", "directions_meta", "directions_context_signature",
    "last_direction_control_signature", "final_script_result", "final_script_meta",
    "final_script_context_signature",
]
SOP1_DYNAMIC_PREFIXES = ("manual_selling_choice_", "effective_points_edit_")


def _storyboard_text(value):
    """Human-facing storyboard text; keep it clean and safe for HTML/SVG."""
    return html.escape(clean_text(value) or "—")


def _render_storyboard_overview(final_script_result, selected_scene="", selected_perspective=""):
    """Render one consolidated shooting-only storyboard overview for the whole script."""
    storyboard = final_script_result.get("storyboard", []) if isinstance(final_script_result, dict) else []
    if not storyboard:
        return

    scene = _storyboard_text(selected_scene)
    perspective = _storyboard_text(selected_perspective)

    st.markdown(
        """
        <style>
        .shoot-board {
            border: 1px solid #e5e7eb;
            border-radius: 16px;
            background: #ffffff;
            overflow: hidden;
            box-shadow: 0 2px 10px rgba(20, 24, 40, 0.05);
        }
        .shoot-board-head {
            padding: 16px 18px;
            background: #f7f8fa;
            border-bottom: 1px solid #e5e7eb;
            font-size: 14px;
            line-height: 1.7;
            color: #374151;
        }
        .shoot-board-head b { color: #111827; }
        .shoot-row {
            display: grid;
            grid-template-columns: 86px 1.05fr 2.2fr 1.35fr;
            border-bottom: 1px solid #edf0f3;
        }
        .shoot-row:last-child { border-bottom: 0; }
        .shoot-cell {
            padding: 13px 12px;
            font-size: 13px;
            line-height: 1.65;
            color: #374151;
            border-right: 1px solid #edf0f3;
            word-break: break-word;
        }
        .shoot-cell:last-child { border-right: 0; }
        .shoot-index {
            background: #fff5f5;
            color: #111827;
            font-weight: 800;
        }
        .shoot-time {
            display: block;
            margin-top: 4px;
            color: #ff4b4b;
            font-size: 12px;
            font-weight: 700;
        }
        .shoot-label {
            display: block;
            margin-bottom: 4px;
            color: #111827;
            font-weight: 800;
            font-size: 12px;
        }
        @media (max-width: 760px) {
            .shoot-row { grid-template-columns: 72px 1fr; }
            .shoot-cell:nth-child(2), .shoot-cell:nth-child(4) { border-right: 0; }
            .shoot-cell:nth-child(3), .shoot-cell:nth-child(4) { grid-column: 2; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    rows = []
    for index, item in enumerate(storyboard):
        sequence = _storyboard_text(item.get("sequence") or str(index + 1))
        time_range = _storyboard_text(item.get("time_range"))
        shot = _storyboard_text(item.get("shot"))
        visual = _storyboard_text(item.get("visual"))
        hand_action = _storyboard_text(item.get("hand_action"))
        rows.append(
            f"""
            <div class="shoot-row">
              <div class="shoot-cell shoot-index">第{sequence}镜<span class="shoot-time">{time_range}</span></div>
              <div class="shoot-cell"><span class="shoot-label">机位 / 景别</span>{shot}</div>
              <div class="shoot-cell"><span class="shoot-label">拍什么</span>{visual}</div>
              <div class="shoot-cell"><span class="shoot-label">怎么拍 / 手部动作</span>{hand_action}</div>
            </div>
            """
        )

    board = f"""
    <div class="shoot-board">
      <div class="shoot-board-head">
        <b>拍摄场景：</b>{scene}　　<b>主视角：</b>{perspective}<br/>
        一张总览完成一条拍摄脚本。这里只保留拍摄人员需要执行的内容，不包含剪辑、音效、设计逻辑。
      </div>
      {''.join(rows)}
    </div>
    """
    st.markdown(board, unsafe_allow_html=True)

def _build_storyboard_pdf(final_script_result, product_name="", selected_scene="", selected_perspective=""):
    """Build one long single-page PDF containing the whole shooting script only."""
    storyboard = final_script_result.get("storyboard", []) if isinstance(final_script_result, dict) else []
    if not storyboard:
        return b""

    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import mm
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    except Exception:
        return b""

    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    except Exception:
        pass

    # One final script = one PDF page. Estimate height from actual text so the
    # page stays compact instead of leaving a large blank area on phones.
    import math
    page_w = 210 * mm
    estimated_rows_mm = 0
    for item in storyboard:
        shot_text = clean_text(item.get("shot")) or "—"
        visual_text = clean_text(item.get("visual")) or "—"
        hand_text = clean_text(item.get("hand_action")) or "—"
        line_count = max(
            2,
            math.ceil(len(shot_text) / 11),
            math.ceil(len(visual_text) / 20),
            math.ceil(len(hand_text) / 12),
        )
        estimated_rows_mm += max(15, 6 + line_count * 5.2)
    page_h = (100 + estimated_rows_mm) * mm

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=(page_w, page_h),
        rightMargin=9 * mm,
        leftMargin=9 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
        title="TikTok拍摄分镜总览",
        author="TikTok爆款视频解析&复盘专用",
    )

    styles = getSampleStyleSheet()
    base_font = "STSong-Light"
    title_style = ParagraphStyle(
        "ShootBoardTitle",
        parent=styles["Title"],
        fontName=base_font,
        fontSize=18,
        leading=24,
        textColor=colors.HexColor("#111827"),
        alignment=TA_CENTER,
        spaceAfter=3 * mm,
    )
    sub_style = ParagraphStyle(
        "ShootBoardSub",
        parent=styles["BodyText"],
        fontName=base_font,
        fontSize=9.5,
        leading=14,
        textColor=colors.HexColor("#6B7280"),
        alignment=TA_CENTER,
        spaceAfter=4 * mm,
    )
    cell_style = ParagraphStyle(
        "ShootCell",
        parent=styles["BodyText"],
        fontName=base_font,
        fontSize=9.2,
        leading=14.5,
        textColor=colors.HexColor("#374151"),
    )
    head_cell_style = ParagraphStyle(
        "ShootHeadCell",
        parent=cell_style,
        fontSize=9.4,
        leading=14.5,
        textColor=colors.HexColor("#111827"),
    )

    product_title = html.escape(clean_text(product_name) or "TikTok 爆款拍摄方案")
    scene = html.escape(clean_text(selected_scene) or "—")
    perspective = html.escape(clean_text(selected_perspective) or "—")

    story = [
        Paragraph("拍摄分镜总览｜一条脚本一张", title_style),
        Paragraph(
            f"{product_title}　|　场景：{scene}　|　视角：{perspective}<br/>"
            "仅保留拍摄执行信息：镜头时间、机位景别、画面、手部动作。",
            sub_style,
        ),
    ]

    header = [
        Paragraph("镜头 / 时间", head_cell_style),
        Paragraph("机位 / 景别", head_cell_style),
        Paragraph("拍什么", head_cell_style),
        Paragraph("怎么拍 / 手部动作", head_cell_style),
    ]
    data = [header]

    for idx, item in enumerate(storyboard):
        seq = html.escape(clean_text(item.get("sequence")) or str(idx + 1))
        time_range = html.escape(clean_text(item.get("time_range")) or "—")
        shot = html.escape(clean_text(item.get("shot")) or "—").replace("\n", "<br/>")
        visual = html.escape(clean_text(item.get("visual")) or "—").replace("\n", "<br/>")
        hand_action = html.escape(clean_text(item.get("hand_action")) or "—").replace("\n", "<br/>")
        data.append([
            Paragraph(f"<b>第{seq}镜</b><br/><font color='#FF4B4B'>{time_range}</font>", cell_style),
            Paragraph(shot, cell_style),
            Paragraph(visual, cell_style),
            Paragraph(hand_action, cell_style),
        ])

    table = Table(
        data,
        colWidths=[25 * mm, 42 * mm, 78 * mm, 47 * mm],
        repeatRows=1,
        hAlign="CENTER",
    )
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F3F4F6")),
        ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#FFF5F5")),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#DDE1E6")),
        ("INNERGRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#E5E7EB")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.append(table)
    story.append(Spacer(1, 3 * mm))
    story.append(Paragraph("现场按第1镜到最后一镜顺序拍摄即可。", sub_style))

    doc.build(story)
    return buf.getvalue()

def _init_sop1_state():
    defaults = {
        "video_analysis_result": None, "video_analysis_meta": {}, "video_batch_signature": "",
        "selected_reference_video_index": None, "directions_result": None, "directions_meta": {},
        "directions_context_signature": "", "selected_direction_index": 0,
        "last_direction_control_signature": "", "final_script_result": None,
        "final_script_meta": {}, "final_script_context_signature": "",
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)

_init_sop1_state()
restore_page_memory(SOP1_PAGE_ID, SOP1_WIDGET_KEYS, SOP1_DYNAMIC_PREFIXES)
render_resume_notice(SOP1_PAGE_ID)
api_key = get_api_key()
client = create_client(api_key) if api_key else None
SOP1_ANALYSIS_TASK_KEY = task_key("sop1_video_analysis")
SOP1_DIRECTIONS_TASK_KEY = task_key("sop1_directions")
SOP1_FINAL_TASK_KEY = task_key("sop1_final_script")


def _create_background_client(api_key_value):
    """后台线程自己创建 Gemini client，避免跨线程复用主页面 client。"""
    if not api_key_value:
        raise RuntimeError("后台任务缺少 Gemini API Key。")
    return create_client(api_key_value)


def _run_analysis_background(api_key_value, task_videos, task_category, task_product_name, task_selling_points):
    bg_client = _create_background_client(api_key_value)
    return analyze_videos(
        bg_client,
        task_videos,
        task_category,
        task_product_name,
        task_selling_points,
    )


def _run_directions_background(
    api_key_value, task_category, task_product_name, task_summary, task_reference_video,
    task_input_selling_points, task_effective_selling_points, task_selling_point_mode,
):
    bg_client = _create_background_client(api_key_value)
    return generate_directions(
        bg_client,
        task_category,
        task_product_name,
        task_summary,
        task_reference_video,
        task_input_selling_points,
        task_effective_selling_points,
        task_selling_point_mode,
    )


def _run_final_background(
    api_key_value, task_category, task_product_name, task_reference_video,
    task_effective_selling_points, task_direction, task_scene, task_perspective,
):
    bg_client = _create_background_client(api_key_value)
    return generate_final_script(
        bg_client,
        task_category,
        task_product_name,
        task_reference_video,
        task_effective_selling_points,
        task_direction,
        task_scene,
        task_perspective,
    )
if not api_key:
    st.error("系统未配置 Gemini API Key，请联系管理员。")

st.caption("爆款拆解｜爆款分析 → 选择卖点 → 3个方向 → 最终拍摄脚本 → 单页拍摄分镜总览")

# --------------------------------------------------------
# ① 产品
# --------------------------------------------------------

st.markdown(
    "### ① 产品信息"
)

c1, c2, c3 = (
    st.columns(
        3
    )
)

with c1:

    tiktok_account = (
        st.text_input(
            "TikTok账号",
            key="analysis_account",
            placeholder="用于历史归档",
        )
    )

with c2:

    category_mode = st.radio(
        "产品品类",
        ["选择现有品类", "自定义输入"],
        horizontal=True,
        key="analysis_category_mode",
    )

with c3:

    product_name = (
        st.text_input(
            "产品名称 / SKU",
            key="analysis_product_name",
            placeholder="例如：隐私印章",
        )
    )

normal_categories = [x for x in PRODUCT_CATEGORIES if x != "其他 / 自定义"]
if category_mode == "自定义输入":
    custom_category = st.text_input(
        "自定义产品品类",
        key="analysis_custom_category",
        placeholder="直接输入，例如：LED阅读灯 / Flat Panel Book Light",
    )
    category = clean_text(custom_category) or "自定义品类（未填写）"
else:
    category = st.selectbox(
        "选择现有产品品类",
        normal_categories,
        key="analysis_category_select",
    )

input_selling_points = (
    st.text_area(
        "我们的真实产品卖点（选填）",
        key="analysis_user_selling_points",
        height=90,
        placeholder=(
            "可不填。"
            "填写后AI只负责比较和推荐，"
            "最终使用哪套卖点由你自己选择。"
        ),
    )
)

# --------------------------------------------------------
# ② 视频
# --------------------------------------------------------

st.markdown(
    "### ② 上传爆款视频"
)

new_uploaded_videos = (
    st.file_uploader(
        "支持同时上传 1-5 条 .mp4 视频",
        type=["mp4"],
        accept_multiple_files=True,
        key="analysis_videos",
    )
)

if new_uploaded_videos:
    remember_uploads(SOP1_PAGE_ID, "benchmark", new_uploaded_videos)

retained_uploaded_videos = get_retained_uploads(SOP1_PAGE_ID, "benchmark")
uploaded_videos = new_uploaded_videos or retained_uploaded_videos

if retained_uploaded_videos and not new_uploaded_videos:
    st.caption(
        "✅ 页面切换前的视频仍已保留："
        + "、".join(video.name for video in retained_uploaded_videos)
    )
    if st.button("清除已保留视频", key="sop1_clear_retained_uploads"):
        clear_retained_uploads(SOP1_PAGE_ID, "benchmark")
        st.session_state["video_batch_signature"] = ""
        st.session_state["video_analysis_result"] = None
        st.session_state["directions_result"] = None
        st.session_state["final_script_result"] = None
        new_workspace_id("SOP1")
        st.rerun()

if uploaded_videos:

    total_mb = (
        sum(
            len(
                video.getvalue()
            )
            for video
            in uploaded_videos
        )
        / 1024
        / 1024
    )

    st.caption(
        f"已上传 {len(uploaded_videos)} 条 "
        f"· 总大小 {total_mb:.2f} MB"
    )

    signature = (
        video_batch_signature(
            uploaded_videos,
            category,
            product_name,
            input_selling_points,
        )
    )

    if (
        st.session_state[
            "video_batch_signature"
        ]
        != signature
    ):

        previous_signature = st.session_state.get("video_batch_signature", "")
        if previous_signature:
            new_workspace_id("SOP1")

        st.session_state[
            "video_batch_signature"
        ] = signature

        st.session_state[
            "video_analysis_result"
        ] = None

        st.session_state[
            "directions_result"
        ] = None

        st.session_state[
            "directions_context_signature"
        ] = ""

        st.session_state[
            "final_script_result"
        ] = None

        st.session_state[
            "final_script_context_signature"
        ] = ""

        st.session_state[
            "selected_reference_video_index"
        ] = None

        st.session_state[
            "selected_direction_index"
        ] = 0

analyze_button = (
    st.button(
        "解析爆款视频",
        type="primary",
        use_container_width=True,

        disabled=(
            client is None
            or not uploaded_videos
            or len(
                uploaded_videos
            )
            > MAX_COMPARE_VIDEOS
        ),

        key="run_video_analysis",
    )
)

# 爆款解析后台结果回收。页面切换时任务仍由线程池继续执行。
analysis_task = take_background_task(SOP1_ANALYSIS_TASK_KEY)
if analysis_task:
    task_context = analysis_task.get("context", {})
    if analysis_task.get("status") == "done":
        result, metadata = analysis_task.get("result")
        if task_context.get("input_signature") == st.session_state.get("video_batch_signature"):
            st.session_state["video_analysis_result"] = result
            st.session_state["video_analysis_meta"] = metadata
            recommended = safe_int(result.get("recommended_reference_video_index", 1), 1)
            st.session_state["selected_reference_video_index"] = recommended
            append_history({
                "record_type": "爆款对比解析",
                "role": task_context.get("role", ""),
                "operator": task_context.get("operator", ""),
                "tiktok_account": task_context.get("tiktok_account", ""),
                "product_category": task_context.get("category", ""),
                "product_name": task_context.get("product_name", ""),
                "input_selling_points": task_context.get("input_selling_points", ""),
                "video_names": task_context.get("video_names", ""),
                "video_count": task_context.get("video_count", ""),
                "model_used": metadata.get("model_used", ""),
                "fallback_used": metadata.get("fallback_used", ""),
                "retry_count": metadata.get("retry_count", ""),
                "analysis_seconds": metadata.get("analysis_seconds", ""),
                "full_output_json": json_dumps(result),
            })
            st.success("爆款视频解析已在后台完成。")
        else:
            st.warning("刚才的后台爆款解析已完成，但当前上传内容已经变化，因此没有覆盖当前页面。")
    elif analysis_task.get("status") == "error":
        st.error(friendly_error(analysis_task.get("error")))

if analyze_button:
    task_videos = get_retained_uploads(SOP1_PAGE_ID, "benchmark") or uploaded_videos
    submitted = submit_background_task(
        SOP1_ANALYSIS_TASK_KEY,
        _run_analysis_background,
        api_key, task_videos, category, product_name, input_selling_points,
        context={
            "input_signature": st.session_state.get("video_batch_signature", ""),
            "role": st.session_state.get("role", ""),
            "operator": st.session_state.get("operator", ""),
            "tiktok_account": tiktok_account,
            "category": category,
            "product_name": product_name,
            "input_selling_points": input_selling_points,
            "video_names": " | ".join(video.name for video in task_videos),
            "video_count": len(task_videos),
        },
    )
    if submitted:
        print(
            f"[SOP1 SUBMIT] videos={len(task_videos)} size_mb="
            f"{sum(len(v.getvalue()) for v in task_videos) / 1024 / 1024:.2f}",
            flush=True,
        )
        st.rerun()
    else:
        st.info("这项爆款解析已经在后台运行，无需重复提交。")

render_background_task_status(
    SOP1_ANALYSIS_TASK_KEY,
    "正在后台逐条拆解爆款视频并提取独立卖点。",
)

analysis_result = (
    st.session_state.get(
        "video_analysis_result"
    )
)

# --------------------------------------------------------
# ③ 拆解
# --------------------------------------------------------

if analysis_result:

    summary = (
        analysis_result.get(
            "comparison_summary",
            {},
        )
    )

    videos = (
        analysis_result.get(
            "videos",
            [],
        )
    )

    st.markdown(
        "### ③ 中文爆款拆解"
    )

    st.info(
        summary.get(
            "one_sentence_core",
            "",
        )
    )

    with st.expander(
        "查看完整爆款拆解",
        expanded=False,
    ):

        st.markdown(
            "**共同爆款脚本路线**"
        )

        st.write(
            summary.get(
                "common_script_route",
                "",
            )
        )

        detail1, detail2 = (
            st.columns(
                2
            )
        )

        with detail1:

            st.markdown(
                "**共同人群画像**"
            )

            st.write(
                summary.get(
                    "common_audience",
                    "",
                )
            )

        with detail2:

            st.markdown(
                "**年龄预估**"
            )

            st.write(
                summary.get(
                    "age_estimate",
                    "",
                )
            )

        st.markdown(
            "**共同前3秒 Hook**"
        )

        st.write(
            summary.get(
                "common_hook_pattern",
                "",
            )
        )

        st.markdown(
            "**共同画面与节奏**"
        )

        st.write(
            summary.get(
                "visual_rhythm",
                "",
            )
        )

        common_strengths = summary.get("common_strengths", [])
        common_weaknesses = summary.get("common_weaknesses", [])
        if common_strengths or common_weaknesses:
            strength_col, weakness_col = st.columns(2)
            with strength_col:
                st.markdown("**共同优势**")
                for item in common_strengths:
                    st.markdown(f"- {item}")
            with weakness_col:
                st.markdown("**共同短板 / 风险**")
                for item in common_weaknesses:
                    st.markdown(f"- {item}")

        st.markdown(
            "**最值得共同吸收的3点**"
        )

        for item in summary.get(
            "top_absorb_points",
            [],
        ):

            st.markdown(
                f"- {item}"
            )

        if len(videos) > 1:

            st.markdown(
                "**多视频关键差异**"
            )

            st.write(
                summary.get(
                    "key_differences",
                    "",
                )
            )

        st.divider()

        st.markdown(
            "**逐条视频拆解**"
        )

        for video in videos:

            with st.expander(
                (
                    f'视频 {video.get("video_index", "")}'
                    f'｜{video.get("filename", "")}'
                ),
                expanded=False,
            ):

                st.markdown(
                    f'**一句话核心：** '
                    f'{video.get("one_sentence_core", "")}'
                )

                st.markdown(
                    "**该视频独立卖点**"
                )

                for item in video.get(
                    "inferred_selling_points",
                    [],
                ):

                    st.markdown(
                        f"- {item}"
                    )

                st.markdown(
                    "**爆款脚本路线**"
                )

                st.write(
                    video.get(
                        "script_route",
                        "",
                    )
                )

                st.markdown(
                    "**人群画像**"
                )

                st.write(
                    video.get(
                        "audience_profile",
                        "",
                    )
                )

                st.markdown(
                    "**年龄预估**"
                )

                st.write(
                    video.get(
                        "age_estimate",
                        "",
                    )
                )

                st.markdown(
                    "**前3秒 Hook**"
                )

                st.write(
                    video.get(
                        "first_3s_hook",
                        "",
                    )
                )

                st.markdown(
                    "**画面与节奏**"
                )

                st.write(
                    video.get(
                        "visual_rhythm",
                        "",
                    )
                )

                subtitle_segments = video.get(
                    "subtitle_segments",
                    [],
                )
                actual_duration = video.get(
                    "actual_duration_seconds",
                    "",
                )

                if subtitle_segments:
                    st.markdown(
                        "**完整英 / 西双语字幕时间轴**"
                    )
                    if actual_duration not in (None, ""):
                        st.caption(
                            f"AI识别原视频实际时长约 {actual_duration} 秒；字幕按原视频完整时长生成，不再按30-40秒截断。"
                        )

                    subtitle_rows = [
                        {
                            "段序": segment.get("segment_no", index + 1),
                            "时间段": segment.get("time_range", ""),
                            "English": segment.get("copy_en", ""),
                            "Español": segment.get("copy_es", ""),
                        }
                        for index, segment in enumerate(subtitle_segments)
                    ]
                    st.dataframe(
                        subtitle_rows,
                        hide_index=True,
                        use_container_width=True,
                    )

                strengths = video.get("strengths", [])
                weaknesses = video.get("weaknesses", [])
                if strengths or weaknesses:
                    strength_col, weakness_col = st.columns(2)
                    with strength_col:
                        st.markdown("**优点 / 做得好的地方**")
                        for item in strengths:
                            st.markdown(f"- {item}")
                    with weakness_col:
                        st.markdown("**短板 / 可优化点**")
                        for item in weaknesses:
                            st.markdown(f"- {item}")

                st.markdown(
                    "**最值得吸收的3点**"
                )

                for item in video.get(
                    "top_absorb_points",
                    [],
                ):

                    st.markdown(
                        f"- {item}"
                    )

                st.markdown(
                    "**参考价值判断**"
                )

                st.write(
                    video.get(
                        "fit_reason",
                        "",
                    )
                )

                st.markdown(
                    f'**推荐指数：** '
                    f'{video.get("recommend_score", "")}'
                )

    # ----------------------------------------------------
    # ④ 主参考视频
    # ----------------------------------------------------

    st.markdown(
        "### ④ 选择主参考视频"
    )

    recommended_index = (
        safe_int(
            analysis_result.get(
                "recommended_reference_video_index",
                1,
            ),
            1,
        )
    )

    video_indices = [

        safe_int(
            video.get(
                "video_index"
            ),
            index + 1,
        )

        for index, video
        in enumerate(
            videos
        )
    ]

    if (
        st.session_state[
            "selected_reference_video_index"
        ]
        not in video_indices
    ):

        st.session_state[
            "selected_reference_video_index"
        ] = recommended_index

    selected_ref_index = (
        st.radio(
            "请选择本次主要参考的视频",

            options=video_indices,

            format_func=lambda value: (
                f'视频{value}｜'
                f'{next((v.get("filename", "") for v in videos if safe_int(v.get("video_index"), 0) == value), "")}'
                + (
                    "（AI推荐）"
                    if value
                    == recommended_index
                    else ""
                )
            ),

            key="selected_reference_video_index",
        )
    )

    chosen_ref_video = next(
        (
            video
            for video
            in videos
            if safe_int(
                video.get(
                    "video_index"
                ),
                0,
            )
            == selected_ref_index
        ),
        videos[0],
    )

    st.caption(
        "AI推荐依据："
    )

    st.write(
        chosen_ref_video.get(
            "fit_reason",
            "",
        )
    )

    # ----------------------------------------------------
    # ⑤ 卖点选择
    # ----------------------------------------------------

    st.markdown(
        "### ⑤ 卖点参考逻辑"
    )

    selected_video_points = (
        chosen_ref_video.get(
            "inferred_selling_points",
            [],
        )
    )

    if not selected_video_points:

        selected_video_points = (
            analysis_result.get(
                "common_inferred_selling_points",
                [],
            )
        )

    viral_points_text = (
        list_to_joined(
            selected_video_points
        )
    )

    st.markdown(
        f'**当前参考：视频{selected_ref_index}'
        f'｜{chosen_ref_video.get("filename", "")}**'
    )

    st.markdown(
        "**该视频推理出的核心卖点：**"
    )

    for item in selected_video_points:

        st.markdown(
            f"- {item}"
        )

    # ----------------------------------------------------
    # 没有人工卖点
    # ----------------------------------------------------

    if not clean_text(
        input_selling_points
    ):

        st.info(
            "你没有填写自己的产品卖点，"
            "因此当前使用该爆款视频推理卖点。"
        )

        selling_point_mode = (
            "viral_first"
        )

        default_effective_points = (
            viral_points_text
        )

    # ----------------------------------------------------
    # 有人工卖点：永远保留选择权
    # ----------------------------------------------------

    else:

        relation = (
            clean_text(
                chosen_ref_video.get(
                    "selling_point_relation",
                    "",
                )
            )
        )

        relation_reason = (
            clean_text(
                chosen_ref_video.get(
                    "selling_point_relation_reason",
                    "",
                )
            )
        )

        suggested_mode = (
            clean_text(
                chosen_ref_video.get(
                    "suggested_mode",
                    "blend",
                )
            )
        )

        if suggested_mode not in {
            "viral_first",
            "user_first",
            "blend",
        }:

            suggested_mode = (
                "blend"
            )

        recommendation_label = {

            "viral_first":
                "以当前爆款视频卖点为主",

            "user_first":
                "以我的真实产品卖点为主",

            "blend":
                "融合两者",

        }[
            suggested_mode
        ]

        st.info(
            f"AI推荐：{recommendation_label}"
        )

        if relation_reason:

            st.caption(
                f"推荐原因：{relation_reason}"
            )

        # =================================================
        # 最重要：
        # AI不给你自动做主
        # 必须人工选择
        # =================================================

        selling_choice_key = (
            f"manual_selling_choice_"
            f"{selected_ref_index}"
        )

        selling_point_mode = (
            st.radio(
                "请由使用人决定本次采用哪套卖点",

                options=[
                    "viral_first",
                    "user_first",
                    "blend",
                ],

                index=None,

                format_func=lambda mode: {

                    "viral_first":
                        "以当前爆款视频卖点为主",

                    "user_first":
                        "以我的真实产品卖点为主",

                    "blend":
                        "融合两者",

                }[mode],

                key=selling_choice_key,
            )
        )

        if selling_point_mode is None:

            default_effective_points = ""

        elif (
            selling_point_mode
            == "viral_first"
        ):

            default_effective_points = (
                viral_points_text
            )

        elif (
            selling_point_mode
            == "user_first"
        ):

            default_effective_points = (
                clean_text(
                    input_selling_points
                )
            )

        else:

            default_effective_points = (
                clean_text(
                    chosen_ref_video.get(
                        "blended_selling_points",
                        "",
                    )
                )
            )

            if not default_effective_points:

                default_effective_points = (
                    f"{viral_points_text}; "
                    f"{clean_text(input_selling_points)}"
                )

    # ----------------------------------------------------
    # 最终卖点人工可编辑
    # ----------------------------------------------------

    if selling_point_mode is not None:

        effective_edit_key = (
            f"effective_points_edit_"
            f"{selected_ref_index}_"
            f"{selling_point_mode}"
        )

        if (
            effective_edit_key
            not in st.session_state
        ):

            st.session_state[
                effective_edit_key
            ] = default_effective_points

        effective_selling_points = (
            st.text_area(
                "最终用于生成脚本的卖点（可编辑）",
                height=110,
                key=effective_edit_key,
                help=(
                    "这里是最终送给AI生成三个方向的卖点。"
                    "你可以继续人工修改。"
                ),
            )
        )

    else:

        effective_selling_points = ""

        st.warning(
            "请先选择上面的卖点参考方式，"
            "再生成3个拍摄方向。"
        )

    current_direction_context = (
        make_signature(
            selected_ref_index,
            chosen_ref_video,
            selling_point_mode,
            effective_selling_points,
        )
    )

    old_direction_context = (
        st.session_state.get(
            "directions_context_signature",
            "",
        )
    )

    if (
        old_direction_context
        and old_direction_context
        != current_direction_context
    ):

        st.session_state[
            "directions_result"
        ] = None

        st.session_state[
            "directions_context_signature"
        ] = ""

        st.session_state[
            "final_script_result"
        ] = None

        st.session_state[
            "final_script_context_signature"
        ] = ""

        st.session_state[
            "selected_direction_index"
        ] = 0

    # ----------------------------------------------------
    # ⑥ 方向
    # ----------------------------------------------------

    st.markdown(
        "### ⑥ 生成 3 个参考方向"
    )

    generate_direction_button = (
        st.button(
            "生成3个拍摄方向",
            type="primary",
            use_container_width=True,

            disabled=(
                client is None
                or selling_point_mode is None
                or not clean_text(
                    effective_selling_points
                )
            ),

            key="generate_directions_button",
        )
    )

    directions_task = take_background_task(SOP1_DIRECTIONS_TASK_KEY)
    if directions_task:
        task_context = directions_task.get("context", {})
        if directions_task.get("status") == "done":
            directions_result, directions_meta = directions_task.get("result")
            if task_context.get("direction_context") == current_direction_context:
                st.session_state["directions_result"] = directions_result
                st.session_state["directions_meta"] = directions_meta
                st.session_state["directions_context_signature"] = current_direction_context
                st.session_state["selected_direction_index"] = 0
                st.session_state["final_script_result"] = None
                append_history({
                    "record_type": "3方向生成",
                    "role": task_context.get("role", ""),
                    "operator": task_context.get("operator", ""),
                    "tiktok_account": task_context.get("tiktok_account", ""),
                    "product_category": task_context.get("category", ""),
                    "product_name": task_context.get("product_name", ""),
                    "input_selling_points": task_context.get("input_selling_points", ""),
                    "inferred_selling_points": task_context.get("viral_points_text", ""),
                    "effective_selling_points": task_context.get("effective_selling_points", ""),
                    "selling_point_mode": task_context.get("selling_point_mode", ""),
                    "reference_video_index": task_context.get("selected_ref_index", ""),
                    "reference_video_name": task_context.get("reference_video_name", ""),
                    "model_used": directions_meta.get("model_used", ""),
                    "fallback_used": directions_meta.get("fallback_used", ""),
                    "retry_count": directions_meta.get("retry_count", ""),
                    "analysis_seconds": directions_meta.get("analysis_seconds", ""),
                    "full_output_json": json_dumps(directions_result),
                })
                st.success("3个拍摄方向已在后台生成。")
            else:
                st.warning("刚才的后台方向生成已完成，但当前参考视频或卖点选择已经变化，因此没有覆盖当前结果。")
        elif directions_task.get("status") == "error":
            st.error(friendly_error(directions_task.get("error")))

    if generate_direction_button:
        submitted = submit_background_task(
            SOP1_DIRECTIONS_TASK_KEY,
            _run_directions_background,
            api_key, category, product_name, summary, chosen_ref_video,
            input_selling_points, effective_selling_points, selling_point_mode,
            context={
                "direction_context": current_direction_context,
                "role": st.session_state.get("role", ""),
                "operator": st.session_state.get("operator", ""),
                "tiktok_account": tiktok_account,
                "category": category,
                "product_name": product_name,
                "input_selling_points": input_selling_points,
                "viral_points_text": viral_points_text,
                "effective_selling_points": effective_selling_points,
                "selling_point_mode": selling_point_mode,
                "selected_ref_index": selected_ref_index,
                "reference_video_name": chosen_ref_video.get("filename", ""),
            },
        )
        if submitted:
            st.rerun()
        else:
            st.info("3个拍摄方向已经在后台生成中，无需重复提交。")

    render_background_task_status(
        SOP1_DIRECTIONS_TASK_KEY,
        "正在后台生成3个不同拍摄方向。",
    )

    directions_result = (
        st.session_state.get(
            "directions_result"
        )
    )

    # ----------------------------------------------------
    # ⑦ 方向选择
    # ----------------------------------------------------

    if directions_result:

        directions = (
            directions_result.get(
                "directions",
                [],
            )
        )

        if directions:

            st.markdown(
                "### ⑦ 选择 1 个方向"
            )

            direction_indices = (
                list(
                    range(
                        len(
                            directions
                        )
                    )
                )
            )

            if (
                st.session_state[
                    "selected_direction_index"
                ]
                not in direction_indices
            ):

                st.session_state[
                    "selected_direction_index"
                ] = 0

            selected_direction_index = (
                st.radio(
                    "请选择你要继续生成最终脚本的方向",

                    options=direction_indices,

                    format_func=lambda index: (
                        f'方向{index + 1}｜'
                        f'{directions[index].get("direction_name", "")}'
                    ),

                    key="selected_direction_index",
                )
            )

            chosen_direction = (
                directions[
                    selected_direction_index
                ]
            )

            # 不再使用 tabs
            # 选哪个只显示哪个

            st.divider()

            st.markdown(
                f'### 方向{selected_direction_index + 1}'
                f'｜{chosen_direction.get("direction_name", "")}'
            )

            st.markdown(
                f'**核心思路：** '
                f'{chosen_direction.get("core_idea", "")}'
            )

            d1, d2 = (
                st.columns(
                    2
                )
            )

            with d1:

                st.markdown(
                    f'**目标人群：** '
                    f'{chosen_direction.get("target_audience", "")}'
                )

                st.markdown(
                    f'**前3秒 Hook：** '
                    f'{chosen_direction.get("hook", "")}'
                )

                st.markdown(
                    f'**产品切入方式：** '
                    f'{chosen_direction.get("product_entry", "")}'
                )

            with d2:

                st.markdown(
                    f'**推荐视角：** '
                    f'{chosen_direction.get("recommended_perspective", "")}'
                )

                st.markdown(
                    f'**推荐小场景：** '
                    f'{chosen_direction.get("recommended_scene", "")}'
                )

            st.markdown(
                "**可吸收点**"
            )

            for item in chosen_direction.get(
                "absorb_points",
                [],
            ):

                st.markdown(
                    f"- {item}"
                )

            st.markdown(
                "**差异化点**"
            )

            for item in chosen_direction.get(
                "differentiation_points",
                [],
            ):

                st.markdown(
                    f"- {item}"
                )

            # ------------------------------------------------
            # 方向改变后刷新场景/视角
            # ------------------------------------------------

            control_signature = (
                make_signature(
                    current_direction_context,
                    selected_direction_index,
                    chosen_direction,
                )
            )

            if (
                st.session_state.get(
                    "last_direction_control_signature",
                    "",
                )
                != control_signature
            ):

                recommended_perspective = (
                    clean_text(
                        chosen_direction.get(
                            "recommended_perspective",
                            "",
                        )
                    )
                )

                if (
                    "第三"
                    in recommended_perspective
                ):

                    st.session_state[
                        "final_selected_perspective"
                    ] = PERSPECTIVE_OPTIONS[1]

                else:

                    st.session_state[
                        "final_selected_perspective"
                    ] = PERSPECTIVE_OPTIONS[0]

                recommended_scene = (
                    clean_text(
                        chosen_direction.get(
                            "recommended_scene",
                            "",
                        )
                    )
                )

                if (
                    recommended_scene
                    in SCENE_LIBRARY
                ):

                    st.session_state[
                        "final_selected_scene"
                    ] = recommended_scene

                else:

                    st.session_state[
                        "final_selected_scene"
                    ] = list(
                        SCENE_LIBRARY.keys()
                    )[0]

                st.session_state[
                    "last_direction_control_signature"
                ] = control_signature

                st.session_state[
                    "final_script_result"
                ] = None

            # ------------------------------------------------
            # ⑧ 最终脚本
            # ------------------------------------------------

            st.markdown(
                "### ⑧ 生成最终拍摄脚本"
            )

            script1, script2 = (
                st.columns(
                    2
                )
            )

            with script1:

                selected_perspective = (
                    st.radio(
                        "拍摄视角",
                        PERSPECTIVE_OPTIONS,
                        key="final_selected_perspective",
                    )
                )

            with script2:

                selected_scene = (
                    st.selectbox(
                        "实际拍摄小场景",
                        list(
                            SCENE_LIBRARY.keys()
                        ),
                        key="final_selected_scene",
                    )
                )

            st.caption(
                "固定限制：真人不露脸 / 不正面出镜；"
                "允许手、手臂和少量身体局部。"
            )

            current_final_context = (
                make_signature(
                    current_direction_context,
                    selected_direction_index,
                    selected_perspective,
                    selected_scene,
                )
            )

            old_final_context = (
                st.session_state.get(
                    "final_script_context_signature",
                    "",
                )
            )

            if (
                old_final_context
                and old_final_context
                != current_final_context
            ):

                st.session_state[
                    "final_script_result"
                ] = None

                st.session_state[
                    "final_script_context_signature"
                ] = ""

            generate_script_button = (
                st.button(
                    "生成最终拍摄脚本",
                    type="primary",
                    use_container_width=True,
                    disabled=(
                        client is None
                    ),
                    key="generate_final_script_button",
                )
            )

            final_task = take_background_task(SOP1_FINAL_TASK_KEY)
            if final_task:
                task_context = final_task.get("context", {})
                if final_task.get("status") == "done":
                    final_result, final_meta = final_task.get("result")
                    if task_context.get("final_context") == current_final_context:
                        st.session_state["final_script_result"] = final_result
                        st.session_state["final_script_meta"] = final_meta
                        st.session_state["final_script_context_signature"] = current_final_context
                        append_history({
                            "record_type": "最终拍摄脚本",
                            "role": task_context.get("role", ""),
                            "operator": task_context.get("operator", ""),
                            "tiktok_account": task_context.get("tiktok_account", ""),
                            "product_category": task_context.get("category", ""),
                            "product_name": task_context.get("product_name", ""),
                            "input_selling_points": task_context.get("input_selling_points", ""),
                            "inferred_selling_points": task_context.get("viral_points_text", ""),
                            "effective_selling_points": task_context.get("effective_selling_points", ""),
                            "selling_point_mode": task_context.get("selling_point_mode", ""),
                            "reference_video_index": task_context.get("selected_ref_index", ""),
                            "reference_video_name": task_context.get("reference_video_name", ""),
                            "direction_name": task_context.get("direction_name", ""),
                            "selected_scene": task_context.get("selected_scene", ""),
                            "selected_perspective": task_context.get("selected_perspective", ""),
                            "model_used": final_meta.get("model_used", ""),
                            "fallback_used": final_meta.get("fallback_used", ""),
                            "retry_count": final_meta.get("retry_count", ""),
                            "analysis_seconds": final_meta.get("analysis_seconds", ""),
                            "full_output_json": json_dumps(final_result),
                        })
                        st.success("最终拍摄脚本已在后台生成。")
                    else:
                        st.warning("刚才的后台脚本生成已完成，但当前方向/场景/视角已经变化，因此没有覆盖当前结果。")
                elif final_task.get("status") == "error":
                    st.error(friendly_error(final_task.get("error")))

            if generate_script_button:
                submitted = submit_background_task(
                    SOP1_FINAL_TASK_KEY,
                    _run_final_background,
                    api_key, category, product_name, chosen_ref_video, effective_selling_points,
                    chosen_direction, selected_scene, selected_perspective,
                    context={
                        "final_context": current_final_context,
                        "role": st.session_state.get("role", ""),
                        "operator": st.session_state.get("operator", ""),
                        "tiktok_account": tiktok_account,
                        "category": category,
                        "product_name": product_name,
                        "input_selling_points": input_selling_points,
                        "viral_points_text": viral_points_text,
                        "effective_selling_points": effective_selling_points,
                        "selling_point_mode": selling_point_mode,
                        "selected_ref_index": selected_ref_index,
                        "reference_video_name": chosen_ref_video.get("filename", ""),
                        "direction_name": chosen_direction.get("direction_name", ""),
                        "selected_scene": selected_scene,
                        "selected_perspective": selected_perspective,
                    },
                )
                if submitted:
                    st.rerun()
                else:
                    st.info("最终拍摄脚本已经在后台生成中，无需重复提交。")

            render_background_task_status(
                SOP1_FINAL_TASK_KEY,
                "正在后台生成中文可执行拍摄脚本。",
            )

    # ----------------------------------------------------
    # ⑨ 最终脚本
    # ----------------------------------------------------

    final_script_result = (
        st.session_state.get(
            "final_script_result"
        )
    )

    if final_script_result:

        st.markdown(
            "### ⑨ 最终拍摄脚本"
        )

        shooting_notes = (
            clean_text(
                final_script_result.get(
                    "shooting_notes",
                    "",
                )
            )
        )

        if shooting_notes:

            st.info(
                shooting_notes
            )

        final_dataframe = (
            final_script_to_df(
                final_script_result
            )
        )

        st.dataframe(
            final_dataframe,
            hide_index=True,
            use_container_width=True,
        )

        excel_data = (
            build_analysis_export_excel(
                analysis_result,
                st.session_state.get(
                    "directions_result"
                ),
                final_script_result,
            )
        )

        st.download_button(
            "一键导出 Excel",

            data=excel_data,

            file_name=(
                "TikTok爆款解析_拍摄脚本_"
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

            type="primary",

            use_container_width=True,
        )

        # ------------------------------------------------
        # ⑩ 拍摄分镜总览（单页）
        # ------------------------------------------------
        st.markdown("### ⑩ 拍摄分镜总览（单页拍摄版）")
        st.caption(
            "一条最终拍摄脚本只生成一张总览；不再拆成多张分镜卡，也不显示剪辑/音效/设计逻辑。"
        )

        _render_storyboard_overview(
            final_script_result,
            selected_scene=selected_scene,
            selected_perspective=selected_perspective,
        )

        storyboard_pdf = _build_storyboard_pdf(
            final_script_result,
            product_name=product_name,
            selected_scene=selected_scene,
            selected_perspective=selected_perspective,
        )

        if storyboard_pdf:
            st.download_button(
                "下载拍摄分镜总览（PDF｜单页）",
                data=storyboard_pdf,
                file_name=(
                    "TikTok拍摄分镜总览_"
                    + datetime.now().strftime("%Y%m%d_%H%M")
                    + ".pdf"
                ),
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )
        else:
            st.caption("PDF组件尚未安装，请确认 requirements.txt 中包含 reportlab。")

# --------------------------------------------------------
# 页面状态与“工作记录”自动保存
# --------------------------------------------------------
snapshot_page_memory(SOP1_PAGE_ID, SOP1_WIDGET_KEYS, SOP1_DYNAMIC_PREFIXES)

_workspace_analysis = st.session_state.get("video_analysis_result") or {}
_workspace_directions = st.session_state.get("directions_result") or {}
_workspace_final = st.session_state.get("final_script_result") or {}
_workspace_ref_index = st.session_state.get("selected_reference_video_index")
_workspace_ref_name = ""
for _video in _workspace_analysis.get("videos", []) if isinstance(_workspace_analysis, dict) else []:
    if safe_int(_video.get("video_index"), 0) == safe_int(_workspace_ref_index, 0):
        _workspace_ref_name = clean_text(_video.get("filename", ""))
        break

_workspace_direction_name = ""
_workspace_direction_index = safe_int(st.session_state.get("selected_direction_index"), 0)
_workspace_direction_list = _workspace_directions.get("directions", []) if isinstance(_workspace_directions, dict) else []
if 0 <= _workspace_direction_index < len(_workspace_direction_list):
    _workspace_direction_name = clean_text(_workspace_direction_list[_workspace_direction_index].get("direction_name", ""))

if _workspace_final:
    _workspace_step = "最终拍摄脚本已生成"
elif _workspace_directions:
    _workspace_step = "3个方向已生成" + (f" · 已选：{_workspace_direction_name}" if _workspace_direction_name else "")
elif _workspace_analysis:
    _workspace_step = "爆款拆解已完成" + (f" · 主参考：{_workspace_ref_name}" if _workspace_ref_name else "")
elif uploaded_videos:
    _workspace_step = "视频已上传，待解析"
else:
    _workspace_step = "产品信息填写中"

_workspace_video_names = " | ".join(video.name for video in (uploaded_videos or []))
_workspace_meaningful = bool(
    uploaded_videos
    or clean_text(tiktok_account)
    or clean_text(product_name)
    or clean_text(input_selling_points)
    or _workspace_analysis
    or _workspace_directions
    or _workspace_final
)

save_workspace_record(
    module="SOP1",
    page_id=SOP1_PAGE_ID,
    page_path="sop1_breakdown.py",
    step=_workspace_step,
    state_keys=SOP1_STATE_KEYS,
    dynamic_prefixes=SOP1_DYNAMIC_PREFIXES,
    tiktok_account=tiktok_account,
    product_category=category,
    product_name=product_name,
    input_selling_points=input_selling_points,
    video_names=_workspace_video_names,
    reference_video_name=_workspace_ref_name,
    direction_name=_workspace_direction_name,
    selected_scene=clean_text(st.session_state.get("final_selected_scene", "")),
    selected_perspective=clean_text(st.session_state.get("final_selected_perspective", "")),
    meaningful=_workspace_meaningful,
)

