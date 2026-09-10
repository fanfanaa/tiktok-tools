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


def _render_storyboard_cards(final_script_result):
    """Render the final script as visual Chinese cards for filming/editing."""
    storyboard = final_script_result.get("storyboard", []) if isinstance(final_script_result, dict) else []
    if not storyboard:
        return

    st.markdown(
        """
        <style>
        .shot-card {
            border: 1px solid #e7e9ee;
            border-radius: 16px;
            background: #ffffff;
            padding: 18px 18px 16px 18px;
            margin: 0 0 16px 0;
            box-shadow: 0 2px 10px rgba(20, 24, 40, 0.05);
            min-height: 340px;
        }
        .shot-card-head {
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:12px;
            margin-bottom:12px;
            padding-bottom:10px;
            border-bottom:1px solid #f0f1f4;
        }
        .shot-card-title {
            font-size:18px;
            font-weight:800;
            color:#1f2937;
        }
        .shot-card-time {
            flex:0 0 auto;
            font-size:13px;
            font-weight:700;
            color:#ff4b4b;
            background:#fff0f0;
            border-radius:999px;
            padding:5px 10px;
        }
        .shot-card-row {
            margin:9px 0;
            font-size:14px;
            line-height:1.65;
            color:#374151;
        }
        .shot-card-label {
            font-weight:800;
            color:#111827;
            margin-right:5px;
        }
        .shot-card-copy {
            margin-top:12px;
            padding:12px 14px;
            border-radius:12px;
            background:#f7f8fa;
            font-size:15px;
            line-height:1.7;
            color:#111827;
            font-weight:650;
        }
        .shot-card-foot {
            margin-top:12px;
            padding-top:10px;
            border-top:1px dashed #e5e7eb;
            font-size:13px;
            line-height:1.6;
            color:#6b7280;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    cols = st.columns(2)
    for index, item in enumerate(storyboard):
        sequence = _storyboard_text(item.get("sequence") or str(index + 1))
        time_range = _storyboard_text(item.get("time_range"))
        shot = _storyboard_text(item.get("shot"))
        visual = _storyboard_text(item.get("visual"))
        hand_action = _storyboard_text(item.get("hand_action"))
        copy_cn = _storyboard_text(item.get("copy_cn"))
        audio = _storyboard_text(item.get("audio"))
        rationale = _storyboard_text(item.get("rationale"))

        card = f"""
        <div class="shot-card">
          <div class="shot-card-head">
            <div class="shot-card-title">🎬 第{sequence}镜</div>
            <div class="shot-card-time">{time_range}</div>
          </div>
          <div class="shot-card-row"><span class="shot-card-label">📷 机位 / 景别：</span>{shot}</div>
          <div class="shot-card-row"><span class="shot-card-label">🖼️ 画面：</span>{visual}</div>
          <div class="shot-card-row"><span class="shot-card-label">✋ 手部动作：</span>{hand_action}</div>
          <div class="shot-card-copy"><span class="shot-card-label">💬 中文字幕 / 中文口播：</span>{copy_cn}</div>
          <div class="shot-card-row"><span class="shot-card-label">🎵 剪辑 / 音效：</span>{audio}</div>
          <div class="shot-card-foot"><span class="shot-card-label">🎯 这一镜为什么拍：</span>{rationale}</div>
        </div>
        """
        with cols[index % 2]:
            st.markdown(card, unsafe_allow_html=True)


def _wrap_storyboard_lines(value, width=24, max_lines=8):
    text = clean_text(value) or "—"
    # Chinese storyboard text is best wrapped by character count, while preserving existing line breaks.
    lines = []
    for raw_line in str(text).splitlines() or [str(text)]:
        raw_line = raw_line.strip()
        if not raw_line:
            continue
        while raw_line:
            lines.append(raw_line[:width])
            raw_line = raw_line[width:]
            if len(lines) >= max_lines:
                break
        if len(lines) >= max_lines:
            break
    if not lines:
        lines = ["—"]
    return lines


def _svg_text_block(x, y, label, value, width=24, font_size=27, line_height=38, max_lines=8):
    lines = _wrap_storyboard_lines(value, width=width, max_lines=max_lines)
    safe_label = html.escape(label)
    parts = [
        f'<text x="{x}" y="{y}" font-size="{font_size}" font-family="Microsoft YaHei, PingFang SC, Noto Sans CJK SC, Arial, sans-serif" fill="#111827">',
        f'<tspan font-weight="700">{safe_label}</tspan>',
    ]
    first = True
    for line in lines:
        safe_line = html.escape(line)
        if first:
            parts.append(f'<tspan dx="10">{safe_line}</tspan>')
            first = False
        else:
            parts.append(f'<tspan x="{x}" dy="{line_height}">{safe_line}</tspan>')
    parts.append('</text>')
    height = max(1, len(lines)) * line_height
    return "".join(parts), height


def _build_storyboard_svg(final_script_result, product_name=""):
    """Create a real downloadable SVG storyboard image; no extra AI/API call required."""
    storyboard = final_script_result.get("storyboard", []) if isinstance(final_script_result, dict) else []
    if not storyboard:
        return b""

    canvas_w = 1600
    margin = 58
    gap = 34
    card_w = (canvas_w - margin * 2 - gap) // 2
    header_h = 160

    cards = []
    for idx, item in enumerate(storyboard):
        fields = [
            ("机位/景别：", item.get("shot", ""), 24, 5),
            ("画面：", item.get("visual", ""), 24, 8),
            ("手部动作：", item.get("hand_action", ""), 24, 6),
            ("中文字幕/口播：", item.get("copy_cn", ""), 22, 6),
            ("剪辑/音效：", item.get("audio", ""), 24, 5),
            ("这一镜为什么拍：", item.get("rationale", ""), 24, 5),
        ]
        estimated = 112
        for _, value, width, max_lines in fields:
            estimated += min(max_lines, max(1, (len(clean_text(value) or "—") + width - 1) // width)) * 38 + 28
        cards.append((item, max(520, estimated)))

    # Row positions use the taller card in each pair.
    row_heights = []
    for i in range(0, len(cards), 2):
        pair = cards[i:i+2]
        row_heights.append(max(h for _, h in pair))
    canvas_h = header_h + margin + sum(row_heights) + gap * max(0, len(row_heights) - 1) + margin

    svg = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{canvas_w}" height="{canvas_h}" viewBox="0 0 {canvas_w} {canvas_h}">',
        '<rect width="100%" height="100%" fill="#f6f7f9"/>',
        f'<text x="{margin}" y="68" font-size="42" font-weight="800" font-family="Microsoft YaHei, PingFang SC, Noto Sans CJK SC, Arial, sans-serif" fill="#111827">中文分镜图｜拍摄 / 剪辑执行版</text>',
        f'<text x="{margin}" y="116" font-size="24" font-family="Microsoft YaHei, PingFang SC, Noto Sans CJK SC, Arial, sans-serif" fill="#6b7280">{html.escape(clean_text(product_name) or "TikTok 爆款拍摄方案")} · 每镜一卡 · 现场按顺序执行</text>',
    ]

    y = header_h
    card_index = 0
    for row_idx, row_h in enumerate(row_heights):
        for col_idx in range(2):
            if card_index >= len(cards):
                break
            item, card_h = cards[card_index]
            x = margin + col_idx * (card_w + gap)
            sequence = html.escape(clean_text(item.get("sequence")) or str(card_index + 1))
            time_range = html.escape(clean_text(item.get("time_range")) or "—")

            svg.append(f'<rect x="{x}" y="{y}" width="{card_w}" height="{card_h}" rx="22" fill="#ffffff" stroke="#e5e7eb" stroke-width="2"/>')
            svg.append(f'<rect x="{x}" y="{y}" width="{card_w}" height="76" rx="22" fill="#fff0f0"/>')
            # Mask lower rounded edge of header for a clean straight divider.
            svg.append(f'<rect x="{x}" y="{y+54}" width="{card_w}" height="22" fill="#fff0f0"/>')
            svg.append(f'<text x="{x+26}" y="{y+50}" font-size="31" font-weight="800" font-family="Microsoft YaHei, PingFang SC, Noto Sans CJK SC, Arial, sans-serif" fill="#111827">第{sequence}镜</text>')
            svg.append(f'<text x="{x+card_w-26}" y="{y+50}" text-anchor="end" font-size="24" font-weight="700" font-family="Microsoft YaHei, PingFang SC, Noto Sans CJK SC, Arial, sans-serif" fill="#ff4b4b">{time_range}</text>')

            cursor_y = y + 118
            sections = [
                ("机位/景别：", item.get("shot", ""), 24, 5),
                ("画面：", item.get("visual", ""), 24, 8),
                ("手部动作：", item.get("hand_action", ""), 24, 6),
                ("中文字幕/口播：", item.get("copy_cn", ""), 22, 6),
                ("剪辑/音效：", item.get("audio", ""), 24, 5),
                ("这一镜为什么拍：", item.get("rationale", ""), 24, 5),
            ]
            for label, value, width, max_lines in sections:
                block, block_h = _svg_text_block(x+26, cursor_y, label, value, width=width, max_lines=max_lines)
                svg.append(block)
                cursor_y += block_h + 28
            card_index += 1
        y += row_h + gap

    svg.append('</svg>')
    return "".join(svg).encode("utf-8")



def _build_storyboard_pdf(final_script_result, product_name=""):
    """Build a phone-friendly Chinese PDF, one storyboard card per page."""
    storyboard = final_script_result.get("storyboard", []) if isinstance(final_script_result, dict) else []
    if not storyboard:
        return b""

    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import mm
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    except Exception:
        return b""

    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    except Exception:
        pass

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        rightMargin=14 * mm,
        leftMargin=14 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
        title="TikTok中文分镜图",
        author="TikTok爆款视频解析&复盘专用",
    )

    styles = getSampleStyleSheet()
    base_font = "STSong-Light"
    title_style = ParagraphStyle(
        "StoryboardTitle",
        parent=styles["Title"],
        fontName=base_font,
        fontSize=19,
        leading=25,
        textColor=colors.HexColor("#111827"),
        alignment=TA_CENTER,
        spaceAfter=5 * mm,
    )
    sub_style = ParagraphStyle(
        "StoryboardSub",
        parent=styles["BodyText"],
        fontName=base_font,
        fontSize=10.5,
        leading=16,
        textColor=colors.HexColor("#6B7280"),
        alignment=TA_CENTER,
        spaceAfter=6 * mm,
    )
    head_style = ParagraphStyle(
        "ShotHead",
        parent=styles["Heading2"],
        fontName=base_font,
        fontSize=16,
        leading=22,
        textColor=colors.HexColor("#111827"),
        spaceAfter=4 * mm,
    )
    label_style = ParagraphStyle(
        "ShotLabel",
        parent=styles["BodyText"],
        fontName=base_font,
        fontSize=11,
        leading=18,
        textColor=colors.HexColor("#111827"),
    )
    value_style = ParagraphStyle(
        "ShotValue",
        parent=styles["BodyText"],
        fontName=base_font,
        fontSize=11,
        leading=18,
        textColor=colors.HexColor("#374151"),
    )
    copy_style = ParagraphStyle(
        "ShotCopy",
        parent=value_style,
        fontSize=12,
        leading=20,
        textColor=colors.HexColor("#111827"),
    )

    product_title = html.escape(clean_text(product_name) or "TikTok 爆款拍摄方案")
    story = []

    for idx, item in enumerate(storyboard):
        seq = html.escape(clean_text(item.get("sequence")) or str(idx + 1))
        time_range = html.escape(clean_text(item.get("time_range")) or "—")

        story.append(Paragraph("中文分镜图｜拍摄 / 剪辑执行版", title_style))
        story.append(Paragraph(f"{product_title} | 第 {seq} 镜 | {time_range}", sub_style))
        story.append(Paragraph(f"第 {seq} 镜　{time_range}", head_style))

        rows = [
            ("机位 / 景别", item.get("shot", ""), False),
            ("画面", item.get("visual", ""), False),
            ("手部动作", item.get("hand_action", ""), False),
            ("中文字幕 / 中文口播", item.get("copy_cn", ""), True),
            ("剪辑 / 音效", item.get("audio", ""), False),
            ("这一镜为什么拍", item.get("rationale", ""), False),
        ]

        table_data = []
        for label, value, is_copy in rows:
            value_text = html.escape(clean_text(value) or "—").replace("\n", "<br/>")
            table_data.append([
                Paragraph(f"<b>{html.escape(label)}</b>", label_style),
                Paragraph(value_text, copy_style if is_copy else value_style),
            ])

        table = Table(table_data, colWidths=[42 * mm, 125 * mm], hAlign="CENTER")
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F7F8FA")),
            ("BACKGROUND", (1, 3), (1, 3), colors.HexColor("#FFF3F3")),
            ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#E5E7EB")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#ECEFF3")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 9),
            ("RIGHTPADDING", (0, 0), (-1, -1), 9),
            ("TOPPADDING", (0, 0), (-1, -1), 9),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
        ]))
        story.append(table)
        story.append(Spacer(1, 4 * mm))
        story.append(Paragraph("拍摄时按页顺序执行；剪辑可直接按每页动作、字幕与节奏提示对照处理。", sub_style))
        if idx < len(storyboard) - 1:
            story.append(PageBreak())

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
if not api_key:
    st.error("系统未配置 Gemini API Key，请联系管理员。")

st.caption("爆款拆解｜爆款分析 → 选择卖点 → 3个方向 → 最终拍摄脚本 → 中文分镜图")

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
        analyze_videos,
        client, task_videos, category, product_name, input_selling_points,
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
            generate_directions,
            client, category, product_name, summary, chosen_ref_video,
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
                    generate_final_script,
                    client, category, product_name, chosen_ref_video, effective_selling_points,
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
        # ⑩ 中文分镜图（给拍摄 / 剪辑人员看）
        # ------------------------------------------------
        st.markdown("### ⑩ 中文分镜图（拍摄 / 剪辑执行版）")
        st.caption(
            "每镜一卡，只保留现场最需要看的信息：机位、画面、动作、中文字幕、剪辑提示。"
            "不是给机器识别，拍摄时可直接按卡片顺序执行。"
        )

        _render_storyboard_cards(final_script_result)

        storyboard_pdf = _build_storyboard_pdf(
            final_script_result,
            product_name=product_name,
        )
        storyboard_svg = _build_storyboard_svg(
            final_script_result,
            product_name=product_name,
        )

        pdf_col, svg_col = st.columns([1.35, 1])
        with pdf_col:
            if storyboard_pdf:
                st.download_button(
                    "下载中文分镜图（PDF｜手机推荐）",
                    data=storyboard_pdf,
                    file_name=(
                        "TikTok中文分镜图_"
                        + datetime.now().strftime("%Y%m%d_%H%M")
                        + ".pdf"
                    ),
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True,
                )
            else:
                st.caption("PDF组件尚未安装，SVG仍可正常使用。")

        with svg_col:
            if storyboard_svg:
                st.download_button(
                    "下载SVG大图（备用）",
                    data=storyboard_svg,
                    file_name=(
                        "TikTok中文分镜图_"
                        + datetime.now().strftime("%Y%m%d_%H%M")
                        + ".svg"
                    ),
                    mime="image/svg+xml",
                    use_container_width=True,
                )

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

