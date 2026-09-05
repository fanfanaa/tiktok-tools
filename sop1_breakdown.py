from datetime import datetime

import streamlit as st

from config import PRODUCT_CATEGORIES, PERSPECTIVE_OPTIONS, SCENE_LIBRARY, MAX_COMPARE_VIDEOS
from common import clean_text, get_api_key, json_dumps, list_to_joined, make_signature, safe_int, video_batch_signature
from gemini_base import create_client, friendly_error
from gemini_sop1 import analyze_videos, generate_directions, generate_final_script
from history_service import append_history
from export_service import final_script_to_df, build_analysis_export_excel


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
api_key = get_api_key()
client = create_client(api_key) if api_key else None
if not api_key:
    st.error("系统未配置 Gemini API Key，请联系管理员。")

st.caption("SOP1｜爆款拆解 → 选择卖点 → 3个方向 → 最终拍摄脚本")

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

    category = (
        st.selectbox(
            "产品品类",
            PRODUCT_CATEGORIES,
            key="analysis_category",
        )
    )

with c3:

    product_name = (
        st.text_input(
            "产品名称 / SKU",
            key="analysis_product_name",
            placeholder="例如：隐私印章",
        )
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

uploaded_videos = (
    st.file_uploader(
        "支持同时上传 1-5 条 .mp4 视频",
        type=["mp4"],
        accept_multiple_files=True,
        key="analysis_videos",
    )
)

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

if analyze_button:

    try:

        with st.spinner(
            "正在逐条拆解视频并提取独立卖点…"
        ):

            (
                result,
                metadata,
            ) = analyze_videos(
                client,
                uploaded_videos,
                category,
                product_name,
                input_selling_points,
            )

        st.session_state[
            "video_analysis_result"
        ] = result

        st.session_state[
            "video_analysis_meta"
        ] = metadata

        recommended = safe_int(
            result.get(
                "recommended_reference_video_index",
                1,
            ),
            1,
        )

        st.session_state[
            "selected_reference_video_index"
        ] = recommended

        append_history(
            {

                "record_type":
                    "爆款对比解析",

                "role":
                    st.session_state[
                        "role"
                    ],

                "operator":
                    st.session_state[
                        "operator"
                    ],

                "tiktok_account":
                    tiktok_account,

                "product_category":
                    category,

                "product_name":
                    product_name,

                "input_selling_points":
                    input_selling_points,

                "video_names":
                    " | ".join(
                        [
                            video.name
                            for video
                            in uploaded_videos
                        ]
                    ),

                "video_count":
                    len(
                        uploaded_videos
                    ),

                "model_used":
                    metadata.get(
                        "model_used",
                        "",
                    ),

                "fallback_used":
                    metadata.get(
                        "fallback_used",
                        "",
                    ),

                "retry_count":
                    metadata.get(
                        "retry_count",
                        "",
                    ),

                "analysis_seconds":
                    metadata.get(
                        "analysis_seconds",
                        "",
                    ),

                "full_output_json":
                    json_dumps(
                        result
                    ),
            }
        )

        st.success(
            "爆款视频解析完成。"
        )

    except Exception as exc:

        st.error(
            friendly_error(
                exc
            )
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

    if generate_direction_button:

        try:

            with st.spinner(
                "正在生成3个不同拍摄方向…"
            ):

                (
                    directions_result,
                    directions_meta,
                ) = generate_directions(
                    client,
                    category,
                    product_name,
                    summary,
                    chosen_ref_video,
                    input_selling_points,
                    effective_selling_points,
                    selling_point_mode,
                )

            st.session_state[
                "directions_result"
            ] = directions_result

            st.session_state[
                "directions_meta"
            ] = directions_meta

            st.session_state[
                "directions_context_signature"
            ] = current_direction_context

            st.session_state[
                "selected_direction_index"
            ] = 0

            st.session_state[
                "final_script_result"
            ] = None

            append_history(
                {

                    "record_type":
                        "3方向生成",

                    "role":
                        st.session_state[
                            "role"
                        ],

                    "operator":
                        st.session_state[
                            "operator"
                        ],

                    "tiktok_account":
                        tiktok_account,

                    "product_category":
                        category,

                    "product_name":
                        product_name,

                    "input_selling_points":
                        input_selling_points,

                    "inferred_selling_points":
                        viral_points_text,

                    "effective_selling_points":
                        effective_selling_points,

                    "selling_point_mode":
                        selling_point_mode,

                    "reference_video_index":
                        selected_ref_index,

                    "reference_video_name":
                        chosen_ref_video.get(
                            "filename",
                            "",
                        ),

                    "model_used":
                        directions_meta.get(
                            "model_used",
                            "",
                        ),

                    "fallback_used":
                        directions_meta.get(
                            "fallback_used",
                            "",
                        ),

                    "retry_count":
                        directions_meta.get(
                            "retry_count",
                            "",
                        ),

                    "analysis_seconds":
                        directions_meta.get(
                            "analysis_seconds",
                            "",
                        ),

                    "full_output_json":
                        json_dumps(
                            directions_result
                        ),
                }
            )

            st.success(
                "3个拍摄方向已生成。"
            )

        except Exception as exc:

            st.error(
                friendly_error(
                    exc
                )
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

            if generate_script_button:

                try:

                    with st.spinner(
                        "正在生成中文可执行拍摄脚本…"
                    ):

                        (
                            final_result,
                            final_meta,
                        ) = generate_final_script(
                            client,
                            category,
                            product_name,
                            chosen_ref_video,
                            effective_selling_points,
                            chosen_direction,
                            selected_scene,
                            selected_perspective,
                        )

                    st.session_state[
                        "final_script_result"
                    ] = final_result

                    st.session_state[
                        "final_script_meta"
                    ] = final_meta

                    st.session_state[
                        "final_script_context_signature"
                    ] = current_final_context

                    append_history(
                        {

                            "record_type":
                                "最终拍摄脚本",

                            "role":
                                st.session_state[
                                    "role"
                                ],

                            "operator":
                                st.session_state[
                                    "operator"
                                ],

                            "tiktok_account":
                                tiktok_account,

                            "product_category":
                                category,

                            "product_name":
                                product_name,

                            "input_selling_points":
                                input_selling_points,

                            "inferred_selling_points":
                                viral_points_text,

                            "effective_selling_points":
                                effective_selling_points,

                            "selling_point_mode":
                                selling_point_mode,

                            "reference_video_index":
                                selected_ref_index,

                            "reference_video_name":
                                chosen_ref_video.get(
                                    "filename",
                                    "",
                                ),

                            "direction_name":
                                chosen_direction.get(
                                    "direction_name",
                                    "",
                                ),

                            "selected_scene":
                                selected_scene,

                            "selected_perspective":
                                selected_perspective,

                            "model_used":
                                final_meta.get(
                                    "model_used",
                                    "",
                                ),

                            "fallback_used":
                                final_meta.get(
                                    "fallback_used",
                                    "",
                                ),

                            "retry_count":
                                final_meta.get(
                                    "retry_count",
                                    "",
                                ),

                            "analysis_seconds":
                                final_meta.get(
                                    "analysis_seconds",
                                    "",
                                ),

                            "full_output_json":
                                json_dumps(
                                    final_result
                                ),
                        }
                    )

                    st.success(
                        "最终拍摄脚本已生成。"
                    )

                except Exception as exc:

                    st.error(
                        friendly_error(
                            exc
                        )
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

