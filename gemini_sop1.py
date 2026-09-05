import json
import os
import tempfile
import time
from pathlib import Path

from google.genai import types

from config import (SOP1_MODEL_CHAIN, MAX_COMPARE_VIDEOS, INLINE_BATCH_MAX_MB,
                    FILE_PROCESS_TIMEOUT_SEC, FILE_POLL_INTERVAL_SEC, SCENE_LIBRARY)
from sop1_schema import VIDEO_ANALYSIS_SCHEMA, DIRECTIONS_SCHEMA, FINAL_SCRIPT_SCHEMA
from common import clean_text, parse_json_output
from gemini_base import generate_resilient as _base_generate_resilient, wait_until_active


def thinking_config():
    return types.ThinkingConfig(thinking_level="minimal")

def generate_resilient(client, contents, config):
    return _base_generate_resilient(client, contents, config, model_chain=SOP1_MODEL_CHAIN)

def build_video_analysis_prompt(
    category,
    product_name,
    filenames,
    input_selling_points,
):

    file_lines = "\n".join(
        [
            f"视频{i + 1}：{filename}"
            for i, filename
            in enumerate(
                filenames
            )
        ]
    )

    return f"""
你是美国 TikTok Shop 爆款短视频分析负责人。

非常重要：

除文件名、必要英文原句之外，
所有分析结果必须使用简体中文。

禁止用英文回答：
- 卖点
- 脚本路线
- 人群
- 年龄
- Hook分析
- 节奏
- 推荐原因
- 差异分析
- 参考价值

如果这些字段输出英文，视为任务失败。

==================================================

本次上传：

{file_lines}

产品品类：

{category}

产品名称：

{product_name}

用户填写的真实产品卖点：

{input_selling_points if clean_text(input_selling_points) else "未填写"}

==================================================

请逐条单独分析视频。

不同视频可能主打完全不同的卖点。

所以每一条视频必须独立输出：

1. 一句话核心

2. inferred_selling_points

单独总结该视频真正主推的3-5条核心卖点。

必须中文。

不能混入其他视频卖点。

3. 爆款脚本路线

4. 人群画像

5. 年龄预估

6. 前3秒Hook

7. 画面与节奏

8. 最值得吸收的3点

9. 参考价值判断

10. 推荐指数 0-100

==================================================

如果用户填写了真实卖点：

请比较：

该视频推理卖点

VS

用户真实卖点

输出：

selling_point_relation

只能是：

similar

different

selling_point_relation_reason：

必须中文。

blended_selling_points：

如果融合，输出中文卖点表达。

suggested_mode：

只能是：

viral_first

user_first

blend

注意：

suggested_mode只是你的推荐。

最终决定权属于使用人。

你不能替使用人做最终选择。

==================================================

如果用户没有填写卖点：

selling_point_relation：

no_user_input

selling_point_relation_reason：

中文说明。

blended_selling_points：

用中文整理该视频推理卖点。

suggested_mode：

viral_first

==================================================

所有视频还要输出：

comparison_summary

包括：

- 一句话共同核心
- 共同爆款脚本路线
- 共同人群
- 年龄预估
- 共同Hook
- 共同画面节奏
- 最值得共同吸收3点
- 多视频关键差异

全部中文。

最后：

recommended_reference_video_index

推荐一条最值得作为主参考的视频。

但这只是AI推荐。

最终由使用人选择。

==================================================

禁止：

- 虚构产品功能
- 虚构TikTok后台数据
- 虚构销量
- 虚构认证
- 虚构医疗效果
- 复制品牌
- 复制原视频完整台词

严格按JSON Schema。
""".strip()

def build_directions_prompt(
    category,
    product_name,
    comparison_summary,
    selected_video,
    input_selling_points,
    effective_selling_points,
    selling_point_mode,
):

    scene_names = "\n".join(
        [
            f"- {scene}"
            for scene
            in SCENE_LIBRARY.keys()
        ]
    )

    return f"""
你是美国 TikTok Shop 中文短视频拍摄策划负责人。

除必要英文产品名之外，
本次输出全部使用简体中文。

禁止用英文输出：
- 方向名
- 核心思路
- 目标人群
- Hook分析
- 产品切入方式
- 推荐视角
- 推荐场景
- 可吸收点
- 差异化点

==================================================

产品：

{category}

{product_name}

用户真实卖点：

{input_selling_points if clean_text(input_selling_points) else "未填写"}

最终由使用人确认的有效卖点：

{effective_selling_points}

用户选择的卖点策略：

{selling_point_mode}

==================================================

主参考视频：

{json.dumps(
    selected_video,
    ensure_ascii=False,
    indent=2
)}

其他视频共同规律：

{json.dumps(
    comparison_summary,
    ensure_ascii=False,
    indent=2
)}

==================================================

硬性拍摄限制：

真人不露脸。

禁止正脸。

禁止主播面对镜头讲话。

允许：

- 手
- 手臂
- 少量身体局部
- 少量背影

拍摄视角只允许：

第一人称 POV

第三人称手部/局部视角

默认优先第一人称。

==================================================

可使用场景：

{scene_names}

recommended_scene必须从上面场景名称选择。

==================================================

生成3个明显不同的创意方向。

不能只是换一句文案。

至少在3个维度形成差异：

- 前3秒Hook
- 痛点
- 手部动作
- Demo顺序
- 产品切入
- 场景
- 情绪
- CTA

每个方向输出：

direction_name
core_idea
target_audience
hook
product_entry
recommended_perspective
recommended_scene
absorb_points
differentiation_points

全部中文。

本阶段不要输出逐秒脚本。

严格按JSON Schema。
""".strip()

def build_final_script_prompt(
    category,
    product_name,
    selected_video,
    effective_selling_points,
    chosen_direction,
    selected_scene,
    selected_perspective,
):

    return f"""
你是美国 TikTok Shop 中文拍摄SOP导演。

这是给中国拍摄团队直接执行的脚本。

非常重要：

除 copy_en 字段之外，
所有字段必须使用简体中文。

包括：

shooting_notes
shot
visual
hand_action
copy_cn
audio
absorb_point
difference_point
rationale

全部必须中文。

只有：

copy_en

使用自然美国英语。

==================================================

产品：

{category}

{product_name}

最终确认卖点：

{effective_selling_points}

主参考视频：

{json.dumps(
    selected_video,
    ensure_ascii=False,
    indent=2
)}

最终选择方向：

{json.dumps(
    chosen_direction,
    ensure_ascii=False,
    indent=2
)}

实际场景：

{selected_scene}

场景说明：

{SCENE_LIBRARY[selected_scene]}

实际视角：

{selected_perspective}

==================================================

真人限制：

不露脸。

不允许正脸。

不允许主播对镜头讲话。

允许：

- 手
- 手臂
- 身体局部
- 少量背影

==================================================

视频时长：

15-40秒。

要求：

- 0-3秒必须强Hook
- 前3秒必须有明显动作
- 产品尽快出现
- 必须真实UGC
- 必须可以在民宿直接拍
- 不需要专业摄影设备
- 每个动作必须具体
- 不要写抽象脚本

例如：

错误：

“展示印章效果”

正确：

“手机固定俯拍，左手压住快递标签右侧，
右手从画面右边拿起印章，
从姓名和地址区域由左向右滚一次，
镜头不断，让观众完整看到文字被遮盖。”

==================================================

每个分镜输出：

sequence

time_range

shot
中文

visual
中文

hand_action
中文

copy_cn
中文口播/字幕参考

copy_en
真正用于美国TikTok的英文字幕/口播

audio
中文说明

absorb_point
中文

difference_point
中文

rationale
中文

==================================================

英文 copy_en 要求：

- 美国口语
- 短句
- TikTok风格
- 不像广告说明书

严格按JSON Schema。
""".strip()

def analyze_videos(
    client,
    uploaded_videos,
    category,
    product_name,
    input_selling_points,
):

    if not uploaded_videos:

        raise ValueError(
            "请至少上传1条视频。"
        )

    if (
        len(uploaded_videos)
        > MAX_COMPARE_VIDEOS
    ):

        raise ValueError(
            f"单次最多上传{MAX_COMPARE_VIDEOS}条视频。"
        )

    started = (
        time.perf_counter()
    )

    total_bytes = sum(
        len(
            video.getvalue()
        )
        for video
        in uploaded_videos
    )

    total_mb = (
        total_bytes
        / 1024
        / 1024
    )

    filenames = [
        video.name
        for video
        in uploaded_videos
    ]

    prompt = (
        build_video_analysis_prompt(
            category,
            product_name,
            filenames,
            input_selling_points,
        )
    )

    remote_files = []
    temp_paths = []

    try:

        parts = []

        if (
            total_mb
            <= INLINE_BATCH_MAX_MB
        ):

            analysis_mode = (
                "多视频快速解析"
            )

            for index, video in enumerate(
                uploaded_videos,
                start=1,
            ):

                parts.append(
                    types.Part.from_text(
                        text=(
                            f"【视频{index}】"
                            f"文件名：{video.name}"
                        )
                    )
                )

                parts.append(
                    types.Part.from_bytes(
                        data=video.getvalue(),
                        mime_type=(
                            video.type
                            or "video/mp4"
                        ),
                    )
                )

        else:

            analysis_mode = (
                "多视频大文件解析"
            )

            for index, video in enumerate(
                uploaded_videos,
                start=1,
            ):

                suffix = (
                    Path(
                        video.name
                    ).suffix
                    or ".mp4"
                )

                with tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=suffix,
                ) as temp:

                    temp.write(
                        video.getvalue()
                    )

                    temp_path = (
                        temp.name
                    )

                temp_paths.append(
                    temp_path
                )

                remote_file = (
                    client.files.upload(
                        file=temp_path
                    )
                )

                remote_file = (
                    wait_until_active(
                        client,
                        remote_file,
                    )
                )

                remote_files.append(
                    remote_file
                )

                parts.append(
                    types.Part.from_text(
                        text=(
                            f"【视频{index}】"
                            f"文件名：{video.name}"
                        )
                    )
                )

                parts.append(
                    types.Part.from_uri(
                        file_uri=remote_file.uri,
                        mime_type=(
                            remote_file.mime_type
                            or "video/mp4"
                        ),
                    )
                )

        parts.append(
            types.Part.from_text(
                text=prompt
            )
        )

        content = types.Content(
            role="user",
            parts=parts,
        )

        config = types.GenerateContentConfig(

            system_instruction=(
                "你必须严格遵守语言规则："
                "本次爆款拆解所有分析字段必须使用简体中文。"
                "不得因为视频是美国TikTok内容就改用英文分析。"
            ),

            thinking_config=(
                thinking_config()
            ),

            max_output_tokens=6200,

            response_mime_type="application/json",

            response_json_schema=(
                VIDEO_ANALYSIS_SCHEMA
            ),
        )

        response, meta = (
            generate_resilient(
                client,
                content,
                config,
            )
        )

        result = (
            parse_json_output(
                response.text
            )
        )

        return (
            result,
            {

                "analysis_mode":
                    analysis_mode,

                "total_size_mb":
                    round(
                        total_mb,
                        2,
                    ),

                "video_count":
                    len(
                        uploaded_videos
                    ),

                "analysis_seconds":
                    round(
                        time.perf_counter()
                        - started,
                        1,
                    ),

                **meta,
            },
        )

    finally:

        for remote_file in remote_files:

            try:

                client.files.delete(
                    name=remote_file.name
                )

            except Exception:
                pass

        for temp_path in temp_paths:

            try:

                if os.path.exists(
                    temp_path
                ):

                    os.remove(
                        temp_path
                    )

            except OSError:
                pass

def generate_directions(
    client,
    category,
    product_name,
    comparison_summary,
    selected_video,
    input_selling_points,
    effective_selling_points,
    selling_point_mode,
):

    started = (
        time.perf_counter()
    )

    prompt = (
        build_directions_prompt(
            category,
            product_name,
            comparison_summary,
            selected_video,
            input_selling_points,
            effective_selling_points,
            selling_point_mode,
        )
    )

    config = types.GenerateContentConfig(

        system_instruction=(
            "除必要英文产品名之外，"
            "三个创意方向全部必须使用简体中文。"
        ),

        thinking_config=(
            thinking_config()
        ),

        max_output_tokens=3200,

        response_mime_type="application/json",

        response_json_schema=(
            DIRECTIONS_SCHEMA
        ),
    )

    response, meta = (
        generate_resilient(
            client,
            prompt,
            config,
        )
    )

    result = (
        parse_json_output(
            response.text
        )
    )

    return (
        result,
        {

            "analysis_seconds":
                round(
                    time.perf_counter()
                    - started,
                    1,
                ),

            **meta,
        },
    )

def generate_final_script(
    client,
    category,
    product_name,
    selected_video,
    effective_selling_points,
    chosen_direction,
    selected_scene,
    selected_perspective,
):

    started = (
        time.perf_counter()
    )

    prompt = (
        build_final_script_prompt(
            category,
            product_name,
            selected_video,
            effective_selling_points,
            chosen_direction,
            selected_scene,
            selected_perspective,
        )
    )

    config = types.GenerateContentConfig(

        system_instruction=(
            "这是给中国团队执行的拍摄脚本。"
            "除copy_en字段之外，"
            "所有字段必须使用简体中文。"
            "copy_en才使用自然美国英语。"
        ),

        thinking_config=(
            thinking_config()
        ),

        max_output_tokens=4800,

        response_mime_type="application/json",

        response_json_schema=(
            FINAL_SCRIPT_SCHEMA
        ),
    )

    response, meta = (
        generate_resilient(
            client,
            prompt,
            config,
        )
    )

    result = (
        parse_json_output(
            response.text
        )
    )

    return (
        result,
        {

            "analysis_seconds":
                round(
                    time.perf_counter()
                    - started,
                    1,
                ),

            **meta,
        },
    )

