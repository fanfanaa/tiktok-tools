import json
import os
import tempfile
import time
from pathlib import Path

from google.genai import types

from config import SOP1_MODEL_CHAIN, MAX_COMPARE_VIDEOS, INLINE_BATCH_MAX_MB, SCENE_LIBRARY
from common import clean_text, parse_json_output
from gemini_base import generate_resilient as _base_generate_resilient, wait_until_active
from sop1_schema import VIDEO_ANALYSIS_SCHEMA, DIRECTIONS_SCHEMA, FINAL_SCRIPT_SCHEMA


def _thinking():
    return types.ThinkingConfig(thinking_level="minimal")


def generate_resilient(client, contents, config):
    """SOP1 使用自己的稳定模型链：3.5 -> 3.1 -> 3.6。"""
    return _base_generate_resilient(
        client,
        contents,
        config,
        model_chain=SOP1_MODEL_CHAIN,
    )


def _scene_description(scene_name):
    value = SCENE_LIBRARY.get(scene_name, "")
    if isinstance(value, dict):
        return value.get("prompt") or value.get("guide") or json.dumps(value, ensure_ascii=False)
    return str(value)


def build_video_analysis_prompt(category, product_name, filenames, input_selling_points):
    file_lines = "\n".join(
        f"视频{i + 1}：{filename}" for i, filename in enumerate(filenames)
    )

    return f"""
你是美国 TikTok Shop 爆款短视频分析负责人。

语言规则：除文件名、必要英文原句外，所有分析结果必须使用简体中文。

本次上传：
{file_lines}

产品品类：{category}
产品名称：{product_name}
用户填写的真实产品卖点：{input_selling_points if clean_text(input_selling_points) else "未填写"}

请逐条单独分析视频。不同视频可能主打完全不同的卖点，不能把不同视频的结论混在一起。

每一条视频必须独立输出：
1. 一句话核心。
2. inferred_selling_points：3-5条真正主推卖点，必须中文。
3. 爆款脚本路线。
4. 人群画像。
5. 年龄预估，并明确这是“预估”。
6. 前3秒Hook。
7. 画面与节奏。
8. strengths：2-5条明确优点。
9. weaknesses：2-5条明确短板。
10. 最值得吸收的3点。
11. 参考价值判断。
12. 推荐指数0-100。

非常重要：
- 爆款不等于完美，必须找短板，不能只写优势。
- strengths / weaknesses 必须基于你实际看到的画面、动作、信息密度、字幕、节奏、产品出现、Proof、CTA 等内容判断。
- 不得用“播放量低、转化差”等没有后台数据支撑的内容当短板。
- 短板重点帮助团队判断：前1秒是否弱、前3秒是否拖、画面是否重复、动作是否单一、产品出现是否太晚、Proof是否不足、字幕是否过少/过密、机位是否单调、CTA是否生硬、信息层级是否不清。

如果用户填写了真实卖点，请比较该视频推理卖点 VS 用户真实卖点，并输出：
- selling_point_relation：只能 similar / different
- selling_point_relation_reason：中文
- blended_selling_points：中文
- suggested_mode：只能 viral_first / user_first / blend
suggested_mode只是推荐，最终决定权属于使用人。

如果用户没有填写卖点：
- selling_point_relation = no_user_input
- selling_point_relation_reason：中文说明
- blended_selling_points：用中文整理该视频推理卖点
- suggested_mode = viral_first

所有视频还要输出 comparison_summary，包括：
- 一句话共同核心
- 共同爆款脚本路线
- 共同人群
- 年龄预估
- 共同Hook
- 共同画面节奏
- common_strengths：这些视频共同做得好的2-5点
- common_weaknesses：这些视频共同存在的2-5个短板/风险
- 最值得共同吸收3点
- 多视频关键差异：不要只写相同点，要明确每条视频谁更强、谁更弱、各自牺牲了什么

最后推荐 recommended_reference_video_index，但只是AI推荐，最终由使用人选择。

禁止虚构产品功能、TikTok后台数据、销量、认证、医疗效果；禁止复制原视频完整台词。
严格按JSON Schema输出。
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
    scene_names = "\n".join(f"- {scene}" for scene in SCENE_LIBRARY.keys())
    return f"""
你是美国 TikTok Shop 中文短视频拍摄策划负责人。
除必要英文产品名之外，本次输出全部使用简体中文。

产品：{category} / {product_name}
用户真实卖点：{input_selling_points if clean_text(input_selling_points) else "未填写"}
最终由使用人确认的有效卖点：{effective_selling_points}
用户选择的卖点策略：{selling_point_mode}

主参考视频：
{json.dumps(selected_video, ensure_ascii=False, indent=2)}

其他视频共同规律：
{json.dumps(comparison_summary, ensure_ascii=False, indent=2)}

硬性拍摄限制：真人不露脸；禁止正脸；禁止主播面对镜头讲话。
允许手、手臂、少量身体局部、少量背影。
拍摄视角只允许第一人称POV、第三人称手部/局部视角；默认优先第一人称。

可使用场景：
{scene_names}
recommended_scene必须从上面场景名称选择。

生成3个明显不同的创意方向，不能只是换一句文案。
至少在3个维度形成差异：前3秒Hook、痛点、手部动作、Demo顺序、产品切入、场景、情绪、CTA。

每个方向输出：direction_name、core_idea、target_audience、hook、product_entry、recommended_perspective、recommended_scene、absorb_points、differentiation_points。
全部中文。本阶段不要输出逐秒脚本。
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
    scene_desc = _scene_description(selected_scene)
    return f"""
你是美国 TikTok Shop 中文拍摄SOP导演。这是给中国拍摄团队直接执行的脚本。

语言规则：除 copy_en 字段外，shooting_notes、shot、visual、hand_action、copy_cn、audio、absorb_point、difference_point、rationale 全部必须中文。只有 copy_en 使用自然美国英语。

产品：{category} / {product_name}
最终确认卖点：{effective_selling_points}

主参考视频：
{json.dumps(selected_video, ensure_ascii=False, indent=2)}

最终选择方向：
{json.dumps(chosen_direction, ensure_ascii=False, indent=2)}

实际场景：{selected_scene}
场景说明：{scene_desc}
实际视角：{selected_perspective}

真人限制：不露脸、不允许正脸、不允许主播对镜头讲话；允许手、手臂、身体局部、少量背影。

视频时长15-40秒。

【画面丰富度硬性要求】
- 输出8-14个分镜，不允许只有少量大段镜头。
- 前3秒至少出现2-3次可感知的画面/动作变化。
- 平均每1.5-3秒至少变化一个维度：机位、景别、手部动作、道具状态、产品状态、前后对比、Proof展示。
- 同一个小场景内完成，不靠频繁换地点制造丰富度。
- 必须混合：POV操作、手部近景、产品微距、结果特写、前后对比/Proof、拿取/放回/开合/滚动/撕/推/按等动作。
- 不要连续多个镜头只是“产品摆拍”或重复同一个动作。
- 每个 visual 必须写清：机位 + 道具 + 动作 + 画面变化/结果，让剪辑能直接按描述找镜头。

【英文字幕/口播硬性要求】
- copy_en 不是一句两三个词的占位字幕，而是给剪辑挑选的“可剪字幕池”。
- 每个分镜通常给5-14个英文单词；关键镜头可给1-2个短句。
- 默认20-30秒视频，整条英文素材总量目标约65-100词；15-20秒约45-70词；30-40秒约90-130词。
- 宁可比最终成片多20%-30%的英文素材，方便剪辑删减，也不要少到无法选择。
- 必须自然美国口语、TikTok风格、短句、可直接上字幕；不要广告说明书语气。
- 不要每个镜头重复同一句卖点，要形成信息递进：Hook → 痛点 → 解决 → Proof → 使用场景/理由 → CTA。

0-3秒必须强Hook且有明显动作；产品尽快出现；必须真实UGC；必须可以在民宿直接拍；不需要专业摄影设备；每个动作必须具体，不要写抽象脚本。

每个分镜输出：
sequence、time_range、shot、visual、hand_action、copy_cn、copy_en、audio、absorb_point、difference_point、rationale。

严格按JSON Schema。
""".strip()


def analyze_videos(client, uploaded_videos, category, product_name, input_selling_points):
    if not uploaded_videos:
        raise ValueError("请至少上传1条视频。")
    if len(uploaded_videos) > MAX_COMPARE_VIDEOS:
        raise ValueError(f"单次最多上传{MAX_COMPARE_VIDEOS}条视频。")

    started = time.perf_counter()
    total_bytes = sum(len(video.getvalue()) for video in uploaded_videos)
    total_mb = total_bytes / 1024 / 1024
    filenames = [video.name for video in uploaded_videos]
    prompt = build_video_analysis_prompt(category, product_name, filenames, input_selling_points)

    remote_files = []
    temp_paths = []
    try:
        parts = []
        if total_mb <= INLINE_BATCH_MAX_MB:
            analysis_mode = "多视频快速解析"
            for index, video in enumerate(uploaded_videos, start=1):
                parts.append(types.Part.from_text(text=f"【视频{index}】文件名：{video.name}"))
                parts.append(types.Part.from_bytes(data=video.getvalue(), mime_type=video.type or "video/mp4"))
        else:
            analysis_mode = "多视频大文件解析"
            for index, video in enumerate(uploaded_videos, start=1):
                suffix = Path(video.name).suffix or ".mp4"
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp:
                    temp.write(video.getvalue())
                    temp_path = temp.name
                temp_paths.append(temp_path)
                remote_file = client.files.upload(file=temp_path)
                remote_file = wait_until_active(client, remote_file)
                remote_files.append(remote_file)
                parts.append(types.Part.from_text(text=f"【视频{index}】文件名：{video.name}"))
                parts.append(types.Part.from_uri(file_uri=remote_file.uri, mime_type=remote_file.mime_type or "video/mp4"))

        parts.append(types.Part.from_text(text=prompt))
        content = types.Content(role="user", parts=parts)
        config = types.GenerateContentConfig(
            system_instruction=(
                "你必须严格遵守语言规则：本次爆款拆解所有分析字段必须使用简体中文。"
                "必须同时指出每条视频的优点和短板，不能只夸优点。"
            ),
            thinking_config=_thinking(),
            max_output_tokens=7600,
            response_mime_type="application/json",
            response_json_schema=VIDEO_ANALYSIS_SCHEMA,
        )
        response, meta = generate_resilient(client, content, config)
        result = parse_json_output(response.text)
        return result, {
            "analysis_mode": analysis_mode,
            "total_size_mb": round(total_mb, 2),
            "video_count": len(uploaded_videos),
            "analysis_seconds": round(time.perf_counter() - started, 1),
            **meta,
        }
    finally:
        for remote_file in remote_files:
            try:
                client.files.delete(name=remote_file.name)
            except Exception:
                pass
        for temp_path in temp_paths:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
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
    started = time.perf_counter()
    prompt = build_directions_prompt(
        category,
        product_name,
        comparison_summary,
        selected_video,
        input_selling_points,
        effective_selling_points,
        selling_point_mode,
    )
    config = types.GenerateContentConfig(
        system_instruction="除必要英文产品名之外，三个创意方向全部必须使用简体中文。",
        thinking_config=_thinking(),
        max_output_tokens=3200,
        response_mime_type="application/json",
        response_json_schema=DIRECTIONS_SCHEMA,
    )
    response, meta = generate_resilient(client, prompt, config)
    result = parse_json_output(response.text)
    return result, {"analysis_seconds": round(time.perf_counter() - started, 1), **meta}


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
    started = time.perf_counter()
    prompt = build_final_script_prompt(
        category,
        product_name,
        selected_video,
        effective_selling_points,
        chosen_direction,
        selected_scene,
        selected_perspective,
    )
    config = types.GenerateContentConfig(
        system_instruction=(
            "这是给中国团队执行的拍摄脚本。除copy_en字段之外所有字段必须使用简体中文。"
            "copy_en必须提供足量、可供剪辑删选的自然美国英语；画面必须丰富，输出8-14个分镜。"
        ),
        thinking_config=_thinking(),
        max_output_tokens=6800,
        response_mime_type="application/json",
        response_json_schema=FINAL_SCRIPT_SCHEMA,
    )
    response, meta = generate_resilient(client, prompt, config)
    result = parse_json_output(response.text)
    return result, {"analysis_seconds": round(time.perf_counter() - started, 1), **meta}
