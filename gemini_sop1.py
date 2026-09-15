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
13. actual_duration_seconds：根据你实际看到的视频时间轴，给出整条视频的真实总时长（秒）。
14. subtitle_segments：给剪辑直接使用的完整英/西双语字幕时间轴，必须从视频开头覆盖到视频结尾。

【完整英/西双语字幕时间轴｜硬性要求】
- subtitle_segments 不是只写前20-40秒，也不是只摘精彩片段；必须按原视频实际完整时长输出。
- 例如视频实际约72秒，最后一条字幕的结束时间必须到约72秒附近；绝不能在30秒、40秒处提前结束。
- 先通过视频时间轴判断 actual_duration_seconds，再生成字幕段；不要把“新拍摄脚本15-40秒”的限制套到原视频拆解上。
- 每段必须包含 segment_no、time_range、copy_en、copy_es。
- time_range 使用清晰格式，例如 0.0-3.2s、3.2-7.0s；时间必须连续递增，不得倒序，不得明显超出视频真实时长。
- 按真实画面/语义变化拆段，通常每段约2-7秒；长句或连续同一语义可适当延长，但不要为了省字把几十秒塞成一段。
- copy_en：自然美国英语，适合TikTok屏幕字幕/口播，可直接给剪辑使用；不是逐字机械转录时，也必须忠实对应当前画面表达。
- copy_es：与 copy_en 同义、同一时间段，使用自然、简洁、适合美国西语受众阅读的拉美西班牙语；不要生硬直译。
- 如果原视频本身没有口播/字幕，也要根据该时间段的画面和脚本意图生成“可剪字幕稿”，但不得虚构产品不存在的功能。
- 输出JSON前必须自检：subtitle_segments 最后一段结束时间应与 actual_duration_seconds 基本一致（允许约1-2秒误差）。
- 对于1分钟以上的视频，必须继续生成到视频真正结束；不要因为输出较长而主动截断。

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

语言规则：shooting_notes、shot、visual、hand_action、copy_cn、audio、absorb_point、difference_point、rationale 全部必须中文。copy_en 使用自然美国英语；copy_es 使用自然、简洁、适合美国西语受众的拉美西班牙语。

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

【画面丰富度 + 简化机位硬性要求】
- 输出8-14个分镜，但整条视频最多只允许2个实际拍摄机位：1个主机位 + 1个补充特写机位。
- 用户选择的“实际视角”就是整条视频的主视角，不要擅自来回切换第一人称、第三人称、俯拍、侧拍、过肩、低机位等复杂视角。
- 默认优先只用1个主机位完成大部分素材；只有产品细节、Proof结果确实看不清时，才允许增加1个固定补充近景。
- 同一个机位连续拍多个动作素材，再通过剪辑切段；不要为了每个分镜都重新移动手机、重新架机位。
- 前3秒至少出现2-3次可感知变化，但优先通过手部动作、道具状态、产品状态、前后对比、推进/拿取/开合等完成，不依赖换机位。
- 平均每1.5-3秒至少变化一个“内容维度”：手部动作、道具状态、产品状态、前后对比、Proof、字幕信息；机位变化不是必选项。
- 同一个小场景内完成，不靠频繁换地点或频繁换角度制造丰富度。
- 禁止把脚本设计成需要频繁俯拍/侧拍/低角度/过肩/微距反复切换的专业拍摄方案。
- shot 字段只使用简洁、统一的机位名称，例如“主机位｜第一人称POV”“主机位｜第三人称手部”“补充特写｜产品结果”，不要写复杂摄影术语。
- 每个 visual 重点写清：道具 + 动作 + 画面变化/结果；除非切到补充特写，否则默认沿用主机位，不要重复描述复杂机位。
- 不要连续多个镜头只是“产品摆拍”或重复同一个动作。

【英/西双语字幕/口播硬性要求】
- 每个分镜必须同时输出 copy_en + copy_es，两者表达同一核心意思、适配同一时间段。
- copy_en 不是一句两三个词的占位字幕，而是给剪辑挑选的“可剪字幕池”。
- copy_es 必须是自然拉美西班牙语，面向美国西语受众；优先自然表达，不做逐词硬翻。
- 每个分镜通常给5-14个英文单词；关键镜头可给1-2个短句。
- 默认20-30秒视频，整条英文素材总量目标约65-100词；15-20秒约45-70词；30-40秒约90-130词。
- 西班牙语字幕与英文字幕逐镜对应，信息量保持相近，但允许为阅读速度做自然压缩。
- 宁可比最终成片多20%-30%的双语素材，方便剪辑删减，也不要少到无法选择。
- 必须自然美国口语、TikTok风格、短句、可直接上字幕；不要广告说明书语气。
- 不要每个镜头重复同一句卖点，要形成信息递进：Hook → 痛点 → 解决 → Proof → 使用场景/理由 → CTA。

0-3秒必须强Hook且有明显动作；产品尽快出现；必须真实UGC；必须可以在民宿直接拍；不需要专业摄影设备。整条视频应尽量“一次架机，多拍动作”，减少重新摆机位和重复布光；每个动作必须具体，不要写抽象脚本。

每个分镜输出：
sequence、time_range、shot、visual、hand_action、copy_cn、copy_en、copy_es、audio、absorb_point、difference_point、rationale。

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
                "你必须严格遵守语言规则：本次爆款拆解的分析字段使用简体中文；完整字幕时间轴必须同时输出自然美国英语和拉美西班牙语。"
                "必须同时指出每条视频的优点和短板，不能只夸优点。"
                "原视频字幕时间轴必须覆盖视频真实完整时长，1分钟以上的视频不得在30-40秒提前截断。"
            ),
            thinking_config=_thinking(),
            max_output_tokens=7800,
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
            "这是给中国团队执行的拍摄脚本。分析与执行字段必须使用简体中文。"
            "copy_en必须提供足量、可供剪辑删选的自然美国英语；copy_es必须提供对应的自然拉美西班牙语。"
            "画面必须丰富，输出8-14个分镜。"
        ),
        thinking_config=_thinking(),
        max_output_tokens=7600,
        response_mime_type="application/json",
        response_json_schema=FINAL_SCRIPT_SCHEMA,
    )
    response, meta = generate_resilient(client, prompt, config)
    result = parse_json_output(response.text)
    return result, {"analysis_seconds": round(time.perf_counter() - started, 1), **meta}
