import os
import tempfile
import time
from pathlib import Path

from google.genai import types

from config import (
    SOP2_MODEL_CHAIN, SOP2_MAX_VIRAL_VIDEOS, SOP2_MAX_OWN_VIDEOS,
    SOP2_MAX_TOTAL_VIDEOS, INLINE_BATCH_MAX_MB
)
from sop2_schema import SOP2_PRE_ANALYSIS_SCHEMA, SOP2_DEEP_COMPARE_SCHEMA
from common import clean_text, parse_json_output
from gemini_base import generate_resilient as _base_generate_resilient, wait_until_active


def _is_transient_error(exc):
    text = f"{type(exc).__name__}: {exc}".upper()
    markers = (
        "429", "500", "502", "503", "504",
        "RESOURCE_EXHAUSTED", "TOO_MANY_REQUESTS", "RATE_LIMIT",
        "UNAVAILABLE", "SERVICE_UNAVAILABLE", "DEADLINE_EXCEEDED",
        "HIGH DEMAND", "TEMPORARILY", "OVERLOADED",
    )
    return any(marker in text for marker in markers)


def _run(client, contents, primary_config, fallback_config):
    """SOP2: Gemini 3.8 优先；高峰/限流时直接切 Gemini 3.5。"""
    primary_model = SOP2_MODEL_CHAIN[0] if SOP2_MODEL_CHAIN else "gemini-3.8-flash"
    fallback_model = "gemini-3.5-flash-lite"

    try:
        response, meta = _base_generate_resilient(
            client,
            contents,
            primary_config,
            model_chain=[primary_model],
            max_attempts_per_model=1,
        )
        meta = dict(meta or {})
        meta.setdefault("primary_model", primary_model)
        meta.setdefault("fallback_used", "")
        return response, meta
    except Exception as exc:
        raw_exc = getattr(exc, "__cause__", None) or exc
        if not _is_transient_error(raw_exc):
            print(
                f"[SOP2 Gemini RAW ERROR] type={type(raw_exc).__name__} error={raw_exc}",
                flush=True,
            )
            raise raw_exc

        print(
            f"[SOP2] {primary_model} busy/unavailable -> fallback {fallback_model}: {raw_exc}",
            flush=True,
        )

        try:
            response, meta = _base_generate_resilient(
                client,
                contents,
                fallback_config,
                model_chain=[fallback_model],
                max_attempts_per_model=2,
            )
            meta = dict(meta or {})
            meta["primary_model"] = primary_model
            meta["fallback_used"] = fallback_model
            return response, meta
        except Exception as fallback_exc:
            raw_fallback = getattr(fallback_exc, "__cause__", None) or fallback_exc
            print(
                f"[SOP2 Gemini FALLBACK ERROR] type={type(raw_fallback).__name__} error={raw_fallback}",
                flush=True,
            )
            raise raw_fallback


def _thinking(level):
    # Gemini 3.8 Flash 支持 low/medium/high。
    return types.ThinkingConfig(thinking_level=level)


def _fallback_thinking():
    # Gemini 3.5 Flash Lite 走最轻 thinking，优先保证高峰期可用性与速度。
    return types.ThinkingConfig(thinking_level="minimal")

def _build_parts(client, labeled_videos):
    total_bytes = sum(len(v.getvalue()) for _, _, v in labeled_videos)
    total_mb = total_bytes / 1024 / 1024
    remote_files = []
    temp_paths = []
    parts = []
    if total_mb <= INLINE_BATCH_MAX_MB:
        mode = "Inline快速线路"
        for group, index, video in labeled_videos:
            parts.append(types.Part.from_text(text=f"【{group}{index}】文件名：{video.name}"))
            parts.append(types.Part.from_bytes(data=video.getvalue(), mime_type=video.type or "video/mp4"))
    else:
        mode = "Files API大文件线路"
        for group, index, video in labeled_videos:
            suffix = Path(video.name).suffix or ".mp4"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(video.getvalue())
                temp_path = tmp.name
            temp_paths.append(temp_path)
            remote = client.files.upload(file=temp_path)
            remote = wait_until_active(client, remote)
            remote_files.append(remote)
            parts.append(types.Part.from_text(text=f"【{group}{index}】文件名：{video.name}"))
            parts.append(types.Part.from_uri(file_uri=remote.uri, mime_type=remote.mime_type or "video/mp4"))
    return parts, remote_files, temp_paths, total_mb, mode


def _cleanup(client, remote_files, temp_paths):
    for remote in remote_files:
        try: client.files.delete(name=remote.name)
        except Exception: pass
    for path in temp_paths:
        try:
            if os.path.exists(path): os.remove(path)
        except OSError: pass


def build_pre_analysis_prompt(category, product_name, user_points, viral_names, own_names):
    return f"""
你是美国 TikTok Shop 多视频对标分析负责人。这里是 SOP2：爆款视频 VS 我们自己已经拍摄的作品。

除文件名和必须保留的英文原句外，所有分析必须使用简体中文。

产品品类：{category}
产品名称：{product_name}
我们的真实产品卖点：{user_points if clean_text(user_points) else '未填写'}
爆款视频：{viral_names}
我的作品：{own_names}

本阶段做“详细预分析 + 优缺点诊断 + 推荐比较组合”，不要只给一句话短评，也不要提前输出最终重剪方案。

【我们的固定拍摄边界｜必须遵守】
- 我们的视频默认不真人露脸、不正面出镜，不把“真人出镜/颜值/表情/口播者个人感染力”作为可执行优势。
- 可使用第一人称 POV；也可使用第三人称手部/肩部/少量身体局部，但不得出现可识别正脸。
- 拍摄环境以真实民宿/家居小场景为主，优先单一小场景内完成，不依赖棚拍或复杂外景。
- 前3秒优先靠强手部动作、产品结果、痛点反差、物体交互、字幕/配音和镜头节奏抓人。
- 如果爆款依赖真人露脸、人物表演或正面口播，允许在“源视频特征”中客观描述，但：
  1) 不得把它列为我们落后的核心原因；
  2) 不得把“增加真人露脸”写成优化建议；
  3) 必须把其中真正可复制的机制翻译成无露脸方案，例如：POV动作、手部演示、物体互动、镜头前后对比、字幕/配音、产品更早出现、Proof更强。
- 优缺点对比只比较我们能够实际调整的维度：Hook机制、画面动作、产品出现时间、Demo/Proof、节奏、字幕/口播信息、镜头丰富度、CTA、场景利用等。

请分别逐条分析爆款视频和我的作品，每条都必须输出足够让运营/剪辑直接判断的内容：
- 一句话核心：不要真的只有一句短句，写成2-3句，说明它在卖什么、靠什么吸引人。
- 脚本路线：按 Hook → 痛点/需求 → 产品切入 → Demo/Proof → CTA 的顺序详细说明，建议120-220字。
- 2-5条主要卖点。
- 前3秒Hook：描述具体画面、动作、字幕/口播、为什么能留人，建议80-160字。
- 画面与节奏：镜头数量/变化、景别、手部动作、产品展示、信息密度、转场节奏，建议100-220字。
- strengths：3-6条，每条必须是“观察到的优点 + 为什么有效”，不能只写形容词。
- weaknesses：3-6条，每条必须指出短板/风险/可优化处；爆款也必须挑短板，不能因为是爆款就只说优点。
- why_it_works：解释这条视频真正起作用的机制，建议100-200字。
- improvement_priority：如果只允许改一个问题，最优先改什么以及原因。
- 作为对比样本的价值。
- 推荐指数0-100。

然后推荐 1 条爆款 + 1 条我的作品作为最值得深入比较的组合。
推荐依据优先考虑：卖点/购买动机相近、产品Demo路径可比、脚本结构可比、最有利于定位“为什么爆款更强或我们哪里更强”。

同时输出 quick_comparison，对 AI 推荐的这一组做预览级优缺点对比：
- 爆款优势 3-6条
- 爆款短板 3-6条
- 我的优势 3-6条
- 我的短板 3-6条
- 最大差距
- 最值得吸收的3-5点
每一条都要具体到画面/节奏/卖点/字幕/动作，不要写“更抓人、节奏更好”这种空话。

重要：推荐只用于提示。最终比较对象必须由使用人自己选择。
不得虚构平台真实数据、销量、认证、产品能力。严格按 JSON Schema 输出。
""".strip()


def pre_analyze(client, viral_videos, own_videos, category, product_name, user_points):
    if not viral_videos or not own_videos:
        raise ValueError("爆款视频和我的作品都至少上传1条。")
    if len(viral_videos) > SOP2_MAX_VIRAL_VIDEOS or len(own_videos) > SOP2_MAX_OWN_VIDEOS:
        raise ValueError("SOP2 单次最多 3 条爆款 + 3 条我的作品。")
    if len(viral_videos) + len(own_videos) > SOP2_MAX_TOTAL_VIDEOS:
        raise ValueError("SOP2 单次最多处理6条视频。")
    started = time.perf_counter()
    labeled = [("爆款视频", i, v) for i, v in enumerate(viral_videos,1)] + [("我的作品", i, v) for i,v in enumerate(own_videos,1)]
    parts, remote_files, temp_paths, total_mb, mode = _build_parts(client, labeled)
    try:
        parts.append(types.Part.from_text(text=build_pre_analysis_prompt(category, product_name, user_points, [v.name for v in viral_videos], [v.name for v in own_videos])))
        content = types.Content(role="user", parts=parts)
        system_instruction = "这是给中国运营团队看的多视频预分析。所有分析字段必须简体中文。你只有推荐权，没有最终选择权。"
        primary_config = types.GenerateContentConfig(
            system_instruction=system_instruction,
            thinking_config=_thinking("low"),
            max_output_tokens=7600,
            response_mime_type="application/json",
            response_json_schema=SOP2_PRE_ANALYSIS_SCHEMA,
        )
        fallback_config = types.GenerateContentConfig(
            system_instruction=system_instruction,
            thinking_config=_fallback_thinking(),
            max_output_tokens=7600,
            response_mime_type="application/json",
            response_json_schema=SOP2_PRE_ANALYSIS_SCHEMA,
        )
        response, meta = _run(client, content, primary_config, fallback_config)
        result = parse_json_output(response.text)
        return result, {"analysis_seconds": round(time.perf_counter()-started,1), "analysis_mode":mode, "total_size_mb":round(total_mb,2), **meta}
    finally:
        _cleanup(client, remote_files, temp_paths)


def build_deep_compare_prompt(category, product_name, user_points, viral_summary, own_summary):
    dimensions = "第一帧、前3秒Hook、核心卖点、痛点表达、产品出现时间、Demo动作、镜头节奏、字幕/口播、信任证明、CTA/成交路径"
    return f"""
你是美国 TikTok Shop 视频复刻与剪辑诊断负责人。
现在请直接观看选中的两条原视频，做 SOP2 深度对比。

产品品类：{category}
产品名称：{product_name}
我们的真实产品卖点：{user_points if clean_text(user_points) else '未填写'}

爆款预分析摘要：{viral_summary}
我的作品预分析摘要：{own_summary}

所有内部分析都必须用简体中文。
目标不是泛泛说“爆款更抓人”，而是告诉剪辑/拍摄人员下一步具体怎么改。

【我们的固定拍摄边界｜深度对比必须遵守】
- 我们不以真人正脸/正面出镜为拍摄方案。默认只允许第一人称 POV、第三人称手部/肩部/少量身体局部，且不出现可识别正脸。
- 场景以真实民宿/家居小场景为主。
- 如果对方爆款有真人露脸、人物表演、正面口播：可以描述这是源视频的呈现方式，但不能把“真人出镜”本身判定为其核心优势，也不能把“我们没有真人出镜”判定为核心短板。
- 必须继续往下拆：真人出镜背后真正产生效果的是哪一个可复制机制，例如情绪表达、信任感、注意力锚点、信息承载、动作演示、节奏切换；然后转换成我们可执行的无露脸方案。
- 所有建议都必须落到我们能调整的内容：POV/第三人称机位、手部动作、产品更早出现、前后对比、Proof、镜头数量和景别变化、字幕/配音、CTA、小场景内道具与路径。
- 禁止建议“让真人露脸”“增加正面口播人物”“靠人物颜值/表情提升”等不符合生产条件的方案。

固定输出：
1. 一句话结论：为什么爆款更强、我们更强在哪里，最核心差距是什么。
2. 爆款脚本路线。
3. 我的脚本路线。
4. 核心差距。
5. 严格输出10个对比维度，顺序必须覆盖：{dimensions}。
   每项包含：爆款表现、我的表现、核心差距、具体建议。每个“表现/差距/建议”都要尽量写到具体画面、动作、字幕、时间位置或信息密度，避免一句话敷衍。
6. 爆款优势 3-6 条：不仅说强，还要说明强在哪、为什么有效。
7. 爆款短板 3-6 条：必须挑出爆款也存在的不足、风险或不可照搬处。
8. 我的优势 3-6 条：指出已经做对、值得保留甚至放大的部分。
9. 我的短板 3-6 条：指出最影响留存/理解/点击/转化的具体问题。
10. 重剪价值只能判断为：高 / 中 / 低，并解释原因。
11. 直接给时间段级重剪计划：
   - 可以保留 keep_segments
   - 建议删除 delete_segments
   - 建议前移/调整 move_segments
   - 必须补拍 reshoot_segments
12. editing_plan：按优先级说明整条如何重剪。
13. optimization_plan：下次拍摄应该吸收什么，避免机械复制爆款。这里必须输出“无露脸可执行版本”；如源视频依赖真人出镜，必须把其底层机制改写成 POV/手部/产品互动方案。

时间码可以根据视频画面估算到可供剪辑定位的程度，但不要假装是逐帧测量；没有明确证据时不要虚构精准帧号。
不要因为爆款这么拍，就强行让我们的产品展示不存在的功能；真实产品卖点优先。
不得虚构平台表现数据、销量、认证、产品功能。严格按 JSON Schema 输出。
""".strip()


def deep_compare(client, viral_video, own_video, category, product_name, user_points, viral_summary, own_summary):
    started = time.perf_counter()
    labeled = [("选中爆款",1,viral_video),("选中我的作品",1,own_video)]
    parts, remote_files, temp_paths, total_mb, mode = _build_parts(client, labeled)
    try:
        parts.append(types.Part.from_text(text=build_deep_compare_prompt(category, product_name, user_points, viral_summary, own_summary)))
        content = types.Content(role="user", parts=parts)
        system_instruction = "这是给中国团队执行的爆款对比、重剪和补拍诊断。所有分析必须简体中文；结论必须具体、可执行。"
        primary_config = types.GenerateContentConfig(
            system_instruction=system_instruction,
            thinking_config=_thinking("medium"),
            max_output_tokens=11000,
            response_mime_type="application/json",
            response_json_schema=SOP2_DEEP_COMPARE_SCHEMA,
        )
        fallback_config = types.GenerateContentConfig(
            system_instruction=system_instruction,
            thinking_config=_fallback_thinking(),
            max_output_tokens=11000,
            response_mime_type="application/json",
            response_json_schema=SOP2_DEEP_COMPARE_SCHEMA,
        )
        response, meta = _run(client, content, primary_config, fallback_config)
        result = parse_json_output(response.text)
        return result, {"analysis_seconds":round(time.perf_counter()-started,1), "analysis_mode":mode, "total_size_mb":round(total_mb,2), **meta}
    finally:
        _cleanup(client, remote_files, temp_paths)
