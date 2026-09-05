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


def _run(client, contents, config):
    return _base_generate_resilient(client, contents, config, model_chain=SOP2_MODEL_CHAIN)


def _thinking(level):
    # Gemini 3.8 Flash: low / medium / high；不使用 minimal
    return types.ThinkingConfig(thinking_level=level)


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

本阶段只做“快速预分析”和“推荐比较组合”，不要做最终重剪方案。

请分别逐条分析爆款视频和我的作品，每条都输出：
- 一句话核心
- 脚本路线
- 2-5条主要卖点
- 前3秒Hook
- 画面与节奏
- 作为对比样本的价值
- 推荐指数0-100

然后推荐 1 条爆款 + 1 条我的作品作为最值得深入比较的组合。
推荐依据优先考虑：卖点/购买动机相近、产品Demo路径可比、脚本结构可比、最有利于定位“为什么爆款更强或我们哪里更强”。

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
        config = types.GenerateContentConfig(
            system_instruction="这是给中国运营团队看的多视频预分析。所有分析字段必须简体中文。你只有推荐权，没有最终选择权。",
            thinking_config=_thinking("low"),
            max_output_tokens=5000,
            response_mime_type="application/json",
            response_json_schema=SOP2_PRE_ANALYSIS_SCHEMA,
        )
        response, meta = _run(client, content, config)
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

固定输出：
1. 一句话结论：为什么爆款更强、我们更强在哪里，最核心差距是什么。
2. 爆款脚本路线。
3. 我的脚本路线。
4. 核心差距。
5. 严格输出10个对比维度，顺序必须覆盖：{dimensions}。
   每项包含：爆款表现、我的表现、核心差距、具体建议。
6. 我的优势 2-6 条。
7. 我的劣势 2-6 条。
8. 重剪价值只能判断为：高 / 中 / 低，并解释原因。
9. 直接给时间段级重剪计划：
   - 可以保留 keep_segments
   - 建议删除 delete_segments
   - 建议前移/调整 move_segments
   - 必须补拍 reshoot_segments
10. editing_plan：按优先级说明整条如何重剪。
11. optimization_plan：下次拍摄应该吸收什么，避免机械复制爆款。

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
        config = types.GenerateContentConfig(
            system_instruction="这是给中国团队执行的爆款对比、重剪和补拍诊断。所有分析必须简体中文；结论必须具体、可执行。",
            thinking_config=_thinking("medium"),
            max_output_tokens=9000,
            response_mime_type="application/json",
            response_json_schema=SOP2_DEEP_COMPARE_SCHEMA,
        )
        response, meta = _run(client, content, config)
        result = parse_json_output(response.text)
        return result, {"analysis_seconds":round(time.perf_counter()-started,1), "analysis_mode":mode, "total_size_mb":round(total_mb,2), **meta}
    finally:
        _cleanup(client, remote_files, temp_paths)
