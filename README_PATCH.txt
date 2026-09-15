SOP1 AI参数异常修复

修复内容：
1. 爆款拆解 max_output_tokens 从 16000 调整为 7800，避免 Gemini Flash-Lite 输出上限导致 400 INVALID_ARGUMENT。
2. 最终脚本 max_output_tokens 从 8200 调整为 7600。
3. subtitle_segments 移除 maxItems=80，降低结构化输出 Schema 复杂度；仍通过提示词要求覆盖原视频完整时长。
4. 保留英/西双字幕、完整时长、简化机位逻辑。

部署：覆盖 main 的 gemini_sop1.py 与 sop1_schema.py。
