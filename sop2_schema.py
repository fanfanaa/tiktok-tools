# SOP2｜爆款 VS 我的作品：独立 Schema

_VIDEO_PRE_ITEM = {
    "type": "object",
    "properties": {
        "video_index": {"type": "integer"},
        "filename": {"type": "string"},
        "one_sentence_core": {"type": "string"},
        "script_route": {"type": "string"},
        "selling_points": {"type": "array", "minItems": 2, "maxItems": 5, "items": {"type": "string"}},
        "first_3s_hook": {"type": "string"},
        "visual_rhythm": {"type": "string"},
        "strengths": {"type": "array", "minItems": 3, "maxItems": 6, "items": {"type": "string"}},
        "weaknesses": {"type": "array", "minItems": 3, "maxItems": 6, "items": {"type": "string"}},
        "why_it_works": {"type": "string"},
        "improvement_priority": {"type": "string"},
        "compare_value": {"type": "string"},
        "recommend_score": {"type": "integer"},
    },
    "required": [
        "video_index", "filename", "one_sentence_core", "script_route", "selling_points",
        "first_3s_hook", "visual_rhythm", "strengths", "weaknesses", "why_it_works",
        "improvement_priority", "compare_value", "recommend_score",
    ],
}

SOP2_PRE_ANALYSIS_SCHEMA = {
    "type": "object",
    "properties": {
        "viral_videos": {"type": "array", "minItems": 1, "maxItems": 3, "items": _VIDEO_PRE_ITEM},
        "own_videos": {"type": "array", "minItems": 1, "maxItems": 3, "items": _VIDEO_PRE_ITEM},
        "recommended_viral_index": {"type": "integer"},
        "recommended_own_index": {"type": "integer"},
        "recommendation_reason": {"type": "string"},
        "quick_comparison": {
            "type": "object",
            "properties": {
                "viral_advantages": {"type": "array", "minItems": 3, "maxItems": 6, "items": {"type": "string"}},
                "viral_weaknesses": {"type": "array", "minItems": 3, "maxItems": 6, "items": {"type": "string"}},
                "own_advantages": {"type": "array", "minItems": 3, "maxItems": 6, "items": {"type": "string"}},
                "own_weaknesses": {"type": "array", "minItems": 3, "maxItems": 6, "items": {"type": "string"}},
                "biggest_gap": {"type": "string"},
                "most_worth_learning": {"type": "array", "minItems": 3, "maxItems": 5, "items": {"type": "string"}},
            },
            "required": [
                "viral_advantages", "viral_weaknesses", "own_advantages", "own_weaknesses",
                "biggest_gap", "most_worth_learning",
            ],
        },
    },
    "required": [
        "viral_videos", "own_videos", "recommended_viral_index", "recommended_own_index",
        "recommendation_reason", "quick_comparison",
    ],
}

SOP2_DEEP_COMPARE_SCHEMA = {
    "type": "object",
    "properties": {
        "one_sentence_conclusion": {"type": "string"},
        "viral_script_route": {"type": "string"},
        "own_script_route": {"type": "string"},
        "core_gap": {"type": "string"},
        "comparison_dimensions": {
            "type": "array", "minItems": 10, "maxItems": 10,
            "items": {
                "type": "object",
                "properties": {
                    "dimension": {"type": "string"},
                    "viral": {"type": "string"},
                    "own": {"type": "string"},
                    "gap": {"type": "string"},
                    "suggestion": {"type": "string"},
                },
                "required": ["dimension", "viral", "own", "gap", "suggestion"],
            },
        },
        "viral_strengths": {"type": "array", "minItems": 3, "maxItems": 6, "items": {"type": "string"}},
        "viral_weaknesses": {"type": "array", "minItems": 3, "maxItems": 6, "items": {"type": "string"}},
        "own_strengths": {"type": "array", "minItems": 3, "maxItems": 6, "items": {"type": "string"}},
        "own_weaknesses": {"type": "array", "minItems": 3, "maxItems": 6, "items": {"type": "string"}},
        "reedit_value": {"type": "string"},
        "reedit_reason": {"type": "string"},
        "keep_segments": {
            "type": "array", "maxItems": 8,
            "items": {"type":"object","properties":{"time_range":{"type":"string"},"content":{"type":"string"},"reason":{"type":"string"}},"required":["time_range","content","reason"]},
        },
        "delete_segments": {
            "type": "array", "maxItems": 8,
            "items": {"type":"object","properties":{"time_range":{"type":"string"},"content":{"type":"string"},"reason":{"type":"string"}},"required":["time_range","content","reason"]},
        },
        "move_segments": {
            "type": "array", "maxItems": 8,
            "items": {"type":"object","properties":{"source_time":{"type":"string"},"target_time":{"type":"string"},"content":{"type":"string"},"reason":{"type":"string"}},"required":["source_time","target_time","content","reason"]},
        },
        "reshoot_segments": {
            "type": "array", "maxItems": 8,
            "items": {"type":"object","properties":{"shot":{"type":"string"},"action":{"type":"string"},"purpose":{"type":"string"}},"required":["shot","action","purpose"]},
        },
        "editing_plan": {"type": "string"},
        "optimization_plan": {"type": "string"},
    },
    "required": [
        "one_sentence_conclusion", "viral_script_route", "own_script_route", "core_gap", "comparison_dimensions",
        "viral_strengths", "viral_weaknesses", "own_strengths", "own_weaknesses", "reedit_value", "reedit_reason",
        "keep_segments", "delete_segments", "move_segments", "reshoot_segments", "editing_plan", "optimization_plan",
    ],
}
