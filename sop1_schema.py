from config import MAX_COMPARE_VIDEOS

# SOP1 独立结构化输出 Schema

VIDEO_ANALYSIS_SCHEMA = {

    "type": "object",

    "properties": {

        "comparison_summary": {

            "type": "object",

            "properties": {

                "one_sentence_core": {
                    "type": "string",
                },

                "common_script_route": {
                    "type": "string",
                },

                "common_audience": {
                    "type": "string",
                },

                "age_estimate": {
                    "type": "string",
                },

                "common_hook_pattern": {
                    "type": "string",
                },

                "visual_rhythm": {
                    "type": "string",
                },

                "top_absorb_points": {

                    "type": "array",
                    "minItems": 3,
                    "maxItems": 3,

                    "items": {
                        "type": "string",
                    },
                },

                "key_differences": {
                    "type": "string",
                },
            },

            "required": [
                "one_sentence_core",
                "common_script_route",
                "common_audience",
                "age_estimate",
                "common_hook_pattern",
                "visual_rhythm",
                "top_absorb_points",
                "key_differences",
            ],
        },

        "common_inferred_selling_points": {

            "type": "array",
            "minItems": 3,
            "maxItems": 5,

            "items": {
                "type": "string",
            },
        },

        "recommended_reference_video_index": {
            "type": "integer",
        },

        "videos": {

            "type": "array",
            "minItems": 1,
            "maxItems": MAX_COMPARE_VIDEOS,

            "items": {

                "type": "object",

                "properties": {

                    "video_index": {
                        "type": "integer",
                    },

                    "filename": {
                        "type": "string",
                    },

                    "one_sentence_core": {
                        "type": "string",
                    },

                    "inferred_selling_points": {

                        "type": "array",
                        "minItems": 3,
                        "maxItems": 5,

                        "items": {
                            "type": "string",
                        },
                    },

                    "script_route": {
                        "type": "string",
                    },

                    "audience_profile": {
                        "type": "string",
                    },

                    "age_estimate": {
                        "type": "string",
                    },

                    "first_3s_hook": {
                        "type": "string",
                    },

                    "visual_rhythm": {
                        "type": "string",
                    },

                    "top_absorb_points": {

                        "type": "array",
                        "minItems": 3,
                        "maxItems": 3,

                        "items": {
                            "type": "string",
                        },
                    },

                    "fit_reason": {
                        "type": "string",
                    },

                    "recommend_score": {
                        "type": "integer",
                    },

                    "selling_point_relation": {
                        "type": "string",
                    },

                    "selling_point_relation_reason": {
                        "type": "string",
                    },

                    "blended_selling_points": {
                        "type": "string",
                    },

                    "suggested_mode": {
                        "type": "string",
                    },
                },

                "required": [
                    "video_index",
                    "filename",
                    "one_sentence_core",
                    "inferred_selling_points",
                    "script_route",
                    "audience_profile",
                    "age_estimate",
                    "first_3s_hook",
                    "visual_rhythm",
                    "top_absorb_points",
                    "fit_reason",
                    "recommend_score",
                    "selling_point_relation",
                    "selling_point_relation_reason",
                    "blended_selling_points",
                    "suggested_mode",
                ],
            },
        },
    },

    "required": [
        "comparison_summary",
        "common_inferred_selling_points",
        "recommended_reference_video_index",
        "videos",
    ],
}

DIRECTIONS_SCHEMA = {

    "type": "object",

    "properties": {

        "directions": {

            "type": "array",
            "minItems": 3,
            "maxItems": 3,

            "items": {

                "type": "object",

                "properties": {

                    "direction_name": {
                        "type": "string",
                    },

                    "core_idea": {
                        "type": "string",
                    },

                    "target_audience": {
                        "type": "string",
                    },

                    "hook": {
                        "type": "string",
                    },

                    "product_entry": {
                        "type": "string",
                    },

                    "recommended_perspective": {
                        "type": "string",
                    },

                    "recommended_scene": {
                        "type": "string",
                    },

                    "absorb_points": {

                        "type": "array",
                        "minItems": 2,
                        "maxItems": 4,

                        "items": {
                            "type": "string",
                        },
                    },

                    "differentiation_points": {

                        "type": "array",
                        "minItems": 2,
                        "maxItems": 4,

                        "items": {
                            "type": "string",
                        },
                    },
                },

                "required": [
                    "direction_name",
                    "core_idea",
                    "target_audience",
                    "hook",
                    "product_entry",
                    "recommended_perspective",
                    "recommended_scene",
                    "absorb_points",
                    "differentiation_points",
                ],
            },
        },
    },

    "required": [
        "directions",
    ],
}

FINAL_SCRIPT_SCHEMA = {

    "type": "object",

    "properties": {

        "shooting_notes": {
            "type": "string",
        },

        "storyboard": {

            "type": "array",
            "minItems": 6,
            "maxItems": 12,

            "items": {

                "type": "object",

                "properties": {

                    "sequence": {
                        "type": "string",
                    },

                    "time_range": {
                        "type": "string",
                    },

                    "shot": {
                        "type": "string",
                    },

                    "visual": {
                        "type": "string",
                    },

                    "hand_action": {
                        "type": "string",
                    },

                    "copy_cn": {
                        "type": "string",
                    },

                    "copy_en": {
                        "type": "string",
                    },

                    "audio": {
                        "type": "string",
                    },

                    "absorb_point": {
                        "type": "string",
                    },

                    "difference_point": {
                        "type": "string",
                    },

                    "rationale": {
                        "type": "string",
                    },
                },

                "required": [
                    "sequence",
                    "time_range",
                    "shot",
                    "visual",
                    "hand_action",
                    "copy_cn",
                    "copy_en",
                    "audio",
                    "absorb_point",
                    "difference_point",
                    "rationale",
                ],
            },
        },
    },

    "required": [
        "shooting_notes",
        "storyboard",
    ],
}
