# 数据复盘独立结构化输出 Schema

REVIEW_SCHEMA = {

    "type": "object",

    "properties": {

        "priority_issue": {
            "type": "string",
        },

        "diagnosis_summary": {
            "type": "string",
        },

        "account_diagnosis": {
            "type": "string",
        },

        "metric_diagnosis": {

            "type": "array",
            "minItems": 5,
            "maxItems": 8,

            "items": {

                "type": "object",

                "properties": {

                    "metric": {
                        "type": "string",
                    },

                    "status": {
                        "type": "string",
                    },

                    "meaning": {
                        "type": "string",
                    },

                    "action": {
                        "type": "string",
                    },
                },

                "required": [
                    "metric",
                    "status",
                    "meaning",
                    "action",
                ],
            },
        },

        "optimized_script": {

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
        "priority_issue",
        "diagnosis_summary",
        "account_diagnosis",
        "metric_diagnosis",
        "optimized_script",
    ],
}
