import hashlib
import json
import pandas as pd
import streamlit as st

def clean_text(value):

    if value is None:
        return ""

    try:

        if pd.isna(value):
            return ""

    except Exception:
        pass

    return str(value).strip()

def get_secret(
    name,
    default="",
):

    try:

        value = st.secrets[name]

    except Exception:

        return default

    if value is None:
        return default

    return str(value).strip()

def get_api_key():

    return get_secret(
        "GEMINI_API_KEY",
        "",
    )

def json_dumps(
    data,
):

    return json.dumps(
        data,
        ensure_ascii=False,
        separators=(",", ":"),
    )

def parse_json_output(
    raw_text,
):

    if not raw_text:

        raise ValueError(
            "AI未返回有效结果。"
        )

    try:

        return json.loads(
            raw_text
        )

    except json.JSONDecodeError as exc:

        raise ValueError(
            "AI返回格式异常，请重新执行。"
        ) from exc

def compact_dict(
    data,
):

    return {

        key: value

        for key, value
        in data.items()

        if value is not None
        and value != ""
    }

def list_to_joined(
    items,
):

    if not items:
        return ""

    return "; ".join(
        [
            clean_text(item)
            for item
            in items
            if clean_text(item)
        ]
    )

def safe_int(
    value,
    default=0,
):

    try:

        return int(
            value
        )

    except Exception:

        return default

def parse_optional_float(
    value,
):

    value = clean_text(
        value
    )

    if not value:
        return None

    try:

        return float(
            value
        )

    except ValueError:

        return None

def make_signature(
    *values,
):

    serialized = []

    for value in values:

        if isinstance(
            value,
            (
                dict,
                list,
            )
        ):

            serialized.append(
                json.dumps(
                    value,
                    ensure_ascii=False,
                    sort_keys=True,
                )
            )

        else:

            serialized.append(
                clean_text(
                    value
                )
            )

    raw = "||".join(
        serialized
    )

    return hashlib.sha256(
        raw.encode(
            "utf-8"
        )
    ).hexdigest()[:24]

def video_batch_signature(
    uploaded_videos,
    category,
    product_name,
    input_selling_points,
):

    hasher = hashlib.sha256()

    for video in uploaded_videos or []:

        hasher.update(
            video.getvalue()
        )

        hasher.update(
            clean_text(
                video.name
            ).encode(
                "utf-8"
            )
        )

    hasher.update(
        clean_text(
            category
        ).encode(
            "utf-8"
        )
    )

    hasher.update(
        clean_text(
            product_name
        ).encode(
            "utf-8"
        )
    )

    hasher.update(
        clean_text(
            input_selling_points
        ).encode(
            "utf-8"
        )
    )

    return hasher.hexdigest()[:24]

