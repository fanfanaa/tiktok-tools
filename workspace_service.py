import copy
import hashlib
import json
import uuid
from io import BytesIO
from typing import Any, Iterable

import streamlit as st

from common import clean_text
from history_service import append_history

_PAGE_MEMORY_ROOT = "_workspace_page_memory"
_LAST_SAVED_ROOT = "_workspace_last_saved_signature"


class RetainedUpload:
    """在同一个 Streamlit 会话中跨页面保留上传文件。"""

    def __init__(self, name: str, mime_type: str, data: bytes):
        self.name = name
        self.type = mime_type or "application/octet-stream"
        self._data = bytes(data)
        self.size = len(self._data)
        self._buffer = BytesIO(self._data)

    def getvalue(self) -> bytes:
        return self._data

    def read(self, *args, **kwargs):
        return self._buffer.read(*args, **kwargs)

    def seek(self, *args, **kwargs):
        return self._buffer.seek(*args, **kwargs)

    def tell(self):
        return self._buffer.tell()


def _memory_for(page_id: str) -> dict:
    root = st.session_state.setdefault(_PAGE_MEMORY_ROOT, {})
    return root.setdefault(page_id, {"widgets": {}, "uploads": {}})


def restore_page_memory(page_id: str, widget_keys: Iterable[str], dynamic_prefixes: Iterable[str] = ()):
    """把稳定副本写回 widget key，避免页面切换后 Streamlit 清掉 widget state。"""
    memory = _memory_for(page_id)
    allowed = set(widget_keys)
    prefixes = tuple(dynamic_prefixes)
    for key, value in memory.get("widgets", {}).items():
        if key in allowed or (prefixes and key.startswith(prefixes)):
            # 只有页面切换导致 widget key 被 Streamlit 清掉时才恢复。
            # 如果当前 key 已存在，说明这是同页交互（radio/selectbox/button rerun），
            # 必须保留用户刚刚的选择，不能用旧快照覆盖。
            if key in st.session_state:
                continue
            try:
                st.session_state[key] = copy.deepcopy(value)
            except Exception:
                st.session_state[key] = value


def snapshot_page_memory(page_id: str, widget_keys: Iterable[str], dynamic_prefixes: Iterable[str] = ()):
    """每次页面正常运行时，把当前 widget 值复制到不会被 Streamlit 清理的稳定区。"""
    memory = _memory_for(page_id)
    allowed = set(widget_keys)
    prefixes = tuple(dynamic_prefixes)
    saved = memory.setdefault("widgets", {})
    for key in list(st.session_state.keys()):
        if key in allowed or (prefixes and key.startswith(prefixes)):
            # uploader 单独保存，避免把 UploadedFile 对象塞进普通状态。
            if key.endswith("_upload") or key in {"analysis_videos", "review_excel"}:
                continue
            value = st.session_state.get(key)
            try:
                saved[key] = copy.deepcopy(value)
            except Exception:
                saved[key] = value


def remember_uploads(page_id: str, slot: str, files):
    if not files:
        return
    memory = _memory_for(page_id)
    memory.setdefault("uploads", {})[slot] = [
        {
            "name": clean_text(getattr(file, "name", "")),
            "type": clean_text(getattr(file, "type", "")) or "application/octet-stream",
            "data": bytes(file.getvalue()),
        }
        for file in files
    ]


def get_retained_uploads(page_id: str, slot: str):
    items = _memory_for(page_id).get("uploads", {}).get(slot, [])
    return [RetainedUpload(item["name"], item.get("type", ""), item["data"]) for item in items]


def clear_retained_uploads(page_id: str, slot: str | None = None):
    memory = _memory_for(page_id)
    if slot is None:
        memory["uploads"] = {}
    else:
        memory.setdefault("uploads", {}).pop(slot, None)


def retained_upload_metadata(page_id: str) -> dict:
    result = {}
    for slot, items in _memory_for(page_id).get("uploads", {}).items():
        result[slot] = [
            {"name": item.get("name", ""), "type": item.get("type", ""), "size": len(item.get("data", b""))}
            for item in items
        ]
    return result


def ensure_workspace_id(module: str) -> str:
    key = f"_workspace_id_{module.lower()}"
    if not st.session_state.get(key):
        st.session_state[key] = f"ws_{module.lower()}_{uuid.uuid4().hex[:10]}"
    return st.session_state[key]


def new_workspace_id(module: str) -> str:
    key = f"_workspace_id_{module.lower()}"
    st.session_state[key] = f"ws_{module.lower()}_{uuid.uuid4().hex[:10]}"
    # 新工作后允许立即写一次。
    st.session_state.setdefault(_LAST_SAVED_ROOT, {}).pop(module, None)
    return st.session_state[key]


def _json_safe(value: Any):
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, dict):
        return {str(k): _json_safe(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_json_safe(v) for v in value]
    return clean_text(value)


def collect_state(keys: Iterable[str], dynamic_prefixes: Iterable[str] = ()) -> dict:
    keys = set(keys)
    prefixes = tuple(dynamic_prefixes)
    result = {}
    for key in list(st.session_state.keys()):
        if key in keys or (prefixes and key.startswith(prefixes)):
            if key.endswith("_upload") or key in {"analysis_videos", "review_excel"}:
                continue
            result[key] = _json_safe(st.session_state.get(key))
    return result


def save_workspace_record(
    *,
    module: str,
    page_id: str,
    page_path: str,
    step: str,
    state_keys: Iterable[str],
    dynamic_prefixes: Iterable[str] = (),
    tiktok_account: str = "",
    product_category: str = "",
    product_name: str = "",
    input_selling_points: str = "",
    video_names: str = "",
    reference_video_name: str = "",
    direction_name: str = "",
    selected_scene: str = "",
    selected_perspective: str = "",
    meaningful: bool = True,
):
    """把当前工作状态作为“工作记录”自动 upsert；操作日志仍由原 append_history 单独保存。"""
    if not meaningful:
        return

    workspace_id = ensure_workspace_id(module)
    payload = {
        "schema": "workspace_v1",
        "workspace_id": workspace_id,
        "module": module,
        "page_id": page_id,
        "page_path": page_path,
        "step": step,
        "state": collect_state(state_keys, dynamic_prefixes),
        "upload_metadata": retained_upload_metadata(page_id),
    }
    payload_text = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    signature = hashlib.sha256(payload_text.encode("utf-8")).hexdigest()
    last_saved = st.session_state.setdefault(_LAST_SAVED_ROOT, {})
    if last_saved.get(module) == signature:
        return

    append_history(
        {
            "record_id": workspace_id,
            "module": module,
            "record_type": "工作记录",
            "role": st.session_state.get("role", ""),
            "operator": st.session_state.get("operator", ""),
            "tiktok_account": tiktok_account,
            "product_category": product_category,
            "product_name": product_name,
            "input_selling_points": input_selling_points,
            "reference_video_name": reference_video_name,
            "direction_name": direction_name,
            "selected_scene": selected_scene,
            "selected_perspective": selected_perspective,
            "video_names": video_names,
            "diagnosis_summary": step,
            "full_output_json": payload_text,
        }
    )
    last_saved[module] = signature


def restore_workspace_record(payload: dict):
    """从 Supabase 工作记录恢复可序列化状态；同会话内的视频字节由 page memory 继续保留。"""
    module = clean_text(payload.get("module", ""))
    workspace_id = clean_text(payload.get("workspace_id", ""))
    if module and workspace_id:
        st.session_state[f"_workspace_id_{module.lower()}"] = workspace_id

    state = payload.get("state", {})
    if isinstance(state, dict):
        for key, value in state.items():
            st.session_state[key] = value

    page_id = clean_text(payload.get("page_id", ""))
    if page_id:
        memory = _memory_for(page_id)
        memory.setdefault("widgets", {}).update(copy.deepcopy(state if isinstance(state, dict) else {}))

    st.session_state["_workspace_resume_notice"] = {
        "page_id": page_id,
        "has_upload_metadata": bool(payload.get("upload_metadata")),
    }


def render_resume_notice(page_id: str):
    notice = st.session_state.get("_workspace_resume_notice")
    if not isinstance(notice, dict) or notice.get("page_id") != page_id:
        return

    has_retained = any(_memory_for(page_id).get("uploads", {}).values())
    if has_retained:
        st.success("已恢复这条工作记录，当前会话中的上传文件也仍然保留，可继续操作。")
    elif notice.get("has_upload_metadata"):
        st.info("已恢复这条工作记录的选择与AI结果。原视频文件不会写入历史库；如需重新运行视频分析，请重新上传原视频。")
    else:
        st.success("已恢复这条工作记录，可继续操作。")

    st.session_state.pop("_workspace_resume_notice", None)
