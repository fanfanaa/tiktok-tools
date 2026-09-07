from __future__ import annotations

import threading
import time
import uuid
from concurrent.futures import Future, ThreadPoolExecutor
from typing import Any, Callable

import streamlit as st


# 线程池与任务表放在模块级：Streamlit 页面切换会 rerun 页面脚本，
# 但同一 Python 进程中的模块与线程池仍然存在，因此后台任务不会因为换页被中断。
_EXECUTOR = ThreadPoolExecutor(max_workers=4, thread_name_prefix="tiktok-ai-bg")
_TASKS: dict[str, dict[str, Any]] = {}
_LOCK = threading.RLock()


def ensure_background_session_id() -> str:
    key = "_background_session_id"
    if not st.session_state.get(key):
        st.session_state[key] = uuid.uuid4().hex
    return st.session_state[key]


def task_key(name: str) -> str:
    return f"{ensure_background_session_id()}::{name}"


def submit_background_task(
    key: str,
    fn: Callable,
    *args,
    context: dict[str, Any] | None = None,
    **kwargs,
) -> bool:
    """提交任务。相同 key 已在运行时不重复提交。"""
    with _LOCK:
        existing = _TASKS.get(key)
        if existing is not None:
            future: Future = existing["future"]
            if not future.done():
                return False

        future = _EXECUTOR.submit(fn, *args, **kwargs)
        _TASKS[key] = {
            "future": future,
            "context": dict(context or {}),
            "submitted_at": time.time(),
        }
        return True


def background_task_status(key: str) -> dict[str, Any]:
    with _LOCK:
        record = _TASKS.get(key)
        if record is None:
            return {"status": "idle"}

        future: Future = record["future"]
        base = {
            "context": dict(record.get("context") or {}),
            "submitted_at": record.get("submitted_at"),
        }

        if future.cancelled():
            return {**base, "status": "cancelled"}
        if future.done():
            try:
                future.result()
            except Exception as exc:
                return {**base, "status": "error", "error": exc}
            return {**base, "status": "done"}
        if future.running():
            return {**base, "status": "running"}
        return {**base, "status": "queued"}


def take_background_task(key: str) -> dict[str, Any] | None:
    """任务完成后一次性取走结果；未完成返回 None。"""
    with _LOCK:
        record = _TASKS.get(key)
        if record is None:
            return None
        future: Future = record["future"]
        if not future.done():
            return None
        _TASKS.pop(key, None)

    payload = {
        "context": dict(record.get("context") or {}),
        "submitted_at": record.get("submitted_at"),
    }
    if future.cancelled():
        return {**payload, "status": "cancelled"}
    try:
        result = future.result()
        return {**payload, "status": "done", "result": result}
    except Exception as exc:
        return {**payload, "status": "error", "error": exc}


def clear_background_task(key: str) -> None:
    with _LOCK:
        record = _TASKS.pop(key, None)
    if record is not None:
        future: Future = record["future"]
        if not future.done():
            future.cancel()


def render_background_task_status(key: str, running_text: str) -> None:
    """显示后台状态；停留当前页时每2秒检查，完成后自动刷新整页。"""
    status = background_task_status(key)
    if status.get("status") not in {"queued", "running"}:
        return

    st.info(running_text + "\n\n可以切换到其他页面，任务会继续在后台运行。")

    # Streamlit 1.62 支持 fragment。它只负责轻量轮询，不执行 AI。
    if hasattr(st, "fragment"):
        @st.fragment(run_every="2s")
        def _poll_background_task():
            current = background_task_status(key)
            if current.get("status") in {"done", "error", "cancelled"}:
                st.rerun()
            else:
                st.caption("后台任务仍在运行…")

        _poll_background_task()
