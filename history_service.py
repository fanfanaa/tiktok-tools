import json
import os
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import pandas as pd
import streamlit as st

from config import HISTORY_FILE, HISTORY_COLUMNS, CN_TZ
from common import clean_text, get_secret

HISTORY_LOCK = threading.Lock()
PENDING_HISTORY_FILE = Path("history_pending.csv")
SUPABASE_TABLE = "tiktok_history"
SUPABASE_TIMEOUT_SEC = 20


def empty_history():
    return pd.DataFrame(columns=HISTORY_COLUMNS)


def normalize_history(dataframe):
    if dataframe is None:
        return empty_history()

    dataframe = dataframe.copy()

    for column in HISTORY_COLUMNS:
        if column not in dataframe.columns:
            dataframe[column] = ""

    dataframe = dataframe.reindex(columns=HISTORY_COLUMNS)

    for column in HISTORY_COLUMNS:
        dataframe[column] = dataframe[column].map(clean_text)

    return dataframe


def _supabase_config():
    url = get_secret("SUPABASE_URL", "").rstrip("/")
    secret_key = get_secret("SUPABASE_SECRET_KEY", "")
    return url, secret_key


def _supabase_ready():
    url, secret_key = _supabase_config()
    return bool(url and secret_key)


def _supabase_request(method, endpoint, payload=None, prefer=""):
    url, secret_key = _supabase_config()

    if not url or not secret_key:
        raise RuntimeError("Supabase 未配置。")

    body = None
    headers = {
        "apikey": secret_key,
        "Accept": "application/json",
        "User-Agent": "TikTok-Video-SOP/1.0",
    }

    if payload is not None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json; charset=utf-8"

    if prefer:
        headers["Prefer"] = prefer

    request = Request(
        url=f"{url}/rest/v1/{endpoint}",
        data=body,
        headers=headers,
        method=method,
    )

    try:
        with urlopen(request, timeout=SUPABASE_TIMEOUT_SEC) as response:
            raw = response.read().decode("utf-8")
            if not raw:
                return None
            return json.loads(raw)

    except HTTPError as exc:
        try:
            detail = exc.read().decode("utf-8", errors="replace")
        except Exception:
            detail = ""
        raise RuntimeError(
            f"Supabase 请求失败（HTTP {exc.code}）：{detail[:300]}"
        ) from exc

    except URLError as exc:
        raise RuntimeError(f"Supabase 网络连接失败：{exc.reason}") from exc


def _load_local_file(path):
    if not path.exists():
        return empty_history()

    try:
        dataframe = pd.read_csv(
            path,
            dtype=str,
            keep_default_na=False,
            encoding="utf-8-sig",
        )
        return normalize_history(dataframe)
    except Exception:
        return empty_history()


def _write_local_file(path, dataframe):
    dataframe = normalize_history(dataframe)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_file = path.with_name(path.stem + ".tmp" + path.suffix)
    dataframe.to_csv(temp_file, index=False, encoding="utf-8-sig")
    os.replace(temp_file, path)


def _prepare_row(record, preserve_timestamps=False):
    row = {
        column: clean_text(record.get(column, ""))
        for column in HISTORY_COLUMNS
    }

    if not row["record_id"]:
        row["record_id"] = uuid.uuid4().hex[:12]

    if not preserve_timestamps or not row["created_at_utc"]:
        now_utc = datetime.now(timezone.utc)
        row["created_at_utc"] = now_utc.isoformat(timespec="seconds")
        row["created_at_cn"] = (
            now_utc
            .astimezone(CN_TZ)
            .strftime("%Y-%m-%d %H:%M:%S")
        )
    elif not row["created_at_cn"]:
        try:
            parsed = datetime.fromisoformat(
                row["created_at_utc"].replace("Z", "+00:00")
            )
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=timezone.utc)
            row["created_at_cn"] = (
                parsed
                .astimezone(CN_TZ)
                .strftime("%Y-%m-%d %H:%M:%S")
            )
        except Exception:
            pass

    return row


def _upsert_rows(rows):
    if not rows:
        return

    _supabase_request(
        "POST",
        f"{SUPABASE_TABLE}?on_conflict=record_id",
        payload=rows,
        prefer="resolution=merge-duplicates,return=minimal",
    )


def _append_pending(row):
    with HISTORY_LOCK:
        current = _load_local_file(PENDING_HISTORY_FILE)
        updated = pd.concat(
            [current, pd.DataFrame([row])],
            ignore_index=True,
        )
        updated = updated.drop_duplicates(subset=["record_id"], keep="last")
        _write_local_file(PENDING_HISTORY_FILE, updated)


def _sync_pending():
    if not _supabase_ready() or not PENDING_HISTORY_FILE.exists():
        return

    pending = _load_local_file(PENDING_HISTORY_FILE)
    if pending.empty:
        try:
            PENDING_HISTORY_FILE.unlink(missing_ok=True)
        except Exception:
            pass
        return

    try:
        rows = pending.to_dict(orient="records")
        for start in range(0, len(rows), 100):
            _upsert_rows(rows[start:start + 100])
        PENDING_HISTORY_FILE.unlink(missing_ok=True)
    except Exception:
        pass


def load_history():
    if _supabase_ready():
        _sync_pending()
        try:
            data = _supabase_request(
                "GET",
                f"{SUPABASE_TABLE}?select=*&order=created_at_utc.asc",
            )
            return normalize_history(pd.DataFrame(data or []))
        except Exception:
            pass

    # 兼容旧版 / 临时网络故障：仍可读取本地历史。
    local_history = _load_local_file(HISTORY_FILE)
    pending_history = _load_local_file(PENDING_HISTORY_FILE)

    if local_history.empty:
        return pending_history
    if pending_history.empty:
        return local_history

    combined = pd.concat(
        [local_history, pending_history],
        ignore_index=True,
    )
    combined = combined.drop_duplicates(subset=["record_id"], keep="last")
    return normalize_history(combined)


def write_history(dataframe):
    """兼容旧调用：优先整批同步到 Supabase，否则写本地。"""
    dataframe = normalize_history(dataframe)

    if _supabase_ready():
        try:
            rows = dataframe.to_dict(orient="records")
            for start in range(0, len(rows), 100):
                _upsert_rows(rows[start:start + 100])
            return
        except Exception:
            pass

    _write_local_file(HISTORY_FILE, dataframe)


def append_history(record):
    row = _prepare_row(record, preserve_timestamps=False)

    if _supabase_ready():
        try:
            _sync_pending()
            _upsert_rows([row])
            return
        except Exception:
            # 不让历史库瞬时故障中断 AI 主流程；先放入本地待同步队列。
            _append_pending(row)
            return

    # 未配置 Supabase 时继续兼容旧版。
    with HISTORY_LOCK:
        current = _load_local_file(HISTORY_FILE)
        updated = pd.concat(
            [current, pd.DataFrame([row])],
            ignore_index=True,
        )
        _write_local_file(HISTORY_FILE, updated)


def import_history_dataframe(dataframe):
    """管理员把旧版下载的历史 CSV 一次性导入 Supabase。重复 record_id 会更新，不会重复新增。"""
    dataframe = normalize_history(dataframe)

    if dataframe.empty:
        return 0

    rows = []
    for record in dataframe.to_dict(orient="records"):
        rows.append(_prepare_row(record, preserve_timestamps=True))

    # 防止旧 CSV 内部本身存在重复 record_id。
    deduped = {}
    for row in rows:
        deduped[row["record_id"]] = row
    rows = list(deduped.values())

    if not _supabase_ready():
        raise RuntimeError("系统尚未配置 Supabase，无法导入。")

    for start in range(0, len(rows), 100):
        _upsert_rows(rows[start:start + 100])

    return len(rows)


def scoped_history():
    dataframe = load_history()

    if st.session_state["role"] == "主账号(Admin)":
        return dataframe

    return dataframe[
        dataframe["operator"] == st.session_state["operator"]
    ].copy()
