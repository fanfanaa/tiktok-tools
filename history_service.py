import os
import threading
import uuid
from datetime import datetime, timezone

import pandas as pd
import streamlit as st

from config import HISTORY_FILE, HISTORY_COLUMNS, CN_TZ
from common import clean_text

HISTORY_LOCK = threading.Lock()

def empty_history():

    return pd.DataFrame(
        columns=HISTORY_COLUMNS
    )

def normalize_history(
    dataframe,
):

    if dataframe is None:

        return empty_history()

    dataframe = (
        dataframe.copy()
    )

    for column in HISTORY_COLUMNS:

        if column not in dataframe.columns:

            dataframe[
                column
            ] = ""

    return dataframe.reindex(
        columns=HISTORY_COLUMNS
    )

def load_history():

    if not HISTORY_FILE.exists():

        return empty_history()

    try:

        dataframe = pd.read_csv(
            HISTORY_FILE,
            dtype=str,
            keep_default_na=False,
            encoding="utf-8-sig",
        )

        return normalize_history(
            dataframe
        )

    except Exception:

        return empty_history()

def write_history(
    dataframe,
):

    dataframe = normalize_history(
        dataframe
    )

    HISTORY_FILE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temp_file = (
        HISTORY_FILE.with_name(
            "history_log.tmp.csv"
        )
    )

    dataframe.to_csv(
        temp_file,
        index=False,
        encoding="utf-8-sig",
    )

    os.replace(
        temp_file,
        HISTORY_FILE,
    )

def append_history(
    record,
):

    row = {

        column:
            clean_text(
                record.get(
                    column,
                    "",
                )
            )

        for column
        in HISTORY_COLUMNS
    }

    if not row[
        "record_id"
    ]:

        row[
            "record_id"
        ] = uuid.uuid4().hex[:12]

    now_utc = (
        datetime.now(
            timezone.utc
        )
    )

    row[
        "created_at_utc"
    ] = now_utc.isoformat(
        timespec="seconds"
    )

    row[
        "created_at_cn"
    ] = (
        now_utc
        .astimezone(
            CN_TZ
        )
        .strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    )

    with HISTORY_LOCK:

        current = load_history()

        updated = pd.concat(
            [
                current,
                pd.DataFrame(
                    [row]
                ),
            ],
            ignore_index=True,
        )

        write_history(
            updated
        )

def scoped_history():

    dataframe = (
        load_history()
    )

    if (
        st.session_state[
            "role"
        ]
        == "主账号(Admin)"
    ):

        return dataframe

    return dataframe[
        dataframe[
            "operator"
        ]
        == st.session_state[
            "operator"
        ]
    ].copy()

