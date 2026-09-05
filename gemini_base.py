import random
import time
from google import genai

from config import MAX_ATTEMPTS_PER_MODEL, FILE_POLL_INTERVAL_SEC, FILE_PROCESS_TIMEOUT_SEC

TRANSIENT_MARKERS = [
    "429",
    "500",
    "502",
    "503",
    "504",

    "RESOURCE_EXHAUSTED",
    "TOO_MANY_REQUESTS",
    "RATE_LIMIT",

    "INTERNAL",
    "UNAVAILABLE",
    "SERVICE_UNAVAILABLE",

    "DEADLINE_EXCEEDED",

    "HIGH DEMAND",
    "TEMPORARILY",
    "OVERLOADED",
]

MODEL_ERROR_MARKERS = [
    "MODEL_NOT_FOUND",
    "MODEL NOT FOUND",
    "NOT_FOUND",
]

def is_transient_error(
    exc,
):

    text = str(
        exc
    ).upper()

    return any(
        marker in text
        for marker
        in TRANSIENT_MARKERS
    )

def is_model_error(
    exc,
):

    text = str(
        exc
    ).upper()

    return any(
        marker in text
        for marker
        in MODEL_ERROR_MARKERS
    )

def friendly_error(
    exc,
):

    text = str(
        exc
    ).upper()

    if (
        "401" in text
        or "UNAUTHENTICATED" in text
    ):

        return (
            "Gemini API Key 认证失败，"
            "请联系管理员检查 Streamlit Secrets。"
        )

    if (
        "400" in text
        or "INVALID_ARGUMENT" in text
    ):

        return (
            "AI请求参数异常，"
            "请联系管理员检查模型调用或结构化输出配置。"
        )

    if (
        "429" in text
        or "RESOURCE_EXHAUSTED" in text
    ):

        return (
            "当前 AI 请求较多，"
            "系统已自动重试并尝试备用线路，"
            "请稍后再次执行。"
        )

    if (
        "503" in text
        or "UNAVAILABLE" in text
        or "HIGH DEMAND" in text
    ):

        return (
            "当前 AI 服务繁忙，"
            "系统已自动重试并尝试备用线路，"
            "请稍后再次执行。"
        )

    return (
        "AI暂时未完成本次任务，"
        "请稍后重新执行。"
    )

def create_client(api_key):
    return genai.Client(api_key=api_key)


def generate_resilient(client, contents, config, model_chain, max_attempts_per_model=MAX_ATTEMPTS_PER_MODEL):
    last_exception = None
    total_attempts = 0
    for model_index, model_name in enumerate(model_chain):
        for attempt_index in range(max_attempts_per_model):
            total_attempts += 1
            try:
                response = client.models.generate_content(model=model_name, contents=contents, config=config)
                return response, {
                    "model_used": model_name,
                    "fallback_used": model_index > 0,
                    "retry_count": max(total_attempts - 1, 0),
                }
            except Exception as exc:
                last_exception = exc
                if is_model_error(exc):
                    break
                if not is_transient_error(exc):
                    raise
                if attempt_index < max_attempts_per_model - 1:
                    delay = 1.3 * (2 ** attempt_index) + random.uniform(0.2, 0.8)
                    time.sleep(delay)
                    continue
                break
    raise RuntimeError(friendly_error(last_exception)) from last_exception

def wait_until_active(
    client,
    uploaded_file,
):

    started = (
        time.monotonic()
    )

    current = (
        uploaded_file
    )

    while True:

        state = getattr(
            current,
            "state",
            None,
        )

        state_name = getattr(
            state,
            "name",
            "",
        )

        if state_name == "ACTIVE":

            return current

        if state_name in {
            "FAILED",
            "ERROR",
        }:

            raise RuntimeError(
                f"视频处理失败：{state_name}"
            )

        if (
            time.monotonic()
            - started
            > FILE_PROCESS_TIMEOUT_SEC
        ):

            raise TimeoutError(
                "视频预处理超时，请压缩视频后重新上传。"
            )

        time.sleep(
            FILE_POLL_INTERVAL_SEC
        )

        current = (
            client.files.get(
                name=current.name
            )
        )
