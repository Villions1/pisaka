"""
ai_client — единый клиент для OCR конспектов и определения порядка страниц.

Поддерживает два бэкенда (выбирается через .env):
  BACKEND=gemini      — Google Gemini API (по умолчанию)
  BACKEND=openrouter  — OpenRouter API (OpenAI-совместимый)

Переменные .env:
  BACKEND          — gemini | openrouter
  PROXY            — прокси, например http://127.0.0.1:2080

  # Gemini
  GEMINI_API_KEY   — ключ Google AI Studio
  GEMINI_MODEL     — модель (по умолчанию gemini-flash-latest)

  # OpenRouter
  OPENROUTER_API_KEY — ключ openrouter.ai
  OPENROUTER_MODEL   — модель (по умолчанию nvidia/nemotron-nano-12b-2-vl:free)

  # Общее
  FONT_PATH        — путь к TTF-шрифту (используется notes_to_docx.py)
  MAX_RETRIES      — попыток при 429/5xx (по умолчанию 5)
"""

from __future__ import annotations

import base64
import json
import mimetypes
import os
import re
import time
from pathlib import Path
from typing import Any

import requests
from dotenv import load_dotenv


# ---------------------------------------------------------------------------
# Промты (одинаковы для обоих бэкендов)
# ---------------------------------------------------------------------------

OCR_SYSTEM_PROMPT = (
    "Ты распознаёшь рукописные конспекты студента по фото. "
    "Твоя задача — выдать ПОЛНЫЙ текст с фотографии и НИЧЕГО, кроме него. "
    "Строгие правила:\n"
    "1. Только обычный связный текст. Никакого Markdown (НЕ используй **, __, #), "
    "никаких заголовков, списков, нумерации, символов *, #, >, кода.\n"
    "2. Формулы НЕ выписывай. Если в конспекте встречается формула, пропусти её "
    "или замени словесным описанием в одну фразу (без специальных символов). "
    "Никаких LaTeX, никаких знаков равенства, интегралов, сумм, стрелок, "
    "греческих букв из формул.\n"
    "3. Схемы, рисунки, таблицы НЕ описывай и не воспроизводи.\n"
    "4. Если отдельные слова или буквы неразборчивы — додумай их по смыслу и "
    "контексту, не отмечая это никак. Никаких [нрзб], скобок с вариантами, "
    "троеточий на месте пропуска.\n"
    "5. ТИПОГРАФИКА (критично): использовать ТОЛЬКО следующие знаки препинания "
    "и символы: точка, запятая, точка с запятой, двоеточие, восклицательный и "
    "вопросительный знаки, обычный дефис-минус, круглые скобки, прямые "
    'ASCII-кавычки ". ЗАПРЕЩЕНЫ: длинное тире, среднее тире, любые кривые/'
    "типографские/ёлочные/немецкие кавычки (никаких символов "
    "U+2013, U+2014, U+2018, U+2019, U+201C, U+201D, U+201E, U+00AB, U+00BB, "
    "U+2026 и подобных). Вместо тире — дефис с пробелами. Вместо "
    "троеточия — три обычные точки. Вместо кавычек — прямые ASCII-кавычки.\n"
    "6. Сохраняй язык оригинала. Сохраняй абзацы переводами строк.\n"
    "7. Никаких вступлений и комментариев. Сразу с первой строки — текст."
)

OCR_USER_PROMPT = (
    "Распознай рукописный конспект на этом фото согласно правилам. "
    "Выдай только текст."
)

ORDER_SYSTEM_PROMPT = (
    "Ты — эксперт по восстановлению порядка страниц рукописного конспекта. "
    "Тебе дают пронумерованные фрагменты текста, распознанные с фотографий "
    "страниц одной тетради, в произвольном порядке. "
    "Твоя задача — определить правильный порядок страниц.\n\n"
    "На что опираешься (по убыванию приоритета):\n"
    "1. Явная нумерация страниц / параграфов / пунктов в тексте.\n"
    "2. Логические связки: фраза обрывается на одной странице и продолжается "
    "на другой.\n"
    "3. Логика изложения: вступление -> определения -> примеры -> выводы.\n\n"
    "ОТВЕТ СТРОГО в формате JSON-массива целых чисел — индексы фрагментов "
    "от первой страницы к последней. Пример: [3,0,2,1]. "
    "Ничего кроме JSON-массива, никаких пояснений, без Markdown-обёртки."
)

# Замены «запрещённых» Unicode-символов на ASCII
_TYPOGRAPHY_FIXES = {
    "—": "-", "–": "-", "−": "-", "‐": "-", "‑": "-", "‒": "-", "―": "-",
    "\u201c": '"', "\u201d": '"', "\u201e": '"', "\u201f": '"',
    "\u00ab": '"', "\u00bb": '"', "\u2039": '"', "\u203a": '"',
    "\u2018": "'", "\u2019": "'", "\u201a": "'", "\u201b": "'",
    "\u2026": "...",
    "\u00a0": " ", "\u202f": " ", "\u2009": " ",
}


def sanitize_for_font(text: str) -> str:
    if not text:
        return text
    # Удаляем лишние пробелы в начале/конце строк и нормализуем whitespace
    lines = []
    for line in text.splitlines():
        # Заменяем все виды пробелов на обычный пробел
        line = re.sub(r"[\s\u00a0\u2000-\u200b]+", " ", line).strip()
        if line:
            lines.append(line.translate(str.maketrans(_TYPOGRAPHY_FIXES)))
        else:
            lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Кэширование
# ---------------------------------------------------------------------------

def get_cache_path(out_dir: Path) -> Path:
    return out_dir / "ocr_cache.json"


def load_cache(out_dir: Path) -> dict[str, str]:
    path = get_cache_path(out_dir)
    if not path.exists():
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_cache(out_dir: Path, cache: dict[str, str]) -> None:
    path = get_cache_path(out_dir)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Конфиг
# ---------------------------------------------------------------------------

DEFAULT_GEMINI_MODEL     = "gemini-flash-latest"
DEFAULT_OR_MODEL         = "nvidia/nemotron-nano-12b-2-vl:free"
GEMINI_URL_TPL = (
    "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
)
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"


def load_config() -> dict[str, Any]:
    here = Path(__file__).resolve().parent
    env_path = here / ".env"
    if env_path.exists():
        load_dotenv(env_path, override=False)

    backend = os.getenv("BACKEND", "gemini").strip().lower()
    proxy   = os.getenv("PROXY", "").strip() or None

    gemini_key   = os.getenv("GEMINI_API_KEY", "").strip()
    gemini_model = os.getenv("GEMINI_MODEL", DEFAULT_GEMINI_MODEL).strip() or DEFAULT_GEMINI_MODEL

    or_key   = os.getenv("OPENROUTER_API_KEY", "").strip()
    or_model = os.getenv("OPENROUTER_MODEL", DEFAULT_OR_MODEL).strip() or DEFAULT_OR_MODEL

    font_path   = os.getenv("FONT_PATH", "").strip()
    max_retries = int(os.getenv("MAX_RETRIES", "5"))

    if backend == "gemini" and not gemini_key:
        raise RuntimeError("GEMINI_API_KEY не задан в .env")
    if backend == "openrouter" and not or_key:
        raise RuntimeError("OPENROUTER_API_KEY не задан в .env")

    return {
        "backend":      backend,
        "proxy":        proxy,
        "gemini_key":   gemini_key,
        "gemini_model": gemini_model,
        "or_key":       or_key,
        "or_model":     or_model,
        "font_path":    font_path,
        "max_retries":  max_retries,
    }


def backend_label(cfg: dict) -> str:
    if cfg["backend"] == "openrouter":
        return f"OpenRouter / {cfg['or_model']}"
    return f"Gemini / {cfg['gemini_model']}"


# ---------------------------------------------------------------------------
# HTTP с retry на 429 / 5xx
# ---------------------------------------------------------------------------

def _proxies(proxy_url: str | None) -> dict[str, str] | None:
    if not proxy_url:
        return None
    return {"http": proxy_url, "https": proxy_url}


def _retry_after(resp: requests.Response) -> float:
    """Извлекает секунды ожидания из заголовка Retry-After или тела 429."""
    # Заголовок Retry-After
    ra = resp.headers.get("Retry-After", "")
    if ra:
        try:
            return max(1.0, float(ra))
        except ValueError:
            pass
    # Тело: "Please retry in 53.1s" / "retry in 53.116336386s"
    try:
        body = resp.text
        m = re.search(r"retry[^0-9]*(\d+(?:\.\d+)?)\s*s", body, re.I)
        if m:
            return max(1.0, float(m.group(1)))
    except Exception:
        pass
    return 60.0  # fallback


def _post_with_retry(
    url: str,
    headers: dict,
    body: dict,
    proxies: dict | None,
    max_retries: int,
    timeout: int = 180,
) -> dict:
    attempt = 0
    while True:
        attempt += 1
        try:
            resp = requests.post(
                url, headers=headers,
                data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
                proxies=proxies,
                timeout=timeout,
            )
        except requests.exceptions.RequestException as e:
            if attempt >= max_retries:
                raise RuntimeError(f"Сетевая ошибка после {attempt} попыток: {e}") from e
            wait = 5.0 * attempt
            print(f"\r  [сеть] попытка {attempt}/{max_retries}, жду {wait:.0f}с... ", end="", flush=True)
            time.sleep(wait)
            continue

        if resp.status_code == 200:
            return resp.json()

        if resp.status_code in (429, 500, 502, 503, 504) and attempt < max_retries:
            wait = _retry_after(resp) if resp.status_code == 429 else 10.0 * attempt
            print(
                f"\r  [HTTP {resp.status_code}] попытка {attempt}/{max_retries}, "
                f"жду {wait:.0f}с... ",
                end="", flush=True,
            )
            time.sleep(wait)
            continue

        raise RuntimeError(f"API HTTP {resp.status_code}: {resp.text[:600]}")


# ---------------------------------------------------------------------------
# Gemini backend
# ---------------------------------------------------------------------------

def _guess_mime(path: Path) -> str:
    mime, _ = mimetypes.guess_type(str(path))
    return mime if (mime and mime.startswith("image/")) else "image/jpeg"


def _gemini_ocr(image_path: Path, cfg: dict) -> str:
    data_b64 = base64.b64encode(Path(image_path).read_bytes()).decode("ascii")
    body = {
        "systemInstruction": {"parts": [{"text": OCR_SYSTEM_PROMPT}]},
        "contents": [{
            "role": "user",
            "parts": [
                {"inlineData": {"mimeType": _guess_mime(image_path), "data": data_b64}},
                {"text": OCR_USER_PROMPT},
            ],
        }],
        "generationConfig": {"temperature": 0.2},
    }
    url = GEMINI_URL_TPL.format(model=cfg["gemini_model"])
    headers = {
        "Content-Type": "application/json",
        "X-goog-api-key": cfg["gemini_key"],
    }
    resp = _post_with_retry(url, headers, body, _proxies(cfg["proxy"]), cfg["max_retries"])
    return _extract_gemini_text(resp)


def _gemini_order(texts: list[str], cfg: dict) -> list[int]:
    user_text = _build_order_user_text(texts)
    body = {
        "systemInstruction": {"parts": [{"text": ORDER_SYSTEM_PROMPT}]},
        "contents": [{"role": "user", "parts": [{"text": user_text}]}],
        "generationConfig": {
            "temperature": 0.0,
            "responseMimeType": "application/json",
            "maxOutputTokens": 2048,
        },
    }
    url = GEMINI_URL_TPL.format(model=cfg["gemini_model"])
    headers = {
        "Content-Type": "application/json",
        "X-goog-api-key": cfg["gemini_key"],
    }
    resp = _post_with_retry(url, headers, body, _proxies(cfg["proxy"]), cfg["max_retries"])
    return _parse_order_json(_extract_gemini_text(resp), len(texts))


def _extract_gemini_text(data: dict) -> str:
    candidates = data.get("candidates") or []
    if not candidates:
        raise RuntimeError(f"Gemini: пустой ответ: {json.dumps(data)[:400]}")
    parts = (candidates[0].get("content") or {}).get("parts") or []
    text = "".join(p.get("text", "") for p in parts if isinstance(p, dict)).strip()
    if not text:
        raise RuntimeError(f"Gemini: ответ без текста: {json.dumps(data)[:400]}")
    return text


# ---------------------------------------------------------------------------
# OpenRouter backend
# ---------------------------------------------------------------------------

def _or_ocr(image_path: Path, cfg: dict) -> str:
    data_b64 = base64.b64encode(Path(image_path).read_bytes()).decode("ascii")
    mime = _guess_mime(image_path)
    body = {
        "model": cfg["or_model"],
        "messages": [
            {"role": "system", "content": OCR_SYSTEM_PROMPT},
            {"role": "user", "content": [
                {"type": "image_url",
                 "image_url": {"url": f"data:{mime};base64,{data_b64}"}},
                {"type": "text", "text": OCR_USER_PROMPT},
            ]},
        ],
        "temperature": 0.2,
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {cfg['or_key']}",
        "HTTP-Referer": "https://github.com/notes-to-docx",
        "X-Title": "notes-to-docx",
    }
    resp = _post_with_retry(OPENROUTER_URL, headers, body, _proxies(cfg["proxy"]), cfg["max_retries"])
    return _extract_or_text(resp)


def _or_order(texts: list[str], cfg: dict) -> list[int]:
    user_text = _build_order_user_text(texts)
    body = {
        "model": cfg["or_model"],
        "messages": [
            {"role": "system", "content": ORDER_SYSTEM_PROMPT},
            {"role": "user",   "content": user_text},
        ],
        "temperature": 0.0,
        "response_format": {"type": "json_object"},
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {cfg['or_key']}",
        "HTTP-Referer": "https://github.com/notes-to-docx",
        "X-Title": "notes-to-docx",
    }
    resp = _post_with_retry(OPENROUTER_URL, headers, body, _proxies(cfg["proxy"]), cfg["max_retries"])
    raw = _extract_or_text(resp)
    # OpenRouter с json_object может вернуть {"order": [0,1,2]} или просто [0,1,2]
    try:
        parsed = json.loads(raw)
        if isinstance(parsed, list):
            raw_list = parsed
        elif isinstance(parsed, dict):
            # ищем первый ключ со списком
            raw_list = next(
                (v for v in parsed.values() if isinstance(v, list)), None
            )
            if raw_list is None:
                raise ValueError("нет массива в ответе")
        else:
            raise ValueError(f"неожиданный тип: {type(parsed)}")
        return _validate_order(raw_list, len(texts))
    except Exception:
        pass
    return _parse_order_json(raw, len(texts))


def _extract_or_text(data: dict) -> str:
    try:
        text = data["choices"][0]["message"]["content"].strip()
        if not text:
            raise ValueError("пустой content")
        return text
    except (KeyError, IndexError, ValueError) as e:
        raise RuntimeError(f"OpenRouter: неожиданный ответ: {json.dumps(data)[:400]}") from e


# ---------------------------------------------------------------------------
# Общие утилиты
# ---------------------------------------------------------------------------

def _build_order_user_text(texts: list[str]) -> str:
    n = len(texts)
    fragments = []
    for i, t in enumerate(texts):
        snippet = t.strip()
        if len(snippet) > 4000:
            snippet = snippet[:2000] + "\n[...]\n" + snippet[-1500:]
        fragments.append(f"--- ФРАГМЕНТ #{i} ---\n{snippet}")
    return (
        f"Всего фрагментов: {n}. Индексы от 0 до {n - 1}. "
        f"Определи правильный порядок и верни JSON-массив индексов длиной {n}.\n\n"
        + "\n\n".join(fragments)
    )


def _validate_order(order: list, n: int) -> list[int]:
    if not isinstance(order, list) or not all(isinstance(x, int) for x in order):
        raise RuntimeError(f"Не список целых: {order!r}")
    if sorted(order) != list(range(n)):
        raise RuntimeError(f"Неверная перестановка 0..{n-1}: {order!r}")
    return order


def _parse_order_json(raw: str, n: int) -> list[int]:
    s = (raw or "").strip()
    s = re.sub(r"^```(?:json)?\s*", "", s)
    s = re.sub(r"\s*```\s*$", "", s)

    candidates = [s]
    lo, hi = s.find("["), s.rfind("]")
    if lo != -1 and hi > lo:
        candidates.append(s[lo:hi + 1])
    digits = re.findall(r"-?\d+", s)
    if digits:
        candidates.append("[" + ",".join(digits) + "]")

    for c in candidates:
        try:
            parsed = json.loads(c)
        except json.JSONDecodeError:
            continue
        if isinstance(parsed, list) and all(isinstance(x, int) for x in parsed):
            return _validate_order(parsed, n)

    raise RuntimeError(f"Не смог распарсить порядок страниц: {raw!r}")


# ---------------------------------------------------------------------------
# Публичный API
# ---------------------------------------------------------------------------

def recognize_notes(image_path: Path, cfg: dict | None = None) -> str:
    """OCR одной страницы конспекта. Возвращает чистый текст."""
    cfg = cfg or load_config()
    if cfg["backend"] == "openrouter":
        text = _or_ocr(image_path, cfg)
    else:
        text = _gemini_ocr(image_path, cfg)
    return sanitize_for_font(text)


def order_pages(texts: list[str], cfg: dict | None = None) -> list[int]:
    """По списку текстов возвращает индексы в правильном порядке (0-based)."""
    cfg = cfg or load_config()
    if len(texts) <= 1:
        return list(range(len(texts)))
    if cfg["backend"] == "openrouter":
        return _or_order(texts, cfg)
    return _gemini_order(texts, cfg)
