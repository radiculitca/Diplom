import os
import json
import re
import time
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "meta-llama/llama-3.3-70b-instruct:free")


def _extract_json(text: str) -> dict:
    try:
        return json.loads(text.strip())
    except json.JSONDecodeError:
        pass
    match = re.search(r'```(?:json)?\s*([\s\S]*?)```', text)
    if match:
        return json.loads(match.group(1).strip())
    raise ValueError("ИИ вернул ответ не в формате JSON")


def _build_group_prompt(answers: list[str], question_name: str) -> str:
    numbered = "\n".join(f"{i + 1}. {a}" for i, a in enumerate(answers))
    return (
        "Ты — эксперт по нормализации текстовых данных.\n"
        f"Вопрос анкеты: «{question_name}»\n\n"
        "Для каждого ответа из списка выполни нормализацию:\n"
        "1. Переведи на русский язык (если ответ на другом языке).\n"
        "2. Исправь явные опечатки в словах (например: «Туркмееистан» → «Туркменистан»).\n"
        "3. Приведи к именительному падежу (например: «России» → «Россия», «Из Индии» → «Индия»).\n"
        "4. Убери предлоги в начале («Из Туркменистана» → «Туркменистан»).\n"
        "5. Нормализуй аббревиатуры («РФ» → «Россия», «КНР» → «Китай»).\n"
        "6. Если ответ явно не относится к теме вопроса (имя человека, случайное слово, число) — "
        "оставь его в точности как есть.\n"
        "ВАЖНО: нормализуй каждый ответ строго по его смыслу. "
        "Не смешивай разные сущности (разные страны, разные понятия).\n\n"
        "Список ответов:\n"
        f"{numbered}\n\n"
        "Верни ТОЛЬКО валидный JSON без пояснений и без markdown:\n"
        '{"1": "нормализованный ответ", "2": "нормализованный ответ", ...}\n'
        "Ключи — номера из списка (строки), значения — нормализованные формы."
    )


def _parse_group_response(raw: str, answers: list[str]) -> list[dict]:
    parsed = _extract_json(raw)
    norm_map = {ans: parsed.get(str(i + 1), ans) for i, ans in enumerate(answers)}
    groups: dict[str, list[str]] = {}
    for ans in answers:
        groups.setdefault(norm_map[ans], []).append(ans)
    return [{"canonical": canon, "members": members} for canon, members in groups.items()]


def group_answers_openrouter(answers: list[str], question_name: str) -> list[dict]:
    key = os.getenv("OPENROUTER_API_KEY", "")
    if not key:
        raise ValueError("OPENROUTER_API_KEY не задан в .env")
    prompt = _build_group_prompt(answers, question_name)
    client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=key)
    for attempt in range(5):
        try:
            response = client.chat.completions.create(
                model=OPENROUTER_MODEL,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=2000,
                temperature=0.3,
            )
            return _parse_group_response(response.choices[0].message.content.strip(), answers)
        except Exception as e:
            if "429" in str(e) and attempt < 4:
                time.sleep(15 * (attempt + 1))
            else:
                raise
