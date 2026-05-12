import io
import os
import time
import ollama
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Cm
from dotenv import load_dotenv

load_dotenv()

OPENROUTER_API_KEY_DEFAULT = os.getenv("OPENROUTER_API_KEY", "")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "meta-llama/llama-3.3-70b-instruct:free")

# ===================== ПРИМЕР СТИЛЯ (FEW-SHOT) =====================

STYLE_EXAMPLE = """
Пример аналитического текста из университетского отчёта:

В текущем учебном году отмечается тенденция к увеличению числа студентов,
положительно оценивающих качество организации образовательного процесса.
Наиболее высокие оценки получили критерии, связанные с логической
последовательностью изучения дисциплин и актуальностью получаемых знаний.
Это свидетельствует о достаточно высоком уровне удовлетворённости
обучающихся содержанием образовательной программы. Вместе с тем часть
опрошенных указывает на необходимость расширения практической составляющей
обучения и повышения гибкости расписания. Полученные результаты позволяют
сделать вывод о сформировавшемся положительном восприятии образовательного
процесса среди студентов. Можно предположить, что сохранение данной
тенденции во многом определяется последовательной работой кафедр по
актуализации содержания учебных дисциплин.
---

Результаты опроса показали, что большинство студентов не проявляют интереса
к участию в научной, творческой и волонтёрской деятельности, демонстрируя
низкую вовлечённость во внеучебную жизнь вуза. Можно предположить, что
это связано с отсутствием предыдущего опыта участия в подобных проектах и
недостаточной информированностью о существующих возможностях. Вместе с тем
следует отметить, что значительная часть опрошенных выразила готовность к
участию при наличии более активной информационной поддержки. Это
свидетельствует о наличии потенциала для повышения вовлечённости студентов
при условии целенаправленной работы со стороны университета.
"""


# ===================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====================

def _p(doc, text, bold=False, size=12, space_before=0, space_after=3):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after = Pt(space_after)
    run = para.add_run(text)
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    return para


def _make_doc():
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(1.5)
    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(12)
    return doc


# ===================== ДАННЫЕ (файл 1) =====================

def generate_data_docx(questions: list) -> bytes:
    """Файл со списками вопросов и статистикой без аналитики."""
    doc = _make_doc()
    _last_section = None

    for q in questions:
        sec = q.get("section")
        sec_name = sec.get("name") if sec else None

        if sec_name and sec_name != _last_section:
            _last_section = sec_name
            if doc.paragraphs:
                doc.add_page_break()
            _p(doc, sec_name, bold=True, size=14, space_after=4)
            if sec.get("description"):
                _p(doc, sec["description"], size=12, space_after=8)

        file_keys = q["file_keys"]
        file_labels = q["file_labels"]
        file_totals = q["file_totals"]
        is_single = len(file_keys) == 1

        _p(doc, f"Вопрос {q['table_num']} – «{q['question_name']}»",
           bold=True, space_before=10, space_after=2)

        for row in q["rows"]:
            if is_single:
                fk = file_keys[0]
                count = row["counts"].get(fk, 0)
                total = file_totals.get(fk, 0)
                pct = f"{count / total * 100:.1f}%" if total > 0 else "—"
                _p(doc, f"  • {row['answer']}: {count} ({pct})", space_after=1)
            else:
                parts = []
                for fk in file_keys:
                    label = file_labels.get(fk, fk)
                    count = row["counts"].get(fk, 0)
                    total = file_totals.get(fk, 0)
                    pct = f"{count / total * 100:.1f}%" if total > 0 else "—"
                    parts.append(f"{label}: {count} ({pct})")
                _p(doc, f"  • {row['answer']}: {'; '.join(parts)}", space_after=1)

        if q.get("show_total", True):
            if is_single:
                fk = file_keys[0]
                _p(doc, f"  Всего: {file_totals.get(fk, 0)}", bold=True, space_after=6)
            else:
                parts = [
                    f"{file_labels.get(fk, fk)}: {file_totals.get(fk, 0)}"
                    for fk in file_keys
                ]
                _p(doc, f"  Всего: {'; '.join(parts)}", bold=True, space_after=6)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()



# ===================== ПРОМПТЫ =====================

def _build_question_prompt(
    question: dict,
    sec_name: str = "",
    sec_description: str = "",
) -> str:
    lines = []

    # Описание раздела — максимальный приоритет, идёт ПЕРВЫМ
    if sec_description:
        lines += [
            "ВАЖНЕЙШЕЕ ТРЕБОВАНИЕ К ЭТОМУ АНАЛИЗУ (высший приоритет над всеми остальными инструкциями):",
            f"{sec_description}",
            "Весь текст ОБЯЗАН соответствовать этому требованию.",
            "",
        ]

    lines += [
        "Ты — аналитик социологических исследований в российском университете.",
        "Пиши в стиле официального аналитического отчёта университета.",
        "",
        "Ниже приведён пример желаемого стиля:",
        STYLE_EXAMPLE,
        "Строго следуй этому стилю.",
        "",
        "ПРАВИЛА НАПИСАНИЯ:",
        "— Не пересказывай статистику подряд.",
        "— Избегай перечисления процентов в каждом предложении.",
        "— Главное — интерпретация результатов и формулирование выводов.",
        "— Текст должен быть похож на раздел аналитического отчёта, а не на сводку данных.",
        "— Используй академические аналитические конструкции:",
        "  «это свидетельствует о», «наблюдается тенденция»,",
        "  «полученные результаты позволяют сделать вывод»,",
        "  «вместе с тем», «следует отметить», «можно предположить».",
        "— Все слова не на русском языке переводи на русский.",
        "",
        "СТРУКТУРА АНАЛИЗА (5–9 предложений):",
        "1. Кратко опиши общую тенденцию распределения ответов.",
        "2. Выдели доминирующие и наименее популярные варианты.",
        "3. Объясни возможные причины наблюдаемого распределения.",
        "4. Сформулируй потенциальные выводы и рекомендации для университета.",
        "5. Если есть различия между группами — интерпретируй их.",
        "",
    ]

    if sec_name:
        lines += [f"Раздел анкеты: «{sec_name}»", ""]

    q = question
    lines += [
        f"Напиши полноценный фрагмент аналитического отчёта в официально-исследовательском стиле "
        f"по вопросу: «{q['question_name']}» (8–12 предложений).",
        "",
        "Статистика ответов:",
    ]
    for row in q["rows"]:
        parts = []
        for fk in q["file_keys"]:
            label = q["file_labels"].get(fk, fk)
            count = row["counts"].get(fk, 0)
            total = q["file_totals"].get(fk, 0)
            pct = round(count / total * 100, 1) if total > 0 else 0
            parts.append(f"{label}: {count} ({pct}%)")
        lines.append(f"  - {row['answer']}: {'; '.join(parts)}")

    return "\n".join(lines)


# ===================== ВЫЗОВ МОДЕЛИ =====================

def _call_ollama(prompt: str) -> str:
    response = ollama.chat(
        model="qwen2.5:7b",
        messages=[{"role": "user", "content": prompt}],
        options={
            "temperature": 0.7,
            "top_p": 0.9,
            "num_predict": 3000,
        },
    )
    return response["message"]["content"].strip()


def _call_openrouter(prompt: str, api_key: str) -> str:
    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
    )
    for attempt in range(5):
        try:
            response = client.chat.completions.create(
                model=OPENROUTER_MODEL,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=3000,
                temperature=0.7,
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            if "429" in str(e) and attempt < 4:
                wait = 15 * (attempt + 1)
                print(f"  -> Rate limit, жду {wait}с...")
                time.sleep(wait)
            else:
                raise


def _call_llm(prompt: str, provider: str = "ollama", api_key: str = "") -> str:
    if provider == "openrouter":
        key = api_key or OPENROUTER_API_KEY_DEFAULT
        if not key:
            raise ValueError("Не указан API-ключ OpenRouter")
        return _call_openrouter(prompt, key)
    return _call_ollama(prompt)


# ===================== АНАЛИТИКА (файл 2) =====================

def generate_analysis_docx(questions: list, progress_callback=None, provider: str = "ollama", api_key: str = "") -> bytes:
    """
    Аналитический файл: один вызов LLM на каждый вопрос.
    Заголовки разделов выводятся при смене раздела.
    """
    doc = _make_doc()
    total_questions = len(questions)
    _last_section = None

    for idx, q in enumerate(questions, start=1):
        sec = q.get("section")
        sec_name = sec.get("name") if sec else ""
        sec_description = sec.get("description", "") if sec else ""

        if progress_callback:
            progress_callback(idx, total_questions, q["question_name"])

        print(f"[{idx}/{total_questions}] Генерация: {q['question_name']}")

        # Заголовок раздела при смене
        if sec_name and sec_name != _last_section:
            _last_section = sec_name
            if idx > 1:
                doc.add_page_break()
            _p(doc, sec_name, bold=True, size=14, space_after=4)
            if sec_description:
                _p(doc, sec_description, size=11, space_after=6)

        # Заголовок вопроса
        _p(
            doc,
            f"Вопрос {q['table_num']} — «{q['question_name']}»",
            bold=True,
            size=12,
            space_before=8,
            space_after=3,
        )

        try:
            prompt = _build_question_prompt(q, sec_name, sec_description)
            analysis = _call_llm(prompt, provider=provider, api_key=api_key)
            _p(doc, analysis, space_after=10)
            print("  -> OK")
            if provider == "openrouter":
                time.sleep(10)
        except Exception as e:
            print(f"  -> ERROR: {e}")
            _p(doc, f"Ошибка генерации аналитики: {e}", bold=True, space_after=8)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ===================== ЕДИНАЯ ТОЧКА ВХОДА =====================

def generate_docx(questions: list, progress_callback=None, provider: str = "ollama", api_key: str = "") -> tuple[bytes, bytes]:
    data_bytes = generate_data_docx(questions)
    analysis_bytes = generate_analysis_docx(
        questions,
        progress_callback=progress_callback,
        provider=provider,
        api_key=api_key,
    )
    return data_bytes, analysis_bytes