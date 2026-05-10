import io
import ollama
from docx import Document
from docx.shared import Pt, Cm


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
Привлечение иностранных студентов является ключевой стратегической задачей,
направленной на повышение позиции университета в международных рейтинговых
системах. Большинство респондентов первоначально выбирали регион или город
для обучения, а затем уже университет. Следует отметить, что в 2024 году
приоритетным критерием выбора образовательного учреждения являлся
университет, тогда как в 2025 году на первый план вышли факторы, связанные
с регионом или городом размещения учебного заведения. Наблюдается тенденция
к концентрации предпочтений абитуриентов вокруг территориальной доступности
вуза. Полученные результаты позволяют сделать вывод о необходимости усиления
работы по продвижению университета за пределами региона.
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

def _build_section_prompt(sec_name: str, sec_description: str, questions: list) -> str:
    lines = [
        "Ты — аналитик социологических исследований в российском университете.",
        "Пиши в стиле официального аналитического отчёта университета.",
        "",
        "Ниже приведён пример желаемого стиля:",
        STYLE_EXAMPLE,
        "Строго следуй этому стилю при написании аналитики.",
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
        "СТРУКТУРА АНАЛИЗА КАЖДОГО ВОПРОСА (8–12 предложений):",
        "1. Кратко опиши общую тенденцию распределения ответов.",
        "2. Выдели доминирующие и наименее популярные варианты.",
        "3. Объясни возможные причины наблюдаемого распределения.",
        "4. Сформулируй потенциальные выводы и рекомендации для университета.",
        "5. Если есть различия между группами — интерпретируй их.",
        "",
    ]

    if sec_description:
        lines += [
            f"Раздел анкеты: «{sec_name}»",
            f"Указания по анализу этого раздела: {sec_description}",
            "",
            "Строго следуй указаниям выше при написании аналитики по каждому вопросу.",
            "",
        ]
    else:
        lines += [
            f"Раздел анкеты: «{sec_name}»",
            "",
        ]

    lines += [
        "Ниже приведены вопросы раздела и статистика ответов.",
        "Напиши полноценный фрагмент аналитического отчёта в официально-исследовательском стиле",
        "по каждому вопросу (8–12 предложений на вопрос).",
        "",
    ]

    for q in questions:
        lines.append(f"Вопрос {q['table_num']} — {q['question_name']}")
        for row in q["rows"]:
            parts = []
            for fk in q["file_keys"]:
                label = q["file_labels"].get(fk, fk)
                count = row["counts"].get(fk, 0)
                total = q["file_totals"].get(fk, 0)
                pct = round(count / total * 100, 1) if total > 0 else 0
                parts.append(f"{label}: {count} ({pct}%)")
            lines.append(f"  - {row['answer']}: {'; '.join(parts)}")
        lines.append("")

    lines += [
        "После анализа каждого вопроса напиши общий аналитический фрагмент по всему разделу",
        "(5–7 предложений): выдели главные инсайты, тенденции и рекомендации.",
    ]

    return "\n".join(lines)


def _build_no_section_prompt(questions: list) -> str:
    """Промпт для вопросов без раздела."""
    lines = [
        "Ты — аналитик социологических исследований в российском университете.",
        "Пиши в стиле официального аналитического отчёта университета.",
        "",
        "Ниже приведён пример желаемого стиля:",
        STYLE_EXAMPLE,
        "Строго следуй этому стилю при написании аналитики.",
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
        "СТРУКТУРА АНАЛИЗА КАЖДОГО ВОПРОСА (8–12 предложений):",
        "1. Кратко опиши общую тенденцию распределения ответов.",
        "2. Выдели доминирующие и наименее популярные варианты.",
        "3. Объясни возможные причины наблюдаемого распределения.",
        "4. Сформулируй потенциальные выводы и рекомендации для университета.",
        "5. Если есть различия между группами — интерпретируй их.",
        "",
        "Ниже приведены вопросы анкеты и статистика ответов.",
        "Напиши полноценный фрагмент аналитического отчёта в официально-исследовательском стиле",
        "по каждому вопросу (8–12 предложений на вопрос).",
        "",
    ]
    for q in questions:
        lines.append(f"Вопрос {q['table_num']} — {q['question_name']}")
        for row in q["rows"]:
            parts = []
            for fk in q["file_keys"]:
                label = q["file_labels"].get(fk, fk)
                count = row["counts"].get(fk, 0)
                total = q["file_totals"].get(fk, 0)
                pct = round(count / total * 100, 1) if total > 0 else 0
                parts.append(f"{label}: {count} ({pct}%)")
            lines.append(f"  - {row['answer']}: {'; '.join(parts)}")
        lines.append("")
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


# ===================== АНАЛИТИКА (файл 2) =====================

def generate_analysis_docx(questions: list, progress_callback=None) -> bytes:
    """
    Аналитический файл: группирует вопросы по разделам,
    на каждый раздел — один вызов LLM.
    """
    doc = _make_doc()

    # Группируем вопросы по разделам, сохраняя порядок
    sections_order = []   # [(sec_name, sec_obj_or_None), ...]
    sections_map = {}     # sec_name -> [questions]
    NO_SECTION = "__NO_SECTION__"

    for q in questions:
        sec = q.get("section")
        sec_name = sec.get("name") if sec else NO_SECTION
        if sec_name not in sections_map:
            sections_map[sec_name] = []
            sections_order.append((sec_name, sec))
        sections_map[sec_name].append(q)

    total_sections = len(sections_order)

    for idx, (sec_name, sec_obj) in enumerate(sections_order, start=1):
        qs = sections_map[sec_name]
        is_no_section = sec_name == NO_SECTION
        display_name = "Вопросы без раздела" if is_no_section else sec_name

        if progress_callback:
            progress_callback(idx, total_sections, display_name)

        print(f"[{idx}/{total_sections}] Генерация аналитики: {display_name}")

        if idx > 1:
            doc.add_page_break()

        _p(doc, display_name, bold=True, size=14, space_after=4)
        if not is_no_section and sec_obj and sec_obj.get("description"):
            _p(doc, sec_obj["description"], size=11, space_after=6)

        try:
            if is_no_section:
                prompt = _build_no_section_prompt(qs)
            else:
                prompt = _build_section_prompt(
                    sec_name,
                    sec_obj.get("description", "") if sec_obj else "",
                    qs,
                )

            analysis = _call_ollama(prompt)
            _p(doc, analysis, space_after=8)
            print("  -> OK")

        except Exception as e:
            print(f"  -> ERROR: {e}")
            _p(doc, f"Ошибка генерации аналитики: {e}", bold=True, space_after=8)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ===================== ЕДИНАЯ ТОЧКА ВХОДА =====================

def generate_docx(questions: list, progress_callback=None) -> tuple[bytes, bytes]:
    """
    Возвращает (data_bytes, analysis_bytes).
    data_bytes     — файл со статистикой
    analysis_bytes — файл с аналитикой по разделам
    """
    data_bytes = generate_data_docx(questions)
    analysis_bytes = generate_analysis_docx(questions, progress_callback=progress_callback)
    return data_bytes, analysis_bytes