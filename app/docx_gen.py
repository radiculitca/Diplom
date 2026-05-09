import io
import ollama
from docx import Document
from docx.shared import Pt, Cm


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
                parts = [f"{file_labels.get(fk, fk)}: {file_totals.get(fk, 0)}" for fk in file_keys]
                _p(doc, f"  Всего: {'; '.join(parts)}", bold=True, space_after=6)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ===================== АНАЛИТИКА (файл 2) =====================

def _build_section_prompt(sec_name: str, sec_description: str, questions: list) -> str:
    lines = [
        "Ты аналитик социологических исследований. Пиши официальным аналитическим стилем.",
        "",
    ]

    if sec_description:
        lines += [
            f"Раздел анкеты: «{sec_name}»",
            f"Указания по анализу этого раздела: {sec_description}",
            "",
            "Строго следуй указаниям выше при написании аналитики по каждому вопросу раздела.",
            "",
        ]
    else:
        lines += [
            f"Раздел анкеты: «{sec_name}»",
            "",
        ]

    lines += [
        "Ниже приведены вопросы раздела и статистика ответов на них.",
        "Сделай аналитический вывод по каждому вопросу (4–6 предложений):",
        "— отметь тенденции и распределение ответов;",
        "— выдели наиболее и наименее популярные варианты;",
        "— если вопрос открытый — выдели уникальные и частые темы;",
        "— если несколько файлов — сравни группы между собой.",
        "Не перечисляй цифры подряд — именно анализируй.",
        "Все слова не на русском языке переводи на русский.",
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
        "После анализа каждого вопроса сделай общий вывод по всему разделу (3–5 предложений): ",
        "выдели главные инсайты и тенденции раздела.",
    ]

    return "\n".join(lines)


def _build_no_section_prompt(questions: list) -> str:
    """Промпт для вопросов без раздела."""
    lines = [
        "Ты аналитик социологических исследований. Пиши официальным аналитическим стилем.",
        "",
        "Ниже приведены вопросы анкеты и статистика ответов.",
        "Сделай аналитический вывод по каждому вопросу (4–6 предложений).",
        "Не перечисляй цифры подряд — именно анализируй.",
        "Все слова не на русском языке переводи на русский.",
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


def _call_ollama(prompt: str) -> str:
    response = ollama.chat(
        model="mistral",
        messages=[{"role": "user", "content": prompt}]
    )
    return response["message"]["content"].strip()


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
                    qs
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
    data_bytes    — файл со статистикой
    analysis_bytes — файл с аналитикой по разделам
    """
    data_bytes = generate_data_docx(questions)
    analysis_bytes = generate_analysis_docx(questions, progress_callback=progress_callback)
    return data_bytes, analysis_bytes