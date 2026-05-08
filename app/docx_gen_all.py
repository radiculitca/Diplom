import io
import ollama

from docx import Document
from docx.shared import Pt, Cm


# ======================================
# WORD HELPER
# ======================================

def _p(doc, text, bold=False, size=12,
       space_before=0, space_after=3):

    para = doc.add_paragraph()

    para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after = Pt(space_after)

    run = para.add_run(text)

    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)

    return para


# ======================================
# СОЗДАНИЕ ОБЩЕГО PROMPT
# ======================================

def build_global_prompt(questions):

    text = """
Ты аналитик социологических исследований.

Ниже представлены результаты анкеты: вопросы и статистика по ним.

Сделай аналитиический вывод по каждому вопросу 3-5 предложений, отметь тендеции, сравни ответы между собой, выдели самые популярные.
Не превращай вывод в перечисление ответов в том или ином порядке, а именно проанализируй их.
Учитывай пердыдущие вопросы если они имеют общую тему.

В коце сделай общий вывод по всему исследованию, выдели главные инсайты и тенденции.

Пиши официальным аналитическим стилем.
Не повторяй все цифры подряд.
"""

    for q in questions:

        

        sec = q.get("section")
        if sec and sec.get("name"):
            section_hint = f" [Раздел: {sec['name']}"
            if sec.get("description"):
                section_hint += f" — {sec['description']}"
            section_hint += "]"
        else:
            section_hint = ""

        text += (
            f"\n\n"
            f"Вопрос {q['table_num']}"
            f"{section_hint}"
            f" — {q['question_name']}\n"
        )

        for row in q["rows"]:

            answer = row["answer"]

            parts = []

            for fk in q["file_keys"]:

                label = q["file_labels"].get(fk, fk)

                count = row["counts"].get(fk, 0)

                total = q["file_totals"].get(fk, 0)

                pct = (
                    round(count / total * 100, 1)
                    if total > 0 else 0
                )

                parts.append(
                    f"{label}: {count} ({pct}%)"
                )

            text += (
                f"- {answer}: "
                f"{'; '.join(parts)}\n"
            )

    return text


# ======================================
# AI ГЕНЕРАЦИЯ
# ======================================

def generate_global_analysis(questions):

    print("Генерация общего аналитического отчета...")

    prompt = build_global_prompt(questions)

    response = ollama.chat(
        model="mistral",
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ]
    )

    print("Аналитика успешно сгенерирована")

    return response["message"]["content"]


# ======================================
# DOCX EXPORT
# ======================================

def generate_docx(questions: list) -> bytes:

    doc = Document()

    _last_section_name = None

    for section in doc.sections:

        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(1.5)

    normal = doc.styles["Normal"]

    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)

    # ======================================
    # ГЕНЕРАЦИЯ ОБЩЕЙ АНАЛИТИКИ
    # ======================================

    try:

        analysis = generate_global_analysis(questions)

        _p(
            doc,
            "Аналитический отчет",
            bold=True,
            size=14,
            space_after=6
        )

        _p(
            doc,
            analysis,
            space_after=10
        )

    except Exception as e:

        print("Ошибка генерации аналитики:", e)

        _p(
            doc,
            "Ошибка генерации аналитического отчета.",
            bold=True,
            space_after=10
        )

    # ======================================
    # ВЫВОД ВОПРОСОВ
    # ======================================

    for q in questions:

        file_keys = q["file_keys"]
        file_labels = q["file_labels"]
        file_totals = q["file_totals"]

        is_single = len(file_keys) == 1

    # Заголовок раздела (вставляется один раз перед первым вопросом раздела)
    sec = q.get("section")
    sec_name = sec.get("name") if sec else None
    if sec_name and sec_name != _last_section_name:
        _last_section_name = sec_name
        doc.add_page_break()
        _p(
            doc,
            sec_name,
            bold=True,
            size=14,
            space_before=0,
            space_after=4
        )
        if sec.get("description"):
            _p(
                doc,
                sec.get("description"),
                bold=False,
                size=12,
                space_before=0,
                space_after=8
            )


        # Заголовок вопроса

        _p(
            doc,
            f"Вопрос {q['table_num']} – "
            f"«{q['question_name']}»",
            bold=True,
            space_before=10,
            space_after=2
        )

        # Ответы

        for row in q["rows"]:

            if is_single:

                fk = file_keys[0]

                count = row["counts"].get(fk, 0)

                total = file_totals.get(fk, 0)

                pct = (
                    f"{count / total * 100:.1f}%"
                    if total > 0 else "—"
                )

                _p(
                    doc,
                    f"• {row['answer']}: "
                    f"{count} ({pct})",
                    space_after=1
                )

            else:

                parts = []

                for fk in file_keys:

                    label = file_labels.get(fk, fk)

                    count = row["counts"].get(fk, 0)

                    total = file_totals.get(fk, 0)

                    pct = (
                        f"{count / total * 100:.1f}%"
                        if total > 0 else "—"
                    )

                    parts.append(
                        f"{label}: {count} ({pct})"
                    )

                _p(
                    doc,
                    f"• {row['answer']}: "
                    f"{'; '.join(parts)}",
                    space_after=1
                )

        # Итого

        if q.get("show_total", True):

            if is_single:

                fk = file_keys[0]

                _p(
                    doc,
                    f"Всего: "
                    f"{file_totals.get(fk, 0)}",
                    bold=True,
                    space_after=6
                )

            else:

                parts = [
                    f"{file_labels.get(fk, fk)}: "
                    f"{file_totals.get(fk, 0)}"
                    for fk in file_keys
                ]

                _p(
                    doc,
                    f"Всего: "
                    f"{'; '.join(parts)}",
                    bold=True,
                    space_after=6
                )

    # ======================================
    # SAVE
    # ======================================

    buf = io.BytesIO()

    doc.save(buf)

    buf.seek(0)

    return buf.getvalue()