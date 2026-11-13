import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List

from docx import Document


VARIANT_RE = re.compile(r"^\s*(?:вариант|variant)\s*№?\s*([\w.-]+)\s*[:.-]?\s*$", re.I)
QUESTION_RE = re.compile(r"^\s*(\d+)\s*[.)]\s*(.+?)\s*$")
OPTION_RE = re.compile(r"^\s*([A-HА-Ж])\s*[.)]\s*(.+?)\s*$", re.I)
ANSWER_RE = re.compile(r"^\s*(\d+)\s*[.)-]\s*(.+?)\s*$")
TEST_TAG_RE = re.compile(r"^\s*\[(?:тест|test)\]\s*", re.I)
OPEN_TAG_RE = re.compile(r"^\s*\[(?:открытый|open)\]\s*", re.I)


class DocumentFormatError(ValueError):
    """Документ не соответствует поддерживаемому формату."""


@dataclass
class ParsedQuestion:
    number: int
    text: str
    question_type: str = "open"
    options: List[dict] = field(default_factory=list)


@dataclass
class ParsedVariant:
    code: str
    questions: List[ParsedQuestion] = field(default_factory=list)


def _document_lines(path: str | Path) -> List[str]:
    document = Document(str(path))
    lines: List[str] = []

    for paragraph in document.paragraphs:
        text = " ".join(paragraph.text.replace("\xa0", " ").split())
        if text:
            lines.append(text)

    # Поддерживаем также документы, где данные оформлены таблицей.
    for table in document.tables:
        for row in table.rows:
            cells = [" ".join(cell.text.replace("\xa0", " ").split()) for cell in row.cells]
            cells = [cell for cell in cells if cell]
            if cells:
                lines.append(". ".join(cells))

    return lines


def parse_questions(path: str | Path) -> List[ParsedVariant]:
    lines = _document_lines(path)
    variants: List[ParsedVariant] = []
    current_variant: ParsedVariant | None = None
    current_question: ParsedQuestion | None = None

    for line in lines:
        variant_match = VARIANT_RE.match(line)
        if variant_match:
            current_variant = ParsedVariant(code=variant_match.group(1))
            variants.append(current_variant)
            current_question = None
            continue

        question_match = QUESTION_RE.match(line)
        if question_match:
            if current_variant is None:
                # Документ без заголовка варианта считается вариантом 1.
                current_variant = ParsedVariant(code="1")
                variants.append(current_variant)

            raw_text = question_match.group(2)
            question_type = "open"
            if TEST_TAG_RE.match(raw_text):
                question_type = "test"
                raw_text = TEST_TAG_RE.sub("", raw_text).strip()
            elif OPEN_TAG_RE.match(raw_text):
                raw_text = OPEN_TAG_RE.sub("", raw_text).strip()

            current_question = ParsedQuestion(
                number=int(question_match.group(1)),
                text=raw_text,
                question_type=question_type,
            )
            current_variant.questions.append(current_question)
            continue

        option_match = OPTION_RE.match(line)
        if option_match and current_question is not None:
            current_question.question_type = "test"
            current_question.options.append(
                {"key": option_match.group(1).upper(), "text": option_match.group(2)}
            )
            continue

        # Продолжение длинного текста вопроса.
        if current_question is not None:
            current_question.text = f"{current_question.text} {line}".strip()

    variants = [variant for variant in variants if variant.questions]
    if not variants:
        raise DocumentFormatError("В документе не найдено ни одного вопроса.")

    for variant in variants:
        numbers = [question.number for question in variant.questions]
        if len(numbers) != len(set(numbers)):
            raise DocumentFormatError(f"В варианте {variant.code} повторяются номера вопросов.")
        for question in variant.questions:
            if question.question_type == "test" and len(question.options) < 2:
                raise DocumentFormatError(
                    f"Вариант {variant.code}, вопрос {question.number}: для теста нужно минимум 2 ответа."
                )

    return variants


def parse_answers(path: str | Path) -> Dict[str, Dict[int, List[str]]]:
    lines = _document_lines(path)
    answers: Dict[str, Dict[int, List[str]]] = {}
    current_variant = "1"

    for line in lines:
        variant_match = VARIANT_RE.match(line)
        if variant_match:
            current_variant = variant_match.group(1)
            answers.setdefault(current_variant, {})
            continue

        answer_match = ANSWER_RE.match(line)
        if not answer_match:
            continue

        answers.setdefault(current_variant, {})
        number = int(answer_match.group(1))
        raw_answer = answer_match.group(2).strip()
        accepted = [item.strip() for item in raw_answer.split("|") if item.strip()]
        if accepted:
            answers[current_variant][number] = accepted

    if not answers or not any(variant_answers for variant_answers in answers.values()):
        raise DocumentFormatError("В документе не найдено ни одного ответа.")

    return answers
