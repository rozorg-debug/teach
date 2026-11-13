from docx import Document

from docx_parser import parse_answers, parse_questions


def save_docx(path, lines):
    document = Document()
    for line in lines:
        document.add_paragraph(line)
    document.save(path)


def test_parses_multiple_variants_and_answers(tmp_path):
    questions_path = tmp_path / "questions.docx"
    save_docx(
        questions_path,
        [
            "ВАРИАНТ 1",
            "1. [ТЕСТ] Выберите букву",
            "A) Первая",
            "B) Вторая",
            "ВАРИАНТ 2",
            "1. [ОТКРЫТЫЙ] Напишите слово",
        ],
    )
    variants = parse_questions(questions_path)
    assert [variant.code for variant in variants] == ["1", "2"]
    assert variants[0].questions[0].question_type == "test"
    assert variants[0].questions[0].options[1]["key"] == "B"
    assert variants[1].questions[0].question_type == "open"

    answers_path = tmp_path / "answers.docx"
    save_docx(answers_path, ["ВАРИАНТ 1", "1. B", "ВАРИАНТ 2", "1. слово | СЛОВО"])
    answers = parse_answers(answers_path)
    assert answers["1"][1] == ["B"]
    assert answers["2"][1] == ["слово", "СЛОВО"]
