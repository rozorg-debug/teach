import io

import pytest
from docx import Document

from app import Attempt, create_app, db


def make_docx(lines):
    document = Document()
    for line in lines:
        document.add_paragraph(line)
    stream = io.BytesIO()
    document.save(stream)
    stream.seek(0)
    return stream


@pytest.fixture()
def app(tmp_path):
    database = tmp_path / "test.db"
    app = create_app(
        {
            "TESTING": True,
            "SECRET_KEY": "test-key",
            "TEACHER_PASSWORD": "teacher-test-password",
            "SQLALCHEMY_DATABASE_URI": f"sqlite:///{database}",
        }
    )
    with app.app_context():
        db.drop_all()
        db.create_all()
    yield app


@pytest.fixture()
def client(app):
    return app.test_client()


def test_complete_teacher_and_student_flow(app, client):
    response = client.post("/teacher/login", data={"password": "teacher-test-password"})
    assert response.status_code == 302

    questions = make_docx(
        [
            "ВАРИАНТ 1",
            "1. [ТЕСТ] Столица Франции?",
            "A) Париж",
            "B) Лондон",
            "2. [ОТКРЫТЫЙ] Сколько будет 2 + 2?",
        ]
    )
    response = client.post(
        "/teacher/upload/questions",
        data={"title": "Проверка", "questions_file": (questions, "questions.docx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert "Загружено" in response.get_data(as_text=True)

    answers = make_docx(["ВАРИАНТ 1", "1. A", "2. 4 | четыре"])
    response = client.post(
        "/teacher/exams/1/answers",
        data={"answers_file": (answers, "answers.docx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert "Сопоставлено ответов: 2" in response.get_data(as_text=True)

    response = client.post("/teacher/exams/1/toggle", follow_redirects=True)
    assert "Опубликована" in response.get_data(as_text=True)
    client.post("/teacher/logout")

    response = client.post(
        "/student/start",
        data={"student_name": "Иванов Иван Иванович", "group_name": "ИС-21", "exam_id": 1},
    )
    assert response.status_code == 302

    response = client.get("/student/question")
    assert "Столица Франции" in response.get_data(as_text=True)
    response = client.post("/student/question", data={"answer": "A"}, follow_redirects=True)
    assert "Сколько будет 2 + 2" in response.get_data(as_text=True)
    response = client.post("/student/question", data={"answer": "четыре"}, follow_redirects=True)
    page = response.get_data(as_text=True)
    assert "100.0%" in page
    assert "оценка" in page

    with app.app_context():
        attempt = Attempt.query.one()
        assert attempt.student_name == "Иванов Иван Иванович"
        assert attempt.group_name == "ИС-21"
        assert attempt.percentage == 100.0
        assert attempt.grade == "5"
        assert len(attempt.responses) == 2
