import csv
import hmac
import io
import os
import random
import re
import tempfile
import time
from datetime import datetime
from functools import wraps

from flask import (
    Flask,
    Response as FlaskResponse,
    abort,
    flash,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import UniqueConstraint, text
from sqlalchemy.exc import SQLAlchemyError
from werkzeug.utils import secure_filename

from docx_parser import DocumentFormatError, parse_answers, parse_questions


db = SQLAlchemy()


class Exam(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    question_doc_name = db.Column(db.String(255), nullable=False)
    answer_doc_name = db.Column(db.String(255))
    active = db.Column(db.Boolean, nullable=False, default=False)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    variants = db.relationship(
        "Variant", backref="exam", lazy=True, cascade="all, delete-orphan", order_by="Variant.id"
    )
    attempts = db.relationship("Attempt", backref="exam", lazy=True, cascade="all, delete-orphan")

    @property
    def answers_ready(self):
        questions = [question for variant in self.variants for question in variant.questions]
        return bool(questions) and all(question.correct_answers for question in questions)

    @property
    def question_count(self):
        return sum(len(variant.questions) for variant in self.variants)


class Variant(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    exam_id = db.Column(db.Integer, db.ForeignKey("exam.id"), nullable=False, index=True)
    code = db.Column(db.String(50), nullable=False)
    questions = db.relationship(
        "Question", backref="variant", lazy=True, cascade="all, delete-orphan", order_by="Question.number"
    )
    __table_args__ = (UniqueConstraint("exam_id", "code", name="uq_exam_variant_code"),)


class Question(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    variant_id = db.Column(db.Integer, db.ForeignKey("variant.id"), nullable=False, index=True)
    number = db.Column(db.Integer, nullable=False)
    text = db.Column(db.Text, nullable=False)
    question_type = db.Column(db.String(20), nullable=False)
    options_json = db.Column(db.JSON, nullable=False, default=list)
    correct_json = db.Column(db.JSON, nullable=False, default=list)
    points = db.Column(db.Float, nullable=False, default=1.0)
    __table_args__ = (UniqueConstraint("variant_id", "number", name="uq_variant_question_number"),)

    @property
    def options(self):
        return self.options_json or []

    @property
    def correct_answers(self):
        return self.correct_json or []

    @property
    def time_limit(self):
        return 40 if self.question_type == "test" else 120


class Attempt(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    exam_id = db.Column(db.Integer, db.ForeignKey("exam.id"), nullable=False, index=True)
    variant_id = db.Column(db.Integer, db.ForeignKey("variant.id"), nullable=False, index=True)
    student_name = db.Column(db.String(200), nullable=False, index=True)
    group_name = db.Column(db.String(100), nullable=False, index=True)
    started_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    submitted_at = db.Column(db.DateTime)
    status = db.Column(db.String(20), nullable=False, default="in_progress")
    score = db.Column(db.Float, nullable=False, default=0.0)
    max_score = db.Column(db.Float, nullable=False, default=0.0)
    percentage = db.Column(db.Float, nullable=False, default=0.0)
    grade = db.Column(db.String(10))
    variant = db.relationship("Variant")
    responses = db.relationship(
        "StudentResponse", backref="attempt", lazy=True, cascade="all, delete-orphan"
    )


class StudentResponse(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    attempt_id = db.Column(db.Integer, db.ForeignKey("attempt.id"), nullable=False, index=True)
    question_id = db.Column(db.Integer, db.ForeignKey("question.id"), nullable=False, index=True)
    answer_text = db.Column(db.Text, nullable=False, default="")
    is_correct = db.Column(db.Boolean, nullable=False, default=False)
    points_awarded = db.Column(db.Float, nullable=False, default=0.0)
    answered_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    time_limit = db.Column(db.Integer, nullable=False)
    timed_out = db.Column(db.Boolean, nullable=False, default=False)
    question = db.relationship("Question")
    __table_args__ = (UniqueConstraint("attempt_id", "question_id", name="uq_attempt_question"),)


def _database_url():
    value = os.getenv("DATABASE_URL", "").strip()
    if not value:
        # Локально SQLite удобен для разработки. На Render отсутствие DATABASE_URL
        # должно быть явной ошибкой, иначе результаты попадут во временный файл.
        if os.getenv("RENDER"):
            raise RuntimeError("DATABASE_URL is required on Render")
        return "sqlite:///control_work.db"
    if value.startswith("postgres://"):
        value = "postgresql+psycopg://" + value[len("postgres://") :]
    elif value.startswith("postgresql://"):
        value = "postgresql+psycopg://" + value[len("postgresql://") :]
    return value


def normalize_answer(value):
    value = (value or "").strip().lower().replace("ё", "е")
    value = re.sub(r"\s+", " ", value)
    return value.strip(" .,:;!?\"'«»")


def answer_is_correct(question, submitted):
    submitted_norm = normalize_answer(submitted)
    if not submitted_norm:
        return False

    accepted = {normalize_answer(item) for item in question.correct_answers}
    if submitted_norm in accepted:
        return True

    if question.question_type == "test":
        option = next(
            (item for item in question.options if normalize_answer(item.get("key")) == submitted_norm),
            None,
        )
        if option and normalize_answer(option.get("text")) in accepted:
            return True

    return False


def grade_for(percentage):
    if percentage >= 90:
        return "5"
    if percentage >= 75:
        return "4"
    if percentage >= 60:
        return "3"
    return "2"


def finalize_attempt(attempt):
    score = sum(response.points_awarded for response in attempt.responses)
    maximum = sum(question.points for question in attempt.variant.questions)
    percentage = round((score / maximum * 100) if maximum else 0, 1)
    attempt.score = round(score, 2)
    attempt.max_score = round(maximum, 2)
    attempt.percentage = percentage
    attempt.grade = grade_for(percentage)
    attempt.status = "completed"
    attempt.submitted_at = datetime.utcnow()
    db.session.commit()


def create_app(test_config=None):
    app = Flask(__name__)
    app.config.update(
        SECRET_KEY=os.getenv("SECRET_KEY", "local-development-only-change-me"),
        SQLALCHEMY_DATABASE_URI=_database_url(),
        SQLALCHEMY_TRACK_MODIFICATIONS=False,
        SQLALCHEMY_ENGINE_OPTIONS={"pool_pre_ping": True, "pool_recycle": 300},
        MAX_CONTENT_LENGTH=16 * 1024 * 1024,
        TEACHER_PASSWORD=os.getenv("TEACHER_PASSWORD", ""),
    )
    if test_config:
        app.config.update(test_config)

    db.init_app(app)
    with app.app_context():
        db.create_all()

    def teacher_required(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if not session.get("teacher_logged_in"):
                return redirect(url_for("teacher_login", next=request.path))
            return view(*args, **kwargs)

        return wrapped

    def save_temp_docx(upload):
        if not upload or not upload.filename:
            raise DocumentFormatError("Выберите DOCX-файл.")
        filename = secure_filename(upload.filename)
        if not filename.lower().endswith(".docx"):
            raise DocumentFormatError("Поддерживаются только файлы .docx.")
        temp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        upload.save(temp.name)
        temp.close()
        return temp.name, filename

    @app.errorhandler(413)
    def too_large(_error):
        flash("Файл слишком большой. Максимальный размер — 16 МБ.", "error")
        return redirect(request.referrer or url_for("index"))

    @app.get("/health/db")
    def database_health():
        try:
            db.session.execute(text("SELECT 1"))
        except SQLAlchemyError:
            db.session.rollback()
            return {"database": "unavailable", "status": "error"}, 503
        return {"database": "connected", "status": "ok"}, 200

    @app.route("/")
    def index():
        active_exams = Exam.query.filter_by(active=True).order_by(Exam.created_at.desc()).all()
        return render_template("index.html", active_exams=active_exams)

    @app.route("/teacher/login", methods=["GET", "POST"])
    def teacher_login():
        if request.method == "POST":
            configured = app.config.get("TEACHER_PASSWORD", "")
            submitted = request.form.get("password", "")
            if not configured:
                flash("На сервере не задана переменная TEACHER_PASSWORD.", "error")
            elif hmac.compare_digest(configured, submitted):
                session.clear()
                session["teacher_logged_in"] = True
                return redirect(url_for("teacher_dashboard"))
            else:
                flash("Неверный пароль.", "error")
        return render_template("teacher_login.html")

    @app.route("/teacher/logout", methods=["POST"])
    def teacher_logout():
        session.clear()
        return redirect(url_for("index"))

    @app.route("/teacher")
    @teacher_required
    def teacher_dashboard():
        exams = Exam.query.order_by(Exam.created_at.desc()).all()
        attempts = Attempt.query.order_by(Attempt.started_at.desc()).all()
        return render_template("teacher_dashboard.html", exams=exams, attempts=attempts)

    @app.route("/teacher/upload/questions", methods=["POST"])
    @teacher_required
    def upload_questions():
        temp_path = None
        try:
            title = request.form.get("title", "").strip()
            if not title:
                raise DocumentFormatError("Укажите название контрольной работы.")
            temp_path, original_name = save_temp_docx(request.files.get("questions_file"))
            parsed_variants = parse_questions(temp_path)

            exam = Exam(title=title, question_doc_name=original_name)
            db.session.add(exam)
            for parsed_variant in parsed_variants:
                variant = Variant(exam=exam, code=parsed_variant.code)
                db.session.add(variant)
                for parsed_question in parsed_variant.questions:
                    db.session.add(
                        Question(
                            variant=variant,
                            number=parsed_question.number,
                            text=parsed_question.text,
                            question_type=parsed_question.question_type,
                            options_json=parsed_question.options,
                            correct_json=[],
                        )
                    )
            db.session.commit()
            flash(
                f"Загружено: {len(parsed_variants)} вариант(ов), {exam.question_count} вопрос(ов). Теперь добавьте ответы.",
                "success",
            )
        except (DocumentFormatError, ValueError) as error:
            db.session.rollback()
            flash(str(error), "error")
        finally:
            if temp_path and os.path.exists(temp_path):
                os.unlink(temp_path)
        return redirect(url_for("teacher_dashboard"))

    @app.route("/teacher/exams/<int:exam_id>/answers", methods=["POST"])
    @teacher_required
    def upload_answers(exam_id):
        exam = db.get_or_404(Exam, exam_id)
        temp_path = None
        try:
            temp_path, original_name = save_temp_docx(request.files.get("answers_file"))
            parsed = parse_answers(temp_path)
            missing = []
            matched = 0
            for variant in exam.variants:
                variant_answers = parsed.get(variant.code, {})
                for question in variant.questions:
                    accepted = variant_answers.get(question.number)
                    if accepted:
                        question.correct_json = accepted
                        matched += 1
                    else:
                        missing.append(f"{variant.code}.{question.number}")

            if missing:
                raise DocumentFormatError(
                    "Не найдены ответы для вопросов: " + ", ".join(missing[:20])
                    + ("…" if len(missing) > 20 else "")
                )
            exam.answer_doc_name = original_name
            db.session.commit()
            flash(f"Файл ответов принят. Сопоставлено ответов: {matched}.", "success")
        except (DocumentFormatError, ValueError) as error:
            db.session.rollback()
            flash(str(error), "error")
        finally:
            if temp_path and os.path.exists(temp_path):
                os.unlink(temp_path)
        return redirect(url_for("teacher_dashboard"))

    @app.route("/teacher/exams/<int:exam_id>/toggle", methods=["POST"])
    @teacher_required
    def toggle_exam(exam_id):
        exam = db.get_or_404(Exam, exam_id)
        if not exam.active and not exam.answers_ready:
            flash("Сначала загрузите полный файл ответов.", "error")
        else:
            exam.active = not exam.active
            db.session.commit()
            flash("Контрольная опубликована." if exam.active else "Контрольная скрыта.", "success")
        return redirect(url_for("teacher_dashboard"))

    @app.route("/teacher/export.csv")
    @teacher_required
    def export_results():
        attempts = Attempt.query.order_by(Attempt.started_at.desc()).all()
        stream = io.StringIO()
        writer = csv.writer(stream, delimiter=";")
        writer.writerow(["ФИО", "Группа", "Контрольная", "Вариант", "Баллы", "Максимум", "%", "Оценка", "Статус", "Дата UTC"])
        for attempt in attempts:
            writer.writerow(
                [
                    attempt.student_name,
                    attempt.group_name,
                    attempt.exam.title,
                    attempt.variant.code,
                    attempt.score,
                    attempt.max_score,
                    attempt.percentage,
                    attempt.grade or "",
                    attempt.status,
                    (attempt.submitted_at or attempt.started_at).strftime("%Y-%m-%d %H:%M:%S"),
                ]
            )
        payload = "\ufeff" + stream.getvalue()
        return FlaskResponse(
            payload,
            mimetype="text/csv; charset=utf-8",
            headers={"Content-Disposition": "attachment; filename=results.csv"},
        )

    @app.route("/student/start", methods=["POST"])
    def student_start():
        student_name = " ".join(request.form.get("student_name", "").split())
        group_name = " ".join(request.form.get("group_name", "").split())
        exam_id = request.form.get("exam_id", type=int)
        exam = db.session.get(Exam, exam_id) if exam_id else None

        if len(student_name) < 3 or len(student_name) > 200:
            flash("Укажите ФИО студента.", "error")
        elif len(group_name) < 1 or len(group_name) > 100:
            flash("Укажите группу.", "error")
        elif not exam or not exam.active or not exam.answers_ready:
            flash("Эта контрольная сейчас недоступна.", "error")
        elif not exam.variants:
            flash("В контрольной нет вариантов.", "error")
        else:
            variant = random.choice(exam.variants)
            attempt = Attempt(
                exam=exam,
                variant=variant,
                student_name=student_name,
                group_name=group_name,
            )
            db.session.add(attempt)
            db.session.commit()
            session.clear()
            session["attempt_id"] = attempt.id
            return redirect(url_for("student_question"))
        return redirect(url_for("index"))

    @app.route("/student/question", methods=["GET", "POST"])
    def student_question():
        attempt_id = session.get("attempt_id")
        attempt = db.session.get(Attempt, attempt_id) if attempt_id else None
        if not attempt:
            flash("Сначала укажите ФИО и группу.", "error")
            return redirect(url_for("index"))
        if attempt.status == "completed":
            return redirect(url_for("student_result", attempt_id=attempt.id))

        answered_ids = {response.question_id for response in attempt.responses}
        question = next(
            (item for item in attempt.variant.questions if item.id not in answered_ids),
            None,
        )
        if question is None:
            finalize_attempt(attempt)
            return redirect(url_for("student_result", attempt_id=attempt.id))

        timer_key = f"question_started_{attempt.id}_{question.id}"
        if request.method == "POST":
            started = float(session.get(timer_key, time.time()))
            elapsed = max(0, time.time() - started)
            timed_out = elapsed > question.time_limit + 3
            submitted = "" if timed_out else request.form.get("answer", "").strip()
            correct = False if timed_out else answer_is_correct(question, submitted)
            response = StudentResponse(
                attempt=attempt,
                question=question,
                answer_text=submitted,
                is_correct=correct,
                points_awarded=question.points if correct else 0,
                time_limit=question.time_limit,
                timed_out=timed_out,
            )
            db.session.add(response)
            session.pop(timer_key, None)
            db.session.commit()
            return redirect(url_for("student_question"))

        session.setdefault(timer_key, time.time())
        session.modified = True
        position = len(answered_ids) + 1
        total = len(attempt.variant.questions)
        return render_template(
            "student_question.html",
            attempt=attempt,
            question=question,
            position=position,
            total=total,
        )

    @app.route("/student/result/<int:attempt_id>")
    def student_result(attempt_id):
        attempt = db.get_or_404(Attempt, attempt_id)
        if session.get("attempt_id") != attempt.id and not session.get("teacher_logged_in"):
            abort(403)
        if attempt.status != "completed":
            return redirect(url_for("student_question"))
        responses = sorted(attempt.responses, key=lambda item: item.question.number)
        return render_template("student_result.html", attempt=attempt, responses=responses)

    return app


app = create_app()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=True)
