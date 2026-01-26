import calendar
import logging
import multiprocessing as mp
import os
import re
import secrets
from datetime import date, datetime, timedelta
from functools import wraps
from uuid import uuid4

import httpx
from flask import Flask, jsonify, redirect, render_template, request, session, url_for, g
from sqlalchemy import (
    Boolean,
    Column,
    Float,
    Integer,
    MetaData,
    String,
    Text,
    Table,
    create_engine,
    func,
    inspect,
    select,
    text,
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

from services.ai_invoice_service import analyze_invoice
from services.storage_service import get_public_url, upload_bytes

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "data.db")
ALLOWED_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png"}
ANALYSIS_TIMEOUT_SECONDS = int(os.getenv("ANALYSIS_TIMEOUT_SECONDS", "120"))
DEFAULT_USER_ID = int(os.getenv("DEFAULT_USER_ID", "1"))
OWNER_EMAIL = (os.getenv("OWNER_EMAIL") or "").strip().lower()
RESEND_API_KEY = os.getenv("RESEND_API_KEY", "")
APP_FROM_EMAIL = os.getenv("APP_FROM_EMAIL", "no-reply@tuapp.com")

_raw_db_url = os.getenv("DATABASE_URL")
DATABASE_URL = _raw_db_url.strip() if _raw_db_url else ""
if not DATABASE_URL:
    DATABASE_URL = f"sqlite:///{DB_PATH}"
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)
if DATABASE_URL.startswith("sqlite"):
    logging.warning("DATABASE_URL no configurada. Usando SQLite local.")

engine = create_engine(DATABASE_URL, pool_pre_ping=True, future=True)
metadata = MetaData()

companies_table = Table(
    "companies",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("user_id", Integer, nullable=False),
    Column("agency_id", Integer, nullable=True),
    Column("display_name", String, nullable=False),
    Column("legal_name", String, nullable=False),
    Column("tax_id", String, nullable=False),
    Column("company_type", String, nullable=False),  # individual | company
    Column("email", String),
    Column("phone", String),
    Column("assigned_user_id", Integer),
    Column("created_at", String, nullable=False),
)

users_table = Table(
    "users",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("email", String, nullable=False, unique=True),
    Column("password_hash", String, nullable=False),
    Column("role", String, nullable=False),  # owner | agency | staff
    Column("plan", String, nullable=False),  # internal | trial | standard | premium
    Column("agency_id", Integer),
    Column("created_at", String, nullable=False),
    Column("is_active", Boolean, nullable=False, server_default="1"),
)

password_resets_table = Table(
    "password_resets",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("user_id", Integer, nullable=False),
    Column("token", String, nullable=False),
    Column("expires_at", String, nullable=False),
    Column("used_at", String),
)

invoices_table = Table(
    "invoices",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("user_id", Integer, nullable=False, server_default=str(DEFAULT_USER_ID)),
    Column("company_id", Integer, nullable=False),
    Column("original_filename", String, nullable=False),
    Column("stored_filename", String, nullable=False),
    Column("invoice_date", String, nullable=False),
    Column("supplier", String, nullable=False),
    Column("base_amount", Float, nullable=False),
    Column("vat_rate", Integer, nullable=False),
    Column("vat_amount", Float),
    Column("total_amount", Float, nullable=False),
    Column("payment_date", String),
    Column("ocr_text", Text),
    Column("expense_category", String, nullable=False, server_default="with_invoice"),
    Column("created_at", String, nullable=False),
)

income_invoices_table = Table(
    "income_invoices",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("user_id", Integer, nullable=False, server_default=str(DEFAULT_USER_ID)),
    Column("company_id", Integer, nullable=False),
    Column("original_filename", String, nullable=False),
    Column("stored_filename", String, nullable=False),
    Column("invoice_date", String, nullable=False),
    Column("client", String, nullable=False),
    Column("base_amount", Float, nullable=False),
    Column("vat_rate", Integer, nullable=False),
    Column("vat_amount", Float),
    Column("total_amount", Float, nullable=False),
    Column("payment_date", String),
    Column("ocr_text", Text),
    Column("created_at", String, nullable=False),
)

facturacion_table = Table(
    "facturacion",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("user_id", Integer, nullable=False, server_default=str(DEFAULT_USER_ID)),
    Column("company_id", Integer, nullable=False),
    Column("mes", Integer, nullable=False),
    Column("anio", Integer, nullable=False),
    Column("invoice_date", String),
    Column("concept", String),
    Column("base_facturada", Float, nullable=False),
    Column("tipo_iva", Integer, nullable=False),
    Column("iva_repercutido", Float, nullable=False),
    Column("total_amount", Float),
)

no_invoice_table = Table(
    "no_invoice_expenses",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("user_id", Integer, nullable=False, server_default=str(DEFAULT_USER_ID)),
    Column("company_id", Integer, nullable=False),
    Column("expense_date", String, nullable=False),
    Column("concept", String, nullable=False),
    Column("amount", Float, nullable=False),
    Column("expense_type", String, nullable=False),
    Column("deductible", Boolean, nullable=False),
    Column("created_at", String, nullable=False),
)

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY") or secrets.token_urlsafe(32)
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.getenv("ENV", "").lower() == "production"


def init_db():
    metadata.create_all(engine)
    inspector = inspect(engine)
    table_names = set(inspector.get_table_names())

    def add_column_if_missing(table_name, column_name, column_type):
        if table_name not in table_names:
            return
        columns = {col["name"] for col in inspector.get_columns(table_name)}
        if column_name in columns:
            return
        with engine.begin() as conn:
            conn.execute(
                text(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}")
            )

    add_column_if_missing("invoices", "user_id", "INTEGER")
    add_column_if_missing("invoices", "company_id", "INTEGER")
    add_column_if_missing("invoices", "payment_date", "VARCHAR")
    if "invoices" in table_names:
        with engine.begin() as conn:
            conn.execute(
                invoices_table.update()
                .where(invoices_table.c.user_id.is_(None))
                .values(user_id=DEFAULT_USER_ID)
            )

    add_column_if_missing("facturacion", "user_id", "INTEGER")
    add_column_if_missing("facturacion", "company_id", "INTEGER")
    add_column_if_missing("facturacion", "invoice_date", "VARCHAR")
    add_column_if_missing("facturacion", "concept", "VARCHAR")
    add_column_if_missing("facturacion", "total_amount", "FLOAT")
    if "facturacion" in table_names:
        with engine.begin() as conn:
            conn.execute(
                facturacion_table.update()
                .where(facturacion_table.c.user_id.is_(None))
                .values(user_id=DEFAULT_USER_ID)
            )

    add_column_if_missing("no_invoice_expenses", "user_id", "INTEGER")
    add_column_if_missing("no_invoice_expenses", "company_id", "INTEGER")
    if "no_invoice_expenses" in table_names:
        with engine.begin() as conn:
            conn.execute(
                no_invoice_table.update()
                .where(no_invoice_table.c.user_id.is_(None))
                .values(user_id=DEFAULT_USER_ID)
            )

    add_column_if_missing("income_invoices", "user_id", "INTEGER")
    add_column_if_missing("income_invoices", "company_id", "INTEGER")
    add_column_if_missing("income_invoices", "payment_date", "VARCHAR")
    if "income_invoices" in table_names:
        with engine.begin() as conn:
            conn.execute(
                income_invoices_table.update()
                .where(income_invoices_table.c.user_id.is_(None))
                .values(user_id=DEFAULT_USER_ID)
            )

    add_column_if_missing("companies", "agency_id", "INTEGER")
    add_column_if_missing("companies", "email", "VARCHAR")
    add_column_if_missing("companies", "phone", "VARCHAR")
    add_column_if_missing("companies", "assigned_user_id", "INTEGER")
    if "companies" in table_names:
        with engine.begin() as conn:
            conn.execute(
                companies_table.update()
                .where(companies_table.c.user_id.is_(None))
                .values(user_id=DEFAULT_USER_ID)
            )
            conn.execute(
                companies_table.update()
                .where(companies_table.c.agency_id.is_(None))
                .values(agency_id=companies_table.c.user_id)
            )

    add_column_if_missing("users", "agency_id", "INTEGER")
    if "users" in table_names:
        with engine.begin() as conn:
            conn.execute(
                users_table.update()
                .where(users_table.c.agency_id.is_(None))
                .values(agency_id=users_table.c.id)
            )


def allowed_file(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTENSIONS


def _row_to_user(row):
    if not row:
        return None
    return {
        "id": row["id"],
        "email": row["email"],
        "role": row["role"],
        "plan": row["plan"],
        "agency_id": row.get("agency_id") if isinstance(row, dict) else None,
        "is_active": bool(row["is_active"]),
    }


def get_current_user():
    user_id = session.get("user_id")
    if not user_id:
        return None
    with engine.connect() as conn:
        row = conn.execute(
            select(
                users_table.c.id,
                users_table.c.email,
                users_table.c.role,
                users_table.c.plan,
                users_table.c.agency_id,
                users_table.c.is_active,
            ).where(users_table.c.id == user_id)
        ).mappings().first()
    user = _row_to_user(row)
    if user and not user["is_active"]:
        return None
    return user


def plan_allows(user, allowed_plans):
    if not user:
        return False
    return user.get("plan") in allowed_plans


def require_plan(allowed_plans):
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            user = getattr(g, "current_user", None)
            if not plan_allows(user, allowed_plans):
                return jsonify({"ok": False, "errors": ["Plan insuficiente."]}), 403
            return view(*args, **kwargs)

        return wrapped

    return decorator


@app.before_request
def load_user_and_enforce_auth():
    g.current_user = get_current_user()
    path = request.path or ""
    if path.startswith("/static/"):
        return None
    if path.startswith("/login") or path.startswith("/register"):
        return None
    if path.startswith("/reset") or path.startswith("/reset-password"):
        return None
    if path.startswith("/health"):
        return None
    if not g.current_user:
        if path.startswith("/api/"):
            return jsonify({"ok": False, "error": "auth_required"}), 401
        return redirect(url_for("login"))
    return None


def parse_amount(value):
    if value is None:
        return None
    cleaned = value.replace("EUR", "").replace("euro", "").strip()
    cleaned = cleaned.replace(".", "").replace(",", ".") if "," in cleaned else cleaned
    try:
        return float(cleaned)
    except ValueError:
        return None


def get_current_user_id():
    if getattr(g, "current_user", None):
        return int(g.current_user["id"])
    try:
        header_value = request.headers.get("X-User-Id")
        if header_value and str(header_value).isdigit():
            return int(header_value)
    except Exception:
        pass
    return DEFAULT_USER_ID


def get_data_owner_id():
    user = g.current_user
    if not user:
        return DEFAULT_USER_ID
    if user.get("role") == "staff":
        return int(user.get("agency_id") or user["id"])
    return int(user["id"])


def _resolve_company_id():
    company_id = None
    if request.is_json:
        payload = request.get_json(silent=True) or {}
        company_id = payload.get("company_id") or payload.get("companyId")
    if company_id is None:
        company_id = request.args.get("company_id") or request.form.get("company_id")
    if company_id is None:
        return None
    try:
        return int(company_id)
    except (TypeError, ValueError):
        return None


def get_company_id(required=True):
    user_id = get_current_user_id()
    user_role = (g.current_user or {}).get("role")
    company_id = _resolve_company_id()
    with engine.connect() as conn:
        if company_id is None:
            if user_role == "staff":
                row = conn.execute(
                    select(companies_table.c.id).where(
                        companies_table.c.assigned_user_id == user_id
                    )
                ).first()
            elif user_role == "owner":
                row = conn.execute(select(companies_table.c.id)).first()
            else:
                row = conn.execute(
                    select(companies_table.c.id).where(
                        companies_table.c.agency_id == user_id
                    )
                ).first()
            if row:
                company_id = int(row[0])
        else:
            if user_role == "staff":
                exists = conn.execute(
                    select(companies_table.c.id)
                    .where(companies_table.c.assigned_user_id == user_id)
                    .where(companies_table.c.id == company_id)
                ).first()
            elif user_role == "owner":
                exists = conn.execute(
                    select(companies_table.c.id).where(
                        companies_table.c.id == company_id
                    )
                ).first()
            else:
                exists = conn.execute(
                    select(companies_table.c.id)
                    .where(companies_table.c.agency_id == user_id)
                    .where(companies_table.c.id == company_id)
                ).first()
            if not exists:
                company_id = None

    if required and company_id is None:
        return None
    return company_id


def is_company_accessible(company_id):
    if company_id is None:
        return False
    user_id = get_current_user_id()
    role = (g.current_user or {}).get("role")
    with engine.connect() as conn:
        if role == "staff":
            exists = conn.execute(
                select(companies_table.c.id)
                .where(companies_table.c.assigned_user_id == user_id)
                .where(companies_table.c.id == company_id)
            ).first()
        elif role == "owner":
            exists = conn.execute(
                select(companies_table.c.id).where(companies_table.c.id == company_id)
            ).first()
        else:
            exists = conn.execute(
                select(companies_table.c.id)
                .where(companies_table.c.agency_id == user_id)
                .where(companies_table.c.id == company_id)
            ).first()
    return bool(exists)


def _validate_nif(nif):
    if not nif:
        return False
    nif = nif.strip().upper()
    match = re.match(r"^(\d{8})([A-Z])$", nif)
    if not match:
        return False
    number, letter = match.groups()
    letters = "TRWAGMYFPDXBNJZSQVHLCKE"
    return letters[int(number) % 23] == letter


def _validate_cif(cif):
    if not cif:
        return False
    cif = cif.strip().upper()
    match = re.match(r"^([ABCDEFGHJKLMNPQRSUVW])(\d{7})([0-9A-J])$", cif)
    if not match:
        return False
    letter, digits, control = match.groups()
    total = 0
    for idx, char in enumerate(digits, start=1):
        n = int(char)
        if idx % 2 == 1:
            n *= 2
            total += n // 10 + n % 10
        else:
            total += n
    control_num = (10 - (total % 10)) % 10
    control_digit = str(control_num)
    control_letter = "JABCDEFGHI"[control_num]
    if letter in "PQRSW":
        return control == control_letter
    if letter in "ABEH":
        return control == control_digit
    return control in {control_digit, control_letter}


def validate_tax_id(tax_id, company_type):
    if company_type == "individual":
        return _validate_nif(tax_id)
    if company_type == "company":
        return _validate_cif(tax_id)
    return False


def resolve_assigned_staff(agency_id, staff_id):
    if not staff_id:
        return None
    try:
        staff_id = int(staff_id)
    except (TypeError, ValueError):
        return None
    with engine.connect() as conn:
        staff = conn.execute(
            select(users_table.c.id)
            .where(users_table.c.id == staff_id)
            .where(users_table.c.role == "staff")
            .where(users_table.c.agency_id == agency_id)
        ).first()
    if not staff:
        return None
    return staff_id


def normalize_date(value):
    if not value:
        return None
    raw = str(value).strip()
    if not raw:
        return None
    try:
        return date.fromisoformat(raw).isoformat()
    except ValueError:
        return None


def compute_payment_date(invoice_date_value, payment_date_value=None):
    payment_date = normalize_date(payment_date_value)
    if payment_date:
        return payment_date
    invoice_date = normalize_date(invoice_date_value)
    if not invoice_date:
        return None
    try:
        base_date = date.fromisoformat(invoice_date)
    except ValueError:
        return None
    return (base_date + timedelta(days=30)).isoformat()


def _parse_period_params():
    year = request.args.get("year") or request.args.get("anio") or request.args.get("año")
    quarter = request.args.get("quarter")
    start_month = request.args.get("start_month")
    end_month = request.args.get("end_month")
    try:
        year = int(year)
    except (TypeError, ValueError):
        year = None
    try:
        quarter = int(quarter) if quarter else None
    except (TypeError, ValueError):
        quarter = None
    try:
        start_month = int(start_month) if start_month else None
    except (TypeError, ValueError):
        start_month = None
    try:
        end_month = int(end_month) if end_month else None
    except (TypeError, ValueError):
        end_month = None
    return year, quarter, start_month, end_month


def _get_months_for_period(year, quarter=None, start_month=None, end_month=None):
    if not year:
        return []
    if quarter in {1, 2, 3, 4}:
        start = (quarter - 1) * 3 + 1
        return [start, start + 1, start + 2]
    if start_month and end_month and 1 <= start_month <= 12 and 1 <= end_month <= 12:
        if start_month <= end_month:
            return list(range(start_month, end_month + 1))
        return list(range(start_month, 13)) + list(range(1, end_month + 1))
    return list(range(1, 13))


def _build_report_totals(user_id, company_id, months, year):
    income_base = 0.0
    income_vat = 0.0
    expense_base = 0.0
    expense_vat = 0.0
    with engine.connect() as conn:
        for month in months:
            prefix = f"{year}-{month:02d}"
            invoice_rows = conn.execute(
                select(
                    invoices_table.c.base_amount,
                    invoices_table.c.vat_amount,
                    invoices_table.c.expense_category,
                )
                .where(invoices_table.c.user_id == user_id)
                .where(invoices_table.c.company_id == company_id)
                .where(invoices_table.c.invoice_date.like(f"{prefix}%"))
            ).mappings().all()
            for row in invoice_rows:
                if row["expense_category"] == "non_deductible":
                    continue
                expense_base += float(row["base_amount"] or 0)
                expense_vat += float(row["vat_amount"] or 0)

            no_invoice_rows = conn.execute(
                select(no_invoice_table.c.amount, no_invoice_table.c.deductible)
                .where(no_invoice_table.c.user_id == user_id)
                .where(no_invoice_table.c.company_id == company_id)
                .where(no_invoice_table.c.expense_date.like(f"{prefix}%"))
            ).mappings().all()
            for row in no_invoice_rows:
                if not row["deductible"]:
                    continue
                expense_base += float(row["amount"] or 0)

            income_invoice_rows = conn.execute(
                select(income_invoices_table.c.base_amount, income_invoices_table.c.vat_amount)
                .where(income_invoices_table.c.user_id == user_id)
                .where(income_invoices_table.c.company_id == company_id)
                .where(income_invoices_table.c.invoice_date.like(f"{prefix}%"))
            ).mappings().all()
            for row in income_invoice_rows:
                income_base += float(row["base_amount"] or 0)
                income_vat += float(row["vat_amount"] or 0)

            billing_rows = conn.execute(
                select(
                    facturacion_table.c.base_facturada,
                    facturacion_table.c.iva_repercutido,
                )
                .where(facturacion_table.c.user_id == user_id)
                .where(facturacion_table.c.company_id == company_id)
                .where(facturacion_table.c.anio == year)
                .where(facturacion_table.c.mes == month)
            ).mappings().all()
            for row in billing_rows:
                income_base += float(row["base_facturada"] or 0)
                income_vat += float(row["iva_repercutido"] or 0)

    return {
        "income_base": round(income_base, 2),
        "income_vat": round(income_vat, 2),
        "expense_base": round(expense_base, 2),
        "expense_vat": round(expense_vat, 2),
        "net_result": round(income_base - expense_base, 2),
        "vat_result": round(income_vat - expense_vat, 2),
    }


def _report_period_label(year, months, quarter=None):
    if quarter in {1, 2, 3, 4}:
        return f"T{quarter} {year}"
    if months:
        if len(months) == 12:
            return f"Año {year}"
        if len(months) == 1:
            return f"{calendar.month_name[months[0]]} {year}"
        return f"{calendar.month_name[months[0]]} - {calendar.month_name[months[-1]]} {year}"
    return str(year)


def _empty_extracted():
    return {
        "provider_name": None,
        "invoice_date": None,
        "payment_date": None,
        "base_amount": None,
        "vat_rate": None,
        "vat_amount": None,
        "total_amount": None,
        "analysis_text": "",
        "validation": {"is_consistent": None, "difference": None},
    }


def _analysis_worker(file_bytes, filename, mime_type, document_type, queue):
    try:
        result = analyze_invoice(
            file_bytes=file_bytes,
            filename=filename,
            mime_type=mime_type,
            document_type=document_type,
        )
        queue.put(result)
    except Exception as exc:
        queue.put({"__error__": str(exc)})


def _analyze_invoice_with_timeout(file_bytes, filename, stored_name, mime_type, document_type="expense"):
    ctx = mp.get_context("spawn")
    queue = ctx.Queue(1)
    process = ctx.Process(
        target=_analysis_worker,
        args=(file_bytes, filename, mime_type, document_type, queue),
    )
    process.start()
    process.join(ANALYSIS_TIMEOUT_SECONDS)

    if process.is_alive():
        process.terminate()
        process.join()
        app.logger.warning(
            "Timeout analizando %s (> %ss). Se pasa a modo manual.",
            stored_name,
            ANALYSIS_TIMEOUT_SECONDS,
        )
        return _empty_extracted()

    if process.exitcode != 0:
        app.logger.warning(
            "Analisis fallido para %s (exitcode %s). Se pasa a modo manual.",
            stored_name,
            process.exitcode,
        )
        return _empty_extracted()

    try:
        result = queue.get_nowait()
    except Exception:
        app.logger.warning(
            "Analisis sin resultado para %s. Se pasa a modo manual.",
            stored_name,
        )
        return _empty_extracted()

    if not isinstance(result, dict) or result.get("__error__"):
        app.logger.warning(
            "Analisis con error para %s. Se pasa a modo manual.",
            stored_name,
        )
        return _empty_extracted()

    return result


def _normalize_email(value):
    return (value or "").strip().lower()


def send_email(to_email, subject, html_content, reply_to=None):
    if not RESEND_API_KEY:
        app.logger.warning("RESEND_API_KEY no configurada. Email no enviado.")
        return False
    payload = {
        "from": APP_FROM_EMAIL,
        "to": [to_email],
        "subject": subject,
        "html": html_content,
    }
    if reply_to:
        payload["reply_to"] = reply_to
    try:
        response = httpx.post(
            "https://api.resend.com/emails",
            headers={"Authorization": f"Bearer {RESEND_API_KEY}"},
            json=payload,
            timeout=20.0,
        )
        response.raise_for_status()
        return True
    except Exception:
        app.logger.exception("Error enviando email a %s", to_email)
        return False


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET":
        return render_template("login.html")
    payload = request.form or request.get_json(silent=True) or {}
    email = _normalize_email(payload.get("email"))
    password = payload.get("password") or ""
    if not email or not password:
        return render_template("login.html", error="Email y contraseña obligatorios.")
    with engine.connect() as conn:
        row = conn.execute(
            select(
                users_table.c.id,
                users_table.c.email,
                users_table.c.password_hash,
                users_table.c.role,
                users_table.c.plan,
                users_table.c.is_active,
            ).where(users_table.c.email == email)
        ).mappings().first()
    if not row or not row["is_active"]:
        return render_template("login.html", error="Credenciales inválidas.")
    if not check_password_hash(row["password_hash"], password):
        return render_template("login.html", error="Credenciales inválidas.")
    session["user_id"] = row["id"]
    return redirect(url_for("index"))


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "GET":
        return render_template("register.html")
    payload = request.form or request.get_json(silent=True) or {}
    email = _normalize_email(payload.get("email"))
    password = payload.get("password") or ""
    if not email or not password:
        return render_template("register.html", error="Email y contraseña obligatorios.")
    if len(password) < 8:
        return render_template(
            "register.html", error="La contraseña debe tener al menos 8 caracteres."
        )
    role = "agency"
    plan = "trial"
    if OWNER_EMAIL and email == OWNER_EMAIL:
        role = "owner"
        plan = "premium"
    with engine.begin() as conn:
        owner_exists = conn.execute(
            select(users_table.c.id).where(users_table.c.role == "owner")
        ).first()
        if owner_exists and role == "owner":
            return render_template("register.html", error="El usuario propietario ya existe.")
        exists = conn.execute(
            select(users_table.c.id).where(users_table.c.email == email)
        ).first()
        if exists:
            return render_template("register.html", error="El email ya está registrado.")
        result = conn.execute(
            users_table.insert().values(
                email=email,
                password_hash=generate_password_hash(password),
                role=role,
                plan=plan,
                agency_id=None,
                created_at=datetime.utcnow().isoformat(),
                is_active=True,
            )
        )
        new_user_id = result.inserted_primary_key[0]
        conn.execute(
            users_table.update()
            .where(users_table.c.id == new_user_id)
            .values(agency_id=new_user_id)
        )
        session["user_id"] = new_user_id
    return redirect(url_for("index"))


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/reset", methods=["GET", "POST"])
@app.route("/reset-password", methods=["GET", "POST"])
def reset_password_request():
    if request.method == "GET":
        return render_template("reset_request.html")
    payload = request.form or request.get_json(silent=True) or {}
    email = _normalize_email(payload.get("email"))
    if not email:
        return render_template("reset_request.html", error="Email obligatorio.")
    with engine.connect() as conn:
        row = conn.execute(
            select(users_table.c.id, users_table.c.email)
            .where(users_table.c.email == email)
            .where(users_table.c.is_active.is_(True))
        ).mappings().first()
    if row:
        token = secrets.token_urlsafe(32)
        expires_at = (datetime.utcnow() + timedelta(hours=1)).isoformat()
        with engine.begin() as conn:
            conn.execute(
                password_resets_table.insert().values(
                    user_id=row["id"],
                    token=token,
                    expires_at=expires_at,
                    used_at=None,
                )
            )
        reset_link = url_for("reset_password", token=token, _external=True)
        html = f"""
        <p>Has solicitado restablecer tu contraseña.</p>
        <p>Enlace válido durante 1 hora:</p>
        <p><a href="{reset_link}">Restablecer contraseña</a></p>
        """
        send_email(email, "Restablece tu contraseña", html)
    return render_template(
        "reset_request.html",
        message="Si el email existe, recibirás un enlace de recuperación.",
    )


@app.route("/reset/<token>", methods=["GET", "POST"])
@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password(token):
    if request.method == "GET":
        return render_template("reset_password.html", token=token)
    payload = request.form or request.get_json(silent=True) or {}
    password = payload.get("password") or ""
    if len(password) < 8:
        return render_template(
            "reset_password.html",
            token=token,
            error="La contraseña debe tener al menos 8 caracteres.",
        )
    now = datetime.utcnow().isoformat()
    with engine.begin() as conn:
        reset_row = conn.execute(
            select(
                password_resets_table.c.id,
                password_resets_table.c.user_id,
                password_resets_table.c.expires_at,
                password_resets_table.c.used_at,
            ).where(password_resets_table.c.token == token)
        ).mappings().first()
        if (
            not reset_row
            or reset_row["used_at"]
            or reset_row["expires_at"] < now
        ):
            return render_template(
                "reset_password.html",
                token=token,
                error="El enlace no es válido o ha caducado.",
            )
        conn.execute(
            users_table.update()
            .where(users_table.c.id == reset_row["user_id"])
            .values(password_hash=generate_password_hash(password))
        )
        conn.execute(
            password_resets_table.update()
            .where(password_resets_table.c.id == reset_row["id"])
            .values(used_at=now)
        )
    return redirect(url_for("login"))


@app.route("/")
def index():
    return render_template("index.html", user=g.current_user)


@app.route("/api/companies")
def list_companies():
    user_id = get_current_user_id()
    user_role = (g.current_user or {}).get("role")
    base_query = select(
        companies_table.c.id,
        companies_table.c.display_name,
        companies_table.c.legal_name,
        companies_table.c.tax_id,
        companies_table.c.company_type,
        companies_table.c.email,
        companies_table.c.phone,
        companies_table.c.assigned_user_id,
    )
    if user_role == "staff":
        base_query = base_query.where(companies_table.c.assigned_user_id == user_id)
    elif user_role == "owner":
        base_query = base_query
    else:
        base_query = base_query.where(companies_table.c.agency_id == user_id)
    with engine.connect() as conn:
        rows = conn.execute(base_query).mappings().all()
    companies = [
        {
            "id": row["id"],
            "display_name": row["display_name"],
            "legal_name": row["legal_name"],
            "tax_id": row["tax_id"],
            "company_type": row["company_type"],
            "email": row["email"],
            "phone": row["phone"],
            "assigned_user_id": row["assigned_user_id"],
        }
        for row in rows
    ]
    return jsonify({"companies": companies})


@app.route("/api/companies", methods=["POST"])
def create_company():
    user_id = get_current_user_id()
    if (g.current_user or {}).get("role") == "staff":
        return jsonify({"ok": False, "errors": ["No autorizado."]}), 403
    payload = request.get_json(silent=True) or {}

    display_name = (payload.get("display_name") or payload.get("displayName") or "").strip()
    legal_name = (payload.get("legal_name") or payload.get("legalName") or "").strip()
    tax_id = (payload.get("tax_id") or payload.get("taxId") or "").strip().upper()
    company_type = payload.get("company_type") or payload.get("companyType") or ""
    email = (payload.get("email") or "").strip()
    phone = (payload.get("phone") or "").strip()
    assigned_user_id = payload.get("assigned_user_id") or payload.get("assignedUserId")

    errors = []
    if not display_name:
        errors.append("Nombre comercial obligatorio.")
    if not legal_name:
        errors.append("Razón social obligatoria.")
    if company_type not in {"individual", "company"}:
        errors.append("Tipo de empresa inválido.")
    if not validate_tax_id(tax_id, company_type):
        errors.append("CIF/NIF inválido.")

    if errors:
        return jsonify({"ok": False, "errors": errors}), 400

    with engine.connect() as conn:
        existing_count = conn.execute(
            select(func.count())
            .select_from(companies_table)
            .where(companies_table.c.agency_id == user_id)
        ).scalar_one()
        exists = conn.execute(
            select(companies_table.c.id)
            .where(companies_table.c.agency_id == user_id)
            .where(companies_table.c.tax_id == tax_id)
        ).first()
    if exists:
        return jsonify({"ok": False, "errors": ["Ya existe una empresa con ese CIF/NIF."]}), 400

    assigned_user_id = resolve_assigned_staff(user_id, assigned_user_id)

    created_at = datetime.utcnow().isoformat()
    with engine.begin() as conn:
        result = conn.execute(
            companies_table.insert().values(
                user_id=user_id,
                agency_id=user_id,
                display_name=display_name,
                legal_name=legal_name,
                tax_id=tax_id,
                company_type=company_type,
                email=email,
                phone=phone,
                assigned_user_id=assigned_user_id,
                created_at=created_at,
            )
        )
        new_id = result.inserted_primary_key[0]
        if existing_count == 0:
            conn.execute(
                invoices_table.update()
                .where(invoices_table.c.user_id == user_id)
                .where(invoices_table.c.company_id.is_(None))
                .values(company_id=new_id)
            )
            conn.execute(
                no_invoice_table.update()
                .where(no_invoice_table.c.user_id == user_id)
                .where(no_invoice_table.c.company_id.is_(None))
                .values(company_id=new_id)
            )
            conn.execute(
                facturacion_table.update()
                .where(facturacion_table.c.user_id == user_id)
                .where(facturacion_table.c.company_id.is_(None))
                .values(company_id=new_id)
            )
            conn.execute(
                income_invoices_table.update()
                .where(income_invoices_table.c.user_id == user_id)
                .where(income_invoices_table.c.company_id.is_(None))
                .values(company_id=new_id)
            )

    return jsonify({"ok": True, "id": new_id})


@app.route("/api/companies/<int:company_id>", methods=["PUT"])
def update_company(company_id):
    user_id = get_current_user_id()
    if (g.current_user or {}).get("role") == "staff":
        return jsonify({"ok": False, "errors": ["No autorizado."]}), 403
    payload = request.get_json(silent=True) or {}

    display_name = (payload.get("display_name") or payload.get("displayName") or "").strip()
    legal_name = (payload.get("legal_name") or payload.get("legalName") or "").strip()
    tax_id = (payload.get("tax_id") or payload.get("taxId") or "").strip().upper()
    company_type = payload.get("company_type") or payload.get("companyType") or ""
    email = (payload.get("email") or "").strip()
    phone = (payload.get("phone") or "").strip()
    assigned_user_id = payload.get("assigned_user_id") or payload.get("assignedUserId")

    errors = []
    if not display_name:
        errors.append("Nombre comercial obligatorio.")
    if not legal_name:
        errors.append("Razón social obligatoria.")
    if company_type not in {"individual", "company"}:
        errors.append("Tipo de empresa inválido.")
    if not validate_tax_id(tax_id, company_type):
        errors.append("CIF/NIF inválido.")

    if errors:
        return jsonify({"ok": False, "errors": errors}), 400

    with engine.connect() as conn:
        exists = conn.execute(
            select(companies_table.c.id)
            .where(companies_table.c.agency_id == user_id)
            .where(companies_table.c.tax_id == tax_id)
            .where(companies_table.c.id != company_id)
        ).first()
    if exists:
        return jsonify({"ok": False, "errors": ["Ya existe una empresa con ese CIF/NIF."]}), 400

    assigned_user_id = resolve_assigned_staff(user_id, assigned_user_id)

    with engine.begin() as conn:
        result = conn.execute(
            companies_table.update()
            .where(companies_table.c.id == company_id)
            .where(companies_table.c.agency_id == user_id)
            .values(
                display_name=display_name,
                legal_name=legal_name,
                tax_id=tax_id,
                company_type=company_type,
                email=email,
                phone=phone,
                assigned_user_id=assigned_user_id,
            )
        )

    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Empresa no encontrada."]}), 404

    return jsonify({"ok": True})


@app.route("/api/companies/<int:company_id>", methods=["DELETE"])
def delete_company(company_id):
    user_id = get_current_user_id()
    if (g.current_user or {}).get("role") == "staff":
        return jsonify({"ok": False, "errors": ["No autorizado."]}), 403
    with engine.begin() as conn:
        result = conn.execute(
            companies_table.delete()
            .where(companies_table.c.id == company_id)
            .where(companies_table.c.agency_id == user_id)
        )
    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Empresa no encontrada."]}), 404
    return jsonify({"ok": True})


@app.route("/api/staff")
def list_staff():
    user_id = get_current_user_id()
    role = (g.current_user or {}).get("role")
    if role == "staff":
        return jsonify({"ok": False, "errors": ["No autorizado."]}), 403
    with engine.connect() as conn:
        rows = conn.execute(
            select(
                users_table.c.id,
                users_table.c.email,
                users_table.c.role,
                users_table.c.is_active,
            )
            .where(users_table.c.agency_id == user_id)
            .where(users_table.c.role == "staff")
        ).mappings().all()
    staff = [
        {
            "id": row["id"],
            "email": row["email"],
            "is_active": bool(row["is_active"]),
        }
        for row in rows
    ]
    return jsonify({"staff": staff})


@app.route("/api/staff", methods=["POST"])
def create_staff():
    user_id = get_current_user_id()
    role = (g.current_user or {}).get("role")
    if role == "staff":
        return jsonify({"ok": False, "errors": ["No autorizado."]}), 403
    payload = request.get_json(silent=True) or {}
    email = _normalize_email(payload.get("email"))
    password = payload.get("password") or ""
    if not email or not password:
        return jsonify({"ok": False, "errors": ["Email y contraseña obligatorios."]}), 400
    if len(password) < 8:
        return jsonify({"ok": False, "errors": ["La contraseña debe tener al menos 8 caracteres."]}), 400

    with engine.begin() as conn:
        exists = conn.execute(
            select(users_table.c.id).where(users_table.c.email == email)
        ).first()
        if exists:
            return jsonify({"ok": False, "errors": ["El email ya está registrado."]}), 400
        agency_plan = conn.execute(
            select(users_table.c.plan).where(users_table.c.id == user_id)
        ).scalar_one_or_none()
        result = conn.execute(
            users_table.insert().values(
                email=email,
                password_hash=generate_password_hash(password),
                role="staff",
                plan=agency_plan or "trial",
                agency_id=user_id,
                created_at=datetime.utcnow().isoformat(),
                is_active=True,
            )
        )
        staff_id = result.inserted_primary_key[0]
    return jsonify({"ok": True, "id": staff_id})


@app.route("/api/staff/<int:staff_id>", methods=["PUT"])
def update_staff(staff_id):
    user_id = get_current_user_id()
    role = (g.current_user or {}).get("role")
    if role == "staff":
        return jsonify({"ok": False, "errors": ["No autorizado."]}), 403
    payload = request.get_json(silent=True) or {}
    password = payload.get("password")
    is_active = payload.get("is_active")
    updates = {}
    if password:
        if len(password) < 8:
            return jsonify({"ok": False, "errors": ["La contraseña debe tener al menos 8 caracteres."]}), 400
        updates["password_hash"] = generate_password_hash(password)
    if is_active is not None:
        updates["is_active"] = bool(is_active)
    if not updates:
        return jsonify({"ok": False, "errors": ["Nada que actualizar."]}), 400
    with engine.begin() as conn:
        result = conn.execute(
            users_table.update()
            .where(users_table.c.id == staff_id)
            .where(users_table.c.agency_id == user_id)
            .where(users_table.c.role == "staff")
            .values(**updates)
        )
    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Usuario no encontrado."]}), 404
    return jsonify({"ok": True})


@app.route("/api/years")
def available_years():
    years = set()
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=False)
    if company_id is None:
        return jsonify({"years": [date.today().year]})
    with engine.connect() as conn:
        invoice_dates = conn.execute(
            select(invoices_table.c.invoice_date)
            .where(invoices_table.c.user_id == data_owner_id)
            .where(invoices_table.c.company_id == company_id)
        ).scalars().all()
        for value in invoice_dates:
            if value:
                try:
                    years.add(int(str(value)[:4]))
                except ValueError:
                    continue
        billing_years = conn.execute(
            select(facturacion_table.c.anio)
            .where(facturacion_table.c.user_id == data_owner_id)
            .where(facturacion_table.c.company_id == company_id)
        ).scalars().all()
        years.update(int(year) for year in billing_years if year)
        income_years = conn.execute(
            select(income_invoices_table.c.invoice_date)
            .where(income_invoices_table.c.user_id == data_owner_id)
            .where(income_invoices_table.c.company_id == company_id)
        ).scalars().all()
        for value in income_years:
            if value:
                try:
                    years.add(int(str(value)[:4]))
                except ValueError:
                    continue

    if not years:
        years = {date.today().year}

    return jsonify({"years": sorted(years)})


@app.route("/api/upload", methods=["POST"])
def upload_invoices():
    data_owner_id = get_data_owner_id()
    if request.is_json:
        payload = request.get_json(silent=True) or {}
        entries = payload.get("entries", [])
        if not entries:
            return jsonify({"ok": False, "errors": ["No se recibieron entradas."]}), 400
        company_id = get_company_id(required=True)
        if company_id is None:
            return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400

        errors = []
        inserted = 0
        with engine.begin() as conn:
            for idx, entry in enumerate(entries):
                original_name = entry.get("originalFilename") or ""
                stored_name = entry.get("storedFilename") or ""
                invoice_date = entry.get("date") or date.today().isoformat()
                entry_company_id = entry.get("company_id") or entry.get("companyId")
                if entry_company_id:
                    try:
                        company_id = int(entry_company_id)
                    except (TypeError, ValueError):
                        errors.append("Empresa inválida.")
                        continue
                    if not is_company_accessible(company_id):
                        errors.append("Empresa inválida.")
                        continue
                supplier = (entry.get("supplier") or "").strip()
                base_amount = parse_amount(str(entry.get("base") or ""))
                vat_rate_raw = str(entry.get("vat") or "").strip()
                vat_amount = parse_amount(str(entry.get("vatAmount") or ""))
                total_amount = parse_amount(str(entry.get("total") or ""))
                payment_date = compute_payment_date(
                    invoice_date,
                    entry.get("paymentDate") or entry.get("payment_date"),
                )
                analysis_text = entry.get("analysisText") or entry.get("ocrText")
                expense_category = entry.get("expenseCategory") or "with_invoice"

                if not stored_name:
                    errors.append(f"Archivo faltante en posición {idx + 1}.")
                    continue
                if not supplier:
                    errors.append(f"Proveedor obligatorio para {original_name}.")
                    continue
                if base_amount is None or base_amount < 0:
                    errors.append(f"Base imponible inválida para {original_name}.")
                    continue
                try:
                    vat_rate_int = int(vat_rate_raw)
                except ValueError:
                    errors.append(f"Tipo de IVA inválido para {original_name}.")
                    continue
                if vat_rate_int not in {0, 4, 10, 21}:
                    errors.append(f"Tipo de IVA inválido para {original_name}.")
                    continue
                if expense_category not in {"with_invoice", "without_invoice", "non_deductible"}:
                    errors.append(f"Tipo de gasto inválido para {original_name}.")
                    continue

                if vat_amount is None:
                    vat_amount = round(base_amount * (vat_rate_int / 100), 2)
                if total_amount is None:
                    total_amount = round(base_amount + vat_amount, 2)

                created_at = datetime.utcnow().isoformat()

                stored_value = (
                    stored_name
                    if stored_name.startswith("http")
                    else get_public_url(stored_name)
                )

                conn.execute(
                    invoices_table.insert().values(
                        user_id=data_owner_id,
                        company_id=company_id,
                        original_filename=original_name,
                        stored_filename=stored_value,
                        invoice_date=invoice_date,
                        supplier=supplier,
                        base_amount=base_amount,
                        vat_rate=vat_rate_int,
                        vat_amount=vat_amount,
                        total_amount=total_amount,
                        payment_date=payment_date,
                        ocr_text=analysis_text,
                        expense_category=expense_category,
                        created_at=created_at,
                    )
                )
                inserted += 1

        return jsonify({"ok": True, "inserted": inserted, "errors": errors})

    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400

    files = request.files.getlist("files")
    dates = request.form.getlist("date")
    suppliers = request.form.getlist("supplier")
    bases = request.form.getlist("base")
    vats = request.form.getlist("vat")
    vat_amounts = request.form.getlist("vatAmount")
    totals = request.form.getlist("total")
    payment_dates = request.form.getlist("paymentDate")

    if not files:
        return jsonify({"ok": False, "errors": ["No se recibieron archivos."]}), 400

    if not (
        len(files)
        == len(dates)
        == len(suppliers)
        == len(bases)
        == len(vats)
        == len(vat_amounts)
        == len(totals)
        == len(payment_dates)
    ):
        return (
            jsonify({"ok": False, "errors": ["Los datos no coinciden con los archivos."]}),
            400,
        )

    errors = []
    inserted = 0
    with engine.begin() as conn:
        for idx, file in enumerate(files):
            if not file or not file.filename:
                errors.append(f"Archivo vacío en posición {idx + 1}.")
                continue

            original_name = os.path.basename(file.filename)
            if not allowed_file(original_name):
                errors.append(f"Tipo de archivo no permitido: {original_name}")
                continue

            invoice_date = dates[idx] or date.today().isoformat()
            payment_date = compute_payment_date(invoice_date, payment_dates[idx] if payment_dates else None)
            supplier = suppliers[idx].strip() if suppliers[idx] else ""
            base_amount = parse_amount(bases[idx])
            vat_rate = vats[idx].strip() if vats[idx] else ""
            vat_amount = parse_amount(vat_amounts[idx])
            total_amount = parse_amount(totals[idx])

            if not supplier:
                errors.append(f"Proveedor obligatorio para {original_name}.")
                continue
            if base_amount is None or base_amount < 0:
                errors.append(f"Base imponible inválida para {original_name}.")
                continue
            try:
                vat_rate_int = int(vat_rate)
            except ValueError:
                errors.append(f"Tipo de IVA inválido para {original_name}.")
                continue
            if vat_rate_int not in {0, 4, 10, 21}:
                errors.append(f"Tipo de IVA inválido para {original_name}.")
                continue

            file_bytes = file.read()
            safe_name = secure_filename(original_name)
            stored_name = f"{uuid4().hex}_{safe_name}"
            try:
                storage_url = upload_bytes(file_bytes, stored_name, file.mimetype)
            except Exception:
                app.logger.exception("Fallo al subir archivo a storage (%s)", stored_name)
                errors.append(f"No se pudo almacenar el archivo {original_name}.")
                continue

            if vat_amount is None:
                vat_amount = round(base_amount * (vat_rate_int / 100), 2)
            if total_amount is None:
                total_amount = round(base_amount + vat_amount, 2)
            created_at = datetime.utcnow().isoformat()

            conn.execute(
                invoices_table.insert().values(
                    user_id=data_owner_id,
                    company_id=company_id,
                    original_filename=original_name,
                    stored_filename=storage_url,
                    invoice_date=invoice_date,
                    supplier=supplier,
                    base_amount=base_amount,
                    vat_rate=vat_rate_int,
                    vat_amount=vat_amount,
                    total_amount=total_amount,
                    payment_date=payment_date,
                    ocr_text=None,
                    expense_category="with_invoice",
                    created_at=created_at,
                )
            )
            inserted += 1

    return jsonify({"ok": True, "inserted": inserted, "errors": errors})


@app.route("/api/analyze-invoice", methods=["POST"])
def analyze_invoice_api():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"ok": False, "errors": ["Archivo no recibido."]}), 400

    original_name = os.path.basename(file.filename)
    if not allowed_file(original_name):
        return jsonify({"ok": False, "errors": ["Tipo de archivo no permitido."]}), 400

    document_type = request.form.get("document_type") or request.args.get("document_type") or "expense"
    company_id = get_company_id(required=False)
    app.logger.info("Solicitud de análisis recibida: %s (%s)", original_name, file.mimetype)
    file_bytes = file.read()
    safe_name = secure_filename(original_name)
    stored_name = f"{uuid4().hex}_{safe_name}"
    try:
        storage_url = upload_bytes(file_bytes, stored_name, file.mimetype)
    except Exception:
        app.logger.exception("Fallo al subir archivo a storage (%s)", stored_name)
        return jsonify({"ok": False, "errors": ["No se pudo almacenar el archivo."]}), 500

    extracted = _analyze_invoice_with_timeout(
        file_bytes,
        original_name,
        stored_name,
        file.mimetype,
        document_type=document_type,
    )

    app.logger.info(
        "AI extracted for %s: provider=%s date=%s payment=%s base=%s vat_rate=%s vat_amount=%s total=%s",
        stored_name,
        extracted.get("provider_name"),
        extracted.get("invoice_date"),
        extracted.get("payment_date"),
        extracted.get("base_amount"),
        extracted.get("vat_rate"),
        extracted.get("vat_amount"),
        extracted.get("total_amount"),
    )

    return jsonify(
        {
            "ok": True,
            "storedFilename": storage_url,
            "originalFilename": original_name,
            "companyId": company_id,
            "extracted": extracted,
        }
    )


@app.route("/api/billing", methods=["POST"])
def create_billing():
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400

    payload = request.get_json(silent=True) or request.form

    month = int(payload.get("month") or 0)
    year = int(payload.get("year") or 0)
    base_amount = parse_amount(str(payload.get("base") or ""))
    vat_rate_raw = str(payload.get("vat") or "").strip()
    concept = (payload.get("concept") or "").strip()
    invoice_date = payload.get("invoice_date") or payload.get("date") or ""

    errors = []
    if month < 1 or month > 12:
        errors.append("Mes inválido.")
    if year < 2000:
        errors.append("Año inválido.")
    if invoice_date:
        normalized_date = normalize_date(invoice_date)
        if normalized_date is None:
            errors.append("Fecha inválida.")
        else:
            invoice_date = normalized_date
    if base_amount is None or base_amount < 0:
        errors.append("Base facturada inválida.")
    try:
        vat_rate = int(vat_rate_raw)
    except ValueError:
        vat_rate = None
    if vat_rate not in {0, 4, 10, 21}:
        errors.append("Tipo de IVA inválido.")

    if errors:
        return jsonify({"ok": False, "errors": errors}), 400

    iva_repercutido = round(base_amount * (vat_rate / 100), 2)
    total_amount = round(base_amount + iva_repercutido, 2)
    if invoice_date:
        try:
            month = int(invoice_date[5:7])
            year = int(invoice_date[:4])
        except (TypeError, ValueError):
            pass

    with engine.begin() as conn:
        conn.execute(
            facturacion_table.insert().values(
                user_id=data_owner_id,
                company_id=company_id,
                mes=month,
                anio=year,
                invoice_date=invoice_date or None,
                concept=concept or None,
                base_facturada=base_amount,
                tipo_iva=vat_rate,
                iva_repercutido=iva_repercutido,
                total_amount=total_amount,
            )
        )

    return jsonify({"ok": True})


@app.route("/api/billing/summary")
def billing_summary():
    month = request.args.get("month", type=int)
    year = request.args.get("year", type=int)
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400

    today = date.today()
    month = month or today.month
    year = year or today.year

    with engine.connect() as conn:
        rows = conn.execute(
            select(
                facturacion_table.c.tipo_iva,
                func.sum(facturacion_table.c.base_facturada).label("base_total"),
                func.sum(facturacion_table.c.iva_repercutido).label("vat_total"),
            )
            .where(
                facturacion_table.c.mes == month,
                facturacion_table.c.anio == year,
                facturacion_table.c.user_id == data_owner_id,
                facturacion_table.c.company_id == company_id,
            )
            .group_by(facturacion_table.c.tipo_iva)
        ).mappings().all()

    base_totals = {0: 0.0, 4: 0.0, 10: 0.0, 21: 0.0}
    vat_totals = {0: 0.0, 4: 0.0, 10: 0.0, 21: 0.0}

    for row in rows:
        vat_rate = int(row["tipo_iva"])
        base_totals[vat_rate] = float(row["base_total"] or 0)
        vat_totals[vat_rate] = float(row["vat_total"] or 0)

    total_vat = round(sum(vat_totals.values()), 2)

    return jsonify(
        {
            "baseTotals": {
                "0": round(base_totals[0], 2),
                "4": round(base_totals[4], 2),
                "10": round(base_totals[10], 2),
                "21": round(base_totals[21], 2),
            },
            "vatTotals": {
                "0": round(vat_totals[0], 2),
                "4": round(vat_totals[4], 2),
                "10": round(vat_totals[10], 2),
                "21": round(vat_totals[21], 2),
            },
            "totalVat": total_vat,
        }
    )


@app.route("/api/invoices")
def list_invoices():
    month = request.args.get("month", type=int)
    year = request.args.get("year", type=int)
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400

    today = date.today()
    month = month or today.month
    year = year or today.year

    _, last_day = calendar.monthrange(year, month)
    start = date(year, month, 1).isoformat()
    end = date(year, month, last_day).isoformat()

    with engine.connect() as conn:
        rows = conn.execute(
            select(
                invoices_table.c.id,
                invoices_table.c.invoice_date,
                invoices_table.c.supplier,
                invoices_table.c.base_amount,
                invoices_table.c.vat_rate,
                invoices_table.c.vat_amount,
                invoices_table.c.total_amount,
                invoices_table.c.payment_date,
                invoices_table.c.original_filename,
                invoices_table.c.expense_category,
            )
            .where(invoices_table.c.user_id == data_owner_id)
            .where(invoices_table.c.company_id == company_id)
            .where(invoices_table.c.invoice_date.between(start, end))
            .order_by(invoices_table.c.invoice_date.desc(), invoices_table.c.id.desc())
        ).mappings().all()

    invoices = [
        {
            "id": row["id"],
            "invoice_date": row["invoice_date"],
            "payment_date": row["payment_date"]
            or compute_payment_date(row["invoice_date"], row["payment_date"]),
            "supplier": row["supplier"],
            "base_amount": float(row["base_amount"]),
            "vat_rate": int(row["vat_rate"]),
            "vat_amount": float(row["vat_amount"]) if row["vat_amount"] is not None else None,
            "total_amount": float(row["total_amount"]),
            "original_filename": row["original_filename"],
            "expense_category": row["expense_category"] or "with_invoice",
        }
        for row in rows
    ]

    return jsonify({"invoices": invoices})


@app.route("/api/payments")
def list_payments():
    month = request.args.get("month", type=int)
    year = request.args.get("year", type=int)
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400

    today = date.today()
    month = month or today.month
    year = year or today.year

    _, last_day = calendar.monthrange(year, month)
    start = date(year, month, 1)
    end = date(year, month, last_day)
    year_start = date(year, 1, 1)
    year_end = date(year, 12, 31)
    buffer_start = (year_start - timedelta(days=31)).isoformat()
    year_start_iso = year_start.isoformat()
    year_end_iso = year_end.isoformat()

    with engine.connect() as conn:
        expense_rows = conn.execute(
            select(
                invoices_table.c.id,
                invoices_table.c.invoice_date,
                invoices_table.c.payment_date,
                invoices_table.c.supplier,
                invoices_table.c.total_amount,
                invoices_table.c.original_filename,
            )
            .where(
                (
                    invoices_table.c.payment_date.between(year_start_iso, year_end_iso)
                )
                | (
                    invoices_table.c.payment_date.is_(None)
                    & invoices_table.c.invoice_date.between(buffer_start, year_end_iso)
                )
            )
            .where(invoices_table.c.user_id == data_owner_id)
            .where(invoices_table.c.company_id == company_id)
            .order_by(invoices_table.c.invoice_date.desc(), invoices_table.c.id.desc())
        ).mappings().all()

        income_rows = conn.execute(
            select(
                income_invoices_table.c.id,
                income_invoices_table.c.invoice_date,
                income_invoices_table.c.payment_date,
                income_invoices_table.c.client,
                income_invoices_table.c.total_amount,
                income_invoices_table.c.original_filename,
            )
            .where(
                (
                    income_invoices_table.c.payment_date.between(year_start_iso, year_end_iso)
                )
                | (
                    income_invoices_table.c.payment_date.is_(None)
                    & income_invoices_table.c.invoice_date.between(buffer_start, year_end_iso)
                )
            )
            .where(income_invoices_table.c.user_id == data_owner_id)
            .where(income_invoices_table.c.company_id == company_id)
            .order_by(income_invoices_table.c.invoice_date.desc(), income_invoices_table.c.id.desc())
        ).mappings().all()

    items = []
    day_totals = {}
    for row in expense_rows:
        payment_date = row["payment_date"] or compute_payment_date(row["invoice_date"], None)
        if not payment_date:
            continue
        try:
            payment_dt = date.fromisoformat(payment_date)
        except ValueError:
            continue
        if payment_dt < start or payment_dt > end:
            continue
        day = payment_dt.day
        amount = float(row["total_amount"] or 0)
        day_totals[day] = round(day_totals.get(day, 0.0) + amount, 2)
        items.append(
            {
                "id": row["id"],
                "counterparty": row["supplier"],
                "concept": row["original_filename"],
                "payment_date": payment_date,
                "amount": amount,
                "type": "expense",
            }
        )

    for row in income_rows:
        payment_date = row["payment_date"] or compute_payment_date(row["invoice_date"], None)
        if not payment_date:
            continue
        try:
            payment_dt = date.fromisoformat(payment_date)
        except ValueError:
            continue
        if payment_dt < start or payment_dt > end:
            continue
        day = payment_dt.day
        amount = float(row["total_amount"] or 0)
        day_totals[day] = round(day_totals.get(day, 0.0) + amount, 2)
        items.append(
            {
                "id": row["id"],
                "counterparty": row["client"],
                "concept": row["original_filename"],
                "payment_date": payment_date,
                "amount": amount,
                "type": "income",
            }
        )

    return jsonify({"items": items, "dayTotals": day_totals})


@app.route("/api/reports/quarterly")
def quarterly_report():
    data_owner_id = get_data_owner_id()
    user_role = (g.current_user or {}).get("role")
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400
    year, quarter, start_month, end_month = _parse_period_params()
    months = _get_months_for_period(year, quarter, start_month, end_month)
    if not months:
        return jsonify({"ok": False, "errors": ["Periodo inválido."]}), 400

    with engine.connect() as conn:
        company_query = select(
            companies_table.c.display_name,
            companies_table.c.legal_name,
            companies_table.c.tax_id,
            companies_table.c.company_type,
            companies_table.c.email,
        ).where(companies_table.c.id == company_id)
        if user_role != "owner":
            company_query = company_query.where(companies_table.c.agency_id == data_owner_id)
        company = conn.execute(
            company_query
        ).mappings().first()
    if not company:
        return jsonify({"ok": False, "errors": ["Empresa no encontrada."]}), 404

    totals = _build_report_totals(data_owner_id, company_id, months, year)
    period_label = _report_period_label(year, months, quarter)
    generated_at = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
    html = f"""
    <html>
      <head>
        <meta charset="utf-8" />
        <title>Informe trimestral {period_label}</title>
        <style>
          body {{ font-family: Arial, sans-serif; margin: 32px; color: #1f2937; }}
          h1 {{ font-size: 20px; margin-bottom: 4px; }}
          p {{ margin: 4px 0; }}
          table {{ border-collapse: collapse; width: 100%; margin-top: 16px; }}
          th, td {{ border: 1px solid #e5e7eb; padding: 8px; text-align: left; }}
          th {{ background: #f9fafb; }}
        </style>
      </head>
      <body>
        <h1>Informe fiscal {period_label}</h1>
        <p><strong>Empresa:</strong> {company["display_name"]} ({company["legal_name"]})</p>
        <p><strong>CIF/NIF:</strong> {company["tax_id"]}</p>
        <p><strong>Generado:</strong> {generated_at}</p>
        <table>
          <tr><th>Concepto</th><th>Importe (€)</th></tr>
          <tr><td>Ingresos base</td><td>{totals["income_base"]:.2f}</td></tr>
          <tr><td>IVA repercutido</td><td>{totals["income_vat"]:.2f}</td></tr>
          <tr><td>Gastos deducibles</td><td>{totals["expense_base"]:.2f}</td></tr>
          <tr><td>IVA soportado</td><td>{totals["expense_vat"]:.2f}</td></tr>
          <tr><td>Resultado neto</td><td>{totals["net_result"]:.2f}</td></tr>
          <tr><td>Resultado IVA</td><td>{totals["vat_result"]:.2f}</td></tr>
        </table>
      </body>
    </html>
    """
    response = app.response_class(html, mimetype="text/html")
    filename = f"informe_{period_label.replace(' ', '_')}.html"
    response.headers["Content-Disposition"] = f"attachment; filename={filename}"
    return response


@app.route("/api/reports/quarterly/email", methods=["POST"])
def quarterly_report_email():
    data_owner_id = get_data_owner_id()
    user_role = (g.current_user or {}).get("role")
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400
    payload = request.get_json(silent=True) or {}
    year = payload.get("year")
    quarter = payload.get("quarter")
    start_month = payload.get("start_month")
    end_month = payload.get("end_month")
    try:
        year = int(year)
    except (TypeError, ValueError):
        year = None
    try:
        quarter = int(quarter) if quarter else None
    except (TypeError, ValueError):
        quarter = None
    try:
        start_month = int(start_month) if start_month else None
    except (TypeError, ValueError):
        start_month = None
    try:
        end_month = int(end_month) if end_month else None
    except (TypeError, ValueError):
        end_month = None

    months = _get_months_for_period(year, quarter, start_month, end_month)
    if not months:
        return jsonify({"ok": False, "errors": ["Periodo inválido."]}), 400

    with engine.connect() as conn:
        company_query = select(
            companies_table.c.display_name,
            companies_table.c.legal_name,
            companies_table.c.tax_id,
            companies_table.c.company_type,
            companies_table.c.email,
        ).where(companies_table.c.id == company_id)
        if user_role != "owner":
            company_query = company_query.where(companies_table.c.agency_id == data_owner_id)
        company = conn.execute(company_query).mappings().first()
        user = conn.execute(
            select(users_table.c.email).where(users_table.c.id == get_current_user_id())
        ).mappings().first()
    if not company:
        return jsonify({"ok": False, "errors": ["Empresa no encontrada."]}), 404
    if not company.get("email"):
        return jsonify({"ok": False, "errors": ["La empresa no tiene email."]}), 400

    totals = _build_report_totals(data_owner_id, company_id, months, year)
    period_label = _report_period_label(year, months, quarter)
    html = f"""
    <h2>Informe fiscal {period_label}</h2>
    <p><strong>Empresa:</strong> {company["display_name"]} ({company["legal_name"]})</p>
    <p><strong>CIF/NIF:</strong> {company["tax_id"]}</p>
    <ul>
      <li>Ingresos base: {totals["income_base"]:.2f} €</li>
      <li>IVA repercutido: {totals["income_vat"]:.2f} €</li>
      <li>Gastos deducibles: {totals["expense_base"]:.2f} €</li>
      <li>IVA soportado: {totals["expense_vat"]:.2f} €</li>
      <li>Resultado neto: {totals["net_result"]:.2f} €</li>
      <li>Resultado IVA: {totals["vat_result"]:.2f} €</li>
    </ul>
    """
    reply_to = user["email"] if user else None
    sent = send_email(
        company["email"],
        f"Informe fiscal {period_label}",
        html,
        reply_to=reply_to,
    )
    if not sent:
        return jsonify({"ok": False, "errors": ["No se pudo enviar el email."]}), 500
    return jsonify({"ok": True})


@app.route("/api/invoices/<int:invoice_id>", methods=["PUT"])
def update_invoice(invoice_id):
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400
    payload = request.get_json(silent=True) or {}

    invoice_date = payload.get("invoice_date") or ""
    payment_date = compute_payment_date(
        invoice_date,
        payload.get("payment_date") or payload.get("paymentDate"),
    )
    supplier = (payload.get("supplier") or "").strip()
    base_amount = parse_amount(str(payload.get("base_amount") or ""))
    vat_rate_raw = str(payload.get("vat_rate") or "").strip()
    vat_amount = parse_amount(str(payload.get("vat_amount") or ""))
    total_amount = parse_amount(str(payload.get("total_amount") or ""))
    expense_category = payload.get("expense_category") or "with_invoice"

    errors = []
    if not invoice_date:
        errors.append("Fecha obligatoria.")
    if not supplier:
        errors.append("Proveedor obligatorio.")
    if base_amount is None or base_amount < 0:
        errors.append("Base imponible inválida.")
    try:
        vat_rate = int(vat_rate_raw)
    except ValueError:
        vat_rate = None
    if vat_rate not in {0, 4, 10, 21}:
        errors.append("Tipo de IVA inválido.")
    if expense_category not in {"with_invoice", "without_invoice", "non_deductible"}:
        errors.append("Tipo de gasto inválido.")

    if errors:
        return jsonify({"ok": False, "errors": errors}), 400

    if vat_amount is None:
        vat_amount = round(base_amount * (vat_rate / 100), 2)
    if total_amount is None:
        total_amount = round(base_amount + vat_amount, 2)

    with engine.begin() as conn:
        result = conn.execute(
            invoices_table.update()
            .where(invoices_table.c.id == invoice_id)
            .where(invoices_table.c.user_id == data_owner_id)
            .where(invoices_table.c.company_id == company_id)
            .values(
                invoice_date=invoice_date,
                supplier=supplier,
                base_amount=base_amount,
                vat_rate=vat_rate,
                vat_amount=vat_amount,
                total_amount=total_amount,
                payment_date=payment_date,
                expense_category=expense_category,
            )
        )

    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Factura no encontrada."]}), 404

    return jsonify(
        {
            "ok": True,
            "invoice": {
                "id": invoice_id,
                "invoice_date": invoice_date,
                "payment_date": payment_date,
                "supplier": supplier,
                "base_amount": base_amount,
                "vat_rate": vat_rate,
                "vat_amount": vat_amount,
                "total_amount": total_amount,
                "expense_category": expense_category,
            },
        }
    )


@app.route("/api/invoices/<int:invoice_id>", methods=["DELETE"])
def delete_invoice(invoice_id):
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400
    with engine.begin() as conn:
        result = conn.execute(
            invoices_table.delete()
            .where(invoices_table.c.id == invoice_id)
            .where(invoices_table.c.user_id == data_owner_id)
            .where(invoices_table.c.company_id == company_id)
        )

    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Factura no encontrada."]}), 404

    return jsonify({"ok": True})


@app.route("/api/income-invoices")
def list_income_invoices():
    month = request.args.get("month", type=int)
    year = request.args.get("year", type=int)
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400

    today = date.today()
    month = month or today.month
    year = year or today.year

    _, last_day = calendar.monthrange(year, month)
    start = date(year, month, 1).isoformat()
    end = date(year, month, last_day).isoformat()

    with engine.connect() as conn:
        rows = conn.execute(
            select(
                income_invoices_table.c.id,
                income_invoices_table.c.invoice_date,
                income_invoices_table.c.payment_date,
                income_invoices_table.c.client,
                income_invoices_table.c.base_amount,
                income_invoices_table.c.vat_rate,
                income_invoices_table.c.vat_amount,
                income_invoices_table.c.total_amount,
                income_invoices_table.c.original_filename,
            )
            .where(income_invoices_table.c.user_id == data_owner_id)
            .where(income_invoices_table.c.company_id == company_id)
            .where(income_invoices_table.c.invoice_date.between(start, end))
            .order_by(income_invoices_table.c.invoice_date.desc(), income_invoices_table.c.id.desc())
        ).mappings().all()

    invoices = [
        {
            "id": row["id"],
            "invoice_date": row["invoice_date"],
            "payment_date": row["payment_date"]
            or compute_payment_date(row["invoice_date"], row["payment_date"]),
            "client": row["client"],
            "base_amount": float(row["base_amount"]),
            "vat_rate": int(row["vat_rate"]),
            "vat_amount": float(row["vat_amount"]) if row["vat_amount"] is not None else None,
            "total_amount": float(row["total_amount"]),
            "original_filename": row["original_filename"],
        }
        for row in rows
    ]

    return jsonify({"invoices": invoices})


@app.route("/api/income-invoices", methods=["POST"])
def create_income_invoices():
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400

    payload = request.get_json(silent=True) or {}
    entries = payload.get("entries", [])
    if not entries:
        return jsonify({"ok": False, "errors": ["No se recibieron entradas."]}), 400

    errors = []
    inserted = 0
    with engine.begin() as conn:
        for entry in entries:
            original_name = entry.get("originalFilename") or ""
            stored_name = entry.get("storedFilename") or ""
            invoice_date = entry.get("date") or entry.get("invoice_date") or date.today().isoformat()
            client = (entry.get("client") or "").strip()
            base_amount = parse_amount(str(entry.get("base") or ""))
            vat_rate_raw = str(entry.get("vat") or "").strip()
            vat_amount = parse_amount(str(entry.get("vatAmount") or ""))
            total_amount = parse_amount(str(entry.get("total") or ""))
            payment_date = compute_payment_date(
                invoice_date,
                entry.get("paymentDate") or entry.get("payment_date"),
            )
            analysis_text = entry.get("analysisText") or entry.get("ocrText")

            if not stored_name:
                errors.append(f"Archivo faltante para {original_name}.")
                continue
            if not client:
                errors.append(f"Cliente obligatorio para {original_name}.")
                continue
            if base_amount is None or base_amount < 0:
                errors.append(f"Base imponible inválida para {original_name}.")
                continue
            try:
                vat_rate_int = int(vat_rate_raw)
            except ValueError:
                errors.append(f"Tipo de IVA inválido para {original_name}.")
                continue
            if vat_rate_int not in {0, 4, 10, 21}:
                errors.append(f"Tipo de IVA inválido para {original_name}.")
                continue

            if vat_amount is None:
                vat_amount = round(base_amount * (vat_rate_int / 100), 2)
            if total_amount is None:
                total_amount = round(base_amount + vat_amount, 2)

            stored_value = (
                stored_name
                if stored_name.startswith("http")
                else get_public_url(stored_name)
            )
            created_at = datetime.utcnow().isoformat()

            conn.execute(
                income_invoices_table.insert().values(
                    user_id=data_owner_id,
                    company_id=company_id,
                    original_filename=original_name,
                    stored_filename=stored_value,
                    invoice_date=invoice_date,
                    client=client,
                    base_amount=base_amount,
                    vat_rate=vat_rate_int,
                    vat_amount=vat_amount,
                    total_amount=total_amount,
                    payment_date=payment_date,
                    ocr_text=analysis_text,
                    created_at=created_at,
                )
            )
            inserted += 1

    return jsonify({"ok": True, "inserted": inserted, "errors": errors})


@app.route("/api/income-invoices/<int:invoice_id>", methods=["PUT"])
def update_income_invoice(invoice_id):
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400

    payload = request.get_json(silent=True) or {}
    invoice_date = payload.get("invoice_date") or ""
    payment_date = compute_payment_date(
        invoice_date,
        payload.get("payment_date") or payload.get("paymentDate"),
    )
    client = (payload.get("client") or "").strip()
    base_amount = parse_amount(str(payload.get("base_amount") or ""))
    vat_rate_raw = str(payload.get("vat_rate") or "").strip()
    vat_amount = parse_amount(str(payload.get("vat_amount") or ""))
    total_amount = parse_amount(str(payload.get("total_amount") or ""))

    errors = []
    if not invoice_date:
        errors.append("Fecha obligatoria.")
    if not client:
        errors.append("Cliente obligatorio.")
    if base_amount is None or base_amount < 0:
        errors.append("Base imponible inválida.")
    try:
        vat_rate = int(vat_rate_raw)
    except ValueError:
        vat_rate = None
    if vat_rate not in {0, 4, 10, 21}:
        errors.append("Tipo de IVA inválido.")

    if errors:
        return jsonify({"ok": False, "errors": errors}), 400

    if vat_amount is None:
        vat_amount = round(base_amount * (vat_rate / 100), 2)
    if total_amount is None:
        total_amount = round(base_amount + vat_amount, 2)

    with engine.begin() as conn:
        result = conn.execute(
            income_invoices_table.update()
            .where(income_invoices_table.c.id == invoice_id)
            .where(income_invoices_table.c.user_id == data_owner_id)
            .where(income_invoices_table.c.company_id == company_id)
            .values(
                invoice_date=invoice_date,
                payment_date=payment_date,
                client=client,
                base_amount=base_amount,
                vat_rate=vat_rate,
                vat_amount=vat_amount,
                total_amount=total_amount,
            )
        )

    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Factura no encontrada."]}), 404

    return jsonify({"ok": True})


@app.route("/api/income-invoices/<int:invoice_id>", methods=["DELETE"])
def delete_income_invoice(invoice_id):
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400

    with engine.begin() as conn:
        result = conn.execute(
            income_invoices_table.delete()
            .where(income_invoices_table.c.id == invoice_id)
            .where(income_invoices_table.c.user_id == data_owner_id)
            .where(income_invoices_table.c.company_id == company_id)
        )
    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Factura no encontrada."]}), 404
    return jsonify({"ok": True})


@app.route("/api/expenses/no-invoice")
def list_no_invoice_expenses():
    month = request.args.get("month", type=int)
    year = request.args.get("year", type=int)
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400

    today = date.today()
    month = month or today.month
    year = year or today.year

    _, last_day = calendar.monthrange(year, month)
    start = date(year, month, 1).isoformat()
    end = date(year, month, last_day).isoformat()

    with engine.connect() as conn:
        rows = conn.execute(
            select(
                no_invoice_table.c.id,
                no_invoice_table.c.expense_date,
                no_invoice_table.c.concept,
                no_invoice_table.c.amount,
                no_invoice_table.c.expense_type,
                no_invoice_table.c.deductible,
            )
            .where(no_invoice_table.c.user_id == data_owner_id)
            .where(no_invoice_table.c.company_id == company_id)
            .where(no_invoice_table.c.expense_date.between(start, end))
            .order_by(no_invoice_table.c.expense_date.desc(), no_invoice_table.c.id.desc())
        ).mappings().all()

    expenses = [
        {
            "id": row["id"],
            "expense_date": row["expense_date"],
            "concept": row["concept"],
            "amount": float(row["amount"]),
            "expense_type": row["expense_type"],
            "deductible": bool(row["deductible"]),
        }
        for row in rows
    ]

    return jsonify({"expenses": expenses})


@app.route("/api/expenses/no-invoice", methods=["POST"])
def create_no_invoice_expense():
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400
    payload = request.get_json(silent=True) or {}

    expense_date = payload.get("expense_date") or ""
    concept = (payload.get("concept") or "").strip()
    amount = parse_amount(str(payload.get("amount") or ""))
    expense_type = payload.get("expense_type") or ""
    deductible = payload.get("deductible")

    errors = []
    if not expense_date:
        errors.append("Fecha obligatoria.")
    if not concept:
        errors.append("Concepto obligatorio.")
    if amount is None or amount < 0:
        errors.append("Importe inválido.")
    if expense_type not in {
        "nomina",
        "seguridad_social",
        "amortizacion",
        "kilometraje",
        "otro",
    }:
        errors.append("Tipo de gasto inválido.")
    if deductible is None:
        deductible = True

    if errors:
        return jsonify({"ok": False, "errors": errors}), 400

    with engine.begin() as conn:
        conn.execute(
            no_invoice_table.insert().values(
                user_id=data_owner_id,
                company_id=company_id,
                expense_date=expense_date,
                concept=concept,
                amount=amount,
                expense_type=expense_type,
                deductible=bool(deductible),
                created_at=datetime.utcnow().isoformat(),
            )
        )

    return jsonify({"ok": True})


@app.route("/api/expenses/no-invoice/<int:expense_id>", methods=["PUT"])
def update_no_invoice_expense(expense_id):
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400
    payload = request.get_json(silent=True) or {}

    expense_date = payload.get("expense_date") or ""
    concept = (payload.get("concept") or "").strip()
    amount = parse_amount(str(payload.get("amount") or ""))
    expense_type = payload.get("expense_type") or ""
    deductible = payload.get("deductible")

    errors = []
    if not expense_date:
        errors.append("Fecha obligatoria.")
    if not concept:
        errors.append("Concepto obligatorio.")
    if amount is None or amount < 0:
        errors.append("Importe inválido.")
    if expense_type not in {
        "nomina",
        "seguridad_social",
        "amortizacion",
        "kilometraje",
        "otro",
    }:
        errors.append("Tipo de gasto inválido.")
    if deductible is None:
        deductible = True

    if errors:
        return jsonify({"ok": False, "errors": errors}), 400

    with engine.begin() as conn:
        result = conn.execute(
            no_invoice_table.update()
            .where(no_invoice_table.c.id == expense_id)
            .where(no_invoice_table.c.user_id == data_owner_id)
            .where(no_invoice_table.c.company_id == company_id)
            .values(
                expense_date=expense_date,
                concept=concept,
                amount=amount,
                expense_type=expense_type,
                deductible=bool(deductible),
            )
        )

    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Gasto no encontrado."]}), 404

    return jsonify(
        {
            "ok": True,
            "expense": {
                "id": expense_id,
                "expense_date": expense_date,
                "concept": concept,
                "amount": amount,
                "expense_type": expense_type,
                "deductible": bool(deductible),
            },
        }
    )


@app.route("/api/expenses/no-invoice/<int:expense_id>", methods=["DELETE"])
def delete_no_invoice_expense(expense_id):
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400
    with engine.begin() as conn:
        result = conn.execute(
            no_invoice_table.delete()
            .where(no_invoice_table.c.id == expense_id)
            .where(no_invoice_table.c.user_id == data_owner_id)
            .where(no_invoice_table.c.company_id == company_id)
        )

    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Gasto no encontrado."]}), 404

    return jsonify({"ok": True})


@app.route("/api/billing/entries")
def billing_entries():
    month = request.args.get("month", type=int)
    year = request.args.get("year", type=int)
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400

    today = date.today()
    month = month or today.month
    year = year or today.year

    with engine.connect() as conn:
        rows = conn.execute(
            select(
                facturacion_table.c.id,
                facturacion_table.c.mes,
                facturacion_table.c.anio,
                facturacion_table.c.invoice_date,
                facturacion_table.c.concept,
                facturacion_table.c.base_facturada,
                facturacion_table.c.tipo_iva,
                facturacion_table.c.iva_repercutido,
                facturacion_table.c.total_amount,
            )
            .where(
                facturacion_table.c.mes == month,
                facturacion_table.c.anio == year,
                facturacion_table.c.user_id == data_owner_id,
                facturacion_table.c.company_id == company_id,
            )
            .order_by(facturacion_table.c.id.desc())
        ).mappings().all()

    entries = [
        {
            "id": row["id"],
            "month": row["mes"],
            "year": row["anio"],
            "invoice_date": row["invoice_date"],
            "concept": row["concept"],
            "base": float(row["base_facturada"]),
            "vat": int(row["tipo_iva"]),
            "vatAmount": float(row["iva_repercutido"]),
            "total": float(row["total_amount"] or 0),
        }
        for row in rows
    ]

    return jsonify({"entries": entries})


@app.route("/api/billing/<int:billing_id>", methods=["PUT"])
def update_billing(billing_id):
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400
    payload = request.get_json(silent=True) or request.form

    base_amount = parse_amount(str(payload.get("base") or ""))
    vat_rate_raw = str(payload.get("vat") or "").strip()

    errors = []
    if base_amount is None or base_amount < 0:
        errors.append("Base facturada inválida.")
    try:
        vat_rate = int(vat_rate_raw)
    except ValueError:
        vat_rate = None
    if vat_rate not in {0, 4, 10, 21}:
        errors.append("Tipo de IVA inválido.")

    if errors:
        return jsonify({"ok": False, "errors": errors}), 400

    iva_repercutido = round(base_amount * (vat_rate / 100), 2)
    total_amount = round(base_amount + iva_repercutido, 2)

    with engine.begin() as conn:
        result = conn.execute(
            facturacion_table.update()
            .where(facturacion_table.c.id == billing_id)
            .where(facturacion_table.c.user_id == data_owner_id)
            .where(facturacion_table.c.company_id == company_id)
            .values(
                base_facturada=base_amount,
                tipo_iva=vat_rate,
                iva_repercutido=iva_repercutido,
                total_amount=total_amount,
            )
        )

    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Registro no encontrado."]}), 404

    return jsonify(
        {
            "ok": True,
            "entry": {
                "id": billing_id,
                "base": base_amount,
                "vat": vat_rate,
                "vatAmount": iva_repercutido,
                "total": total_amount,
            },
        }
    )


@app.route("/api/billing/<int:billing_id>", methods=["DELETE"])
def delete_billing(billing_id):
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400
    with engine.begin() as conn:
        result = conn.execute(
            facturacion_table.delete()
            .where(facturacion_table.c.id == billing_id)
            .where(facturacion_table.c.user_id == data_owner_id)
            .where(facturacion_table.c.company_id == company_id)
        )

    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Registro no encontrado."]}), 404

    return jsonify({"ok": True})


@app.route("/api/summary")
def summary():
    month = request.args.get("month", type=int)
    year = request.args.get("year", type=int)
    data_owner_id = get_data_owner_id()
    company_id = get_company_id(required=True)
    if company_id is None:
        return jsonify({"ok": False, "errors": ["Empresa no seleccionada."]}), 400

    today = date.today()
    month = month or today.month
    year = year or today.year

    _, last_day = calendar.monthrange(year, month)
    start = date(year, month, 1).isoformat()
    end = date(year, month, last_day).isoformat()

    with engine.connect() as conn:
        rows = conn.execute(
            select(
                invoices_table.c.invoice_date,
                invoices_table.c.supplier,
                invoices_table.c.total_amount,
                invoices_table.c.vat_rate,
                invoices_table.c.base_amount,
            )
            .where(invoices_table.c.user_id == data_owner_id)
            .where(invoices_table.c.company_id == company_id)
            .where(invoices_table.c.invoice_date.between(start, end))
            .order_by(invoices_table.c.invoice_date)
        ).mappings().all()

    daily_totals = {day: 0.0 for day in range(1, last_day + 1)}
    supplier_totals = {}
    vat_totals = {0: 0.0, 4: 0.0, 10: 0.0, 21: 0.0}
    total_spent = 0.0

    for row in rows:
        try:
            row_date = date.fromisoformat(row["invoice_date"])
        except ValueError:
            continue
        day = row_date.day
        amount = float(row["total_amount"])
        base_amount = float(row["base_amount"])
        vat_rate = int(row["vat_rate"])

        daily_totals[day] += amount
        total_spent += amount

        supplier = row["supplier"]
        supplier_totals[supplier] = supplier_totals.get(supplier, 0.0) + amount

        vat_totals[vat_rate] += base_amount * (vat_rate / 100)

    cumulative = []
    running = 0.0
    for day in range(1, last_day + 1):
        running += daily_totals[day]
        cumulative.append(round(running, 2))

    suppliers = list(supplier_totals.keys())
    supplier_values = [round(supplier_totals[name], 2) for name in suppliers]

    vat_total_deductible = round(sum(vat_totals.values()), 2)

    return jsonify(
        {
            "days": list(range(1, last_day + 1)),
            "cumulative": cumulative,
            "suppliers": suppliers,
            "supplierTotals": supplier_values,
            "totalSpent": round(total_spent, 2),
            "vatTotals": {
                "0": round(vat_totals[0], 2),
                "4": round(vat_totals[4], 2),
                "10": round(vat_totals[10], 2),
                "21": round(vat_totals[21], 2),
            },
            "vatTotalDeductible": vat_total_deductible,
        }
    )


init_db()

if __name__ == "__main__":
    app.run()
