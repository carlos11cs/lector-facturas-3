import calendar
import logging
import multiprocessing as mp
import os
from datetime import date, datetime
from uuid import uuid4

from flask import Flask, jsonify, render_template, request
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
    select,
)
from werkzeug.utils import secure_filename

from services.ai_invoice_service import analyze_invoice
from services.storage_service import get_public_url, upload_bytes

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "data.db")
ALLOWED_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png"}
ANALYSIS_TIMEOUT_SECONDS = int(os.getenv("ANALYSIS_TIMEOUT_SECONDS", "120"))

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

invoices_table = Table(
    "invoices",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("original_filename", String, nullable=False),
    Column("stored_filename", String, nullable=False),
    Column("invoice_date", String, nullable=False),
    Column("supplier", String, nullable=False),
    Column("base_amount", Float, nullable=False),
    Column("vat_rate", Integer, nullable=False),
    Column("vat_amount", Float),
    Column("total_amount", Float, nullable=False),
    Column("ocr_text", Text),
    Column("expense_category", String, nullable=False, server_default="with_invoice"),
    Column("created_at", String, nullable=False),
)

facturacion_table = Table(
    "facturacion",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("mes", Integer, nullable=False),
    Column("anio", Integer, nullable=False),
    Column("base_facturada", Float, nullable=False),
    Column("tipo_iva", Integer, nullable=False),
    Column("iva_repercutido", Float, nullable=False),
)

no_invoice_table = Table(
    "no_invoice_expenses",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("expense_date", String, nullable=False),
    Column("concept", String, nullable=False),
    Column("amount", Float, nullable=False),
    Column("expense_type", String, nullable=False),
    Column("deductible", Boolean, nullable=False),
    Column("created_at", String, nullable=False),
)

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)


def init_db():
    metadata.create_all(engine)


def allowed_file(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTENSIONS


def parse_amount(value):
    if value is None:
        return None
    cleaned = value.replace("EUR", "").replace("euro", "").strip()
    cleaned = cleaned.replace(".", "").replace(",", ".") if "," in cleaned else cleaned
    try:
        return float(cleaned)
    except ValueError:
        return None


def _empty_extracted():
    return {
        "provider_name": None,
        "invoice_date": None,
        "base_amount": None,
        "vat_rate": None,
        "vat_amount": None,
        "total_amount": None,
        "analysis_text": "",
        "validation": {"is_consistent": None, "difference": None},
    }


def _analysis_worker(file_bytes, filename, mime_type, queue):
    try:
        result = analyze_invoice(
            file_bytes=file_bytes,
            filename=filename,
            mime_type=mime_type,
        )
        queue.put(result)
    except Exception as exc:
        queue.put({"__error__": str(exc)})


def _analyze_invoice_with_timeout(file_bytes, filename, stored_name, mime_type):
    ctx = mp.get_context("spawn")
    queue = ctx.Queue(1)
    process = ctx.Process(
        target=_analysis_worker,
        args=(file_bytes, filename, mime_type, queue),
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


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/years")
def available_years():
    years = set()
    with engine.connect() as conn:
        invoice_dates = conn.execute(select(invoices_table.c.invoice_date)).scalars().all()
        for value in invoice_dates:
            if value:
                try:
                    years.add(int(str(value)[:4]))
                except ValueError:
                    continue
        billing_years = conn.execute(select(facturacion_table.c.anio)).scalars().all()
        years.update(int(year) for year in billing_years if year)

    if not years:
        years = {date.today().year}

    return jsonify({"years": sorted(years)})


@app.route("/api/upload", methods=["POST"])
def upload_invoices():
    if request.is_json:
        payload = request.get_json(silent=True) or {}
        entries = payload.get("entries", [])
        if not entries:
            return jsonify({"ok": False, "errors": ["No se recibieron entradas."]}), 400

        errors = []
        inserted = 0
        with engine.begin() as conn:
            for idx, entry in enumerate(entries):
                original_name = entry.get("originalFilename") or ""
                stored_name = entry.get("storedFilename") or ""
                invoice_date = entry.get("date") or date.today().isoformat()
                supplier = (entry.get("supplier") or "").strip()
                base_amount = parse_amount(str(entry.get("base") or ""))
                vat_rate_raw = str(entry.get("vat") or "").strip()
                vat_amount = parse_amount(str(entry.get("vatAmount") or ""))
                total_amount = parse_amount(str(entry.get("total") or ""))
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
                        original_filename=original_name,
                        stored_filename=stored_value,
                        invoice_date=invoice_date,
                        supplier=supplier,
                        base_amount=base_amount,
                        vat_rate=vat_rate_int,
                        vat_amount=vat_amount,
                        total_amount=total_amount,
                        ocr_text=analysis_text,
                        expense_category=expense_category,
                        created_at=created_at,
                    )
                )
                inserted += 1

        return jsonify({"ok": True, "inserted": inserted, "errors": errors})

    files = request.files.getlist("files")
    dates = request.form.getlist("date")
    suppliers = request.form.getlist("supplier")
    bases = request.form.getlist("base")
    vats = request.form.getlist("vat")
    vat_amounts = request.form.getlist("vatAmount")
    totals = request.form.getlist("total")

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
                    original_filename=original_name,
                    stored_filename=storage_url,
                    invoice_date=invoice_date,
                    supplier=supplier,
                    base_amount=base_amount,
                    vat_rate=vat_rate_int,
                    vat_amount=vat_amount,
                    total_amount=total_amount,
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

    file_bytes = file.read()
    safe_name = secure_filename(original_name)
    stored_name = f"{uuid4().hex}_{safe_name}"
    try:
        storage_url = upload_bytes(file_bytes, stored_name, file.mimetype)
    except Exception:
        app.logger.exception("Fallo al subir archivo a storage (%s)", stored_name)
        return jsonify({"ok": False, "errors": ["No se pudo almacenar el archivo."]}), 500

    extracted = _analyze_invoice_with_timeout(file_bytes, original_name, stored_name, file.mimetype)

    app.logger.info(
        "AI extracted for %s: provider=%s date=%s base=%s vat_rate=%s vat_amount=%s total=%s",
        stored_name,
        extracted.get("provider_name"),
        extracted.get("invoice_date"),
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
            "extracted": extracted,
        }
    )


@app.route("/api/billing", methods=["POST"])
def create_billing():
    payload = request.get_json(silent=True) or request.form

    month = int(payload.get("month") or 0)
    year = int(payload.get("year") or 0)
    base_amount = parse_amount(str(payload.get("base") or ""))
    vat_rate_raw = str(payload.get("vat") or "").strip()

    errors = []
    if month < 1 or month > 12:
        errors.append("Mes inválido.")
    if year < 2000:
        errors.append("Año inválido.")
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

    with engine.begin() as conn:
        conn.execute(
            facturacion_table.insert().values(
                mes=month,
                anio=year,
                base_facturada=base_amount,
                tipo_iva=vat_rate,
                iva_repercutido=iva_repercutido,
            )
        )

    return jsonify({"ok": True})


@app.route("/api/billing/summary")
def billing_summary():
    month = request.args.get("month", type=int)
    year = request.args.get("year", type=int)

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
                invoices_table.c.original_filename,
                invoices_table.c.expense_category,
            )
            .where(invoices_table.c.invoice_date.between(start, end))
            .order_by(invoices_table.c.invoice_date.desc(), invoices_table.c.id.desc())
        ).mappings().all()

    invoices = [
        {
            "id": row["id"],
            "invoice_date": row["invoice_date"],
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


@app.route("/api/invoices/<int:invoice_id>", methods=["PUT"])
def update_invoice(invoice_id):
    payload = request.get_json(silent=True) or {}

    invoice_date = payload.get("invoice_date") or ""
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
            .values(
                invoice_date=invoice_date,
                supplier=supplier,
                base_amount=base_amount,
                vat_rate=vat_rate,
                vat_amount=vat_amount,
                total_amount=total_amount,
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
    with engine.begin() as conn:
        result = conn.execute(
            invoices_table.delete().where(invoices_table.c.id == invoice_id)
        )

    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Factura no encontrada."]}), 404

    return jsonify({"ok": True})


@app.route("/api/expenses/no-invoice")
def list_no_invoice_expenses():
    month = request.args.get("month", type=int)
    year = request.args.get("year", type=int)

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
    with engine.begin() as conn:
        result = conn.execute(no_invoice_table.delete().where(no_invoice_table.c.id == expense_id))

    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Gasto no encontrado."]}), 404

    return jsonify({"ok": True})


@app.route("/api/billing/entries")
def billing_entries():
    month = request.args.get("month", type=int)
    year = request.args.get("year", type=int)

    today = date.today()
    month = month or today.month
    year = year or today.year

    with engine.connect() as conn:
        rows = conn.execute(
            select(
                facturacion_table.c.id,
                facturacion_table.c.mes,
                facturacion_table.c.anio,
                facturacion_table.c.base_facturada,
                facturacion_table.c.tipo_iva,
                facturacion_table.c.iva_repercutido,
            )
            .where(facturacion_table.c.mes == month, facturacion_table.c.anio == year)
            .order_by(facturacion_table.c.id.desc())
        ).mappings().all()

    entries = [
        {
            "id": row["id"],
            "month": row["mes"],
            "year": row["anio"],
            "base": float(row["base_facturada"]),
            "vat": int(row["tipo_iva"]),
            "vatAmount": float(row["iva_repercutido"]),
        }
        for row in rows
    ]

    return jsonify({"entries": entries})


@app.route("/api/billing/<int:billing_id>", methods=["PUT"])
def update_billing(billing_id):
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

    with engine.begin() as conn:
        result = conn.execute(
            facturacion_table.update()
            .where(facturacion_table.c.id == billing_id)
            .values(
                base_facturada=base_amount,
                tipo_iva=vat_rate,
                iva_repercutido=iva_repercutido,
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
            },
        }
    )


@app.route("/api/billing/<int:billing_id>", methods=["DELETE"])
def delete_billing(billing_id):
    with engine.begin() as conn:
        result = conn.execute(facturacion_table.delete().where(facturacion_table.c.id == billing_id))

    if result.rowcount == 0:
        return jsonify({"ok": False, "errors": ["Registro no encontrado."]}), 404

    return jsonify({"ok": True})


@app.route("/api/summary")
def summary():
    month = request.args.get("month", type=int)
    year = request.args.get("year", type=int)

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
