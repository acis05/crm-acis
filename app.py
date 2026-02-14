import os
import io
import secrets
from functools import wraps
from decimal import Decimal
from collections import Counter, defaultdict
from datetime import datetime, date, timedelta

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    send_file,
    session,
)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from flask import send_from_directory


# =====================================================
# APP CONFIG
# =====================================================
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key-change-me")

db_url = os.environ.get("DATABASE_URL", "sqlite:///crm.db")

if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)

if db_url.startswith("postgresql://"):
    db_url = db_url.replace("postgresql://", "postgresql+psycopg://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)


# =====================================================
# UPLOAD CONFIG
# =====================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

ALLOWED_EXT = {"pdf", "jpg", "jpeg", "png"}


# =====================================================
# TENANT HELPER
# =====================================================
def current_company_id():
    cid = session.get("company_id")
    return int(cid) if cid else None


def tenant_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not current_company_id():
            return redirect(url_for("tenant_login", next=request.path))
        return fn(*args, **kwargs)
    return wrapper


# =====================================================
# MODELS
# =====================================================
class Company(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(160), nullable=False)
    access_code = db.Column(db.String(40), unique=True, nullable=False, index=True)
    pin_hash = db.Column(db.String(255), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def set_pin(self, pin: str):
        self.pin_hash = generate_password_hash(pin)

    def check_pin(self, pin: str):
        return check_password_hash(self.pin_hash, pin)


class LeadSource(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey("company.id"), nullable=False)
    name = db.Column(db.String(120), nullable=False)


class Need(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey("company.id"), nullable=False)
    name = db.Column(db.String(120), nullable=False)


class Progress(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey("company.id"), nullable=False)
    name = db.Column(db.String(120), nullable=False)


class FollowUpStage(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey("company.id"), nullable=False)
    name = db.Column(db.String(120), nullable=False)


class Customer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey("company.id"), nullable=False)

    name = db.Column(db.String(160), nullable=False)
    salesman_name = db.Column(db.String(160))
    address = db.Column(db.Text)
    phone_wa = db.Column(db.String(80))
    email = db.Column(db.String(160))
    pic = db.Column(db.String(160))

    lead_source_id = db.Column(db.Integer, db.ForeignKey("lead_source.id"))
    need_id = db.Column(db.Integer, db.ForeignKey("need.id"))
    progress_id = db.Column(db.Integer, db.ForeignKey("progress.id"))

    prospect_date = db.Column(db.Date)
    estimated_value = db.Column(db.Numeric(14, 2))
    prospect_next_followup_date = db.Column(db.Date)

    note_followup_awal = db.Column(db.Text)
    note_followup_lanjutan = db.Column(db.Text)
    management_comment = db.Column(db.Text)


class FollowUpLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey("company.id"), nullable=False)
    customer_id = db.Column(db.Integer, db.ForeignKey("customer.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    note = db.Column(db.Text)


class Attachment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey("company.id"), nullable=False)
    customer_id = db.Column(db.Integer, db.ForeignKey("customer.id"), nullable=False)
    original_filename = db.Column(db.String(255))
    stored_filename = db.Column(db.String(255))


# =====================================================
# SEED DEFAULT COMPANY
# =====================================================
def ensure_seed():
    if Company.query.count() == 0:
        comp = Company(
            name="Default Company",
            access_code="BP-DEMO",
        )
        comp.set_pin("123456")
        db.session.add(comp)
        db.session.commit()
        print("Default tenant created: BP-DEMO / 123456")


# =====================================================
# AUTH ROUTES (TENANT ONLY)
# =====================================================
@app.get("/tenant-login")
def tenant_login():
    return render_template("tenant_login.html")


@app.post("/tenant-login")
def tenant_login_post():
    access_code = request.form.get("access_code", "").strip()
    pin = request.form.get("pin", "").strip()

    company = Company.query.filter_by(access_code=access_code).first()

    if not company or not company.check_pin(pin):
        flash("Access Code / PIN salah.", "error")
        return redirect(url_for("tenant_login"))

    session["company_id"] = company.id
    flash(f"Login berhasil: {company.name}", "success")
    return redirect(url_for("home"))


@app.get("/logout")
def logout():
    session.pop("company_id", None)
    flash("Logout berhasil.", "success")
    return redirect(url_for("tenant_login"))


# =====================================================
# DASHBOARD
# =====================================================
@app.get("/")
@tenant_required
def home():
    cid = current_company_id()
    customers_count = Customer.query.filter_by(company_id=cid).count()

    return render_template(
        "home.html",
        customers_count=customers_count,
    )


# =====================================================
# CUSTOMERS
# =====================================================
@app.get("/customers")
@tenant_required
def customers_list():
    cid = current_company_id()
    customers = Customer.query.filter_by(company_id=cid).all()
    return render_template("customers_list.html", customers=customers)


@app.get("/customers/new")
@tenant_required
def customers_new():
    return render_template("customer_form.html", c=None, is_edit=False)


@app.post("/customers/new")
@tenant_required
def customers_create():
    cid = current_company_id()

    c = Customer(
        company_id=cid,
        name=request.form.get("name"),
        phone_wa=request.form.get("phone_wa"),
    )

    db.session.add(c)
    db.session.commit()
    flash("Customer dibuat.", "success")
    return redirect(url_for("customers_list"))


# =====================================================
# INIT
# =====================================================
with app.app_context():
    db.create_all()
    ensure_seed()


if __name__ == "__main__":
    app.run(debug=True)
