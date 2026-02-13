import os
import io
from datetime import datetime
from urllib.parse import urlparse, urljoin

from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from flask_sqlalchemy import SQLAlchemy

from flask_login import (
    LoginManager,
    UserMixin,
    login_user,
    logout_user,
    login_required,
    current_user,
)
from werkzeug.security import generate_password_hash, check_password_hash

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# -----------------------
# App & DB Config
# -----------------------
db_url = os.environ.get("DATABASE_URL", "sqlite:///crm.db")

# Railway/Heroku kadang pakai "postgres://"
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)

# Paksa SQLAlchemy pakai psycopg v3 (bukan psycopg2)
if db_url.startswith("postgresql://"):
    db_url = db_url.replace("postgresql://", "postgresql+psycopg://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = db_url


# -----------------------
# Login Manager
# -----------------------
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"  # endpoint function name
login_manager.login_message_category = "error"


@login_manager.unauthorized_handler
def unauthorized():
    """
    Hindari redirect loop. Kalau belum login, selalu lempar ke /login?next=...
    """
    next_url = request.full_path if request.query_string else request.path
    # Jangan bikin next jadi /login lagi (biar nggak loop)
    if next_url.startswith("/login"):
        next_url = "/"
    return redirect(url_for("login", next=next_url))


def is_safe_url(target: str) -> bool:
    """
    Cegah open redirect. Untuk lokal biasanya aman, tapi ini best practice.
    """
    if not target:
        return False
    ref_url = urlparse(request.host_url)
    test_url = urlparse(urljoin(request.host_url, target))
    return test_url.scheme in ("http", "https") and ref_url.netloc == test_url.netloc


# -----------------------
# Models
# -----------------------
class LeadSource(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), unique=True, nullable=False)


class Need(db.Model):  # Produk/jasa dibutuhkan
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), unique=True, nullable=False)


class Progress(db.Model):  # Progress (master)
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), unique=True, nullable=False)


class FollowUpStage(db.Model):  # Progress follow up (master tahap)
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), unique=True, nullable=False)


class Customer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(160), nullable=False)
    salesman_name = db.Column(db.String(160), nullable=True)

    address = db.Column(db.Text, nullable=True)
    phone_wa = db.Column(db.String(80), nullable=True)
    email = db.Column(db.String(160), nullable=True)
    pic = db.Column(db.String(160), nullable=True)

    lead_source_id = db.Column(db.Integer, db.ForeignKey("lead_source.id"), nullable=True)
    need_id = db.Column(db.Integer, db.ForeignKey("need.id"), nullable=True)

    progress_id = db.Column(db.Integer, db.ForeignKey("progress.id"), nullable=True)  # <--- baru

    # 3 kolom catatan panjang
    note_followup_awal = db.Column(db.Text, nullable=True)
    note_followup_lanjutan = db.Column(db.Text, nullable=True)
    management_comment = db.Column(db.Text, nullable=True)

    lead_source = db.relationship("LeadSource")
    need = db.relationship("Need")
    progress = db.relationship("Progress")  # <--- baru

    followups = db.relationship("FollowUpLog", backref="customer", cascade="all, delete-orphan")


class FollowUpLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey("customer.id"), nullable=False)

    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    stage_id = db.Column(db.Integer, db.ForeignKey("follow_up_stage.id"), nullable=True)
    note = db.Column(db.Text, nullable=True)

    stage = db.relationship("FollowUpStage")


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)

    def set_password(self, password: str):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)


@login_manager.user_loader
def load_user(user_id):
    # SQLAlchemy 2.x recommends Session.get, tapi query.get masih oke untuk simpel
    return User.query.get(int(user_id))


# -----------------------
# DB init / seed
# -----------------------
def ensure_seed_data():
    # Seed tahap default
    if FollowUpStage.query.count() == 0:
        defaults = ["New", "Contacted", "Qualified", "Proposal", "Negotiation", "Won", "Lost"]
        for n in defaults:
            db.session.add(FollowUpStage(name=n))
        db.session.commit()

    # Seed user admin default kalau belum ada user
    if User.query.count() == 0:
        admin_user = os.environ.get("ADMIN_USERNAME", "admin")
        admin_pass = os.environ.get("ADMIN_PASSWORD", "admin12345")
        u = User(username=admin_user)
        u.set_password(admin_pass)
        db.session.add(u)
        db.session.commit()
        print(f"[seed] created admin user: {admin_user}")


@app.before_request
def init_db_once():
    # Simple approach tanpa migration
    db.create_all()
    ensure_seed_data()


# -----------------------
# Masters CRUD (generic) - map
# -----------------------
MASTER_MAP = {
    "sources": (LeadSource, "Sumber Prospek"),
    "needs": (Need, "Produk/Jasa Dibutuhkan"),
    "progress": (Progress, "Progress"),  # <--- ganti ini
    "stages": (FollowUpStage, "Tahap Progress Follow Up"),
}


# -----------------------
# Routes - Auth
# -----------------------
@app.get("/login")
def login():
    if current_user.is_authenticated:
        return redirect(url_for("home"))

    # kalau sudah ada next dari unauthorized_handler, tampilkan form biasa
    return render_template("login.html")


@app.post("/login")
def login_post():
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")

    u = User.query.filter_by(username=username).first()
    if not u or not u.check_password(password):
        flash("Username / password salah.", "error")
        return redirect(url_for("login"))

    login_user(u)
    flash("Login berhasil.", "success")

    next_url = request.args.get("next")
    if next_url and is_safe_url(next_url):
        return redirect(next_url)
    return redirect(url_for("home"))


@app.get("/logout")
@login_required
def logout():
    logout_user()
    flash("Logout berhasil.", "success")
    return redirect(url_for("login"))


# -----------------------
# Routes - Dashboard
# -----------------------
@app.get("/")
@login_required
def home():
    customers_count = Customer.query.count()
    latest_followups = FollowUpLog.query.order_by(FollowUpLog.created_at.desc()).limit(10).all()
    return render_template(
        "home.html",
        customers_count=customers_count,
        latest_followups=latest_followups
    )


# -----------------------
# Routes - Customers
# -----------------------
@app.get("/customers")
@login_required
def customers_list():
    q = request.args.get("q", "").strip()
    query = Customer.query
    if q:
        like = f"%{q}%"
        query = query.filter(
            db.or_(
                Customer.name.ilike(like),
                Customer.pic.ilike(like),
                Customer.phone_wa.ilike(like),
                Customer.email.ilike(like),
                Customer.salesman_name.ilike(like),
            )
        )
    customers = query.order_by(Customer.id.desc()).all()
    return render_template("customers_list.html", customers=customers, q=q)


@app.get("/customers/new")
@login_required
def customers_new():
    return render_customer_form(Customer(), is_edit=False)


@app.post("/customers/new")
@login_required
def customers_create():
    c = Customer(
        name=request.form.get("name", "").strip(),
        salesman_name=request.form.get("salesman_name", "").strip(),
        address=request.form.get("address", "").strip(),
        phone_wa=request.form.get("phone_wa", "").strip(),
        email=request.form.get("email", "").strip(),
        pic=request.form.get("pic", "").strip(),
        lead_source_id=to_int_or_none(request.form.get("lead_source_id")),
        need_id=to_int_or_none(request.form.get("need_id")),
        progress_id=to_int_or_none(request.form.get("progress_id")),  # <---

        note_followup_awal=request.form.get("note_followup_awal", "").strip(),
        note_followup_lanjutan=request.form.get("note_followup_lanjutan", "").strip(),
        management_comment=request.form.get("management_comment", "").strip(),
    )

    if not c.name:
        flash("Nama customer wajib diisi.", "error")
        return render_customer_form(c, is_edit=False)

    db.session.add(c)
    db.session.commit()
    flash("Customer dibuat.", "success")
    return redirect(url_for("customer_detail", customer_id=c.id))


@app.get("/customers/<int:customer_id>")
@login_required
def customer_detail(customer_id: int):
    c = Customer.query.get_or_404(customer_id)
    stages = FollowUpStage.query.order_by(FollowUpStage.name.asc()).all()
    followups = FollowUpLog.query.filter_by(customer_id=c.id).order_by(FollowUpLog.created_at.desc()).all()
    return render_template("customer_detail.html", c=c, stages=stages, followups=followups)


@app.get("/customers/<int:customer_id>/edit")
@login_required
def customers_edit(customer_id: int):
    c = Customer.query.get_or_404(customer_id)
    return render_customer_form(c, is_edit=True)


@app.post("/customers/<int:customer_id>/edit")
@login_required
def customers_update(customer_id: int):
    c = Customer.query.get_or_404(customer_id)

    c.name = request.form.get("name", "").strip()
    c.salesman_name = request.form.get("salesman_name", "").strip()
    c.address = request.form.get("address", "").strip()
    c.phone_wa = request.form.get("phone_wa", "").strip()
    c.email = request.form.get("email", "").strip()
    c.pic = request.form.get("pic", "").strip()
    c.lead_source_id = to_int_or_none(request.form.get("lead_source_id"))
    c.need_id = to_int_or_none(request.form.get("need_id"))
    c.progress_id = to_int_or_none(request.form.get("progress_id"))

    c.note_followup_awal = request.form.get("note_followup_awal", "").strip()
    c.note_followup_lanjutan = request.form.get("note_followup_lanjutan", "").strip()
    c.management_comment = request.form.get("management_comment", "").strip()
    
    if not c.name:
        flash("Nama customer wajib diisi.", "error")
        return render_customer_form(c, is_edit=True)

    db.session.commit()
    flash("Customer diupdate.", "success")
    return redirect(url_for("customer_detail", customer_id=c.id))


@app.post("/customers/<int:customer_id>/delete")
@login_required
def customers_delete(customer_id: int):
    c = Customer.query.get_or_404(customer_id)
    db.session.delete(c)
    db.session.commit()
    flash("Customer dihapus.", "success")
    return redirect(url_for("customers_list"))


def render_customer_form(c: Customer, is_edit: bool):
    sources = LeadSource.query.order_by(LeadSource.name.asc()).all()
    needs = Need.query.order_by(Need.name.asc()).all()
    progresses = Progress.query.order_by(Progress.name.asc()).all()
    return render_template(
        "customer_form.html",
        c=c,
        is_edit=is_edit,
        sources=sources,
        needs=needs,
        progresses=progresses,
    )


# -----------------------
# Routes - Followup Log
# -----------------------
@app.post("/customers/<int:customer_id>/followups")
@login_required
def followups_add(customer_id: int):
    c = Customer.query.get_or_404(customer_id)
    stage_id = to_int_or_none(request.form.get("stage_id"))
    note = request.form.get("note", "").strip()

    fu = FollowUpLog(customer_id=c.id, stage_id=stage_id, note=note)
    db.session.add(fu)
    db.session.commit()
    flash("Follow up dicatat.", "success")
    return redirect(url_for("customer_detail", customer_id=c.id))


@app.post("/followups/<int:followup_id>/delete")
@login_required
def followups_delete(followup_id: int):
    fu = FollowUpLog.query.get_or_404(followup_id)
    cid = fu.customer_id
    db.session.delete(fu)
    db.session.commit()
    flash("Log follow up dihapus.", "success")
    return redirect(url_for("customer_detail", customer_id=cid))


# -----------------------
# Routes - Masters
# -----------------------
@app.get("/masters/<string:key>")
@login_required
def masters_list(key: str):
    model, title = get_master_model(key)
    items = model.query.order_by(model.name.asc()).all()
    return render_template("masters_list.html", key=key, title=title, items=items)


@app.get("/masters/<string:key>/new")
@login_required
def masters_new(key: str):
    _, title = get_master_model(key)
    return render_template("master_form.html", key=key, title=title, item=None, is_edit=False)


@app.post("/masters/<string:key>/new")
@login_required
def masters_create(key: str):
    model, title = get_master_model(key)
    name = request.form.get("name", "").strip()

    if not name:
        flash("Nama wajib diisi.", "error")
        return render_template("master_form.html", key=key, title=title, item=None, is_edit=False)

    if model.query.filter_by(name=name).first():
        flash("Nama sudah ada.", "error")
        return render_template("master_form.html", key=key, title=title, item=None, is_edit=False)

    item = model(name=name)
    db.session.add(item)
    db.session.commit()
    flash("Data master ditambahkan.", "success")
    return redirect(url_for("masters_list", key=key))


@app.get("/masters/<string:key>/<int:item_id>/edit")
@login_required
def masters_edit(key: str, item_id: int):
    model, title = get_master_model(key)
    item = model.query.get_or_404(item_id)
    return render_template("master_form.html", key=key, title=title, item=item, is_edit=True)


@app.post("/masters/<string:key>/<int:item_id>/edit")
@login_required
def masters_update(key: str, item_id: int):
    model, title = get_master_model(key)
    item = model.query.get_or_404(item_id)
    name = request.form.get("name", "").strip()

    if not name:
        flash("Nama wajib diisi.", "error")
        return render_template("master_form.html", key=key, title=title, item=item, is_edit=True)

    exists = model.query.filter(model.name == name, model.id != item.id).first()
    if exists:
        flash("Nama sudah dipakai item lain.", "error")
        return render_template("master_form.html", key=key, title=title, item=item, is_edit=True)

    item.name = name
    db.session.commit()
    flash("Data master diupdate.", "success")
    return redirect(url_for("masters_list", key=key))


@app.post("/masters/<string:key>/<int:item_id>/delete")
@login_required
def masters_delete(key: str, item_id: int):
    model, _ = get_master_model(key)
    item = model.query.get_or_404(item_id)
    db.session.delete(item)
    db.session.commit()
    flash("Data master dihapus.", "success")
    return redirect(url_for("masters_list", key=key))


# -----------------------
# Routes - Export
# -----------------------
@app.get("/export/customers.xlsx")
@login_required
def export_customers_xlsx():
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Customers"

    headers = [
        "ID",
        "Nama",
        "Salesman",
        "Alamat",
        "Telp/WA",
        "Email",
        "PIC",
        "Sumber Prospek",
        "Produk/Jasa Dibutuhkan",
        "Progress",
        "Catatan Follow Up Awal",
        "Catatan Follow Up Lanjutan",
        "Komen Manajemen",
    ]
    ws1.append(headers)

    customers = Customer.query.order_by(Customer.id.asc()).all()
    for c in customers:
        ws1.append([
            c.id,
            c.name,
            c.salesman_name or "",
            c.address or "",
            c.phone_wa or "",
            c.email or "",
            c.pic or "",
            c.lead_source.name if c.lead_source else "",
            c.need.name if c.need else "",
            c.progress.name if c.progress else "",
            c.note_followup_awal or "",
            c.note_followup_lanjutan or "",
            c.management_comment or "",
        ])

    for col in range(1, len(headers) + 1):
        ws1.column_dimensions[get_column_letter(col)].width = 22

    ws2 = wb.create_sheet("FollowUps")
    headers2 = ["ID", "Customer ID", "Customer", "Tanggal", "Stage", "Catatan"]
    ws2.append(headers2)

    logs = FollowUpLog.query.order_by(FollowUpLog.created_at.desc()).all()
    for fu in logs:
        ws2.append([
            fu.id,
            fu.customer_id,
            fu.customer.name if fu.customer else "",
            fu.created_at.strftime("%Y-%m-%d %H:%M"),
            fu.stage.name if fu.stage else "",
            fu.note or "",
        ])

    for col in range(1, len(headers2) + 1):
        ws2.column_dimensions[get_column_letter(col)].width = 24

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    return send_file(
        bio,
        as_attachment=True,
        download_name="customers_export.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# -----------------------
# Helpers
# -----------------------
def get_master_model(key: str):
    key = (key or "").strip().lower()  # <--- TAMBAH INI

    if key not in MASTER_MAP:
        flash("Master tidak dikenal.", "error")
        # kalau ini kepanggil saat belum login, unauthorized_handler akan handle
        return MASTER_MAP["sources"]
    return MASTER_MAP[key]


def to_int_or_none(v):
    try:
        if v is None or str(v).strip() == "":
            return None
        return int(v)
    except ValueError:
        return None

# -----------------------
# User Management
# -----------------------
@app.get("/users")
@login_required
def users_list():
    users = User.query.order_by(User.username.asc()).all()
    return render_template("users_list.html", users=users)

@app.get("/users/new")
@login_required
def users_new():
    return render_template("user_form.html", user=None)

@app.post("/users/new")
@login_required
def users_create():
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")

    if not username or not password:
        flash("Username dan password wajib diisi.", "error")
        return redirect(url_for("users_new"))

    if User.query.filter_by(username=username).first():
        flash("Username sudah ada.", "error")
        return redirect(url_for("users_new"))

    u = User(username=username)
    u.set_password(password)
    db.session.add(u)
    db.session.commit()

    flash("User berhasil dibuat.", "success")
    return redirect(url_for("users_list"))


# -----------------------
# Run local
# -----------------------
if __name__ == "__main__":
    app.run(debug=True)

