import os
import io
from datetime import datetime, date, timedelta
from urllib.parse import urlparse, urljoin
from decimal import Decimal
from collections import Counter, defaultdict

from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
from flask import send_from_directory

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
app = Flask(__name__)

app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key-change-me")

db_url = os.environ.get("DATABASE_URL", "sqlite:///crm.db")

# Railway/Heroku kadang pakai "postgres://", SQLAlchemy maunya "postgresql://"
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)

# Paksa SQLAlchemy pakai psycopg v3 (bukan psycopg2)
if db_url.startswith("postgresql://"):
    db_url = db_url.replace("postgresql://", "postgresql+psycopg://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"
login_manager.login_message_category = "error"

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
    prospect_next_followup_date = db.Column(db.Date, nullable=True)  # rencana FU berikutnya

    # 3 kolom catatan panjang
    note_followup_awal = db.Column(db.Text, nullable=True)
    note_followup_lanjutan = db.Column(db.Text, nullable=True)
    management_comment = db.Column(db.Text, nullable=True)

    lead_source = db.relationship("LeadSource")
    need = db.relationship("Need")
    progress = db.relationship("Progress")  # <--- baru

    followups = db.relationship("FollowUpLog", backref="customer", cascade="all, delete-orphan")

    prospect_date = db.Column(db.Date, nullable=True)
    estimated_value = db.Column(db.Numeric(14, 2), nullable=True)  # rupiah, aman besar

    from decimal import Decimal
    from sqlalchemy import func


class FollowUpLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey("customer.id"), nullable=False)

    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    stage_id = db.Column(db.Integer, db.ForeignKey("follow_up_stage.id"), nullable=True)
    note = db.Column(db.Text, nullable=True)

    stage = db.relationship("FollowUpStage")

class Attachment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey("customer.id"), nullable=False)

    original_filename = db.Column(db.String(255), nullable=False)
    stored_filename = db.Column(db.String(255), nullable=False)
    mime_type = db.Column(db.String(120), nullable=True)
    size_bytes = db.Column(db.Integer, nullable=True)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    customer = db.relationship("Customer", backref=db.backref("attachments", cascade="all, delete-orphan"))

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

def is_won_progress(name: str) -> bool:
    if not name:
        return False
    p = name.lower()
    return any(k in p for k in ["won", "closing", "closed", "deal", "berhasil", "success"])

def is_lost_progress(name: str) -> bool:
    if not name:
        return False
    p = name.lower()
    return any(k in p for k in ["lost", "gagal", "failed", "cancel", "batal"])

def is_offer_progress(name: str) -> bool:
    if not name:
        return False
    p = name.lower()
    return any(k in p for k in ["proposal", "penawaran", "quotation", "offer"])

def to_decimal_or_zero(v):
    try:
        if v is None or v == "":
            return Decimal("0")
        return Decimal(str(v))
    except Exception:
        return Decimal("0")

def parse_date_yyyy_mm_dd(s: str):
    s = (s or "").strip()
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        return None

def week_range(today: date):
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    return monday, sunday

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
from sqlalchemy import func

@app.get("/")
@login_required
def home():
    customers_count = Customer.query.count()
    latest_followups = FollowUpLog.query.order_by(FollowUpLog.created_at.desc()).limit(10).all()

    # ambil semua customers yang punya progress (untuk hitung won/lost)
    rows = db.session.query(
        Customer.id,
        Customer.salesman_name,
        Customer.estimated_value,
        Progress.name
    ).outerjoin(Progress, Customer.progress_id == Progress.id).all()

    total = 0
    won_count = 0
    lost_count = 0
    won_value = Decimal("0")
    lost_value = Decimal("0")

    sales_won = {}  # salesman -> count won

    for _id, salesman, est, prog_name in rows:
        total += 1
        is_won, is_lost = progress_flag(prog_name)

        if is_won:
            won_count += 1
            if est:
                won_value += Decimal(est)
            key = (salesman or "Unknown").strip() or "Unknown"
            sales_won[key] = sales_won.get(key, 0) + 1

        elif is_lost:
            lost_count += 1
            if est:
                lost_value += Decimal(est)

    other_count = max(total - won_count - lost_count, 0)
    not_won = max(total - won_count, 0)
    not_lost = max(total - lost_count, 0)

    # ===== Estimasi Nilai: Sudah Penawaran vs Belum Penawaran =====
    offer_keywords = ("proposal", "penawaran", "quotation", "quote", "offer")

    # ambil semua customer + progress (biar bisa cek keyword)
    all_customers = Customer.query.all()

    offer_value = Decimal("0")
    not_offer_value = Decimal("0")

    for c in all_customers:
        val = c.estimated_value or Decimal("0")
        p = (c.progress.name if c.progress else "").lower()

        is_offer = any(k in p for k in offer_keywords)
        if is_offer:
            offer_value += val
        else:
            not_offer_value += val

    pie_offer_labels = ["Sudah Penawaran", "Belum Penawaran"]
    pie_offer_values = [float(offer_value), float(not_offer_value)]


    # Top sales
    top_sales = sorted(sales_won.items(), key=lambda x: x[1], reverse=True)[:10]
    top_sales_labels = [x[0] for x in top_sales]
    top_sales_values = [x[1] for x in top_sales]

    dashboard = {
        "total": total,
        "won_count": won_count,
        "lost_count": lost_count,
        "other_count": other_count,
        "not_won": not_won,
        "not_lost": not_lost,
        "won_value": int(won_value),
        "lost_value": int(lost_value),
        "top_sales_labels": top_sales_labels,
        "top_sales_values": top_sales_values,
    }

   # ===== PIE: sumber prospek =====
    source_rows = (
        db.session.query(LeadSource.name, func.count(Customer.id))
        .join(Customer, Customer.lead_source_id == LeadSource.id)
        .group_by(LeadSource.name)
        .order_by(func.count(Customer.id).desc())
        .all()
    )
    pie_sources_labels = [r[0] for r in source_rows]
    pie_sources_values = [int(r[1]) for r in source_rows]

    # ===== PIE: produk/jasa (kebutuhan) =====
    need_rows = (
        db.session.query(Need.name, func.count(Customer.id))
        .join(Customer, Customer.need_id == Need.id)
        .group_by(Need.name)
        .order_by(func.count(Customer.id).desc())
        .all()
    )
    pie_needs_labels = [r[0] for r in need_rows]
    pie_needs_values = [int(r[1]) for r in need_rows]

    # Reminder followup minggu ini
    today = date.today()
    d1, d2 = week_range(today)
    followup_week = (
        Customer.query
        .filter(Customer.prospect_next_followup_date.isnot(None))
        .filter(Customer.prospect_next_followup_date >= d1)
        .filter(Customer.prospect_next_followup_date <= d2)
        .order_by(Customer.prospect_next_followup_date.asc())
        .limit(50)
        .all()
    )

        # Top sumber prospek yg menghasilkan closing (count + total nilai)
    won_customers = (
        Customer.query
        .join(Customer.progress, isouter=True)
        .join(Customer.lead_source, isouter=True)
        .all()
    )

    source_won_count = Counter()
    source_won_value = defaultdict(Decimal)

    for c in won_customers:
        pname = c.progress.name if getattr(c, "progress", None) else ""
        if is_won_progress(pname):
            sname = c.lead_source.name if c.lead_source else "(Tanpa Sumber)"
            source_won_count[sname] += 1
            source_won_value[sname] += to_decimal_or_zero(getattr(c, "estimated_value", None))

    # ambil top 10
    top_sources = sorted(source_won_count.items(), key=lambda x: x[1], reverse=True)[:10]
    top_sources_labels = [k for k, _ in top_sources]
    top_sources_values = [v for _, v in top_sources]
    top_sources_value_sum = [float(source_won_value[k]) for k in top_sources_labels]


    return render_template(
        "home.html",
        customers_count=customers_count,
        latest_followups=latest_followups,
        followup_week=followup_week,
        dashboard=dashboard,
        pie_offer_labels=pie_offer_labels,
        pie_offer_values=pie_offer_values,
        week_start=d1,
        week_end=d2,
        top_sources_won_labels=top_sources_labels,
        top_sources_won_values=top_sources_values,
        top_sources_won_value_sum=top_sources_value_sum,
    )


# -----------------------
# Routes - Customers
# -----------------------
from datetime import datetime, date

def to_date_or_none(s: str):
    try:
        if not s:
            return None
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        return None

@app.get("/customers")
@login_required
def customers_list():
    q = (request.args.get("q", "") or "").strip()
    date_from = (request.args.get("date_from", "") or "").strip()
    date_to = (request.args.get("date_to", "") or "").strip()

    query = Customer.query

    if q:
        like = f"%{q}%"
        query = query.filter(
            db.or_(
                Customer.name.ilike(like),
                Customer.pic.ilike(like),
                Customer.phone_wa.ilike(like),
                Customer.email.ilike(like),
            )
        )

    df = to_date_or_none(date_from)
    dt = to_date_or_none(date_to)

    if df:
        query = query.filter(Customer.prospect_date >= df)
    if dt:
        query = query.filter(Customer.prospect_date <= dt)

    customers = query.order_by(Customer.id.desc()).all()

    return render_template(
        "customers_list.html",
        customers=customers,
        q=q,
        date_from=date_from,
        date_to=date_to,
    )


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
        prospect_date=to_date_or_none(request.form.get("prospect_date")),
        address=request.form.get("address", "").strip(),
        phone_wa=request.form.get("phone_wa", "").strip(),
        email=request.form.get("email", "").strip(),
        pic=request.form.get("pic", "").strip(),
        lead_source_id=to_int_or_none(request.form.get("lead_source_id")),
        need_id=to_int_or_none(request.form.get("need_id")),
        estimated_value=to_decimal_or_none(request.form.get("estimated_value")),
        progress_id=to_int_or_none(request.form.get("progress_id")),  # <---
        c.prospect_next_followup_date = to_date_or_none(request.form.get("prospect_next_followup_date"))


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

@app.post("/customers/<int:customer_id>/attachments")
@login_required
def attachments_upload(customer_id: int):
    c = Customer.query.get_or_404(customer_id)

    f = request.files.get("file")
    if not f or f.filename.strip() == "":
        flash("Pilih file dulu.", "error")
        return redirect(url_for("customer_detail", customer_id=c.id))

    filename = secure_filename(f.filename)
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext not in ALLOWED_EXT:
        flash("File harus PDF/JPG/PNG.", "error")
        return redirect(url_for("customer_detail", customer_id=c.id))

    # simpan dengan nama unik
    ts = datetime.utcnow().strftime("%Y%m%d%H%M%S%f")
    stored = f"{c.id}_{ts}_{filename}"
    path = os.path.join(UPLOAD_DIR, stored)

    f.save(path)
    size = os.path.getsize(path)

    att = Attachment(
        customer_id=c.id,
        original_filename=filename,
        stored_filename=stored,
        mime_type=f.mimetype,
        size_bytes=size
    )
    db.session.add(att)
    db.session.commit()

    flash("Lampiran berhasil diupload.", "success")
    return redirect(url_for("customer_detail", customer_id=c.id))


@app.get("/attachments/<int:att_id>/download")
@login_required
def attachments_download(att_id: int):
    att = Attachment.query.get_or_404(att_id)
    # (opsional) pastikan customer ada
    return send_from_directory(
        UPLOAD_DIR,
        att.stored_filename,
        as_attachment=True,
        download_name=att.original_filename,
    )


@app.post("/attachments/<int:att_id>/delete")
@login_required
def attachments_delete(att_id: int):
    att = Attachment.query.get_or_404(att_id)
    cid = att.customer_id

    # hapus file fisik
    try:
        os.remove(os.path.join(UPLOAD_DIR, att.stored_filename))
    except Exception:
        pass

    db.session.delete(att)
    db.session.commit()
    flash("Lampiran dihapus.", "success")
    return redirect(url_for("customer_detail", customer_id=cid))

@app.get("/customers/import")
@login_required
def customers_import_page():
    return render_template("customers_import.html")


@app.post("/customers/import")
@login_required
def customers_import_post():
    f = request.files.get("file")
    if not f or f.filename.strip() == "":
        flash("Pilih file Excel (.xlsx) dulu.", "error")
        return redirect(url_for("customers_import_page"))

    if not f.filename.lower().endswith(".xlsx"):
        flash("File harus .xlsx", "error")
        return redirect(url_for("customers_import_page"))

    wb = Workbook()
    bio = io.BytesIO(f.read())
    from openpyxl import load_workbook
    wb = load_workbook(bio)
    ws = wb.active

    # header -> index (lower)
    header_row = [str(c.value or "").strip() for c in ws[1]]
    header_map = {h.lower(): i for i, h in enumerate(header_row)}

    def get_cell(row, key):
        idx = header_map.get(key.lower())
        if idx is None:
            return ""
        v = row[idx].value
        return "" if v is None else str(v).strip()

    created = 0
    updated = 0

    for r in ws.iter_rows(min_row=2):
        name = get_cell(r, "Nama")
        if not name:
            continue

        email = get_cell(r, "Email")
        phone = get_cell(r, "Telp/WA") or get_cell(r, "WA/Telp")
        salesman = get_cell(r, "Salesman")
        pic = get_cell(r, "PIC")
        address = get_cell(r, "Alamat")
        source_name = get_cell(r, "Sumber Prospek")
        need_name = get_cell(r, "Produk/Jasa") or get_cell(r, "Produk/Jasa Dibutuhkan")
        progress_name = get_cell(r, "Progress")
        est_value = get_cell(r, "Estimasi Nilai")
        prospect_date = get_cell(r, "Tgl Prospek")
        fu_next = get_cell(r, "Rencana FU") or get_cell(r, "Next Follow Up") or get_cell(r, "Rencana Follow Up")
        fu_awal = get_cell(r, "FU Awal") or get_cell(r, "Catatan Follow Up Awal")
        fu_lanjutan = get_cell(r, "FU Lanjutan") or get_cell(r, "Catatan Follow Up Lanjutan")
        km = get_cell(r, "Komen Manajemen")

        # masters: source / need / progress (kalau belum ada, buat)
        lead_source_id = None
        if source_name:
            ls = LeadSource.query.filter_by(name=source_name).first()
            if not ls:
                ls = LeadSource(name=source_name)
                db.session.add(ls)
                db.session.flush()
            lead_source_id = ls.id

        need_id = None
        if need_name:
            nd = Need.query.filter_by(name=need_name).first()
            if not nd:
                nd = Need(name=need_name)
                db.session.add(nd)
                db.session.flush()
            need_id = nd.id

        progress_id = None
        if progress_name:
            pg = Progress.query.filter_by(name=progress_name).first()
            if not pg:
                pg = Progress(name=progress_name)
                db.session.add(pg)
                db.session.flush()
            progress_id = pg.id

        # cari existing by email/phone
        c = None
        if email:
            c = Customer.query.filter_by(email=email).first()
        if not c and phone:
            c = Customer.query.filter_by(phone_wa=phone).first()

        # parse tanggal
        pd = parse_date_yyyy_mm_dd(prospect_date)
        ndt = parse_date_yyyy_mm_dd(fu_next)

        # parse nilai
        ev = None
        if est_value:
            # bersihin "Rp" dan koma
            cleaned = est_value.replace("Rp", "").replace(".", "").replace(",", "").strip()
            try:
                ev = Decimal(cleaned)
            except Exception:
                ev = None

        if c:
            c.name = name
            c.salesman_name = salesman
            c.pic = pic
            c.address = address
            c.phone_wa = phone
            c.email = email
            c.lead_source_id = lead_source_id
            c.need_id = need_id
            c.progress_id = progress_id
            c.estimated_value = ev
            c.prospect_date = pd
            c.prospect_next_followup_date = ndt
            c.note_followup_awal = fu_awal
            c.note_followup_lanjutan = fu_lanjutan
            c.management_comment = km
            updated += 1
        else:
            c = Customer(
                name=name,
                salesman_name=salesman,
                pic=pic,
                address=address,
                phone_wa=phone,
                email=email,
                lead_source_id=lead_source_id,
                need_id=need_id,
                progress_id=progress_id,
                estimated_value=ev,
                prospect_date=pd,
                prospect_next_followup_date=ndt,
                note_followup_awal=fu_awal,
                note_followup_lanjutan=fu_lanjutan,
                management_comment=km,
            )
            db.session.add(c)
            created += 1

    db.session.commit()
    flash(f"Import selesai. Created: {created}, Updated: {updated}", "success")
    return redirect(url_for("customers_list"))


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
    c.prospect_date = to_date_or_none(request.form.get("prospect_date"))
    c.address = request.form.get("address", "").strip()
    c.phone_wa = request.form.get("phone_wa", "").strip()
    c.email = request.form.get("email", "").strip()
    c.pic = request.form.get("pic", "").strip()
    c.lead_source_id = to_int_or_none(request.form.get("lead_source_id"))
    c.need_id = to_int_or_none(request.form.get("need_id"))
    c.estimated_value = to_decimal_or_none(request.form.get("estimated_value"))
    c.progress_id = to_int_or_none(request.form.get("progress_id"))
    c.prospect_next_followup_date = to_date_or_none(request.form.get("prospect_next_followup_date"))


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

@app.context_processor
def inject_helpers():
    return dict(fmt_idr=fmt_idr)


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

def to_date_or_none(v: str):
    try:
        v = (v or "").strip()
        if not v:
            return None
        # expect YYYY-MM-DD dari input type="date"
        return datetime.strptime(v, "%Y-%m-%d").date()
    except Exception:
        return None

def to_decimal_or_none(v: str):
    try:
        v = (v or "").strip()
        if not v:
            return None
        # buang pemisah ribuan umum: 1.234.567 atau 1,234,567
        v = v.replace(".", "").replace(",", "")
        return Decimal(v)
    except Exception:
        return None

def fmt_idr(x):
    try:
        if x is None:
            return "-"
        n = int(Decimal(x))
        return "Rp {:,}".format(n).replace(",", ".")
    except Exception:
        return "-"

def progress_flag(progress_name: str):
    p = (progress_name or "").lower()
    is_won = any(k in p for k in ["won", "closing", "closed", "deal", "berhasil", "success"])
    is_lost = any(k in p for k in ["lost", "gagal", "failed", "cancel", "batal"])
    return is_won, is_lost


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

def ensure_schema():
    engine = db.engine
    dialect = engine.dialect.name  # "sqlite" / "postgresql"

    def col_exists_pg(conn, table: str, col: str) -> bool:
        q = """
        SELECT 1
        FROM information_schema.columns
        WHERE table_schema='public'
          AND table_name=%s
          AND column_name=%s
        LIMIT 1
        """
        return conn.exec_driver_sql(q, (table, col)).first() is not None

    with engine.begin() as conn:
        if dialect == "sqlite":
            cols = [r[1] for r in conn.exec_driver_sql("PRAGMA table_info(customer)").fetchall()]
            if "prospect_next_followup_date" not in cols:
                conn.exec_driver_sql("ALTER TABLE customer ADD COLUMN prospect_next_followup_date DATE")
        else:
            # PostgreSQL (Railway)
            if not col_exists_pg(conn, "customer", "prospect_next_followup_date"):
                conn.exec_driver_sql("ALTER TABLE customer ADD COLUMN prospect_next_followup_date DATE")

# -----------------------
# Init DB (create tables + schema + seed)
# -----------------------
with app.app_context():
    db.create_all()
    ensure_schema()
    ensure_seed_data()
