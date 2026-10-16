"""
Harly Click GLOBAL v4 - المتجر الشامل
✅ تسويق بالعمولة (خصم 10% + عمولة 5% + مكافحة احتيال + تعليق 14 يوم)
✅ المحفظة برقم المحوّل + تحكم الإدارة + تفعيل/تعطيل الميزة
✅ صفحة الإرجاع بقبول/رفض الإدارة
✅ لوحة تحليلات شاملة
✅ فاتورة PDF-جاهزة للطباعة
✅ تحكم كامل في حسابات العملاء والطلبات والمنتجات
✅ كل التغييرات تظهر فوراً (بيانات حية من قاعدة البيانات)
حساب الإدارة: بياناته في ملف .env (مش مكتوبة في الكود) — شوف CHANGELOG_AR.txt
✅ v5: أقسام بتتحكم فيها من الإدارة (منها المنتجات الغذائية) + طلب السلة عبر واتساب + تحسينات أمان
✅ v6: وحدة بيع وحد أدنى للمنتج + حفظ طلبات واتساب + شحن بالمحافظة + تصدير/بوليصة شحن + نسخ احتياطي يومي
       + ضغط صور وSEO + تحليلات موسّعة + دعم فني + تحكم كامل في العملاء وسجل المدفوعات
"""
from flask import (Flask, render_template, request, redirect,
                   url_for, session, flash, g, jsonify, Response,
                   send_from_directory)
import sqlite3
import os
import re
import csv
import io
import secrets
from datetime import date, datetime, timedelta
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash
import math
import hmac
import gzip
import json
import shutil
import threading
import time
import hashlib
from urllib.parse import quote, urlparse
from xml.sax.saxutils import escape as xml_escape
from markupsafe import Markup

try:                      # ضغط الصور عند الرفع (لو Pillow مش متسطبة الرفع بيشتغل عادي من غير ضغط)
    from PIL import Image, ImageOps
    Image.MAX_IMAGE_PIXELS = 40_000_000
except Exception:         # pragma: no cover
    Image = None
try:                      # تصدير Excel (لو مش متسطبة التصدير CSV بيفضل شغال)
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
except Exception:         # pragma: no cover
    Workbook = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _load_env_file(path):
    """يقرأ ملف .env (لو موجود) ويحط القيم في متغيرات البيئة — من غير أي مكتبة إضافية"""
    try:
        with open(path, encoding="utf-8-sig") as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k, v = k.strip(), v.strip()
                if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
                    v = v[1:-1]
                os.environ.setdefault(k, v)
    except OSError:
        pass


_load_env_file(os.path.join(BASE_DIR, ".env"))


def _get_secret_key():
    """مفتاح الجلسات: من SECRET_KEY في .env، أو يتولّد مرة واحدة ويتحفظ في .secret_key"""
    key = os.environ.get("SECRET_KEY", "").strip()
    if key:
        return key
    path = os.path.join(BASE_DIR, ".secret_key")
    try:
        with open(path, encoding="utf-8") as fh:
            k = fh.read().strip()
            if k:
                return k
    except OSError:
        pass
    k = secrets.token_hex(32)
    try:
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(k)
        os.chmod(path, 0o600)
    except OSError:
        pass
    return k


app = Flask(__name__)
app.secret_key = _get_secret_key()
app.config.update(
    SEND_FILE_MAX_AGE_DEFAULT=timedelta(days=30),   # كاش للصور والـ CSS (الـ CSS ليه رقم إصدار)
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("COOKIE_SECURE", "0") == "1",
    PERMANENT_SESSION_LIFETIME=timedelta(days=14),
    MAX_CONTENT_LENGTH=8 * 1024 * 1024,   # أقصى حجم رفع 8MB
)
DB = os.path.join(BASE_DIR, "harly_click.db")
UPLOAD_DIR = os.path.join(BASE_DIR, "static", "uploads", "products")

# بيانات الإدارة بتتقرا من ملف .env (مش مكتوبة هنا)
ADMIN_EMAIL = (os.environ.get("ADMIN_EMAIL") or "").strip().lower()
ADMIN_PASS = os.environ.get("ADMIN_PASS", "")              # اختياري (نص عادي)
ADMIN_PASS_HASH = os.environ.get("ADMIN_PASS_HASH", "")    # الأفضل: هاش مش نص
TRANSFER_NUMBER = os.environ.get("TRANSFER_NUMBER", "01222370381")
WHATSAPP_NUMBER = os.environ.get("WHATSAPP_NUMBER", "01222370381")   # رقم استقبال طلبات واتساب
COUNTRY_CODE = os.environ.get("COUNTRY_CODE", "20")                  # مصر
TRUST_PROXY = os.environ.get("TRUST_PROXY", "1") == "1"

PAY_METHODS = {"vodafone": "فودافون كاش 📱", "orange": "أورانج كاش 📱",
               "etisalat": "اتصالات كاش 📱", "instapay": "انستاباي 🏦"}
# أقسام السوبر ماركت (الاسم، الإيموجي، يظهر في الصفحة الرئيسية؟) — الإدارة بتتحكم فيها بعد كده من «الأقسام»
SUPERMARKET_CATEGORIES = [
    ("بقالة وأساسيات", "🛒", 1), ("ألبان وأجبان وبيض", "🥛", 1), ("لحوم ودواجن وأسماك", "🍗", 0),
    ("خضروات وفاكهة", "🥬", 1), ("مخبوزات وفطور", "🥖", 0), ("مشروبات", "🥤", 1),
    ("سناكس وحلويات", "🍫", 0), ("معلبات ومجمدات", "🥫", 0),
    ("منظفات ومستلزمات المنزل", "🧴", 0), ("عناية شخصية", "🧼", 0)]
CATEGORIES = [c[0] for c in SUPERMARKET_CATEGORIES]
# طرق الدفع وقت الشراء + أسماء كل الطرق (بتتعرض في الفواتير والتقارير)
CHECKOUT_METHODS = [("cod", "💵 كاش عند الاستلام"), ("balance", "💰 رصيد المحفظة"), ("partial", "💸 دفع جزئي من الرصيد")]
PAY_LABELS = dict(PAY_METHODS, **dict(CHECKOUT_METHODS))
FOOD_CATEGORY = CATEGORIES[0]
# أقسام ومنتجات النسخ القديمة (بتتشال مرة واحدة تلقائياً لو لسه تجريبية ومحدش لمسها)
OLD_DEFAULT_CATEGORIES = ["إلكترونيات", "ملابس", "أحذية", "إكسسوارات", "منزل", "رياضة", "منتجات غذائية"]
OLD_DEMO_PRODUCTS = [("سماعات بلوتوث لاسلكية", 499), ("ساعة ذكية Harly S7", 1299), ("جاكيت جلد كلاسيك", 899),
                     ("حذاء رياضي إير فلكس", 649), ("حقيبة ظهر عصرية", 349), ("ماوس جيمنج احترافي", 279),
                     ("طقم أدوات مطبخ 12 قطعة", 459), ("كرة قدم احترافية", 199), ("شاحن سريع 65W", 399),
                     ("نظارة شمسية UV400", 299), ("تيشيرت قطن 100%", 149), ("جهاز تتبع لياقة", 549)]
# منتجات سوبر ماركت تجريبية (الاسم، الوصف، السعر، السعر القديم، المخزون، القسم، إيموجي، مميز، فلاش، الوحدة، الحد الأدنى)
SUPERMARKET_DEMO = [
    ("أرز مصري فاخر", "أرز مصري حبة طويلة نظيف وجاهز للطبخ.", 42, 48, 300, "بقالة وأساسيات", "🍚", 1, 0, "كيلو", 1),
    ("سكر أبيض ناعم", "سكر أبيض ناعم نقي للاستخدام اليومي.", 32, 0, 250, "بقالة وأساسيات", "🧂", 0, 0, "كيلو", 1),
    ("زيت عباد الشمس", "زيت عباد الشمس الصافي للطبخ والقلي.", 95, 110, 120, "بقالة وأساسيات", "🫒", 0, 1, "زجاجة", 1),
    ("مكرونة سباجتي 400 جم", "مكرونة من القمح الصلب تستوي في 8 دقايق.", 14, 0, 500, "بقالة وأساسيات", "🍝", 0, 0, "كيس", 5),
    ("لبن كامل الدسم 1 لتر", "لبن طازج معقم كامل الدسم.", 48, 0, 150, "ألبان وأجبان وبيض", "🥛", 1, 0, "علبة", 2),
    ("جبنة رومي", "جبنة رومي قديمة بطعم مميز.", 260, 290, 40, "ألبان وأجبان وبيض", "🧀", 0, 0, "كيلو", 1),
    ("بيض بلدي (30 بيضة)", "بيض طازج من المزرعة.", 150, 165, 60, "ألبان وأجبان وبيض", "🥚", 0, 1, "كرتونة", 1),
    ("زبادي طبيعي", "زبادي طبيعي كامل الدسم.", 9, 0, 400, "ألبان وأجبان وبيض", "🍶", 0, 0, "علبة", 6),
    ("فراخ بيضاء طازجة", "فراخ بيضاء منظفة وجاهزة للطبخ.", 105, 0, 80, "لحوم ودواجن وأسماك", "🍗", 1, 0, "كيلو", 1),
    ("لحمة كندوز", "لحمة كندوز طازجة مفرومة أو قطع.", 430, 470, 50, "لحوم ودواجن وأسماك", "🥩", 0, 1, "كيلو", 1),
    ("سمك بلطي", "سمك بلطي طازج منظف.", 90, 0, 45, "لحوم ودواجن وأسماك", "🐟", 0, 0, "كيلو", 1),
    ("طماطم طازجة", "طماطم حمرا طازجة للسلطة والطبخ.", 22, 0, 200, "خضروات وفاكهة", "🍅", 0, 0, "كيلو", 2),
    ("بطاطس", "بطاطس مناسبة للقلي والسلق.", 20, 0, 250, "خضروات وفاكهة", "🥔", 0, 0, "كيلو", 2),
    ("موز", "موز طازج ناضج وحلو.", 40, 45, 120, "خضروات وفاكهة", "🍌", 1, 1, "كيلو", 1),
    ("برتقال", "برتقال عصير حلو.", 28, 0, 150, "خضروات وفاكهة", "🍊", 0, 0, "كيلو", 2),
    ("عسل نحل طبيعي", "عسل نحل طبيعي 100% بدون إضافات.", 320, 360, 60, "مخبوزات وفطور", "🍯", 1, 0, "علبة", 1),
    ("مربى فراولة", "مربى فراولة بقطع الفاكهة.", 38, 0, 90, "مخبوزات وفطور", "🍓", 0, 0, "علبة", 1),
    ("كورن فليكس", "رقائق الذرة المقرمشة للفطار.", 75, 85, 80, "مخبوزات وفطور", "🥣", 0, 0, "علبة", 1),
    ("مياه معدنية", "مياه معدنية نقية (كرتونة 12 زجاجة).", 80, 0, 100, "مشروبات", "💧", 0, 0, "كرتونة", 1),
    ("شاي أسود", "شاي أسود فاخر 100 فتلة.", 95, 105, 120, "مشروبات", "🍵", 1, 0, "علبة", 1),
    ("عصير برتقال 1 لتر", "عصير برتقال طبيعي بدون سكر مضاف.", 42, 0, 100, "مشروبات", "🧃", 0, 0, "علبة", 3),
    ("شيبس بطاطس", "شيبس بطاطس مقرمش بالملح.", 10, 0, 300, "سناكس وحلويات", "🥔", 0, 0, "كيس", 6),
    ("شوكولاتة ألواح", "علبة ألواح شوكولاتة بالحليب (12 قطعة).", 150, 170, 70, "سناكس وحلويات", "🍫", 0, 1, "علبة", 1),
    ("تونة معلبة", "تونة قطع في زيت نباتي.", 62, 0, 200, "معلبات ومجمدات", "🥫", 0, 0, "علبة", 3),
    ("بسلة مجمدة", "بسلة خضرا مجمدة جاهزة للطبخ.", 38, 0, 90, "معلبات ومجمدات", "🫛", 0, 0, "كيس", 2),
    ("مسحوق غسيل 3 كجم", "مسحوق غسيل أوتوماتيك برائحة منعشة.", 190, 215, 70, "منظفات ومستلزمات المنزل", "🧺", 0, 1, "كيس", 1),
    ("سائل غسيل أطباق", "سائل أطباق بقوة تنظيف عالية.", 45, 0, 110, "منظفات ومستلزمات المنزل", "🧴", 0, 0, "زجاجة", 1),
    ("شامبو للشعر", "شامبو بتركيبة لطيفة للاستخدام اليومي.", 92, 0, 80, "عناية شخصية", "🧴", 0, 0, "زجاجة", 1),
    ("صابون (4 قطع)", "صابون معطر للجسم عبوة 4 قطع.", 60, 70, 100, "عناية شخصية", "🧼", 0, 0, "علبة", 1),
]
ORDER_STATUSES = ["قيد التجهيز", "تم الشحن", "تم التوصيل", "ملغي"]
UNITS = ["قطعة", "كيلو", "علبة", "كرتونة", "كيس", "زجاجة", "لتر"]   # وحدات البيع (الافتراضي: قطعة)
DEFAULT_UNIT = "قطعة"
WA_STATUSES = ["جديد", "تم التواصل", "تم التأكيد", "ملغي"]
TICKET_STATUSES = ["جديدة", "قيد المتابعة", "تم الحل"]
TX_TYPES = {"recharge": "📲 شحن", "purchase": "🛒 شراء", "admin": "⚙️ إدارة",
            "affiliate": "🤝 عمولة", "refund": "🔄 استرداد"}
# المحافظات + مصاريف الشحن المبدئية (أرقام مثال — بتعدّلها من الإدارة ← الشحن)
GOVERNORATES = [("القاهرة", 40), ("الجيزة", 40), ("القليوبية", 45), ("الإسكندرية", 55),
                ("الشرقية", 55), ("الدقهلية", 55), ("الغربية", 55), ("المنوفية", 55),
                ("البحيرة", 55), ("كفر الشيخ", 55), ("دمياط", 55), ("بورسعيد", 60),
                ("الإسماعيلية", 60), ("السويس", 60), ("الفيوم", 60), ("بني سويف", 60),
                ("المنيا", 75), ("أسيوط", 75), ("سوهاج", 75), ("قنا", 75), ("الأقصر", 75),
                ("أسوان", 75), ("البحر الأحمر", 95), ("الوادي الجديد", 95), ("مطروح", 95),
                ("شمال سيناء", 95), ("جنوب سيناء", 95)]
BACKUP_DIR = os.path.join(BASE_DIR, "backups")
BACKUP_NAME_RE = re.compile(r"^harly_click_\d{4}-\d{2}-\d{2}_\d{6}_[a-z]+\.db\.gz$")
SUPPORT_PHONE = os.environ.get("SUPPORT_PHONE", "01222370381")
SUPPORT_EMAIL = os.environ.get("SUPPORT_EMAIL", "250iie3@gmail.com")
SITE_URL = (os.environ.get("SITE_URL") or "").strip().rstrip("/")   # لو الموقع على دومين (للـ SEO)

AFF_BUYER_DISCOUNT = 10      # خصم المشتري %
AFF_COMMISSION = 5           # عمولة المسوق %
AFF_HOLD_DAYS = 14           # تعليق العمولة 14 يوم
RETURN_DAYS = 14             # مدة طلب الإرجاع


# ================== قاعدة البيانات ==================
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


SCHEMA = """
CREATE TABLE IF NOT EXISTS users(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    email TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL,
    phone TEXT,
    address TEXT,
    balance REAL DEFAULT 0,
    is_admin INTEGER DEFAULT 0,
    active INTEGER DEFAULT 1,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS products(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    description TEXT,
    price REAL NOT NULL,
    old_price REAL,
    stock INTEGER DEFAULT 10,
    sold INTEGER DEFAULT 0,
    category TEXT,
    emoji TEXT DEFAULT '📦',
    image TEXT DEFAULT '',
    featured INTEGER DEFAULT 0,
    flash INTEGER DEFAULT 0,
    unit TEXT DEFAULT 'قطعة',
    min_order INTEGER DEFAULT 1,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS orders(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER REFERENCES users(id),
    total REAL NOT NULL,
    discount REAL DEFAULT 0,
    coupon TEXT,
    payment_method TEXT DEFAULT 'cod',
    paid REAL DEFAULT 0,
    status TEXT DEFAULT 'قيد التجهيز',
    address TEXT,
    phone TEXT,
    shipping REAL DEFAULT 0,
    governorate TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS order_items(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id INTEGER REFERENCES orders(id),
    product_id INTEGER REFERENCES products(id),
    qty INTEGER,
    price REAL
);
CREATE TABLE IF NOT EXISTS wishlist(
    user_id INTEGER REFERENCES users(id),
    product_id INTEGER REFERENCES products(id),
    PRIMARY KEY(user_id, product_id)
);
CREATE TABLE IF NOT EXISTS coupons(
    code TEXT PRIMARY KEY,
    percent INTEGER NOT NULL,
    active INTEGER DEFAULT 1,
    expires TEXT
);
CREATE TABLE IF NOT EXISTS wallet_requests(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER REFERENCES users(id),
    amount REAL NOT NULL,
    method TEXT NOT NULL,
    sender_number TEXT DEFAULT '',
    status TEXT DEFAULT 'pending',
    note TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    handled_at TEXT
);
CREATE TABLE IF NOT EXISTS wallet_transactions(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER REFERENCES users(id),
    amount REAL NOT NULL,
    type TEXT NOT NULL,
    note TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
-- ===== التسويق بالعمولة =====
CREATE TABLE IF NOT EXISTS marketers(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER UNIQUE REFERENCES users(id),
    code TEXT UNIQUE NOT NULL,
    active INTEGER DEFAULT 0,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS affiliate_orders(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id INTEGER REFERENCES orders(id),
    marketer_id INTEGER REFERENCES marketers(id),
    buyer_email TEXT,
    buyer_ip TEXT,
    order_total REAL,
    buyer_discount REAL,
    commission REAL,
    status TEXT DEFAULT 'pending',   -- pending / confirmed / cancelled
    flag INTEGER DEFAULT 0,
    flag_reason TEXT,
    hold_until TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
-- ===== الإرجاع =====
CREATE TABLE IF NOT EXISTS returns(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id INTEGER REFERENCES orders(id),
    user_id INTEGER REFERENCES users(id),
    reason TEXT,
    status TEXT DEFAULT 'pending',   -- pending / approved / rejected
    admin_note TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    handled_at TEXT
);
-- ===== إعدادات الموقع (تفعيل/تعطيل المميزات) =====
CREATE TABLE IF NOT EXISTS settings(
    key TEXT PRIMARY KEY,
    value TEXT
);
-- ===== الأقسام (بتتحكم فيها الإدارة: إضافة/تعديل/إيقاف/حذف) =====
CREATE TABLE IF NOT EXISTS categories(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    emoji TEXT DEFAULT '',
    active INTEGER DEFAULT 1,
    show_home INTEGER DEFAULT 0,
    sort_order INTEGER DEFAULT 0,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
-- ===== محاولات الدخول الغلط (حماية من التخمين) =====
CREATE TABLE IF NOT EXISTS login_attempts(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ip TEXT,
    email TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
-- ===== مصاريف الشحن حسب المحافظة =====
CREATE TABLE IF NOT EXISTS shipping_zones(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    governorate TEXT UNIQUE NOT NULL,
    fee REAL DEFAULT 0,
    active INTEGER DEFAULT 1,
    sort_order INTEGER DEFAULT 0
);
-- ===== طلبات واتساب (بتتسجل أول ما العميل يدوس الزرار) =====
CREATE TABLE IF NOT EXISTS whatsapp_orders(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    customer_name TEXT,
    phone TEXT,
    address TEXT,
    governorate TEXT,
    items_json TEXT,
    subtotal REAL DEFAULT 0,
    discount REAL DEFAULT 0,
    shipping REAL DEFAULT 0,
    total REAL DEFAULT 0,
    source TEXT DEFAULT 'cart',
    status TEXT DEFAULT 'جديد',
    admin_note TEXT DEFAULT '',
    fingerprint TEXT,
    ip TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
-- ===== الدعم الفني =====
CREATE TABLE IF NOT EXISTS support_tickets(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    name TEXT,
    email TEXT,
    phone TEXT,
    subject TEXT,
    message TEXT,
    status TEXT DEFAULT 'جديدة',
    admin_reply TEXT DEFAULT '',
    replied_at TEXT,
    ip TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
-- ===== فهارس للسرعة =====
CREATE INDEX IF NOT EXISTS idx_products_category ON products(category);
CREATE INDEX IF NOT EXISTS idx_orders_user ON orders(user_id);
CREATE INDEX IF NOT EXISTS idx_order_items_order ON order_items(order_id);
CREATE INDEX IF NOT EXISTS idx_wtx_user ON wallet_transactions(user_id);
CREATE INDEX IF NOT EXISTS idx_wreq_user ON wallet_requests(user_id);
CREATE INDEX IF NOT EXISTS idx_wa_status ON whatsapp_orders(status);
CREATE INDEX IF NOT EXISTS idx_ticket_user ON support_tickets(user_id);
"""

DEFAULT_SETTINGS = {"wallet_sender": "1", "affiliate": "1", "returns": "1",
                    "whatsapp_cart": "1", "whatsapp_number": WHATSAPP_NUMBER,
                    "shipping_enabled": "1", "free_shipping_min": "1000",
                    "support_phone": SUPPORT_PHONE, "support_email": SUPPORT_EMAIL,
                    "support_hours": "يومياً 9ص - 11م", "store_address": "شارع التحرير، الدقي، الجيزة",
                    "backup_keep": "14", "transfer_number": TRANSFER_NUMBER,
                    "support_whatsapp": SUPPORT_PHONE, "support_whatsapp_on": "1",
                    "pm_vodafone_on": "1", "pm_orange_on": "1", "pm_etisalat_on": "1", "pm_instapay_on": "1",
                    "pm_vodafone_number": "", "pm_orange_number": "", "pm_etisalat_number": "", "pm_instapay_number": "",
                    "co_cod_on": "1", "co_balance_on": "1", "co_partial_on": "1"}


def get_setting(key, default=""):
    r = get_db().execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return r["value"] if r else default


def _seed_sold(db):
    """أرقام مبيعات تجريبية للمنتجات التجريبية (بالاسم عشان أرقام الـ id مش ثابتة)"""
    for name, sold in [("أرز مصري فاخر", 45), ("لبن كامل الدسم 1 لتر", 38), ("فراخ بيضاء طازجة", 30),
                       ("موز", 26), ("عسل نحل طبيعي", 22), ("شاي أسود", 18)]:
        db.execute("UPDATE products SET sold=? WHERE name=?", (sold, name))


def init_db():
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    db = sqlite3.connect(DB)
    db.executescript(SCHEMA)
    db.row_factory = sqlite3.Row
    for stmt in ["ALTER TABLE users ADD COLUMN active INTEGER DEFAULT 1",
                 "ALTER TABLE users ADD COLUMN balance REAL DEFAULT 0",
                 "ALTER TABLE products ADD COLUMN image TEXT DEFAULT ''",
                 "ALTER TABLE products ADD COLUMN sold INTEGER DEFAULT 0",
                 "ALTER TABLE products ADD COLUMN flash INTEGER DEFAULT 0",
                 "ALTER TABLE products ADD COLUMN emoji TEXT DEFAULT '📦'",
                 "ALTER TABLE orders ADD COLUMN payment_method TEXT DEFAULT 'cod'",
                 "ALTER TABLE orders ADD COLUMN paid REAL DEFAULT 0",
                 "ALTER TABLE orders ADD COLUMN discount REAL DEFAULT 0",
                 "ALTER TABLE orders ADD COLUMN coupon TEXT",
                 "ALTER TABLE users ADD COLUMN phone TEXT",
                 "ALTER TABLE users ADD COLUMN address TEXT",
                 "ALTER TABLE wallet_requests ADD COLUMN sender_number TEXT DEFAULT ''",
                 "ALTER TABLE products ADD COLUMN unit TEXT DEFAULT 'قطعة'",
                 "ALTER TABLE products ADD COLUMN min_order INTEGER DEFAULT 1",
                 "ALTER TABLE orders ADD COLUMN shipping REAL DEFAULT 0",
                 "ALTER TABLE orders ADD COLUMN governorate TEXT"]:
        try:
            db.execute(stmt)
        except Exception:
            pass
    for k, v in DEFAULT_SETTINGS.items():
        db.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (k, v))
    # محافظات الشحن: بتتزرع مرة واحدة وبعد كده الإدارة هي اللي بتتحكم
    if db.execute("SELECT COUNT(*) c FROM shipping_zones").fetchone()["c"] == 0:
        db.executemany("INSERT OR IGNORE INTO shipping_zones(governorate,fee,active,sort_order) VALUES(?,?,1,?)",
                       [(g, fee, i) for i, (g, fee) in enumerate(GOVERNORATES, 1)])
    db.execute("UPDATE products SET unit=? WHERE unit IS NULL OR unit=''", (DEFAULT_UNIT,))
    db.execute("UPDATE products SET min_order=1 WHERE min_order IS NULL OR min_order<1")
    try:                       # أسرع وأأمن مع أكتر من زائر في نفس الوقت (والنسخ الاحتياطي)
        db.execute("PRAGMA journal_mode=WAL")
        db.execute("PRAGMA synchronous=NORMAL")
    except Exception:
        pass
    # حساب الإدارة (البيانات من ملف .env — مش من الكود)
    admin_pw_hash = ADMIN_PASS_HASH or (generate_password_hash(ADMIN_PASS) if ADMIN_PASS else "")
    if ADMIN_EMAIL:
        admin = db.execute("SELECT id FROM users WHERE email=?", (ADMIN_EMAIL,)).fetchone()
        if admin:
            if admin_pw_hash:
                db.execute("UPDATE users SET password=?, is_admin=1, active=1 WHERE email=?",
                           (admin_pw_hash, ADMIN_EMAIL))
            else:
                db.execute("UPDATE users SET is_admin=1, active=1 WHERE email=?", (ADMIN_EMAIL,))
        else:
            old = db.execute("SELECT id FROM users WHERE email='admin@harlyclick.local' AND is_admin=1").fetchone()
            if old and admin_pw_hash:     # الحساب المؤقت القديم يتحوّل لحساب الإدارة الجديد (بدل ما يفضل حساب تاني)
                db.execute("UPDATE users SET email=?, password=?, active=1 WHERE id=?", (ADMIN_EMAIL, admin_pw_hash, old["id"]))
            else:
                if not admin_pw_hash:
                    tmp = secrets.token_urlsafe(10)
                    admin_pw_hash = generate_password_hash(tmp)
                    print("⚠️ اتعمل حساب إدارة جديد — كلمة السر المؤقتة (غيّرها): %s" % tmp)
                db.execute("INSERT INTO users(name,email,password,is_admin,active) VALUES(?,?,?,1,1)",
                           ("الإدارة", ADMIN_EMAIL, admin_pw_hash))
    elif not db.execute("SELECT 1 FROM users WHERE is_admin=1").fetchone():
        tmp = secrets.token_urlsafe(10)
        db.execute("INSERT INTO users(name,email,password,is_admin,active) VALUES(?,?,?,1,1)",
                   ("الإدارة", "admin@harlyclick.local", generate_password_hash(tmp)))
        print("⚠️ مفيش ADMIN_EMAIL في .env — اتعمل حساب admin@harlyclick.local / %s (غيّرهم)" % tmp)
    # منتجات تجريبية (سوبر ماركت) — لقاعدة جديدة بس
    if db.execute("SELECT COUNT(*) c FROM products").fetchone()["c"] == 0:
        db.executemany("""INSERT INTO products(name,description,price,old_price,stock,category,emoji,featured,flash,unit,min_order)
                          VALUES(?,?,?,?,?,?,?,?,?,?,?)""", SUPERMARKET_DEMO)
        _seed_sold(db)
    if db.execute("SELECT COUNT(*) c FROM coupons").fetchone()["c"] == 0:
        db.executemany("INSERT INTO coupons(code,percent,expires) VALUES(?,?,?)",
                       [("HARLY10", 10, "2027-12-31"), ("WELCOME20", 20, "2026-12-31")])
    # الأقسام: بتتعمل مرة واحدة (الأقسام القديمة + المنتجات الغذائية) وبعد كده الإدارة هي اللي بتتحكم
    if db.execute("SELECT COUNT(*) c FROM categories").fetchone()["c"] == 0:
        order = 0
        names = list(CATEGORIES) + [FOOD_CATEGORY]
        for (c,) in db.execute("SELECT DISTINCT category FROM products WHERE category IS NOT NULL AND category!=''").fetchall():
            if c not in names:
                names.append(c)
        meta = {n: (e, h) for n, e, h in SUPERMARKET_CATEGORIES}
        for c in names:
            order += 1
            db.execute("INSERT OR IGNORE INTO categories(name,emoji,active,show_home,sort_order) VALUES(?,?,?,?,?)",
                       (c, meta.get(c, ("", 0))[0], 1, meta.get(c, ("", 0))[1], order))
    # تحويل المتجر لسوبر ماركت (مرة واحدة): يشيل الأقسام/المنتجات التجريبية القديمة اللي محدش لمسها
    if not db.execute("SELECT 1 FROM settings WHERE key='supermarket_v1'").fetchone():
        for n, price in OLD_DEMO_PRODUCTS:
            db.execute("""DELETE FROM products WHERE name=? AND price=? AND (image IS NULL OR image='')
                          AND id NOT IN (SELECT product_id FROM order_items)""", (n, price))
        db.execute("DELETE FROM wishlist WHERE product_id NOT IN (SELECT id FROM products)")
        order = (db.execute("SELECT COALESCE(MAX(sort_order),0) m FROM categories").fetchone()["m"] or 0)
        for n, e, h in SUPERMARKET_CATEGORIES:
            order += 1
            db.execute("INSERT OR IGNORE INTO categories(name,emoji,active,show_home,sort_order) VALUES(?,?,1,?,?)", (n, e, h, order))
        for n in OLD_DEFAULT_CATEGORIES:
            db.execute("DELETE FROM categories WHERE name=? AND name NOT IN (SELECT DISTINCT category FROM products WHERE category IS NOT NULL)", (n,))
        for i, (n, e, h) in enumerate(SUPERMARKET_CATEGORIES, 1):    # ترتيب الأقسام الجديدة أولاً
            db.execute("UPDATE categories SET sort_order=? WHERE name=?", (i, n))
        if db.execute("SELECT COUNT(*) c FROM products").fetchone()["c"] == 0:
            db.executemany("""INSERT INTO products(name,description,price,old_price,stock,category,emoji,featured,flash,unit,min_order)
                              VALUES(?,?,?,?,?,?,?,?,?,?,?)""", SUPERMARKET_DEMO)
            _seed_sold(db)
        db.execute("INSERT INTO settings(key,value) VALUES('supermarket_v1','1')")
    db.commit()
    db.close()


# ================== أدوات مساعدة ==================
@app.context_processor
def inject_globals():
    cart = session.get("cart", {})
    cart_count = sum(cart.values()) if cart else 0
    cats, all_cats, emojis, wa_on = list(CATEGORIES), list(CATEGORIES), {}, False
    site = {"phone": SUPPORT_PHONE, "email": SUPPORT_EMAIL, "hours": "", "address": "", "whatsapp": WHATSAPP_NUMBER,
            "wa_support": SUPPORT_PHONE, "wa_on": False, "wa_link": ""}
    try:
        site = site_info()
    except Exception:
        pass
    try:
        rows = category_rows()
        all_cats = [r["name"] for r in rows]
        cats = [r["name"] for r in rows if r["active"]]
        emojis = {r["name"]: (r["emoji"] or "") for r in rows}
        wa_on = get_setting("whatsapp_cart", "1") == "1"
        if cart:   # عدّاد السلة يحسب المنتجات المتاحة بس (مش الأقسام الموقوفة)
            ids = [int(k) for k in cart if str(k).isdigit()]
            if ids:
                ok = get_db().execute("SELECT id FROM products WHERE id IN (%s) AND %s" %
                                      (",".join("?" * len(ids)), VISIBLE_SQL), ids).fetchall()
                cart_count = sum(cart.get(str(r["id"]), 0) for r in ok)
    except Exception:
        pass
    return dict(cart_count=cart_count, categories=cats, all_categories=all_cats,
                category_emoji=emojis, whatsapp_on=wa_on,
                is_admin=session.get("is_admin", 0), user_name=session.get("user_name"),
                coupon=session.get("coupon"), pay_methods=PAY_LABELS, site=site,
                canonical_url=canonical_url(), site_url=site_base_url(), units=UNITS)


def login_required(f):
    @wraps(f)
    def wrapper(*a, **kw):
        if "user_id" not in session:
            flash("سجّل دخولك الأول يا هندسة 🔑", "warning")
            return redirect(url_for("login", next=request.path))
        return f(*a, **kw)
    return wrapper


def admin_required(f):
    @wraps(f)
    def wrapper(*a, **kw):
        if not session.get("is_admin"):
            flash("الصفحة دي للإدارة بس 🚫", "danger")
            return redirect(url_for("home"))
        return f(*a, **kw)
    return wrapper


# ================== الأمان وأدوات مساعدة (جديد v5) ==================
VISIBLE_SQL = "COALESCE(category,'') NOT IN (SELECT name FROM categories WHERE active=0)"


def category_rows(active_only=False):
    sql = "SELECT * FROM categories" + (" WHERE active=1" if active_only else "") + " ORDER BY sort_order, id"
    return get_db().execute(sql).fetchall()


def category_is_active(name):
    """القسم الموقوف بيخفي منتجاته من الموقع. قسم مش مسجّل (قديم) بيفضل ظاهر."""
    r = get_db().execute("SELECT active FROM categories WHERE name=?", (name or "",)).fetchone()
    return True if r is None else bool(r["active"])


def csrf_token():
    if "_csrf" not in session:
        session["_csrf"] = secrets.token_hex(24)
    return session["_csrf"]


app.jinja_env.globals["csrf_token"] = csrf_token


def _tokens_match(a, b):
    return bool(a) and bool(b) and hmac.compare_digest(str(a), str(b))


def client_ip():
    if TRUST_PROXY:
        return request.headers.get("X-Forwarded-For", request.remote_addr or "").split(",")[0].strip()
    return request.remote_addr or ""


def to_int(v, default=0, lo=None, hi=None):
    try:
        n = int(str(v).strip())
    except (TypeError, ValueError):
        n = default
    if lo is not None and n < lo:
        n = lo
    if hi is not None and n > hi:
        n = hi
    return n


def to_float(v, default=None):
    try:
        n = float(str(v).strip())
    except (TypeError, ValueError):
        return default
    return n if math.isfinite(n) else default


def fmt_money(x):
    x = float(x or 0)
    return str(int(x)) if x == int(x) else ("%.2f" % x)


def safe_back(default):
    """رجوع للصفحة السابقة بس لو من نفس الموقع (يمنع open-redirect)"""
    ref = request.referrer
    if ref and urlparse(ref).netloc == request.host:
        return ref
    return default


def safe_next(nxt):
    if nxt and nxt.startswith("/") and not nxt.startswith("//") and "\\" not in nxt:
        return nxt
    return ""


def intl_number(num):
    """01222370381 -> 201222370381 (صيغة واتساب الدولية)"""
    d = re.sub(r"\D", "", num or "")
    if d.startswith("00"):
        d = d[2:]
    if d.startswith("0"):
        d = COUNTRY_CODE + d[1:]
    return d


ALLOWED_IMG_EXT = {".jpg", ".jpeg", ".png", ".gif", ".webp"}


def save_product_image(file_storage, prefix):
    """يقبل صور حقيقية بس (امتداد + توقيع الملف) ← وبيضغطها تلقائياً (WebP + نسخة مصغّرة للكروت)"""
    ext = os.path.splitext(file_storage.filename or "")[1].lower() or ".jpg"
    if ext not in ALLOWED_IMG_EXT:
        raise ValueError("صيغة الصورة غير مدعومة (JPG / PNG / WEBP / GIF بس) ❌")
    head = file_storage.stream.read(16)
    file_storage.stream.seek(0)
    ok = (head[:3] == b"\xff\xd8\xff" or head[:8] == b"\x89PNG\r\n\x1a\n" or
          head[:6] in (b"GIF87a", b"GIF89a") or (head[:4] == b"RIFF" and head[8:12] == b"WEBP"))
    if not ok:
        raise ValueError("الملف ده مش صورة صالحة ❌")
    rnd = os.urandom(2).hex()
    if Image is None:      # Pillow مش موجودة: حفظ زي ما هي
        name = "%s_%s%s" % (prefix, rnd, ext)
        file_storage.save(os.path.join(UPLOAD_DIR, name))
        return name
    try:
        im = Image.open(file_storage.stream)
        im.load()
        im = ImageOps.exif_transpose(im)
        if im.mode not in ("RGB", "RGBA"):
            im = im.convert("RGBA" if "transparency" in im.info or im.mode in ("LA", "PA", "P") else "RGB")
    except Exception:
        raise ValueError("الملف ده مش صورة صالحة ❌")
    stem = "%s_%s" % (prefix, rnd)
    main = im.copy()
    main.thumbnail((1200, 1200), Image.LANCZOS)
    main.save(os.path.join(UPLOAD_DIR, stem + ".webp"), "WEBP", quality=80, method=4)
    thumb = im.copy()
    thumb.thumbnail((480, 480), Image.LANCZOS)
    thumb.save(os.path.join(UPLOAD_DIR, stem + "_t.webp"), "WEBP", quality=75, method=4)
    return stem + ".webp"


def delete_product_image(name):
    """يمسح الصورة ونسختها المصغّرة من الديسك"""
    if not name:
        return
    stem = os.path.splitext(name)[0]
    for fn in (name, stem + "_t.webp"):
        try:
            os.remove(os.path.join(UPLOAD_DIR, os.path.basename(fn)))
        except OSError:
            pass


def img_src(name, thumb=False):
    """رابط الصورة (النسخة المصغّرة للكروت لو موجودة)"""
    if thumb and name:
        t = os.path.splitext(name)[0] + "_t.webp"
        if os.path.exists(os.path.join(UPLOAD_DIR, t)):
            name = t
    return url_for("static", filename="uploads/products/" + name)


def unit_label(p):
    """' / كيلو' للوحدات غير الافتراضية (فاضي لـ 'قطعة')"""
    try:
        u = p["unit"] or DEFAULT_UNIT
    except (KeyError, IndexError):
        u = DEFAULT_UNIT
    return (" / " + u) if u != DEFAULT_UNIT else ""


def static_v(filename):
    try:
        v = int(os.path.getmtime(os.path.join(BASE_DIR, "static", filename)))
    except OSError:
        v = 0
    return url_for("static", filename=filename, v=v)


app.jinja_env.globals.update(img_src=img_src, unit_label=unit_label, static_v=static_v)
app.jinja_env.filters["money"] = fmt_money


def parse_product_form(f, db):
    """يتحقق من بيانات المنتج ويرجّع (data, error)"""
    name = (f.get("name") or "").strip()
    if not name or len(name) > 150:
        return None, "اسم المنتج مطلوب (لحد 150 حرف) ❌"
    price = to_float(f.get("price"))
    if price is None or price <= 0 or price > 10000000:
        return None, "السعر لازم يكون رقم صحيح أكبر من صفر ❌"
    old_raw = (f.get("old_price") or "").strip()
    old_price = None
    if old_raw:
        old_price = to_float(old_raw)
        if old_price is None or old_price <= 0 or old_price > 10000000:
            return None, "السعر قبل الخصم غير صحيح ❌"
    stock = to_int(f.get("stock"), -1)
    if stock < 0 or stock > 1000000:
        return None, "المخزون لازم يكون رقم صفر أو أكتر ❌"
    category = (f.get("category") or "").strip()
    if not db.execute("SELECT 1 FROM categories WHERE name=?", (category,)).fetchone():
        return None, "القسم ده مش موجود ❌"
    unit = (f.get("unit") or DEFAULT_UNIT).strip()
    if unit not in UNITS:
        return None, "وحدة البيع غير صحيحة ❌"
    min_order = to_int(f.get("min_order"), 1)
    if min_order < 1 or min_order > 10000:
        return None, "الحد الأدنى للطلب لازم يكون رقم من 1 لـ 10000 ❌"
    return {"unit": unit, "min_order": min_order, "name": name, "description": (f.get("description") or "").strip()[:2000],
            "price": price, "old_price": old_price, "stock": stock, "category": category,
            "emoji": ((f.get("emoji") or "").strip() or "📦")[:4],
            "featured": 1 if f.get("featured") else 0, "flash": 1 if f.get("flash") else 0}, None


# --- حماية الدخول من التخمين ---
LOGIN_WINDOW_MIN = 15


def login_blocked(db, ip, email):
    w = "-%d minutes" % LOGIN_WINDOW_MIN
    q = lambda sql, p: db.execute(sql, p).fetchone()["c"]
    a = q("SELECT COUNT(*) c FROM login_attempts WHERE email=? AND ip=? AND created_at>=datetime('now',?)", (email, ip, w))
    b = q("SELECT COUNT(*) c FROM login_attempts WHERE ip=? AND created_at>=datetime('now',?)", (ip, w))
    c = q("SELECT COUNT(*) c FROM login_attempts WHERE email=? AND created_at>=datetime('now',?)", (email, w))
    return a >= 5 or b >= 20 or c >= 15


# --- حماية CSRF + التحقق من الحساب في كل طلب + هيدرز الأمان ---
# روابط GET بتنفّذ إجراء (حذف/موافقة/تفعيل) لازم تحمل توكن _t
GET_ACTION_ENDPOINTS = {"delete_product": None,
                        "admin_wallet": ("approve", "reject"),
                        "admin_affiliates": ("toggle", "confirm", "cancel"),
                        "admin_coupons": ("toggle",)}


@app.before_request
def security_gate():
    if request.endpoint in (None, "static"):
        return
    if request.method in ("POST", "PUT", "PATCH", "DELETE"):
        sent = request.form.get("csrf_token") or request.headers.get("X-CSRF-Token")
        if not _tokens_match(sent, session.get("_csrf")):
            flash("الطلب مش آمن أو الصفحة قديمة — جرّب تاني 🔒", "danger")
            return redirect(safe_back(url_for("home")))
    elif request.method == "GET" and request.endpoint in GET_ACTION_ENDPOINTS:
        params = GET_ACTION_ENDPOINTS[request.endpoint]
        if params is None or any(k in request.args for k in params):
            if not _tokens_match(request.args.get("_t"), session.get("_csrf")):
                flash("الطلب مش آمن — افتح الصفحة تاني وجرّب 🔒", "danger")
                return redirect(url_for("home"))
    if "user_id" in session:
        u = get_db().execute("SELECT active, is_admin FROM users WHERE id=?", (session["user_id"],)).fetchone()
        if (not u) or (not u["active"]):
            session.clear()
            flash("الحساب موقوف أو غير موجود 🚫", "danger")
            return redirect(url_for("home"))
        if session.get("is_admin") != u["is_admin"]:
            session["is_admin"] = u["is_admin"]


@app.after_request
def security_headers(resp):
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("X-Frame-Options", "DENY")
    resp.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    resp.headers.setdefault("Permissions-Policy", "geolocation=(), microphone=(), camera=()")
    resp.headers.setdefault("Content-Security-Policy",
                            "default-src 'self'; img-src 'self' data:; style-src 'self' 'unsafe-inline'; "
                            "script-src 'self' 'unsafe-inline'; connect-src 'self'; form-action 'self'; "
                            "frame-ancestors 'none'; base-uri 'self'")
    if app.config.get("SESSION_COOKIE_SECURE"):
        resp.headers.setdefault("Strict-Transport-Security", "max-age=31536000; includeSubDomains")
    if "user_id" in session and request.endpoint != "static":
        resp.headers["Cache-Control"] = "no-store, private"
    if request.path.startswith(NOINDEX_PREFIXES) or (request.endpoint == "shop" and (request.args.get("q") or request.args.get("sort"))):
        resp.headers["X-Robots-Tag"] = "noindex, nofollow"
    return maybe_gzip(resp)


@app.errorhandler(404)
def err_404(e):
    return render_template("error.html", code=404, msg="الصفحة دي مش موجودة 😕"), 404


@app.errorhandler(413)
def err_413(e):
    return render_template("error.html", code=413, msg="الملف أكبر من الحجم المسموح (8MB) 📁"), 413


@app.errorhandler(500)
def err_500(e):
    return render_template("error.html", code=500, msg="حصل خطأ غير متوقع — جرّب تاني بعد شوية 🛠️"), 500


def adjust_balance(db, user_id, amount, ttype, note=""):
    db.execute("UPDATE users SET balance = balance + ? WHERE id=?", (amount, user_id))
    db.execute("INSERT INTO wallet_transactions(user_id,amount,type,note) VALUES(?,?,?,?)",
               (user_id, amount, ttype, note))


def cart_details():
    """يرجع: items, subtotal, خصم الكوبون, خصم المسوق, الإجمالي"""
    db = get_db()
    cart = session.get("cart", {})
    items, subtotal = [], 0
    for pid, qty in cart.items():
        p = db.execute("SELECT * FROM products WHERE id=?", (pid,)).fetchone()
        if p and category_is_active(p["category"]):   # منتجات القسم الموقوف مش بتظهر في السلة
            items.append({"p": p, "qty": qty})
            subtotal += p["price"] * qty
    # كوبون عادي
    coupon_discount = 0
    coupon = session.get("coupon")
    if coupon:
        c = db.execute("SELECT * FROM coupons WHERE code=? AND active=1", (coupon,)).fetchone()
        if c and (not c["expires"] or c["expires"] >= str(date.today())):
            coupon_discount = round(subtotal * c["percent"] / 100, 2)
    # كود المسوق بالعمولة (يغطي الكوبون العادي)
    aff_discount = 0
    aff = validate_aff_code(session.get("aff_code"), quiet=True)
    if aff:
        aff_discount = round(subtotal * AFF_BUYER_DISCOUNT / 100, 2)
        coupon_discount = 0
    discount = max(coupon_discount, aff_discount)
    return items, subtotal, discount, round(subtotal - discount, 2)


# ============ مكافحة احتيال التسويق بالعمولة ============
def validate_aff_code(code, quiet=False, buyer_email=""):
    """يرجع marketer row أو None + يطبق قواعد مكافحة الاحتيال"""
    if not code:
        return None
    db = get_db()
    if get_setting("affiliate", "1") != "1":
        return None
    m = db.execute("""SELECT m.*, u.email, u.name FROM marketers m
                      JOIN users u ON u.id=m.user_id
                      WHERE m.code=? AND m.active=1""", (code.strip().upper(),)).fetchone()
    if not m:
        if not quiet:
            flash("كود المسوق غير صحيح أو موقوف ❌", "danger")
        return None
    # 1) المسوق ميعملش طلب بكوده الخاص
    if session.get("user_id") == m["user_id"]:
        if not quiet:
            flash("مينفعش تستخدم كودك الخاص في طلباتك 🚫", "danger")
        return None
    ip = client_ip()
    email = buyer_email or (session.get("user_email") or "")
    # 2) نفس الإيميل يستخدم نفس الكود 3 مرات بس في الشهر
    if email:
        cnt = db.execute("""SELECT COUNT(*) c FROM affiliate_orders
                            WHERE marketer_id=? AND buyer_email=? AND created_at>=datetime('now','-30 days')""",
                         (m["id"], email)).fetchone()["c"]
        if cnt >= 3:
            if not quiet:
                flash("الكود ده استُخدم بأقصى حد مسموح من حسابك 🛡️", "danger")
            return None
    # 3) نفس IP ميستخدمش نفس الكود مرتين خلال 24 ساعة
    dup_ip = db.execute("""SELECT 1 FROM affiliate_orders
                           WHERE marketer_id=? AND buyer_ip=? AND created_at>=datetime('now','-1 day')""",
                        (m["id"], ip)).fetchone()
    if dup_ip:
        if not quiet:
            flash("تم رصد نشاط مشبوه — الكود محظور مؤقتاً من هذا الجهاز 🛡️", "danger")
        return None
    return m


def process_matured_affiliates(db):
    """تعليق الكاش باك 14 يوم: اللي عدّى ولم يُرجع المنتج يتأكد ويتحول للرصيد"""
    rows = db.execute("""SELECT * FROM affiliate_orders
                         WHERE status='pending' AND hold_until<=?""",
                      (str(date.today()),)).fetchall()
    for a in rows:
        o = db.execute("SELECT status FROM orders WHERE id=?", (a["order_id"],)).fetchone()
        ret = db.execute("SELECT 1 FROM returns WHERE order_id=? AND status='approved'",
                         (a["order_id"],)).fetchone()
        if (not o) or o["status"] == "ملغي" or ret:
            db.execute("UPDATE affiliate_orders SET status='cancelled' WHERE id=?", (a["id"],))
        else:
            mk = db.execute("SELECT user_id FROM marketers WHERE id=?", (a["marketer_id"],)).fetchone()
            adjust_balance(db, mk["user_id"], a["commission"], "affiliate",
                           "عمولة طلب #%d (بعد فترة التعليق)" % a["order_id"])
            db.execute("UPDATE affiliate_orders SET status='confirmed' WHERE id=?", (a["id"],))


# ================== الصفحات العامة ==================
@app.route("/")
def home():
    db = get_db()
    if request.args.get("ref"):
        session["aff_code"] = request.args.get("ref", "").strip().upper()[:30]
        session.modified = True
        flash("كود الخصم %s هيتطبق تلقائياً على طلبك 🎉 خصم %d%%" % (session["aff_code"], AFF_BUYER_DISCOUNT), "success")
    featured = db.execute("SELECT * FROM products WHERE featured=1 AND " + VISIBLE_SQL + " LIMIT 8").fetchall()
    flash_deals = db.execute("SELECT * FROM products WHERE flash=1 AND old_price IS NOT NULL AND " + VISIBLE_SQL + " LIMIT 6").fetchall()
    best = db.execute("SELECT * FROM products WHERE " + VISIBLE_SQL + " ORDER BY sold DESC LIMIT 6").fetchall()
    latest = db.execute("SELECT * FROM products WHERE " + VISIBLE_SQL + " ORDER BY id DESC LIMIT 8").fetchall()
    # أقسام مستقلة في الرئيسية (زي المنتجات الغذائية) — بتظهر بس لو القسم مفعّل وفيه منتجات
    home_sections = []
    for c in db.execute("SELECT * FROM categories WHERE active=1 AND show_home=1 ORDER BY sort_order, id").fetchall():
        prods = db.execute("SELECT * FROM products WHERE category=? ORDER BY id DESC LIMIT 8", (c["name"],)).fetchall()
        if prods:
            home_sections.append({"name": c["name"], "emoji": c["emoji"] or "", "products": prods})
    return render_template("index.html", featured=featured, flash_deals=flash_deals,
                           best=best, latest=latest, home_sections=home_sections, ld=site_ld())


@app.route("/shop")
def shop():
    db = get_db()
    if request.args.get("ref"):
        session["aff_code"] = request.args.get("ref", "").strip().upper()[:30]
        session.modified = True
    q = request.args.get("q", "").strip()
    cat = request.args.get("cat", "").strip()
    sort = request.args.get("sort", "new")
    page = to_int(request.args.get("page", 1), 1, lo=1)
    per = 9
    sql = "SELECT * FROM products WHERE 1=1 AND " + VISIBLE_SQL
    params = []
    if q:
        sql += " AND (name LIKE ? OR description LIKE ?)"
        params += ["%" + q + "%", "%" + q + "%"]
    if cat:
        sql += " AND category = ?"
        params.append(cat)
    sql += {"new": " ORDER BY id DESC", "cheap": " ORDER BY price ASC",
            "exp": " ORDER BY price DESC", "name": " ORDER BY name",
            "best": " ORDER BY sold DESC"}.get(sort, " ORDER BY id DESC")
    total_rows = db.execute("SELECT COUNT(*) c FROM (" + sql + ")", params).fetchone()["c"]
    pages = max((total_rows + per - 1) // per, 1)
    page = min(page, pages)
    products = db.execute(sql + " LIMIT ? OFFSET ?", params + [per, (page - 1) * per]).fetchall()
    return render_template("shop.html", products=products, q=q, cat=cat, sort=sort,
                           page=page, pages=pages, total_rows=total_rows)


@app.route("/autocomplete")
def autocomplete():
    q = request.args.get("q", "").strip()
    if len(q) < 2:
        return jsonify([])
    rows = get_db().execute("SELECT id,name FROM products WHERE name LIKE ? AND " + VISIBLE_SQL + " LIMIT 6",
                            ("%" + q + "%",)).fetchall()
    return jsonify([{"id": r["id"], "name": r["name"]} for r in rows])


@app.route("/product/<int:pid>")
def product(pid):
    db = get_db()
    p = db.execute("SELECT * FROM products WHERE id=?", (pid,)).fetchone()
    if not p:
        flash("المنتج مش موجود 😕", "danger")
        return redirect(url_for("shop"))
    if not category_is_active(p["category"]):
        flash("القسم ده متوقف مؤقتاً ⏸️", "warning")
        return redirect(url_for("shop"))
    related = db.execute("SELECT * FROM products WHERE category=? AND id!=? LIMIT 4", (p["category"], pid)).fetchall()
    in_wish = False
    if "user_id" in session:
        in_wish = bool(db.execute("SELECT 1 FROM wishlist WHERE user_id=? AND product_id=?",
                                  (session["user_id"], pid)).fetchone())
    return render_template("product.html", p=p, related=related, in_wish=in_wish, ld=product_ld(p))


# ================== التسويق بالعمولة ==================
@app.route("/affiliate")
@login_required
def affiliate():
    db = get_db()
    if get_setting("affiliate", "1") != "1":
        flash("قسم التسويق بالعمولة متوقف مؤقتاً ⏸️", "warning")
        return redirect(url_for("home"))
    m = db.execute("SELECT * FROM marketers WHERE user_id=?", (session["user_id"],)).fetchone()
    process_matured_affiliates(db)
    db.commit()
    stats, rows = None, []
    if m:
        stats = {
            "code": m["code"],
            "active": m["active"],
            "orders": db.execute("SELECT COUNT(*) c FROM affiliate_orders WHERE marketer_id=?", (m["id"],)).fetchone()["c"],
            "pending": db.execute("SELECT COALESCE(SUM(commission),0) s FROM affiliate_orders WHERE marketer_id=? AND status='pending'", (m["id"],)).fetchone()["s"],
            "confirmed": db.execute("SELECT COALESCE(SUM(commission),0) s FROM affiliate_orders WHERE marketer_id=? AND status='confirmed'", (m["id"],)).fetchone()["s"],
            "link": url_for("home", ref=m["code"], _external=True),
        }
        rows = db.execute("""SELECT a.*, o.status ostatus FROM affiliate_orders a
                             JOIN orders o ON o.id=a.order_id
                             WHERE a.marketer_id=? ORDER BY a.id DESC LIMIT 30""", (m["id"],)).fetchall()
    return render_template("affiliate.html", m=m, stats=stats, rows=rows, hold_days=AFF_HOLD_DAYS)


@app.route("/affiliate/join", methods=["POST"])
@login_required
def affiliate_join():
    db = get_db()
    if get_setting("affiliate", "1") != "1":
        flash("القسم متوقف مؤقتاً ⏸️", "warning")
        return redirect(url_for("affiliate"))
    if db.execute("SELECT 1 FROM marketers WHERE user_id=?", (session["user_id"],)).fetchone():
        flash("انت مسجل مسوق بالفعل ✅", "info")
        return redirect(url_for("affiliate"))
    code = "HC-" + secrets.token_hex(3).upper()
    db.execute("INSERT INTO marketers(user_id,code) VALUES(?,?)", (session["user_id"], code))
    db.commit()
    flash("تم تسجيلك! كودك: %s — هيشتغل بعد موافقة الإدارة ⏳" % code, "success")
    return redirect(url_for("affiliate"))


@app.route("/affiliate/apply_code", methods=["POST"])
def affiliate_apply_code():
    code = request.form.get("code", "").strip().upper()
    if validate_aff_code(code):
        session["aff_code"] = code
        session.modified = True
        flash("كود '%s' اتفعّل! خصم %d%% على طلبك 🎉" % (code, AFF_BUYER_DISCOUNT), "success")
    return redirect(safe_back(url_for("cart")))


@app.route("/affiliate/remove_code")
def affiliate_remove_code():
    session.pop("aff_code", None)
    session.modified = True
    return redirect(safe_back(url_for("cart")))


# ================== المحفظة ==================
@app.route("/wallet")
@login_required
def wallet():
    db = get_db()
    u = db.execute("SELECT balance FROM users WHERE id=?", (session["user_id"],)).fetchone()
    reqs = db.execute("SELECT * FROM wallet_requests WHERE user_id=? ORDER BY id DESC LIMIT 20",
                      (session["user_id"],)).fetchall()
    txs = db.execute("SELECT * FROM wallet_transactions WHERE user_id=? ORDER BY id DESC LIMIT 20",
                     (session["user_id"],)).fetchall()
    return render_template("wallet.html", balance=u["balance"], reqs=reqs, txs=txs,
                           methods=recharge_methods(only_on=True),
                           sender_on=get_setting("wallet_sender", "1") == "1")


@app.route("/wallet/request", methods=["POST"])
@login_required
def wallet_request():
    amount = to_float(request.form.get("amount", 0), 0) or 0
    method = request.form.get("method", "")
    if amount <= 0 or amount > 100000:
        flash("اكتب مبلغ صحيح 💰", "danger")
        return redirect(url_for("wallet"))
    if method not in [m["key"] for m in recharge_methods(only_on=True)]:
        flash("وسيلة التحويل دي مش متاحة حالياً — اختار وسيلة تانية ❌", "danger")
        return redirect(url_for("wallet"))
    return render_template("wallet_transfer.html", amount=amount, method=method,
                           method_name=PAY_METHODS[method], transfer_number=transfer_number_for(method),
                           sender_on=get_setting("wallet_sender", "1") == "1")


@app.route("/wallet/confirm", methods=["POST"])
@login_required
def wallet_confirm():
    amount = to_float(request.form.get("amount", 0), 0) or 0
    method = request.form.get("method", "")
    sender = request.form.get("sender_number", "").strip()
    if amount <= 0 or amount > 100000 or method not in [m["key"] for m in recharge_methods(only_on=True)]:
        flash("بيانات الشحن غير صحيحة أو الوسيلة متوقفة ❌", "danger")
        return redirect(url_for("wallet"))
    if get_setting("wallet_sender", "1") == "1":
        if not re.match(r"^01[0-9]{9}$", sender):
            flash("اكتب رقم المحوّل منه بشكل صحيح (11 رقم يبدأ بـ 01) 📱", "danger")
            return redirect(url_for("wallet"))
    db = get_db()
    db.execute("INSERT INTO wallet_requests(user_id,amount,method,sender_number) VALUES(?,?,?,?)",
               (session["user_id"], amount, method, sender))
    db.commit()
    flash("تم استلام طلبك ✅ الإدارة هتراجع رقم المحوّل وتضيف الرصيد", "success")
    return redirect(url_for("wallet"))


# ================== المفضلة ==================
@app.route("/wishlist")
@login_required
def wishlist():
    rows = get_db().execute("SELECT p.* FROM wishlist w JOIN products p ON p.id=w.product_id WHERE w.user_id=? AND COALESCE(p.category,'') NOT IN (SELECT name FROM categories WHERE active=0)", (session["user_id"],)).fetchall()
    return render_template("wishlist.html", products=rows)


@app.route("/wish/toggle/<int:pid>")
def toggle_wish(pid):
    if "user_id" not in session:
        flash("سجّل دخولك الأول عشان تضيف للمفضلة ❤️", "warning")
        return redirect(url_for("login", next=safe_back(url_for("home"))))
    db = get_db()
    if db.execute("SELECT 1 FROM wishlist WHERE user_id=? AND product_id=?",
                  (session["user_id"], pid)).fetchone():
        db.execute("DELETE FROM wishlist WHERE user_id=? AND product_id=?", (session["user_id"], pid))
        flash("اتشال من المفضلة 💔", "info")
    else:
        db.execute("INSERT INTO wishlist(user_id,product_id) VALUES(?,?)", (session["user_id"], pid))
        flash("اتضاف للمفضلة ❤️", "success")
    db.commit()
    return redirect(safe_back(url_for("wishlist")))


# ================== السلة والكوبونات ==================
@app.route("/cart")
def cart():
    items, subtotal, discount, total = cart_details()
    db = get_db()
    aff = None
    if session.get("aff_code"):
        aff = db.execute("""SELECT m.*, u.name FROM marketers m JOIN users u ON u.id=m.user_id
                            WHERE m.code=? AND m.active=1""", (session["aff_code"],)).fetchone()
    return render_template("cart.html", items=items, subtotal=subtotal,
                           discount=discount, total=total, aff=aff,
                           ship=shipping_quote(total, session.get("gov")), zones=active_zones())


@app.route("/cart/gov", methods=["POST"])
def cart_set_gov():
    """اختيار المحافظة في السلة (لحساب الشحن)"""
    gov = (request.form.get("governorate") or "").strip()
    if gov and get_db().execute("SELECT 1 FROM shipping_zones WHERE governorate=? AND active=1", (gov,)).fetchone():
        session["gov"] = gov
    elif not gov:
        session.pop("gov", None)
    session.modified = True
    return redirect(safe_back(url_for("cart")))


@app.route("/add/<int:pid>", methods=["POST"])
def add_to_cart(pid):
    db = get_db()
    p = db.execute("SELECT * FROM products WHERE id=?", (pid,)).fetchone()
    if not p or not category_is_active(p["category"]):
        flash("المنتج ده مش متاح حالياً 😕", "warning")
        return redirect(safe_back(url_for("shop")))
    mn = p["min_order"] or 1
    unit = p["unit"] or DEFAULT_UNIT
    if p["stock"] <= 0:
        flash("المنتج ده سيعود قريباً ⏳", "warning")
        return redirect(safe_back(url_for("shop")))
    if p["stock"] < mn:
        flash("المتاح (%d %s) أقل من الحد الأدنى للطلب (%d %s) 😕" % (p["stock"], unit, mn, unit), "warning")
        return redirect(safe_back(url_for("shop")))
    qty = to_int(request.form.get("qty", 1), 1, lo=1, hi=99)
    cart = session.get("cart", {})
    new_qty = min(cart.get(str(pid), 0) + qty, 99)
    if new_qty < mn:
        new_qty = mn
        flash("الحد الأدنى لطلب «%s» هو %d %s — اتظبطت الكمية ⚠️" % (p["name"], mn, unit), "info")
    if new_qty > p["stock"]:
        new_qty = p["stock"]
        flash("المتاح من المنتج ده %d %s بس — اتظبطت الكمية ⚠️" % (p["stock"], unit), "warning")
    cart[str(pid)] = new_qty
    session["cart"] = cart
    session.modified = True
    flash("اتضاف للسلة ✅", "success")
    return redirect(safe_back(url_for("cart")))


@app.route("/update_cart/<int:pid>", methods=["POST"])
def update_cart(pid):
    qty = to_int(request.form.get("qty", 1), 1, hi=99)
    cart = session.get("cart", {})
    p = get_db().execute("SELECT stock, min_order, unit FROM products WHERE id=?", (pid,)).fetchone()
    if qty > 0 and p:
        mn, unit = p["min_order"] or 1, p["unit"] or DEFAULT_UNIT
        if qty < mn:
            qty = mn
            flash("الحد الأدنى للطلب %d %s — لو عايز تشيله دوس ✖ ⚠️" % (mn, unit), "info")
        if qty > p["stock"]:
            qty = max(p["stock"], 0)
            flash("المتاح %d %s بس ⚠️" % (qty, unit), "warning")
            if qty < mn:
                qty = 0
    if qty <= 0:
        cart.pop(str(pid), None)
    else:
        cart[str(pid)] = qty
    session["cart"] = cart
    session.modified = True
    return redirect(url_for("cart"))


@app.route("/remove/<int:pid>")
def remove_from_cart(pid):
    cart = session.get("cart", {})
    cart.pop(str(pid), None)
    session["cart"] = cart
    session.modified = True
    return redirect(url_for("cart"))


@app.route("/cart/whatsapp", methods=["GET", "POST"])
def cart_whatsapp():
    """يسجّل الطلب في لوحة الإدارة وبعدها يحوّل كل منتجات السلة لرسالة واتساب جاهزة"""
    if get_setting("whatsapp_cart", "1") != "1":
        flash("الطلب عبر واتساب متوقف مؤقتاً ⏸️", "warning")
        return redirect(url_for("cart"))
    items, subtotal, discount, total = cart_details()
    if not items:
        flash("سلتك فاضية 🛒", "warning")
        return redirect(url_for("shop"))
    lines = [{"id": it["p"]["id"], "name": it["p"]["name"], "unit": it["p"]["unit"] or DEFAULT_UNIT,
              "qty": it["qty"], "price": it["p"]["price"]} for it in items]
    return wa_redirect(lines, subtotal, discount, "checkout" if request.method == "POST" else "cart")


@app.route("/product/<int:pid>/whatsapp")
def product_whatsapp(pid):
    """زرار واتساب في صفحة المنتج: طلب المنتج ده لوحده بالكمية المختارة"""
    if get_setting("whatsapp_cart", "1") != "1":
        flash("الطلب عبر واتساب متوقف مؤقتاً ⏸️", "warning")
        return redirect(url_for("product", pid=pid))
    p = get_db().execute("SELECT * FROM products WHERE id=?", (pid,)).fetchone()
    if not p or not category_is_active(p["category"]):
        flash("المنتج ده مش متاح حالياً 😕", "warning")
        return redirect(url_for("shop"))
    mn, unit = p["min_order"] or 1, p["unit"] or DEFAULT_UNIT
    if p["stock"] < max(mn, 1):
        flash("المنتج غير متاح بالكمية المطلوبة حالياً 😕", "warning")
        return redirect(url_for("product", pid=pid))
    qty = max(to_int(request.args.get("qty", mn), mn, lo=1, hi=99), mn)
    qty = min(qty, p["stock"])
    lines = [{"id": p["id"], "name": p["name"], "unit": unit, "qty": qty, "price": p["price"]}]
    return wa_redirect(lines, p["price"] * qty, 0, "product", product_url=url_for("product", pid=pid, _external=True))


@app.route("/apply_coupon", methods=["POST"])
def apply_coupon():
    code = request.form.get("code", "").strip().upper()
    c = get_db().execute("SELECT * FROM coupons WHERE code=? AND active=1", (code,)).fetchone()
    if c and (not c["expires"] or c["expires"] >= str(date.today())):
        session["coupon"] = code
        flash("كوبون '%s' اتفعّل! خصم %d%% 🎉" % (code, c["percent"]), "success")
    else:
        session.pop("coupon", None)
        flash("الكوبون ده مش شغال أو منتهي ❌", "danger")
    session.modified = True
    return redirect(url_for("cart"))


@app.route("/remove_coupon")
def remove_coupon():
    session.pop("coupon", None)
    session.modified = True
    return redirect(url_for("cart"))


# ================== الطلبات والدفع ==================
@app.route("/checkout", methods=["GET", "POST"])
@login_required
def checkout():
    db = get_db()
    items, subtotal, discount, total = cart_details()
    if not items:
        flash("سلتك فاضية 🛒", "warning")
        return redirect(url_for("shop"))
    u = db.execute("SELECT * FROM users WHERE id=?", (session["user_id"],)).fetchone()
    aff = None
    if session.get("aff_code"):
        aff = validate_aff_code(session["aff_code"], quiet=True)
    zones = active_zones()
    if request.method == "POST":
        address = request.form.get("address", "").strip()
        phone = request.form.get("phone", "").strip()
        gov = (request.form.get("governorate") or "").strip()
        if not address or len(address) > 300 or not re.match(r"^[0-9+\-\s]{7,20}$", phone):
            flash("اكتب عنوان ورقم موبايل صحيحين 📍", "danger")
            return redirect(url_for("checkout"))
        for it in items:   # التأكد من المخزون والحد الأدنى قبل تأكيد الطلب
            unit = it["p"]["unit"] or DEFAULT_UNIT
            if it["qty"] > it["p"]["stock"]:
                flash("الكمية المطلوبة من «%s» أكبر من المتاح (%d) ❌" % (it["p"]["name"], it["p"]["stock"]), "danger")
                return redirect(url_for("cart"))
            if it["qty"] < (it["p"]["min_order"] or 1):
                flash("الحد الأدنى لطلب «%s» هو %d %s ❌" % (it["p"]["name"], it["p"]["min_order"], unit), "danger")
                return redirect(url_for("cart"))
        # التحقق من كود المسوق قبل حساب الإجمالي
        aff_row = None
        if session.get("aff_code"):
            aff_row = validate_aff_code(session["aff_code"], buyer_email=u["email"])
            if not aff_row:
                session.pop("aff_code", None)
                items, subtotal, discount, total = cart_details()
        # الشحن حسب المحافظة (والمجاني فوق الحد اللي الإدارة محدده)
        zone_names = [z["governorate"] for z in zones]
        shipping_on = get_setting("shipping_enabled", "1") == "1"
        if zone_names and (shipping_on and gov not in zone_names):
            flash("اختار المحافظة عشان نحسب الشحن 🚚", "danger")
            return redirect(url_for("checkout"))
        if gov not in zone_names:
            gov = ""
        ship = shipping_quote(total, gov)
        shipping = ship["fee"]
        grand = round(total + shipping, 2)
        method = request.form.get("payment", "")
        if not checkout_methods().get(method):
            flash("اختار طريقة دفع متاحة 💳", "danger")
            return redirect(url_for("checkout"))
        paid = 0.0
        if method == "balance":
            if u["balance"] < grand:
                flash("رصيدك مش كافي! محتاج تشتري رصيد من صفحة المحفظة 💰", "danger")
                return redirect(url_for("wallet"))
            paid = grand
        elif method == "partial":
            paid = to_float(request.form.get("partial_amount", 0), 0) or 0
            if paid <= 0 or paid >= grand:
                flash("اكتب مبلغ الدفع الجزئي (أقل من الإجمالي) ❌", "danger")
                return redirect(url_for("checkout"))
            if paid > u["balance"]:
                flash("المبلغ أكبر من رصيدك المتاح (%.2f ج.م) 💰" % u["balance"], "danger")
                return redirect(url_for("checkout"))
        cur = db.execute("""INSERT INTO orders(user_id,total,discount,coupon,payment_method,paid,address,phone,shipping,governorate)
                            VALUES(?,?,?,?,?,?,?,?,?,?)""",
                         (session["user_id"], grand, discount,
                          session.get("aff_code") or session.get("coupon"),
                          method, paid, address, phone, shipping, gov))
        oid = cur.lastrowid
        for it in items:
            db.execute("INSERT INTO order_items(order_id,product_id,qty,price) VALUES(?,?,?,?)",
                       (oid, it["p"]["id"], it["qty"], it["p"]["price"]))
            db.execute("UPDATE products SET stock=stock-?, sold=sold+? WHERE id=?",
                       (it["qty"], it["qty"], it["p"]["id"]))
        if paid > 0:
            adjust_balance(db, session["user_id"], -paid, "purchase", "دفع طلب #%d" % oid)
        # تسجيل عمولة المسوق (معلّقة 14 يوم) — على قيمة المنتجات بدون الشحن
        if aff_row:
            ip = client_ip()
            comm = round(total * AFF_COMMISSION / 100, 2)
            hold = (date.today() + timedelta(days=AFF_HOLD_DAYS)).isoformat()
            db.execute("""INSERT INTO affiliate_orders(order_id,marketer_id,buyer_email,buyer_ip,
                          order_total,buyer_discount,commission,hold_until)
                          VALUES(?,?,?,?,?,?,?,?)""",
                       (oid, aff_row["id"], u["email"], ip, total, discount, comm, hold))
        db.commit()
        session["cart"] = {}
        session["gov"] = gov
        session.pop("coupon", None)
        session.pop("aff_code", None)
        session.modified = True
        flash("تم تأكيد طلبك رقم #%d 🎉" % oid, "success")
        return redirect(url_for("order_done", oid=oid))
    gov = session.get("gov") if session.get("gov") in [z["governorate"] for z in zones] else ""
    ship = shipping_quote(total, gov)
    return render_template("checkout.html", items=items, subtotal=subtotal,
                           discount=discount, total=total, u=u, balance=u["balance"], aff=aff,
                           zones=zones, ship=ship, gov=gov, co=checkout_methods(),
                           grand=round(total + ship["fee"], 2))


@app.route("/order_done/<int:oid>")
@login_required
def order_done(oid):
    o = get_db().execute("SELECT user_id FROM orders WHERE id=?", (oid,)).fetchone()
    if not o or (o["user_id"] != session["user_id"] and not session.get("is_admin")):
        flash("الطلب مش متاح 🚫", "danger")
        return redirect(url_for("home"))
    return render_template("order_done.html", oid=oid)


@app.route("/invoice/<int:oid>")
@login_required
def invoice(oid):
    db = get_db()
    o = db.execute("SELECT o.*, u.name uname, u.email uemail FROM orders o LEFT JOIN users u ON u.id=o.user_id WHERE o.id=?", (oid,)).fetchone()
    if not o or (o["user_id"] != session["user_id"] and not session.get("is_admin")):
        flash("الفاتورة مش متاحة 🚫", "danger")
        return redirect(url_for("home"))
    items = db.execute("""SELECT oi.*, p.name, p.emoji, p.unit FROM order_items oi
                          JOIN products p ON p.id=oi.product_id WHERE oi.order_id=?""", (oid,)).fetchall()
    aff = db.execute("""SELECT a.*, m.code FROM affiliate_orders a JOIN marketers m ON m.id=a.marketer_id
                        WHERE a.order_id=?""", (oid,)).fetchone()
    return render_template("invoice.html", o=o, items=items, aff=aff,
                           pay_methods=PAY_LABELS, now=datetime.now().strftime("%Y-%m-%d %H:%M"))


@app.route("/my_orders")
@login_required
def my_orders():
    db = get_db()
    orders = db.execute("SELECT * FROM orders WHERE user_id=? ORDER BY id DESC",
                        (session["user_id"],)).fetchall()
    details, returns_map = {}, {}
    for o in orders:
        details[o["id"]] = db.execute(
            "SELECT oi.*, p.name, p.emoji, p.image, p.unit FROM order_items oi JOIN products p ON p.id=oi.product_id WHERE oi.order_id=?",
            (o["id"],)).fetchall()
        returns_map[o["id"]] = db.execute("SELECT * FROM returns WHERE order_id=?", (o["id"],)).fetchone()
    return render_template("my_orders.html", orders=orders, details=details,
                           returns_map=returns_map, return_days=RETURN_DAYS,
                           steps=["قيد التجهيز", "تم الشحن", "تم التوصيل"],
                           pay_methods=PAY_LABELS)


# ================== الإرجاع ==================
@app.route("/returns")
@login_required
def returns():
    if get_setting("returns", "1") != "1":
        flash("الإرجاع متوقف مؤقتاً ⏸️", "warning")
        return redirect(url_for("home"))
    db = get_db()
    rows = db.execute("""SELECT r.*, o.total ototal, o.status ostatus FROM returns r
                         JOIN orders o ON o.id=r.order_id
                         WHERE r.user_id=? ORDER BY r.id DESC""", (session["user_id"],)).fetchall()
    return render_template("returns.html", rows=rows)


@app.route("/returns/request/<int:oid>", methods=["POST"])
@login_required
def return_request(oid):
    if get_setting("returns", "1") != "1":
        flash("الإرجاع متوقف مؤقتاً ⏸️", "warning")
        return redirect(url_for("home"))
    db = get_db()
    o = db.execute("SELECT * FROM orders WHERE id=? AND user_id=?", (oid, session["user_id"])).fetchone()
    reason = request.form.get("reason", "").strip()
    if not o:
        flash("الطلب مش موجود 🚫", "danger")
    elif o["status"] != "تم التوصيل":
        flash("الإرجاع متاح بعد التوصيل بس 📦", "danger")
    elif o["created_at"][:10] < (date.today() - timedelta(days=RETURN_DAYS)).isoformat():
        flash("عدّت مدة الإرجاع (%d يوم) ⏰" % RETURN_DAYS, "danger")
    elif db.execute("SELECT 1 FROM returns WHERE order_id=?", (oid,)).fetchone():
        flash("فيه طلب إرجاع موجود بالفعل لهذا الطلب ⏳", "info")
    elif not reason:
        flash("اكتب سبب الإرجاع ✍️", "danger")
    else:
        db.execute("INSERT INTO returns(order_id,user_id,reason) VALUES(?,?,?)", (oid, session["user_id"], reason))
        db.commit()
        flash("تم إرسال طلب الإرجاع ✅ هيتم المراجعة من الإدارة", "success")
    return redirect(url_for("my_orders"))


# ================== الحسابات ==================
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        name, email = request.form["name"].strip(), request.form["email"].lower().strip()
        pw = request.form["password"]
        db = get_db()
        if not name or len(name) > 80:
            flash("اكتب اسمك (لحد 80 حرف) ✍️", "danger")
            return redirect(url_for("register"))
        if len(pw) < 8:
            flash("كلمة السر لازم تكون 8 حروف على الأقل 🔒", "danger")
            return redirect(url_for("register"))
        if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
            flash("الإيميل مش صحيح 📧", "danger")
            return redirect(url_for("register"))
        if db.execute("SELECT id FROM users WHERE email=?", (email,)).fetchone():
            flash("الإيميل مسجل قبل كده 📧", "danger")
            return redirect(url_for("register"))
        db.execute("INSERT INTO users(name,email,password) VALUES(?,?,?)",
                   (name, email, generate_password_hash(pw)))
        db.commit()
        flash("تم إنشاء حسابك! سجّل دخولك دلوقتي 🎉", "success")
        return redirect(url_for("login"))
    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form["email"].lower().strip()
        db = get_db()
        ip = client_ip()
        if login_blocked(db, ip, email):   # حماية من تخمين كلمات السر
            flash("محاولات دخول غلط كتير 🛡️ استنى %d دقيقة وحاول تاني" % LOGIN_WINDOW_MIN, "danger")
            return redirect(url_for("login"))
        u = db.execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()
        if u and check_password_hash(u["password"], request.form["password"]):
            if not u["active"]:
                flash("الحساب موقوف — تواصل مع الإدارة 🚫", "danger")
                return redirect(url_for("login"))
            db.execute("DELETE FROM login_attempts WHERE email=? AND ip=?", (email, ip))
            db.commit()
            session.update(user_id=u["id"], user_name=u["name"], is_admin=u["is_admin"],
                           user_email=u["email"])
            flash("أهلاً بيك يا " + u["name"] + " 👋", "success")
            return redirect(safe_next(request.args.get("next")) or url_for("home"))
        db.execute("INSERT INTO login_attempts(ip,email) VALUES(?,?)", (ip, email))
        db.execute("DELETE FROM login_attempts WHERE created_at < datetime('now','-1 day')")
        db.commit()
        flash("بيانات الدخول غلط ❌", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("تم تسجيل الخروج، مع السلامة 👋", "info")
    return redirect(url_for("home"))


@app.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    db = get_db()
    u = db.execute("SELECT * FROM users WHERE id=?", (session["user_id"],)).fetchone()
    if request.method == "POST":
        f = request.form
        if f.get("new_password"):
            if not check_password_hash(u["password"], f["current_password"]):
                flash("كلمة السر الحالية غلط ❌", "danger")
                return redirect(url_for("profile"))
            if len(f["new_password"]) < 8:
                flash("كلمة السر الجديدة لازم تكون 8 حروف على الأقل 🔒", "danger")
                return redirect(url_for("profile"))
            db.execute("UPDATE users SET password=? WHERE id=?",
                       (generate_password_hash(f["new_password"]), session["user_id"]))
        db.execute("UPDATE users SET name=?, phone=?, address=? WHERE id=?",
                   (f["name"], f.get("phone"), f.get("address"), session["user_id"]))
        db.commit()
        session["user_name"] = f["name"]
        flash("تم تحديث بياناتك ✅", "success")
        return redirect(url_for("profile"))
    stats = {
        "orders": db.execute("SELECT COUNT(*) c FROM orders WHERE user_id=?", (session["user_id"],)).fetchone()["c"],
        "spent": db.execute("SELECT COALESCE(SUM(total),0) s FROM orders WHERE user_id=?", (session["user_id"],)).fetchone()["s"],
        "wish": db.execute("SELECT COUNT(*) c FROM wishlist WHERE user_id=?", (session["user_id"],)).fetchone()["c"],
        "balance": u["balance"],
    }
    return render_template("profile.html", u=u, stats=stats)


@app.route("/profile/export")
@login_required
def profile_export():
    """تصدير بياناتي (حق حماية البيانات)"""
    db = get_db()
    u = db.execute("SELECT id,name,email,phone,address,balance,created_at FROM users WHERE id=?", (session["user_id"],)).fetchone()
    orders = db.execute("SELECT id,total,status,created_at FROM orders WHERE user_id=?", (session["user_id"],)).fetchall()
    data = {"بيانات_الحساب": dict(u), "الطلبات": [dict(o) for o in orders]}
    return Response(__import__("json").dumps(data, ensure_ascii=False, indent=2),
                    mimetype="application/json",
                    headers={"Content-Disposition": "attachment; filename=my_data.json"})


@app.route("/profile/delete", methods=["POST"])
@login_required
def profile_delete():
    """حذف حسابي نهائياً"""
    db = get_db()
    uid = session["user_id"]
    if session.get("is_admin"):
        flash("حساب الإدارة مينفعش يتحذف 🚫", "danger")
        return redirect(url_for("profile"))
    pw = request.form.get("password", "")
    u = db.execute("SELECT password FROM users WHERE id=?", (uid,)).fetchone()
    if not check_password_hash(u["password"], pw):
        flash("كلمة السر غلط — الحذف متوقف ❌", "danger")
        return redirect(url_for("profile"))
    delete_user_data(db, uid)
    db.commit()
    session.clear()
    flash("تم حذف حسابك نهائياً — بنتمنى نشوفك تاني 🥺", "info")
    return redirect(url_for("home"))


# ================== صفحات إضافية ==================
@app.route("/about")
def about():
    return render_template("about.html")


# ================== لوحة تحكم الإدارة ==================
@app.route("/admin")
@admin_required
def admin():
    db = get_db()
    stats = {
        "products": db.execute("SELECT COUNT(*) c FROM products").fetchone()["c"],
        "orders": db.execute("SELECT COUNT(*) c FROM orders").fetchone()["c"],
        "users": db.execute("SELECT COUNT(*) c FROM users WHERE is_admin=0").fetchone()["c"],
        "revenue": db.execute("SELECT COALESCE(SUM(total),0) s FROM orders WHERE status!='ملغي'").fetchone()["s"],
        "pending_wallet": db.execute("SELECT COUNT(*) c FROM wallet_requests WHERE status='pending'").fetchone()["c"],
        "pending_returns": db.execute("SELECT COUNT(*) c FROM returns WHERE status='pending'").fetchone()["c"],
        "pending_aff": db.execute("SELECT COUNT(*) c FROM affiliate_orders WHERE status='pending'").fetchone()["c"],
        "pending_marketers": db.execute("SELECT COUNT(*) c FROM marketers WHERE active=0").fetchone()["c"],
        "balances_sum": db.execute("SELECT COALESCE(SUM(balance),0) s FROM users WHERE is_admin=0").fetchone()["s"],
        "wa_new": db.execute("SELECT COUNT(*) c FROM whatsapp_orders WHERE status='جديد'").fetchone()["c"],
        "tickets_new": db.execute("SELECT COUNT(*) c FROM support_tickets WHERE status='جديدة'").fetchone()["c"],
    }
    recent = db.execute("SELECT * FROM orders ORDER BY id DESC LIMIT 10").fetchall()
    low_stock = db.execute("SELECT * FROM products WHERE stock<=5 ORDER BY stock").fetchall()
    return render_template("admin/dashboard.html", stats=stats, recent=recent, low_stock=low_stock)


# ----- المنتجات -----
@app.route("/admin/products", methods=["GET", "POST"])
@admin_required
def admin_products():
    db = get_db()
    if request.method == "POST":
        f = request.form
        d, err = parse_product_form(f, db)
        if err:
            flash(err, "danger")
            return redirect(url_for("admin_products"))
        image = ""
        img = request.files.get("image")
        if img and img.filename:
            try:
                image = save_product_image(img, "p%d" % int(datetime.now().timestamp()))
            except ValueError as e:
                flash(str(e), "danger")
                return redirect(url_for("admin_products"))
        db.execute("INSERT INTO products(name,description,price,old_price,stock,category,emoji,image,featured,flash,unit,min_order) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                   (d["name"], d["description"], d["price"], d["old_price"],
                    d["stock"], d["category"], d["emoji"], image,
                    d["featured"], d["flash"], d["unit"], d["min_order"]))
        db.commit()
        flash("تم إضافة المنتج ✅ ظهر فوراً على الموقع 🚀", "success")
        return redirect(url_for("admin_products"))
    products = db.execute("SELECT * FROM products ORDER BY id DESC").fetchall()
    return render_template("admin/products.html", products=products)


@app.route("/admin/product/edit/<int:pid>", methods=["GET", "POST"])
@admin_required
def edit_product(pid):
    db = get_db()
    p = db.execute("SELECT * FROM products WHERE id=?", (pid,)).fetchone()
    if not p:
        return redirect(url_for("admin_products"))
    if request.method == "POST":
        f = request.form
        d, err = parse_product_form(f, db)
        if err:
            flash(err, "danger")
            return redirect(url_for("edit_product", pid=pid))
        image = p["image"]
        img = request.files.get("image")
        if img and img.filename:
            try:
                image = save_product_image(img, "p%d" % pid)
            except ValueError as e:
                flash(str(e), "danger")
                return redirect(url_for("edit_product", pid=pid))
        db.execute("UPDATE products SET name=?,description=?,price=?,old_price=?,stock=?,category=?,emoji=?,image=?,featured=?,flash=?,unit=?,min_order=? WHERE id=?",
                   (d["name"], d["description"], d["price"], d["old_price"],
                    d["stock"], d["category"], d["emoji"], image,
                    d["featured"], d["flash"], d["unit"], d["min_order"], pid))
        db.commit()
        if image != p["image"]:
            delete_product_image(p["image"])   # مسح الصورة القديمة
        flash("تم تحديث المنتج ✅ ظهر فوراً على الموقع 🚀", "success")
        return redirect(url_for("admin_products"))
    return render_template("admin/edit_product.html", p=p)


@app.route("/admin/product/delete/<int:pid>")
@admin_required
def delete_product(pid):
    db = get_db()
    p = db.execute("SELECT image FROM products WHERE id=?", (pid,)).fetchone()
    try:
        db.execute("DELETE FROM products WHERE id=?", (pid,))
        db.commit()
    except sqlite3.IntegrityError:
        db.rollback()
        flash("المنتج عليه طلبات سابقة ومينفعش يتحذف — صفّر المخزون أو أوقف القسم بدل الحذف ⚠️", "warning")
        return redirect(url_for("admin_products"))
    if p and p["image"]:
        delete_product_image(p["image"])
    flash("تم حذف المنتج 🗑️", "info")
    return redirect(url_for("admin_products"))


# ----- الطلبات -----
@app.route("/admin/orders")
@admin_required
def admin_orders():
    db = get_db()
    status_filter = request.args.get("status", "")
    dfrom, dto = valid_date(request.args.get("from")), valid_date(request.args.get("to"))
    orders = fetch_orders(db, status_filter, dfrom, dto)
    details = order_items_map(db, [o["id"] for o in orders])
    return render_template("admin/orders.html", orders=orders, details=details,
                           pay_methods=PAY_LABELS, statuses=ORDER_STATUSES,
                           status_filter=status_filter, dfrom=dfrom, dto=dto)


@app.route("/admin/order_status/<int:oid>", methods=["POST"])
@admin_required
def order_status(oid):
    db = get_db()
    st = request.form["status"]
    db.execute("UPDATE orders SET status=? WHERE id=?", (st, oid))
    # لو الطلب اتلغي → إلغاء عمولة المسوق المعلقة
    if st == "ملغي":
        db.execute("UPDATE affiliate_orders SET status='cancelled' WHERE order_id=? AND status='pending'", (oid,))
    db.commit()
    flash("تم تحديث حالة الطلب ✅ (ظهرت للعميل فوراً) 🚀", "success")
    return redirect(safe_back(url_for("admin_orders")))


# ----- العملاء: تحكم كامل -----
@app.route("/admin/users", methods=["GET", "POST"])
@admin_required
def admin_users():
    db = get_db()
    if request.method == "POST":
        f = request.form
        action = f.get("action")
        back = safe_next(f.get("next")) or url_for("admin_users")
        if action == "balance_zero_all":      # تصفير أرصدة كل العملاء (بتأكيد كتابي + نسخة احتياطية تلقائية)
            if (f.get("confirm_word") or "").strip() != "تصفير":
                flash("اكتب كلمة «تصفير» للتأكيد ❌", "danger")
                return redirect(back)
            make_backup("predelete")
            rows = db.execute("SELECT id, balance FROM users WHERE is_admin=0 AND balance!=0").fetchall()
            for r in rows:
                adjust_balance(db, r["id"], -r["balance"], "admin", "تصفير جماعي للأرصدة من الإدارة")
            db.commit()
            flash("تم تصفير أرصدة %d عميل ✅ (واتعملت نسخة احتياطية قبلها)" % len(rows), "success")
            return redirect(back)
        uid = to_int(f.get("user_id"), 0)
        u = db.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
        if not u:
            return redirect(url_for("admin_users"))
        note = (f.get("note") or "").strip()[:100]
        label = "تعديل يدوي من الإدارة" + ((" — " + note) if note else "")
        if action == "balance":
            amount = to_float(f.get("amount"))
            if amount is None or amount <= 0 or amount > 10000000:
                flash("اكتب مبلغ صحيح ❌", "danger")
                return redirect(back)
            if f.get("dir") == "deduct":
                amount = -abs(amount)
            if amount < 0 and u["balance"] + amount < 0:
                flash("مينفعش تخصم أكتر من رصيده (%d ج.م) ❌" % u["balance"], "danger")
                return redirect(back)
            adjust_balance(db, uid, amount, "admin", label)
            flash("تم تحديث الرصيد بمقدار %+.2f ج.م ✅" % amount, "success")
        elif action == "balance_set":
            target = to_float(f.get("amount"))
            if target is None or target < 0 or target > 10000000:
                flash("اكتب رصيد صحيح (صفر أو أكتر) ❌", "danger")
                return redirect(back)
            diff = round(target - (u["balance"] or 0), 2)
            if diff == 0:
                flash("الرصيد هو نفسه — مفيش تغيير", "info")
            else:
                adjust_balance(db, uid, diff, "admin", "ضبط الرصيد إلى %s ج.م%s" % (fmt_money(target), (" — " + note) if note else ""))
                flash("تم ضبط الرصيد على %s ج.م ✅" % fmt_money(target), "success")
        elif action == "balance_zero":
            if (u["balance"] or 0) == 0:
                flash("الرصيد صفر بالفعل", "info")
            else:
                adjust_balance(db, uid, -u["balance"], "admin", "تصفير الرصيد من الإدارة" + ((" — " + note) if note else ""))
                flash("تم حذف رصيد العميل (تصفير) ✅", "success")
        elif action == "edit":
            name = (f.get("name") or "").strip()
            email = (f.get("email") or "").lower().strip()
            if not name or len(name) > 80:
                flash("اكتب الاسم (لحد 80 حرف) ✍️", "danger")
                return redirect(back)
            if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
                flash("الإيميل مش صحيح 📧", "danger")
                return redirect(back)
            if f.get("reset_password") and len(f["reset_password"]) < 8:
                flash("كلمة السر الجديدة لازم تكون 8 حروف على الأقل 🔒", "danger")
                return redirect(back)
            if db.execute("SELECT 1 FROM users WHERE email=? AND id!=?", (email, uid)).fetchone():
                flash("الإيميل مستخدم لحساب تاني ❌", "danger")
                return redirect(back)
            db.execute("UPDATE users SET name=?, email=?, phone=?, address=? WHERE id=?",
                       (name, email, (f.get("phone") or "")[:30], (f.get("address") or "")[:300], uid))
            if f.get("reset_password"):
                db.execute("UPDATE users SET password=? WHERE id=?",
                           (generate_password_hash(f["reset_password"]), uid))
            flash("تم تحديث بيانات العميل ✅", "success")
        elif action == "toggle":
            if u["is_admin"]:
                flash("مينفعش تقفل حساب إدارة 🚫", "danger")
            else:
                db.execute("UPDATE users SET active=1-active WHERE id=?", (uid,))
                flash("تم %s الحساب ✅" % ("تفعيل" if not u["active"] else "إيقاف"), "success")
        elif action == "delete":
            if u["is_admin"]:
                flash("مينفعش تحذف حساب إدارة 🚫", "danger")
            else:
                delete_user_data(db, uid)
                back = url_for("admin_users")
                flash("تم حذف الحساب نهائياً 🗑️ (طلباته القديمة فضلت محفوظة للمحاسبة)", "info")
        db.commit()
        return redirect(back)
    users = db.execute("""SELECT u.*, (SELECT COUNT(*) FROM orders o WHERE o.user_id=u.id) orders_count
                          FROM users u ORDER BY u.id DESC""").fetchall()
    return render_template("admin/users.html", users=users)


@app.route("/admin/users/<int:uid>")
@admin_required
def admin_user_detail(uid):
    """صفحة تفاصيل عميل: تعديل بياناته + التحكم في رصيده + سجله + حذفه"""
    db = get_db()
    u = db.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    if not u:
        flash("العميل مش موجود ❌", "danger")
        return redirect(url_for("admin_users"))
    txs = db.execute("SELECT * FROM wallet_transactions WHERE user_id=? ORDER BY id DESC LIMIT 100", (uid,)).fetchall()
    orders = db.execute("SELECT * FROM orders WHERE user_id=? ORDER BY id DESC LIMIT 50", (uid,)).fetchall()
    reqs = db.execute("SELECT * FROM wallet_requests WHERE user_id=? ORDER BY id DESC LIMIT 30", (uid,)).fetchall()
    return render_template("admin/user_detail.html", u=u, txs=txs, orders=orders, reqs=reqs,
                           tx_types=TX_TYPES, pay_methods=PAY_LABELS)


@app.route("/admin/users/export")
@admin_required
def admin_users_export():
    """تصدير بيانات العملاء CSV"""
    db = get_db()
    users = db.execute("""SELECT id,name,email,phone,address,balance,active,created_at
                          FROM users WHERE is_admin=0 ORDER BY id""").fetchall()
    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(["ID", "الاسم", "الإيميل", "الموبايل", "العنوان", "الرصيد", "الحالة", "تاريخ التسجيل"])
    for u in users:
        w.writerow([csv_safe(x) for x in (u["id"], u["name"], u["email"], u["phone"] or "", u["address"] or "",
                    u["balance"], "نشط" if u["active"] else "موقوف", u["created_at"])])
    return Response("\ufeff" + out.getvalue(), mimetype="text/csv; charset=utf-8",
                    headers={"Content-Disposition": "attachment; filename=customers.csv"})


# ----- المحفظة (إدارة) -----
@app.route("/admin/wallet")
@admin_required
def admin_wallet():
    db = get_db()
    if request.args.get("approve"):
        rid = int(request.args["approve"])
        r = db.execute("SELECT * FROM wallet_requests WHERE id=?", (rid,)).fetchone()
        if r and r["status"] == "pending":
            adjust_balance(db, r["user_id"], r["amount"], "recharge",
                           "شحن رصيد عبر %s (من رقم: %s)" % (PAY_METHODS.get(r["method"], r["method"]), r["sender_number"] or "غير مسجل"))
            db.execute("UPDATE wallet_requests SET status='approved', handled_at=? WHERE id=?",
                       (datetime.now().strftime("%Y-%m-%d %H:%M"), rid))
            db.commit()
            flash("تمت الموافقة وإضافة %.2f ج.م ✅" % r["amount"], "success")
        return redirect(url_for("admin_wallet"))
    if request.args.get("reject"):
        rid = int(request.args["reject"])
        db.execute("UPDATE wallet_requests SET status='rejected', handled_at=? WHERE id=?",
                   (datetime.now().strftime("%Y-%m-%d %H:%M"), rid))
        db.commit()
        flash("تم رفض الطلب ❌", "info")
        return redirect(url_for("admin_wallet"))
    reqs = db.execute("""SELECT wr.*, u.name uname, u.email FROM wallet_requests wr
                         JOIN users u ON u.id=wr.user_id ORDER BY wr.id DESC""").fetchall()
    txs = db.execute("""SELECT wt.*, u.name uname FROM wallet_transactions wt
                        JOIN users u ON u.id=wt.user_id ORDER BY wt.id DESC LIMIT 30""").fetchall()
    return render_template("admin/wallet.html", reqs=reqs, txs=txs,
                           pay_methods=PAY_LABELS, sender_on=get_setting("wallet_sender", "1") == "1")


# ----- الإرجاع (إدارة) -----
@app.route("/admin/returns", methods=["GET", "POST"])
@admin_required
def admin_returns():
    db = get_db()
    if request.method == "POST":
        f = request.form
        rid = int(f["return_id"])
        action = f["action"]
        r = db.execute("SELECT * FROM returns WHERE id=?", (rid,)).fetchone()
        if r and r["status"] == "pending":
            now = datetime.now().strftime("%Y-%m-%d %H:%M")
            if action == "approve":
                db.execute("UPDATE returns SET status='approved', admin_note=?, handled_at=? WHERE id=?",
                           (f.get("note", ""), now, rid))
                # رجّع المخزون
                items = db.execute("SELECT * FROM order_items WHERE order_id=?", (r["order_id"],)).fetchall()
                for it in items:
                    db.execute("UPDATE products SET stock=stock+?, sold=sold-? WHERE id=?",
                               (it["qty"], it["qty"], it["product_id"]))
                # استرداد للمحفظة لو الطلب مدفوع
                o = db.execute("SELECT * FROM orders WHERE id=?", (r["order_id"],)).fetchone()
                if o and o["paid"] > 0 and f.get("refund"):
                    adjust_balance(db, r["user_id"], o["paid"], "refund", "استرداد قيمة الإرجاع للطلب #%d" % o["id"])
                # إلغاء عمولة المسوق
                db.execute("UPDATE affiliate_orders SET status='cancelled' WHERE order_id=? AND status='pending'", (r["order_id"],))
                flash("تمت الموافقة على الإرجاع ✅ (المخزون رجع والعمولة اتلغت)", "success")
            else:
                db.execute("UPDATE returns SET status='rejected', admin_note=?, handled_at=? WHERE id=?",
                           (f.get("note", ""), now, rid))
                flash("تم رفض طلب الإرجاع ❌", "info")
            db.commit()
        return redirect(url_for("admin_returns"))
    rows = db.execute("""SELECT r.*, u.name uname, u.email, o.total ototal, o.status ostatus
                         FROM returns r JOIN users u ON u.id=r.user_id
                         JOIN orders o ON o.id=r.order_id ORDER BY r.id DESC""").fetchall()
    return render_template("admin/returns.html", rows=rows)


# ----- التسويق بالعمولة (إدارة) -----
@app.route("/admin/affiliates")
@admin_required
def admin_affiliates():
    db = get_db()
    # إجراءات
    if request.args.get("toggle"):
        mid = int(request.args["toggle"])
        db.execute("UPDATE marketers SET active=1-active WHERE id=?", (mid,))
        db.commit()
        return redirect(url_for("admin_affiliates"))
    if request.args.get("confirm"):  # تأكيد عمولة يدوي
        aid = int(request.args["confirm"])
        a = db.execute("SELECT * FROM affiliate_orders WHERE id=? AND status='pending'", (aid,)).fetchone()
        if a:
            mk = db.execute("SELECT user_id FROM marketers WHERE id=?", (a["marketer_id"],)).fetchone()
            adjust_balance(db, mk["user_id"], a["commission"], "affiliate", "عمولة طلب #%d (تأكيد يدوي)" % a["order_id"])
            db.execute("UPDATE affiliate_orders SET status='confirmed' WHERE id=?", (aid,))
            db.commit()
            flash("تم تأكيد العمولة وتحويلها للرصيد ✅", "success")
        return redirect(url_for("admin_affiliates"))
    if request.args.get("cancel"):
        aid = int(request.args["cancel"])
        db.execute("UPDATE affiliate_orders SET status='cancelled' WHERE id=? AND status='pending'", (aid,))
        db.commit()
        flash("تم إلغاء العمولة ❌", "info")
        return redirect(url_for("admin_affiliates"))
    process_matured_affiliates(db)
    db.commit()
    marketers = db.execute("""SELECT m.*, u.name, u.email,
        (SELECT COUNT(*) FROM affiliate_orders a WHERE a.marketer_id=m.id) orders_count,
        (SELECT COALESCE(SUM(commission),0) FROM affiliate_orders a WHERE a.marketer_id=m.id AND a.status='confirmed') paid,
        (SELECT COALESCE(SUM(commission),0) FROM affiliate_orders a WHERE a.marketer_id=m.id AND a.status='pending') pending
        FROM marketers m JOIN users u ON u.id=m.user_id ORDER BY m.id DESC""").fetchall()
    mid = request.args.get("mid", type=int)
    aff_orders = []
    if mid:
        aff_orders = db.execute("""SELECT a.*, m.code, u.name mname FROM affiliate_orders a
                                   JOIN marketers m ON m.id=a.marketer_id
                                   JOIN users u ON u.id=m.user_id
                                   WHERE a.marketer_id=? ORDER BY a.id DESC""", (mid,)).fetchall()
    flagged = db.execute("""SELECT a.*, m.code FROM affiliate_orders a JOIN marketers m ON m.id=a.marketer_id
                            WHERE a.flag=1 ORDER BY a.id DESC LIMIT 20""").fetchall()
    return render_template("admin/affiliates.html", marketers=marketers,
                           aff_orders=aff_orders, flagged=flagged, mid=mid,
                           disc=AFF_BUYER_DISCOUNT, comm=AFF_COMMISSION, hold=AFF_HOLD_DAYS)


# ----- الأقسام (إضافة / تعديل / إيقاف / حذف) -----
@app.route("/admin/categories", methods=["GET", "POST"])
@admin_required
def admin_categories():
    db = get_db()
    if request.method == "POST":
        f = request.form
        action = f.get("action")
        cid = to_int(f.get("cat_id"), 0)
        name = (f.get("name") or "").strip()
        emoji = (f.get("emoji") or "").strip()[:4]
        show_home = 1 if f.get("show_home") else 0
        cat = db.execute("SELECT * FROM categories WHERE id=?", (cid,)).fetchone() if cid else None
        if action == "add":
            if not (2 <= len(name) <= 40):
                flash("اسم القسم من 2 لـ 40 حرف ❌", "danger")
            elif db.execute("SELECT 1 FROM categories WHERE name=?", (name,)).fetchone():
                flash("القسم ده موجود قبل كده ❌", "danger")
            else:
                nxt = db.execute("SELECT COALESCE(MAX(sort_order),0)+1 n FROM categories").fetchone()["n"]
                db.execute("INSERT INTO categories(name,emoji,active,show_home,sort_order) VALUES(?,?,1,?,?)",
                           (name, emoji, show_home, nxt))
                flash("تم إضافة القسم ✅ وظهر فوراً على الموقع 🚀", "success")
        elif not cat:
            flash("القسم مش موجود ❌", "danger")
        elif action == "edit":
            if not (2 <= len(name) <= 40):
                flash("اسم القسم من 2 لـ 40 حرف ❌", "danger")
            elif db.execute("SELECT 1 FROM categories WHERE name=? AND id!=?", (name, cid)).fetchone():
                flash("فيه قسم تاني بنفس الاسم ❌", "danger")
            else:
                if name != cat["name"]:   # تغيير الاسم يتنقل لكل منتجات القسم
                    db.execute("UPDATE products SET category=? WHERE category=?", (name, cat["name"]))
                db.execute("UPDATE categories SET name=?, emoji=?, show_home=?, sort_order=? WHERE id=?",
                           (name, emoji, show_home, to_int(f.get("sort_order"), cat["sort_order"], lo=0, hi=9999), cid))
                flash("تم حفظ القسم ✅", "success")
        elif action == "toggle":
            db.execute("UPDATE categories SET active=1-active WHERE id=?", (cid,))
            flash("تم %s القسم «%s» ✅" % ("تفعيل" if not cat["active"] else "إيقاف", cat["name"]), "success")
        elif action == "delete":
            used = db.execute("SELECT COUNT(*) c FROM products WHERE category=?", (cat["name"],)).fetchone()["c"]
            if used:
                flash("القسم فيه %d منتج — انقلهم لقسم تاني أو اكتفي بإيقاف القسم ⚠️" % used, "warning")
            else:
                db.execute("DELETE FROM categories WHERE id=?", (cid,))
                flash("تم حذف القسم 🗑️", "info")
        db.commit()
        return redirect(url_for("admin_categories"))
    cats = db.execute("""SELECT c.*, (SELECT COUNT(*) FROM products p WHERE p.category=c.name) pcount
                         FROM categories c ORDER BY c.sort_order, c.id""").fetchall()
    return render_template("admin/categories.html", cats=cats)


# ----- الكوبونات -----
@app.route("/admin/coupons", methods=["GET", "POST"])
@admin_required
def admin_coupons():
    db = get_db()
    if request.method == "POST":
        f = request.form
        pct = to_int(f.get("percent"), 0)
        if not (1 <= pct <= 100) or not f.get("code", "").strip():
            flash("اكتب كود ونسبة خصم من 1 لـ 100 ❌", "danger")
            return redirect(url_for("admin_coupons"))
        try:
            db.execute("INSERT INTO coupons(code,percent,expires) VALUES(?,?,?)",
                       (f["code"].strip().upper(), pct, f.get("expires") or None))
            db.commit()
            flash("تم إنشاء الكوبون 🎟️", "success")
        except sqlite3.IntegrityError:
            flash("الكود ده موجود قبل كده ❌", "danger")
        return redirect(url_for("admin_coupons"))
    if request.args.get("toggle"):
        db.execute("UPDATE coupons SET active = 1 - active WHERE code=?", (request.args["toggle"],))
        db.commit()
        return redirect(url_for("admin_coupons"))
    coupons = db.execute("SELECT * FROM coupons").fetchall()
    return render_template("admin/coupons.html", coupons=coupons)


# ================== v6: أدوات عامة (شحن / تصدير / واتساب / SEO) ==================
def site_info():
    """بيانات التواصل (من الإعدادات) — بتظهر في الفوتر وصفحة الدعم والفاتورة"""
    keys = ("support_phone", "support_email", "support_hours", "store_address", "whatsapp_number",
            "support_whatsapp", "support_whatsapp_on")
    rows = {r["key"]: r["value"] for r in get_db().execute(
        "SELECT key,value FROM settings WHERE key IN (%s)" % ",".join("?" * len(keys)), keys).fetchall()}
    return {"phone": rows.get("support_phone") or SUPPORT_PHONE,
            "email": rows.get("support_email") or SUPPORT_EMAIL,
            "hours": rows.get("support_hours") or "",
            "address": rows.get("store_address") or "",
            "whatsapp": rows.get("whatsapp_number") or WHATSAPP_NUMBER,
            "wa_support": rows.get("support_whatsapp") or rows.get("support_phone") or SUPPORT_PHONE,
            "wa_on": rows.get("support_whatsapp_on", "1") == "1",
            "wa_link": "https://wa.me/%s?text=%s" % (
                intl_number(rows.get("support_whatsapp") or rows.get("support_phone") or SUPPORT_PHONE),
                quote("أهلاً Harly Click 👋 محتاج مساعدة"))}


def site_base_url():
    return SITE_URL or request.url_root.rstrip("/")


def canonical_url():
    base = site_base_url() + request.path
    if request.endpoint == "shop":
        parts = []
        if request.args.get("cat"):
            parts.append("cat=" + quote(request.args["cat"]))
        if to_int(request.args.get("page"), 1) > 1:
            parts.append("page=%d" % to_int(request.args.get("page"), 1))
        if parts:
            base += "?" + "&".join(parts)
    return base


def active_zones():
    return get_db().execute("SELECT * FROM shipping_zones WHERE active=1 ORDER BY sort_order, id").fetchall()


def shipping_quote(amount, gov=None):
    """يحسب الشحن: حسب المحافظة، ومجاني لو المبلغ وصل للحد اللي الإدارة محدده"""
    enabled = get_setting("shipping_enabled", "1") == "1"
    free_min = to_float(get_setting("free_shipping_min", "0"), 0) or 0
    q = {"enabled": enabled, "gov": gov or "", "fee": 0.0, "known": False, "free": False,
         "free_min": free_min, "remaining": 0.0}
    if not enabled:
        q["known"] = True
        return q
    free = free_min > 0 and amount >= free_min
    q["free"] = free
    if free_min > 0 and not free:
        q["remaining"] = round(free_min - amount, 2)
    if gov:
        z = get_db().execute("SELECT fee FROM shipping_zones WHERE governorate=? AND active=1", (gov,)).fetchone()
        if z:
            q["known"] = True
            q["fee"] = 0.0 if free else float(z["fee"] or 0)
    elif free:
        q["known"] = True
    return q


def recharge_methods(only_on=False):
    """وسائل شحن الرصيد: التفعيل + رقم التحويل لكل وسيلة (لو فاضي بيستخدم الرقم العام)"""
    default_num = (get_setting("transfer_number", TRANSFER_NUMBER) or TRANSFER_NUMBER).strip()
    out = []
    for k, label in PAY_METHODS.items():
        on = get_setting("pm_%s_on" % k, "1") == "1"
        num = (get_setting("pm_%s_number" % k, "") or "").strip()
        if only_on and not on:
            continue
        out.append({"key": k, "label": label, "on": on, "number": num or default_num, "own_number": num})
    return out


def transfer_number_for(method):
    for m in recharge_methods():
        if m["key"] == method:
            return m["number"]
    return TRANSFER_NUMBER


def checkout_methods():
    return {k: get_setting("co_%s_on" % k, "1") == "1" for k, _ in CHECKOUT_METHODS}


def valid_date(v):
    v = (v or "").strip()
    try:
        return datetime.strptime(v, "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        return ""


def fetch_orders(db, status="", dfrom="", dto=""):
    sql = """SELECT o.*, u.name uname, u.email FROM orders o
             LEFT JOIN users u ON u.id=o.user_id WHERE 1=1"""
    params = []
    if status:
        sql += " AND o.status=?"
        params.append(status)
    if dfrom:
        sql += " AND date(o.created_at)>=?"
        params.append(dfrom)
    if dto:
        sql += " AND date(o.created_at)<=?"
        params.append(dto)
    return db.execute(sql + " ORDER BY o.id DESC", params).fetchall()


def order_items_map(db, ids):
    out = {i: [] for i in ids}
    for k in range(0, len(ids), 500):
        chunk = ids[k:k + 500]
        rows = db.execute("""SELECT oi.*, p.name, p.unit FROM order_items oi
                             JOIN products p ON p.id=oi.product_id
                             WHERE oi.order_id IN (%s)""" % ",".join("?" * len(chunk)), chunk).fetchall()
        for r in rows:
            out[r["order_id"]].append(r)
    return out


def delete_user_data(db, uid):
    """حذف حساب عميل بأمان: الطلبات القديمة بتفضل محفوظة للمحاسبة (من غير ربط بالحساب)"""
    for t in ("wishlist", "wallet_requests", "wallet_transactions", "returns"):
        db.execute("DELETE FROM %s WHERE user_id=?" % t, (uid,))
    m = db.execute("SELECT id FROM marketers WHERE user_id=?", (uid,)).fetchone()
    if m:
        db.execute("DELETE FROM affiliate_orders WHERE marketer_id=?", (m["id"],))
        db.execute("DELETE FROM marketers WHERE id=?", (m["id"],))
    for t in ("orders", "whatsapp_orders", "support_tickets"):
        db.execute("UPDATE %s SET user_id=NULL WHERE user_id=?" % t, (uid,))
    db.execute("DELETE FROM users WHERE id=?", (uid,))


def csv_safe(v):
    """يمنع حقن المعادلات في Excel (خلية تبدأ بـ = + - @)"""
    if isinstance(v, str) and v[:1] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + v
    return v


def export_response(sheets, base_name, fmt):
    """sheets = [(عنوان, [رؤوس], [صفوف])] → ملف Excel (أو CSV للشيت الأول)"""
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    if fmt == "xlsx" and Workbook is not None:
        wb = Workbook()
        wb.remove(wb.active)
        for title, headers, rows in sheets:
            ws = wb.create_sheet(re.sub(r"[\[\]:*?/\\]", " ", title)[:31])
            ws.sheet_view.rightToLeft = True
            ws.append(headers)
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="6C2BD9")
                c.alignment = Alignment(horizontal="center")
            widths = [len(str(h)) + 2 for h in headers]
            for r in rows:
                r = [csv_safe(x) for x in r]
                ws.append(r)
                for i, x in enumerate(r):
                    widths[i] = max(widths[i], min(len(str(x)), 50) + 2)
            for i, w in enumerate(widths, 1):
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
            ws.freeze_panes = "A2"
        bio = io.BytesIO()
        wb.save(bio)
        return Response(bio.getvalue(),
                        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        headers={"Content-Disposition": "attachment; filename=%s_%s.xlsx" % (base_name, stamp)})
    title, headers, rows = sheets[0]
    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(headers)
    for r in rows:
        w.writerow([csv_safe(x) for x in r])
    return Response("\ufeff" + out.getvalue(), mimetype="text/csv; charset=utf-8",
                    headers={"Content-Disposition": "attachment; filename=%s_%s.csv" % (base_name, stamp)})


def cod_amount(o):
    """المطلوب تحصيله من العميل عند الاستلام"""
    if o["payment_method"] == "balance":
        return 0.0
    return round(max((o["total"] or 0) - (o["paid"] or 0), 0), 2)


# ---------- باركود Code39 (SVG بدون مكتبات) لبوليصة الشحن ----------
CODE39 = {
    "0": "nnnwwnwnn", "1": "wnnwnnnnw", "2": "nnwwnnnnw", "3": "wnwwnnnnn", "4": "nnnwwnnnw",
    "5": "wnnwwnnnn", "6": "nnwwwnnnn", "7": "nnnwnnwnw", "8": "wnnwnnwnn", "9": "nnwwnnwnn",
    "A": "wnnnnwnnw", "B": "nnwnnwnnw", "C": "wnwnnwnnn", "D": "nnnnwwnnw", "E": "wnnnwwnnn",
    "F": "nnwnwwnnn", "G": "nnnnnwwnw", "H": "wnnnnwwnn", "I": "nnwnnwwnn", "J": "nnnnwwwnn",
    "K": "wnnnnnnww", "L": "nnwnnnnww", "M": "wnwnnnnwn", "N": "nnnnwnnww", "O": "wnnnwnnwn",
    "P": "nnwnwnnwn", "Q": "nnnnnnwww", "R": "wnnnnnwwn", "S": "nnwnnnwwn", "T": "nnnnwnwwn",
    "U": "wwnnnnnnw", "V": "nwwnnnnnw", "W": "wwwnnnnnn", "X": "nwnnwnnnw", "Y": "wwnnwnnnn",
    "Z": "nwwnwnnnn", "-": "nwnnnnwnw", ".": "wwnnnnwnn", " ": "nwwnnnwnn", "*": "nwnnwnwnn"}


def code39_svg(text, narrow=1.6, wide=4.0, height=46):
    text = re.sub(r"[^0-9A-Z\-\. ]", "", str(text).upper())
    x, rects = 0.0, []
    for ch in "*" + text + "*":
        for i, e in enumerate(CODE39[ch]):
            w = wide if e == "w" else narrow
            if i % 2 == 0:
                rects.append('<rect x="%.2f" y="0" width="%.2f" height="%d"/>' % (x, w, height))
            x += w
        x += narrow
    return Markup('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 %.2f %d" width="%d" height="%d" '
                  'fill="#000" role="img" aria-label="barcode %s">%s</svg>'
                  % (x, height, round(x), height, text, "".join(rects)))


# ---------- واتساب: تسجيل الطلب + بناء الرسالة ----------
def save_wa_order(db, lines, subtotal, discount, ship, source, cust):
    """يسجّل طلب واتساب في لوحة الإدارة (بيمنع التكرار لو العميل داس الزرار كذا مرة)"""
    ip = client_ip()
    if db.execute("SELECT COUNT(*) c FROM whatsapp_orders WHERE ip=? AND created_at>=datetime('now','-1 hour')",
                  (ip,)).fetchone()["c"] >= 20:
        return None
    items_json = json.dumps(lines, ensure_ascii=False)
    fp = hashlib.sha1(("%s|%s|%s|%s|%s" % (session.get("user_id") or ip, items_json, ship["gov"],
                                          cust.get("phone", ""), source)).encode("utf-8")).hexdigest()
    old = db.execute("SELECT id FROM whatsapp_orders WHERE fingerprint=? AND created_at>=datetime('now','-10 minutes')",
                     (fp,)).fetchone()
    if old:
        return old["id"]
    total = round(subtotal - discount + ship["fee"], 2)
    cur = db.execute("""INSERT INTO whatsapp_orders(user_id,customer_name,phone,address,governorate,items_json,
                        subtotal,discount,shipping,total,source,fingerprint,ip)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                     (session.get("user_id"), cust.get("name", ""), cust.get("phone", ""), cust.get("address", ""),
                      ship["gov"], items_json, subtotal, discount, ship["fee"], total, source, fp, ip))
    db.commit()
    return cur.lastrowid


def wa_redirect(lines, subtotal, discount, source, product_url=""):
    db = get_db()
    subtotal, discount = round(subtotal, 2), round(discount or 0, 2)
    gov = (request.form.get("governorate") or request.args.get("gov") or session.get("gov") or "").strip()
    if gov and not db.execute("SELECT 1 FROM shipping_zones WHERE governorate=? AND active=1", (gov,)).fetchone():
        gov = ""
    if gov:
        session["gov"] = gov
        session.modified = True
    ship = shipping_quote(round(subtotal - discount, 2), gov)
    phone = (request.form.get("phone") or "").strip()
    if not re.match(r"^[0-9+\-\s]{7,20}$", phone):
        phone = ""
    cust = {"name": session.get("user_name") or "", "phone": phone,
            "address": (request.form.get("address") or "").strip()[:300]}
    oid = save_wa_order(db, lines, subtotal, discount, ship, source, cust)
    number = intl_number(get_setting("whatsapp_number", "") or WHATSAPP_NUMBER)
    head = ["أهلاً Harly Click 👋"]
    if oid:
        head.append("🧾 مرجع الطلب: WA-%d" % oid)
    head += ["عايز أطلب المنتجات دي:", ""]
    if cust["name"]:
        head.insert(len(head) - 1, "الاسم: %s" % cust["name"])
    tail = ["", "المجموع الفرعي: %s ج.م" % fmt_money(subtotal)]
    if discount:
        tail.append("الخصم: -%s ج.م" % fmt_money(discount))
    if ship["enabled"]:
        if ship["gov"]:
            tail.append("المحافظة: %s" % ship["gov"])
            tail.append("الشحن: %s" % ("مجاني 🎉" if ship["fee"] == 0 else "%s ج.م" % fmt_money(ship["fee"])))
        elif ship["free"]:
            tail.append("الشحن: مجاني 🎉")
        else:
            tail.append("المحافظة: (هحددها معاكم) — الشحن بيتحدد حسب المحافظة")
    tail.append("الإجمالي: %s ج.م" % fmt_money(round(subtotal - discount + ship["fee"], 2)))
    if cust["phone"]:
        tail.append("موبايلي: %s" % cust["phone"])
    if cust["address"]:
        tail.append("العنوان: %s" % cust["address"])
    if product_url:
        tail.append("رابط المنتج: %s" % product_url)
    out, used = [], len("\n".join(head + tail))
    for i, ln in enumerate(lines, 1):
        unit = "" if ln["unit"] == DEFAULT_UNIT else " " + ln["unit"]
        line = "%d) %s × %d%s = %s ج.م" % (i, ln["name"], ln["qty"], unit, fmt_money(ln["price"] * ln["qty"]))
        if used + len(line) > 1200:     # حد طول الرابط
            out.append("... و %d منتج تانيين (هبعتهم في رسالة تانية)" % (len(lines) - i + 1))
            break
        out.append(line)
        used += len(line) + 1
    return redirect("https://wa.me/%s?text=%s" % (number, quote("\n".join(head + out + tail))))


# ---------- SEO ----------
def product_ld(p):
    base = site_base_url()
    d = {"@context": "https://schema.org", "@type": "Product", "name": p["name"],
         "description": ((p["description"] or p["name"]).strip())[:300], "sku": str(p["id"]),
         "category": p["category"], "brand": {"@type": "Brand", "name": "Harly Click"},
         "offers": {"@type": "Offer", "url": base + url_for("product", pid=p["id"]), "priceCurrency": "EGP",
                    "price": "%.2f" % p["price"], "itemCondition": "https://schema.org/NewCondition",
                    "availability": "https://schema.org/InStock" if p["stock"] > 0 else "https://schema.org/OutOfStock"}}
    if p["image"]:
        d["image"] = [base + img_src(p["image"])]
    return d


def site_ld():
    base = site_base_url()
    info = site_info()
    return [{"@context": "https://schema.org", "@type": "WebSite", "name": "Harly Click", "url": base + "/",
             "potentialAction": {"@type": "SearchAction", "target": base + "/shop?q={search_term_string}",
                                 "query-input": "required name=search_term_string"}},
            {"@context": "https://schema.org", "@type": "Organization", "name": "Harly Click", "url": base + "/",
             "email": info["email"],
             "contactPoint": {"@type": "ContactPoint", "telephone": "+" + intl_number(info["phone"]),
                              "contactType": "customer service", "availableLanguage": "Arabic"}}]


NOINDEX_PREFIXES = ("/admin", "/cart", "/checkout", "/login", "/register", "/profile", "/wallet", "/wishlist",
                    "/my_orders", "/returns", "/affiliate", "/order_done", "/invoice", "/autocomplete", "/wish/",
                    "/add/", "/remove", "/update_cart", "/apply_coupon", "/remove_coupon", "/logout")


def maybe_gzip(resp):
    """ضغط صفحات HTML/JSON تلقائياً (أسرع تحميل)"""
    try:
        if (resp.status_code != 200 or resp.direct_passthrough or "Content-Encoding" in resp.headers
                or "Content-Disposition" in resp.headers
                or "gzip" not in request.headers.get("Accept-Encoding", "")
                or not (resp.mimetype or "").startswith(("text/html", "application/json", "application/xml", "text/xml"))):
            return resp
        data = resp.get_data()
        if len(data) < 600:
            return resp
        resp.set_data(gzip.compress(data, compresslevel=5))
        resp.headers["Content-Encoding"] = "gzip"
        resp.headers["Content-Length"] = str(len(resp.get_data()))
        resp.headers.add("Vary", "Accept-Encoding")
    except Exception:
        pass
    return resp


@app.route("/robots.txt")
def robots_txt():
    base = site_base_url()
    lines = ["User-agent: *"] + ["Disallow: " + p for p in
             ("/admin", "/cart", "/checkout", "/login", "/register", "/profile", "/wallet", "/wishlist",
              "/my_orders", "/returns", "/affiliate", "/order_done", "/invoice", "/autocomplete")] + \
            ["Allow: /", "", "Sitemap: %s/sitemap.xml" % base]
    return Response("\n".join(lines) + "\n", mimetype="text/plain")


@app.route("/sitemap.xml")
def sitemap_xml():
    db = get_db()
    base = site_base_url()
    urls = [base + "/", base + url_for("shop"), base + url_for("about"), base + url_for("contact")]
    for c in category_rows(active_only=True):
        urls.append(base + url_for("shop", cat=c["name"]))
    for r in db.execute("SELECT id FROM products WHERE " + VISIBLE_SQL + " ORDER BY id DESC LIMIT 5000").fetchall():
        urls.append(base + url_for("product", pid=r["id"]))
    body = "".join("<url><loc>%s</loc></url>" % xml_escape(u) for u in urls)
    return Response('<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">'
                    + body + "</urlset>", mimetype="application/xml")


# ================== v6: النسخ الاحتياطي ==================
_backup_started = False


def _backup_files():
    out = []
    if os.path.isdir(BACKUP_DIR):
        for fn in os.listdir(BACKUP_DIR):
            if BACKUP_NAME_RE.match(fn):
                st = os.stat(os.path.join(BACKUP_DIR, fn))
                out.append({"name": fn, "size": st.st_size, "tag": fn.rsplit("_", 1)[1].split(".")[0],
                            "time": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")})
    return sorted(out, key=lambda x: x["name"], reverse=True)


def make_backup(tag="manual"):
    """نسخة احتياطية متسقة من القاعدة (SQLite backup API) مضغوطة gzip — بترجّع اسم الملف"""
    if not os.path.exists(DB):
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    tag = re.sub(r"[^a-z]", "", tag.lower()) or "manual"
    name = "harly_click_%s_%s.db.gz" % (datetime.now().strftime("%Y-%m-%d_%H%M%S"), tag)
    tmp = os.path.join(BACKUP_DIR, name + ".tmp.db")
    try:
        src, dst = sqlite3.connect(DB), sqlite3.connect(tmp)
        try:
            src.backup(dst)
        finally:
            dst.close()
            src.close()
        with open(tmp, "rb") as fi, gzip.open(os.path.join(BACKUP_DIR, name), "wb", compresslevel=6) as fo:
            shutil.copyfileobj(fi, fo)
        return name
    except Exception as e:   # pragma: no cover
        print("backup failed:", e)
        return None
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


def prune_backups(keep_auto=14, keep_other=30):
    groups = {}
    for f in _backup_files():
        groups.setdefault("auto" if f["tag"] == "auto" else "other", []).append(f)
    for g, files in groups.items():
        for f in files[(keep_auto if g == "auto" else keep_other):]:
            try:
                os.remove(os.path.join(BACKUP_DIR, f["name"]))
            except OSError:
                pass


def ensure_daily_backup():
    """نسخة تلقائية مرة في اليوم (قفل بملف عشان أكتر من worker ميعملوش نسخة مكررة)"""
    today = date.today().isoformat()
    os.makedirs(BACKUP_DIR, exist_ok=True)
    if any(f["name"].startswith("harly_click_%s_" % today) and f["tag"] == "auto" for f in _backup_files()):
        return
    lock = os.path.join(BACKUP_DIR, ".lock-" + today)
    try:
        os.close(os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY))
    except FileExistsError:
        return
    name = make_backup("auto")
    if not name:
        try:
            os.remove(lock)
        except OSError:
            pass
        return
    keep = 14
    try:
        c = sqlite3.connect(DB)
        r = c.execute("SELECT value FROM settings WHERE key='backup_keep'").fetchone()
        c.close()
        keep = to_int(r[0], 14, lo=1, hi=90) if r else 14
    except Exception:
        pass
    prune_backups(keep_auto=keep)
    for fn in os.listdir(BACKUP_DIR):     # تنضيف أقفال الأيام القديمة
        if fn.startswith(".lock-") and fn != ".lock-" + today:
            try:
                os.remove(os.path.join(BACKUP_DIR, fn))
            except OSError:
                pass


def _backup_loop():
    time.sleep(8)
    while True:
        try:
            ensure_daily_backup()
        except Exception as e:   # pragma: no cover
            print("daily backup error:", e)
        time.sleep(3600)


def start_backup_scheduler():
    global _backup_started
    if _backup_started or os.environ.get("DISABLE_BACKUP_THREAD") == "1":
        return
    _backup_started = True
    threading.Thread(target=_backup_loop, daemon=True, name="daily-backup").start()


@app.route("/admin/backups", methods=["GET", "POST"])
@admin_required
def admin_backups():
    if request.method == "POST":
        action = request.form.get("action")
        if action == "now":
            name = make_backup("manual")
            flash("تم عمل نسخة احتياطية ✅ %s" % name if name else "فشل عمل النسخة ❌", "success" if name else "danger")
        elif action == "delete":
            name = request.form.get("name", "")
            if BACKUP_NAME_RE.match(name) and os.path.exists(os.path.join(BACKUP_DIR, name)):
                os.remove(os.path.join(BACKUP_DIR, name))
                flash("تم حذف النسخة 🗑️", "info")
        return redirect(url_for("admin_backups"))
    return render_template("admin/backups.html", files=_backup_files(),
                           keep=get_setting("backup_keep", "14"), backup_dir=BACKUP_DIR)


@app.route("/admin/backups/download/<name>")
@admin_required
def admin_backup_download(name):
    if not BACKUP_NAME_RE.match(name) or not os.path.exists(os.path.join(BACKUP_DIR, name)):
        flash("النسخة مش موجودة ❌", "danger")
        return redirect(url_for("admin_backups"))
    return send_from_directory(BACKUP_DIR, name, as_attachment=True)


# ================== v6: تصدير الطلبات + بوليصة الشحن ==================
def orders_export_rows(db, orders, details):
    rows = []
    for o in orders:
        items = "؛ ".join("%s × %d%s" % (it["name"], it["qty"],
                          "" if (it["unit"] or DEFAULT_UNIT) == DEFAULT_UNIT else " " + it["unit"])
                          for it in details.get(o["id"], []))
        qty = sum(it["qty"] for it in details.get(o["id"], []))
        ship = o["shipping"] or 0
        rows.append([o["id"], o["created_at"], o["status"], o["uname"] or "حساب محذوف", o["email"] or "",
                     o["phone"] or "", o["governorate"] or "", o["address"] or "", items, qty,
                     round((o["total"] or 0) - ship + (o["discount"] or 0), 2), o["discount"] or 0, ship,
                     o["total"], PAY_LABELS.get(o["payment_method"], "كاش عند الاستلام"),
                     o["paid"] or 0, cod_amount(o)])
    return rows


ORDERS_HEADERS = ["رقم الطلب", "التاريخ", "الحالة", "العميل", "الإيميل", "الموبايل", "المحافظة", "العنوان",
                  "المنتجات", "عدد القطع", "المجموع الفرعي", "الخصم", "الشحن", "الإجمالي", "طريقة الدفع",
                  "المدفوع", "المطلوب تحصيله"]


@app.route("/admin/orders/export")
@admin_required
def admin_orders_export():
    db = get_db()
    orders = fetch_orders(db, request.args.get("status", ""), valid_date(request.args.get("from")),
                          valid_date(request.args.get("to")))[:20000]
    details = order_items_map(db, [o["id"] for o in orders])
    fmt = "xlsx" if request.args.get("format") == "xlsx" else "csv"
    return export_response([("الطلبات", ORDERS_HEADERS, orders_export_rows(db, orders, details))], "orders", fmt)


@app.route("/admin/orders/waybills")
@admin_required
def admin_waybills():
    """بوليصة شحن قابلة للطباعة (طلب واحد: ?ids=5 — أو كل الطلبات المفلترة)"""
    db = get_db()
    ids = [int(x) for x in re.findall(r"\d+", request.args.get("ids", ""))][:200]
    if ids:
        orders = db.execute("""SELECT o.*, u.name uname, u.email FROM orders o LEFT JOIN users u ON u.id=o.user_id
                               WHERE o.id IN (%s) ORDER BY o.id""" % ",".join("?" * len(ids)), ids).fetchall()
    else:
        orders = fetch_orders(db, request.args.get("status", ""), valid_date(request.args.get("from")),
                              valid_date(request.args.get("to")))[:200]
        orders = [o for o in orders if o["status"] != "ملغي"]
    if not orders:
        flash("مفيش طلبات للطباعة 🚫", "warning")
        return redirect(url_for("admin_orders"))
    details = order_items_map(db, [o["id"] for o in orders])
    bills = [{"o": o, "items": details.get(o["id"], []), "cod": cod_amount(o),
              "barcode": code39_svg("HC%05d" % o["id"]), "code": "HC%05d" % o["id"]} for o in orders]
    return render_template("admin/waybill.html", bills=bills, store=site_info(),
                           now=datetime.now().strftime("%Y-%m-%d %H:%M"))


# ================== v6: طلبات واتساب (إدارة) ==================
def _wa_rows(db, status=""):
    sql = "SELECT * FROM whatsapp_orders" + (" WHERE status=?" if status else "") + " ORDER BY id DESC LIMIT 1000"
    out = []
    for r in db.execute(sql, (status,) if status else ()).fetchall():
        d = dict(r)
        try:
            d["items"] = json.loads(r["items_json"] or "[]")
        except ValueError:
            d["items"] = []
        d["wa_link"] = ("https://wa.me/%s" % intl_number(r["phone"])) if r["phone"] else ""
        out.append(d)
    return out


@app.route("/admin/whatsapp", methods=["GET", "POST"])
@admin_required
def admin_whatsapp():
    db = get_db()
    if request.method == "POST":
        f = request.form
        wid = to_int(f.get("id"), 0)
        action = f.get("action")
        if action == "status" and f.get("status") in WA_STATUSES:
            db.execute("UPDATE whatsapp_orders SET status=? WHERE id=?", (f["status"], wid))
            flash("تم تحديث حالة الطلب ✅", "success")
        elif action == "note":
            db.execute("UPDATE whatsapp_orders SET admin_note=? WHERE id=?", ((f.get("note") or "").strip()[:500], wid))
            flash("تم حفظ الملاحظة ✅", "success")
        elif action == "delete":
            db.execute("DELETE FROM whatsapp_orders WHERE id=?", (wid,))
            flash("تم حذف الطلب 🗑️", "info")
        db.commit()
        return redirect(safe_back(url_for("admin_whatsapp")))
    status = request.args.get("status", "")
    counts = {r["status"]: r["c"] for r in db.execute("SELECT status, COUNT(*) c FROM whatsapp_orders GROUP BY status").fetchall()}
    return render_template("admin/whatsapp_orders.html", rows=_wa_rows(db, status if status in WA_STATUSES else ""),
                           statuses=WA_STATUSES, status_filter=status, counts=counts)


@app.route("/admin/whatsapp/export")
@admin_required
def admin_whatsapp_export():
    rows = []
    for r in _wa_rows(get_db(), request.args.get("status", "")):
        items = "؛ ".join("%s × %d%s" % (i["name"], i["qty"], "" if i["unit"] == DEFAULT_UNIT else " " + i["unit"]) for i in r["items"])
        rows.append(["WA-%d" % r["id"], r["created_at"], r["status"], r["customer_name"] or "زائر", r["phone"] or "",
                     r["governorate"] or "", r["address"] or "", items, r["subtotal"], r["discount"], r["shipping"],
                     r["total"], {"cart": "السلة", "product": "صفحة المنتج", "checkout": "صفحة الدفع"}.get(r["source"], r["source"]),
                     r["admin_note"] or ""])
    heads = ["المرجع", "التاريخ", "الحالة", "العميل", "الموبايل", "المحافظة", "العنوان", "المنتجات", "المجموع الفرعي",
             "الخصم", "الشحن", "الإجمالي", "المصدر", "ملاحظة الإدارة"]
    return export_response([("طلبات واتساب", heads, rows)], "whatsapp_orders",
                           "xlsx" if request.args.get("format") == "xlsx" else "csv")


# ================== v6: الشحن (إدارة) ==================
@app.route("/admin/shipping", methods=["GET", "POST"])
@admin_required
def admin_shipping():
    db = get_db()
    if request.method == "POST":
        f = request.form
        action = f.get("action")
        if f.get("delete_zone"):
            db.execute("DELETE FROM shipping_zones WHERE id=?", (to_int(f.get("delete_zone"), 0),))
            flash("تم حذف المكان 🗑️ (الطلبات القديمة محفوظة زي ما هي)", "info")
        elif action == "add":
            name = re.sub(r"\s+", " ", (f.get("name") or "").strip())
            fee = to_float(f.get("fee"))
            if not (2 <= len(name) <= 40):
                flash("اكتب اسم المكان (من 2 لـ 40 حرف) ✍️", "danger")
            elif fee is None or fee < 0 or fee > 100000:
                flash("اكتب مصاريف شحن صحيحة ❌", "danger")
            elif db.execute("SELECT 1 FROM shipping_zones WHERE governorate=?", (name,)).fetchone():
                flash("المكان ده موجود بالفعل ❌", "danger")
            else:
                nxt = (db.execute("SELECT COALESCE(MAX(sort_order),0) m FROM shipping_zones").fetchone()["m"] or 0) + 1
                db.execute("INSERT INTO shipping_zones(governorate,fee,active,sort_order) VALUES(?,?,?,?)",
                           (name, fee, 1 if f.get("active", "1") else 0, nxt))
                flash("تم إضافة «%s» ✅" % name, "success")
        elif action == "bulk":
            fee = to_float(f.get("bulk_fee"))
            if fee is None or fee < 0 or fee > 100000:
                flash("اكتب مصاريف شحن صحيحة ❌", "danger")
            else:
                db.execute("UPDATE shipping_zones SET fee=?", (fee,))
                flash("تم ضبط مصاريف كل الأماكن على %s ج.م ✅" % fmt_money(fee), "success")
        else:
            bad = 0
            for z in db.execute("SELECT id FROM shipping_zones").fetchall():
                fee = to_float(f.get("fee_%d" % z["id"]))
                if fee is None or fee < 0 or fee > 100000:
                    bad += 1
                    continue
                db.execute("UPDATE shipping_zones SET fee=?, active=?, sort_order=? WHERE id=?",
                           (fee, 1 if f.get("active_%d" % z["id"]) else 0,
                            to_int(f.get("sort_%d" % z["id"]), 0, lo=0, hi=9999), z["id"]))
                name = re.sub(r"\s+", " ", (f.get("name_%d" % z["id"]) or "").strip())
                if name:
                    if not (2 <= len(name) <= 40) or db.execute("SELECT 1 FROM shipping_zones WHERE governorate=? AND id!=?", (name, z["id"])).fetchone():
                        bad += 1
                    else:
                        db.execute("UPDATE shipping_zones SET governorate=? WHERE id=?", (name, z["id"]))
            fm = to_float(f.get("free_min"), 0)
            if fm is None or fm < 0 or fm > 10000000:
                fm, bad = 0, bad + 1
            db.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('shipping_enabled',?)", ("1" if f.get("enabled") else "0",))
            db.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('free_shipping_min',?)", (fmt_money(fm),))
            flash("تم حفظ إعدادات الشحن ✅" + (" (فيه %d قيمة غلط اتتجاهلت ⚠️)" % bad if bad else ""), "success" if not bad else "warning")
        db.commit()
        return redirect(url_for("admin_shipping"))
    zones = db.execute("SELECT * FROM shipping_zones ORDER BY sort_order, id").fetchall()
    return render_template("admin/shipping.html", zones=zones,
                           enabled=get_setting("shipping_enabled", "1") == "1",
                           free_min=get_setting("free_shipping_min", "0"))


# ================== طرق الدفع وأرقام التحويل (إدارة) ==================
@app.route("/admin/payment_settings", methods=["GET", "POST"])
@admin_required
def admin_payment_settings():
    db = get_db()
    num_re = re.compile(r"^[0-9A-Za-z@._\-+ ]{5,60}$")

    def put(k, v):
        db.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (k, v))
    if request.method == "POST":
        f = request.form
        if not any(f.get("co_" + k) for k, _ in CHECKOUT_METHODS):
            flash("لازم تسيب طريقة دفع واحدة على الأقل شغالة وقت الشراء ❌", "danger")
            return redirect(url_for("admin_payment_settings"))
        glob = (f.get("transfer_number") or "").strip()
        if not num_re.match(glob):
            flash("رقم التحويل العام غير صحيح ❌", "danger")
            return redirect(url_for("admin_payment_settings"))
        bad = []
        put("transfer_number", glob)
        for k, label in PAY_METHODS.items():
            put("pm_%s_on" % k, "1" if f.get("on_" + k) else "0")
            n = (f.get("num_" + k) or "").strip()
            if not n or num_re.match(n):
                put("pm_%s_number" % k, n)
            else:
                bad.append(label)
        for k, _ in CHECKOUT_METHODS:
            put("co_%s_on" % k, "1" if f.get("co_" + k) else "0")
        db.commit()
        flash("تم حفظ طرق الدفع ✅" + (" — أرقام غير صحيحة اتجاهلناها: " + "، ".join(bad) if bad else ""), "success" if not bad else "warning")
        return redirect(url_for("admin_payment_settings"))
    return render_template("admin/payment_settings.html", methods=recharge_methods(), co=checkout_methods(),
                           checkout_list=CHECKOUT_METHODS,
                           glob=get_setting("transfer_number", TRANSFER_NUMBER) or TRANSFER_NUMBER)


# ================== v6: الدعم الفني ==================
SUPPORT_SUBJECTS = ["استفسار عن منتج", "متابعة طلب", "شحن وتوصيل", "الدفع والمحفظة",
                    "إرجاع أو استبدال", "شكوى أو اقتراح", "أخرى"]


@app.route("/contact", methods=["GET", "POST"])
def contact():
    db = get_db()
    if request.method == "POST":
        f = request.form
        if (f.get("website") or "").strip():          # فخ للبوتات (حقل مخفي)
            flash("وصلتنا رسالتك ✉️ هنرد عليك في أقرب وقت!", "success")
            return redirect(url_for("contact"))
        name, email = (f.get("name") or "").strip(), (f.get("email") or "").strip().lower()
        phone, msg = (f.get("phone") or "").strip(), (f.get("msg") or "").strip()
        subject = f.get("subject") if f.get("subject") in SUPPORT_SUBJECTS else "أخرى"
        if not name or len(name) > 80:
            flash("اكتب اسمك (لحد 80 حرف) ✍️", "danger")
        elif not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
            flash("الإيميل مش صحيح 📧", "danger")
        elif phone and not re.match(r"^[0-9+\-\s]{7,20}$", phone):
            flash("رقم الموبايل مش صحيح 📱", "danger")
        elif not (5 <= len(msg) <= 2000):
            flash("اكتب رسالتك (من 5 لـ 2000 حرف) ✍️", "danger")
        elif db.execute("SELECT COUNT(*) c FROM support_tickets WHERE ip=? AND created_at>=datetime('now','-1 hour')",
                        (client_ip(),)).fetchone()["c"] >= 5:
            flash("بعتّ رسايل كتير — استنى شوية وجرّب تاني 🛡️", "danger")
        else:
            cur = db.execute("""INSERT INTO support_tickets(user_id,name,email,phone,subject,message,ip)
                                VALUES(?,?,?,?,?,?,?)""",
                             (session.get("user_id"), name, email, phone, subject, msg, client_ip()))
            db.commit()
            flash("وصلتنا رسالتك ✉️ رقم التذكرة #%d — هنرد عليك في أقرب وقت!" % cur.lastrowid, "success")
        return redirect(url_for("contact"))
    mine = []
    if session.get("user_id"):
        mine = db.execute("SELECT * FROM support_tickets WHERE user_id=? ORDER BY id DESC LIMIT 10",
                          (session["user_id"],)).fetchall()
    u = db.execute("SELECT name,email,phone FROM users WHERE id=?", (session["user_id"],)).fetchone() if session.get("user_id") else None
    return render_template("contact.html", subjects=SUPPORT_SUBJECTS, mine=mine, u=u,
                           free_min=to_float(get_setting("free_shipping_min", "0"), 0) or 0,
                           wa_link=site_info()["wa_link"])


@app.route("/admin/support", methods=["GET", "POST"])
@admin_required
def admin_support():
    db = get_db()
    if request.method == "POST":
        f = request.form
        tid = to_int(f.get("id"), 0)
        action = f.get("action")
        if action == "reply":
            reply = (f.get("reply") or "").strip()[:2000]
            st = f.get("status") if f.get("status") in TICKET_STATUSES else "قيد المتابعة"
            db.execute("UPDATE support_tickets SET admin_reply=?, status=?, replied_at=? WHERE id=?",
                       (reply, st if reply else st, datetime.now().strftime("%Y-%m-%d %H:%M") if reply else None, tid))
            flash("تم حفظ الرد ✅ (بيظهر للعميل في صفحة الدعم لو مسجّل دخول)", "success")
        elif action == "status" and f.get("status") in TICKET_STATUSES:
            db.execute("UPDATE support_tickets SET status=? WHERE id=?", (f["status"], tid))
            flash("تم تحديث الحالة ✅", "success")
        elif action == "delete":
            db.execute("DELETE FROM support_tickets WHERE id=?", (tid,))
            flash("تم حذف التذكرة 🗑️", "info")
        db.commit()
        return redirect(safe_back(url_for("admin_support")))
    status = request.args.get("status", "")
    sql = "SELECT * FROM support_tickets" + (" WHERE status=?" if status in TICKET_STATUSES else "") + " ORDER BY id DESC LIMIT 500"
    rows = db.execute(sql, (status,) if status in TICKET_STATUSES else ()).fetchall()
    counts = {r["status"]: r["c"] for r in db.execute("SELECT status, COUNT(*) c FROM support_tickets GROUP BY status").fetchall()}
    return render_template("admin/support.html", rows=rows, statuses=TICKET_STATUSES, status_filter=status,
                           counts=counts, intl=intl_number)


# ================== v6: سجل المدفوعات (عرض / حذف / تصدير) ==================
def _pay_filters(args):
    return {"kind": args.get("kind") if args.get("kind") in ("tx", "req", "orders") else "tx",
            "q": (args.get("q") or "").strip()[:60], "type": (args.get("type") or "").strip(),
            "from": valid_date(args.get("from")), "to": valid_date(args.get("to"))}


def _pay_query(fl, kind=None):
    """يرجّع (SELECT, FROM+WHERE, params) حسب النوع والفلاتر"""
    kind = kind or fl["kind"]
    params, where = [], ["1=1"]
    if kind == "tx":
        frm = "FROM wallet_transactions t JOIN users u ON u.id=t.user_id"
        col = "t.created_at"
        sel = "SELECT t.id, t.user_id, u.name uname, u.email, t.amount, t.type, t.note, t.created_at "
        if fl["type"] in TX_TYPES:
            where.append("t.type=?"); params.append(fl["type"])
    elif kind == "req":
        frm = "FROM wallet_requests t JOIN users u ON u.id=t.user_id"
        col = "t.created_at"
        sel = "SELECT t.id, t.user_id, u.name uname, u.email, t.amount, t.method, t.sender_number, t.status, t.created_at, t.handled_at "
        if fl["type"] in ("pending", "approved", "rejected"):
            where.append("t.status=?"); params.append(fl["type"])
    else:
        frm = "FROM orders t LEFT JOIN users u ON u.id=t.user_id"
        col = "t.created_at"
        sel = "SELECT t.id, t.user_id, u.name uname, u.email, t.total, t.paid, t.payment_method, t.status, t.governorate, t.created_at "
        if fl["type"] in ("cod", "balance", "partial"):
            where.append("t.payment_method=?"); params.append(fl["type"])
    if fl["q"]:
        where.append("(u.name LIKE ? OR u.email LIKE ?)")
        params += ["%" + fl["q"] + "%"] * 2
    if fl["from"]:
        where.append("date(%s)>=?" % col); params.append(fl["from"])
    if fl["to"]:
        where.append("date(%s)<=?" % col); params.append(fl["to"])
    return sel, frm + " WHERE " + " AND ".join(where), params


def _pay_sheets(db, fl, kinds):
    sheets = []
    for k in kinds:
        sel, fw, params = _pay_query(fl, k)
        rows = db.execute(sel + fw + " ORDER BY t.id DESC LIMIT 50000", params).fetchall()
        if k == "tx":
            sheets.append(("حركات الرصيد", ["#", "العميل", "الإيميل", "المبلغ", "النوع", "ملاحظة", "التاريخ"],
                           [[r["id"], r["uname"], r["email"], r["amount"], TX_TYPES.get(r["type"], r["type"]),
                             r["note"] or "", r["created_at"]] for r in rows]))
        elif k == "req":
            st = {"pending": "معلق", "approved": "موافق", "rejected": "مرفوض"}
            sheets.append(("طلبات الشحن", ["#", "العميل", "الإيميل", "المبلغ", "الوسيلة", "رقم المحوّل", "الحالة",
                                          "تاريخ الطلب", "تاريخ المراجعة"],
                           [[r["id"], r["uname"], r["email"], r["amount"], PAY_METHODS.get(r["method"], r["method"]),
                             r["sender_number"] or "", st.get(r["status"], r["status"]), r["created_at"],
                             r["handled_at"] or ""] for r in rows]))
        else:
            sheets.append(("مدفوعات الطلبات", ["رقم الطلب", "العميل", "الإيميل", "الإجمالي", "المدفوع", "المتبقي",
                                               "طريقة الدفع", "الحالة", "المحافظة", "التاريخ"],
                           [[r["id"], r["uname"] or "حساب محذوف", r["email"] or "", r["total"], r["paid"] or 0,
                             round(max((r["total"] or 0) - (r["paid"] or 0), 0), 2),
                             PAY_LABELS.get(r["payment_method"], "كاش عند الاستلام"), r["status"],
                             r["governorate"] or "", r["created_at"]] for r in rows]))
    return sheets


@app.route("/admin/payments", methods=["GET", "POST"])
@admin_required
def admin_payments():
    db = get_db()
    if request.method == "POST":
        f = request.form
        fl = _pay_filters(f)
        kind, scope = fl["kind"], f.get("scope")
        back = url_for("admin_payments", kind=kind, q=fl["q"], type=fl["type"], **{"from": fl["from"], "to": fl["to"]})
        if kind == "orders":
            flash("مدفوعات الطلبات جزء من الطلبات نفسها — مينفعش تتحذف من هنا 🚫", "danger")
            return redirect(back)
        table, extra = ("wallet_transactions", "") if kind == "tx" else ("wallet_requests", " AND status!='pending'")
        if scope == "selected":
            ids = [to_int(x, 0) for x in f.getlist("ids")]
            ids = [i for i in ids if i > 0][:5000]
            if not ids:
                flash("اختار سجلات الأول ☑️", "warning")
                return redirect(back)
            make_backup("predelete")
            db.execute("DELETE FROM %s WHERE id IN (%s)%s" % (table, ",".join("?" * len(ids)), extra), ids)
        elif scope in ("filtered", "all"):
            need = "حذف" if scope == "filtered" else "حذف الكل"
            if (f.get("confirm_word") or "").strip() != need:
                flash("اكتب «%s» للتأكيد ❌" % need, "danger")
                return redirect(back)
            make_backup("predelete")
            if scope == "all":
                db.execute("DELETE FROM %s WHERE 1=1%s" % (table, extra))
            else:
                _, fw, params = _pay_query(fl)
                db.execute("DELETE FROM %s WHERE id IN (SELECT t.id %s)%s" % (table, fw, extra), params)
        else:
            return redirect(back)
        db.commit()
        flash("تم الحذف ✅ (واتعملت نسخة احتياطية قبلها — الأرصدة الحالية للعملاء ماتغيّرتش)", "success")
        return redirect(back)
    fl = _pay_filters(request.args)
    sel, fw, params = _pay_query(fl)
    page = to_int(request.args.get("page"), 1, lo=1)
    per = 50
    total_rows = db.execute("SELECT COUNT(*) c " + fw, params).fetchone()["c"]
    pages = max((total_rows + per - 1) // per, 1)
    page = min(page, pages)
    rows = db.execute(sel + fw + " ORDER BY t.id DESC LIMIT ? OFFSET ?", params + [per, (page - 1) * per]).fetchall()
    if fl["kind"] == "tx":
        agg = db.execute("SELECT COALESCE(SUM(CASE WHEN t.amount>0 THEN t.amount END),0) a, "
                         "COALESCE(SUM(CASE WHEN t.amount<0 THEN t.amount END),0) b " + fw, params).fetchone()
        sums = {"in": agg["a"], "out": agg["b"]}
    elif fl["kind"] == "req":
        sums = {"in": db.execute("SELECT COALESCE(SUM(t.amount),0) s " + fw + " AND t.status='approved'", params).fetchone()["s"], "out": 0}
    else:
        agg = db.execute("SELECT COALESCE(SUM(t.total),0) a, COALESCE(SUM(t.paid),0) b " + fw, params).fetchone()
        sums = {"in": agg["b"], "out": agg["a"]}
    return render_template("admin/payments.html", fl=fl, rows=rows, page=page, pages=pages, total_rows=total_rows,
                           sums=sums, tx_types=TX_TYPES, pay_methods=PAY_LABELS)


@app.route("/admin/payments/export")
@admin_required
def admin_payments_export():
    fl = _pay_filters(request.args)
    kind = request.args.get("kind")
    kinds = ["tx", "req", "orders"] if kind == "all" else [fl["kind"]]
    fmt = "xlsx" if request.args.get("format") == "xlsx" else "csv"
    return export_response(_pay_sheets(get_db(), fl, kinds), "payments", fmt)


# ================== التحليلات الشاملة (موسّعة v6) ==================
def _analytics_data(db, days):
    since = "-%d days" % (days - 1)
    ok = "status!='ملغي'"
    one = lambda sql, p=(): db.execute(sql, p).fetchone()[0]
    kpi = {
        "revenue": one("SELECT COALESCE(SUM(total),0) FROM orders WHERE " + ok),
        "orders": one("SELECT COUNT(*) FROM orders"),
        "users": one("SELECT COUNT(*) FROM users WHERE is_admin=0"),
        "avg": one("SELECT COALESCE(AVG(total),0) FROM orders WHERE " + ok),
        "aff_paid": one("SELECT COALESCE(SUM(commission),0) FROM affiliate_orders WHERE status='confirmed'"),
        "aff_pending": one("SELECT COALESCE(SUM(commission),0) FROM affiliate_orders WHERE status='pending'"),
        "returns": one("SELECT COUNT(*) FROM returns WHERE status='approved'"),
        "wallet_in": one("SELECT COALESCE(SUM(amount),0) FROM wallet_requests WHERE status='approved'"),
        "shipping": one("SELECT COALESCE(SUM(shipping),0) FROM orders WHERE " + ok),
        "wa_total": one("SELECT COUNT(*) FROM whatsapp_orders"),
        "wa_new": one("SELECT COUNT(*) FROM whatsapp_orders WHERE status='جديد'"),
        "tickets_open": one("SELECT COUNT(*) FROM support_tickets WHERE status!='تم الحل'"),
        "low_stock": one("SELECT COUNT(*) FROM products WHERE stock<=5"),
        "range_orders": one("SELECT COUNT(*) FROM orders WHERE date(created_at)>=date('now',?)", (since,)),
        "range_revenue": one("SELECT COALESCE(SUM(total),0) FROM orders WHERE " + ok + " AND date(created_at)>=date('now',?)", (since,)),
        "range_users": one("SELECT COUNT(*) FROM users WHERE is_admin=0 AND date(created_at)>=date('now',?)", (since,)),
        "balances": one("SELECT COALESCE(SUM(balance),0) FROM users WHERE is_admin=0"),
    }
    grp = "date(created_at)" if days <= 31 else "strftime('%Y-W%W', created_at)"
    daily = db.execute("SELECT %s d, SUM(total) s, COUNT(*) c FROM orders WHERE date(created_at)>=date('now',?) AND %s "
                       "GROUP BY d ORDER BY d" % (grp, ok), (since,)).fetchall()
    d = {
        "kpi": kpi, "daily": daily,
        "maxd": max([r["s"] for r in daily], default=1) or 1,
        "status_counts": db.execute("SELECT status, COUNT(*) c FROM orders GROUP BY status").fetchall(),
        "top_products": db.execute("SELECT name, emoji, sold, price FROM products ORDER BY sold DESC LIMIT 5").fetchall(),
        "top_customers": db.execute("""SELECT u.name, u.email, COUNT(o.id) orders_count, SUM(o.total) spent
                                       FROM orders o JOIN users u ON u.id=o.user_id
                                       GROUP BY o.user_id ORDER BY spent DESC LIMIT 5""").fetchall(),
        "top_marketers": db.execute("""SELECT u.name, m.code, COUNT(a.id) orders_count,
                                       COALESCE(SUM(CASE WHEN a.status='confirmed' THEN a.commission END),0) paid
                                       FROM marketers m JOIN users u ON u.id=m.user_id
                                       LEFT JOIN affiliate_orders a ON a.marketer_id=m.id
                                       GROUP BY m.id ORDER BY orders_count DESC LIMIT 5""").fetchall(),
        "cat_sales": db.execute("""SELECT p.category, SUM(oi.qty) q, SUM(oi.qty*oi.price) s
                                   FROM order_items oi JOIN products p ON p.id=oi.product_id
                                   GROUP BY p.category ORDER BY s DESC""").fetchall(),
        "gov_sales": db.execute("""SELECT COALESCE(NULLIF(governorate,''),'غير محدد') g, COUNT(*) c, SUM(total) s
                                   FROM orders WHERE %s AND date(created_at)>=date('now',?)
                                   GROUP BY g ORDER BY s DESC LIMIT 10""" % ok, (since,)).fetchall(),
        "pay_split": db.execute("""SELECT payment_method m, COUNT(*) c, SUM(total) s FROM orders
                                   WHERE %s AND date(created_at)>=date('now',?) GROUP BY m ORDER BY c DESC""" % ok, (since,)).fetchall(),
        "range_top": db.execute("""SELECT p.name, p.emoji, SUM(oi.qty) q, SUM(oi.qty*oi.price) s
                                   FROM order_items oi JOIN orders o ON o.id=oi.order_id
                                   JOIN products p ON p.id=oi.product_id
                                   WHERE o.%s AND date(o.created_at)>=date('now',?)
                                   GROUP BY oi.product_id ORDER BY s DESC LIMIT 5""" % ok, (since,)).fetchall(),
    }
    return d


def _analytics_days():
    d = to_int(request.args.get("days"), 14)
    return d if d in (7, 14, 30, 90, 365) else 14


@app.route("/admin/analytics")
@admin_required
def admin_analytics():
    days = _analytics_days()
    data = _analytics_data(get_db(), days)
    return render_template("admin/analytics.html", days=days, pay_methods=PAY_LABELS, **data)


@app.route("/admin/analytics/export")
@admin_required
def admin_analytics_export():
    days = _analytics_days()
    d = _analytics_data(get_db(), days)
    k = d["kpi"]
    summary = [["إجمالي المبيعات (ج.م)", round(k["revenue"], 2)], ["عدد الطلبات", k["orders"]],
               ["متوسط قيمة الطلب", round(k["avg"], 2)], ["عدد العملاء", k["users"]],
               ["إيراد مصاريف الشحن", round(k["shipping"], 2)], ["طلبات واتساب (الكل)", k["wa_total"]],
               ["طلبات واتساب جديدة", k["wa_new"]], ["تذاكر دعم مفتوحة", k["tickets_open"]],
               ["أرصدة العملاء الحالية", round(k["balances"], 2)],
               ["طلبات آخر %d يوم" % days, k["range_orders"]], ["مبيعات آخر %d يوم" % days, round(k["range_revenue"], 2)],
               ["عملاء جدد آخر %d يوم" % days, k["range_users"]]]
    sheets = [("ملخص", ["المؤشر", "القيمة"], summary),
              ("المبيعات", ["الفترة", "المبيعات", "عدد الطلبات"], [[r["d"], round(r["s"], 2), r["c"]] for r in d["daily"]]),
              ("المحافظات", ["المحافظة", "عدد الطلبات", "المبيعات"], [[r["g"], r["c"], round(r["s"], 2)] for r in d["gov_sales"]]),
              ("أفضل المنتجات", ["المنتج", "الكمية", "الإيراد"], [[r["name"], r["q"], round(r["s"], 2)] for r in d["range_top"]])]
    return export_response(sheets, "analytics", "xlsx" if request.args.get("format") == "xlsx" else "csv")


# ================== إعدادات المميزات (v6: + بيانات التواصل والنسخ الاحتياطي) ==================
SETTING_FLAGS = ["wallet_sender", "affiliate", "returns", "whatsapp_cart", "support_whatsapp_on"]


@app.route("/admin/settings", methods=["GET", "POST"])
@admin_required
def admin_settings():
    db = get_db()

    def put(k, v):
        db.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (k, v))
    if request.method == "POST":
        f = request.form
        for k in SETTING_FLAGS:
            put(k, "1" if f.get(k) else "0")
        warn = []
        wa = intl_number(f.get("whatsapp_number", ""))
        if wa and 10 <= len(wa) <= 15:
            put("whatsapp_number", f.get("whatsapp_number", "").strip())
        elif f.get("whatsapp_number", "").strip():
            warn.append("رقم واتساب")
        ph = (f.get("support_phone") or "").strip()
        if re.match(r"^[0-9+\-\s]{7,20}$", ph):
            put("support_phone", ph)
        elif ph:
            warn.append("رقم الدعم")
        em = (f.get("support_email") or "").strip()
        if re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", em):
            put("support_email", em)
        elif em:
            warn.append("إيميل الدعم")
        sw = (f.get("support_whatsapp") or "").strip()
        if 10 <= len(intl_number(sw)) <= 15:
            put("support_whatsapp", sw)
        elif sw:
            warn.append("رقم واتساب الدعم")
        put("support_hours", (f.get("support_hours") or "").strip()[:80])
        put("store_address", (f.get("store_address") or "").strip()[:160])
        put("backup_keep", str(to_int(f.get("backup_keep"), 14, lo=1, hi=90)))
        db.commit()
        if warn:
            flash("القيم دي غير صحيحة واتجاهلناها: %s — اتحفظ الباقي ⚠️" % "، ".join(warn), "warning")
        flash("تم حفظ الإعدادات ✅", "success")
        return redirect(url_for("admin_settings"))
    s = {k: get_setting(k, "1") == "1" for k in SETTING_FLAGS}
    return render_template("admin/settings.html", s=s,
                           wa_number=get_setting("whatsapp_number", "") or WHATSAPP_NUMBER,
                           info=site_info(), backup_keep=get_setting("backup_keep", "14"))


# تجهيز قاعدة البيانات (آمن يتكرر) — بيشتغل كمان لو الموقع متشغّل عبر WSGI/gunicorn
try:
    init_db()
    start_backup_scheduler()
except Exception as _e:   # pragma: no cover
    print("init_db warning:", _e)

if __name__ == "__main__":
    init_db()
    # الـ debug متقفل افتراضياً (خطر أمني) — لتشغيله محلياً: FLASK_DEBUG=1 في .env
    app.run(debug=os.environ.get("FLASK_DEBUG", "0") == "1",
            host=os.environ.get("HOST", "127.0.0.1"),
            port=to_int(os.environ.get("PORT", 5000), 5000))