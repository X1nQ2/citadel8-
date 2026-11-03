"""Онлайн-магазин «Генацидная цитадель номер 8» — Flask + Access (.accdb).

Запуск:  python server.py   →  http://127.0.0.1:5000
База:    database.accdb рядом с server.py (или переменная окружения DB_PATH).
"""
import json
import os
import random
import re
import secrets
import threading
import time
from contextlib import contextmanager
from collections import Counter
from datetime import date, datetime
from functools import wraps

import pyodbc
from flask import (Flask, abort, flash, jsonify, redirect, render_template,
                   request, session, url_for)
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(BASE_DIR, "static", "book_images")
DB_PATH = os.environ.get("DB_PATH", os.path.join(BASE_DIR, "database.accdb"))
CONN_STR = f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};Dbq={DB_PATH};"

BRAND = "Генацидная цитадель номер 8"
ROLE_ADMIN = "Админ"
ROLE_USER = "Покупатель"
STATUS_NEW, STATUS_DONE, STATUS_CANCELED = "В работе", "Выполнен", "Отменён"
ORDER_STATUSES = [STATUS_NEW, STATUS_DONE, STATUS_CANCELED]
IMAGE_EXT = ("jpg", "jpeg", "png", "webp")
HASH_PREFIXES = ("scrypt:", "pbkdf2:")
MAX_QTY = 20
BOOK_OF_DAY_DAILY = True   # True: «Книга дня» одна на сутки; False: новая случайная при каждом обновлении
HAS_RATING = True          # есть ли таблица [Рейтинг] (определяется при старте)

# Реквизиты и контакты магазина — ЗАМЕНИТЕ на свои (их видно в подвале и на страницах документов)
SHOP = {
    "email": os.environ.get("SHOP_EMAIL", "support@citadel8.example"),
    "phone": os.environ.get("SHOP_PHONE", "+7 (000) 000-00-00"),
    "hours": os.environ.get("SHOP_HOURS", "ежедневно с 9:00 до 21:00"),
    "address": os.environ.get("SHOP_ADDRESS", ""),   # адрес офиса/пункта выдачи (необязательно)
    "requisites": os.environ.get("SHOP_REQUISITES", ""),  # ИП/ООО, ИНН, ОГРН (необязательно)
    "delivery": os.environ.get(
        "SHOP_DELIVERY",
        "Сроки и стоимость доставки зависят от региона и уточняются при подтверждении заказа."),
    "legal_date": "29.10.2026",
}
HAS_ADDR = True    # есть ли в [Заказ] колонка «Адрес доставки» (определяется при старте)
HAS_ORDER_META = True
HAS_ACTIVE = True  # есть ли в [Товар] колонка Активен (определяется при старте)


def hash_pw(pwd):
    return generate_password_hash(pwd, method="pbkdf2:sha256")


def active_filter(prefix="t."):
    return f"({prefix}[Активен] IS NULL OR {prefix}[Активен] = True)" if HAS_ACTIVE else "1=1"


def _secret_key() -> str:
    key = os.environ.get("SECRET_KEY")
    if key:
        return key
    path = os.path.join(BASE_DIR, "secret.key")
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            f.write(secrets.token_hex(32))
    with open(path, encoding="utf-8") as f:
        return f.read().strip()


app = Flask(__name__)  # статика только из ./static — база и код недоступны по URL
app.secret_key = _secret_key()
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    MAX_CONTENT_LENGTH=3 * 1024 * 1024,
)


# ====================== БД ======================
@contextmanager
def db():
    """Соединение с коммитом при успехе и откатом при ошибке."""
    conn = pyodbc.connect(CONN_STR)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def last_id(cur, table, id_col):
    """ID только что вставленной строки (в рамках этого соединения)."""
    try:
        cur.execute("SELECT @@IDENTITY")
        val = cur.fetchone()[0]
        if val:
            return int(val)
    except pyodbc.Error:
        pass
    cur.execute(f"SELECT MAX({id_col}) FROM [{table}]")
    return int(cur.fetchone()[0])


def ensure_schema():
    """Идемпотентная подготовка структуры. Данные не удаляются."""
    with db() as conn:
        cur = conn.cursor()

        def ok(sql):
            try:
                cur.execute(sql)
                return True
            except pyodbc.Error:
                return False

        # 1. Логин / Пароль / Роль у покупателей
        if not ok("SELECT Логин, Пароль, Роль FROM [Покупатель] WHERE 1=0"):
            ok("ALTER TABLE [Покупатель] ADD COLUMN [Логин] TEXT(50)")
            ok("ALTER TABLE [Покупатель] ADD COLUMN [Пароль] TEXT(255)")
            ok("ALTER TABLE [Покупатель] ADD COLUMN [Роль] TEXT(50)")

        # 2. Поле пароля должно вмещать хеш (255 символов)
        try:
            cols = {r.column_name.lower(): r.column_size
                    for r in cur.columns(table="Покупатель")}
            if cols.get("пароль", 255) < 255:
                ok("ALTER TABLE [Покупатель] ALTER COLUMN [Пароль] TEXT(255)")
        except pyodbc.Error:
            pass

        # 3. Мягкое удаление книг (сохраняет историю заказов)
        if not ok("SELECT Активен FROM [Товар] WHERE 1=0"):
            ok("ALTER TABLE [Товар] ADD COLUMN [Активен] YESNO")
            ok("UPDATE [Товар] SET [Активен] = True")

        # 4. Адрес доставки в заказе (старые заказы получают адрес покупателя)
        if not ok("SELECT [Адрес доставки] FROM [Заказ] WHERE 1=0"):
            if ok("ALTER TABLE [Заказ] ADD COLUMN [Адрес доставки] TEXT(255)"):
                ok("UPDATE [Заказ] AS z INNER JOIN [Покупатель] AS p ON z.id_user = p.id_user "
                   "SET z.[Адрес доставки] = p.[Адресс]")

        # 5. Данные доставки и оплаты. Миграция безопасна для существующих заказов.
        order_columns = [
            ("Код доставки", "TEXT(24)"),
            ("Дата и время", "DATETIME"),
            ("Способ доставки", "TEXT(60)"),
            ("Способ оплаты", "TEXT(60)"),
            ("Статус оплаты", "TEXT(40)"),
            ("Комментарий", "TEXT(255)"),
        ]
        for column, sql_type in order_columns:
            if not ok(f"SELECT [{column}] FROM [Заказ] WHERE 1=0"):
                ok(f"ALTER TABLE [Заказ] ADD COLUMN [{column}] {sql_type}")

        # Заполняем кодами старые заказы, у которых поле пустое.
        try:
            cur.execute("SELECT id_order FROM [Заказ] WHERE [Код доставки] IS NULL OR [Код доставки] = ''")
            for (old_id,) in cur.fetchall():
                code = make_delivery_code(cur)
                cur.execute("UPDATE [Заказ] SET [Код доставки] = ? WHERE id_order = ?", code, old_id)
            cur.execute("UPDATE [Заказ] SET [Способ доставки] = 'Курьерская доставка' "
                        "WHERE [Способ доставки] IS NULL OR [Способ доставки] = ''")
            cur.execute("UPDATE [Заказ] SET [Способ оплаты] = 'При получении' "
                        "WHERE [Способ оплаты] IS NULL OR [Способ оплаты] = ''")
            cur.execute("UPDATE [Заказ] SET [Статус оплаты] = 'Ожидает оплаты' "
                        "WHERE [Статус оплаты] IS NULL OR [Статус оплаты] = ''")
            ok("CREATE UNIQUE INDEX UX_Order_DeliveryCode ON [Заказ] ([Код доставки])")
        except pyodbc.Error as e:
            print(f"[warn] метаданные заказов: {e}")

        # 6. Оценки книг (1–5, одна оценка пользователя на книгу)
        if not ok("SELECT id_tov FROM [Рейтинг] WHERE 1=0"):
            ok("CREATE TABLE [Рейтинг] (id_rating COUNTER PRIMARY KEY, id_tov LONG, "
               "id_user LONG, Балл INTEGER, Дата DATETIME)")

        # 6. Открытые пароли → хеши
        cur.execute("SELECT id_user, Пароль FROM [Покупатель] WHERE Пароль IS NOT NULL")
        for uid, pwd in cur.fetchall():
            if pwd and not pwd.startswith(HASH_PREFIXES):
                cur.execute("UPDATE [Покупатель] SET Пароль = ? WHERE id_user = ?",
                            hash_pw(pwd), uid)


def detect_schema():
    global HAS_ACTIVE, HAS_ADDR, HAS_RATING
    try:
        with db() as conn:
            conn.cursor().execute("SELECT id_tov, Балл FROM [Рейтинг] WHERE 1=0")
        HAS_RATING = True
    except pyodbc.Error:
        HAS_RATING = False
        print("[warn] таблица [Рейтинг] недоступна: оценки отключены")
    try:
        with db() as conn:
            conn.cursor().execute("SELECT [Адрес доставки] FROM [Заказ] WHERE 1=0")
        HAS_ADDR = True
    except pyodbc.Error:
        HAS_ADDR = False
        print("[warn] колонка [Заказ].[Адрес доставки] недоступна: адрес берётся из профиля")
    try:
        with db() as conn:
            conn.cursor().execute("SELECT Активен FROM [Товар] WHERE 1=0")
        HAS_ACTIVE = True
    except pyodbc.Error:
        HAS_ACTIVE = False
        print("[warn] колонка [Товар].Активен недоступна: удаление книг будет полным")


# ====================== Вспомогательное ======================
@app.template_filter("rub")
def rub(value):
    try:
        v = float(value)
    except (TypeError, ValueError):
        return "—"
    s = f"{v:,.0f}" if v == int(v) else f"{v:,.2f}"
    return s.replace(",", " ") + " ₽"


@app.template_filter("status_cls")
def status_cls(value):
    return {STATUS_DONE: "done", STATUS_CANCELED: "canceled"}.get(value, "")


@app.template_filter("fmtdate")
def fmtdate(value):
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    return str(value or "—")[:10]


@app.template_filter("fmtdatetime")
def fmtdatetime(value):
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y · %H:%M")
    return str(value or "—")


def cover_url(book_id):
    for ext in IMAGE_EXT:
        if os.path.exists(os.path.join(IMG_DIR, f"{book_id}.{ext}")):
            return url_for("static", filename=f"book_images/{book_id}.{ext}")
    return url_for("static", filename="book_images/default.jpg")


def sniff_image(head: bytes):
    if head.startswith(b"\xff\xd8\xff"):
        return "jpg"
    if head.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if head[:4] == b"RIFF" and head[8:12] == b"WEBP":
        return "webp"
    return None


def safe_next(target):
    if target and target.startswith("/") and not target.startswith("//"):
        return target
    return url_for("catalog")


def fetch_books(cur, book_id=None):
    sql = """
        SELECT t.id_tov, t.Название, t.ISBN, t.Цена, t.Колво,
               a.ФИО, j.Название, i.Название, t.[Год издания]
        FROM (((([Товар] AS t
        LEFT JOIN [АвторТовар] AS ta ON t.id_tov = ta.id_tov)
        LEFT JOIN [Автор] AS a ON ta.id_aut = a.id_aut)
        LEFT JOIN [ЖанрТовар] AS jt ON t.id_tov = jt.id_tov)
        LEFT JOIN [Жанр] AS j ON jt.id_janr = j.id_janr)
        LEFT JOIN [Издательство] AS i ON t.id_izd = i.id_izd
        WHERE """ + active_filter("t.") + """
    """
    params = []
    if book_id is not None:
        sql += " AND t.id_tov = ?"
        params.append(book_id)
    sql += " ORDER BY t.id_tov"
    cur.execute(sql, params)

    books = {}
    for r in cur.fetchall():
        b = books.setdefault(r[0], {
            "id": r[0], "title": (r[1] or "").strip(), "isbn": r[2] or "—",
            "price": float(r[3] or 0), "stock": int(r[4] or 0),
            "authors": [], "genres": [],
            "publisher": r[7] or "—", "year": r[8] or "—",
            "cover": cover_url(r[0]),
        })
        author = (r[5] or "").strip()
        genre = (r[6] or "").strip()
        if author and author not in b["authors"]:
            b["authors"].append(author)
        if genre and genre not in b["genres"]:
            b["genres"].append(genre)
    out = []
    for b in books.values():
        b["author"] = ", ".join(sorted(b["authors"])) or "Не указан"
        b["genre"] = ", ".join(sorted(b["genres"])) or "Без жанра"
        del b["authors"], b["genres"]
        out.append(b)
    return out



# ====================== Корзина ======================
# Корзина привязана к пользователю и хранится в carts.json: переживает выход и перезапуск.
CARTS_PATH = os.path.join(BASE_DIR, "carts.json")
_carts_lock = threading.Lock()


def _load_carts():
    try:
        with open(CARTS_PATH, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


CARTS = _load_carts()  # {"id_user": {"id_tov": qty}}


def get_cart():
    raw = CARTS.get(str(session.get("user_id")), {})
    return {int(k): int(v) for k, v in raw.items()}


def save_cart(cart):
    uid = str(session["user_id"])
    with _carts_lock:
        if cart:
            CARTS[uid] = {str(k): int(v) for k, v in cart.items()}
        else:
            CARTS.pop(uid, None)
        tmp = CARTS_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(CARTS, f)
        os.replace(tmp, CARTS_PATH)


class CartError(Exception):
    pass


# ====================== Безопасность ======================
@app.before_request
def csrf_protect():
    if request.method == "POST":
        token = session.get("csrf")
        if not token or not secrets.compare_digest(token, request.form.get("csrf", "")):
            abort(400)


@app.context_processor
def inject():
    session.setdefault("csrf", secrets.token_hex(16))
    return dict(
        csrf=session["csrf"],
        brand=BRAND,
        is_admin=session.get("role") == ROLE_ADMIN,
        cur_login=session.get("login"),
        cur_role=session.get("role"),
        cart_count=sum(get_cart().values()) if "user_id" in session else 0,
        has_addr=HAS_ADDR,
        has_rating=HAS_RATING,
        shop=SHOP,
        STATUS_NEW=STATUS_NEW,
        STATUS_DONE=STATUS_DONE,
        STATUS_CANCELED=STATUS_CANCELED,
    )


@app.after_request
def headers(resp):
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    if request.endpoint in ("login", "register"):
        resp.headers["Cache-Control"] = "no-store"
    return resp


def login_required(f):
    @wraps(f)
    def wrapper(*a, **k):
        if "user_id" not in session:
            flash("Войдите или зарегистрируйтесь, чтобы оформить заказ")
            return redirect(url_for("login", next=request.path))
        return f(*a, **k)
    return wrapper


def admin_required(f):
    @wraps(f)
    def wrapper(*a, **k):
        if "user_id" not in session:
            return redirect(url_for("login", next=request.path))
        if session.get("role") != ROLE_ADMIN:  # проверка прав на сервере
            abort(403)
        return f(*a, **k)
    return wrapper


FAILS = {}  # защита от подбора пароля: 5 ошибок за 5 минут


def throttled(key):
    now = time.time()
    FAILS[key] = [t for t in FAILS.get(key, []) if now - t < 300]
    return len(FAILS[key]) >= 5


def start_session(user_id, login, role):
    session.clear()
    session.update(user_id=user_id, login=login, role=role or ROLE_USER,
                   csrf=secrets.token_hex(16))


def verify_pw(stored, pwd):
    if not stored:
        return False
    if stored.startswith(HASH_PREFIXES):
        return check_password_hash(stored, pwd)
    return secrets.compare_digest(stored, pwd)


def make_delivery_code(cur):
    """Человекочитаемый уникальный код выдачи; проверяется непосредственно в БД."""
    alphabet = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"
    for _ in range(30):
        code = "DLV-" + "".join(secrets.choice(alphabet) for _ in range(8))
        cur.execute("SELECT COUNT(*) FROM [Заказ] WHERE [Код доставки] = ?", code)
        if int(cur.fetchone()[0] or 0) == 0:
            return code
    raise RuntimeError("Не удалось сгенерировать уникальный код доставки")


def insert_order(cur, user_id, total, address, delivery_method="Курьерская доставка",
                 payment_method="При получении", comment=""):
    """Создаёт заказ, возвращает (id_order, номер, код доставки)."""
    now = datetime.now()
    number_placeholder = "ORD-000"
    delivery_code = make_delivery_code(cur)
    cols = ("id_user, [Номер заказа], [Дата создания], Статус, Сумма, "
            "[Код доставки], [Дата и время], [Способ доставки], [Способ оплаты], "
            "[Статус оплаты], [Комментарий]")
    vals = [user_id, number_placeholder, now.strftime("%d.%m.%Y"), STATUS_NEW, total,
            delivery_code, now, delivery_method, payment_method,
            "Ожидает оплаты" if payment_method == "При получении" else "Ожидает подтверждения",
            comment]
    if HAS_ADDR:
        cols += ", [Адрес доставки]"
        vals.append(address)
    cur.execute(f"INSERT INTO [Заказ] ({cols}) VALUES ({', '.join('?' * len(vals))})", *vals)
    id_order = last_id(cur, "Заказ", "id_order")
    number = f"ORD-{id_order:03d}"
    cur.execute("UPDATE [Заказ] SET [Номер заказа] = ? WHERE id_order = ?", number, id_order)
    return id_order, number, delivery_code


# ====================== Каталог ======================
def rating_stats(cur):
    """{id_tov: (средняя оценка, число оценок)}"""
    if not HAS_RATING:
        return {}
    try:
        cur.execute("SELECT id_tov, AVG(Балл), COUNT(*) FROM [Рейтинг] GROUP BY id_tov")
        return {r[0]: (float(r[1] or 0), int(r[2] or 0)) for r in cur.fetchall()}
    except pyodbc.Error as e:
        print(f"[warn] рейтинг недоступен: {e}")
        return {}


def my_ratings(cur, user_id):
    if not HAS_RATING or not user_id:
        return {}
    try:
        cur.execute("SELECT id_tov, Балл FROM [Рейтинг] WHERE id_user = ?", user_id)
        return {r[0]: int(r[1] or 0) for r in cur.fetchall()}
    except pyodbc.Error:
        return {}


def purchase_stats(cur, user_id):
    """(популярность книг по заказам, множество id книг, купленных этим пользователем)"""
    popularity, mine = Counter(), set()
    try:
        cur.execute("SELECT p.id_tov, z.id_user, p.Колво FROM [ПозицияЗаказа] AS p "
                    "INNER JOIN [Заказ] AS z ON p.id_order = z.id_order WHERE z.Статус <> ?",
                    STATUS_CANCELED)
        for tid, uid, qty in cur.fetchall():
            popularity[tid] += int(qty or 1)
            if user_id and uid == user_id:
                mine.add(tid)
    except pyodbc.Error as e:
        print(f"[warn] статистика покупок недоступна: {e}")
    return popularity, mine


def book_of_day(books):
    pool = sorted((b for b in books if b["stock"] > 0), key=lambda b: b["id"])
    if not pool:
        return None
    if BOOK_OF_DAY_DAILY:
        return random.Random(date.today().toordinal()).choice(pool)
    return random.choice(pool)


def recommend(books, popularity, mine, liked, limit=4):
    """Подборка: по жанрам/авторам купленного и высоко оценённого; иначе — популярное.
    Возвращает (список книг, персональная ли подборка)."""
    by_id = {b["id"]: b for b in books}
    taste = Counter()
    for bid in set(mine) | set(liked):
        b = by_id.get(bid)
        if b:
            for g in b["genre"].split(","):
                taste[g.strip()] += 1
            for a in b["author"].split(","):
                taste["@" + a.strip()] += 2
    personal = bool(taste)
    scored = []
    for b in books:
        if b["stock"] < 1 or b["id"] in mine:
            continue
        score = random.random() * 0.5  # небольшая случайность: подборка не застывает
        if personal:
            score += sum(taste[g.strip()] for g in b["genre"].split(","))
            score += sum(taste["@" + a.strip()] for a in b["author"].split(","))
            if score < 1:  # ничего общего со вкусом пользователя
                continue
        if b["votes"]:
            score += b["rating"] - 3
        score += 0.2 * popularity.get(b["id"], 0)
        scored.append((score, b))
    scored.sort(key=lambda x: -x[0])
    return [b for _, b in scored[:limit]], personal


@app.get("/")
def catalog():
    uid = session.get("user_id")
    with db() as conn:
        cur = conn.cursor()
        books = fetch_books(cur)
        stats = rating_stats(cur)
        mine_r = my_ratings(cur, uid)
        popularity, bought = purchase_stats(cur, uid)
    for b in books:
        avg, n = stats.get(b["id"], (0.0, 0))
        b["rating"], b["votes"], b["mine"] = round(avg, 1), n, mine_r.get(b["id"], 0)
    genres = sorted({g.strip() for b in books for g in b["genre"].split(",")})
    liked = {i for i, v in mine_r.items() if v >= 4}
    recs, personal = recommend(books, popularity, bought, liked)
    day = book_of_day(books)
    random.shuffle(books)  # при каждом обновлении страницы порядок новый
    resp = app.make_response(render_template(
        "catalog.html", books=books, genres=genres, day=day, recs=recs, personal=personal))
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.post("/rate/<int:id_tov>")
@login_required
def rate(id_tov):
    ajax = request.headers.get("X-Requested-With") == "fetch"
    try:
        score = int(request.form.get("score", ""))
    except ValueError:
        score = 0

    def out(ok, msg, code=200, **kw):
        if ajax:
            return jsonify(ok=ok, message=msg, **kw), code
        flash(msg)
        return redirect(url_for("catalog"))

    if not 1 <= score <= 5:
        return out(False, "Оценка — число от 1 до 5", 400)
    if not HAS_RATING:
        return out(False, "Оценки временно недоступны", 503)
    with db() as conn:
        cur = conn.cursor()
        if not fetch_books(cur, id_tov):
            return out(False, "Такой книги нет", 404)
        cur.execute("SELECT id_rating FROM [Рейтинг] WHERE id_tov = ? AND id_user = ?",
                    id_tov, session["user_id"])
        row = cur.fetchone()
        if row:
            cur.execute("UPDATE [Рейтинг] SET Балл = ?, Дата = ? WHERE id_rating = ?",
                        score, datetime.now(), row[0])
        else:
            cur.execute("INSERT INTO [Рейтинг] (id_tov, id_user, Балл, Дата) VALUES (?, ?, ?, ?)",
                        id_tov, session["user_id"], score, datetime.now())
        avg, n = rating_stats(cur).get(id_tov, (float(score), 1))
    return out(True, "Спасибо за оценку!", avg=round(avg, 1), votes=n, mine=score)


@app.get("/33.html")
def old_index():
    return redirect(url_for("catalog"), 301)


# ====================== Регистрация / вход ======================
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "GET":
        return render_template("register.html", error=None, form={})

    f = {k: (request.form.get(k) or "").strip()
         for k in ("login", "fio", "address", "phone")}
    password = request.form.get("password") or ""
    password2 = request.form.get("password2") or ""
    agree = request.form.get("agree")

    def fail(msg, code=400):
        return render_template("register.html", error=msg, form=dict(f, agree=agree)), code

    if not agree:
        return fail("Необходимо принять пользовательское соглашение и политику конфиденциальности")
    if not f["fio"] or not f["address"] or not f["phone"]:
        return fail("Заполните ФИО, адрес и телефон")
    if len(re.sub(r"\D", "", f["phone"])) < 10:
        return fail("Укажите корректный телефон")
    if not re.fullmatch(r"[\w.\-@]{3,50}", f["login"]):
        return fail("Логин: 3–50 символов (буквы, цифры, . _ - @)")
    if len(password) < 6:
        return fail("Пароль от 6 символов")
    if password != password2:
        return fail("Пароли не совпадают")

    with db() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id_user FROM [Покупатель] WHERE Логин = ?", f["login"])
        if cur.fetchone():
            return fail("Логин уже занят", 409)
        cur.execute(
            "INSERT INTO [Покупатель] (ФИО, [Адресс], Телефон, Логин, Пароль, Роль) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            f["fio"], f["address"], f["phone"], f["login"],
            hash_pw(password), ROLE_USER)
        uid = last_id(cur, "Покупатель", "id_user")
    start_session(uid, f["login"], ROLE_USER)
    return redirect(url_for("catalog"))


@app.route("/login", methods=["GET", "POST"])
def login():
    nxt = request.args.get("next") or request.form.get("next") or ""
    if request.method == "GET":
        return render_template("login.html", error=None, next=nxt)

    login_ = (request.form.get("login") or "").strip()
    password = request.form.get("password") or ""
    key = f"{request.remote_addr}|{login_.lower()}"
    if throttled(key):
        return render_template("login.html", next=nxt,
                               error="Слишком много попыток. Подождите 5 минут"), 429

    with db() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id_user, Пароль, Роль FROM [Покупатель] WHERE Логин = ?", login_)
        row = cur.fetchone()
        good = False
        if row and row[1]:
            if row[1].startswith(HASH_PREFIXES):
                good = check_password_hash(row[1], password)
            elif secrets.compare_digest(row[1], password):  # старая запись без хеша
                good = True
                cur.execute("UPDATE [Покупатель] SET Пароль = ? WHERE id_user = ?",
                            hash_pw(password), row[0])
    if not good:
        FAILS.setdefault(key, []).append(time.time())
        return render_template("login.html", next=nxt,
                               error="Неверный логин или пароль"), 401

    FAILS.pop(key, None)
    start_session(row[0], login_, row[2])
    return redirect(safe_next(nxt))


@app.get("/login.html")
def old_login():
    return redirect(url_for("login"), 301)


@app.get("/register.html")
def old_register():
    return redirect(url_for("register"), 301)


@app.route("/logout", methods=["GET", "POST"])
def logout():
    session.clear()
    return redirect(url_for("catalog"))


# ====================== Заказ ======================
@app.route("/order/<int:id_tov>", methods=["GET", "POST"])
@login_required
def order(id_tov):
    with db() as conn:
        cur = conn.cursor()
        books = fetch_books(cur, id_tov)
        if not books:
            abort(404)
        book = books[0]
        cur.execute("SELECT ФИО, [Адресс], Телефон FROM [Покупатель] WHERE id_user = ?",
                    session["user_id"])
        p = cur.fetchone()
        profile = {"fio": p[0] or "", "address": p[1] or "", "phone": p[2] or ""} if p else {}

        if request.method == "GET":
            return render_template("order.html", book=book, form=profile, error=None, qty=1)

        form = {k: (request.form.get(k) or "").strip() for k in
                ("fio", "phone", "address", "delivery_method", "payment_method", "comment")}
        form["delivery_method"] = form["delivery_method"] or "Курьерская доставка"
        form["payment_method"] = form["payment_method"] or "При получении"
        try:
            qty = int(request.form.get("qty", "1"))
        except ValueError:
            qty = 0

        def fail(msg, code=400):
            return render_template("order.html", book=book, form=form,
                                   error=msg, qty=max(qty, 1)), code

        if not all(form[k] for k in ("fio", "phone", "address")):
            return fail("Заполните ФИО, телефон и адрес")
        if form["delivery_method"] not in ("Курьерская доставка", "Самовывоз"):
            return fail("Выберите способ доставки")
        if form["payment_method"] != "При получении":
            return fail("Выберите способ оплаты")
        if qty < 1 or qty > 20:
            return fail("Количество: от 1 до 20")
        if book["stock"] < qty:
            return fail(f"На складе только {book['stock']} экз.")

        # атомарное списание: не уйдём в минус при одновременных заказах
        cur.execute(
            "UPDATE [Товар] SET Колво = Колво - ? "
            "WHERE id_tov = ? AND Колво >= ? AND " + active_filter(""),
            qty, id_tov, qty)
        if cur.rowcount != 1:
            return fail("Том только что закончился", 409)

        total = book["price"] * qty  # цена берётся из БД, а не из формы
        cur.execute("UPDATE [Покупатель] SET ФИО = ?, [Адресс] = ?, Телефон = ? WHERE id_user = ?",
                    form["fio"], form["address"], form["phone"], session["user_id"])
        id_order, number, delivery_code = insert_order(
            cur, session["user_id"], total, form["address"],
            form["delivery_method"], form["payment_method"], form["comment"])
        try:  # состав заказа (если структура таблицы отличается — заказ всё равно сохранится)
            cur.execute(
                "INSERT INTO [ПозицияЗаказа] (id_order, id_tov, Колво, Стоимость) "
                "VALUES (?, ?, ?, ?)", id_order, id_tov, qty, total)
        except pyodbc.Error as e:
            print(f"[warn] ПозицияЗаказа не записана: {e}")
    flash(f"Заказ {number} оформлен. Код доставки: {delivery_code}")
    return redirect(url_for("orders"))


@app.get("/order.html")
def old_order():
    id_tov = request.args.get("id_tov") or request.args.get("id")
    if id_tov and id_tov.isdigit():
        return redirect(url_for("order", id_tov=int(id_tov)))
    return redirect(url_for("catalog"))


@app.get("/orders")
@login_required
def orders():
    is_adm = session.get("role") == ROLE_ADMIN
    with db() as conn:
        rows = load_orders(conn.cursor(), None if is_adm else session["user_id"])
    return render_template("orders.html", orders=rows, statuses=ORDER_STATUSES)


@app.post("/orders/<int:id_order>/cancel")
@login_required
def order_cancel(id_order):
    is_adm = session.get("role") == ROLE_ADMIN
    try:
        with db() as conn:
            cancel_order(conn.cursor(), id_order, None if is_adm else session["user_id"])
    except CartError as e:
        flash(str(e))
    else:
        flash("Заказ отменён, тома возвращены на склад")
    return redirect(url_for("orders"))


@app.post("/orders/<int:id_order>/address")
@login_required
def order_address(id_order):
    addr = (request.form.get("address") or "").strip()
    if not addr or len(addr) > 255:
        flash("Укажите адрес доставки (до 255 символов)")
        return redirect(url_for("orders"))
    if not HAS_ADDR:
        flash("Изменение адреса в заказе недоступно: в базе нет нужного поля")
        return redirect(url_for("orders"))
    is_adm = session.get("role") == ROLE_ADMIN
    with db() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id_user, Статус FROM [Заказ] WHERE id_order = ?", id_order)
        row = cur.fetchone()
        if not row or (not is_adm and row[0] != session["user_id"]):
            abort(404)
        if row[1] != STATUS_NEW:
            flash("Адрес можно изменить, пока заказ в работе")
            return redirect(url_for("orders"))
        cur.execute("UPDATE [Заказ] SET [Адрес доставки] = ? WHERE id_order = ?", addr, id_order)
    flash("Адрес доставки обновлён")
    return redirect(url_for("orders"))


@app.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    with db() as conn:
        cur = conn.cursor()
        error = None
        if request.method == "POST":
            action = request.form.get("action")
            if action == "info":
                f = {k: (request.form.get(k) or "").strip() for k in ("fio", "phone", "address")}
                if not all(f.values()):
                    error = "Заполните ФИО, телефон и адрес"
                elif len(re.sub(r"\D", "", f["phone"])) < 10:
                    error = "Укажите корректный телефон"
                else:
                    cur.execute("UPDATE [Покупатель] SET ФИО = ?, [Адресс] = ?, Телефон = ? "
                                "WHERE id_user = ?", f["fio"], f["address"], f["phone"],
                                session["user_id"])
                    flash("Данные профиля сохранены")
                    return redirect(url_for("profile"))
            elif action == "password":
                cur.execute("SELECT Пароль FROM [Покупатель] WHERE id_user = ?", session["user_id"])
                row = cur.fetchone()
                old = request.form.get("old") or ""
                new = request.form.get("new") or ""
                if not row or not verify_pw(row[0], old):
                    error = "Текущий пароль неверный"
                elif len(new) < 6:
                    error = "Новый пароль от 6 символов"
                elif new != request.form.get("new2"):
                    error = "Пароли не совпадают"
                else:
                    cur.execute("UPDATE [Покупатель] SET Пароль = ? WHERE id_user = ?",
                                hash_pw(new), session["user_id"])
                    flash("Пароль изменён")
                    return redirect(url_for("profile"))
        cur.execute("SELECT ФИО, [Адресс], Телефон FROM [Покупатель] WHERE id_user = ?",
                    session["user_id"])
        p = cur.fetchone()
    form = {"fio": p[0] or "", "address": p[1] or "", "phone": p[2] or ""} if p else {}
    if request.method == "POST" and request.form.get("action") == "info":
        form = {k: (request.form.get(k) or "").strip() for k in ("fio", "phone", "address")}
    return render_template("profile.html", form=form, error=error), (400 if error else 200)


# ====================== Корзина: страницы ======================
def load_orders(cur, user_id=None):
    """Заказы с составом и адресом. user_id=None — все заказы (для админа)."""
    addr = "z.[Адрес доставки]" if HAS_ADDR else "''"
    sql = ("SELECT z.id_order, z.[Номер заказа], z.[Дата создания], z.Статус, z.Сумма, "
           f"{addr}, p.ФИО, p.Телефон, p.[Адресс], z.[Код доставки], z.[Дата и время], "
           "z.[Способ доставки], z.[Способ оплаты], z.[Статус оплаты], z.[Комментарий] "
           "FROM [Заказ] AS z LEFT JOIN [Покупатель] AS p ON z.id_user = p.id_user")
    params = []
    if user_id is not None:
        sql += " WHERE z.id_user = ?"
        params.append(user_id)
    cur.execute(sql + " ORDER BY z.id_order DESC", *params)
    orders_ = {}
    for r in cur.fetchall():
        orders_[r[0]] = {"id": r[0], "number": r[1], "date": r[2], "status": r[3],
                         "total": r[4], "address": (r[5] or r[8] or "").strip(),
                         "fio": r[6], "phone": r[7], "items": [],
                          "delivery_code": r[9] or "—", "created_at": r[10] or r[2],
                          "delivery_method": r[11] or "Курьерская доставка",
                          "payment_method": r[12] or "При получении",
                          "payment_status": r[13] or "Ожидает оплаты",
                          "comment": r[14] or ""}
    if not orders_:
        return []
    try:
        sql = ("SELECT z.id_order, t.id_tov, t.Название, p.Колво, p.Стоимость "
               "FROM ([Заказ] AS z INNER JOIN [ПозицияЗаказа] AS p ON z.id_order = p.id_order) "
               "LEFT JOIN [Товар] AS t ON p.id_tov = t.id_tov")
        params = []
        if user_id is not None:
            sql += " WHERE z.id_user = ?"
            params.append(user_id)
        cur.execute(sql + " ORDER BY z.id_order DESC", *params)
        for oid, tid, title, qty, cost in cur.fetchall():
            if oid in orders_:
                orders_[oid]["items"].append({
                    "id": tid, "title": (title or "Том удалён").strip(),
                    "qty": qty or 1, "cost": cost or 0})
    except pyodbc.Error as e:
        print(f"[warn] состав заказов недоступен: {e}")
    return list(orders_.values())


def cancel_order(cur, id_order, user_id=None):
    """Отмена заказа с возвратом томов на склад. user_id=None — вызывает админ."""
    cur.execute("SELECT id_user, Статус FROM [Заказ] WHERE id_order = ?", id_order)
    row = cur.fetchone()
    if not row or (user_id is not None and row[0] != user_id):
        abort(404)
    if row[1] == STATUS_CANCELED:
        raise CartError("Заказ уже отменён")
    if user_id is not None and row[1] != STATUS_NEW:
        raise CartError("Выполненный заказ отменить нельзя")
    # статус меняется атомарно: двойной клик не вернёт товары на склад дважды
    cur.execute("UPDATE [Заказ] SET Статус = ? WHERE id_order = ? AND Статус = ?",
                STATUS_CANCELED, id_order, row[1])
    if cur.rowcount != 1:
        raise CartError("Статус заказа уже изменился, обновите страницу")
    n = 0
    try:
        cur.execute("SELECT id_tov, Колво FROM [ПозицияЗаказа] WHERE id_order = ?", id_order)
        for id_tov, qty in cur.fetchall():
            cur.execute("UPDATE [Товар] SET Колво = Колво + ? WHERE id_tov = ?", qty or 0, id_tov)
            n += 1
    except pyodbc.Error as e:
        print(f"[warn] возврат на склад не выполнен: {e}")
    return n


def cart_page(form=None, error=None, code=200):
    with db() as conn:
        cur = conn.cursor()
        cart_ = get_cart()
        items, fixed = [], False
        for id_tov, qty in cart_.items():
            found = fetch_books(cur, id_tov)
            if not found or found[0]["stock"] < 1:
                fixed = True  # том удалён или закончился
                continue
            b = found[0]
            q = min(qty, b["stock"], MAX_QTY)
            fixed = fixed or q != qty
            b["qty"], b["sum"] = q, b["price"] * q
            items.append(b)
        if fixed:
            save_cart({b["id"]: b["qty"] for b in items})
            flash("Корзина обновлена: часть томов закончилась или недоступна")
        if form is None:
            cur.execute("SELECT ФИО, [Адресс], Телефон FROM [Покупатель] WHERE id_user = ?",
                        session["user_id"])
            p = cur.fetchone()
            form = {"fio": p[0] or "", "address": p[1] or "", "phone": p[2] or ""} if p else {}
        history = load_orders(cur, session["user_id"])
    total = sum(b["sum"] for b in items)
    return render_template("cart.html", items=items, total=total, form=form,
                           error=error, history=history), code


@app.get("/cart")
@login_required
def cart():
    return cart_page()


@app.post("/cart/add/<int:id_tov>")
@login_required
def cart_add(id_tov):
    try:
        qty = max(1, min(int(request.form.get("qty", "1")), MAX_QTY))
    except ValueError:
        qty = 1
    with db() as conn:
        found = fetch_books(conn.cursor(), id_tov)
    ok = False
    if not found or found[0]["stock"] < 1:
        msg = "Этого тома нет в наличии"
    else:
        b, cart_ = found[0], get_cart()
        have = cart_.get(id_tov, 0)
        new = min(have + qty, b["stock"], MAX_QTY)
        if new == have:
            msg = f"«{b['title']}»: больше добавить нельзя (в корзине максимум)"
        else:
            cart_[id_tov] = new
            save_cart(cart_)
            ok, msg = True, f"«{b['title']}» добавлен в корзину"
    if request.headers.get("X-Requested-With") == "fetch":
        return jsonify(ok=ok, message=msg, count=sum(get_cart().values()))
    flash(msg)
    return redirect(safe_next(request.form.get("next")))


@app.post("/cart/update/<int:id_tov>")
@login_required
def cart_update(id_tov):
    cart_ = get_cart()
    try:
        qty = int(request.form.get("qty", "1"))
    except ValueError:
        qty = cart_.get(id_tov, 1)
    if id_tov in cart_:
        if qty < 1:
            del cart_[id_tov]
        else:
            cart_[id_tov] = min(qty, MAX_QTY)
        save_cart(cart_)
    return redirect(url_for("cart"))


@app.post("/cart/remove/<int:id_tov>")
@login_required
def cart_remove(id_tov):
    cart_ = get_cart()
    if cart_.pop(id_tov, None) is not None:
        save_cart(cart_)
    return redirect(url_for("cart"))


@app.post("/cart/checkout")
@login_required
def cart_checkout():
    form = {k: (request.form.get(k) or "").strip() for k in
            ("fio", "phone", "address", "delivery_method", "payment_method", "comment")}
    form["delivery_method"] = form["delivery_method"] or "Курьерская доставка"
    form["payment_method"] = form["payment_method"] or "При получении"
    cart_ = get_cart()
    if not cart_:
        flash("Корзина пуста")
        return redirect(url_for("cart"))
    if not all(form[k] for k in ("fio", "phone", "address")):
        return cart_page(form, "Заполните ФИО, телефон и адрес", 400)
    if form["delivery_method"] not in ("Курьерская доставка", "Самовывоз"):
        return cart_page(form, "Выберите способ доставки", 400)
    if form["payment_method"] != "При получении":
        return cart_page(form, "Выберите способ оплаты", 400)

    try:
        with db() as conn:  # любая CartError откатывает всё: заказ либо целиком, либо никак
            cur = conn.cursor()
            lines, total = [], 0.0
            for id_tov, qty in cart_.items():
                found = fetch_books(cur, id_tov)
                if not found:
                    raise CartError("Один из томов больше недоступен")
                b = found[0]
                qty = min(qty, MAX_QTY)
                if b["stock"] < qty:
                    raise CartError(f"«{b['title']}»: на складе только {b['stock']} экз.")
                lines.append((id_tov, qty, b["price"] * qty))  # цена из БД, не из формы
                total += b["price"] * qty

            cur.execute("UPDATE [Покупатель] SET ФИО = ?, [Адресс] = ?, Телефон = ? WHERE id_user = ?",
                        form["fio"], form["address"], form["phone"], session["user_id"])
            id_order, number, delivery_code = insert_order(
                cur, session["user_id"], total, form["address"],
                form["delivery_method"], form["payment_method"], form["comment"])
            for id_tov, qty, cost in lines:
                cur.execute(
                    "UPDATE [Товар] SET Колво = Колво - ? WHERE id_tov = ? AND Колво >= ? AND "
                    + active_filter(""), qty, id_tov, qty)
                if cur.rowcount != 1:
                    raise CartError("Один из томов только что закончился")
                try:
                    cur.execute(
                        "INSERT INTO [ПозицияЗаказа] (id_order, id_tov, Колво, Стоимость) "
                        "VALUES (?, ?, ?, ?)", id_order, id_tov, qty, cost)
                except pyodbc.Error as e:
                    print(f"[warn] ПозицияЗаказа не записана: {e}")
    except CartError as e:
        return cart_page(form, str(e), 409)
    save_cart({})
    flash(f"Заказ {number} оформлен. Код доставки: {delivery_code}")
    return redirect(url_for("orders"))


# ====================== Админ ======================
def _refs(cur):
    cur.execute("SELECT id_janr, Название FROM [Жанр] ORDER BY Название")
    genres = [(r[0], r[1]) for r in cur.fetchall()]
    cur.execute("SELECT id_izd, Название FROM [Издательство] ORDER BY Название")
    pubs = [(r[0], r[1]) for r in cur.fetchall()]
    return genres, pubs


@app.get("/admin")
@admin_required
def admin_dashboard():
    term = (request.args.get("q") or "").strip().lower()
    genre_filter = (request.args.get("genre") or "").strip()
    stock_filter = (request.args.get("stock") or "all").strip()
    sort = (request.args.get("sort") or "newest").strip()
    with db() as conn:
        cur = conn.cursor()
        books = fetch_books(cur)
        genres, pubs = _refs(cur)

    total_books = len(books)
    total_available = sum(1 for b in books if b["stock"] > 0)
    total_empty = sum(1 for b in books if b["stock"] == 0)
    if term:
        books = [b for b in books if term in " ".join(
            [str(b["id"]), b["title"], b["author"], b["isbn"], b["genre"], b["publisher"]]
        ).lower()]
    if genre_filter:
        books = [b for b in books if genre_filter in [x.strip() for x in b["genre"].split(",")]]
    if stock_filter == "available":
        books = [b for b in books if b["stock"] > 0]
    elif stock_filter == "empty":
        books = [b for b in books if b["stock"] == 0]
    elif stock_filter == "low":
        books = [b for b in books if 0 < b["stock"] <= 5]

    sorters = {
        "newest": lambda b: b["id"],
        "title": lambda b: str(b["title"]).lower(),
        "price_asc": lambda b: b["price"],
        "price_desc": lambda b: b["price"],
        "stock_asc": lambda b: b["stock"],
        "stock_desc": lambda b: b["stock"],
    }
    if sort in sorters:
        books.sort(key=sorters[sort], reverse=sort in ("newest", "price_desc", "stock_desc"))
    else:
        sort = "newest"
    stats = {
        "total": total_books,
        "shown": len(books),
        "available": total_available,
        "empty": total_empty,
    }
    return render_template("admin.html", books=books, genres=genres, pubs=pubs,
                           term=term, genre_filter=genre_filter, stock_filter=stock_filter,
                           sort=sort, stats=stats)


@app.post("/admin/update/<int:id_tov>")
@admin_required
def admin_update_book(id_tov):
    f = {k: (request.form.get(k) or "").strip() for k in ("title", "author", "isbn", "year", "price", "stock", "genre_id", "publisher_id")}
    try:
        year, price, stock = int(f["year"]), float(f["price"].replace(",", ".")), int(f["stock"])
        genre_id, pub_id = int(f["genre_id"]), int(f["publisher_id"])
        if not f["title"] or not f["author"] or not 1000 <= year <= datetime.now().year + 1 or price <= 0 or stock < 0:
            raise ValueError
    except ValueError:
        flash("Проверьте название, автора, год, цену и остаток.")
        return redirect(url_for("admin_dashboard"))
    cover_ext = None
    cover_data = b""
    cover = request.files.get("cover")
    if cover and cover.filename:
        cover_data = cover.read()
        cover_ext = sniff_image(cover_data[:12])
        if not cover_ext:
            flash("Обложка должна быть JPG, PNG или WEBP.")
            return redirect(url_for("admin_dashboard", q=request.form.get("return_q", "")))
    with db() as conn:
        cur = conn.cursor()
        genres, pubs = _refs(cur)
        if genre_id not in {x[0] for x in genres} or pub_id not in {x[0] for x in pubs}:
            flash("Выберите существующие жанр и издательство.")
            return redirect(url_for("admin_dashboard"))
        cur.execute("SELECT id_aut FROM [Автор] WHERE ФИО = ?", f["author"])
        author_row = cur.fetchone()
        if author_row: author_id = author_row[0]
        else:
            cur.execute("INSERT INTO [Автор] (ФИО) VALUES (?)", f["author"])
            author_id = last_id(cur, "Автор", "id_aut")
        cur.execute("UPDATE [Товар] SET Название = ?, ISBN = ?, [Год издания] = ?, Цена = ?, Колво = ?, id_izd = ? WHERE id_tov = ?", f["title"], f["isbn"] or None, year, price, stock, pub_id, id_tov)
        cur.execute("DELETE FROM [АвторТовар] WHERE id_tov = ?", id_tov)
        cur.execute("INSERT INTO [АвторТовар] (id_aut, id_tov) VALUES (?, ?)", author_id, id_tov)
        cur.execute("DELETE FROM [ЖанрТовар] WHERE id_tov = ?", id_tov)
        cur.execute("INSERT INTO [ЖанрТовар] (id_janr, id_tov) VALUES (?, ?)", genre_id, id_tov)
    if cover_ext:
        os.makedirs(IMG_DIR, exist_ok=True)
        for old_ext in IMAGE_EXT:
            old_path = os.path.join(IMG_DIR, f"{id_tov}.{old_ext}")
            if os.path.exists(old_path): os.remove(old_path)
        with open(os.path.join(IMG_DIR, f"{id_tov}.{cover_ext}"), "wb") as fh:
            fh.write(cover_data)
    flash(f"Издание «{f['title']}» сохранено.")
    return redirect(url_for("admin_dashboard", q=request.form.get("return_q", "")))


@app.route("/admin/add", methods=["GET", "POST"])
@admin_required
def admin_add():
    with db() as conn:
        cur = conn.cursor()
        genres, pubs = _refs(cur)
        if request.method == "GET":
            return render_template("admin_add.html", genres=genres, pubs=pubs, form={})

        f = {k: (request.form.get(k) or "").strip() for k in
             ("title", "author", "isbn", "year", "price", "stock", "genre_id", "publisher_id")}

        def fail(msg):
            return render_template("admin_add.html", genres=genres, pubs=pubs,
                                   form=f, error=msg), 400

        try:
            year = int(f["year"])
            price = float(f["price"].replace(",", "."))
            stock = int(f["stock"])
            genre_id = int(f["genre_id"])
            pub_id = int(f["publisher_id"])
        except ValueError:
            return fail("Год, цена, количество, жанр и издательство — числа")
        if not f["title"] or not f["author"]:
            return fail("Укажите название и автора")
        if not 1000 <= year <= datetime.now().year + 1:
            return fail("Некорректный год издания")
        if price <= 0 or stock < 0:
            return fail("Цена должна быть больше 0, количество — не меньше 0")
        if genre_id not in {g[0] for g in genres} or pub_id not in {p[0] for p in pubs}:
            return fail("Выберите жанр и издательство из списка")

        ext = None
        cover = request.files.get("cover")
        data = b""
        if cover and cover.filename:
            data = cover.read()
            ext = sniff_image(data[:12])
            if not ext:
                return fail("Обложка: допустимы JPG, PNG, WEBP")

        cur.execute("SELECT id_aut FROM [Автор] WHERE ФИО = ?", f["author"])
        row = cur.fetchone()
        if row:
            id_aut = row[0]
        else:
            cur.execute("INSERT INTO [Автор] (ФИО) VALUES (?)", f["author"])
            id_aut = last_id(cur, "Автор", "id_aut")

        cur.execute(
            "INSERT INTO [Товар] (ISBN, Название, Цена, [Год издания], Колво, id_izd"
            + (", [Активен]) VALUES (?, ?, ?, ?, ?, ?, True)" if HAS_ACTIVE
               else ") VALUES (?, ?, ?, ?, ?, ?)"),
            f["isbn"] or None, f["title"], price, year, stock, pub_id)
        id_tov = last_id(cur, "Товар", "id_tov")
        cur.execute("INSERT INTO [АвторТовар] (id_aut, id_tov) VALUES (?, ?)", id_aut, id_tov)
        cur.execute("INSERT INTO [ЖанрТовар] (id_janr, id_tov) VALUES (?, ?)", genre_id, id_tov)

    if ext:
        os.makedirs(IMG_DIR, exist_ok=True)
        with open(os.path.join(IMG_DIR, f"{id_tov}.{ext}"), "wb") as fh:
            fh.write(data)
    flash(f"Том «{f['title']}» добавлен в каталог")
    return redirect(url_for("admin_dashboard"))


@app.post("/admin/edit/<int:id_tov>")
@admin_required
def admin_edit(id_tov):
    try:
        price = float((request.form.get("price") or "").replace(",", "."))
        stock = int(request.form.get("stock") or "")
        if price <= 0 or stock < 0:
            raise ValueError
    except ValueError:
        flash("Неверные значения цены или остатка")
        return redirect(url_for("admin_dashboard"))
    with db() as conn:
        conn.cursor().execute("UPDATE [Товар] SET Цена = ?, Колво = ? WHERE id_tov = ?",
                              price, stock, id_tov)
    flash("Данные тома обновлены")
    return redirect(url_for("admin_dashboard"))


@app.post("/admin/delete/<int:id_tov>")
@admin_required
def admin_delete(id_tov):
    # мягкое удаление: том исчезает из каталога, история заказов остаётся целой
    with db() as conn:
        cur = conn.cursor()
        if HAS_ACTIVE:
            cur.execute("UPDATE [Товар] SET [Активен] = False WHERE id_tov = ?", id_tov)
        else:
            try:
                cur.execute("DELETE FROM [АвторТовар] WHERE id_tov = ?", id_tov)
                cur.execute("DELETE FROM [ЖанрТовар] WHERE id_tov = ?", id_tov)
                cur.execute("DELETE FROM [Товар] WHERE id_tov = ?", id_tov)
            except pyodbc.Error:
                conn.rollback()
                flash("Нельзя удалить: том есть в заказах")
                return redirect(url_for("admin_dashboard"))
    flash("Том удалён из каталога")
    return redirect(url_for("admin_dashboard"))


@app.post("/admin/order/<int:id_order>/status")
@admin_required
def admin_order_status(id_order):
    status = request.form.get("status")
    if status not in ORDER_STATUSES:
        abort(400)
    try:
        with db() as conn:
            cur = conn.cursor()
            if status == STATUS_CANCELED:
                cancel_order(cur, id_order)  # вернёт тома на склад
            else:
                cur.execute("UPDATE [Заказ] SET Статус = ?, [Дата и время] = Now(), "
                             "[Статус оплаты] = ? WHERE id_order = ? AND Статус <> ?",
                            status, ("Оплачено" if status == STATUS_DONE else "Ожидает оплаты"),
                            id_order, STATUS_CANCELED)
    except CartError as e:
        flash(str(e))
    return redirect(url_for("orders"))


# ====================== Информационные страницы ======================
@app.get("/terms")
def terms():
    return render_template("terms.html")


@app.get("/privacy")
def privacy():
    return render_template("privacy.html")


@app.get("/contacts")
def contacts():
    return render_template("contacts.html")


@app.get("/delivery")
def delivery():
    return render_template("delivery.html")


# ====================== Ошибки ======================
def _error(code, title, text):
    return render_template("error.html", code=code, title=title, text=text), code


@app.errorhandler(400)
def e400(_):
    return _error(400, "Запрос отклонён", "Страница устарела. Обновите её и повторите действие.")


@app.errorhandler(403)
def e403(_):
    return _error(403, "Доступ запрещён", "Нужны права администратора.")


@app.errorhandler(404)
def e404(_):
    return _error(404, "Не найдено", "Такой страницы или тома нет.")


@app.errorhandler(pyodbc.Error)
def edb(e):
    print(f"[db error] {e}")
    return _error(503, "База данных недоступна", "Попробуйте позже.")


_init_lock = threading.Lock()
_inited = False


def init_once():
    """Подготовка БД при первом запросе (работает и при запуске не через python server.py)."""
    global _inited
    with _init_lock:
        if _inited:
            return
        _inited = True
        try:
            ensure_schema()
        except pyodbc.Error as e:
            print(f"[warn] ensure_schema: {e}")
        detect_schema()
        os.makedirs(IMG_DIR, exist_ok=True)


@app.before_request
def _boot():
    if not _inited:
        init_once()


if __name__ == "__main__":
    init_once()
    # debug выключен: отладчик Werkzeug позволяет выполнить код на сервере
    app.run(debug=False, host=os.environ.get("HOST", "0.0.0.0"),
            port=int(os.environ.get("PORT", "3000")))
