import os
import sys
from flask import Flask, render_template, request, jsonify, redirect, session
import pyodbc
from datetime import datetime

# Определение базовой папки
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

app = Flask(__name__,
            template_folder=os.path.join(BASE_DIR, 'templates'),
            static_folder=BASE_DIR,
            static_url_path='')

app.secret_key = 'change-me-to-random-string'

DB_PATH = os.path.join(BASE_DIR, "database.accdb")
CONN_STR = f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};Dbq={DB_PATH};"


def get_db_connection():
    try:
        conn = pyodbc.connect(CONN_STR)
        return conn
    except Exception as e:
        print(f"Ошибка подключения к СУБД Access: {e}")
        return None


def ensure_auth_columns():
    """Добавляет Логин, Пароль и Роль в [Покупатель], если их нет в структуре таблицы."""
    conn = get_db_connection()
    if not conn:
        return
    try:
        cur = conn.cursor()
        try:
            cur.execute("SELECT Логин, Пароль, Роль FROM [Покупатель] WHERE 1=0")
        except Exception:
            conn.rollback()
            try:
                cur.execute("ALTER TABLE [Покупатель] ADD COLUMN [Логин] TEXT(50)")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE [Покупатель] ADD COLUMN [Пароль] TEXT(255)")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE [Покупатель] ADD COLUMN [Роль] TEXT(50)")
            except Exception:
                pass
            conn.commit()
    except Exception as e:
        print(f"Ошибка миграции [Покупатель]: {e}")
    finally:
        conn.close()


@app.after_request
def no_cache(resp):
    if request.path in ('/login.html', '/register.html'):
        resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        resp.headers['Pragma'] = 'no-cache'
        resp.headers['Expires'] = '0'
    return resp


@app.route('/register.html', methods=['GET', 'POST'])
def register():
    if request.method == 'GET':
        return render_template('register.html', error=None)

    login = (request.form.get('login') or '').strip()
    password = request.form.get('password') or ''
    password2 = request.form.get('password2') or ''
    fio = (request.form.get('fio') or '').strip()
    address = (request.form.get('address') or '').strip()
    phone = (request.form.get('phone') or '').strip()

    if not fio or not address or not phone:
        return render_template('register.html', error='Заполните ФИО, адрес и телефон'), 400
    if len(login) < 3 or len(password) < 6:
        return render_template('register.html', error='Логин от 3 символов, пароль от 6'), 400
    if password != password2:
        return render_template('register.html', error='Пароли не совпадают'), 400

    conn = get_db_connection()
    if not conn:
        return render_template('register.html', error='База данных недоступна'), 500
    try:
        cur = conn.cursor()
        cur.execute("SELECT id_user FROM [Покупатель] WHERE Логин = ?", (login,))
        if cur.fetchone():
            return render_template('register.html', error='Логин уже занят'), 409

        cur.execute(
            "INSERT INTO [Покупатель] (ФИО, [Адресс], Телефон, Логин, Пароль, Роль) VALUES (?, ?, ?, ?, ?, ?)",
            (fio, address, phone, login, password, 'Покупатель'))
        conn.commit()

        cur.execute("SELECT id_user, Роль FROM [Покупатель] WHERE Логин = ?", (login,))
        row = cur.fetchone()
        session['user_id'] = row[0]
        session['login'] = login
        session['role'] = row[1] if row[1] else 'Покупатель'
        return redirect('/')
    except Exception as e:
        print(f"Ошибка регистрации: {e}")
        return render_template('register.html', error='Ошибка сохранения в БД'), 500
    finally:
        conn.close()


@app.route('/login.html', methods=['GET', 'POST'])
def login_page():
    if request.method == 'GET':
        return render_template('login.html', error=None)

    login = (request.form.get('login') or '').strip()
    password = request.form.get('password') or ''
    conn = get_db_connection()
    if not conn:
        return render_template('login.html', error='База данных недоступна'), 500
    try:
        cur = conn.cursor()
        cur.execute("SELECT id_user, Пароль, Роль FROM [Покупатель] WHERE Логин = ?", (login,))
        row = cur.fetchone()
        if not row or row[1] != password:
            return render_template('login.html', error='Неверный логин или пароль'), 401

        session['user_id'] = row[0]
        session['login'] = login
        session['role'] = row[2] if row[2] else 'Покупатель'
        return redirect('/')
    finally:
        conn.close()


@app.route('/logout')
def logout():
    session.clear()
    return redirect('/')


@app.route('/api/me')
def me():
    return jsonify({
        "login": session.get('login'),
        "role": session.get('role')
    })


@app.route('/')
@app.route('/33.html')
def index_page():
    return render_template('33.html')


@app.route('/order.html')
def order_page():
    return render_template('order.html')


@app.route('/api/books', methods=['GET'])
def get_books():
    conn = get_db_connection()
    if not conn:
        return jsonify({"error": "База данных временно недоступна"}), 500

    cursor = conn.cursor()
    sql = """
        SELECT 
            t.id_tov, 
            t.Название, 
            t.ISBN, 
            t.Цена, 
            t.Колво,
            a.ФИО,
            j.Название,
            i.Название,
            t.[Год издания]
        FROM (((([Товар] AS t
        LEFT JOIN [АвторТовар] AS at ON t.id_tov = at.id_tov)
        LEFT JOIN [Автор] AS a ON at.id_aut = a.id_aut)
        LEFT JOIN [ЖанрТовар] AS jt ON t.id_tov = jt.id_tov)
        LEFT JOIN [Жанр] AS j ON jt.id_janr = j.id_janr)
        LEFT JOIN [Издательство] AS i ON t.id_izd = i.id_izd
    """
    try:
        cursor.execute(sql)
        rows = cursor.fetchall()

        books_dict = {}
        for row in rows:
            book_id = row[0]
            author_val = row[5].strip() if row[5] else "Не указан"
            genre_val = row[6].strip() if row[6] else "Классика"

            if book_id not in books_dict:
                books_dict[book_id] = {
                    "id": book_id,
                    "title": row[1],
                    "isbn": row[2],
                    "price": float(row[3]) if row[3] is not None else 0.0,
                    "stock": row[4],
                    "authors_set": {author_val},
                    "genres_set": {genre_val},
                    "publisher": row[7] if row[7] else "Цитадель",
                    "year": row[8] if row[8] else "—"
                }
            else:
                books_dict[book_id]["authors_set"].add(author_val)
                books_dict[book_id]["genres_set"].add(genre_val)

        final_books = []
        for b_id, b_data in books_dict.items():
            b_data["author"] = ", ".join(sorted(list(b_data["authors_set"]))).strip()
            b_data["genre"] = ", ".join(sorted(list(b_data["genres_set"]))).strip()
            del b_data["authors_set"]
            del b_data["genres_set"]
            final_books.append(b_data)

        return jsonify(final_books)
    except Exception as e:
        print(f"Ошибка выполнения SQL-запроса: {e}")
        return jsonify({"error": "Ошибка при чтении таблиц"}), 500
    finally:
        conn.close()


@app.route('/api/books/add', methods=['POST'])
def add_book():
    if session.get('role') != 'Админ':
        return jsonify({"error": "Доступ запрещен. Требуются права администратора"}), 403

    data = request.form if request.form else request.get_json()

    title = data.get('title')
    isbn = data.get('isbn')
    price = data.get('price')
    stock = data.get('stock')
    year = data.get('year')
    author_name = data.get('author')
    genre_id = data.get('genre_id')
    publisher_id = data.get('publisher_id')

    if not title or not price or not stock or not author_name or not genre_id or not publisher_id or not year:
        return jsonify({"error": "Заполните все обязательные поля!"}), 400

    conn = get_db_connection()
    if not conn:
        return jsonify({"error": "База данных недоступна"}), 500

    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id_aut FROM [Автор] WHERE ФИО = ?", (author_name,))
        author_row = cursor.fetchone()

        if author_row:
            id_aut = author_row[0]
        else:
            cursor.execute("INSERT INTO [Автор] (ФИО) VALUES (?)", (author_name,))
            conn.commit()
            cursor.execute("SELECT MAX(id_aut) FROM [Автор]")
            id_aut = cursor.fetchone()[0]

        sql_insert_goods = """
            INSERT INTO [Товар] (ISBN, Название, Цена, [Год издания], Колво, id_izd) 
            VALUES (?, ?, ?, ?, ?, ?)
        """
        cursor.execute(sql_insert_goods, (isbn, title, float(price), int(year), int(stock), int(publisher_id)))
        conn.commit()

        cursor.execute("SELECT MAX(id_tov) FROM [Товар]")
        id_tov = cursor.fetchone()[0]

        cursor.execute("INSERT INTO [АвторТовар] (id_aut, id_tov) VALUES (?, ?)", (int(id_aut), int(id_tov)))
        cursor.execute("INSERT INTO [ЖанрТовар] (id_janr, id_tov) VALUES (?, ?)", (int(genre_id), int(id_tov)))

        conn.commit()
        return redirect('/')
    except Exception as e:
        print(f"Ошибка сохранения книги в Access: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()


@app.route('/api/checkout', methods=['POST'])
def checkout():
    fio = request.form.get('fio')
    phone = request.form.get('phone')
    address = request.form.get('address')
    id_tov = request.form.get('id_tov')
    price = request.form.get('price')

    if not fio or not phone or not address or not id_tov or not price:
        return '<h3>Ошибка: Заполните все поля!</h3><a href="/">В каталог</a>', 400

    conn = get_db_connection()
    if not conn:
        return '<h3>Ошибка: База данных недоступна</h3>', 500

    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Колво FROM [Товар] WHERE id_tov = ?", (int(id_tov),))
        book_row = cursor.fetchone()

        if not book_row:
            return '<h3>Ошибка: Товар не найден в базе данных!</h3><a href="/">В каталог</a>', 404

        current_stock = book_row[0]
        if current_stock <= 0:
            return '<h3>Ошибка: Извините, данный том закончился на складе!</h3><a href="/">В каталог</a>', 400

        cursor.execute("UPDATE [Товар] SET Колво = Колво - 1 WHERE id_tov = ?", (int(id_tov),))

        id_user = session.get('user_id')
        if id_user:
            cursor.execute("UPDATE [Покупатель] SET ФИО = ?, [Адресс] = ?, Телефон = ? WHERE id_user = ?",
                           (fio, address, phone, int(id_user)))
        else:
            cursor.execute("INSERT INTO [Покупатель] (ФИО, [Адресс], Телефон, Роль) VALUES (?, ?, ?, ?)",
                           (fio, address, phone, 'Покупатель'))
            conn.commit()
            cursor.execute("SELECT MAX(id_user) FROM [Покупатель]")
            id_user = cursor.fetchone()[0]
        conn.commit()

        current_date = datetime.now().strftime('%d.%m.%Y')
        cursor.execute("SELECT COUNT(id_order) FROM [Заказ]")
        order_count = cursor.fetchone()[0] + 1
        order_num = f"ORD-{str(order_count).zfill(3)}"

        sql_order = """
            INSERT INTO [Заказ] (id_user, [Номер заказа], [Дата создания], Статус, Сумма) 
            VALUES (?, ?, ?, ?, ?)
        """
        cursor.execute(sql_order, (int(id_user), order_num, current_date, 'В работе', float(price)))
        conn.commit()

        return redirect('/?payment=success')
    except Exception as e:
        conn.rollback()
        print(f"Критическая ошибка СУБД Access: {e}")
        return f"""
            <div style="font-family: sans-serif; padding: 30px; color: #721c24; background: #f8d7da; border: 1px solid #f5c6cb;">
                <h3>Ошибка базы данных Access при сохранении:</h3>
                <p><code>{str(e)}</code></p>
                <a href="/">Вернуться назад</a>
            </div>
        """, 500
    finally:
        conn.close()


@app.route('/api/books/delete/<int:book_id>', methods=['POST'])
def delete_book(book_id):
    if session.get('role') != 'Админ':
        return jsonify({"error": "Доступ запрещен. Требуются права администратора"}), 403

    conn = get_db_connection()
    if not conn:
        return jsonify({"error": "База данных временно недоступна"}), 500

    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM [АвторТовар] WHERE id_tov = ?", (book_id,))
        cursor.execute("DELETE FROM [ЖанрТовар] WHERE id_tov = ?", (book_id,))
        cursor.execute("DELETE FROM [Товар] WHERE id_tov = ?", (book_id,))
        conn.commit()
        return jsonify({"success": True, "message": "Издание удалено из архива"})
    except Exception as e:
        print(f"Ошибка при удалении книги: {e}")
        return jsonify({"error": "Не удалось удалить книгу"}), 500
    finally:
        conn.close()


if __name__ == '__main__':
    ensure_auth_columns()
    app.run(debug=True, host='0.0.0.0', port=5000)
