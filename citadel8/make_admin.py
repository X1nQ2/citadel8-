"""Выдать / снять роль администратора.

  python make_admin.py ЛОГИН          -> роль «Админ»
  python make_admin.py ЛОГИН --user   -> роль «Покупатель»
"""
import sys

import pyodbc

from server import CONN_STR, ROLE_ADMIN, ROLE_USER

if len(sys.argv) < 2:
    sys.exit(__doc__)

login = sys.argv[1]
role = ROLE_USER if "--user" in sys.argv else ROLE_ADMIN
with pyodbc.connect(CONN_STR, autocommit=True) as conn:
    n = conn.execute("UPDATE [Покупатель] SET Роль = ? WHERE Логин = ?", role, login).rowcount
print(f"Готово: {login} → {role}" if n else "Пользователь с таким логином не найден.")
