# models.py — Banco de dados SQLite para LH TECH
import sqlite3
import os
import json
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash

DB_PATH = os.path.join(os.path.dirname(__file__), "database.db")

# Empresas/módulos disponíveis
COMPANIES = {
    "elleve": {
        "name": "Colégio Elleve",
        "icon": "school",
        "color": "#4BB6E5",
        "description": "Gerador de relatórios de ponto — Colégio Elleve LTDA",
        "active": True,
    },
    "aquarela": {
        "name": "Aquarela Kids",
        "icon": "child_care",
        "color": "#f59f00",
        "description": "Gerador de relatórios de ponto — Aquarela Kids",
        "active": True,
    },
    "mrcontabil": {
        "name": "MR Contábil",
        "icon": "account_balance",
        "color": "#51cf66",
        "description": "Gerador de relatórios de ponto — MR Organização Contábil LTDA",
        "active": True,
    },
}


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Cria tabelas e admin padrão se não existirem."""
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            full_name TEXT NOT NULL DEFAULT '',
            role TEXT NOT NULL DEFAULT 'user',
            companies TEXT NOT NULL DEFAULT '[]',
            tools TEXT NOT NULL DEFAULT '[]',
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL DEFAULT (datetime('now')),
            last_login TEXT,
            last_ip TEXT
        )
    """)
    conn.commit()

    # Migração simples: garante que colunas novas existam caso o BD já tenha sido criado
    for col_def in [
        "ALTER TABLE users ADD COLUMN last_login TEXT",
        "ALTER TABLE users ADD COLUMN last_ip TEXT",
        "ALTER TABLE users ADD COLUMN tools TEXT NOT NULL DEFAULT '[]'",
    ]:
        try:
            conn.execute(col_def)
            conn.commit()
        except sqlite3.OperationalError:
            pass  # Coluna já existe

    # Cria admin padrão se não existir
    admin = conn.execute("SELECT id FROM users WHERE username = ?", ("admin",)).fetchone()
    if not admin:
        create_user(conn, "admin", "admin123", "Administrador", "admin",
                     list(COMPANIES.keys()))
    conn.close()


# ─── CRUD ────────────────────────────────────────────


def create_user(conn, username, password, full_name, role="user", companies=None, tools=None):
    pw_hash = generate_password_hash(password)
    companies_json = json.dumps(companies or [])
    tools_json = json.dumps(tools or [])
    conn.execute(
        "INSERT INTO users (username, password_hash, full_name, role, companies, tools) VALUES (?, ?, ?, ?, ?, ?)",
        (username, pw_hash, full_name, role, companies_json, tools_json)
    )
    conn.commit()


def get_user_by_id(user_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    conn.close()
    if row:
        return _row_to_dict(row)
    return None


def get_user_by_username(username):
    conn = get_db()
    row = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
    conn.close()
    if row:
        return _row_to_dict(row)
    return None


def verify_password(user_dict, password):
    return check_password_hash(user_dict["password_hash"], password)


def list_users():
    conn = get_db()
    rows = conn.execute("SELECT * FROM users ORDER BY username").fetchall()
    conn.close()
    return [_row_to_dict(r) for r in rows]


def update_user(user_id, full_name=None, role=None, companies=None, tools=None, password=None, active=None):
    conn = get_db()
    fields = []
    values = []

    if full_name is not None:
        fields.append("full_name = ?")
        values.append(full_name)
    if role is not None:
        fields.append("role = ?")
        values.append(role)
    if companies is not None:
        fields.append("companies = ?")
        values.append(json.dumps(companies))
    if tools is not None:
        fields.append("tools = ?")
        values.append(json.dumps(tools))
    if password is not None:
        fields.append("password_hash = ?")
        values.append(generate_password_hash(password))
    if active is not None:
        fields.append("active = ?")
        values.append(1 if active else 0)

    if fields:
        values.append(user_id)
        conn.execute(f"UPDATE users SET {', '.join(fields)} WHERE id = ?", values)
        conn.commit()
    conn.close()


def delete_user(user_id):
    conn = get_db()
    conn.execute("DELETE FROM users WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()


def update_last_login(user_id, ip):
    conn = get_db()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        "UPDATE users SET last_login = ?, last_ip = ? WHERE id = ?",
        (now, ip, user_id)
    )
    conn.commit()
    conn.close()


def _row_to_dict(row):
    d = dict(row)
    try:
        d["companies"] = json.loads(d.get("companies", "[]"))
    except Exception:
        d["companies"] = []
    try:
        d["tools"] = json.loads(d.get("tools", "[]"))
    except Exception:
        d["tools"] = []
    return d
