# admin.py — Blueprint do painel administrativo
from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required, current_user
from functools import wraps
from models import list_users, create_user, get_user_by_id, update_user, delete_user, get_db, COMPANIES

admin_bp = Blueprint("admin", __name__, url_prefix="/admin")


def admin_required(f):
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if not current_user.is_admin:
            flash("Acesso restrito a administradores.", "error")
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)
    return decorated


# ─── Listar usuários ─────────────────────────────────


@admin_bp.route("/users")
@admin_required
def users_list():
    users = list_users()
    return render_template("admin_users.html", users=users, companies=COMPANIES)


# ─── Criar usuário ───────────────────────────────────


@admin_bp.route("/users/new", methods=["GET", "POST"])
@admin_required
def user_new():
    if request.method == "POST":
        username = request.form.get("username", "").strip().lower()
        password = request.form.get("password", "")
        full_name = request.form.get("full_name", "").strip()
        role = request.form.get("role", "user")
        companies = request.form.getlist("companies")

        if not username or not password:
            flash("Usuário e senha são obrigatórios.", "error")
            return render_template("admin_user_form.html", mode="new",
                                   user={}, companies=COMPANIES)

        from models import get_user_by_username
        if get_user_by_username(username):
            flash("Esse nome de usuário já existe.", "error")
            return render_template("admin_user_form.html", mode="new",
                                   user={}, companies=COMPANIES)

        conn = get_db()
        create_user(conn, username, password, full_name, role, companies)
        conn.close()
        flash(f"Usuário '{username}' criado com sucesso!", "success")
        return redirect(url_for("admin.users_list"))

    return render_template("admin_user_form.html", mode="new",
                           user={}, companies=COMPANIES)


# ─── Editar usuário ──────────────────────────────────


@admin_bp.route("/users/<int:user_id>/edit", methods=["GET", "POST"])
@admin_required
def user_edit(user_id):
    user = get_user_by_id(user_id)
    if not user:
        flash("Usuário não encontrado.", "error")
        return redirect(url_for("admin.users_list"))

    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        role = request.form.get("role", "user")
        companies = request.form.getlist("companies")
        password = request.form.get("password", "").strip()
        active = request.form.get("active") == "on"

        kwargs = {
            "full_name": full_name,
            "role": role,
            "companies": companies,
            "active": active,
        }
        if password:
            kwargs["password"] = password

        update_user(user_id, **kwargs)
        flash(f"Usuário '{user['username']}' atualizado!", "success")
        return redirect(url_for("admin.users_list"))

    return render_template("admin_user_form.html", mode="edit",
                           user=user, companies=COMPANIES)


# ─── Excluir usuário ─────────────────────────────────


@admin_bp.route("/users/<int:user_id>/delete", methods=["POST"])
@admin_required
def user_delete(user_id):
    user = get_user_by_id(user_id)
    if not user:
        flash("Usuário não encontrado.", "error")
    elif user["username"] == "admin":
        flash("Não é possível excluir o admin padrão.", "error")
    else:
        delete_user(user_id)
        flash(f"Usuário '{user['username']}' excluído.", "success")
    return redirect(url_for("admin.users_list"))
