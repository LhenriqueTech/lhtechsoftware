# auth.py — Blueprint de autenticação
from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from models import get_user_by_id, get_user_by_username, verify_password, update_last_login

auth_bp = Blueprint("auth", __name__)

# ─── Flask-Login User class ─────────────────────────


class User(UserMixin):
    def __init__(self, user_dict):
        self.id = user_dict["id"]
        self.username = user_dict["username"]
        self.full_name = user_dict["full_name"]
        self.role = user_dict["role"]
        self.companies = user_dict["companies"]
        self.tools = user_dict.get("tools", [])
        self.active = user_dict["active"]

    @property
    def is_admin(self):
        return self.role == "admin"

    @property
    def is_fiscal(self):
        return self.is_admin or "conversor" in self.tools

    def has_company(self, slug):
        return self.is_admin or slug in self.companies



def init_login_manager(app):
    login_manager = LoginManager()
    login_manager.login_view = "auth.login"
    login_manager.login_message = "Faça login para continuar."
    login_manager.login_message_category = "info"
    login_manager.init_app(app)

    @login_manager.user_loader
    def load_user(user_id):
        data = get_user_by_id(int(user_id))
        if data and data["active"]:
            return User(data)
        return None


# ─── Rotas ───────────────────────────────────────────


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        user_data = get_user_by_username(username)
        if not user_data:
            flash("Usuário ou senha inválidos.", "error") # Consolidated message
            return render_template("login.html")

        if not user_data["active"]:
            flash("Esta conta está desativada. Entre em contato com o administrador.", "error")
            return render_template("login.html")

        if not verify_password(user_data, password):
            flash("Usuário ou senha inválidos.", "error") # Consolidated message
            return render_template("login.html")

        user = User(user_data)
        login_user(user, remember=True)

        # Registra Auditoria (Login + IP)
        try:
            ip = request.headers.get('X-Forwarded-For', request.remote_addr)
            if ',' in ip: ip = ip.split(',')[0].strip()
            update_last_login(user.id, ip)
        except Exception:
            pass # Não trava o login se falhar o registro de IP

        next_page = request.args.get("next")
        return redirect(next_page or url_for("dashboard"))

    return render_template("login.html")


@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("auth.login"))
