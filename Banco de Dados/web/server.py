# server.py — Flask App Principal — LH TECH
import os
import secrets
from flask import Flask, redirect, url_for, render_template
from flask_login import login_required, current_user

from models import init_db, COMPANIES
from auth import auth_bp, init_login_manager
from admin import admin_bp
from modules import modules_bp

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

# Registra blueprints
app.register_blueprint(auth_bp)
app.register_blueprint(admin_bp)
app.register_blueprint(modules_bp)

# Configura Flask-Login
init_login_manager(app)

# Inicializa banco de dados
with app.app_context():
    init_db()


# Context processor — torna 'companies' disponível em todos os templates
@app.context_processor
def inject_companies():
    if current_user.is_authenticated:
        if current_user.is_admin:
            return {"companies": COMPANIES}
        return {"companies": {k: v for k, v in COMPANIES.items() if k in current_user.companies}}
    return {"companies": {}}


# ─── Rotas principais ───────────────────────────────


@app.route("/")
def index():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    return redirect(url_for("auth.login"))


@app.route("/dashboard")
@login_required
def dashboard():
    # Filtra empresas que o usuário pode acessar
    if current_user.is_admin:
        user_companies = COMPANIES
    else:
        user_companies = {k: v for k, v in COMPANIES.items() if k in current_user.companies}
    return render_template("dashboard.html", companies=user_companies)


@app.errorhandler(403)
def forbidden(e):
    return render_template("error.html", code=403, message="Acesso negado."), 403


@app.errorhandler(404)
def not_found(e):
    return render_template("error.html", code=404, message="Página não encontrada."), 404


if __name__ == "__main__":
    print("=" * 50)
    print("  LH TECH — Sistema de Relatórios (Web)")
    print("  Acesse: http://localhost:5000")
    print("  Login padrão: admin / admin123")
    print("=" * 50)
    app.run(debug=True, host="0.0.0.0", port=5000)
