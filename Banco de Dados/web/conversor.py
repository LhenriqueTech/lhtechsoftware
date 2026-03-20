# conversor.py — Blueprint do Conversor UTF-8 → ANSI (somente cargo Fiscal/Admin)
import os
import sys
import io
import zipfile
import uuid
from functools import wraps
from concurrent.futures import ThreadPoolExecutor

from flask import Blueprint, render_template, request, jsonify, send_file, redirect, url_for, flash
from flask_login import login_required, current_user

# Importa a lógica de conversão do converter.py (na pasta pai)
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from converter import convert_file

conversor_bp = Blueprint("conversor", __name__)

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads", "conversor")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output",  "conversor")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ─── Decorator de acesso Fiscal ──────────────────────────────────────────────

def fiscal_required(f):
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if not current_user.is_fiscal:
            flash("Acesso restrito a usuários com cargo Fiscal.", "error")
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)
    return decorated


# ─── Página principal ────────────────────────────────────────────────────────

@conversor_bp.route("/conversor")
@fiscal_required
def conversor_page():
    return render_template("module_conversor.html")


# ─── Converter ───────────────────────────────────────────────────────────────

@conversor_bp.route("/conversor/converter", methods=["POST"])
@fiscal_required
def conversor_run():
    """
    Recebe um ou mais arquivos .txt (UTF-8), converte para ANSI (cp1252)
    e devolve um ZIP para download.
    """
    files = request.files.getlist("txts")
    if not files:
        return jsonify({"error": "Nenhum arquivo enviado."}), 400

    # Filtra apenas .txt
    valid = [f for f in files if f.filename.lower().endswith(".txt")]
    if not valid:
        return jsonify({"error": "Envie arquivos .txt."}), 400

    session_id = str(uuid.uuid4())
    in_dir  = os.path.join(UPLOAD_DIR, session_id)
    out_dir = os.path.join(OUTPUT_DIR, session_id)
    os.makedirs(in_dir,  exist_ok=True)
    os.makedirs(out_dir, exist_ok=True)

    # Salva os arquivos recebidos
    saved_pairs = []
    for f in valid:
        safe_name = os.path.basename(f.filename).replace("..", "_")
        in_path   = os.path.join(in_dir, safe_name)
        out_name  = "ANSI_" + safe_name
        out_path  = os.path.join(out_dir, out_name)
        f.save(in_path)
        saved_pairs.append((in_path, out_path, out_name))

    # Converte em paralelo usando a função do converter.py
    errors = []
    with ThreadPoolExecutor() as executor:
        futures = {
            executor.submit(convert_file, inp, outp): name
            for inp, outp, name in saved_pairs
        }
        for future, name in futures.items():
            try:
                future.result()
            except Exception as e:
                errors.append(f"{name}: {e}")

    if errors:
        return jsonify({"error": "Erros na conversão:\n" + "\n".join(errors)}), 500

    # Empacota em ZIP na memória
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for _, outp, out_name in saved_pairs:
            if os.path.exists(outp):
                zf.write(outp, out_name)
    zip_buffer.seek(0)

    zip_filename = f"ANSI_convertidos_{session_id[:8]}.zip"
    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name=zip_filename,
        mimetype="application/zip",
    )
