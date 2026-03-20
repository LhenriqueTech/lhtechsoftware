# modules.py — Blueprint dos módulos por empresa
import sys
import os
import uuid
import json
import zipfile
from datetime import date
from pathlib import Path
from typing import Dict, Any, Optional, Callable

from flask import Blueprint, render_template, request, jsonify, send_file, abort
from flask_login import login_required, current_user

# Importa processadores
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import lh_processor as elleve_processor
import aquarela_processor as aquarela_processor
import maisrazao as mrcontabil_processor

PROCESSORS = {
    "elleve": elleve_processor,
    "aquarela": aquarela_processor,
    # mrcontabil usa rotas próprias abaixo (fluxo diferente: PDFs + datas + CPFs)
}

from models import COMPANIES

modules_bp = Blueprint("modules", __name__)

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


def make_target_provider(model: Dict[str, Any]) -> Callable[[str, date], Optional[float]]:
    defaults = model.get("defaults", {})
    overrides = model.get("overrides", {})

    def _prov(nome: str, dt: date) -> Optional[float]:
        key = (nome or "").strip().upper()
        per_day = overrides.get(key, {})
        raw = per_day.get(dt.isoformat())
        if raw is not None:
            try:
                return float(raw)
            except Exception:
                return None
        week_arr = defaults.get(key)
        if week_arr and 0 <= dt.weekday() < 7:
            try:
                return float(week_arr[dt.weekday()])
            except Exception:
                return None
        return None

    return _prov


# ─── Página do módulo ────────────────────────────────


@modules_bp.route("/modulo/<slug>")
@login_required
def module_page(slug):
    if slug not in COMPANIES:
        abort(404)
    if not current_user.has_company(slug):
        abort(403)

    company = COMPANIES[slug]

    if not company["active"]:
        return render_template("module_placeholder.html", company=company, slug=slug)

    # MR Contábil usa template próprio (fluxo de PDFs)
    if slug == "mrcontabil":
        return render_template("module_mrcontabil.html", company=company, slug=slug)

    return render_template("module.html", company=company, slug=slug)


# ─── Upload (xlsx — fluxo padrão) ────────────────────


@modules_bp.route("/modulo/<slug>/upload", methods=["POST"])
@login_required
def module_upload(slug):
    if slug not in COMPANIES or not current_user.has_company(slug):
        return jsonify({"error": "Acesso negado."}), 403

    processor = PROCESSORS.get(slug)
    if not processor:
        return jsonify({"error": "Módulo não implementado ainda."}), 501

    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado."}), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Nome de arquivo vazio."}), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".xls"):
        return jsonify({"error": "Formato inválido. Envie um arquivo .xlsx"}), 400

    file_id = str(uuid.uuid4())
    safe_name = f"{file_id}{ext}"
    save_path = os.path.join(UPLOAD_DIR, safe_name)
    f.save(save_path)

    try:
        summary = processor.quick_preview(save_path)
        names = sorted(summary.get("names", []), key=str.lower)
        predefined = {}
        for name in names:
            norm_key = processor._norm(name)
            cfg = processor.PREDEFINED_PEOPLE.get(norm_key)
            if cfg:
                predefined[name] = {
                    "cpf": cfg.get("cpf", "000.000.000-00"),
                    "week_times": cfg.get("week_times", [])
                }
        return jsonify({
            "file_id": file_id,
            "filename": f.filename,
            "names": names,
            "month_year": summary.get("month_year"),
            "sheets": summary.get("sheets", []),
            "predefined": predefined
        })

    except Exception as e:
        return jsonify({"error": f"Erro ao ler o arquivo: {e}"}), 400


# ─── Gerar (xlsx — fluxo padrão) ─────────────────────


@modules_bp.route("/modulo/<slug>/generate", methods=["POST"])
@login_required
def module_generate(slug):
    if slug not in COMPANIES or not current_user.has_company(slug):
        return jsonify({"error": "Acesso negado."}), 403

    processor = PROCESSORS.get(slug)
    if not processor:
        return jsonify({"error": "Módulo não implementado ainda."}), 501

    data = request.get_json()
    if not data or "file_id" not in data:
        return jsonify({"error": "file_id não informado."}), 400

    file_id = data["file_id"]
    in_file = None
    for fname in os.listdir(UPLOAD_DIR):
        if fname.startswith(file_id):
            in_file = os.path.join(UPLOAD_DIR, fname)
            break

    if not in_file or not os.path.exists(in_file):
        return jsonify({"error": "Arquivo não encontrado. Faça upload novamente."}), 404

    try:
        jornada_model = data.get("jornada_model", {"defaults": {}, "overrides": {}})
        provider = make_target_provider(jornada_model)

        out_subdir = os.path.join(OUTPUT_DIR, file_id)
        os.makedirs(out_subdir, exist_ok=True)

        out_path = processor.process_file(
            in_file, out_subdir,
            progress_callback=lambda p, m="": None,
            target_hours_provider=provider,
        )
        out_filename = os.path.basename(out_path)
        return jsonify({
            "success": True,
            "download_url": f"/modulo/{slug}/download/{file_id}/{out_filename}"
        })

    except Exception as e:
        return jsonify({"error": f"Erro ao gerar relatório: {e}"}), 500


# ─── Download (fluxo padrão) ─────────────────────────


@modules_bp.route("/modulo/<slug>/download/<file_id>/<filename>")
@login_required
def module_download(slug, file_id, filename):
    file_path = os.path.join(OUTPUT_DIR, file_id, filename)
    if not os.path.exists(file_path):
        return jsonify({"error": "Arquivo não encontrado."}), 404
    return send_file(file_path, as_attachment=True, download_name=filename)


# ═══════════════════════════════════════════════════════════════
# MR CONTÁBIL — Fluxo próprio (PDFs + datas + CPFs)
# ═══════════════════════════════════════════════════════════════

MR_CPF_DEFAULT = {
    "michelle silva de moraes do nascimento": "CPF: 293.480.658-82",
    "graziela delamura silva": "CPF: 435.842.658-19",
    "micaela araujo": "CPF: 528.570.258-58",
    "tahiara conceição": "CPF: 418.202.198-36",
}


@modules_bp.route("/modulo/mrcontabil/upload-pdfs", methods=["POST"])
@login_required
def mrcontabil_upload_pdfs():
    """Recebe múltiplos PDFs, salva na pasta temporária e retorna session_id."""
    if not current_user.has_company("mrcontabil"):
        return jsonify({"error": "Acesso negado."}), 403

    files = request.files.getlist("pdfs")
    if not files:
        return jsonify({"error": "Nenhum PDF enviado."}), 400

    session_id = str(uuid.uuid4())
    session_dir = os.path.join(UPLOAD_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)

    saved = []
    for f in files:
        if f.filename and f.filename.lower().endswith(".pdf"):
            safe = f.filename.replace("..", "_")
            path = os.path.join(session_dir, safe)
            f.save(path)
            saved.append(f.filename)

    if not saved:
        return jsonify({"error": "Nenhum PDF válido encontrado."}), 400

    return jsonify({"session_id": session_id, "files": saved})


@modules_bp.route("/modulo/mrcontabil/generate", methods=["POST"])
@login_required
def mrcontabil_generate():
    """Processa os PDFs e gera os relatórios .xlsx, retorna ZIP para download."""
    if not current_user.has_company("mrcontabil"):
        return jsonify({"error": "Acesso negado."}), 403

    data = request.get_json()
    if not data:
        return jsonify({"error": "Payload inválido."}), 400

    session_id = data.get("session_id")
    start_date = data.get("start_date", "").strip()
    end_date = data.get("end_date", "").strip()
    cpf_dict = data.get("cpf_dict", MR_CPF_DEFAULT)

    if not session_id or not start_date or not end_date:
        return jsonify({"error": "session_id, start_date e end_date são obrigatórios."}), 400

    pdf_folder = os.path.join(UPLOAD_DIR, session_id)
    if not os.path.isdir(pdf_folder):
        return jsonify({"error": "Sessão não encontrada. Faça o upload dos PDFs novamente."}), 404

    out_subdir = os.path.join(OUTPUT_DIR, session_id)
    os.makedirs(out_subdir, exist_ok=True)

    try:
        generated = mrcontabil_processor.process_pdfs(
            pdf_folder=pdf_folder,
            output_folder=out_subdir,
            start_date=start_date,
            end_date=end_date,
            cpf_dict=cpf_dict,
        )
    except Exception as e:
        return jsonify({"error": f"Erro ao processar PDFs: {e}"}), 500

    if not generated:
        return jsonify({"error": "Nenhum dado encontrado nos PDFs para o período informado."}), 400

    # Cria ZIP com todos os relatórios
    zip_name = f"relatorios_mrcontabil_{session_id[:8]}.zip"
    zip_path = os.path.join(out_subdir, zip_name)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in generated:
            zf.write(f, os.path.basename(f))

    return jsonify({
        "success": True,
        "count": len(generated),
        "download_url": f"/modulo/mrcontabil/download-zip/{session_id}/{zip_name}"
    })


@modules_bp.route("/modulo/mrcontabil/download-zip/<session_id>/<filename>")
@login_required
def mrcontabil_download_zip(session_id, filename):
    file_path = os.path.join(OUTPUT_DIR, session_id, filename)
    if not os.path.exists(file_path):
        return jsonify({"error": "Arquivo não encontrado."}), 404
    return send_file(file_path, as_attachment=True, download_name=filename)
