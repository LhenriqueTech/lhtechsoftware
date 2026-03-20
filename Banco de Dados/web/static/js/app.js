// ═══════════════════════════════════════════════
// LH TECH — Sistema de Relatórios (Web) v2.0
// app.js — Lógica do módulo de upload/geração
// ═══════════════════════════════════════════════

// Slug do módulo e nome, injetados pelo template
const MODULE_SLUG = window.MODULE_SLUG || "elleve";
const MODULE_NAME = window.MODULE_NAME || "Módulo";

// ── Estado global ──
const state = {
    fileId: null,
    filename: null,
    names: [],
    predefined: {},
    jornadaModel: { defaults: {}, overrides: {}, cpfs: {} },
    currentPerson: null,
};

// ── Helpers DOM ──
const $ = (id) => document.getElementById(id);

const dropZone       = $("drop-zone");
const fileInput      = $("file-input");
const fileInfo       = $("file-info");
const fileName       = $("file-name");
const fileMeta       = $("file-meta");
const btnRemoveFile  = $("btn-remove-file");

const previewSection = $("preview-section");
const previewMonth   = $("preview-month");
const previewCount   = $("preview-count");
const previewSheets  = $("preview-sheets");
const namesList      = $("names-list");

const actionsSection = $("actions-section");
const btnJornadas    = $("btn-jornadas");
const btnGenerate    = $("btn-generate");

const progressContainer = $("progress-container");
const progressFill      = $("progress-fill");
const progressText      = $("progress-text");

const downloadBox    = $("download-box");
const btnDownload    = $("btn-download");

// Modal
const modalOverlay   = $("modal-overlay");
const modalClose     = $("modal-close");
const jornadaPerson  = $("jornada-person");
const jornadaCpf     = $("jornada-cpf");
const exceptionsBody = $("exceptions-body");
const btnAddException = $("btn-add-exception");
const btnApply       = $("btn-apply-jornadas");
const btnLoadJson    = $("btn-load-json");
const btnSaveJson    = $("btn-save-json");
const jsonFileInput  = $("json-file-input");

// ── API base URL para o módulo ──
const API = `/modulo/${MODULE_SLUG}`;

// ═══════════════════════════════════════════════
// UPLOAD
// ═══════════════════════════════════════════════

if (dropZone) {
    dropZone.addEventListener("click", () => fileInput.click());

    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("drag-over");
    });

    dropZone.addEventListener("dragleave", () => {
        dropZone.classList.remove("drag-over");
    });

    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("drag-over");
        if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });
}

if (fileInput) {
    fileInput.addEventListener("change", () => {
        if (fileInput.files.length) handleFile(fileInput.files[0]);
    });
}

if (btnRemoveFile) {
    btnRemoveFile.addEventListener("click", () => resetState());
}

async function handleFile(file) {
    const formData = new FormData();
    formData.append("file", file);

    dropZone.style.display = "none";
    fileInfo.style.display = "flex";
    fileName.textContent = file.name;
    fileMeta.textContent = "Enviando e analisando…";
    previewSection.style.display = "none";
    actionsSection.style.display = "none";
    downloadBox.style.display = "none";
    progressContainer.style.display = "none";

    try {
        const res = await fetch(`${API}/upload`, { method: "POST", body: formData });
        const data = await res.json();

        if (!res.ok || data.error) {
            toast(data.error || "Erro ao enviar arquivo.", "error");
            resetState();
            return;
        }

        state.fileId     = data.file_id;
        state.filename   = data.filename;
        state.names      = data.names || [];
        state.predefined = data.predefined || {};

        initJornadaModelFromPredefined();

        fileMeta.textContent = `${formatBytes(file.size)} • ${state.names.length} funcionário(s)`;

        previewMonth.textContent  = data.month_year || "N/D";
        previewCount.textContent  = state.names.length;
        previewSheets.textContent = (data.sheets || []).join(", ") || "N/D";

        namesList.innerHTML = "";
        state.names.forEach((n) => {
            const span = document.createElement("span");
            span.className = "name-tag";
            span.textContent = n;
            namesList.appendChild(span);
        });

        previewSection.style.display = "block";
        actionsSection.style.display = "block";

    } catch (err) {
        toast("Erro de conexão com o servidor.", "error");
        resetState();
    }
}

function initJornadaModelFromPredefined() {
    for (const name of state.names) {
        const key = name.trim().toUpperCase();
        const pred = state.predefined[name];

        if (!state.jornadaModel.defaults[key]) {
            if (pred && pred.week_times) {
                state.jornadaModel.defaults[key] = pred.week_times.map(timeStrToFloat);
            } else {
                state.jornadaModel.defaults[key] = [8.8, 8.8, 8.8, 8.8, 8.8, 0, 0];
            }
        }
        if (!state.jornadaModel.cpfs[key]) {
            if (pred && pred.cpf) state.jornadaModel.cpfs[key] = pred.cpf;
        }
        if (!state.jornadaModel.overrides[key]) {
            state.jornadaModel.overrides[key] = {};
        }
    }
}

function resetState() {
    state.fileId = null;
    state.filename = null;
    state.names = [];
    state.predefined = {};

    dropZone.style.display = "";
    fileInfo.style.display = "none";
    previewSection.style.display = "none";
    actionsSection.style.display = "none";
    downloadBox.style.display = "none";
    progressContainer.style.display = "none";
    fileInput.value = "";
}

// ═══════════════════════════════════════════════
// GERAR RELATÓRIOS
// ═══════════════════════════════════════════════

if (btnGenerate) {
    btnGenerate.addEventListener("click", async () => {
        if (!state.fileId) { toast("Nenhum arquivo selecionado.", "error"); return; }

        btnGenerate.disabled = true;
        btnJornadas.disabled = true;
        progressContainer.style.display = "block";
        downloadBox.style.display = "none";
        progressFill.style.width = "0%";
        progressText.textContent = "Iniciando processamento...";

        let fakeProgress = 0;
        const progressInterval = setInterval(() => {
            fakeProgress = Math.min(fakeProgress + Math.random() * 8, 85);
            progressFill.style.width = fakeProgress + "%";
            progressText.textContent = `Processando… ${Math.round(fakeProgress)}%`;
        }, 300);

        try {
            const res = await fetch(`${API}/generate`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    file_id: state.fileId,
                    jornada_model: state.jornadaModel,
                }),
            });
            const data = await res.json();
            clearInterval(progressInterval);

            if (!res.ok || data.error) {
                progressFill.style.width = "0%";
                progressText.textContent = "Erro!";
                toast(data.error || "Erro ao gerar relatório.", "error");
            } else {
                progressFill.style.width = "100%";
                progressText.textContent = "Concluído!";
                btnDownload.href = data.download_url;
                downloadBox.style.display = "flex";
                toast("Relatório gerado com sucesso!", "success");
            }
        } catch (err) {
            clearInterval(progressInterval);
            toast("Erro de conexão.", "error");
        }

        btnGenerate.disabled = false;
        btnJornadas.disabled = false;
    });
}

// ═══════════════════════════════════════════════
// MODAL — JORNADAS
// ═══════════════════════════════════════════════

if (btnJornadas)   btnJornadas.addEventListener("click", () => openModal());
if (modalClose)    modalClose.addEventListener("click", () => closeModal());
if (modalOverlay)  modalOverlay.addEventListener("click", (e) => { if (e.target === modalOverlay) closeModal(); });

function openModal() {
    jornadaPerson.innerHTML = "";
    state.names.forEach((n) => {
        const opt = document.createElement("option");
        opt.value = n;
        opt.textContent = n;
        jornadaPerson.appendChild(opt);
    });
    if (state.names.length) loadPersonFields(state.names[0]);
    modalOverlay.style.display = "flex";
    document.body.style.overflow = "hidden";
}

function closeModal() {
    captureCurrentPerson();
    modalOverlay.style.display = "none";
    document.body.style.overflow = "";
}

if (jornadaPerson) {
    jornadaPerson.addEventListener("change", () => {
        captureCurrentPerson();
        loadPersonFields(jornadaPerson.value);
    });
}

function loadPersonFields(name) {
    state.currentPerson = name;
    const key = name.trim().toUpperCase();
    const defaults = state.jornadaModel.defaults[key] || [8.8, 8.8, 8.8, 8.8, 8.8, 0, 0];
    for (let i = 0; i < 7; i++) $("jw-" + i).value = floatToHhmm(defaults[i] || 0);
    jornadaCpf.value = state.jornadaModel.cpfs[key] || "";
    exceptionsBody.innerHTML = "";
    const overrides = state.jornadaModel.overrides[key] || {};
    for (const [dt, hours] of Object.entries(overrides).sort()) addExceptionRow(dt, floatToHhmm(hours));
}

function captureCurrentPerson() {
    if (!state.currentPerson) return;
    const key = state.currentPerson.trim().toUpperCase();
    const arr = [];
    for (let i = 0; i < 7; i++) arr.push(hhmmToFloat($("jw-" + i).value));
    state.jornadaModel.defaults[key] = arr;
    const cpf = jornadaCpf.value.trim();
    if (cpf) state.jornadaModel.cpfs[key] = cpf;
    const ov = {};
    exceptionsBody.querySelectorAll("tr").forEach((tr) => {
        const inputs = tr.querySelectorAll("input");
        const dt = inputs[0]?.value?.trim();
        const hh = inputs[1]?.value?.trim();
        if (dt) ov[dt] = hhmmToFloat(hh || "0");
    });
    state.jornadaModel.overrides[key] = ov;
}

if (btnAddException) {
    btnAddException.addEventListener("click", () => {
        addExceptionRow(new Date().toISOString().split("T")[0], "00:00");
    });
}

function addExceptionRow(dateStr, hhmm) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
        <td><input type="date" value="${dateStr}"></td>
        <td><input type="text" value="${hhmm}" placeholder="hh:mm"></td>
        <td><button class="btn-icon btn-remove-exc" title="Remover"><span class="material-icons-round">delete</span></button></td>
    `;
    tr.querySelector(".btn-remove-exc").addEventListener("click", () => tr.remove());
    exceptionsBody.appendChild(tr);
}

if (btnApply) {
    btnApply.addEventListener("click", () => {
        captureCurrentPerson();
        toast("Configurações de jornadas aplicadas!", "success");
        closeModal();
    });
}

// JSON import/export
if (btnLoadJson)  btnLoadJson.addEventListener("click", () => jsonFileInput.click());
if (jsonFileInput) {
    jsonFileInput.addEventListener("change", () => {
        const file = jsonFileInput.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                state.jornadaModel = JSON.parse(e.target.result);
                toast("Jornadas carregadas do JSON.", "success");
                if (state.currentPerson) loadPersonFields(state.currentPerson);
            } catch (err) {
                toast("Erro ao parsear JSON.", "error");
            }
        };
        reader.readAsText(file);
        jsonFileInput.value = "";
    });
}

if (btnSaveJson) {
    btnSaveJson.addEventListener("click", () => {
        captureCurrentPerson();
        const blob = new Blob([JSON.stringify(state.jornadaModel, null, 2)], { type: "application/json" });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "jornadas.json";
        a.click();
        URL.revokeObjectURL(a.href);
        toast("JSON salvo!", "info");
    });
}

// ═══════════════════════════════════════════════
// UTILITÁRIOS
// ═══════════════════════════════════════════════

function hhmmToFloat(s) {
    s = (s || "").trim();
    if (!s) return 0;
    if (s.includes(":")) {
        const parts = s.split(":");
        return (parseInt(parts[0]) || 0) + (parseInt(parts[1]) || 0) / 60.0;
    }
    return parseFloat(s.replace(",", ".")) || 0;
}

function floatToHhmm(x) {
    if (!x && x !== 0) return "00:00";
    const totalMin = Math.round(x * 60);
    return `${String(Math.floor(totalMin / 60)).padStart(2, "0")}:${String(totalMin % 60).padStart(2, "0")}`;
}

function timeStrToFloat(s) {
    s = (s || "").trim();
    if (!s) return 0;
    const parts = s.split(":");
    return (parseInt(parts[0]) || 0) + (parseInt(parts[1]) || 0) / 60.0;
}

function formatBytes(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

function toast(msg, type = "info") {
    const container = $("toast-container");
    if (!container) return;
    const div = document.createElement("div");
    div.className = `toast ${type}`;
    div.textContent = msg;
    container.appendChild(div);
    setTimeout(() => div.remove(), 4000);
}
