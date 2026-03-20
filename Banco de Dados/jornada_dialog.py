from __future__ import annotations
import json
import unicodedata
from datetime import date
from typing import Dict, Any

from PySide6 import QtCore, QtGui, QtWidgets

# Tenta importar as configurações fixas do código (lh_processor.PREDEFINED_PEOPLE)
try:
    import lh_processor as processor
except ImportError:
    processor = None

WEEK_LABELS = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]

# Jornada padrão 08:48 (44h semanais) em HORAS (float)
DEFAULT_8848 = [8 + 48/60.0, 8 + 48/60.0, 8 + 48/60.0, 8 + 48/60.0, 8 + 48/60.0, 0.0, 0.0]
DEFAULT_CPF = "000.000.000-00"


def _norm(s: str) -> str:
    """Normaliza string removendo acentos e deixando maiúscula."""
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.strip().upper()


def _week_times_to_floats(week_times) -> list[float]:
    """
    Converte lista de strings HH:MM ou HH:MM:SS em horas (float).
    Ignora segundos (usa apenas HH e MM).
    """
    arr: list[float] = []
    for s in (week_times or []):
        s = (s or "").strip()
        if not s:
            arr.append(0.0)
            continue
        try:
            parts = s.split(":")
            hh = int(parts[0])
            mm = int(parts[1]) if len(parts) > 1 else 0
            arr.append(hh + mm / 60.0)
        except Exception:
            arr.append(0.0)
    # garante 7 posições
    while len(arr) < 7:
        arr.append(0.0)
    return arr[:7]


# Carrega PREDEFINED_PEOPLE do lh_processor, se existir
if processor is not None and hasattr(processor, "PREDEFINED_PEOPLE"):
    _PREDEFINED_FROM_CODE: Dict[str, Dict[str, Any]] = {}
    for name, cfg in processor.PREDEFINED_PEOPLE.items():
        key = _norm(name)
        _PREDEFINED_FROM_CODE[key] = cfg
else:
    _PREDEFINED_FROM_CODE = {}


def hhmm_to_float(s: str) -> float:
    s = (s or "").strip()
    if not s:
        return 0.0
    if ":" in s:
        try:
            parts = s.split(":")
            hh = int(parts[0])
            mm = int(parts[1]) if len(parts) > 1 else 0
            return hh + mm / 60.0
        except Exception:
            return 0.0
    try:
        return float(s.replace(",", "."))
    except Exception:
        return 0.0


def float_to_hhmm(x: float) -> str:
    if x is None:
        return "00:00"
    total_min = int(round(x * 60))
    hh = total_min // 60
    mm = total_min % 60
    return f"{hh:02d}:{mm:02d}"


class JornadaDialog(QtWidgets.QDialog):
    """Editor para defaults (Seg..Dom), CPF e exceções por data, por colaborador."""
    modelChanged = QtCore.Signal(dict)  # emite o modelo completo ao clicar "Aplicar"

    def __init__(self, parent=None, people: list[str] | None = None, initial_model: Dict[str, Any] | None = None):
        super().__init__(parent)
        self.setWindowTitle("Configurar Jornadas (key-users)")
        self.resize(720, 560)

        self.people = sorted(people or [], key=lambda s: s.upper())
        self.model = initial_model or {"defaults": {}, "overrides": {}, "cpfs": {}}

        # topo: seleção de colaborador
        self.cmbPeople = QtWidgets.QComboBox()
        self.cmbPeople.addItems(self.people)
        self.cmbPeople.currentIndexChanged.connect(self._load_person_fields)

        # campo de CPF
        self.edCPF = QtWidgets.QLineEdit()
        self.edCPF.setPlaceholderText("CPF (000.000.000-00)")
        self.edCPF.setMaxLength(20)

        # defaults seg..dom
        defaults_layout = QtWidgets.QGridLayout()
        self.edWeek: list[QtWidgets.QLineEdit] = []
        for i, lab in enumerate(WEEK_LABELS):
            defaults_layout.addWidget(QtWidgets.QLabel(lab), 0, i)
            ed = QtWidgets.QLineEdit()
            ed.setPlaceholderText("hh:mm")
            ed.setFixedWidth(80)
            ed.setAlignment(QtCore.Qt.AlignCenter)
            self.edWeek.append(ed)
            defaults_layout.addWidget(ed, 1, i)

        grpDef = QtWidgets.QGroupBox("Padrão por dia da semana")
        grpDef.setLayout(defaults_layout)

        # tabela de exceções
        self.tbl = QtWidgets.QTableWidget(0, 2)
        self.tbl.setHorizontalHeaderLabels(["Data (YYYY-MM-DD)", "Horas (hh:mm)"])
        self.tbl.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tbl.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked
            | QtWidgets.QAbstractItemView.EditKeyPressed
        )

        # Ajustes de espaçamento da coluna "Data"
        hdr = self.tbl.horizontalHeader()
        hdr.setStretchLastSection(True)
        hdr.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        hdr.setMinimumSectionSize(120)
        self.tbl.setColumnWidth(0, 160)

        btnAdd = QtWidgets.QPushButton("Adicionar exceção")
        btnDel = QtWidgets.QPushButton("Remover selecionada")
        btnAdd.clicked.connect(self._add_exception)
        btnDel.clicked.connect(self._remove_selected)

        grpExc = QtWidgets.QGroupBox("Exceções por data")
        excLay = QtWidgets.QVBoxLayout()
        excLay.addWidget(self.tbl, 1)
        rowBtns = QtWidgets.QHBoxLayout()
        rowBtns.addWidget(btnAdd)
        rowBtns.addWidget(btnDel)
        rowBtns.addStretch(1)
        excLay.addLayout(rowBtns)
        grpExc.setLayout(excLay)

        # botões de rodapé
        btnOpen = QtWidgets.QPushButton("Abrir JSON…")
        btnSave = QtWidgets.QPushButton("Salvar JSON…")
        btnApply = QtWidgets.QPushButton("Aplicar")
        btnClose = QtWidgets.QPushButton("Fechar")

        btnOpen.clicked.connect(self._open_json)
        btnSave.clicked.connect(self._save_json)
        btnApply.clicked.connect(self._emit_model)
        btnClose.clicked.connect(self.close)

        # layout raiz
        root = QtWidgets.QVBoxLayout(self)

        top = QtWidgets.QHBoxLayout()
        top.addWidget(QtWidgets.QLabel("Colaborador:"))
        top.addWidget(self.cmbPeople, 1)

        top_cpf = QtWidgets.QHBoxLayout()
        top_cpf.addWidget(QtWidgets.QLabel("CPF:"))
        top_cpf.addWidget(self.edCPF, 1)

        root.addLayout(top)
        root.addLayout(top_cpf)
        root.addWidget(grpDef)
        root.addWidget(grpExc, 1)

        bottom = QtWidgets.QHBoxLayout()
        bottom.addWidget(btnOpen)
        bottom.addWidget(btnSave)
        bottom.addStretch(1)
        bottom.addWidget(btnApply)
        bottom.addWidget(btnClose)
        root.addLayout(bottom)

        if self.people:
            self._load_person_fields()

    # ------------ helpers ------------
    def _person_key(self) -> str:
        return (self.cmbPeople.currentText() or "").strip()

    def _load_person_fields(self):
        key = self._person_key()
        norm_key = _norm(key)

        # === Jornada padrão (defaults) ===
        defaults = self.model.get("defaults", {})
        arr = defaults.get(key)

        if arr is None:
            # tenta buscar do PREDEFINED_PEOPLE do código
            cfg = _PREDEFINED_FROM_CODE.get(norm_key)
            if cfg and "week_times" in cfg:
                arr = _week_times_to_floats(cfg["week_times"])

        if arr is None:
            # se ainda não encontrou, usa padrão 08:48 seg–sex
            arr = DEFAULT_8848[:]

        for i in range(7):
            self.edWeek[i].setText(float_to_hhmm(arr[i]))

        # === CPF ===
        cpfs = self.model.get("cpfs", {})
        cpf_val = cpfs.get(key)

        if not cpf_val:
            cfg = _PREDEFINED_FROM_CODE.get(norm_key)
            if cfg and cfg.get("cpf"):
                cpf_val = str(cfg["cpf"])

        if not cpf_val:
            cpf_val = DEFAULT_CPF

        self.edCPF.setText(cpf_val)

        # === Exceções ===
        self.tbl.setRowCount(0)
        ov_all = self.model.get("overrides", {})
        ov = ov_all.get(key, {})
        for dstr, hours in sorted(ov.items()):
            self._append_row(dstr, float_to_hhmm(hours))

    def _append_row(self, dstr: str, hhmm: str):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)
        it0 = QtWidgets.QTableWidgetItem(dstr)
        it1 = QtWidgets.QTableWidgetItem(hhmm)
        it0.setTextAlignment(QtCore.Qt.AlignCenter)
        it1.setTextAlignment(QtCore.Qt.AlignCenter)
        self.tbl.setItem(r, 0, it0)
        self.tbl.setItem(r, 1, it1)

    def _add_exception(self):
        self._append_row(date.today().isoformat(), "00:00")

    def _remove_selected(self):
        rows = {idx.row() for idx in self.tbl.selectedIndexes()}
        for r in sorted(rows, reverse=True):
            self.tbl.removeRow(r)

    def _capture_person_to_model(self):
        key = self._person_key()
        if not key:
            return

        # jornada semanal
        arr = [hhmm_to_float(self.edWeek[i].text()) for i in range(7)]
        self.model.setdefault("defaults", {})[key] = arr

        # CPF
        cpf = (self.edCPF.text() or "").strip()
        if cpf:
            self.model.setdefault("cpfs", {})[key] = cpf

        # exceções por data
        m: Dict[str, float] = {}
        for r in range(self.tbl.rowCount()):
            d = (self.tbl.item(r, 0).text() if self.tbl.item(r, 0) else "").strip()
            h = (self.tbl.item(r, 1).text() if self.tbl.item(r, 1) else "").strip()
            if d:
                m[d] = hhmm_to_float(h)
        self.model.setdefault("overrides", {})[key] = m

    # ------------ JSON I/O ------------
    def _open_json(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Abrir jornadas.json", "", "JSON (*.json)")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                self.model = json.load(f)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Não foi possível abrir:\n{e}")
            return
        self._load_person_fields()

    def _save_json(self):
        self._capture_person_to_model()
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Salvar jornadas.json", "jornadas.json", "JSON (*.json)")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.model, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Não foi possível salvar:\n{e}")

    # ------------ Aplicar ------------
    def _emit_model(self):
        self._capture_person_to_model()
        self.modelChanged.emit(self.model)
        QtWidgets.QMessageBox.information(self, "OK", "Configurações aplicadas.")
