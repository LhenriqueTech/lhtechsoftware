# app.py
import sys
import os
import json
from pathlib import Path
from datetime import date
from typing import Dict, Any, Callable, Optional

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import Signal, Slot
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QMessageBox, QLabel, QVBoxLayout, QWidget
)

from styles import APP_QSS, COLORS
import lh_processor as processor
from jornada_dialog import JornadaDialog  # janelinha de jornadas

APP_NAME = "Gerador de Relatórios LH TECH"


def make_target_provider(model: Dict[str, Any]) -> Callable[[str, date], Optional[float]]:
    """
    Constrói uma função que recebe (nome, data) e devolve horas alvo (float) para a célula I (em horas).
    - model["defaults"][NOME] = [seg, ter, qua, qui, sex, sab, dom] (cada item em horas, ex.: 8, 7.5, 0, ...)
    - model["overrides"][NOME][YYYY-MM-DD] = horas (ex.: 7.0, 9.0, 0)
    Se não houver override nem default, retorna None e o processor usa fallback (8h dias úteis, 0h fds).
    """
    defaults = model.get("defaults", {})
    overrides = model.get("overrides", {})

    def _prov(nome: str, dt: date) -> Optional[float]:
        key = (nome or "").strip().upper()
        # override por data específica (formato iso)
        per_day = overrides.get(key, {})
        raw = per_day.get(dt.isoformat())
        if raw is not None:
            try:
                return float(raw)
            except Exception:
                return None

        # default por dia da semana
        week_arr = defaults.get(key)
        if week_arr and 0 <= dt.weekday() < 7:
            try:
                return float(week_arr[dt.weekday()])
            except Exception:
                return None

        return None

    return _prov


class Worker(QtCore.QObject):
    finished = Signal(str)
    progress = Signal(int)
    error = Signal(str)

    def __init__(self, in_file: str, out_dir: str, target_provider=None):
        super().__init__()
        self.in_file = in_file
        self.out_dir = out_dir
        self.target_provider = target_provider

    @Slot()
    def run(self):
        try:
            def _progress(pct, msg=""):
                self.progress.emit(int(pct))

            out_file = processor.process_file(
                self.in_file,
                self.out_dir,
                progress_callback=_progress,
                target_hours_provider=self.target_provider,  # passa jornada por pessoa/dia
            )
            self.finished.emit(out_file)
        except Exception as e:
            self.error.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setMinimumSize(860, 580)

        # estados
        self.selected_file: Optional[str] = None
        self.out_dir = str(Path.home())
        self.jornada_model: Dict[str, Any] = {"defaults": {}, "overrides": {}}
        self.preview_names = []  # nomes detectados (ordenados)

        # UI principal
        self.root = QWidget()
        self.setCentralWidget(self.root)
        self.vbox = QVBoxLayout(self.root)

        # header
        self.logo = QLabel()
        logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
        if os.path.exists(logo_path):
            pix = QtGui.QPixmap(logo_path).scaledToHeight(64, QtCore.Qt.SmoothTransformation)
            self.logo.setPixmap(pix)
        self.title = QLabel(
            f"<h2>Gerador de Relatórios <span style='color:{COLORS['primary']}'>LH TECH</span></h2>"
        )
        self.title.setTextFormat(QtCore.Qt.RichText)
        self.title.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)

        header = QtWidgets.QHBoxLayout()
        header.addWidget(self.logo)
        header.addWidget(self.title)
        header.addStretch()
        self.vbox.addLayout(header)

        # Drop area
        self.drop_area = QLabel("Arraste o arquivo base (.xlsx) aqui ou clique em 'Selecionar arquivo'")
        self.drop_area.setAcceptDrops(True)
        self.drop_area.setAlignment(QtCore.Qt.AlignCenter)
        self.drop_area.setFixedHeight(140)
        self.drop_area.setObjectName("dropArea")
        self.vbox.addWidget(self.drop_area)

        # Buttons
        hbox = QtWidgets.QHBoxLayout()
        self.btn_select = QtWidgets.QPushButton("Selecionar arquivo")
        self.btn_select.clicked.connect(self.select_file)

        self.btn_jornadas = QtWidgets.QPushButton("Jornadas… (key-users)")
        self.btn_jornadas.clicked.connect(self.open_jornadas_dialog)

        self.btn_load_jornadas = QtWidgets.QPushButton("Carregar jornadas.json")
        self.btn_load_jornadas.clicked.connect(self.load_jornadas_json)

        self.btn_generate = QtWidgets.QPushButton("Gerar Relatórios")
        self.btn_generate.setEnabled(False)
        self.btn_generate.clicked.connect(self.generate)

        self.btn_save = QtWidgets.QPushButton("Abrir pasta de saída")
        self.btn_save.setEnabled(False)
        self.btn_save.clicked.connect(self.open_out_folder)

        hbox.addWidget(self.btn_select)
        hbox.addWidget(self.btn_jornadas)
        hbox.addWidget(self.btn_load_jornadas)
        hbox.addStretch()
        hbox.addWidget(self.btn_generate)
        hbox.addWidget(self.btn_save)
        self.vbox.addLayout(hbox)

        # Preview
        self.preview_label = QLabel("Preview: nenhum arquivo selecionado")
        self.preview_label.setWordWrap(True)
        self.vbox.addWidget(self.preview_label)

        # Progress
        self.progress = QtWidgets.QProgressBar()
        self.progress.setValue(0)
        self.vbox.addWidget(self.progress)

        # Footer
        footer = QLabel("LH TECH")
        footer.setAlignment(QtCore.Qt.AlignRight)
        footer.setStyleSheet("color: #3B5D7A;")
        self.vbox.addWidget(footer)

        # DnD
        self.drop_area.installEventFilter(self)

        # Estilo
        self.setStyleSheet(APP_QSS)

    # --- Drag & Drop ---
    def eventFilter(self, obj, event):
        if obj is self.drop_area:
            if event.type() == QtCore.QEvent.DragEnter and event.mimeData().hasUrls():
                event.acceptProposedAction()
                return True
            if event.type() == QtCore.QEvent.Drop:
                urls = event.mimeData().urls()
                if urls:
                    self.load_file(urls[0].toLocalFile())
                return True
        return super().eventFilter(obj, event)

    # --- Arquivo base ---
    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Selecione o arquivo base (.xlsx)", str(Path.home()),
            "Excel Files (*.xlsx *.xlsm *.xls)"
        )
        if file_path:
            self.load_file(file_path)

    def load_file(self, path):
        self.selected_file = path
        self.preview_label.setText(f"Selecionado: {path}\n\nPreview automático: extraindo nomes e mês...")
        try:
            summary = processor.quick_preview(path)
            # nomes ordenados alfabeticamente também na janelinha
            self.preview_names = sorted(summary.get('names', []), key=str.lower)
            txt = (
                f"Arquivo: {Path(path).name}\n"
                f"Mês/Ano detectado: {summary.get('month_year', 'N/D')}\n"
                f"Abas detectadas: {', '.join(summary.get('sheets', []))}\n"
                f"Funcionários detectados: {', '.join(self.preview_names)}"
            )
            self.preview_label.setText(txt)
            self.btn_generate.setEnabled(True)
        except Exception as e:
            self.preview_label.setText(f"Erro ao ler o arquivo: {e}")
            self.btn_generate.setEnabled(False)

    # --- Jornadas (dialog) ---
    def open_jornadas_dialog(self):
        dlg = JornadaDialog(self, people=self.preview_names, initial_model=self.jornada_model)
        dlg.modelChanged.connect(self._on_jornada_model_changed)
        dlg.exec()

    def _on_jornada_model_changed(self, model: Dict[str, Any]):
        self.jornada_model = model

    # --- Jornadas via JSON ---
    def load_jornadas_json(self):
        path, _ = QFileDialog.getOpenFileName(self, "Carregar jornadas.json", str(Path.home()), "JSON (*.json)")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                self.jornada_model = json.load(f)
            QMessageBox.information(self, "OK", "Jornadas carregadas com sucesso.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível carregar o JSON:\n{e}")

    # --- Gerar ---
    def generate(self):
        if not self.selected_file:
            QMessageBox.warning(self, "Aviso", "Nenhum arquivo selecionado.")
            return

        out_dir = QFileDialog.getExistingDirectory(self, "Selecione pasta para salvar", self.out_dir)
        if not out_dir:
            return
        self.out_dir = out_dir

        provider = make_target_provider(self.jornada_model)

        # desabilita durante o processamento
        for btn in (self.btn_generate, self.btn_select, self.btn_save, self.btn_load_jornadas, self.btn_jornadas):
            btn.setEnabled(False)
        self.progress.setValue(0)

        self.thread = QtCore.QThread()
        self.worker = Worker(self.selected_file, self.out_dir, target_provider=provider)
        self.worker.moveToThread(self.thread)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.thread.started.connect(self.worker.run)
        self.thread.start()

    @Slot(int)
    def on_progress(self, pct):
        self.progress.setValue(pct)

    @Slot(str)
    def on_finished(self, out_file):
        self.progress.setValue(100)
        for btn in (self.btn_generate, self.btn_select, self.btn_save, self.btn_load_jornadas, self.btn_jornadas):
            btn.setEnabled(True)
        QMessageBox.information(self, "Concluído", f"Relatórios gerados em:\n{out_file}")
        try:
            self.thread.quit()
            self.thread.wait()
        except:
            pass

    @Slot(str)
    def on_error(self, msg):
        QMessageBox.critical(self, "Erro", f"Ocorreu um erro:\n{msg}")
        for btn in (self.btn_generate, self.btn_select, self.btn_save, self.btn_load_jornadas, self.btn_jornadas):
            btn.setEnabled(True)
        self.progress.setValue(0)
        try:
            self.thread.quit()
            self.thread.wait()
        except:
            pass

    def open_out_folder(self):
        if self.out_dir:
            QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(self.out_dir))


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
