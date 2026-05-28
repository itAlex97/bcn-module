import os
import sys

from PySide6 import QtCore, QtWidgets, QtGui

from export import guardar_resultado
from main import calcular_resultado


class BomWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("BCN BOM Comparison")
        self.resize(920, 520)

        self.result_df = None
        self.result_path = None

        self._setup_ui()
        self._apply_stylesheet()

    def _setup_ui(self):
        main_widget = QtWidgets.QWidget()
        main_layout = QtWidgets.QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(12)

        title = QtWidgets.QLabel("BOM Comparison")
        title_font = title.font()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        main_layout.addWidget(title)

        main_layout.addWidget(self._create_input_section())
        main_layout.addWidget(self._create_action_section())
        main_layout.addWidget(self._create_results_section(), 1)

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

    def _create_input_section(self):
        container = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        layout.addWidget(QtWidgets.QLabel("Archivo Excel principal (hojas fijas)"))

        self.input_edit = QtWidgets.QLineEdit()
        self.input_edit.setPlaceholderText("Archivo Excel principal (hojas fijas)")
        self.input_edit.setMinimumHeight(34)
        input_row = QtWidgets.QHBoxLayout()
        input_row.addWidget(self.input_edit, 1)
        input_btn = self._create_button("Buscar", self.select_input, style="secondary")
        input_row.addWidget(input_btn)
        layout.addLayout(input_row)

        layout.addWidget(QtWidgets.QLabel("BD componentes (Excel)"))

        self.components_edit = QtWidgets.QLineEdit()
        self.components_edit.setPlaceholderText("BD componentes (Excel)")
        self.components_edit.setMinimumHeight(34)
        comp_row = QtWidgets.QHBoxLayout()
        comp_row.addWidget(self.components_edit, 1)
        comp_btn = self._create_button("Buscar", self.select_components, style="secondary")
        comp_row.addWidget(comp_btn)
        layout.addLayout(comp_row)

        container.setLayout(layout)
        return container

    def _create_action_section(self):
        layout = QtWidgets.QHBoxLayout()
        layout.setSpacing(8)

        self.process_btn = self._create_button("Procesar", self.run_process, style="primary")
        self.process_btn.setMinimumHeight(36)
        self.process_btn.setMinimumWidth(110)

        self.save_btn = self._create_button(
            "Guardar resultado", self.save_result, style="success"
        )
        self.save_btn.setEnabled(False)
        self.save_btn.setMinimumHeight(36)
        self.save_btn.setMinimumWidth(140)

        self.open_btn = self._create_button(
            "Abrir resultado", self.open_result, style="secondary"
        )
        self.open_btn.setEnabled(False)
        self.open_btn.setMinimumHeight(36)
        self.open_btn.setMinimumWidth(130)

        layout.addWidget(self.process_btn)
        layout.addWidget(self.save_btn)
        layout.addWidget(self.open_btn)
        layout.addStretch()

        self.status_label = QtWidgets.QLabel("Listo")
        status_font = self.status_label.font()
        status_font.setPointSize(9)
        self.status_label.setFont(status_font)
        layout.addWidget(self.status_label, 0, QtCore.Qt.AlignmentFlag.AlignVCenter)

        container = QtWidgets.QWidget()
        container.setLayout(layout)
        return container

    def _create_results_section(self):
        container = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        container.setLayout(layout)
        return container

    def _create_button(self, text, callback, style="primary"):
        btn = QtWidgets.QPushButton(text)
        btn.clicked.connect(callback)
        btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))

        btn.setObjectName(style)
        return btn

    def _apply_stylesheet(self):
        self.setStyleSheet(
            "QMainWindow { background-color: #F6F7FB; }"
            "QLabel { color: #1F2937; }"
            "QLineEdit { background: #FFFFFF; border: 1px solid #D6DAE3; "
            "border-radius: 6px; padding: 6px 8px; }"
            "QPlainTextEdit { background: #FFFFFF; border: 1px solid #E1E4EA; "
            "border-radius: 6px; padding: 8px; font-family: Consolas; font-size: 10px; }"
            "QPushButton { background: #FFFFFF; border: 1px solid #D6DAE3; "
            "border-radius: 6px; padding: 6px 12px; color: #1F2937; }"
            "QPushButton:hover { background: #EEF2F7; }"
            "QPushButton:disabled { color: #9CA3AF; background: #F3F4F6; }"
            "QPushButton#primary { background: #2563EB; color: #FFFFFF; border: none; }"
            "QPushButton#primary:hover { background: #1D4ED8; }"
            "QPushButton#success { background: #16A34A; color: #FFFFFF; border: none; }"
            "QPushButton#success:hover { background: #15803D; }"
        )

    def select_input(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Selecciona el Excel de entrada",
            "",
            "Excel (*.xlsx *.xls);;Todos (*.*)",
        )
        if path:
            self.input_edit.setText(path)

    def select_components(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Selecciona el Excel de componentes",
            "",
            "Excel (*.xlsx *.xls);;Todos (*.*)",
        )
        if path:
            self.components_edit.setText(path)

    def run_process(self):
        in_path = self.input_edit.text().strip()
        comp_path = self.components_edit.text().strip()

        if not in_path:
            QtWidgets.QMessageBox.warning(self, "Falta archivo", "Selecciona un archivo de entrada.")
            return
        if not comp_path:
            QtWidgets.QMessageBox.warning(self, "Falta BD", "Selecciona el Excel de componentes.")
            return

        try:
            self.status_label.setText("Procesando...")
            QtWidgets.QApplication.processEvents()
            self.log_message("Cargando datos...")
            (
                df_resultado,
                charted_partes,
                charted_modelos,
                mfg_partes,
                mfg_modelos,
                resumen,
            ) = calcular_resultado(in_path, ruta_componentes=comp_path)
            self.log_message(f"Charted: {charted_partes} partes, {charted_modelos} modelos")
            self.log_message(f"MFGPro:  {mfg_partes} partes, {mfg_modelos} modelos")
            if df_resultado.empty:
                self.log_message("No se encontraron diferencias entre los BOMs.")
            else:
                self.log_message(f"Diferencias: {len(df_resultado)} -> {resumen}")
            self.result_df = df_resultado
            self.save_btn.setEnabled(True)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Error", str(exc))
            self.status_label.setText("Error")
            return

        self.status_label.setText("Completado")
        QtWidgets.QMessageBox.information(self, "Listo", "Comparacion finalizada.")

    def save_result(self):
        if self.result_df is None:
            QtWidgets.QMessageBox.warning(self, "Sin resultado", "Primero ejecuta el proceso.")
            return

        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Guardar resultado",
            "Resultado_BOM.xlsx",
            "Excel (*.xlsx)",
        )
        if not path:
            return

        try:
            guardar_resultado(self.result_df, path)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Error", str(exc))
            return

        self.result_path = path
        self.open_btn.setEnabled(True)
        self.log_message(f"Resultado guardado: {path}")

    def open_result(self):
        if not self.result_path or not os.path.exists(self.result_path):
            QtWidgets.QMessageBox.warning(self, "No encontrado", "No existe el archivo guardado.")
            return
        os.startfile(self.result_path)

    def log_message(self, message):
        self.log.appendPlainText(message)


def run_app():
    app = QtWidgets.QApplication(sys.argv)
    win = BomWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    run_app()
