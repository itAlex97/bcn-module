import os
import tkinter as tk
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk

from export import guardar_resultado
from main import calcular_resultado


class BomApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("BCN BOM Comparison")
        self.geometry("820x520")
        self.resizable(True, True)

        self.input_path = tk.StringVar()
        self.components_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Listo")
        self.result_df = None
        self.result_path = None

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        ttk.Label(self, text="Archivo Excel (hojas fijas):").grid(
            row=0, column=0, sticky="w", **pad
        )
        ttk.Entry(self, textvariable=self.input_path, width=58).grid(
            row=1, column=0, sticky="w", **pad
        )
        ttk.Button(self, text="Buscar", command=self._select_input).grid(
            row=1, column=1, sticky="w", **pad
        )

        ttk.Label(self, text="BD componentes (Excel):").grid(
            row=2, column=0, sticky="w", **pad
        )
        ttk.Entry(self, textvariable=self.components_path, width=58).grid(
            row=3, column=0, sticky="w", **pad
        )
        ttk.Button(self, text="Buscar", command=self._select_components).grid(
            row=3, column=1, sticky="w", **pad
        )

        self.process_button = ttk.Button(self, text="Procesar", command=self._run, bootstyle="primary")
        self.process_button.grid(row=4, column=0, sticky="w", **pad)

        self.save_button = ttk.Button(
            self,
            text="Guardar resultado",
            command=self._save_result,
            state="disabled",
            bootstyle="success",
        )
        self.save_button.grid(row=4, column=1, sticky="w", **pad)

        self.open_button = ttk.Button(
            self,
            text="Abrir resultado",
            command=self._open_result,
            state="disabled",
        )
        self.open_button.grid(row=5, column=1, sticky="w", **pad)

        ttk.Label(self, textvariable=self.status_text, bootstyle="secondary").grid(
            row=5, column=0, sticky="w", **pad
        )

        self.log = tk.Text(self, width=78, height=9, state="disabled", relief="solid", bd=1)
        self.log.grid(row=6, column=0, columnspan=2, sticky="w", **pad)

    def _select_input(self):
        path = filedialog.askopenfilename(
            title="Selecciona el Excel de entrada",
            filetypes=[("Excel", "*.xlsx;*.xls"), ("Todos", "*.*")],
        )
        if path:
            self.input_path.set(path)

    def _select_components(self):
        path = filedialog.askopenfilename(
            title="Selecciona el Excel de componentes",
            filetypes=[("Excel", "*.xlsx;*.xls"), ("Todos", "*.*")],
        )
        if path:
            self.components_path.set(path)

    def _run(self):
        in_path = self.input_path.get().strip()
        comp_path = self.components_path.get().strip()

        if not in_path:
            messagebox.showerror("Falta archivo", "Selecciona un archivo de entrada.")
            return
        if not comp_path:
            messagebox.showerror("Falta BD", "Selecciona el Excel de componentes.")
            return

        try:
            self.status_text.set("Procesando...")
            self.update_idletasks()
            self._log("Cargando datos...")
            (
                df_resultado,
                charted_partes,
                charted_modelos,
                mfg_partes,
                mfg_modelos,
                resumen,
            ) = calcular_resultado(in_path, ruta_componentes=comp_path)
            self._log(f"Charted: {charted_partes} partes, {charted_modelos} modelos")
            self._log(f"MFGPro:  {mfg_partes} partes, {mfg_modelos} modelos")
            if df_resultado.empty:
                self._log("No se encontraron diferencias entre los BOMs.")
            else:
                self._log(f"Diferencias: {len(df_resultado)} -> {resumen}")
            self.result_df = df_resultado
            self.save_button.config(state="normal")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            self.status_text.set("Error")
            return

        self.status_text.set("Completado")
        messagebox.showinfo("Listo", "Comparacion finalizada.")

    def _save_result(self):
        if self.result_df is None:
            messagebox.showerror("Sin resultado", "Primero ejecuta el proceso.")
            return

        path = filedialog.asksaveasfilename(
            title="Guardar resultado",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="Resultado_BOM.xlsx",
        )
        if not path:
            return

        try:
            guardar_resultado(self.result_df, path)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        self.result_path = path
        self.open_button.config(state="normal")
        self._log(f"Resultado guardado: {path}")

    def _open_result(self):
        if not self.result_path or not os.path.exists(self.result_path):
            messagebox.showerror("No encontrado", "No existe el archivo guardado.")
            return
        os.startfile(self.result_path)

    def _log(self, message):
        self.log.config(state="normal")
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.log.config(state="disabled")


def run_app():
    app = BomApp()
    app.mainloop()
