import tkinter as tk
from tkinter import filedialog, messagebox
import os


class BaseAppRelatorio:
    """Shared tkinter GUI skeleton for the report-generation apps."""

    FILE_CONFIGS = {
        'export':        {'label': 'Arquivo Export:',           'btn_text': 'Selecionar Export'},
        'wms':           {'label': 'Arquivo WMS:',              'btn_text': 'Selecionar WMS'},
        'gbs':           {'label': 'Arquivo GBS (Neo Coelba):', 'btn_text': 'Selecionar GBS'},
        'classificacao': {'label': 'Arquivo Classifica\u00e7\u00e3o:',  'btn_text': 'Selecionar Classifica\u00e7\u00e3o'},
        'm2n':           {'label': 'Arquivo ME2N:',             'btn_text': 'Selecionar ME2N'},
    }

    def __init__(self, root):
        self.root = root
        self.root.title("logistic_machine - Gabriel Passos")
        self.root.geometry("650x520")

        self.arquivos = {
            key: {'label': cfg['label'], 'path': '', 'btn_text': cfg['btn_text']}
            for key, cfg in self.FILE_CONFIGS.items()
        }

        self.setup_ui()

    def setup_ui(self):
        tk.Label(self.root, text="Painel de Processamento",
                 font=("Arial", 14, "bold"), pady=20).pack()

        frame_files = tk.Frame(self.root, padx=20)
        frame_files.pack(fill="x")

        self.labels_status = {}

        for key, info in self.arquivos.items():
            row = tk.Frame(frame_files, pady=5)
            row.pack(fill="x")

            tk.Label(row, text=info['label'], width=22, anchor="w").pack(side="left")
            tk.Button(row, text=info['btn_text'], width=22,
                      command=lambda k=key: self.selecionar_arquivo(k)).pack(side="left", padx=10)

            status = tk.Label(row, text="N\u00e3o selecionado", fg="red",
                              wraplength=220, justify="left")
            status.pack(side="left")
            self.labels_status[key] = status

        self.btn_processar = tk.Button(
            self.root, text="GERAR RELAT\u00d3RIO FINAL",
            command=self.processar,
            bg="#4CAF50", fg="white",
            font=("Arial", 12, "bold"), pady=15, state="disabled"
        )
        self.btn_processar.pack(pady=30)

        tk.Label(self.root,
                 text="Selecione todos os arquivos para habilitar o processamento.",
                 font=("Arial", 8, "italic")).pack()

    def selecionar_arquivo(self, key):
        path = filedialog.askopenfilename(
            title=f"Selecione: {self.arquivos[key]['label']}",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.arquivos[key]['path'] = path
            self.labels_status[key].config(text=os.path.basename(path), fg="green")
            if all(info['path'] for info in self.arquivos.values()):
                self.btn_processar.config(state="normal")

    def save_with_dialog(self, df, default_filename="Recebimento_t\u00e9cnico_final.xlsx"):
        """Prompt the user to pick a save location and write the DataFrame."""
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if output_path:
            df.to_excel(output_path, index=False)
            messagebox.showinfo("Sucesso", f"Relat\u00f3rio gerado com sucesso!\n{output_path}")

    def processar(self):
        raise NotImplementedError("Subclasses must implement processar()")
