import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
import os


class AppRelatorio:
    def __init__(self, root):
        self.root = root
        self.root.title("logistic_machine - Gabriel Passos")
        self.root.geometry("650x520")

        # Dicionário para armazenar os caminhos dos arquivos
        self.arquivos = {
            'export':        {'label': 'Arquivo Export:',        'path': '', 'btn_text': 'Selecionar Export'},
            'wms':           {'label': 'Arquivo WMS:',           'path': '', 'btn_text': 'Selecionar WMS'},
            'gbs':           {'label': 'Arquivo GBS (Neo Coelba):','path': '', 'btn_text': 'Selecionar GBS'},
            'classificacao': {'label': 'Arquivo Classificação:', 'path': '', 'btn_text': 'Selecionar Classificação'},
            'm2n':           {'label': 'Arquivo ME2N:',          'path': '', 'btn_text': 'Selecionar ME2N'},
        }

        self.setup_ui()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
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

            status = tk.Label(row, text="Não selecionado", fg="red",
                              wraplength=220, justify="left")
            status.pack(side="left")
            self.labels_status[key] = status

        self.btn_processar = tk.Button(
            self.root, text="GERAR RELATÓRIO FINAL",
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
            if not path.lower().endswith(('.xlsx', '.xls')):
                messagebox.showwarning("Aviso", "Selecione um arquivo Excel (.xlsx ou .xls).")
                return
            self.arquivos[key]['path'] = os.path.normpath(path)
            self.labels_status[key].config(text=os.path.basename(path), fg="green")
            if all(info['path'] for info in self.arquivos.values()):
                self.btn_processar.config(state="normal")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def encontrar_coluna(self, df, possiveis_nomes):
        """Retorna o primeiro nome encontrado nas colunas do DataFrame."""
        for nome in possiveis_nomes:
            if nome and nome in df.columns:
                return nome
        return None

    def padronizar_pedido(self, serie):
        """Converte para string, remove espaços e zeros à esquerda."""
        return serie.astype(str).str.strip().str.lstrip('0')

    # ------------------------------------------------------------------
    # Processamento principal
    # ------------------------------------------------------------------
    def processar(self):
        try:
            # ── 1. Carregar arquivos ───────────────────────────────────
            df_export        = pd.read_excel(self.arquivos['export']['path'])
            df_wms           = pd.read_excel(self.arquivos['wms']['path'])
            df_gbs           = pd.read_excel(self.arquivos['gbs']['path'])
            df_classificacao = pd.read_excel(self.arquivos['classificacao']['path'])
            df_m2n           = pd.read_excel(self.arquivos['m2n']['path'])

            # ── 2. Mapeamento de colunas do WMS ───────────────────────
            ncols = len(df_wms.columns)
            mapeamento_wms = {
                'nf':     ['NF', 'Nota Fiscal', 'Nº NF-e', 'Nf',
                           df_wms.columns[-1]  if ncols >= 1 else None],
                'qtd':    ['Quantidade', 'Qtd', 'Qnt recebida', 'Qtd.',
                           df_wms.columns[-2]  if ncols >= 2 else None],
                'desc':   ['Descrição', 'Desc. Material', 'Descrição do material',
                           df_wms.columns[-4]  if ncols >= 4 else None],
                'cod':    ['Código', 'Cód. Material', 'Cód do material', 'Material',
                           df_wms.columns[-5]  if ncols >= 5 else None],
                'data':   ['Data', 'Data de recebimento', 'Data Recebimento',
                           df_wms.columns[-9]  if ncols >= 9 else None],
                'centro': ['Centro', 'Depósito', 'Unidade',
                           df_wms.columns[-12] if ncols >= 12 else None],
            }

            col_nf_wms     = self.encontrar_coluna(df_wms, mapeamento_wms['nf'])
            col_qtd_wms    = self.encontrar_coluna(df_wms, mapeamento_wms['qtd'])
            col_desc_wms   = self.encontrar_coluna(df_wms, mapeamento_wms['desc'])
            col_cod_wms    = self.encontrar_coluna(df_wms, mapeamento_wms['cod'])
            col_data_wms   = self.encontrar_coluna(df_wms, mapeamento_wms['data'])
            col_centro_wms = self.encontrar_coluna(df_wms, mapeamento_wms['centro'])

            if not col_nf_wms or not col_cod_wms:
                messagebox.showerror("Erro", "Não foi possível identificar colunas essenciais no WMS.")
                return

            # ── 3. Colunas do ME2N ────────────────────────────────────
            col_pedido_m2n       = 'Doc.compra'
            col_aprovacao        = self.encontrar_coluna(df_m2n, ['Lib', 'Aprovacao do Pedido', 'Aprovação'])
            col_grupo_compradores= self.encontrar_coluna(df_m2n, ['GCm', 'Grupo Compradores', 'Grp. Compradores'])

            if col_pedido_m2n not in df_m2n.columns:
                messagebox.showerror("Erro", f"Coluna '{col_pedido_m2n}' não encontrada no ME2N.\n"
                                             f"Colunas disponíveis: {list(df_m2n.columns)}")
                return
            if not col_aprovacao or not col_grupo_compradores:
                messagebox.showerror("Erro", "Não foi possível identificar 'Aprovação do Pedido' ou "
                                             "'Grupo de Compradores' no ME2N.\n"
                                             f"Colunas disponíveis: {list(df_m2n.columns)}")
                return

            # ── 4. Padronização das colunas de cruzamento ─────────────
            df_export['Nº NF-e']       = df_export['Nº NF-e'].astype(str).str.strip()
            df_wms[col_nf_wms]         = df_wms[col_nf_wms].astype(str).str.strip()
            df_gbs['Nº NF-e']          = df_gbs['Nº NF-e'].astype(str).str.strip()

            # Pedido de Compras: remove zeros à esquerda para garantir match
            df_gbs['Pedido de Compras']      = self.padronizar_pedido(df_gbs['Pedido de Compras'])
            df_m2n[col_pedido_m2n]           = self.padronizar_pedido(df_m2n[col_pedido_m2n])

            # ── 5. Selecionar colunas relevantes de cada base ─────────
            colunas_wms = [c for c in [col_nf_wms, col_cod_wms, col_desc_wms,
                                       col_qtd_wms, col_data_wms, col_centro_wms] if c]
            df_wms_sel = df_wms[colunas_wms]

            df_gbs_sel = (
                df_gbs[['Nº NF-e', 'Observação (GBS)', 'Pedido de Compras']]
                .drop_duplicates(subset='Nº NF-e')
            )

            df_m2n_sel = (
                df_m2n[[col_pedido_m2n, col_aprovacao, col_grupo_compradores]]
                .drop_duplicates(subset=col_pedido_m2n)
            )

            # ── 6. Merges ─────────────────────────────────────────────
            # 6.1 Export ← WMS (por NF)
            df_final = pd.merge(df_export, df_wms_sel,
                                left_on='Nº NF-e', right_on=col_nf_wms, how='left')

            # 6.2 + GBS (por NF) — traz Observação e Pedido de Compras
            df_final = pd.merge(df_final, df_gbs_sel,
                                left_on='Nº NF-e', right_on='Nº NF-e', how='left')

            # Padronizar Pedido de Compras no df_final antes de cruzar com ME2N
            df_final['Pedido de Compras'] = self.padronizar_pedido(df_final['Pedido de Compras'])

            # 6.3 + ME2N (por Pedido de Compras) — traz Aprovação e Grupo de Compradores
            df_final = pd.merge(df_final, df_m2n_sel,
                                left_on='Pedido de Compras', right_on=col_pedido_m2n, how='left')

            # Remove coluna duplicada do pedido vinda do ME2N
            if col_pedido_m2n in df_final.columns and col_pedido_m2n != 'Pedido de Compras':
                df_final = df_final.drop(columns=[col_pedido_m2n])

            # ── 7. Valor Total Ajustado ───────────────────────────────
            df_final = df_final.sort_values(by='Nº NF-e')
            df_final['Valor total Ajustado'] = np.where(
                df_final['Nº NF-e'] != df_final['Nº NF-e'].shift(),
                df_final['Valor total'], 0
            )

            # ── 8. Status Real ────────────────────────────────────────
            df_final['Status Real'] = df_final[col_nf_wms].apply(
                lambda x: "RECEBIDO (WMS)" if pd.notna(x) else "PENDENTE / EM TRÂNSITO"
            )

            # ── 9. Renomear colunas ───────────────────────────────────
            rename_map = {
                col_cod_wms:          'Cód do material',
                col_desc_wms:         'Descrição do material',
                col_qtd_wms:          'Qnt recebida',
                col_data_wms:         'Data de recebimento',
                'Observação (GBS)':   'Observação GBS',
            }
            df_final = df_final.rename(columns={k: v for k, v in rename_map.items() if k})

            # ── 10. Planejador (por código do material) ───────────────
            df_planejador = (
                df_classificacao[['Material', 'Planejador']]
                .drop_duplicates(subset='Material')
            )
            df_final['Cód do material'] = (
                pd.to_numeric(df_final['Cód do material'], errors='coerce').astype('Int64')
            )
            df_planejador['Material'] = (
                pd.to_numeric(df_planejador['Material'], errors='coerce').astype('Int64')
            )
            df_final = pd.merge(df_final, df_planejador,
                                left_on='Cód do material', right_on='Material',
                                how='left').drop(columns=['Material'])

            # ── 11. Verificação no console ────────────────────────────
            total         = len(df_final)
            cruzados_wms  = df_final['Status Real'].eq("RECEBIDO (WMS)").sum()
            cruzados_m2n  = df_final[col_aprovacao].notna().sum()
            print(f"✅ Total de linhas     : {total}")
            print(f"✅ Recebidos (WMS)     : {cruzados_wms}")
            print(f"✅ Cruzados c/ ME2N    : {cruzados_m2n}")
            print(f"⚠️  Sem match ME2N     : {total - cruzados_m2n}")

            # ── 12. Salvar ────────────────────────────────────────────
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile="Recebimento_técnico_final.xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if output_path:
                df_final.to_excel(output_path, index=False)
                messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso!\n{output_path}")

        except Exception:
            messagebox.showerror("Erro", "Erro durante o processamento. Verifique se os arquivos selecionados estão corretos e tente novamente.")


# ── Entry point ───────────────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    app = AppRelatorio(root)
    root.mainloop()
