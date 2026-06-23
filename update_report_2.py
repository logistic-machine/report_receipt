import logging
import traceback

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
import os

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


class AppRelatorio:
    def __init__(self, root):
        self.root = root
        self.root.title("logistic_machine - Gabriel Passos")
        self.root.geometry("650x520")

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
            self.arquivos[key]['path'] = path
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

    def normalizar_fornecedor(self, nome):
        """Extrai a primeira palavra do fornecedor para chave de cruzamento."""
        if pd.isna(nome):
            return "DESCONHECIDO"
        nome_str = str(nome).strip().upper()
        partes = nome_str.split()
        return partes[0] if partes else "DESCONHECIDO"

    def carregar_arquivo(self, key):
        """Load an Excel file with proper error reporting."""
        path = self.arquivos[key]['path']
        label = self.arquivos[key]['label']

        if not path:
            raise FileNotFoundError(f"Nenhum arquivo selecionado para '{label}'.")

        if not os.path.isfile(path):
            raise FileNotFoundError(
                f"Arquivo não encontrado para '{label}': {path}"
            )

        df = pd.read_excel(path)
        if df.empty:
            raise ValueError(f"Arquivo '{label}' está vazio: {path}")

        logger.info("Arquivo '%s' carregado: %d linhas, %d colunas", label, len(df), len(df.columns))
        return df

    def validar_colunas(self, df, colunas, descricao):
        """Validate that required columns exist in a DataFrame."""
        faltando = [c for c in colunas if c not in df.columns]
        if faltando:
            raise KeyError(
                f"Colunas obrigatórias ausentes em '{descricao}': {faltando}\n"
                f"Colunas disponíveis: {list(df.columns)}"
            )

    # ------------------------------------------------------------------
    # Processamento principal
    # ------------------------------------------------------------------
    def processar(self):
        try:
            self._processar_interno()
        except FileNotFoundError as e:
            logger.error("Arquivo não encontrado: %s", e)
            messagebox.showerror("Arquivo Não Encontrado", str(e))
        except KeyError as e:
            logger.error("Coluna não encontrada: %s", e)
            messagebox.showerror("Coluna Ausente", str(e))
        except ValueError as e:
            logger.error("Erro de validação: %s", e)
            messagebox.showerror("Erro de Validação", str(e))
        except PermissionError as e:
            logger.error("Permissão negada: %s", e)
            messagebox.showerror("Permissão Negada", str(e))
        except Exception as e:
            logger.error("Erro inesperado: %s\n%s", e, traceback.format_exc())
            messagebox.showerror(
                "Erro Inesperado",
                f"Ocorreu um erro não esperado durante o processamento:\n\n"
                f"{type(e).__name__}: {e}\n\n"
                f"Verifique o log para mais detalhes."
            )

    def _processar_interno(self):
        # ── 1. Carregar arquivos ───────────────────────────────────
        df_export = self.carregar_arquivo('export')
        df_wms = self.carregar_arquivo('wms')
        df_gbs = self.carregar_arquivo('gbs')
        df_classificacao = self.carregar_arquivo('classificacao')
        df_m2n = self.carregar_arquivo('m2n')

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

        if not col_nf_wms:
            raise KeyError(
                "Não foi possível identificar a coluna de Nota Fiscal no WMS.\n"
                f"Colunas disponíveis: {list(df_wms.columns)}"
            )
        if not col_cod_wms:
            raise KeyError(
                "Não foi possível identificar a coluna de Código do Material no WMS.\n"
                f"Colunas disponíveis: {list(df_wms.columns)}"
            )

        colunas_wms_opcionais_ausentes = []
        if not col_qtd_wms:
            colunas_wms_opcionais_ausentes.append("Quantidade")
        if not col_desc_wms:
            colunas_wms_opcionais_ausentes.append("Descrição")
        if not col_data_wms:
            colunas_wms_opcionais_ausentes.append("Data")
        if not col_centro_wms:
            colunas_wms_opcionais_ausentes.append("Centro")

        if colunas_wms_opcionais_ausentes:
            logger.warning(
                "Colunas opcionais não encontradas no WMS: %s. "
                "Essas informações ficarão ausentes no relatório.",
                colunas_wms_opcionais_ausentes,
            )

        # ── 3. Colunas do ME2N ────────────────────────────────────
        col_pedido_m2n = 'Doc.compra'
        col_aprovacao = self.encontrar_coluna(df_m2n, ['Lib', 'Aprovacao do Pedido', 'Aprovação'])
        col_grupo_compradores = self.encontrar_coluna(df_m2n, ['GCm', 'Grupo Compradores', 'Grp. Compradores'])

        if col_pedido_m2n not in df_m2n.columns:
            raise KeyError(
                f"Coluna '{col_pedido_m2n}' não encontrada no ME2N.\n"
                f"Colunas disponíveis: {list(df_m2n.columns)}"
            )
        if not col_aprovacao:
            raise KeyError(
                "Não foi possível identificar a coluna 'Aprovação do Pedido' no ME2N.\n"
                f"Colunas disponíveis: {list(df_m2n.columns)}"
            )
        if not col_grupo_compradores:
            raise KeyError(
                "Não foi possível identificar a coluna 'Grupo de Compradores' no ME2N.\n"
                f"Colunas disponíveis: {list(df_m2n.columns)}"
            )

        # ── 4. Validate required columns in Export, GBS, Classificação ──
        self.validar_colunas(df_export, ['Nº NF-e', 'Valor total'], "Export")
        self.validar_colunas(df_gbs, ['Nº NF-e', 'Observação (GBS)', 'Pedido de Compras'], "GBS")
        self.validar_colunas(df_classificacao, ['Material', 'Planejador'], "Classificação")

        # ── 5. Fornecedor normalization ───────────────────────────
        possiveis_fornecedores = [
            'CNPJ Fornecedor', 'Fornecedor', 'Emitente',
            'Nome emissor', 'Razão Social', 'Desc. Fornecedor',
        ]

        col_forn_export = self.encontrar_coluna(df_export, possiveis_fornecedores)
        if col_forn_export:
            df_export['Forn_Norm'] = df_export[col_forn_export].apply(self.normalizar_fornecedor)
        else:
            logger.warning(
                "Coluna de fornecedor não encontrada no Export. "
                "Usando valor padrão 'SEM_FORN' para chave de cruzamento."
            )
            df_export['Forn_Norm'] = "SEM_FORN"

        df_export['Chave_NF_Forn'] = df_export['Nº NF-e'].astype(str).str.strip() + "-" + df_export['Forn_Norm']

        col_forn_wms = self.encontrar_coluna(df_wms, possiveis_fornecedores)
        if col_forn_wms:
            df_wms['Forn_Norm'] = df_wms[col_forn_wms].apply(self.normalizar_fornecedor)
        else:
            logger.warning(
                "Coluna de fornecedor não encontrada no WMS. "
                "Usando valor padrão 'SEM_FORN' para chave de cruzamento."
            )
            df_wms['Forn_Norm'] = "SEM_FORN"

        df_wms['Chave_NF_Forn'] = df_wms[col_nf_wms].astype(str).str.strip() + "-" + df_wms['Forn_Norm']

        col_forn_gbs = self.encontrar_coluna(df_gbs, possiveis_fornecedores)
        if col_forn_gbs:
            df_gbs['Forn_Norm'] = df_gbs[col_forn_gbs].apply(self.normalizar_fornecedor)
        else:
            logger.warning(
                "Coluna de fornecedor não encontrada no GBS. "
                "Usando valor padrão 'SEM_FORN' para chave de cruzamento."
            )
            df_gbs['Forn_Norm'] = "SEM_FORN"

        df_gbs['Chave_NF_Forn'] = df_gbs['Nº NF-e'].astype(str).str.strip() + "-" + df_gbs['Forn_Norm']

        # ── 6. Prepare WMS item key ───────────────────────────────
        df_wms['Material_Str'] = df_wms[col_cod_wms].astype(str).str.strip().str.lstrip('0')
        df_wms['Chave_Item'] = df_wms['Chave_NF_Forn'] + "-" + df_wms['Material_Str']

        # GBS: deduplicate by composite key
        df_gbs_sel = df_gbs[['Chave_NF_Forn', 'Observação (GBS)', 'Pedido de Compras']].drop_duplicates(subset='Chave_NF_Forn')
        df_gbs_sel['Pedido de Compras'] = self.padronizar_pedido(df_gbs_sel['Pedido de Compras'])

        # ME2N
        df_m2n_sel = df_m2n[[col_pedido_m2n, col_aprovacao, col_grupo_compradores]].drop_duplicates(subset=col_pedido_m2n)
        df_m2n_sel[col_pedido_m2n] = self.padronizar_pedido(df_m2n_sel[col_pedido_m2n])

        # ── 7. Merges ─────────────────────────────────────────────
        wms_merge_cols = [c for c in [col_nf_wms, col_cod_wms, col_desc_wms, col_qtd_wms, col_data_wms, col_centro_wms, 'Chave_NF_Forn'] if c]
        df_final = pd.merge(df_export, df_wms[wms_merge_cols],
                            left_on='Chave_NF_Forn', right_on='Chave_NF_Forn', how='left')

        df_final = pd.merge(df_final, df_gbs_sel, on='Chave_NF_Forn', how='left')

        df_final['Pedido de Compras'] = self.padronizar_pedido(df_final['Pedido de Compras'])
        df_final = pd.merge(df_final, df_m2n_sel, left_on='Pedido de Compras', right_on=col_pedido_m2n, how='left')

        # ── 8. Valor Total Ajustado ───────────────────────────────
        df_final = df_final.sort_values(by=['Chave_NF_Forn', col_cod_wms])
        df_final['Valor total Ajustado'] = np.where(
            df_final['Chave_NF_Forn'] != df_final['Chave_NF_Forn'].shift(),
            df_final['Valor total'], 0
        )

        # ── 9. Status Real ────────────────────────────────────────
        df_final['Status Real'] = np.where(
            df_final[col_nf_wms].notna(), "RECEBIDO (WMS)", "PENDENTE / EM TRÂNSITO"
        )

        # ── 10. Planejador ────────────────────────────────────────
        df_planejador = df_classificacao[['Material', 'Planejador']].drop_duplicates(subset='Material')
        df_final['Material_Merge'] = pd.to_numeric(df_final[col_cod_wms], errors='coerce').astype('Int64')
        df_planejador['Material'] = pd.to_numeric(df_planejador['Material'], errors='coerce').astype('Int64')

        coerce_nulos = df_final['Material_Merge'].isna().sum()
        if coerce_nulos > 0:
            logger.warning(
                "%d valores de código do material não puderam ser convertidos para numérico "
                "e serão tratados como nulos no cruzamento com Classificação.",
                coerce_nulos,
            )

        df_final = pd.merge(df_final, df_planejador, left_on='Material_Merge', right_on='Material', how='left')

        # ── 11. Rename and cleanup ────────────────────────────────
        rename_map = {}
        if col_cod_wms:
            rename_map[col_cod_wms] = 'Cód do material'
        if col_desc_wms:
            rename_map[col_desc_wms] = 'Descrição do material'
        if col_qtd_wms:
            rename_map[col_qtd_wms] = 'Qnt recebida'
        if col_data_wms:
            rename_map[col_data_wms] = 'Data de recebimento'
        rename_map['Observação (GBS)'] = 'Observação GBS'

        df_final = df_final.rename(columns=rename_map)

        cols_to_drop = ['Forn_Norm', 'Chave_NF_Forn', col_nf_wms, col_pedido_m2n, 'Material_Merge', 'Material']
        df_final = df_final.drop(columns=[c for c in cols_to_drop if c in df_final.columns])

        # ── 12. Save ──────────────────────────────────────────────
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Recebimento_técnico_final.xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not output_path:
            logger.info("Salvamento cancelado pelo usuário.")
            messagebox.showwarning("Cancelado", "Nenhum arquivo foi salvo.")
            return

        try:
            df_final.to_excel(output_path, index=False)
        except PermissionError:
            raise PermissionError(
                f"Sem permissão para salvar em '{output_path}'.\n"
                f"Verifique se o arquivo está aberto em outro programa."
            )

        logger.info("Relatório salvo em: %s", output_path)
        messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso!\n{output_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AppRelatorio(root)
    root.mainloop()
