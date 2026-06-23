import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import messagebox

from shared.base_app import BaseAppRelatorio
from shared.utils import (
    encontrar_coluna,
    padronizar_pedido,
    normalize_nf_column,
    resolve_wms_columns,
    compute_adjusted_total,
    compute_status_real,
    build_wms_rename_map,
    merge_planejador,
)


class AppRelatorio(BaseAppRelatorio):

    def processar(self):
        try:
            # 1. Carregar arquivos
            df_export        = pd.read_excel(self.arquivos['export']['path'])
            df_wms           = pd.read_excel(self.arquivos['wms']['path'])
            df_gbs           = pd.read_excel(self.arquivos['gbs']['path'])
            df_classificacao = pd.read_excel(self.arquivos['classificacao']['path'])
            df_m2n           = pd.read_excel(self.arquivos['m2n']['path'])

            # 2. Mapeamento de colunas do WMS
            wms_cols   = resolve_wms_columns(df_wms)
            col_nf_wms = wms_cols['nf']
            col_cod_wms = wms_cols['cod']

            if not col_nf_wms or not col_cod_wms:
                messagebox.showerror("Erro", "Não foi possível identificar colunas essenciais no WMS.")
                return

            # 3. Colunas do ME2N
            col_pedido_m2n        = 'Doc.compra'
            col_aprovacao         = encontrar_coluna(df_m2n, ['Lib', 'Aprovacao do Pedido', 'Aprovação'])
            col_grupo_compradores = encontrar_coluna(df_m2n, ['GCm', 'Grupo Compradores', 'Grp. Compradores'])

            if col_pedido_m2n not in df_m2n.columns:
                messagebox.showerror("Erro", f"Coluna '{col_pedido_m2n}' não encontrada no ME2N.\n"
                                             f"Colunas disponíveis: {list(df_m2n.columns)}")
                return
            if not col_aprovacao or not col_grupo_compradores:
                messagebox.showerror("Erro", "Não foi possível identificar 'Aprovação do Pedido' ou "
                                             "'Grupo de Compradores' no ME2N.\n"
                                             f"Colunas disponíveis: {list(df_m2n.columns)}")
                return

            # 4. Padronização das colunas de cruzamento
            normalize_nf_column(df_export, 'Nº NF-e')
            normalize_nf_column(df_wms, col_nf_wms)
            normalize_nf_column(df_gbs, 'Nº NF-e')
            df_gbs['Pedido de Compras'] = padronizar_pedido(df_gbs['Pedido de Compras'])
            df_m2n[col_pedido_m2n]      = padronizar_pedido(df_m2n[col_pedido_m2n])

            # 5. Selecionar colunas relevantes de cada base
            colunas_wms = [c for c in [col_nf_wms, wms_cols['cod'], wms_cols['desc'],
                                       wms_cols['qtd'], wms_cols['data'], wms_cols['centro']] if c]
            df_wms_sel = df_wms[colunas_wms]

            df_gbs_sel = (
                df_gbs[['Nº NF-e', 'Observação (GBS)', 'Pedido de Compras']]
                .drop_duplicates(subset='Nº NF-e')
            )

            df_m2n_sel = (
                df_m2n[[col_pedido_m2n, col_aprovacao, col_grupo_compradores]]
                .drop_duplicates(subset=col_pedido_m2n)
            )

            # 6. Merges
            df_final = pd.merge(df_export, df_wms_sel,
                                left_on='Nº NF-e', right_on=col_nf_wms, how='left')
            df_final = pd.merge(df_final, df_gbs_sel,
                                left_on='Nº NF-e', right_on='Nº NF-e', how='left')
            df_final['Pedido de Compras'] = padronizar_pedido(df_final['Pedido de Compras'])
            df_final = pd.merge(df_final, df_m2n_sel,
                                left_on='Pedido de Compras', right_on=col_pedido_m2n, how='left')

            if col_pedido_m2n in df_final.columns and col_pedido_m2n != 'Pedido de Compras':
                df_final = df_final.drop(columns=[col_pedido_m2n])

            # 7. Valor Total Ajustado + Status Real
            df_final = df_final.sort_values(by='Nº NF-e')
            df_final = compute_adjusted_total(df_final, group_col='Nº NF-e')
            df_final = compute_status_real(df_final, col_nf_wms)

            # 8. Renomear colunas
            rename_map = build_wms_rename_map(wms_cols)
            rename_map['Observação (GBS)'] = 'Observação GBS'
            df_final = df_final.rename(columns={k: v for k, v in rename_map.items() if k})

            # 9. Planejador
            df_final = merge_planejador(df_final, df_classificacao, 'Cód do material')

            # 10. Verificação no console
            total        = len(df_final)
            cruzados_wms = df_final['Status Real'].eq("RECEBIDO (WMS)").sum()
            cruzados_m2n = df_final[col_aprovacao].notna().sum()
            print(f"Total de linhas     : {total}")
            print(f"Recebidos (WMS)     : {cruzados_wms}")
            print(f"Cruzados c/ ME2N    : {cruzados_m2n}")
            print(f"Sem match ME2N      : {total - cruzados_m2n}")

            # 11. Salvar
            self.save_with_dialog(df_final)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante o processamento:\n\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AppRelatorio(root)
    root.mainloop()
