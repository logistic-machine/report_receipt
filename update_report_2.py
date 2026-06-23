import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import messagebox

from shared.base_app import BaseAppRelatorio
from shared.utils import (
    encontrar_coluna,
    padronizar_pedido,
    normalizar_fornecedor,
    resolve_wms_columns,
    compute_adjusted_total,
    compute_status_real,
    build_wms_rename_map,
    merge_planejador,
    add_supplier_key_columns,
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
            wms_cols    = resolve_wms_columns(df_wms)
            col_nf_wms  = wms_cols['nf']
            col_cod_wms = wms_cols['cod']

            if not col_nf_wms or not col_cod_wms:
                messagebox.showerror("Erro", "Não foi possível identificar colunas essenciais no WMS.")
                return

            col_pedido_m2n        = 'Doc.compra'
            col_aprovacao         = encontrar_coluna(df_m2n, ['Lib', 'Aprovacao do Pedido', 'Aprovação'])
            col_grupo_compradores = encontrar_coluna(df_m2n, ['GCm', 'Grupo Compradores', 'Grp. Compradores'])

            # 3. Chave composta NF + Fornecedor
            possiveis_fornecedores = [
                'CNPJ Fornecedor', 'Fornecedor', 'Emitente',
                'Nome emissor', 'Razão Social', 'Desc. Fornecedor',
            ]

            add_supplier_key_columns(df_export, 'Nº NF-e', possiveis_fornecedores)
            add_supplier_key_columns(df_wms, col_nf_wms, possiveis_fornecedores)
            add_supplier_key_columns(df_gbs, 'Nº NF-e', possiveis_fornecedores)

            df_wms['Material_Str'] = df_wms[col_cod_wms].astype(str).str.strip().str.lstrip('0')
            df_wms['Chave_Item']   = df_wms['Chave_NF_Forn'] + "-" + df_wms['Material_Str']

            # 4. Selecionar colunas relevantes
            df_gbs_sel = (
                df_gbs[['Chave_NF_Forn', 'Observação (GBS)', 'Pedido de Compras']]
                .drop_duplicates(subset='Chave_NF_Forn')
            )
            df_gbs_sel['Pedido de Compras'] = padronizar_pedido(df_gbs_sel['Pedido de Compras'])

            df_m2n_sel = (
                df_m2n[[col_pedido_m2n, col_aprovacao, col_grupo_compradores]]
                .drop_duplicates(subset=col_pedido_m2n)
            )
            df_m2n_sel[col_pedido_m2n] = padronizar_pedido(df_m2n_sel[col_pedido_m2n])

            # 5. Merges
            col_qtd_wms    = wms_cols['qtd']
            col_desc_wms   = wms_cols['desc']
            col_data_wms   = wms_cols['data']
            col_centro_wms = wms_cols['centro']

            df_final = pd.merge(
                df_export,
                df_wms[[col_nf_wms, col_cod_wms, col_desc_wms, col_qtd_wms,
                         col_data_wms, col_centro_wms, 'Chave_NF_Forn']],
                left_on='Chave_NF_Forn', right_on='Chave_NF_Forn', how='left',
            )
            df_final = pd.merge(df_final, df_gbs_sel, on='Chave_NF_Forn', how='left')
            df_final['Pedido de Compras'] = padronizar_pedido(df_final['Pedido de Compras'])
            df_final = pd.merge(df_final, df_m2n_sel,
                                left_on='Pedido de Compras', right_on=col_pedido_m2n, how='left')

            # 6. Valor Total Ajustado + Status Real
            df_final = df_final.sort_values(by=['Chave_NF_Forn', col_cod_wms])
            df_final = compute_adjusted_total(df_final, group_col='Chave_NF_Forn')
            df_final = compute_status_real(df_final, col_nf_wms)

            # 7. Planejador (before rename so we can reference the original column)
            df_final['Material_Merge'] = (
                pd.to_numeric(df_final[col_cod_wms], errors='coerce').astype('Int64')
            )
            df_final = merge_planejador(df_final, df_classificacao, 'Material_Merge')

            # 8. Renomear colunas
            rename_map = build_wms_rename_map(wms_cols)
            rename_map['Observação (GBS)'] = 'Observação GBS'
            df_final = df_final.rename(columns=rename_map)

            # 9. Remover colunas auxiliares
            cols_to_drop = [
                'Forn_Norm', 'Chave_NF_Forn', col_nf_wms,
                col_pedido_m2n, 'Material_Merge', 'Material',
            ]
            df_final = df_final.drop(columns=[c for c in cols_to_drop if c in df_final.columns])

            # 10. Salvar
            self.save_with_dialog(df_final)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante o processamento:\n\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AppRelatorio(root)
    root.mainloop()
