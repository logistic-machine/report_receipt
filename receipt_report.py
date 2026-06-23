import pandas as pd
import numpy as np

from shared.utils import normalize_nf_column, compute_adjusted_total, compute_status_real

df_export = pd.read_excel
df_wms = pd.read_excel

col_nf_wms = df_wms.columns[-1]
col_qtd_wms = df_wms.columns[-2]
col_desc_wms = df_wms.columns[-4]
col_cod_wms = df_wms.columns[-5]
col_data_wms = df_wms.columns[-9]
col_centro_wms = df_wms.columns[-12]

normalize_nf_column(df_export, 'Nº NF-e')
normalize_nf_column(df_wms, col_nf_wms)

df_wms_selecionado = df_wms[[col_nf_wms, col_cod_wms, col_desc_wms, col_qtd_wms, col_data_wms, col_centro_wms]]
df_final = pd.merge(df_export, df_wms_selecionado, left_on='Nº NF-e', right_on=col_nf_wms, how='left')

df_final = df_final.sort_values(by='Nº NF-e')

df_final = compute_adjusted_total(df_final, group_col='Nº NF-e')
df_final = compute_status_real(df_final, col_nf_wms)

df_final = df_final.rename(columns={
    col_cod_wms: 'CÓDIGO PRODUTO',
    col_desc_wms: 'DESCRIÇÃO PRODUTO',
    col_qtd_wms: 'QTD RECEBIDA',
    col_data_wms: 'DATA RECEBIMENTO',
})

if 'Nº NF-e' != col_nf_wms:
    df_final = df_final.drop(columns=[col_nf_wms])

df_final.to_excel('Recebimento_técnico.xlsx', index=False)

print("Relatório gerado! O valor total agora só aparece na primeira linha de cada nota.")
