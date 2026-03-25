import pandas as pd
import numpy as np

df_export = pd.read_excel
df_wms = pd.read_excel

col_nf_wms = df_wms.columns[-1]
col_qtd_wms = df_wms.columns[-2]
col_desc_wms = df_wms.columns[-4]
col_cod_wms = df_wms.columns[-5]
col_data_wms = df_wms.columns[-9] 
col_centro_wms = df_wms.columns[-12]

df_export['Nº NF-e'] = df_export['Nº NF-e'].astype(str).str.strip()
df_wms[col_nf_wms] = df_wms[col_nf_wms].astype(str).str.strip()

df_wms_selecionado = df_wms[[col_nf_wms, col_cod_wms, col_desc_wms, col_qtd_wms, col_data_wms, col_centro_wms]]
df_final = pd.merge(df_export, df_wms_selecionado, left_on='Nº NF-e', right_on=col_nf_wms, how='left')

df_final = df_final.sort_values(by='Nº NF-e')

df_final['Valor total Ajustado'] = np.where(df_final['Nº NF-e'] != df_final['Nº NF-e'].shift(), df_final['Valor total'], 0)

df_final['Status Real'] = df_final[col_nf_wms].apply(
    lambda x: "RECEBIDO (WMS)" if pd.notna(x) else "PENDENTE / EM TRÂNSITO"
)

df_final = df_final.rename(columns={
    col_cod_wms: 'CÓDIGO PRODUTO',
    col_desc_wms: 'DESCRIÇÃO PRODUTO',
    col_qtd_wms: 'QTD RECEBIDA',
    col_data_wms: 'DATA RECEBIMENTO',
})

if 'Nº NF-e' != col_nf_wms:
    df_final = df_final.drop(columns=[col_nf_wms])

df_final.to_excel('Recebimento_técnico.xlsx', index=False)

print("✅ Relatório gerado! O valor total agora só aparece na primeira linha de cada nota.")
