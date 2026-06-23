import pandas as pd
import numpy as np


def encontrar_coluna(df, possiveis_nomes):
    """Return the first matching column name found in the DataFrame."""
    for nome in possiveis_nomes:
        if nome and nome in df.columns:
            return nome
    return None


def padronizar_pedido(serie):
    """Convert to string, strip whitespace, and remove leading zeros."""
    return serie.astype(str).str.strip().str.lstrip('0')


def normalizar_fornecedor(nome):
    """Extract the first word of the supplier name as a cross-reference key."""
    if pd.isna(nome):
        return "DESCONHECIDO"
    nome_str = str(nome).strip().upper()
    return nome_str.split()[0] if nome_str.split() else "DESCONHECIDO"


def normalize_nf_column(df, col_name):
    """Cast a NF column to string and strip whitespace in place."""
    df[col_name] = df[col_name].astype(str).str.strip()
    return df


def build_wms_column_mapping(df_wms):
    """Build the standard WMS column-name mapping with positional fallbacks."""
    ncols = len(df_wms.columns)
    return {
        'nf':     ['NF', 'Nota Fiscal', 'N\u00ba NF-e', 'Nf',
                   df_wms.columns[-1] if ncols >= 1 else None],
        'qtd':    ['Quantidade', 'Qtd', 'Qnt recebida', 'Qtd.',
                   df_wms.columns[-2] if ncols >= 2 else None],
        'desc':   ['Descri\u00e7\u00e3o', 'Desc. Material', 'Descri\u00e7\u00e3o do material',
                   df_wms.columns[-4] if ncols >= 4 else None],
        'cod':    ['C\u00f3digo', 'C\u00f3d. Material', 'C\u00f3d do material', 'Material',
                   df_wms.columns[-5] if ncols >= 5 else None],
        'data':   ['Data', 'Data de recebimento', 'Data Recebimento',
                   df_wms.columns[-9] if ncols >= 9 else None],
        'centro': ['Centro', 'Dep\u00f3sito', 'Unidade',
                   df_wms.columns[-12] if ncols >= 12 else None],
    }


def resolve_wms_columns(df_wms, mapping=None):
    """Resolve WMS column names. Returns ``{key: resolved_name}``."""
    if mapping is None:
        mapping = build_wms_column_mapping(df_wms)
    return {
        key: encontrar_coluna(df_wms, candidates)
        for key, candidates in mapping.items()
    }


def compute_adjusted_total(df, group_col, value_col='Valor total'):
    """Show the total value only on the first row of each group."""
    df['Valor total Ajustado'] = np.where(
        df[group_col] != df[group_col].shift(),
        df[value_col], 0
    )
    return df


def compute_status_real(df, col_nf_wms):
    """Add 'Status Real' column based on whether the WMS NF column has data."""
    df['Status Real'] = np.where(
        df[col_nf_wms].notna(),
        "RECEBIDO (WMS)",
        "PENDENTE / EM TR\u00c2NSITO"
    )
    return df


def build_wms_rename_map(wms_cols):
    """Build standard rename mapping for resolved WMS columns."""
    mapping = {}
    if wms_cols.get('cod'):
        mapping[wms_cols['cod']] = 'C\u00f3d do material'
    if wms_cols.get('desc'):
        mapping[wms_cols['desc']] = 'Descri\u00e7\u00e3o do material'
    if wms_cols.get('qtd'):
        mapping[wms_cols['qtd']] = 'Qnt recebida'
    if wms_cols.get('data'):
        mapping[wms_cols['data']] = 'Data de recebimento'
    return mapping


def merge_planejador(df_final, df_classificacao, material_col):
    """Merge 'Planejador' into *df_final* by matching material codes as integers."""
    df_planejador = (
        df_classificacao[['Material', 'Planejador']]
        .drop_duplicates(subset='Material')
    )
    df_final[material_col] = (
        pd.to_numeric(df_final[material_col], errors='coerce').astype('Int64')
    )
    df_planejador['Material'] = (
        pd.to_numeric(df_planejador['Material'], errors='coerce').astype('Int64')
    )
    df_final = pd.merge(
        df_final, df_planejador,
        left_on=material_col, right_on='Material', how='left'
    )
    if material_col != 'Material' and 'Material' in df_final.columns:
        df_final = df_final.drop(columns=['Material'])
    return df_final


def add_supplier_key_columns(df, nf_col, fornecedor_candidates):
    """Add 'Forn_Norm' and 'Chave_NF_Forn' composite-key columns."""
    col_forn = encontrar_coluna(df, fornecedor_candidates)
    if col_forn:
        df['Forn_Norm'] = df[col_forn].apply(normalizar_fornecedor)
    else:
        df['Forn_Norm'] = "SEM_FORN"
    df['Chave_NF_Forn'] = df[nf_col].astype(str).str.strip() + "-" + df['Forn_Norm']
    return df
