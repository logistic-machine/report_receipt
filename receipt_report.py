import sys
import pandas as pd
import numpy as np


def carregar_arquivo(caminho, descricao):
    """Load an Excel file and raise a clear error if it fails."""
    try:
        df = pd.read_excel(caminho)
    except FileNotFoundError:
        raise FileNotFoundError(
            f"Arquivo '{descricao}' nao encontrado: {caminho}"
        )
    except Exception as e:
        raise RuntimeError(
            f"Erro ao ler arquivo '{descricao}' ({caminho}): {e}"
        ) from e

    if df.empty:
        raise ValueError(f"Arquivo '{descricao}' esta vazio: {caminho}")
    return df


def validar_colunas_minimas(df, colunas_requeridas, descricao):
    """Ensure all required columns exist in the DataFrame."""
    faltando = [c for c in colunas_requeridas if c not in df.columns]
    if faltando:
        raise KeyError(
            f"Colunas obrigatorias ausentes no arquivo '{descricao}': {faltando}. "
            f"Colunas disponiveis: {list(df.columns)}"
        )


def obter_coluna_por_indice(df, indice_negativo, descricao_coluna):
    """Safely get a column by negative index, raising a clear error if out of range."""
    ncols = len(df.columns)
    if ncols < abs(indice_negativo):
        raise IndexError(
            f"Arquivo WMS tem apenas {ncols} colunas, mas a coluna "
            f"'{descricao_coluna}' requer pelo menos {abs(indice_negativo)} colunas."
        )
    return df.columns[indice_negativo]


def gerar_relatorio(caminho_export, caminho_wms, caminho_saida="Recebimento_tecnico.xlsx"):
    """Main report generation with proper error handling and validation."""

    # 1. Load files
    df_export = carregar_arquivo(caminho_export, "Export")
    df_wms = carregar_arquivo(caminho_wms, "WMS")

    # 2. Validate Export has required columns
    validar_colunas_minimas(df_export, ['Nº NF-e', 'Valor total'], "Export")

    # 3. Map WMS columns by position with validation
    col_nf_wms = obter_coluna_por_indice(df_wms, -1, "NF")
    col_qtd_wms = obter_coluna_por_indice(df_wms, -2, "Quantidade")
    col_desc_wms = obter_coluna_por_indice(df_wms, -4, "Descricao")
    col_cod_wms = obter_coluna_por_indice(df_wms, -5, "Codigo")
    col_data_wms = obter_coluna_por_indice(df_wms, -9, "Data")
    col_centro_wms = obter_coluna_por_indice(df_wms, -12, "Centro")

    # 4. Standardize NF columns for matching
    df_export['Nº NF-e'] = df_export['Nº NF-e'].astype(str).str.strip()
    df_wms[col_nf_wms] = df_wms[col_nf_wms].astype(str).str.strip()

    # 5. Select and merge
    df_wms_selecionado = df_wms[[col_nf_wms, col_cod_wms, col_desc_wms, col_qtd_wms, col_data_wms, col_centro_wms]]
    df_final = pd.merge(df_export, df_wms_selecionado, left_on='Nº NF-e', right_on=col_nf_wms, how='left')

    df_final = df_final.sort_values(by='Nº NF-e')

    # 6. Adjusted total value (only first row per NF)
    df_final['Valor total Ajustado'] = np.where(
        df_final['Nº NF-e'] != df_final['Nº NF-e'].shift(),
        df_final['Valor total'], 0
    )

    # 7. Status based on WMS match
    df_final['Status Real'] = df_final[col_nf_wms].apply(
        lambda x: "RECEBIDO (WMS)" if pd.notna(x) else "PENDENTE / EM TRANSITO"
    )

    # 8. Rename columns
    df_final = df_final.rename(columns={
        col_cod_wms: 'CODIGO PRODUTO',
        col_desc_wms: 'DESCRICAO PRODUTO',
        col_qtd_wms: 'QTD RECEBIDA',
        col_data_wms: 'DATA RECEBIMENTO',
    })

    # 9. Drop duplicate NF column if distinct from Export's
    if 'Nº NF-e' != col_nf_wms:
        df_final = df_final.drop(columns=[col_nf_wms])

    # 10. Save output
    try:
        df_final.to_excel(caminho_saida, index=False)
    except PermissionError:
        raise PermissionError(
            f"Sem permissao para salvar em '{caminho_saida}'. "
            f"Verifique se o arquivo esta aberto em outro programa."
        )
    except Exception as e:
        raise RuntimeError(
            f"Erro ao salvar relatorio em '{caminho_saida}': {e}"
        ) from e

    total = len(df_final)
    recebidos = df_final['Status Real'].eq("RECEBIDO (WMS)").sum()
    print(f"Relatorio gerado: {caminho_saida}")
    print(f"  Total de linhas: {total}")
    print(f"  Recebidos (WMS): {recebidos}")
    print(f"  Pendentes:       {total - recebidos}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(
            "Uso: python receipt_report.py <arquivo_export.xlsx> <arquivo_wms.xlsx> "
            "[arquivo_saida.xlsx]",
            file=sys.stderr,
        )
        sys.exit(1)

    caminho_export = sys.argv[1]
    caminho_wms = sys.argv[2]
    caminho_saida = sys.argv[3] if len(sys.argv) > 3 else "Recebimento_tecnico.xlsx"

    try:
        gerar_relatorio(caminho_export, caminho_wms, caminho_saida)
    except (FileNotFoundError, ValueError, KeyError, IndexError, PermissionError, RuntimeError) as e:
        print(f"ERRO: {e}", file=sys.stderr)
        sys.exit(1)
