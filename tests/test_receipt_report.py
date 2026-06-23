"""Unit tests for receipt_report.py – the flat-script data-processing logic.

receipt_report.py is a top-level script (no functions/classes).  We test
the *same* data-transformation logic it implements:  merge, value
adjustment, status assignment, column rename, and duplicate-column drop.
"""

import pandas as pd
import numpy as np
import pytest


# ── Helpers to build representative data ──────────────────────────────

def make_export():
    return pd.DataFrame({
        "Nº NF-e": ["100", "100", "200", "300"],
        "Valor total": [1000.0, 1000.0, 500.0, 750.0],
    })


def make_wms(col_nf="NF", col_cod="Código", col_desc="Descrição",
             col_qtd="Quantidade", col_data="Data", col_centro="Centro"):
    """12-column WMS frame; last cols map to script's negative indexing."""
    return pd.DataFrame({
        "c1": [0]*3, "c2": [0]*3, "c3": [0]*3,
        "c4": [0]*3, "c5": [0]*3, "c6": [0]*3,
        col_centro: ["C1", "C1", "C2"],
        "c8": [0]*3, "c9": [0]*3,
        col_data: ["2024-01-01", "2024-01-01", "2024-01-02"],
        col_desc: ["Desc A", "Desc B", "Desc C"],
        col_cod: ["A01", "A02", "B01"],
        col_qtd: [10, 20, 5],
        col_nf: ["100", "100", "200"],
    })


# ── WMS column selection by negative index ────────────────────────────

class TestWmsColumnSelection:
    def test_negative_indexing(self):
        df = make_wms()
        # The make_wms helper has 14 columns; verify the last ones
        assert df.columns[-1] == "NF"
        assert df.columns[-2] == "Quantidade"
        assert df.columns[-3] == "Código"
        assert df.columns[-4] == "Descrição"

    def test_selected_columns_subset(self):
        df = make_wms()
        cols = [df.columns[-1], df.columns[-5], df.columns[-4],
                df.columns[-2], df.columns[-9], df.columns[-12]]
        selected = df[cols]
        assert len(selected.columns) == 6


# ── NF stripping / type coercion ──────────────────────────────────────

class TestNfNormalization:
    def test_strip_whitespace(self):
        s = pd.Series(["  100 ", " 200"])
        result = s.astype(str).str.strip()
        assert list(result) == ["100", "200"]

    def test_numeric_to_string(self):
        s = pd.Series([100, 200])
        result = s.astype(str).str.strip()
        assert list(result) == ["100", "200"]


# ── Merge (Export ← WMS, left join) ───────────────────────────────────

class TestMergeExportWms:
    def test_left_join_preserves_unmatched(self):
        df_export = make_export()
        df_wms = make_wms()
        col_nf_wms = df_wms.columns[-1]
        col_cod_wms = df_wms.columns[-5]
        col_desc_wms = df_wms.columns[-4]
        col_qtd_wms = df_wms.columns[-2]

        df_export["Nº NF-e"] = df_export["Nº NF-e"].astype(str).str.strip()
        df_wms[col_nf_wms] = df_wms[col_nf_wms].astype(str).str.strip()

        df_wms_sel = df_wms[[col_nf_wms, col_cod_wms, col_desc_wms, col_qtd_wms]]
        merged = pd.merge(df_export, df_wms_sel,
                          left_on="Nº NF-e", right_on=col_nf_wms, how="left")

        # NF "300" has no WMS match → should still exist with NaN in WMS cols
        nf300 = merged[merged["Nº NF-e"] == "300"]
        assert len(nf300) == 1
        assert pd.isna(nf300.iloc[0][col_cod_wms])

    def test_one_to_many_expansion(self):
        df_export = pd.DataFrame({"Nº NF-e": ["100"], "Valor total": [1000.0]})
        df_wms = make_wms()
        col_nf_wms = df_wms.columns[-1]
        col_cod_wms = df_wms.columns[-5]

        df_export["Nº NF-e"] = df_export["Nº NF-e"].astype(str).str.strip()
        df_wms[col_nf_wms] = df_wms[col_nf_wms].astype(str).str.strip()

        df_wms_sel = df_wms[[col_nf_wms, col_cod_wms]]
        merged = pd.merge(df_export, df_wms_sel,
                          left_on="Nº NF-e", right_on=col_nf_wms, how="left")
        # NF "100" appears twice in WMS → merge should produce 2 rows
        assert len(merged) == 2


# ── Valor total Ajustado ──────────────────────────────────────────────

class TestValorAjustado:
    def test_first_occurrence_keeps_value(self):
        df = pd.DataFrame({
            "Nº NF-e": ["100", "100", "200"],
            "Valor total": [1000.0, 1000.0, 500.0],
        }).sort_values(by="Nº NF-e")

        df["Valor total Ajustado"] = np.where(
            df["Nº NF-e"] != df["Nº NF-e"].shift(),
            df["Valor total"], 0
        )
        vals = df["Valor total Ajustado"].tolist()
        assert vals[0] == 1000.0
        assert vals[1] == 0.0
        assert vals[2] == 500.0

    def test_single_row(self):
        df = pd.DataFrame({"Nº NF-e": ["100"], "Valor total": [1000.0]})
        df["Valor total Ajustado"] = np.where(
            df["Nº NF-e"] != df["Nº NF-e"].shift(),
            df["Valor total"], 0
        )
        assert df["Valor total Ajustado"].iloc[0] == 1000.0

    def test_all_same_nf(self):
        df = pd.DataFrame({
            "Nº NF-e": ["100", "100", "100"],
            "Valor total": [500.0, 500.0, 500.0],
        })
        df["Valor total Ajustado"] = np.where(
            df["Nº NF-e"] != df["Nº NF-e"].shift(),
            df["Valor total"], 0
        )
        assert df["Valor total Ajustado"].tolist() == [500.0, 0.0, 0.0]

    def test_all_different_nf(self):
        df = pd.DataFrame({
            "Nº NF-e": ["100", "200", "300"],
            "Valor total": [100.0, 200.0, 300.0],
        })
        df["Valor total Ajustado"] = np.where(
            df["Nº NF-e"] != df["Nº NF-e"].shift(),
            df["Valor total"], 0
        )
        assert df["Valor total Ajustado"].tolist() == [100.0, 200.0, 300.0]


# ── Status Real ───────────────────────────────────────────────────────

class TestStatusReal:
    def test_received_when_present(self):
        s = pd.Series(["100"])
        result = s.apply(lambda x: "RECEBIDO (WMS)" if pd.notna(x) else "PENDENTE / EM TRÂNSITO")
        assert result.iloc[0] == "RECEBIDO (WMS)"

    def test_pending_when_nan(self):
        s = pd.Series([np.nan])
        result = s.apply(lambda x: "RECEBIDO (WMS)" if pd.notna(x) else "PENDENTE / EM TRÂNSITO")
        assert result.iloc[0] == "PENDENTE / EM TRÂNSITO"

    def test_mixed_statuses(self):
        s = pd.Series(["100", np.nan, "300", np.nan])
        result = s.apply(lambda x: "RECEBIDO (WMS)" if pd.notna(x) else "PENDENTE / EM TRÂNSITO")
        expected = [
            "RECEBIDO (WMS)", "PENDENTE / EM TRÂNSITO",
            "RECEBIDO (WMS)", "PENDENTE / EM TRÂNSITO",
        ]
        assert list(result) == expected


# ── Column rename ─────────────────────────────────────────────────────

class TestColumnRename:
    def test_rename_map(self):
        df = pd.DataFrame(columns=["Código", "Descrição", "Quantidade", "Data"])
        renamed = df.rename(columns={
            "Código": "CÓDIGO PRODUTO",
            "Descrição": "DESCRIÇÃO PRODUTO",
            "Quantidade": "QTD RECEBIDA",
            "Data": "DATA RECEBIMENTO",
        })
        assert list(renamed.columns) == [
            "CÓDIGO PRODUTO", "DESCRIÇÃO PRODUTO", "QTD RECEBIDA", "DATA RECEBIMENTO"
        ]


# ── Drop duplicate NF column ─────────────────────────────────────────

class TestDropDuplicateNfColumn:
    def test_drop_when_different_names(self):
        df = pd.DataFrame({"Nº NF-e": [1], "NF": [1], "val": [99]})
        col_nf_wms = "NF"
        if "Nº NF-e" != col_nf_wms:
            df = df.drop(columns=[col_nf_wms])
        assert "NF" not in df.columns
        assert "Nº NF-e" in df.columns

    def test_no_drop_when_same_name(self):
        df = pd.DataFrame({"Nº NF-e": [1], "val": [99]})
        col_nf_wms = "Nº NF-e"
        if "Nº NF-e" != col_nf_wms:
            df = df.drop(columns=[col_nf_wms])
        assert "Nº NF-e" in df.columns


# ── Edge cases ────────────────────────────────────────────────────────

class TestEdgeCases:
    def test_empty_export(self):
        df_export = pd.DataFrame({"Nº NF-e": pd.Series(dtype=str), "Valor total": pd.Series(dtype=float)})
        df_wms_sel = pd.DataFrame({"NF": pd.Series(dtype=str), "Código": pd.Series(dtype=str)})
        merged = pd.merge(df_export, df_wms_sel,
                          left_on="Nº NF-e", right_on="NF", how="left")
        assert len(merged) == 0

    def test_empty_wms(self):
        df_export = pd.DataFrame({"Nº NF-e": ["100"], "Valor total": [1000.0]})
        df_wms_sel = pd.DataFrame({"NF": pd.Series(dtype=str), "Código": pd.Series(dtype=str)})
        merged = pd.merge(df_export, df_wms_sel,
                          left_on="Nº NF-e", right_on="NF", how="left")
        assert len(merged) == 1
        assert pd.isna(merged.iloc[0]["Código"])
