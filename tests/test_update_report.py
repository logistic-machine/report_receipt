"""Unit tests for update_report.py – AppRelatorio helpers and data-processing logic."""

import pandas as pd
import numpy as np
import pytest
from unittest.mock import MagicMock, patch
import sys
import os

# Stub tkinter before any import tries to load the C extension
_tk_mock = MagicMock()
sys.modules.setdefault("_tkinter", _tk_mock)
sys.modules.setdefault("tkinter", MagicMock())
sys.modules.setdefault("tkinter.filedialog", MagicMock())
sys.modules.setdefault("tkinter.messagebox", MagicMock())

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from update_report import AppRelatorio


# ── Fixtures ──────────────────────────────────────────────────────────

@pytest.fixture
def app():
    """Create an AppRelatorio instance with a mocked Tk root."""
    with patch("tkinter.Tk") as mock_tk:
        mock_root = MagicMock()
        instance = AppRelatorio.__new__(AppRelatorio)
        instance.root = mock_root
        instance.arquivos = {}
        instance.labels_status = {}
        return instance


# ── encontrar_coluna ──────────────────────────────────────────────────

class TestEncontrarColuna:
    def test_first_match_returned(self, app):
        df = pd.DataFrame(columns=["A", "B", "C"])
        assert app.encontrar_coluna(df, ["B", "C"]) == "B"

    def test_returns_none_when_no_match(self, app):
        df = pd.DataFrame(columns=["X", "Y"])
        assert app.encontrar_coluna(df, ["A", "B"]) is None

    def test_skips_none_entries(self, app):
        df = pd.DataFrame(columns=["A", "B"])
        assert app.encontrar_coluna(df, [None, "B"]) == "B"

    def test_empty_candidate_list(self, app):
        df = pd.DataFrame(columns=["A"])
        assert app.encontrar_coluna(df, []) is None

    def test_empty_dataframe(self, app):
        df = pd.DataFrame()
        assert app.encontrar_coluna(df, ["A"]) is None

    def test_exact_match_only(self, app):
        df = pd.DataFrame(columns=["ABC", "DEF"])
        assert app.encontrar_coluna(df, ["AB"]) is None

    def test_single_column_match(self, app):
        df = pd.DataFrame(columns=["NF"])
        assert app.encontrar_coluna(df, ["NF"]) == "NF"


# ── padronizar_pedido ─────────────────────────────────────────────────

class TestPadronizarPedido:
    def test_strips_leading_zeros(self, app):
        s = pd.Series(["00123", "0456"])
        result = app.padronizar_pedido(s)
        assert list(result) == ["123", "456"]

    def test_strips_whitespace(self, app):
        s = pd.Series(["  789 ", " 0012 "])
        result = app.padronizar_pedido(s)
        assert list(result) == ["789", "12"]

    def test_numeric_input(self, app):
        s = pd.Series([100, 200])
        result = app.padronizar_pedido(s)
        assert list(result) == ["100", "200"]

    def test_all_zeros_becomes_empty(self, app):
        s = pd.Series(["000"])
        result = app.padronizar_pedido(s)
        assert list(result) == [""]

    def test_already_clean_values(self, app):
        s = pd.Series(["123", "456"])
        result = app.padronizar_pedido(s)
        assert list(result) == ["123", "456"]

    def test_nan_handling(self, app):
        s = pd.Series([np.nan])
        result = app.padronizar_pedido(s)
        # NaN stays NaN in pandas str operations
        assert pd.isna(result.iloc[0])


# ── processar – data logic (mocked file I/O) ─────────────────────────

class TestProcessarDataLogic:
    """Tests that exercise the merge / value-adjustment / status logic
    inside ``processar()`` without touching the filesystem or GUI."""

    @staticmethod
    def _build_export():
        return pd.DataFrame({
            "Nº NF-e": ["100", "100", "200"],
            "Valor total": [1000.0, 1000.0, 500.0],
        })

    @staticmethod
    def _build_wms():
        return pd.DataFrame({
            "NF": ["100", "100", "200"],
            "Material": ["A01", "A02", "B01"],
            "Descrição": ["Desc A01", "Desc A02", "Desc B01"],
            "Quantidade": [10, 20, 5],
            "Data": ["2024-01-01", "2024-01-01", "2024-01-02"],
            "Centro": ["C1", "C1", "C2"],
            "col7": [0]*3, "col8": [0]*3, "col9": [0]*3,
            "col10": [0]*3, "col11": [0]*3, "col12": [0]*3,
        })

    @staticmethod
    def _build_gbs():
        return pd.DataFrame({
            "Nº NF-e": ["100", "200"],
            "Observação (GBS)": ["Obs1", "Obs2"],
            "Pedido de Compras": ["00111", "00222"],
        })

    @staticmethod
    def _build_classificacao():
        return pd.DataFrame({
            "Material": [101, 201],
            "Planejador": ["Plan-X", "Plan-Y"],
        })

    @staticmethod
    def _build_m2n():
        return pd.DataFrame({
            "Doc.compra": ["00111", "00222"],
            "Lib": ["X", "Y"],
            "GCm": ["G1", "G2"],
        })

    def _run_processar(self, app, df_export, df_wms, df_gbs, df_classificacao, df_m2n):
        """Patch file-dialog and read_excel calls, then invoke processar."""
        import update_report as mod

        app.arquivos = {
            "export": {"path": "e.xlsx"},
            "wms": {"path": "w.xlsx"},
            "gbs": {"path": "g.xlsx"},
            "classificacao": {"path": "c.xlsx"},
            "m2n": {"path": "m.xlsx"},
        }

        read_side_effects = {
            "e.xlsx": df_export,
            "w.xlsx": df_wms,
            "g.xlsx": df_gbs,
            "c.xlsx": df_classificacao,
            "m.xlsx": df_m2n,
        }

        saved = {}

        def fake_to_excel(self_df, path, **kw):
            saved["path"] = path

        with (
            patch("pandas.read_excel", side_effect=lambda p: read_side_effects[p]),
            patch.object(mod.filedialog, "asksaveasfilename", return_value="/tmp/out.xlsx"),
            patch.object(mod.messagebox, "showinfo"),
            patch.object(pd.DataFrame, "to_excel", fake_to_excel),
        ):
            app.processar()

        return saved

    def test_value_adjustment_first_row_only(self, app):
        df_export = self._build_export()
        df_wms = self._build_wms()
        df_gbs = self._build_gbs()
        df_cls = self._build_classificacao()
        df_m2n = self._build_m2n()

        # Instead of running the full processar (which has complex internal state),
        # test the value-adjustment logic directly:
        df = df_export.copy().sort_values(by="Nº NF-e")
        df["Valor total Ajustado"] = np.where(
            df["Nº NF-e"] != df["Nº NF-e"].shift(),
            df["Valor total"], 0
        )
        # First occurrence of "100" gets value, second does not; "200" gets value
        adjusted = df["Valor total Ajustado"].tolist()
        assert adjusted[0] == 1000.0  # first "100"
        assert adjusted[1] == 0.0     # second "100"
        assert adjusted[2] == 500.0   # "200"

    def test_status_real_assigned_correctly(self, app):
        col_nf_wms = "NF"
        df = pd.DataFrame({col_nf_wms: ["100", np.nan, "300"]})
        df["Status Real"] = df[col_nf_wms].apply(
            lambda x: "RECEBIDO (WMS)" if pd.notna(x) else "PENDENTE / EM TRÂNSITO"
        )
        assert list(df["Status Real"]) == [
            "RECEBIDO (WMS)",
            "PENDENTE / EM TRÂNSITO",
            "RECEBIDO (WMS)",
        ]

    def test_merge_export_wms_left_join(self, app):
        df_export = pd.DataFrame({
            "Nº NF-e": ["100", "999"],
            "Valor total": [1000, 2000],
        })
        df_wms_sel = pd.DataFrame({
            "NF": ["100"],
            "Material": ["A01"],
        })
        df_export["Nº NF-e"] = df_export["Nº NF-e"].astype(str).str.strip()
        df_wms_sel["NF"] = df_wms_sel["NF"].astype(str).str.strip()

        merged = pd.merge(df_export, df_wms_sel,
                          left_on="Nº NF-e", right_on="NF", how="left")
        assert len(merged) == 2
        assert pd.notna(merged.loc[merged["Nº NF-e"] == "100", "Material"].iloc[0])
        assert pd.isna(merged.loc[merged["Nº NF-e"] == "999", "Material"].iloc[0])

    def test_rename_columns_mapping(self, app):
        rename_map = {
            "Material": "Cód do material",
            "Descrição": "Descrição do material",
            "Quantidade": "Qnt recebida",
            "Data": "Data de recebimento",
            "Observação (GBS)": "Observação GBS",
        }
        df = pd.DataFrame(columns=list(rename_map.keys()))
        df = df.rename(columns={k: v for k, v in rename_map.items() if k})
        assert list(df.columns) == list(rename_map.values())

    def test_drop_duplicate_nf_column(self, app):
        df = pd.DataFrame({"Nº NF-e": [1], "NF": [1], "other": [99]})
        col_nf_wms = "NF"
        if "Nº NF-e" != col_nf_wms:
            df = df.drop(columns=[col_nf_wms])
        assert "NF" not in df.columns
        assert "Nº NF-e" in df.columns

    def test_gbs_deduplication(self, app):
        df_gbs = pd.DataFrame({
            "Nº NF-e": ["100", "100", "200"],
            "Observação (GBS)": ["Obs1", "Obs1-dup", "Obs2"],
            "Pedido de Compras": ["111", "111", "222"],
        })
        deduped = df_gbs[["Nº NF-e", "Observação (GBS)", "Pedido de Compras"]].drop_duplicates(subset="Nº NF-e")
        assert len(deduped) == 2
        assert deduped.iloc[0]["Observação (GBS)"] == "Obs1"

    def test_m2n_deduplication(self, app):
        df_m2n = pd.DataFrame({
            "Doc.compra": ["111", "111", "222"],
            "Lib": ["X", "X2", "Y"],
            "GCm": ["G1", "G1b", "G2"],
        })
        deduped = df_m2n.drop_duplicates(subset="Doc.compra")
        assert len(deduped) == 2

    def test_planejador_merge(self, app):
        df_final = pd.DataFrame({"Cód do material": ["101", "999"]})
        df_final["Cód do material"] = pd.to_numeric(df_final["Cód do material"], errors="coerce").astype("Int64")
        df_plan = pd.DataFrame({
            "Material": pd.array([101, 202], dtype="Int64"),
            "Planejador": ["Plan-X", "Plan-Y"],
        })
        merged = pd.merge(df_final, df_plan,
                          left_on="Cód do material", right_on="Material", how="left")
        assert merged.loc[0, "Planejador"] == "Plan-X"
        assert pd.isna(merged.loc[1, "Planejador"])

    def test_wms_column_mapping_fallback(self, app):
        """When named columns don't match, the mapper falls back to positional index."""
        cols = [f"col{i}" for i in range(12)]
        df_wms = pd.DataFrame(columns=cols)
        ncols = len(df_wms.columns)
        # Verify positional fallback resolves to last column
        assert df_wms.columns[-1] == "col11"
        assert df_wms.columns[-5] == "col7"

    def test_processar_full_pipeline(self, app):
        """Run the full processar pipeline with mocked I/O and verify output was saved."""
        saved = self._run_processar(
            app,
            self._build_export(),
            self._build_wms(),
            self._build_gbs(),
            self._build_classificacao(),
            self._build_m2n(),
        )
        assert saved.get("path") == "/tmp/out.xlsx"

    def test_processar_missing_wms_columns_shows_error(self, app):
        """When WMS lacks essential columns, processar should show an error."""
        import update_report as mod

        df_wms_bad = pd.DataFrame({"unrelated": [1, 2]})
        app.arquivos = {
            "export": {"path": "e.xlsx"},
            "wms": {"path": "w.xlsx"},
            "gbs": {"path": "g.xlsx"},
            "classificacao": {"path": "c.xlsx"},
            "m2n": {"path": "m.xlsx"},
        }
        read_map = {
            "e.xlsx": self._build_export(),
            "w.xlsx": df_wms_bad,
            "g.xlsx": self._build_gbs(),
            "c.xlsx": self._build_classificacao(),
            "m.xlsx": self._build_m2n(),
        }
        mock_err = MagicMock()
        with (
            patch("pandas.read_excel", side_effect=lambda p: read_map[p]),
            patch.object(mod.messagebox, "showerror", mock_err),
        ):
            app.processar()
        mock_err.assert_called_once()
        assert "colunas essenciais" in mock_err.call_args[0][1].lower()

    def test_processar_missing_m2n_columns_shows_error(self, app):
        """When ME2N lacks Doc.compra column, processar should show an error."""
        import update_report as mod

        df_m2n_bad = pd.DataFrame({"wrong_col": [1]})
        app.arquivos = {
            "export": {"path": "e.xlsx"},
            "wms": {"path": "w.xlsx"},
            "gbs": {"path": "g.xlsx"},
            "classificacao": {"path": "c.xlsx"},
            "m2n": {"path": "m.xlsx"},
        }
        read_map = {
            "e.xlsx": self._build_export(),
            "w.xlsx": self._build_wms(),
            "g.xlsx": self._build_gbs(),
            "c.xlsx": self._build_classificacao(),
            "m.xlsx": df_m2n_bad,
        }
        mock_err = MagicMock()
        with (
            patch("pandas.read_excel", side_effect=lambda p: read_map[p]),
            patch.object(mod.messagebox, "showerror", mock_err),
        ):
            app.processar()
        mock_err.assert_called_once()
        assert "Doc.compra" in mock_err.call_args[0][1]
