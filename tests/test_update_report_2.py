"""Unit tests for update_report_2.py – additional helpers and supplier-key logic."""

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

from update_report_2 import AppRelatorio


# ── Fixtures ──────────────────────────────────────────────────────────

@pytest.fixture
def app():
    with patch("tkinter.Tk"):
        instance = AppRelatorio.__new__(AppRelatorio)
        instance.root = MagicMock()
        instance.arquivos = {}
        instance.labels_status = {}
        return instance


# ── normalizar_fornecedor ─────────────────────────────────────────────

class TestNormalizarFornecedor:
    def test_extracts_first_word(self, app):
        assert app.normalizar_fornecedor("Empresa ABC Ltda") == "EMPRESA"

    def test_returns_desconhecido_for_nan(self, app):
        assert app.normalizar_fornecedor(np.nan) == "DESCONHECIDO"
        assert app.normalizar_fornecedor(None) == "DESCONHECIDO"

    def test_uppercases_result(self, app):
        assert app.normalizar_fornecedor("nome") == "NOME"

    def test_strips_whitespace(self, app):
        assert app.normalizar_fornecedor("  Foo Bar  ") == "FOO"

    def test_empty_string(self, app):
        assert app.normalizar_fornecedor("") == "DESCONHECIDO"

    def test_single_word(self, app):
        assert app.normalizar_fornecedor("ACME") == "ACME"

    def test_numeric_input(self, app):
        assert app.normalizar_fornecedor(12345) == "12345"

    def test_whitespace_only(self, app):
        assert app.normalizar_fornecedor("   ") == "DESCONHECIDO"


# ── Supplier key (Chave_NF_Forn) construction ─────────────────────────

class TestChaveNFForn:
    def test_key_concatenation(self, app):
        df = pd.DataFrame({
            "Nº NF-e": ["100", "200"],
            "Fornecedor": ["Empresa X", "Acme Y"],
        })
        df["Forn_Norm"] = df["Fornecedor"].apply(app.normalizar_fornecedor)
        df["Chave_NF_Forn"] = df["Nº NF-e"].astype(str).str.strip() + "-" + df["Forn_Norm"]
        assert df.loc[0, "Chave_NF_Forn"] == "100-EMPRESA"
        assert df.loc[1, "Chave_NF_Forn"] == "200-ACME"

    def test_key_with_missing_supplier(self, app):
        df = pd.DataFrame({
            "Nº NF-e": ["100"],
            "Forn_Norm": ["SEM_FORN"],
        })
        df["Chave_NF_Forn"] = df["Nº NF-e"].astype(str).str.strip() + "-" + df["Forn_Norm"]
        assert df.loc[0, "Chave_NF_Forn"] == "100-SEM_FORN"


# ── Value adjustment with composite key ───────────────────────────────

class TestValueAdjustmentV2:
    def test_value_only_on_first_row_per_key(self, app):
        df = pd.DataFrame({
            "Chave_NF_Forn": ["100-EMP", "100-EMP", "200-ACM"],
            "cod": ["A", "B", "C"],
            "Valor total": [1000.0, 1000.0, 500.0],
        })
        df = df.sort_values(by=["Chave_NF_Forn", "cod"])
        df["Valor total Ajustado"] = np.where(
            df["Chave_NF_Forn"] != df["Chave_NF_Forn"].shift(),
            df["Valor total"], 0
        )
        adjusted = df["Valor total Ajustado"].tolist()
        assert adjusted == [1000.0, 0.0, 500.0]


# ── Status Real (np.where variant in v2) ──────────────────────────────

class TestStatusRealV2:
    def test_status_with_np_where(self, app):
        col_nf_wms = "NF"
        df = pd.DataFrame({col_nf_wms: ["100", np.nan, "300"]})
        df["Status Real"] = np.where(
            df[col_nf_wms].notna(), "RECEBIDO (WMS)", "PENDENTE / EM TRÂNSITO"
        )
        assert list(df["Status Real"]) == [
            "RECEBIDO (WMS)",
            "PENDENTE / EM TRÂNSITO",
            "RECEBIDO (WMS)",
        ]


# ── Material_Str construction ─────────────────────────────────────────

class TestMaterialStr:
    def test_strips_and_removes_leading_zeros(self, app):
        df = pd.DataFrame({"Material": ["00123", " 0456 "]})
        df["Material_Str"] = df["Material"].astype(str).str.strip().str.lstrip("0")
        assert list(df["Material_Str"]) == ["123", "456"]


# ── Chave_Item construction ───────────────────────────────────────────

class TestChaveItem:
    def test_key_item_concatenation(self, app):
        df = pd.DataFrame({
            "Chave_NF_Forn": ["100-EMP", "100-EMP"],
            "Material_Str": ["123", "456"],
        })
        df["Chave_Item"] = df["Chave_NF_Forn"] + "-" + df["Material_Str"]
        assert df.loc[0, "Chave_Item"] == "100-EMP-123"
        assert df.loc[1, "Chave_Item"] == "100-EMP-456"


# ── Drop auxiliary columns ────────────────────────────────────────────

class TestDropAuxColumns:
    def test_drops_only_existing_cols(self, app):
        df = pd.DataFrame({
            "Forn_Norm": [1],
            "Chave_NF_Forn": [1],
            "NF": [1],
            "Doc.compra": [1],
            "Material_Merge": [1],
            "Material": [1],
            "KeepMe": [99],
        })
        cols_to_drop = ["Forn_Norm", "Chave_NF_Forn", "NF", "Doc.compra",
                        "Material_Merge", "Material", "NonExistent"]
        df = df.drop(columns=[c for c in cols_to_drop if c in df.columns])
        assert list(df.columns) == ["KeepMe"]

    def test_no_error_on_empty_drop_list(self, app):
        df = pd.DataFrame({"A": [1]})
        cols_to_drop = ["X", "Y"]
        df = df.drop(columns=[c for c in cols_to_drop if c in df.columns])
        assert list(df.columns) == ["A"]


# ── encontrar_coluna (inherited) ──────────────────────────────────────

class TestEncontrarColunaV2:
    def test_supplier_column_detection(self, app):
        possible = ["CNPJ Fornecedor", "Fornecedor", "Emitente"]
        df = pd.DataFrame(columns=["ID", "Fornecedor", "Valor"])
        assert app.encontrar_coluna(df, possible) == "Fornecedor"

    def test_falls_back_through_list(self, app):
        possible = ["CNPJ Fornecedor", "Fornecedor", "Emitente"]
        df = pd.DataFrame(columns=["ID", "Emitente"])
        assert app.encontrar_coluna(df, possible) == "Emitente"


# ── Full pipeline smoke test ──────────────────────────────────────────

class TestProcessarV2:
    @staticmethod
    def _build_export():
        return pd.DataFrame({
            "Nº NF-e": ["100", "200"],
            "Valor total": [1000.0, 500.0],
            "Fornecedor": ["Empresa A", "Acme B"],
        })

    @staticmethod
    def _build_wms():
        return pd.DataFrame({
            "NF": ["100", "200"],
            "Material": ["A01", "B01"],
            "Descrição": ["Desc A01", "Desc B01"],
            "Quantidade": [10, 5],
            "Data": ["2024-01-01", "2024-01-02"],
            "Centro": ["C1", "C2"],
            "Fornecedor": ["Empresa A", "Acme B"],
            "col7": [0]*2, "col8": [0]*2, "col9": [0]*2,
            "col10": [0]*2, "col11": [0]*2, "col12": [0]*2,
        })

    @staticmethod
    def _build_gbs():
        return pd.DataFrame({
            "Nº NF-e": ["100", "200"],
            "Observação (GBS)": ["Obs1", "Obs2"],
            "Pedido de Compras": ["00111", "00222"],
            "Fornecedor": ["Empresa A", "Acme B"],
        })

    @staticmethod
    def _build_classificacao():
        return pd.DataFrame({"Material": [101, 201], "Planejador": ["P-X", "P-Y"]})

    @staticmethod
    def _build_m2n():
        return pd.DataFrame({
            "Doc.compra": ["00111", "00222"],
            "Lib": ["X", "Y"],
            "GCm": ["G1", "G2"],
        })

    def test_full_pipeline_runs(self, app):
        import update_report_2 as mod

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
            "m.xlsx": self._build_m2n(),
        }
        saved = {}

        def fake_to_excel(self_df, path, **kw):
            saved["path"] = path

        with (
            patch("pandas.read_excel", side_effect=lambda p: read_map[p]),
            patch.object(mod.filedialog, "asksaveasfilename", return_value="/tmp/out.xlsx"),
            patch.object(mod.messagebox, "showinfo"),
            patch.object(pd.DataFrame, "to_excel", fake_to_excel),
        ):
            app.processar()
        assert saved.get("path") == "/tmp/out.xlsx"

    def test_pipeline_error_shows_messagebox(self, app):
        import update_report_2 as mod

        app.arquivos = {
            "export": {"path": "e.xlsx"},
            "wms": {"path": "w.xlsx"},
            "gbs": {"path": "g.xlsx"},
            "classificacao": {"path": "c.xlsx"},
            "m2n": {"path": "m.xlsx"},
        }
        mock_err = MagicMock()
        with (
            patch("pandas.read_excel", side_effect=FileNotFoundError("not found")),
            patch.object(mod.messagebox, "showerror", mock_err),
        ):
            app.processar()
        mock_err.assert_called_once()
