"""Testes dos dados e da divisão (TED 03). Rodar: pytest -q"""
import json
from pathlib import Path

import pandas as pd
import pytest

from modelagem import (ALVO, ANO_CORTE, COL_APARELHO, COL_TEXTO, PriorPorAparelho,
                       carregar_dados, catalogo_modelos, dividir_temporal, normalizar_texto,
                       topk_acuracia)

RAIZ = Path(__file__).resolve().parents[1]


@pytest.fixture(scope="module")
def df():
    return carregar_dados()


def test_normalizar_texto():
    assert normalizar_texto("  Won't TURN on!!  123 http://x.com ") == "won't turn on"
    assert normalizar_texto("Tela não liga — ação") == "tela não liga ação"
    assert normalizar_texto("") == ""


def test_dados_basicos(df):
    assert len(df) == 4019
    assert df["id_origem"].is_unique
    assert df[ALVO].nunique() == 13 and "Indeterminado" not in set(df[ALVO])
    assert df[COL_TEXTO].str.len().gt(0).all()
    assert df[COL_APARELHO].notna().all()


def test_divisao_temporal_sem_vazamento(df):
    tr, te = dividir_temporal(df)
    assert (len(tr), len(te)) == (2224, 1795)
    assert tr["ano_evento"].max() <= ANO_CORTE < te["ano_evento"].min()
    assert set(tr["id_origem"]).isdisjoint(te["id_origem"])
    assert len(tr) + len(te) == len(df)


def test_divisao_salva_confere_com_codigo(df):
    sal = pd.read_csv(RAIZ / "data" / "processed" / "divisao_treino_teste_ted03.csv")
    tr, te = dividir_temporal(df)
    col = [c for c in sal.columns if "divis" in c.lower() or c.lower() in ("conjunto", "split")][0]
    assert set(sal.loc[sal[col].astype(str).str.lower().str.startswith("trein"), "id_origem"]) == set(tr["id_origem"])
    assert set(sal.loc[sal[col].astype(str).str.lower().str.startswith("test"), "id_origem"]) == set(te["id_origem"])


def test_baseline_prior_probabilidades_validas(df):
    tr, te = dividir_temporal(df)
    m = PriorPorAparelho().fit(tr, tr[ALVO])
    p = m.predict_proba(te)
    assert p.shape == (len(te), 13)
    assert abs(p.sum(axis=1) - 1).max() < 1e-9


def test_topk_monotono():
    import numpy as np
    y = ["a", "b", "c"]; classes = ["a", "b", "c"]
    s = np.array([[.5, .3, .2], [.5, .3, .2], [.1, .2, .7]])
    assert topk_acuracia(y, s, classes, 1) == pytest.approx(2 / 3)
    assert topk_acuracia(y, s, classes, 2) == pytest.approx(1.0)
    assert topk_acuracia(y, s, classes, 1) <= topk_acuracia(y, s, classes, 2) <= topk_acuracia(y, s, classes, 3)


def test_catalogo_tem_modelos_esperados():
    assert len(catalogo_modelos()) >= 6


def test_resultados_json_coerentes():
    r = json.loads((RAIZ / "reports" / "resultados_modelagem.json").read_text(encoding="utf8"))
    assert r  # não vazio
