"""
inferencia.py — Uso do classificador de falhas (TED 03) no protótipo do MVP.

Dado o sintoma relatado e o tipo de aparelho, devolve as k causas mais prováveis,
com probabilidade, e um indicador de confiança. O limiar (0,35) vem da curva
acurácia × cobertura da Seção 19.4 do Documento Técnico v3.0 (≈80% de acerto top-1
em ≈78% dos casos). Fora dele, o assistente deve pedir mais informações ao técnico.

Exemplo:
    from inferencia import sugerir_causas
    sugerir_causas("won't turn on, no charging light", "Notebook")
"""
from __future__ import annotations

from functools import lru_cache
from pathlib import Path

import joblib
import numpy as np
import pandas as pd

from modelagem import COL_APARELHO, COL_TEXTO, normalizar_texto

RAIZ = Path(__file__).resolve().parents[1]
MODELO_PADRAO = RAIZ / "models" / "classificador_falhas.joblib"
LIMIAR_CONFIANCA = 0.35
APARELHOS = ("Celular", "Notebook", "Tablet", "Desktop")


@lru_cache(maxsize=2)
def _carregar(caminho: str):
    return joblib.load(caminho)


def _cobertura_vocabulario(modelo, texto_norm: str) -> float:
    """Fração das palavras do texto que o modelo conhece (vocabulário TF-IDF de palavras)."""
    vocab = modelo.named_steps["prep"].named_transformers_["w"].vocabulary_
    palavras = texto_norm.split()
    return sum(p in vocab for p in palavras) / len(palavras) if palavras else 0.0


def sugerir_causas(texto: str, aparelho: str, k: int = 3, modelo=None, limiar: float = LIMIAR_CONFIANCA) -> dict:
    """Devolve as k causas mais prováveis para (texto, aparelho).

    Levanta ValueError para texto vazio/sem letras, aparelho fora de APARELHOS
    (use "Desconhecido" se não souber) ou k < 1. Se nenhuma palavra do texto é
    conhecida pelo modelo, a confiança é marcada como baixa, mesmo que a
    probabilidade do 1º lugar passe do limiar.
    """
    if not isinstance(texto, str) or not any(c.isalpha() for c in texto):
        raise ValueError("texto deve ser uma string com o sintoma descrito (ao menos uma letra).")
    if aparelho not in APARELHOS and aparelho != "Desconhecido":
        raise ValueError(f"aparelho inválido: {aparelho!r}. Use um de {APARELHOS} ou 'Desconhecido'.")
    if k < 1:
        raise ValueError("k deve ser >= 1.")
    modelo = modelo or _carregar(str(MODELO_PADRAO))
    texto_norm = normalizar_texto(texto)
    if not texto_norm:
        raise ValueError("texto sem conteúdo útil após a normalização.")
    X = pd.DataFrame({COL_TEXTO: [texto_norm], COL_APARELHO: [aparelho]})
    proba = modelo.predict_proba(X)[0]
    ordem = np.argsort(-proba)[:k]
    cobertura = _cobertura_vocabulario(modelo, texto_norm)
    alta = bool(proba.max() >= limiar and cobertura > 0)
    return {
        "sugestoes": [{"categoria_falha": str(modelo.classes_[i]), "probabilidade": round(float(proba[i]), 3)} for i in ordem],
        "confianca_alta": alta,
        "cobertura_vocabulario": round(cobertura, 2),
        "acao_recomendada": ("Iniciar pelos testes da 1ª causa e seguir a ordem sugerida."
                             if alta else
                             "Confiança baixa: fazer perguntas complementares antes de sugerir testes."),
    }
