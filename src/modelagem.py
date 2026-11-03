"""
modelagem.py — Funções de preparação, treinamento e avaliação da TED 03
Projeto: Assistente de Diagnóstico Inteligente da Inforcell (UNIBALSAS)
Disciplina: Big Data e Ciência de Dados — TED 03 (Modelagem Preditiva e Machine Learning)

Problema de ML: classificação multiclasse de texto
    sintoma relatado (texto livre) [+ categoria do aparelho]  ->  categoria de falha (taxonomia unificada)

As decisões implementadas aqui estão justificadas nas Seções 15 a 21 do
Documento Técnico v3.0 (docs/TED03.pdf / TED03.docx).
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.base import BaseEstimator, ClassifierMixin
from sklearn.compose import ColumnTransformer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (accuracy_score, balanced_accuracy_score, f1_score,
                             precision_score, recall_score)
from sklearn.model_selection import (GridSearchCV, StratifiedGroupKFold,
                                     train_test_split)
from sklearn.naive_bayes import ComplementNB
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder
from sklearn.svm import LinearSVC

RAIZ = Path(__file__).resolve().parents[1]
PROCESSED = RAIZ / "data" / "processed"

SEED = 42
ANO_CORTE = 2018          # treino: eventos até 2018; teste: 2019 e 2020
ALVO = "categoria_falha"
COL_TEXTO = "texto"
COL_APARELHO = "categoria"

# mesmo padrão usado na EDA (Seção 11.5 do documento v2.0) para o sintoma "não liga"
PADRAO_NAO_LIGA = re.compile(
    r"(?:won'?t|will not|does ?n[o']?t|doesnt|not|no|cannot|can'?t)\s+(?:turn|switch|power|boot|start)(?:ing|s)?\s*(?:on|up)?|dead|no power|"
    r"gaat niet aan|start niet|geht nicht (?:mehr )?an|startet nicht|ne s'?allume|ne démarre", re.I)


# --------------------------------------------------------------------------- dados
def normalizar_texto(t: str) -> str:
    """Minúsculas, remove números e pontuação, preserva letras acentuadas (multilíngue)."""
    t = unicodedata.normalize("NFKC", str(t)).lower()
    t = re.sub(r"https?://\S+", " ", t)
    t = re.sub(r"[^\w\s']", " ", t)
    t = re.sub(r"\d+", " ", t).replace("_", " ")
    return re.sub(r"\s+", " ", t).strip()


def carregar_dados() -> pd.DataFrame:
    """Une falhas_modelagem.csv (TED 02) à data do evento e prepara as colunas de modelagem."""
    mod = pd.read_csv(PROCESSED / "falhas_modelagem.csv")
    base_b = pd.read_csv(PROCESSED / "falhas_rotuladas_limpo.csv",
                         usecols=["id_origem", "quest", "fonte", "data_evento"])
    df = mod.merge(base_b, on="id_origem", how="left", validate="one_to_one")
    df["ano_evento"] = pd.to_datetime(df["data_evento"]).dt.year
    df[COL_APARELHO] = df[COL_APARELHO].fillna("Desconhecido")   # 1 registro sem categoria
    df[COL_TEXTO] = df["problema"].map(normalizar_texto)
    df["n_palavras"] = df[COL_TEXTO].str.split().str.len()
    return df.reset_index(drop=True)


def dividir_temporal(df: pd.DataFrame, ano_corte: int = ANO_CORTE):
    """Treino = eventos até ano_corte; teste = eventos posteriores (respeita o tempo)."""
    tr = df[df["ano_evento"] <= ano_corte].copy()
    te = df[df["ano_evento"] > ano_corte].copy()
    return tr.reset_index(drop=True), te.reset_index(drop=True)


def dividir_aleatorio(df: pd.DataFrame, frac_teste: float, seed: int = SEED):
    """Divisão estratificada pelo rótulo (referência para medir o efeito da deriva temporal)."""
    tr, te = train_test_split(df, test_size=frac_teste, stratify=df[ALVO], random_state=seed)
    return tr.reset_index(drop=True), te.reset_index(drop=True)


# --------------------------------------------------------------------------- baselines
class PriorPorAparelho(BaseEstimator, ClassifierMixin):
    """Linha de base: probabilidade a priori da falha dado só o tipo de aparelho
    (é a ideia de 'probabilidade inicial' proposta na Seção 11.5 da TED 02)."""

    def __init__(self, usar_aparelho: bool = True):
        self.usar_aparelho = usar_aparelho

    def fit(self, X, y):
        y = np.asarray(y)
        self.classes_ = np.unique(y)
        glob = pd.Series(y).value_counts().reindex(self.classes_).fillna(0).values + 1.0
        self.prior_ = glob / glob.sum()
        self.por_aparelho_ = {}
        if self.usar_aparelho:
            for ap, idx in pd.Series(np.arange(len(y))).groupby(X[COL_APARELHO].values):
                c = pd.Series(y[idx.values]).value_counts().reindex(self.classes_).fillna(0).values + 1.0
                self.por_aparelho_[ap] = c / c.sum()
        return self

    def predict_proba(self, X):
        if not self.usar_aparelho:
            return np.tile(self.prior_, (len(X), 1))
        return np.vstack([self.por_aparelho_.get(a, self.prior_) for a in X[COL_APARELHO].values])

    def predict(self, X):
        return self.classes_[self.predict_proba(X).argmax(axis=1)]


# --------------------------------------------------------------------------- modelos
def _word():
    return TfidfVectorizer(analyzer="word", ngram_range=(1, 2), min_df=2, sublinear_tf=True)


def _char():
    return TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 5), min_df=2, sublinear_tf=True)


def _prep(palavras=True, caracteres=True, aparelho=False):
    partes = []
    if palavras:
        partes.append(("w", _word(), COL_TEXTO))
    if caracteres:
        partes.append(("c", _char(), COL_TEXTO))
    if aparelho:
        partes.append(("d", OneHotEncoder(handle_unknown="ignore"), [COL_APARELHO]))
    return ColumnTransformer(partes)


def _lr():
    return LogisticRegression(class_weight="balanced", max_iter=3000, random_state=SEED)


def catalogo_modelos() -> dict:
    """Retorna {nome: (pipeline, grade de hiperparâmetros)}.
    Todos usam TF-IDF e pesos de classe balanceados (quando o algoritmo permite)."""
    c_lr = {"clf__C": [1, 3, 10, 30]}
    return {
        "LR · palavras": (Pipeline([("prep", _prep(True, False)), ("clf", _lr())]), c_lr),
        "LR · caracteres": (Pipeline([("prep", _prep(False, True)), ("clf", _lr())]), c_lr),
        "LR · palavras+caracteres": (Pipeline([("prep", _prep(True, True)), ("clf", _lr())]), c_lr),
        "LR · palavras+caracteres+aparelho": (Pipeline([("prep", _prep(True, True, True)), ("clf", _lr())]), c_lr),
        "SVM linear · palavras+caracteres": (
            Pipeline([("prep", _prep(True, True)),
                      ("clf", LinearSVC(class_weight="balanced", random_state=SEED, max_iter=10000))]),
            {"clf__C": [0.03, 0.1, 0.3, 1]}),
        "Naive Bayes (Complement) · palavras+caracteres": (
            Pipeline([("prep", _prep(True, True)), ("clf", ComplementNB())]),
            {"clf__alpha": [0.03, 0.1, 0.3, 1]}),
    }


def validacao_cruzada(df_treino: pd.DataFrame, n_splits: int = 5):
    """5 dobras estratificadas por rótulo e agrupadas por texto idêntico,
    para que o mesmo texto nunca apareça no treino e na validação ao mesmo tempo."""
    grupos = pd.factorize(df_treino[COL_TEXTO])[0]
    cv = StratifiedGroupKFold(n_splits=n_splits, shuffle=True, random_state=SEED)
    return list(cv.split(df_treino, df_treino[ALVO], grupos))


def treinar_e_selecionar(nome: str, pipe: Pipeline, grade: dict, df_treino: pd.DataFrame, cv) -> GridSearchCV:
    gs = GridSearchCV(pipe, grade, scoring="f1_macro", cv=cv, n_jobs=-1, refit=True)
    gs.fit(df_treino[[COL_TEXTO, COL_APARELHO]], df_treino[ALVO])
    return gs


# --------------------------------------------------------------------------- avaliação
def pontuacoes(modelo, X) -> np.ndarray:
    """Pontuação por classe (probabilidade ou margem) na ordem de modelo.classes_."""
    if hasattr(modelo, "predict_proba"):
        return modelo.predict_proba(X)
    return modelo.decision_function(X)


def topk_acuracia(y_true, scores, classes, k: int) -> float:
    classes = np.asarray(classes)
    topo = np.argsort(-scores, axis=1)[:, :k]
    return float(np.mean([y in classes[t] for y, t in zip(y_true, topo)]))


def metricas(y_true, y_pred, scores=None, classes=None) -> dict:
    m = {
        "acuracia": accuracy_score(y_true, y_pred),
        "acuracia_balanceada": balanced_accuracy_score(y_true, y_pred),
        "precisao_macro": precision_score(y_true, y_pred, average="macro", zero_division=0),
        "revocacao_macro": recall_score(y_true, y_pred, average="macro", zero_division=0),
        "f1_macro": f1_score(y_true, y_pred, average="macro", zero_division=0),
        "f1_ponderado": f1_score(y_true, y_pred, average="weighted", zero_division=0),
    }
    if scores is not None:
        for k in (2, 3, 5):
            m[f"top{k}"] = topk_acuracia(y_true, scores, classes, k)
    return {k: float(v) for k, v in m.items()}


def bootstrap_ic(y_true, y_pred, scores=None, classes=None, n=1000, seed=SEED, base_pred=None):
    """IC 95% por bootstrap (reamostragem do conjunto de teste). Se base_pred for dado,
    devolve também o IC da diferença pareada de F1-macro em relação à linha de base."""
    rng = np.random.default_rng(seed)
    y_true = np.asarray(y_true); y_pred = np.asarray(y_pred)
    N = len(y_true); acc, f1, t3, dif = [], [], [], []
    for _ in range(n):
        i = rng.integers(0, N, N)
        acc.append(accuracy_score(y_true[i], y_pred[i]))
        f1.append(f1_score(y_true[i], y_pred[i], average="macro", zero_division=0))
        if scores is not None:
            t3.append(topk_acuracia(y_true[i], scores[i], classes, 3))
        if base_pred is not None:
            dif.append(f1[-1] - f1_score(y_true[i], np.asarray(base_pred)[i], average="macro", zero_division=0))
    q = lambda v: [float(np.percentile(v, 2.5)), float(np.percentile(v, 97.5))]
    out = {"acuracia": q(acc), "f1_macro": q(f1)}
    if scores is not None:
        out["top3"] = q(t3)
    if base_pred is not None:
        out["dif_f1_macro_vs_base"] = q(dif)
    return out


def curva_confianca(y_true, y_pred, conf, cortes=np.arange(0.0, 0.91, 0.05)) -> pd.DataFrame:
    """Classificação seletiva: acurácia e cobertura se só respondermos quando conf >= corte."""
    y_true = np.asarray(y_true); y_pred = np.asarray(y_pred)
    linhas = []
    for c in cortes:
        sel = conf >= c
        linhas.append({"corte": float(c), "cobertura": float(sel.mean()),
                       "acuracia": float((y_true[sel] == y_pred[sel]).mean()) if sel.any() else np.nan,
                       "n": int(sel.sum())})
    return pd.DataFrame(linhas)


def faixa_tamanho(n: int) -> str:
    return "1–3 palavras" if n <= 3 else ("4–6 palavras" if n <= 6 else "7+ palavras")


def principais_termos(modelo_pipeline: Pipeline, classes, top: int = 6, prefixo: str = "w") -> dict:
    """Termos (palavras) de maior peso por classe em um pipeline com regressão logística."""
    prep = modelo_pipeline.named_steps["prep"]
    clf = modelo_pipeline.named_steps["clf"]
    nomes = prep.get_feature_names_out()
    idx = [i for i, n in enumerate(nomes) if n.startswith(f"{prefixo}__")]
    termos = {}
    for k, c in enumerate(classes):
        coef = clf.coef_[k][idx]
        melhores = np.argsort(-coef)[:top]
        termos[c] = [nomes[idx[j]].split("__", 1)[1] for j in melhores]
    return termos
