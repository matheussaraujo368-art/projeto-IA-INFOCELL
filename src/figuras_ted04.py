"""
figuras_ted04.py — Figuras e números novos da TED 04 (integração, ética, LGPD e Data Storytelling).

Gera:
- reports/figures/fig22_funil_dados.png        do dado bruto ao conjunto de teste
- reports/figures/fig23_representatividade.png  países e idiomas da Base A
- reports/resultados_ted04.json                 números citados no documento final

Uso (na raiz do repositório):  python src/figuras_ted04.py
"""
import json
import re
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

from graficos import CINZA, SERIES, TXT, TXT2, aplicar_estilo, salvar

RAIZ = Path(__file__).resolve().parents[1]
PROC = RAIZ / "data" / "processed"


def carregar():
    reparos = pd.read_csv(PROC / "reparos_limpo.csv", low_memory=False)
    falhas = pd.read_csv(PROC / "falhas_rotuladas_limpo.csv", low_memory=False)
    modelagem = pd.read_csv(PROC / "falhas_modelagem.csv")
    modelo = json.loads((RAIZ / "reports" / "resultados_modelagem.json").read_text(encoding="utf8"))
    return reparos, falhas, modelagem, modelo


def fig_funil(reparos, falhas, modelagem, modelo):
    etapas = [
        ("Base ORDS v0.3 completa (todas as categorias)", 305_649, CINZA),
        ("Celulares, notebooks, desktops e tablets", 23_612, SERIES[0]),
        ("Base A – reparos tratados", len(reparos), SERIES[0]),
        ("Base B – falhas rotuladas (quests)", len(falhas), SERIES[2]),
        ("Aptos para modelagem (com texto e causa)", len(modelagem), SERIES[2]),
        ("Treino (eventos até 2018)", modelo["n_treino"], SERIES[3]),
        ("Teste (eventos de 2019 e 2020)", modelo["n_teste"], SERIES[1]),
    ]
    fig, ax = plt.subplots(figsize=(8.6, 4.2))
    y = list(range(len(etapas)))[::-1]
    for yi, (rot, n, cor) in zip(y, etapas):
        ax.barh(yi, n, color=cor, height=0.62, log=True)
        ax.text(n * 1.12, yi, f"{n:,}".replace(",", "."), va="center", fontsize=9.5, color=TXT, fontweight="bold")
    ax.set_yticks(y, [e[0] for e in etapas], fontsize=9.5)
    ax.set_xscale("log")
    ax.set_xlim(1_000, 1_500_000)
    ax.set_xlabel("Registros (escala logarítmica)")
    ax.grid(axis="y", visible=False)
    ax.set_title("Do dado bruto ao modelo: quantos registros restam em cada etapa")
    fig.text(0.01, -0.02, "Azul: TED 02 (Base A) · Verde: TED 02 (Base B) · Amarelo e laranja: TED 03 (divisão temporal)",
             fontsize=8.5, color=TXT2)
    return salvar(fig, "fig22_funil_dados")


def fig_representatividade(reparos):
    paises = reparos["pais"].fillna("?").value_counts(normalize=True) * 100
    top = paises.head(8)
    outros = paises.iloc[8:].sum()
    nomes = {"GBR": "Reino Unido", "DEU": "Alemanha", "BEL": "Bélgica", "CAN": "Canadá", "NLD": "Países Baixos",
             "FRA": "França", "DNK": "Dinamarca", "NOR": "Noruega", "USA": "EUA", "ITA": "Itália", "IRL": "Irlanda"}
    rot = [nomes.get(p, p) for p in top.index] + ["Outros 20 países", "Brasil"]
    val = list(top.values) + [outros, 0.0]

    texto = reparos[reparos["tem_descricao"] == True]  # noqa: E712
    idi = texto["idioma"].fillna("indefinido").value_counts(normalize=True) * 100
    mapa = {"en": "Inglês", "de": "Alemão", "nl": "Holandês", "fr": "Francês"}
    rot2 = [mapa[k] for k in ["en", "de", "nl", "fr"] if k in idi.index]
    val2 = [idi[k] for k in ["en", "de", "nl", "fr"] if k in idi.index]
    resto = 100 - sum(val2) - idi.get("indefinido", 0)
    rot2 += ["Outros idiomas", "Indefinido", "Português"]
    val2 += [resto, idi.get("indefinido", 0), 0.0]

    fig, (a1, a2) = plt.subplots(1, 2, figsize=(9.6, 4.0), gridspec_kw={"wspace": 0.55})
    for ax, r, v, tit in [(a1, rot, val, "País do evento de reparo (% da Base A)"),
                          (a2, rot2, val2, "Idioma do relato (% dos relatos com texto)")]:
        cores = [SERIES[0]] * (len(r) - 1) + [SERIES[7]]
        yy = list(range(len(r)))[::-1]
        ax.barh(yy, v, color=cores, height=0.62)
        for yi, vi in zip(yy, v):
            ax.text(vi + 0.6, yi, f"{vi:.1f}%".replace(".", ","), va="center", fontsize=8.5, color=TXT)
        ax.set_yticks(yy, r, fontsize=9)
        ax.set_xlim(0, max(v) * 1.25)
        ax.grid(axis="y", visible=False)
        ax.set_title(tit, fontsize=10.5)
    fig.suptitle("Quem está representado nos dados: nenhum registro do Brasil e nenhum relato em português",
                 x=0.01, ha="left", fontsize=11.5, fontweight="bold", color=TXT, y=1.03)
    return salvar(fig, "fig23_representatividade")


# Padrões simples de dados pessoais. A varredura é indicativa: não encontra todo nome próprio.
PADROES = {
    "e-mail": r"[\w.+-]+@[\w-]+\.[\w.-]+",
    "endereço web (URL)": r"https?://|www\.",
    "sequência longa de dígitos (telefone ou documento)": r"(?<!\d)\d{9,}(?!\d)",
    "pronome de tratamento seguido de nome": r"\b(?:Mr|Mrs|Ms|Miss|Herr|Frau)\.? [A-Z][a-z]+",
}


def varredura_privacidade(reparos, falhas, modelagem):
    out = {}
    for nome, base in [("reparos_limpo", reparos), ("falhas_rotuladas_limpo", falhas), ("falhas_modelagem", modelagem)]:
        t = base["problema"].dropna().astype(str)
        out[nome] = {"textos": int(len(t)), **{k: int(t.str.contains(p, regex=True).sum()) for k, p in PADROES.items()}}
    return out


def main():
    aplicar_estilo()
    reparos, falhas, modelagem, modelo = carregar()
    fig_funil(reparos, falhas, modelagem, modelo)
    fig_representatividade(reparos)
    paises = reparos["pais"].value_counts(normalize=True)
    texto = reparos[reparos["tem_descricao"] == True]  # noqa: E712
    res = {
        "n_paises": int(reparos["pais"].nunique()),
        "registros_brasil": int((reparos["pais"] == "BRA").sum()),
        "pct_paises_top3": round(float(paises.head(3).sum() * 100), 1),
        "paises_top3": list(paises.head(3).index),
        "pct_fonte_restart": round(float((reparos["fonte"] == "The Restart Project").mean() * 100), 1),
        "pct_relatos_portugues": round(float((texto["idioma"] == "pt").mean() * 100), 1),
        "colunas_reparos": list(reparos.columns),
        "privacidade": varredura_privacidade(reparos, falhas, modelagem),
    }
    (RAIZ / "reports" / "resultados_ted04.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf8")
    print(json.dumps({k: v for k, v in res.items() if k != "colunas_reparos"}, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
