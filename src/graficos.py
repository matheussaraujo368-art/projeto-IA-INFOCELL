"""
graficos.py — Estilo visual padronizado para as figuras da TED 02.
Paleta categórica validada para daltonismo (ordem fixa, nunca reciclada).
"""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

FIG_DIR = Path(__file__).resolve().parents[1] / "reports" / "figures"
FIG_DIR.mkdir(parents=True, exist_ok=True)

SERIES = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
AZUL_SEQ = ["#cde2fb", "#9ec5f4", "#6da7ec", "#3987e5", "#256abf", "#184f95", "#0d366b"]
CINZA = "#a3a29c"
TXT = "#0b0b0b"
TXT2 = "#52514e"
GRADE = "#e4e3de"

# cor fixa por entidade (a cor segue a entidade, não a posição)
COR_STATUS = {
    "Consertado": SERIES[0],
    "Reparável (não concluído)": SERIES[2],
    "Fim de vida": SERIES[1],
    "Desconhecido": CINZA,
}
COR_CATEGORIA = {"Notebook": SERIES[0], "Celular": SERIES[1], "Tablet": SERIES[2], "Desktop": SERIES[3]}


def aplicar_estilo():
    plt.rcParams.update({
        "figure.dpi": 110, "savefig.dpi": 200, "figure.facecolor": "white",
        "axes.facecolor": "white", "axes.edgecolor": GRADE, "axes.labelcolor": TXT2,
        "axes.titlesize": 12, "axes.titleweight": "bold", "axes.titlecolor": TXT,
        "axes.titlelocation": "left", "axes.spines.top": False, "axes.spines.right": False,
        "axes.grid": True, "grid.color": GRADE, "grid.linewidth": 0.6,
        "xtick.color": TXT2, "ytick.color": TXT2, "font.size": 10,
        "font.family": "DejaVu Sans", "legend.frameon": False, "lines.linewidth": 2,
    })


def salvar(fig, nome: str):
    caminho = FIG_DIR / f"{nome}.png"
    fig.savefig(caminho, bbox_inches="tight", facecolor="white")
    return caminho
