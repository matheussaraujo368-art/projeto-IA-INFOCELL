"""Baixa o arquivo agregado completo da Open Repair Data (58 MB, não versionado no Git)."""
import sys, urllib.request
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))
import preparacao as prep

destino = prep.RAW / prep.ARQ_AGREGADO
print("Baixando", prep.URL_AGREGADO)
urllib.request.urlretrieve(prep.URL_AGREGADO, destino)
print("Salvo em", destino)
