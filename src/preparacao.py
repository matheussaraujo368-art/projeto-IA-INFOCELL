"""
preparacao.py — Funções de carga, limpeza e saneamento da base Open Repair Data
Projeto: Assistente de Diagnóstico Inteligente da Inforcell (UNIBALSAS)
Disciplina: Big Data e Ciência de Dados — TED 02

Todas as decisões de tratamento implementadas aqui estão justificadas na
Seção 8 do Documento Técnico v2.0 (docs/TED02.pdf).
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]
RAW = RAIZ / "data" / "raw"
PROCESSED = RAIZ / "data" / "processed"

# Versão da base utilizada (repositório github.com/openrepair/data)
VERSAO_ORDS = "v0.3"
VERSAO_EXPORT = "202507"  # dados até 31/07/2025, publicados em out/2025
ARQ_AGREGADO = f"OpenRepairData_{VERSAO_ORDS}_aggregate_{VERSAO_EXPORT}.csv"
URL_AGREGADO = (
    "https://raw.githubusercontent.com/openrepair/data/main/aggregated/"
    f"{VERSAO_EXPORT}/{ARQ_AGREGADO}"
)
ARQ_BRUTO_FILTRADO = RAW / f"ora_{VERSAO_EXPORT}_dispositivos_bruto.csv"

CATEGORIAS_ALVO = ["Mobile", "Laptop", "Desktop computer", "Tablet"]
CATEGORIA_PT = {
    "Mobile": "Celular",
    "Laptop": "Notebook",
    "Desktop computer": "Desktop",
    "Tablet": "Tablet",
}
STATUS_PT = {
    "Fixed": "Consertado",
    "Repairable": "Reparável (não concluído)",
    "End of life": "Fim de vida",
    "Unknown": "Desconhecido",
}
BARREIRA_PT = {
    "Spare parts too expensive": "Peças caras demais",
    "Repair information not available": "Falta de informação de reparo",
    "Spare parts not available": "Peças indisponíveis",
    "Lack of equipment": "Falta de equipamento",
    "No way to open product": "Produto não pode ser aberto",
    "Item too worn out": "Item muito desgastado",
}

# ---------------------------------------------------------------------------
# 1. CARGA
# ---------------------------------------------------------------------------

def carregar_agregado_completo(caminho: Path | str | None = None) -> pd.DataFrame:
    """Lê o CSV agregado completo (305.649 registros) — baixado do GitHub se preciso."""
    caminho = Path(caminho) if caminho else RAW / ARQ_AGREGADO
    fonte = caminho if caminho.exists() else URL_AGREGADO
    return pd.read_csv(fonte, dtype=str, keep_default_na=False, na_values=[""])


def carregar_bruto() -> pd.DataFrame:
    """Subconjunto bruto (sem nenhum tratamento) das 4 categorias do escopo da Inforcell.

    Se o arquivo filtrado ainda não existir em data/raw, ele é gerado a partir
    do agregado completo. Tudo é lido como texto para que a tipagem seja uma
    decisão explícita da etapa de limpeza.
    """
    if not ARQ_BRUTO_FILTRADO.exists():
        full = carregar_agregado_completo()
        sub = full[full["product_category"].isin(CATEGORIAS_ALVO)]
        sub.to_csv(ARQ_BRUTO_FILTRADO, index=False)
    return pd.read_csv(ARQ_BRUTO_FILTRADO, dtype=str, keep_default_na=False, na_values=[""])


def carregar_quests() -> dict[str, pd.DataFrame]:
    """Bases rotuladas com tipo de falha (quests MobiFix, CompCat e TabiCat)."""
    q = RAW / "quests"
    ler = lambda f: pd.read_csv(q / f, dtype=str, keep_default_na=False, na_values=[""])
    return {
        "mobifix": ler("mobifix.csv"),
        "compcat_laptops": ler("compcat_laptops.csv"),
        "compcat_desktops": ler("compcat_desktops.csv"),
        "tabicat": ler("tabicat.csv"),
    }


# ---------------------------------------------------------------------------
# 2. FUNÇÕES AUXILIARES DE LIMPEZA
# ---------------------------------------------------------------------------

# Itens classificados como "Tablet" pelos parceiros, mas que NÃO são tablets
# (GPS, e-readers, estações meteorológicas...). Fora do escopo da Inforcell.
PADRAO_FORA_ESCOPO = re.compile(
    r"e-?reader|e-?book|kindle(?! fire)|kobo|tolino|pocketbook|navigat|navi(?:gation|gerät)|"
    r"\bgps\b|satnav|weather|wetter|bicycle computer|fahrradcomputer|\bpda\b|"
    r"gps tracker|m[eé]t[eé]o|liseuse|smartwatch|walkie|pager|dect|schnurlos|cordless",
    re.IGNORECASE,
)
MARCAS_FORA_ESCOPO = {"TomTom", "Garmin", "Kobo", "Tolino", "PocketBook", "Navigon",
                      "Oregon Scientific", "TFA Dostmann", "Mio", "Cateye", "Sigma", "Alecto"}

# Normalização de marcas: variações de grafia, erros de digitação e nomes de
# produto usados no lugar da marca.
SINONIMOS_MARCA = {
    "apple": "Apple", "appel": "Apple", "aple": "Apple", "iphone": "Apple",
    "i phone": "Apple", "i-phone": "Apple", "ipad": "Apple", "i pad": "Apple",
    "macbook": "Apple", "macbook pro": "Apple", "macbook air": "Apple", "imac": "Apple",
    "mac": "Apple", "apple macbook": "Apple", "apple iphone": "Apple", "iphone 6": "Apple",
    "hp": "HP", "h.p.": "HP", "h p": "HP", "hewlett packard": "HP", "hewlett-packard": "HP",
    "hp compaq": "HP", "compaq": "HP", "hp6": "HP", "h[": "HP",
    "samsung": "Samsung", "samsung galaxy": "Samsung", "galaxy": "Samsung", "samson": "Samsung",
    "lenovo": "Lenovo", "lanovo": "Lenovo", "lenova": "Lenovo", "lenovo thinkpad": "Lenovo",
    "lenovo yoga": "Lenovo", "thinkpad": "Lenovo", "ibm": "IBM/Lenovo",
    "acer": "Acer", "aser": "Acer", "acer aspire": "Acer", "aspire": "Acer", "emachines": "Acer",
    "e-machines": "Acer", "packard bell": "Acer", "packardbell": "Acer",
    "asus": "Asus", "azus": "Asus", "google asus": "Asus",
    "dell": "Dell", "del": "Dell", "toshiba": "Toshiba", "sony": "Sony", "vaio": "Sony",
    "sony vaio": "Sony", "sony ericsson": "Sony",
    "huawei": "Huawei", "honor": "Huawei", "xiaomi": "Xiaomi", "redmi": "Xiaomi",
    "pocophone": "Xiaomi", "poco": "Xiaomi", "mi": "Xiaomi",
    "motorola": "Motorola", "moto": "Motorola", "nokia": "Nokia", "microsoft": "Microsoft",
    "surface": "Microsoft", "microsoft surface": "Microsoft",
    "amazon": "Amazon", "kindle": "Amazon", "kindle fire": "Amazon", "amazon fire": "Amazon",
    "medion": "Medion", "médion": "Medion", "medion akoya": "Medion",
    "fujitsu": "Fujitsu", "fujitsu siemens": "Fujitsu", "siemens": "Siemens",
    "lg": "LG", "lg electronics": "LG", "google": "Google", "pixel": "Google",
    "oneplus": "OnePlus", "one plus": "OnePlus", "fairphone": "Fairphone",
    "msi": "MSI", "htc": "HTC", "alcatel": "Alcatel", "doro": "Doro", "wiko": "Wiko",
    "oppo": "Oppo", "zte": "ZTE", "gigaset": "Gigaset", "panasonic": "Panasonic",
    "chrome book": "Desconhecida", "chromebook": "Desconhecida", "chrome": "Desconhecida",
    "android": "Desconhecida", "windows": "Desconhecida", "laptop": "Desconhecida",
    "unknown": "Desconhecida", "ukendt": "Desconhecida", "unbekannt": "Desconhecida",
    "onbekend": "Desconhecida", "inconnu": "Desconhecida", "inconnue": "Desconhecida",
    "xx": "Desconhecida", "?": "Desconhecida", "-": "Desconhecida", "n/a": "Desconhecida",
    "none": "Desconhecida", "various": "Desconhecida", "diverse": "Desconhecida",
    "eigenbau": "Desconhecida", "no brand": "Desconhecida", "generic": "Desconhecida",
}


def normalizar_marca(valor) -> str:
    if pd.isna(valor):
        return "Desconhecida"
    v = re.sub(r"\s+", " ", str(valor)).strip()
    chave = v.lower().strip(" .,;")
    if chave in SINONIMOS_MARCA:
        return SINONIMOS_MARCA[chave]
    # Heurística: marca composta ("HP Pavilion", "Apple iPad Air") -> primeira palavra
    primeira = chave.split(" ")[0]
    if primeira in SINONIMOS_MARCA and SINONIMOS_MARCA[primeira] != "Desconhecida":
        return SINONIMOS_MARCA[primeira]
    if len(chave) <= 1 or not re.search(r"[a-z]", chave):
        return "Desconhecida"
    return v[:1].upper() + v[1:] if v.islower() else v


# Descrições "vazias de conteúdo" — não informam o sintoma
PLACEHOLDERS_PROBLEMA = {
    "", ".", "-", "?", "??", "???", "unknown", "unknown.", "n/a", "na", "none", "x",
    "onbekend", "onbekend.", "unbekannt", "inconnu", "ukendt", "test", "nvt", "divers",
}


def limpar_texto_problema(valor):
    """Normaliza espaços, remove artefatos de exportação ("~", ". ." etc.)."""
    if pd.isna(valor):
        return np.nan
    t = str(valor)
    t = t.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    t = re.sub(r"https?://\S+", " ", t)  # URLs não descrevem sintoma
    t = re.sub(r"\s*~\s*", " - ", t)  # separador herdado de partner_product_category
    t = re.sub(r"(\s*\.\s*){2,}", ". ", t)  # ". ." repetido
    t = re.sub(r"\s+", " ", t).strip(" .-;,")
    if t.lower() in PLACEHOLDERS_PROBLEMA or len(t) < 3:
        return np.nan
    return t


# ----- Detecção de idioma (dicionário de palavras funcionais) --------------
# Justificativa: a coluna `language` citada na TED01 não existe na versão
# v0.3 agregada; bibliotecas externas de detecção não estavam disponíveis no
# ambiente. Usamos um detector transparente e validado contra os rótulos de
# idioma da TabiCat (ver notebook 02).
STOPWORDS = {
    "en": "the and not is was to of it on in with no for does doesn't won't wont will has have broken working "
          "work screen battery charge charging replaced cracked needs slow need but from fixed turn switch "
          "power advice when after new keyboard phone laptop cleaned stopped port loose".split(),
    "nl": "de het een niet werkt en van is op met geen te voor defect kapot scherm accu laadt meer "
          "doet gaat opgeladen vervangen nieuwe maar wel ook bij na aan nog kan".split(),
    "de": "der die das nicht und ist mit ein eine kein keine geht funktioniert defekt lädt mehr "
          "auf von zu bildschirm akku ausgetauscht wurde neu startet kaputt nur noch bei".split(),
    "fr": "le la les ne pas est et de du des un une plus fonctionne marche écran batterie charge "
          "cassé remplacé ordinateur problème s'allume allume sur avec pour".split(),
    "da": "ikke og er det en på virker skærm batteri oplader lader med til af har kan fejl "
          "tænder skiftet ny".split(),
    "it": "non il la funziona schermo batteria rotto si accende di che per con è".split(),
    "es": "no el la funciona pantalla batería roto enciende de que por con se carga".split(),
    "no": "ikke og er det virker skjerm batteri lader lades med til av har kan".split(),
}
_STOP_SETS = {k: set(v) for k, v in STOPWORDS.items()}


IDIOMA_POR_PAIS = {  # países monolíngues (fallback quando o texto é curto/ambíguo)
    "GBR": "en", "USA": "en", "AUS": "en", "NZL": "en", "IRL": "en", "JEY": "en",
    "DEU": "de", "AUT": "de", "NLD": "nl", "FRA": "fr", "DNK": "da", "NOR": "no",
    "ITA": "it", "ESP": "es",
}


def idioma_com_fallback(textos: pd.Series, paises: pd.Series) -> pd.Series:
    """Detecta pelo texto; se 'indefinido', usa o idioma oficial do país (se único)."""
    det = textos.map(detectar_idioma)
    fb = paises.map(IDIOMA_POR_PAIS)
    usar = (det == "indefinido") & fb.notna()
    return det.where(~usar, fb)


def detectar_idioma(texto) -> str:
    if pd.isna(texto):
        return "sem texto"
    tokens = re.findall(r"[a-zà-ÿæøåäöüß']+", str(texto).lower())
    if not tokens:
        return "indefinido"
    pont = {lang: sum(t in s for t in tokens) for lang, s in _STOP_SETS.items()}
    melhor = max(pont, key=pont.get)
    if pont[melhor] == 0:
        return "indefinido"
    # desempate: se empate entre línguas, marca como indefinido
    if list(pont.values()).count(pont[melhor]) > 1:
        return "indefinido"
    return melhor


# ---------------------------------------------------------------------------
# 3. BASE A — REPAROS (contexto: dispositivo, idade, desfecho, barreira)
# ---------------------------------------------------------------------------

def limpar_reparos(bruto: pd.DataFrame, log: list | None = None) -> pd.DataFrame:
    """Aplica os tratamentos da Base A e registra cada etapa em `log`."""
    log = log if log is not None else []
    df = bruto.copy()
    n0 = len(df)
    log.append(("Registros brutos (4 categorias)", n0))

    # 3.1 Tipos de dados --------------------------------------------------
    df["event_date"] = pd.to_datetime(df["event_date"], errors="coerce")
    df["year_of_manufacture"] = pd.to_numeric(df["year_of_manufacture"], errors="coerce")
    df["product_age"] = pd.to_numeric(df["product_age"], errors="coerce")

    # 3.2 Escopo: itens fora do perfil da Inforcell -----------------------
    fora = df["partner_product_category"].fillna("").str.contains(PADRAO_FORA_ESCOPO) | (
        (df["product_category"] == "Tablet") & df["brand"].isin(MARCAS_FORA_ESCOPO)
    )
    log.append(("Removidos: itens fora de escopo (GPS, e-reader, estação meteorológica...)", int(fora.sum())))
    df = df[~fora].copy()

    # 3.3 Marca -----------------------------------------------------------
    df["marca"] = df["brand"].map(normalizar_marca)

    # 3.4 Ano de fabricação e idade --------------------------------------
    ano_evento = df["event_date"].dt.year
    limite_inf = df["product_category"].map(
        {"Mobile": 1985, "Tablet": 2000, "Laptop": 1985, "Desktop computer": 1977}
    )
    invalido = (df["year_of_manufacture"] > ano_evento) | (df["year_of_manufacture"] < limite_inf)
    log.append(("Ano de fabricação inválido -> ausente (futuro ou anterior à tecnologia)", int(invalido.sum())))
    df.loc[invalido, "year_of_manufacture"] = np.nan
    # a idade é recalculada a partir do ano (garante consistência entre as colunas)
    df["idade_anos"] = (ano_evento - df["year_of_manufacture"]).astype("Int64")
    df["year_of_manufacture"] = df["year_of_manufacture"].astype("Int64")
    df["faixa_idade"] = pd.cut(
        df["idade_anos"].astype(float), bins=[-0.1, 2, 5, 8, 12, 100],
        labels=["0–2 anos", "3–5 anos", "6–8 anos", "9–12 anos", "13+ anos"],
    )
    df["idade_atipica"] = df["idade_anos"].astype(float) > 20  # outlier mantido e sinalizado

    # 3.5 Texto do problema ----------------------------------------------
    df["problema"] = df["problem"].map(limpar_texto_problema)
    novos_nulos = int(df["problema"].isna().sum() - df["problem"].isna().sum())
    log.append(("Descrições sem conteúdo (placeholder) -> ausente", novos_nulos))
    df["tem_descricao"] = df["problema"].notna()
    df["n_palavras"] = df["problema"].fillna("").str.split().str.len().astype(int)
    df["idioma"] = idioma_com_fallback(df["problema"], df["country"])

    # 3.6 Categorias traduzidas e barreira ------------------------------
    df["categoria"] = df["product_category"].map(CATEGORIA_PT)
    df["status_reparo"] = df["repair_status"].map(STATUS_PT)
    df["consertado"] = np.where(df["repair_status"] == "Unknown", np.nan,
                                (df["repair_status"] == "Fixed").astype(float))
    barreira = df["repair_barrier_if_end_of_life"].map(BARREIRA_PT)
    df["barreira_reparo"] = np.where(
        df["repair_status"] == "End of life", barreira.fillna("Não informada"), "Não se aplica")

    # 3.7 Duplicatas -------------------------------------------------------
    chave = ["data_provider", "country", "partner_product_category", "product_category", "brand",
             "year_of_manufacture", "repair_status", "group_identifier", "event_date", "problem"]
    df["registro_em_lote"] = df.duplicated(subset=chave, keep=False)
    dup_exatos = int(df.duplicated(subset=chave, keep="first").sum())
    log.append(("Registros idênticos (exceto id) — mantidos e sinalizados", dup_exatos))

    # 3.8 Atributos derivados de data -------------------------------------
    df["ano_evento"] = df["event_date"].dt.year.astype("Int64")

    # 3.9 Seleção final de colunas (remove redundantes/irrelevantes) -----
    colunas = ["id", "data_provider", "country", "categoria", "marca", "year_of_manufacture",
               "idade_anos", "faixa_idade", "idade_atipica", "status_reparo", "consertado",
               "barreira_reparo", "event_date", "ano_evento", "problema", "tem_descricao",
               "n_palavras", "idioma", "registro_em_lote"]
    df = df[colunas].rename(columns={"data_provider": "fonte", "country": "pais",
                                     "year_of_manufacture": "ano_fabricacao",
                                     "event_date": "data_evento"})
    for c in ["fonte", "pais", "categoria", "marca", "status_reparo", "barreira_reparo", "idioma"]:
        df[c] = df[c].astype("category")
    log.append(("Registros na Base A tratada", len(df)))
    return df.reset_index(drop=True)


# ---------------------------------------------------------------------------
# 4. BASE B — FALHAS ROTULADAS (sintoma -> tipo de falha)
# ---------------------------------------------------------------------------

# Taxonomia unificada (PT-BR). As três quests usam listas de falhas diferentes;
# aqui elas são harmonizadas em categorias próximas da rotina de bancada da
# Inforcell e dos passos do diagnóstico guiado previsto no MVP.
TAXONOMIA = {
    # Energia
    "Power/battery": "Energia / bateria", "Charger": "Energia / bateria",
    # Tela
    "Screen": "Tela", "Integrated screen": "Tela",
    # Conectores
    "USB/charging port": "Conectores e portas", "Ports/slots/connectors": "Conectores e portas",
    "Headphone jack": "Conectores e portas", "Sim card slot": "Conectores e portas",
    "Memory card slot": "Conectores e portas",
    # Botões e entrada
    "On/Off button": "Botões / teclado / touchpad", "Volume buttons": "Botões / teclado / touchpad",
    "Other buttons": "Botões / teclado / touchpad", "Integrated keyboard": "Botões / teclado / touchpad",
    "Integrated pointing device": "Botões / teclado / touchpad",
    # Áudio, câmera e mídia
    "Speaker/amplifier": "Áudio / câmera / mídia", "Microphone": "Áudio / câmera / mídia",
    "Camera": "Áudio / câmera / mídia", "Integrated media component": "Áudio / câmera / mídia",
    "Bluetooth": "Áudio / câmera / mídia", "Integrated optical drive": "Áudio / câmera / mídia",
    # Placa
    "System board": "Placa / dano interno", "Internal damage": "Placa / dano interno",
    "Liquid damage": "Dano por líquido",
    "External damage": "Carcaça / dano externo", "Case/chassis": "Carcaça / dano externo",
    "Storage problem": "Armazenamento", "Internal storage": "Armazenamento",
    # Software
    "Software update": "Software / sistema", "Software issue": "Software / sistema",
    "Operating system": "Software / sistema", "Configuration": "Software / sistema",
    "Help/configuration": "Software / sistema", "Virus/malware": "Software / sistema",
    "Access issue": "Software / sistema",
    "Boot": "Inicialização (boot)", "Stuck booting": "Inicialização (boot)",
    "Performance": "Desempenho / superaquecimento", "Overheating": "Desempenho / superaquecimento",
    # Não informativos
    "Other": "Outros", "Multiple": "Outros",
    "Unknown": "Indeterminado", "Poor data": "Indeterminado",
}

_COMPCAT_TIPOS = RAW / "quests" / "compcat_fault_types.csv"


def _normaliza_para_chave(t: str) -> str:
    t = unicodedata.normalize("NFKD", str(t).lower())
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9 ]+", " ", t).strip()


def montar_falhas(quests: dict[str, pd.DataFrame], log: list | None = None) -> pd.DataFrame:
    log = log if log is not None else []
    partes = []

    m = quests["mobifix"].copy()
    m["event_date"] = pd.to_datetime(m["event_date"], errors="coerce")
    idade = pd.to_numeric(m["age"], errors="coerce")
    m["year_of_manufacture"] = m["event_date"].dt.year - idade  # idade -> ano (padroniza com as demais)
    m["quest"] = "MobiFix"
    m = m.rename(columns={"id": "id_origem"})
    partes.append(m)

    for nome, rot in [("compcat_laptops", "CompCat"), ("compcat_desktops", "CompCat"), ("tabicat", "TabiCat")]:
        q = quests[nome].copy()
        q["event_date"] = pd.to_datetime(q["event_date"], errors="coerce")
        # TabiCat usa "????" para ano desconhecido
        q["year_of_manufacture"] = pd.to_numeric(q["year_of_manufacture"].replace("????", np.nan), errors="coerce")
        q["quest"] = rot
        q = q.rename(columns={"id_ords": "id_origem"})
        partes.append(q)

    cols = ["id_origem", "quest", "data_provider", "country", "product_category", "brand",
            "repair_status", "year_of_manufacture", "event_date", "problem", "fault_type"]
    df = pd.concat([p[cols] for p in partes], ignore_index=True)
    log.append(("Registros rotulados brutos (MobiFix + CompCat + TabiCat)", len(df)))

    ano_evento = df["event_date"].dt.year
    inval = (df["year_of_manufacture"] > ano_evento) | (df["year_of_manufacture"] < 1977)
    df.loc[inval, "year_of_manufacture"] = np.nan
    log.append(("Ano de fabricação inválido -> ausente", int(inval.sum())))

    df["categoria"] = df["product_category"].map(CATEGORIA_PT)
    df["marca"] = df["brand"].map(normalizar_marca)
    df["status_reparo"] = df["repair_status"].map(STATUS_PT)
    df["consertado"] = np.where(df["repair_status"] == "Unknown", np.nan,
                                (df["repair_status"] == "Fixed").astype(float))
    df["tipo_falha_original"] = df["fault_type"]
    df["categoria_falha"] = df["fault_type"].map(TAXONOMIA)
    sem_map = df["categoria_falha"].isna().sum()
    assert sem_map == 0, f"{sem_map} tipos de falha sem mapeamento: {df.loc[df.categoria_falha.isna(), 'fault_type'].unique()}"
    df["rotulo_informativo"] = ~df["categoria_falha"].isin(["Indeterminado"])

    df["problema"] = df["problem"].map(limpar_texto_problema)
    df["tem_descricao"] = df["problema"].notna()
    df["n_palavras"] = df["problema"].fillna("").str.split().str.len().astype(int)
    df["idioma"] = idioma_com_fallback(df["problema"], df["country"])
    df["texto_norm"] = df["problema"].fillna("").map(_normaliza_para_chave)
    df["ano_fabricacao"] = df["year_of_manufacture"].astype("Int64")
    df["idade_anos"] = (ano_evento - df["year_of_manufacture"]).astype("Int64")

    # Duplicatas: mesmo texto + mesmo rótulo + mesma categoria de aparelho.
    # Mantidas na base descritiva, mas marcadas para exclusão no conjunto de
    # modelagem (evita que o futuro modelo "decore" frases repetidas).
    df["duplicata_texto"] = df["tem_descricao"] & df.duplicated(
        subset=["texto_norm", "categoria_falha", "categoria"], keep="first")
    log.append(("Duplicatas de texto+rótulo (sinalizadas)", int(df["duplicata_texto"].sum())))

    df["apto_modelagem"] = df["tem_descricao"] & df["rotulo_informativo"] & ~df["duplicata_texto"]
    log.append(("Registros aptos para modelagem (texto + rótulo informativo, sem duplicata)",
                int(df["apto_modelagem"].sum())))

    out = df[["id_origem", "quest", "data_provider", "country", "categoria", "marca", "ano_fabricacao",
              "idade_anos", "status_reparo", "consertado", "event_date", "problema", "tem_descricao",
              "n_palavras", "idioma", "tipo_falha_original", "categoria_falha", "rotulo_informativo",
              "duplicata_texto", "apto_modelagem"]].rename(
        columns={"data_provider": "fonte", "country": "pais", "event_date": "data_evento"})
    for c in ["quest", "fonte", "pais", "categoria", "marca", "status_reparo", "idioma",
              "tipo_falha_original", "categoria_falha"]:
        out[c] = out[c].astype("category")
    return out


# ---------------------------------------------------------------------------
# 5. PERFIL DE QUALIDADE (usado nos notebooks e no documento)
# ---------------------------------------------------------------------------

def perfil_qualidade(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({
        "tipo": df.dtypes.astype(str),
        "nao_nulos": df.notna().sum(),
        "ausentes": df.isna().sum(),
        "%_ausentes": (df.isna().mean() * 100).round(1),
        "valores_distintos": df.nunique(),
    })


def cramers_v(x: pd.Series, y: pd.Series) -> float:
    """V de Cramér com correção de viés (Bergsma, 2013)."""
    from scipy.stats import chi2_contingency
    tab = pd.crosstab(x, y)
    if tab.shape[0] < 2 or tab.shape[1] < 2:
        return np.nan
    chi2 = chi2_contingency(tab, correction=False)[0]
    n = tab.values.sum()
    phi2 = chi2 / n
    r, k = tab.shape
    phi2c = max(0, phi2 - (k - 1) * (r - 1) / (n - 1))
    rc, kc = r - (r - 1) ** 2 / (n - 1), k - (k - 1) ** 2 / (n - 1)
    return float(np.sqrt(phi2c / max(1e-9, min(kc - 1, rc - 1))))
