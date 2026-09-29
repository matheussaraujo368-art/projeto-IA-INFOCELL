# Assistente de Diagnóstico Inteligente — Inforcell

Projeto multidisciplinar do curso de **Sistemas de Informação da UNIBALSAS (Balsas/MA)** para a **Inforcell**, assistência técnica de celulares e computadores em Balsas/MA.

**Equipe:** Rafael R Pontara · Matheus de Brito Araujo · Pedro Junior Ubirajara · Pedro Paulo T. Araujo

## Problema e MVP
Um mesmo sintoma ("o celular não liga") pode ter várias causas, e o diagnóstico depende da experiência individual do técnico. O MVP é uma plataforma web que une a gestão de Ordens de Serviço a um **assistente de IA que conduz o técnico por um diagnóstico guiado** (sintoma → causas possíveis → testes → diagnóstico final).

## Disciplina de Big Data e Ciência de Dados — evolução
| Etapa | Entrega | Documento |
|---|---|---|
| TED 01 | Definição do problema de dados e seleção da base **Open Repair Data** | `docs/TED01.pdf` (v1.0) |
| TED 02 | Limpeza, saneamento e análise exploratória | `docs/TED02.pdf` / `docs/TED02.docx` (v2.0) |

### Principais resultados da TED 02
- A base atual (ORDS v0.3) tem 305.649 registros e **não tem** `fault_type`; o projeto passou a usar **duas bases**: Reparos (22.533 registros tratados) e Falhas rotuladas (6.657 registros das quests MobiFix, CompCat e TabiCat, com taxonomia unificada de 14 categorias).
- O sintoma **"não liga"** aparece associado a **11 tipos de falha** — ~80% em energia/bateria e inicialização —, confirmando com dados o problema da Inforcell e sugerindo a ordem dos testes do diagnóstico guiado.
- Falhas eletrônicas (placa, energia, boot) têm as menores taxas de conserto (33–39%); a idade do aparelho quase não influencia o desfecho.
- 4.019 registros ficaram aptos para a modelagem da TED 03.

## Estrutura
```
projeto-inforcell/
├── README.md
├── requirements.txt
├── docs/
│   ├── TED01.pdf
│   ├── TED02.pdf / TED02.docx      Documento Técnico v2.0
│   └── build/                      scripts que geram o documento a partir dos resultados
├── data/
│   ├── raw/                        dados originais sem alteração (+ script de download)
│   │   └── quests/
│   └── processed/                  bases tratadas + dicionário de dados
├── notebooks/
│   ├── TED01/
│   └── TED02/
│       ├── 01_carga_inspecao.ipynb
│       ├── 02_limpeza_saneamento.ipynb
│       └── 03_analise_exploratoria.ipynb
├── src/
│   ├── preparacao.py               regras de limpeza (reutilizáveis)
│   └── graficos.py                 estilo e paleta dos gráficos
└── reports/
    ├── figures/                    fig01 … fig12
    └── resultados_eda.json         números usados no documento
```

## Como reproduzir
```bash
pip install -r requirements.txt
cd notebooks/TED02
jupyter notebook   # executar 01 → 02 → 03, nesta ordem
```
O notebook 02 regrava `data/processed/`; o 03 regrava `reports/figures/` e `reports/resultados_eda.json`.
Para regenerar o documento: `node docs/build/gerar_documento.js` (requer o pacote npm `docx`).

## Fonte dos dados e licença
Open Repair Alliance — https://openrepair.org/open-data/ · https://github.com/openrepair/data
Licença **CC BY-SA 4.0**. Detalhes da versão utilizada em `data/raw/README.md`.
