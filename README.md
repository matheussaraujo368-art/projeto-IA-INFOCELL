# Assistente de Diagnóstico Inteligente — Inforcell

Projeto multidisciplinar do curso de **Sistemas de Informação da UNIBALSAS (Balsas/MA)** para a **Inforcell**, assistência técnica de celulares e computadores em Balsas/MA.

**Equipe:** Rafael R Pontara · Matheus de Brito Araujo · Pedro Paulo T. Araujo (Pedro Junior Ubirajara participou das TEDs 01 e 02)

**Tecnologias:** Python (pandas, NumPy, SciPy, scikit-learn, Matplotlib), Jupyter, pytest; documentos gerados com Node.js (docx, pptxgenjs) e LibreOffice.

## Problema e MVP
Um mesmo sintoma ("o celular não liga") pode ter várias causas, e o diagnóstico depende da experiência individual do técnico. O MVP é uma plataforma web que une a gestão de Ordens de Serviço a um **assistente de IA que conduz o técnico por um diagnóstico guiado** (sintoma → causas possíveis → testes → diagnóstico final).

## Disciplina de Big Data e Ciência de Dados — evolução
| Etapa | Entrega | Documento |
|---|---|---|
| TED 01 | Definição do problema de dados e seleção da base **Open Repair Data** | `docs/TED01.pdf` (v1.0) |
| TED 02 | Limpeza, saneamento e análise exploratória | `docs/TED02.pdf` / `docs/TED02.docx` (v2.0) |
| TED 03 | Modelagem preditiva e Machine Learning | `docs/TED03.pdf` / `docs/TED03.docx` (v3.0) |
| TED 04 | Integração dos resultados, ética, LGPD e Data Storytelling | **`docs/documento-final.pdf`** (Versão Final) e `presentation/apresentacao-final.pdf` |

**Comece por aqui:** o [Documento Técnico – Versão Final](docs/documento-final.pdf) reúne as quatro etapas em 13 seções, e a [apresentação](presentation/apresentacao-final.pdf) conta a história dos resultados em 13 slides. As versões 1.0 a 3.0 continuam em `docs/`, com todos os detalhes de cada etapa.

### Principais conclusões (TED 04)
- Os dados confirmam o problema da Inforcell: o sintoma “não liga” se distribui por 11 categorias de falha, mas cerca de 80% são energia/bateria ou inicialização.
- O modelo serve para **ordenar hipóteses** (causa certa entre as 3 primeiras em 89,5% dos casos), não para dar o diagnóstico final.
- Não há registros do Brasil nem relatos em português: antes de usar o modelo na Inforcell, é preciso avaliá-lo com dados próprios.
- Ética e LGPD: a base pública tem risco baixo, mas não nulo (poucos nomes em texto livre); o sistema em produção trata dados de clientes e precisa de finalidade, retenção e direitos do titular definidos.
- O modelo (`sugerir_causas`) ainda **não está integrado** ao sistema web; a integração é o primeiro passo de evolução.

### Principais resultados da TED 03
- Problema: classificação multiclasse **sintoma (texto) + aparelho → categoria de falha** (13 classes), com saída em probabilidades. São as 14 categorias da taxonomia da TED 02 menos "Indeterminado", que não informa a causa e foi excluída da modelagem (Seção 15.4 do documento).
- Divisão temporal (treino ≤ 2018: 2.224; teste 2019–2020: 1.795). Modelo: regressão logística sobre TF-IDF de palavras + caracteres + aparelho.
- Teste: acurácia 71,6%, F1-macro 0,654 (IC95% 0,619–0,683), top-3 89,5% — contra 17,5% e 0,053 da linha de base.
- Limitação central: nenhum texto em português na base; confiança < 0,35 deve acionar perguntas complementares.
- Uso no MVP: `src/inferencia.py` → `sugerir_causas(texto, aparelho)`.

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
│   ├── TED03.pdf / TED03.docx      Documento Técnico v3.0
│   ├── documento-final.pdf / .docx Documento Técnico – Versão Final (TED 04)
│   └── build/                      scripts que geram o documento a partir dos resultados
├── presentation/
│   └── apresentacao-final.pdf / .pptx   apresentação de Data Storytelling (TED 04)
├── data/
│   ├── raw/                        dados originais sem alteração (+ script de download)
│   │   └── quests/
│   └── processed/                  bases tratadas + dicionário de dados
├── notebooks/
│   ├── TED01/                      vazia: a TED 01 foi de seleção e avaliação da base, sem código
│   ├── TED02/
│   │   ├── 01_carga_inspecao.ipynb
│   │   ├── 02_limpeza_saneamento.ipynb
│   │   └── 03_analise_exploratoria.ipynb
│   └── TED03/
│       ├── 01_definicao_preparacao.ipynb
│       ├── 02_treinamento_avaliacao.ipynb
│       └── 03_analise_erros_aplicabilidade.ipynb
├── src/
│   ├── preparacao.py               regras de limpeza (reutilizáveis)
│   ├── graficos.py                 estilo e paleta dos gráficos
│   ├── modelagem.py                TED 03: divisão, modelos, métricas
│   ├── inferencia.py               TED 03: sugerir_causas(texto, aparelho)
│   └── figuras_ted04.py            TED 04: figuras 22 e 23 e números de privacidade
├── models/                         classificadores treinados (.joblib)
├── tests/                          testes automatizados (pytest -q)
└── reports/
    ├── figures/                    fig01 … fig23
    ├── resultados_eda.json         números da TED 02
    ├── resultados_modelagem.json   números da TED 03
    └── resultados_ted04.json       números da TED 04
```

## Como reproduzir
```bash
pip install -r requirements.txt
cd notebooks/TED02
jupyter notebook   # executar 01 → 02 → 03, nesta ordem
```
O notebook 02 regrava `data/processed/`; o 03 regrava `reports/figures/` e `reports/resultados_eda.json`.
Para a TED 03: `cd notebooks/TED03` e executar 01 → 02 → 03 (gera `models/`, `reports/figures/fig13…fig21` e `reports/resultados_modelagem.json`).
Para conferir o ambiente e o código: `pytest -q` na raiz (24 testes).
Para a TED 04: `cd src && python figuras_ted04.py` (gera `fig22`, `fig23` e `reports/resultados_ted04.json`).
Para regenerar os documentos: `node docs/build/gerar_documento.js` (v2.0), `node docs/build/gerar_documento_v3.js` (v3.0) e `node docs/build/gerar_documento_final.js` (Versão Final), que requerem o pacote npm `docx`; o sumário e o PDF são gerados com `docs/build/atualizar_sumario.py` (Python do LibreOffice). A apresentação: `node docs/build/gerar_apresentacao.js` (pacote npm `pptxgenjs`) e conversão para PDF pelo LibreOffice.

## Fonte dos dados e licença
Open Repair Alliance — https://openrepair.org/open-data/ · https://github.com/openrepair/data
Licença **CC BY-SA 4.0**. Detalhes da versão utilizada em `data/raw/README.md`.
