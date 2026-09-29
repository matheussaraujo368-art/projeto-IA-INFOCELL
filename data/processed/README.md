# Dados tratados (TED 02)

Gerados por `notebooks/TED02/02_limpeza_saneamento.ipynb` usando as regras de `src/preparacao.py`.

| Arquivo | Registros | Descrição |
|---|---|---|
| `reparos_limpo.csv` | 22.533 | **Base A – Reparos**: contexto do atendimento (aparelho, marca, idade, desfecho, barreira) |
| `falhas_rotuladas_limpo.csv` | 6.657 | **Base B – Falhas rotuladas**: sintoma → tipo de falha (MobiFix + CompCat + TabiCat) |
| `falhas_modelagem.csv` | 4.019 | Subconjunto da Base B com texto, rótulo informativo e sem duplicatas (para a TED 03) |
| `taxonomia_falhas.csv` | 41 | Mapeamento rótulo original → categoria unificada (14 categorias) |

## Dicionário — Base A (`reparos_limpo.csv`)
| Coluna | Tipo | Descrição |
|---|---|---|
| id | texto | Identificador ORDS original |
| fonte | categórico | Organização que coletou o dado |
| pais | categórico | País do evento (ISO-3) |
| categoria | categórico | Notebook, Celular, Tablet, Desktop |
| marca | categórico | Marca padronizada ("Desconhecida" se não informada) |
| ano_fabricacao | inteiro (nulo) | Ano validado (inválidos → nulo) |
| idade_anos | inteiro (nulo) | ano_evento − ano_fabricacao |
| faixa_idade | categórico ordinal | 0–2, 3–5, 6–8, 9–12, 13+ anos |
| idade_atipica | booleano | idade > 20 anos (outlier mantido) |
| status_reparo | categórico | Consertado, Reparável (não concluído), Fim de vida, Desconhecido |
| consertado | 0/1 (nulo) | 1 = consertado; nulo quando status desconhecido |
| barreira_reparo | categórico | Motivo do fim de vida / "Não informada" / "Não se aplica" |
| data_evento, ano_evento | data, inteiro | Data do evento de reparo |
| problema | texto | Sintoma após limpeza |
| tem_descricao | booleano | Há descrição do sintoma |
| n_palavras | inteiro | Tamanho da descrição |
| idioma | categórico | en, de, fr, nl, no, it, es, da, indefinido, sem texto |
| registro_em_lote | booleano | Registro idêntico a outro (exceto id) — reparo em lote |

## Dicionário — Base B (`falhas_rotuladas_limpo.csv`)
Mesmas colunas de contexto da Base A (`fonte`, `pais`, `categoria`, `marca`, `ano_fabricacao`, `idade_anos`, `status_reparo`, `consertado`, `data_evento`, `problema`, `tem_descricao`, `n_palavras`, `idioma`), mais:

| Coluna | Descrição |
|---|---|
| id_origem | ID da quest (MobiFix) ou ID ORDS (CompCat/TabiCat) |
| quest | MobiFix, CompCat ou TabiCat |
| tipo_falha_original | Rótulo original em inglês |
| categoria_falha | Categoria unificada em português (14 categorias) |
| rotulo_informativo | Falso quando a categoria é "Indeterminado" |
| duplicata_texto | Mesmo texto + mesmo rótulo + mesmo aparelho já visto |
| apto_modelagem | tem_descricao ∧ rotulo_informativo ∧ ¬duplicata_texto |
