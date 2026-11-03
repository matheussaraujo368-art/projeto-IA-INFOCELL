// Gera docs/TED03.docx — Documento Técnico v3.0 (preserva a TED 01 e a TED 02 e acrescenta a TED 03)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, AlignmentType,
  HeadingLevel, LevelFormat, WidthType, ShadingType, BorderStyle, PageBreak, TableOfContents,
  Footer, Header, PageNumber, ExternalHyperlink,
} = require("docx");

const ROOT = path.resolve(__dirname, "..", "..");
const FIG = path.join(ROOT, "reports", "figures");
const R = JSON.parse(fs.readFileSync(path.join(ROOT, "reports", "resultados_eda.json"), "utf8"));

const FONT = "Arial";
const W = 9071; // largura útil A4 com margens ABNT (3 cm esq., 2 cm dir.)
const fmt = (n, d = 1) => Number(n).toLocaleString("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d });
const int = (n) => Number(n).toLocaleString("pt-BR");

// ---------------------------------------------------------------- helpers
function runs(text, base = {}) {
  // **negrito** e _itálico_ simples
  const out = [];
  const re = /(\*\*[^*]+\*\*|~~[^~]+~~)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const t = m[0];
    if (t.startsWith("**")) out.push(new TextRun({ text: t.slice(2, -2), bold: true, ...base }));
    else out.push(new TextRun({ text: t.slice(2, -2), italics: true, ...base }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...base }));
  return out;
}
const P = (text, o = {}) => new Paragraph({
  children: runs(text, o.run || {}), alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { line: 360, after: 120, ...(o.spacing || {}) }, indent: o.noIndent ? undefined : { firstLine: 709 },
  keepNext: o.keepNext,
});
const PC = (text, o = {}) => new Paragraph({ children: runs(text, o.run || {}), alignment: AlignmentType.CENTER, spacing: { line: 360, after: o.after ?? 0, before: o.before ?? 0 } });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)], pageBreakBefore: false });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(t)] });
const B = (text, lvl = 0) => new Paragraph({ numbering: { reference: "bullets", level: lvl }, children: runs(text), alignment: AlignmentType.JUSTIFIED, spacing: { line: 360, after: 60 } });
const N = (text, ref = "num") => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(text), alignment: AlignmentType.JUSTIFIED, spacing: { line: 360, after: 60 } });
const brk = () => new Paragraph({ children: [new PageBreak()] });

let nFig = 0, nTab = 0;
const border = { style: BorderStyle.SINGLE, size: 4, color: "A6A6A6" };
const borders = { top: border, bottom: border, left: border, right: border };

function caption(kind, title) {
  const n = kind === "Figura" ? ++nFig : ++nTab;
  return new Paragraph({ keepNext: true, alignment: AlignmentType.CENTER, spacing: { before: 200, after: 80, line: 240 },
    children: [new TextRun({ text: `${kind} ${n} – `, bold: true, size: 20 }), new TextRun({ text: title, size: 20 })] });
}
const fonte = (t = "Elaborado pela equipe (2026), a partir de Open Repair Alliance (2025).") =>
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240, line: 240 }, children: [new TextRun({ text: "Fonte: ", bold: true, size: 18 }), new TextRun({ text: t, size: 18 })] });

function table(headers, rows, widths, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const cell = (txt, head, i) => new TableCell({
    borders, width: { size: widths[i], type: WidthType.DXA },
    shading: head ? { fill: "DCE6F2", type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    children: [new Paragraph({ alignment: (o.align && o.align[i]) || (head ? AlignmentType.CENTER : AlignmentType.LEFT), spacing: { line: 240 },
      children: runs(String(txt), { size: o.size || 19, bold: head || undefined }) })],
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, true, i)) }),
      ...rows.map((r) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, false, i)) }))],
  });
}

function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}
function figura(nome, titulo, larguraPx = 600) {
  const file = path.join(FIG, nome + ".png");
  const { w, h } = pngSize(file);
  const width = larguraPx, height = Math.round((h / w) * larguraPx);
  return [caption("Figura", titulo),
    new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, children: [new ImageRun({ type: "png", data: fs.readFileSync(file), transformation: { width, height }, altText: { title: titulo, description: titulo, name: nome } })] }),
    fonte()];
}
function nota(titulo, texto) {
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: W, type: WidthType.DXA },
      borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.SINGLE, size: 24, color: "2A78D6" } },
      shading: { fill: "EEF4FB", type: ShadingType.CLEAR, color: "auto" }, margins: { top: 100, bottom: 100, left: 180, right: 140 },
      children: [new Paragraph({ spacing: { line: 276, after: 60 }, children: [new TextRun({ text: titulo, bold: true, size: 20, color: "1C5CAB" })] }),
        new Paragraph({ spacing: { line: 276 }, alignment: AlignmentType.JUSTIFIED, children: runs(texto, { size: 20 }) })],
    })] })],
  });
}
const sp = () => new Paragraph({ spacing: { after: 120 }, children: [] });

// ---------------------------------------------------------------- dados
const st = Object.fromEntries(R.status_categoria.map((r) => [r.categoria, r]));
const idade = Object.fromEntries(R.idade_desc.map((r) => [r.categoria, r]));
const cats = ["Notebook", "Celular", "Tablet", "Desktop"];
const nl = R.nao_liga;
const barr = R.barreiras; const nInf = R.n_fim_vida - barr["Não informada"];
const anual = Object.fromEntries(R.anual.map((r) => [r.ano_evento, r.Notebook + r.Celular + r.Tablet + r.Desktop]));
const quedaPand = (1 - anual[2020] / anual[2019]) * 100;
const tf = Object.fromEntries(R.taxa_por_falha.map((r) => [r.categoria_falha, r]));
const fc = R.falha_por_cat;
const marcas = Object.fromEntries(R.marcas.map((r) => [r.marca, r]));

// ---------------------------------------------------------------- conteúdo
const C = [];

// Capa
C.push(PC("UNIBALSAS – CENTRO UNIVERSITÁRIO DE BALSAS", { run: { bold: true } }), PC("Curso de Sistemas de Informação"), PC("Disciplina de Big Data e Ciência de Dados"),
  new Paragraph({ spacing: { before: 1600 }, children: [] }),
  ...["Rafael R Pontara", "Matheus de Brito Araujo", "Pedro Paulo T. Araujo"].map((n) => PC(n)),
  new Paragraph({ spacing: { before: 1800 }, children: [] }),
  PC("Seleção, Preparação, Análise Exploratória e Modelagem Preditiva de Dados para o Assistente de Diagnóstico Inteligente da Inforcell", { run: { bold: true, size: 28 }, after: 200 }),
  PC("Documento Técnico do Projeto — Versão 3.0"), PC("TED 01 (Seleção do conjunto de dados) + TED 02 (Limpeza, saneamento e análise exploratória) + TED 03 (Modelagem preditiva e Machine Learning)", { run: { size: 22 } }),
  new Paragraph({ spacing: { before: 3000 }, children: [] }), PC("Balsas – MA"), PC("2026"), brk());

// Folha de rosto
C.push(...["Rafael R Pontara", "Matheus de Brito Araujo", "Pedro Paulo T. Araujo"].map((n) => PC(n)),
  new Paragraph({ spacing: { before: 2200 }, children: [] }),
  PC("Seleção, Preparação, Análise Exploratória e Modelagem Preditiva de Dados para o Assistente de Diagnóstico Inteligente da Inforcell", { run: { bold: true, size: 26 }, after: 600 }),
  new Paragraph({ indent: { left: 4535 }, alignment: AlignmentType.JUSTIFIED, spacing: { line: 240 }, children: runs("Documento técnico apresentado à disciplina de Big Data e Ciência de Dados do curso de Sistemas de Informação da UNIBALSAS – Balsas/MA, como requisito parcial de avaliação, referente às Etapas 1 (TED 01 – definição do problema de dados e seleção do conjunto de dados), 2 (TED 02 – limpeza, saneamento e análise exploratória) e 3 (TED 03 – modelagem preditiva e Machine Learning) do projeto desenvolvido para a empresa Inforcell.", { size: 20 }) }),
  new Paragraph({ spacing: { before: 4200 }, children: [] }), PC("Balsas – MA"), PC("2026"), brk());

// Histórico de versões
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: "HISTÓRICO DE VERSÕES", bold: true })] }),
  table(["Versão", "Etapa", "Data", "Conteúdo"], [
    ["1.0", "TED 01", "2026", "Definição do problema de dados, necessidades de dados e seleção do conjunto Open Repair Data."],
    ["2.0", "TED 02", "26/09/2026", "Inspeção da base, revisão da seleção (Seção 6), limpeza e saneamento com justificativas (Seções 7 a 10), análise exploratória (Seções 11 e 12) e relação com o problema e o MVP (Seção 13). Seções 1 a 5 preservadas da versão 1.0."],
    ["3.0", "TED 03", "03/11/2026", "Definição do problema de Machine Learning (Seção 15), preparação para a modelagem (16), técnicas e processo de treinamento (17), resultados (18), interpretação (19), limitações (20), relação com o projeto e o MVP (21), registro de decisões e reprodução (22) e considerações finais (23). Seções 1 a 14 preservadas da versão 2.0."],
  ], [900, 1100, 1300, 5771], { align: [AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] }),
  brk(),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: "SUMÁRIO", bold: true })] }),
  new TableOfContents("Sumário", { hyperlink: true, headingStyleRange: "1-2" }), brk());

// ============================ PARTE I — TED 01 (preservada) ============================
C.push(H1("1 IDENTIFICAÇÃO DA EQUIPE"),
  P("Este documento apresenta as Etapas 1 (TED 01), 2 (TED 02) e 3 (TED 03) do projeto desenvolvido na disciplina de Big Data e Ciência de Dados. A Etapa 1 tratou da seleção e avaliação de um conjunto de dados público relacionado ao problema e ao MVP definidos anteriormente na disciplina de Inteligência Artificial, para a empresa Inforcell, assistência técnica de celulares e computadores localizada em Balsas – MA. A Etapa 2 tratou da limpeza, do saneamento e da análise exploratória desse conjunto de dados, e a Etapa 3 desenvolve e avalia um modelo de Machine Learning sobre os dados preparados."),
  table(["Item", "Descrição"], [
    ["Instituição", "UNIBALSAS – Balsas/MA"], ["Curso", "Sistemas de Informação"], ["Disciplina", "Big Data e Ciência de Dados"],
    ["Organização / público-alvo atendido", "Inforcell — Assistência técnica de celulares e computadores, Balsas/MA"],
    ["Etapas", "TED 01 — Seleção e avaliação de um conjunto de dados; TED 02 — Limpeza, saneamento e análise exploratória; TED 03 — Modelagem preditiva e Machine Learning"],
    ["Repositório", "Estrutura descrita no Apêndice A (docs/, data/raw, data/processed, notebooks/TED02 e TED03, src/, models/)"],
  ], [3000, 6071]), sp(),
  P("**Equipe:**", { noIndent: true }),
  B("Rafael R Pontara"), B("Matheus de Brito Araujo"), B("Pedro Paulo T. Araujo"), sp(),
  P("**Nota:** Pedro Junior Ubirajara participou das Etapas 1 (TED 01) e 2 (TED 02) e deixou o projeto antes da Etapa 3; a autoria das Seções 1 a 14, preservadas das versões anteriores, inclui sua contribuição.", { noIndent: true }));

C.push(H1("2 RETOMADA DO PROBLEMA E DO MVP"),
  P("Conforme definido na Etapa 1 da disciplina de Inteligência Artificial, o problema identificado junto à Inforcell está relacionado à dificuldade dos técnicos em identificar as possíveis causas de falhas em equipamentos eletrônicos. Um mesmo sintoma — por exemplo, “um computador não liga” ou “um celular não liga” — pode estar relacionado a diferentes fatores (alimentação, bateria, memória, placa-mãe, software, entre outros), exigindo que o técnico realize uma sequência de testes até chegar à causa provável. Esse processo depende hoje quase exclusivamente da experiência individual de cada técnico, e o uso do WhatsApp como ferramenta de apoio ao gerenciamento das Ordens de Serviço (OS) dificulta a centralização das informações."),
  P("A solução proposta é uma plataforma web que une o gerenciamento de Ordens de Serviço a um assistente de Inteligência Artificial capaz de conduzir o técnico por um diagnóstico estruturado, passo a passo: identificação do equipamento, levantamento das possíveis causas, definição do teste a ser realizado, registro do resultado, interpretação do resultado pela IA e indicação do próximo procedimento, até o registro do diagnóstico final."),
  P("O Produto Mínimo Viável (MVP) definido para validar essa proposta contempla o cadastro e autenticação de usuários, o cadastro de clientes e equipamentos, a criação e o acompanhamento de Ordens de Serviço, e um assistente de IA capaz de receber o sintoma relatado, apresentar possíveis causas, conduzir o diagnóstico guiado e registrar o histórico de diagnósticos e reparos."));

C.push(H1("3 NECESSIDADES DE DADOS DA SOLUÇÃO"),
  P("A partir do problema e do MVP retomados acima, a equipe identificou que, para apoiar, desenvolver e avaliar a solução proposta, seriam necessários dados que representem:"),
  B("sintomas relatados por equipamentos eletrônicos, descritos em linguagem natural, da forma como um técnico ou cliente relataria o problema;"),
  B("o tipo ou categoria de falha associada a cada sintoma relatado;"),
  B("a categoria e a marca do equipamento envolvido (celular, notebook, desktop, tablet);"),
  B("o resultado do reparo (se o problema foi solucionado ou não), permitindo aproximar o comportamento de um “diagnóstico final”."),
  P("Esses dados apoiariam tanto a compreensão do problema (quais sintomas são mais recorrentes e quais causas eles costumam ter) quanto uma futura avaliação da qualidade do diagnóstico guiado proposto no MVP."));

C.push(H1("4 CONJUNTO DE DADOS SELECIONADO"),
  H2("4.1 Identificação"),
  table(["Item", "Descrição"], [
    ["Conjunto de dados", "Open Repair Data"],
    ["Fonte / origem", "Open Repair Alliance (ORA) — aliança internacional de organizações que promovem eventos comunitários de reparo (Restart Project, Repair Café Foundation, Anstiftung, entre outras), seguindo o padrão aberto Open Repair Data Standard (ORDS)"],
    ["Licença / condição de uso", "Dados abertos, disponibilizados publicamente pela Open Repair Alliance (licença CC BY-SA 4.0, confirmada na TED 02)"],
    ["Portal oficial (download)", "https://openrepair.org/open-data/"], ["Repositório com os dados brutos", "https://github.com/openrepair/data"],
  ], [3000, 6071]), sp(),
  H2("4.2 Descrição do conjunto de dados"),
  P("A base reúne registros de tentativas de reparo de diversos tipos de dispositivos eletrônicos (celulares, laptops, desktops, tablets e outros eletrônicos de consumo), coletados por voluntários durante eventos reais de reparo comunitário. Cada registro contém uma descrição textual do problema relatado pelo dono do equipamento, a classificação da falha atribuída por quem realizou o reparo e o desfecho da tentativa de reparo. Entre as fontes públicas avaliadas pela equipe, esta é a que mais se aproxima da lógica “sintoma → causa → resultado” que o assistente de IA da Inforcell pretende automatizar."),
  H2("4.3 Principais atributos/variáveis (conforme levantamento da TED 01)"),
  table(["Atributo", "Descrição"], [
    ["id_ords", "Identificador único do registro"], ["data_provider", "Organização que coletou o dado"], ["country", "País do evento de reparo"],
    ["event_date", "Data do evento de reparo"], ["product_category", "Categoria do produto (ex.: celular, laptop, desktop)"], ["brand", "Marca do equipamento"],
    ["year_of_manufacture", "Ano de fabricação do equipamento"], ["problem", "Descrição textual do sintoma relatado (dado em linguagem natural)"],
    ["fault_type", "Tipo de falha identificada pelo reparador"], ["repair_status", "Resultado do reparo (consertado, não consertado, entre outros)"], ["language", "Idioma da descrição do problema"],
  ], [3000, 6071]), sp(),
  nota("Atualização da versão 2.0", "A inspeção realizada na TED 02 mostrou que, na versão atual da base agregada (ORDS v0.3), os campos **fault_type** e **language** não existem, e que o identificador se chama **id**. O tipo de falha está disponível apenas nas bases rotuladas das ~~quests~~ publicadas no mesmo repositório. A revisão está registrada e justificada na Seção 6."),
  H2("4.4 Quantidade de registros e formato"),
  P("O conjunto consolidado da Open Repair Alliance reúne aproximadamente 48 mil registros de reparos, distribuídos entre diversas categorias de produto. Como o problema da Inforcell está focado em celulares e computadores, a equipe pretende filtrar apenas os registros cuja product_category corresponda a essas categorias, o que resulta em um subconjunto de alguns milhares de registros. Os dados são disponibilizados em arquivos no formato CSV, seguindo a estrutura do padrão ORDS."),
  nota("Atualização da versão 2.0", `A exportação utilizada na TED 02 (202507, dados até 31/07/2025) possui **305.649** registros no total e **23.612** nas quatro categorias de interesse (Mobile, Laptop, Desktop computer e Tablet), quantidade bem superior à estimada na TED 01.`),
  H2("4.5 Justificativa da escolha"),
  P("Diferentemente de bases voltadas à manutenção industrial, que tratam de sensores de máquinas fabris, a Open Repair Data trata diretamente de eletrônicos de consumo com falhas relatadas em linguagem natural, o que é conceitualmente próximo do que o assistente de IA da Inforcell precisa processar: um sintoma é relatado e o sistema precisa associá-lo a uma causa provável. Além disso, trata-se de uma base pública, com licença aberta, documentada, atualizada e de fácil acesso, atendendo aos critérios de disponibilidade e confiabilidade exigidos pela atividade."),
  H2("4.6 Relação entre os dados e o problema"),
  P("Os campos problem (sintoma em texto livre) e fault_type (causa/categoria da falha) permitem investigar, de forma exploratória, como sintomas relatados se relacionam a diferentes causas — exatamente a dificuldade central identificada na Inforcell. O campo repair_status também permite observar a taxa de sucesso do diagnóstico, algo comparável à métrica-chave de “redução do tempo/erro no diagnóstico” definida no Lean Canvas do projeto."),
  H2("4.7 Limitações identificadas inicialmente"),
  B("Os dados vêm de eventos comunitários de reparo (voluntários), não de uma assistência técnica comercial como a Inforcell — o contexto e o nível de detalhe técnico podem diferir."),
  B("A maior parte das descrições de sintoma está em inglês ou em outros idiomas (alemão, holandês), exigindo tradução ou tratamento para uso no contexto brasileiro."),
  B("O campo fault_type é uma classificação geral, sem o nível de granularidade técnica (ex.: componente elétrico específico) que o assistente da Inforcell pretende oferecer."),
  B("Há distribuição desigual entre categorias de produto, podendo haver poucos registros no subconjunto específico de celulares e computadores após o filtro."),
  P("Conforme observado no próprio enunciado da atividade, essa seleção é preliminar: caso as etapas seguintes revelem que a base não é adequada, a equipe poderá substituí-la, desde que a decisão seja apresentada e justificada."));

C.push(H1("5 CONSIDERAÇÕES DA ETAPA 1"),
  P("A seleção do conjunto de dados Open Repair Data permite à equipe avançar para as próximas etapas da disciplina de Big Data e Ciência de Dados com uma base pública, relacionada ao domínio do problema da Inforcell, ainda que com limitações que deverão ser tratadas ao longo do processo de análise (idioma, granularidade da classificação de falhas e diferença de contexto entre reparo comunitário e assistência técnica comercial). A relação Problema → Necessidades de dados → Base selecionada → Justificativa de utilização fica assim demonstrada, cabendo às próximas etapas aprofundar o tratamento e a análise exploratória dos dados."));

// ============================ PARTE II — TED 02 ============================
C.push(brk(), H1("6 REVISÃO DA SELEÇÃO APÓS A INSPEÇÃO DA BASE"),
  P("Antes de qualquer tratamento, a equipe carregou a versão mais recente da base diretamente do repositório oficial e comparou sua estrutura com o que foi descrito na TED 01. A comparação revelou diferenças relevantes, que levaram à revisão de algumas decisões. Conforme orienta o enunciado da TED 02, as alterações estão registradas e justificadas na Tabela 1."),
  caption("Tabela", "Decisões da TED 01 revistas na TED 02"),
  table(["Decisão / premissa da TED 01", "Situação encontrada", "Decisão na TED 02 e justificativa"], [
    ["Base com ~48 mil registros", "305.649 registros (ORDS v0.3, exportação 202507); 23.612 nas 4 categorias de interesse", "Manter a base. O volume maior é positivo: amplia a amostra de celulares e computadores."],
    ["Campo fault_type na base principal", "Inexistente na base agregada v0.3", "Incorporar as bases rotuladas das quests MobiFix (celulares), CompCat (notebooks e desktops) e TabiCat (tablets), do mesmo repositório, que trazem o tipo de falha atribuído por voluntários. Sem esse campo não há par sintoma → causa."],
    ["Campo language", "Inexistente", "Criar o atributo idioma por detecção automática, validada contra os rótulos manuais da TabiCat (Seção 9.3)."],
    ["Uma única base", "As informações necessárias estão em arquivos distintos", "Trabalhar com duas bases complementares: **Base A – Reparos** (contexto do atendimento) e **Base B – Falhas rotuladas** (sintoma → tipo de falha)."],
    ["Categorias: celular, laptop, desktop, tablet", "A categoria Tablet inclui GPS, e-readers e estações meteorológicas", "Manter as 4 categorias, removendo os itens que não são tablets (Seção 9.1)."],
    ["Classificação de falha única", "Quests usam listas diferentes (26 e 20 tipos)", "Criar taxonomia unificada em português com 14 categorias (Seção 9.2), alinhada à rotina de bancada da Inforcell."],
  ], [2400, 2800, 3871], { size: 18 }),
  fonte("Elaborado pela equipe (2026)."),
  P("A decisão de **manter** o Open Repair Data foi tomada porque, mesmo com as diferenças, a combinação das duas bases atende às quatro necessidades de dados da Seção 3: sintoma em linguagem natural, tipo de falha, categoria/marca do aparelho e desfecho do reparo. Nenhuma outra base pública avaliada oferece o par sintoma → falha para eletrônicos de consumo."));

C.push(H1("7 PROCESSO DE PREPARAÇÃO DOS DADOS"),
  P("O processo seguiu o fluxo proposto na atividade: **dados brutos → inspeção → limpeza e preparação → análise exploratória → descoberta de padrões → compreensão do problema**. Todo o processo foi implementado em Python (pandas, NumPy, SciPy e Matplotlib) e organizado para ser reprodutível:"),
  B("**src/preparacao.py** — módulo com todas as regras de carga, limpeza, harmonização e validação. Concentrar as regras em um único módulo permite reaplicá-las a novas exportações da ORA e, futuramente, aos dados de Ordens de Serviço da própria Inforcell;"),
  B("**notebooks/TED02/01_carga_inspecao.ipynb** — carga e diagnóstico de qualidade da base bruta;"),
  B("**notebooks/TED02/02_limpeza_saneamento.ipynb** — aplicação dos tratamentos, validações e exportação das bases tratadas;"),
  B("**notebooks/TED02/03_analise_exploratoria.ipynb** — estatísticas, gráficos e testes. Todos os números citados neste documento são exportados pelo notebook para reports/resultados_eda.json."),
  P("A base bruta foi lida integralmente como texto. Essa escolha é intencional: impede que o pandas converta valores automaticamente (por exemplo, transformando o marcador “????” em erro ou anos em números decimais) e faz com que cada conversão de tipo seja uma decisão explícita e documentada da etapa de limpeza."),
  P("A versão utilizada é identificada por: repositório github.com/openrepair/data, arquivo aggregated/202507/OpenRepairData_v0.3_aggregate_202507.csv (commit 90eba807, de 22/03/2026) e arquivos da pasta quests/. O arquivo completo tem 58 MB; por isso, o repositório do projeto versiona apenas o subconjunto bruto das 4 categorias (4 MB, sem nenhuma alteração) e um script que baixa o arquivo completo quando necessário."));

C.push(H1("8 PROBLEMAS DE QUALIDADE IDENTIFICADOS"),
  P("A inspeção (notebook 01) avaliou estrutura, tipos, completude, duplicidade, consistência e plausibilidade dos valores. A Tabela 2 resume os problemas encontrados e a Figura 1 mostra a completude dos campos."),
  caption("Tabela", "Problemas de qualidade identificados na base bruta (23.612 registros, 4 categorias)"),
  table(["#", "Problema", "Campo(s)", "Evidência"], [
    ["1", "Estrutura diferente da prevista", "base agregada", "Sem fault_type e language; 305.649 registros no total"],
    ["2", "Tipos de dados inadequados", "event_date, year_of_manufacture, product_age", "Datas e anos armazenados como texto"],
    ["3", "Dados ausentes em grande volume", "year_of_manufacture, product_age", "69,1% e 69,2% ausentes"],
    ["4", "Descrição do sintoma ausente", "problem", "24,1% ausente; há também placeholders (“?”, “Unknown”, “-”)"],
    ["5", "Ausência estrutural", "repair_barrier_if_end_of_life", "95% ausente — só se aplica a itens em fim de vida"],
    ["6", "Registros duplicados", "todos, exceto id", "2.671 linhas idênticas; IDs sequenciais de reparos em lote"],
    ["7", "Valores impossíveis", "year_of_manufacture", "Ano posterior ao evento; notebook de 1957; tablet de 1969"],
    ["8", "Outliers de idade", "product_age", "161 registros acima do limite de Tukey (19 anos); máximo de 67 anos"],
    ["9", "Categorização inconsistente", "product_category", "“Tablet” contém GPS, e-readers, estações meteorológicas, ciclocomputadores"],
    ["10", "Grafias diferentes da mesma marca", "brand", "942 marcas distintas; ex.: Apple, apple, appel, iPhone, I phone"],
    ["11", "Texto multilíngue e com ruído", "problem", "EN, DE, FR, NL, NO, IT, ES, DA; URLs; frases de 2 a 6.022 caracteres"],
    ["12", "Atributos redundantes/irrelevantes", "product_category_id, product_age, group_identifier", "Relação 1:1 com a categoria; idade derivável do ano; grupo mistura códigos e nomes"],
    ["13", "Marcadores textuais de ausência", "year_of_manufacture (TabiCat)", "479 valores “????”"],
    ["14", "Rótulos pouco informativos", "fault_type (quests)", "“Unknown” e “Poor data” em 26,1% dos registros rotulados"],
  ], [450, 2500, 2450, 3671], { size: 17, align: [AlignmentType.CENTER] }),
  fonte("Elaborado pela equipe (2026), a partir do notebook 01."),
  ...figura("fig01_ausentes", "Percentual de valores ausentes por campo (base bruta)", 560),
  P("**Interpretação.** A ausência de ano de fabricação em cerca de 7 de cada 10 registros é o problema de completude mais sério, porque limita qualquer conclusão sobre a relação entre idade do aparelho e sucesso do reparo. Já a ausência da barreira de reparo não é um defeito: o campo só faz sentido quando o aparelho foi declarado em fim de vida. Diferenciar ausência estrutural de ausência real evitou um tratamento equivocado (preencher barreiras para aparelhos que foram consertados)."));

C.push(H1("9 TRATAMENTOS REALIZADOS E JUSTIFICATIVAS"),
  P("Cada problema foi avaliado individualmente, considerando o uso que o dado terá no projeto. Em vez de aplicar correções automáticas, a equipe escolheu, para cada caso, entre corrigir, transformar, preencher, remover ou manter a informação."),
  H2("9.1 Base A – Reparos"),
  caption("Tabela", "Tratamentos aplicados à Base A e justificativas"),
  table(["Problema", "Ação", "Tratamento", "Justificativa"], [
    ["Itens que não são tablets", "Remover", "1.079 registros removidos por palavras-chave na categoria do parceiro (GPS, e-reader, navegação, estação meteorológica etc.)", "Não são atendidos pela Inforcell e reduziriam artificialmente a taxa de conserto de tablets. Não há categoria adequada para recategorizá-los."],
    ["Tipos inadequados", "Transformar", "Data → datetime; anos e idade → inteiro anulável (Int64)", "Permite cálculos temporais; o tipo Int64 preserva a ausência sem transformá-la em zero."],
    ["Ano impossível", "Corrigir para ausente", "18 anos (futuros ou anteriores à tecnologia) convertidos em nulo", "O restante do registro (sintoma, desfecho) continua útil; remover a linha desperdiçaria informação."],
    ["Idade redundante", "Transformar", "Idade recalculada a partir do ano validado; product_age descartada", "Garante consistência entre as duas colunas."],
    ["69% sem idade", "Manter ausente", "Sem imputação; análises de idade usam só registros com valor", "Imputar média ou mediana criaria uma idade artificial em 2 de cada 3 registros e apagaria a relação idade × desfecho que se pretende estudar."],
    ["Outliers de idade", "Manter e sinalizar", "Flag idade_atipica (> 20 anos; 77 registros)", "São aparelhos antigos reais (Commodore, Nokia antigos), não erros de digitação."],
    ["Marcas com grafias distintas", "Corrigir", "Dicionário de sinônimos: 942 → 685 marcas; “Unknown” → “Desconhecida”", "Sem padronização, a Apple apareceria dividida em mais de dez marcas."],
    ["Texto com ruído", "Transformar", "Normalização de espaços, remoção de URLs e artefatos; placeholders → nulo", "Descrições como “?” não informam sintoma e distorceriam a análise textual."],
    ["Texto ausente", "Manter e sinalizar", "Flag tem_descricao", "Esses registros ainda servem para estatísticas de desfecho, marca e idade."],
    ["Barreira ausente", "Preencher", "“Não se aplica” (não está em fim de vida) ou “Não informada”", "Distingue ausência estrutural de ausência real."],
    ["Duplicatas", "Manter e sinalizar", "Flag registro_em_lote (2.669 cópias)", "Têm IDs distintos e representam aparelhos diferentes recondicionados em lote — removê-las subestimaria o volume real."],
    ["Status “Unknown”", "Manter", "Excluído do denominador da taxa de conserto", "Não é possível afirmar se foi consertado ou não."],
    ["Colunas redundantes", "Remover", "product_category_id, group_identifier, partner_product_category, product_age", "Não agregam informação após o filtro de escopo."],
  ], [1750, 1200, 2860, 3261], { size: 16 }),
  fonte("Elaborado pela equipe (2026)."),
  H2("9.2 Base B – Falhas rotuladas e taxonomia unificada"),
  P(`A Base B foi construída a partir das três quests (MobiFix: 1.916 celulares; CompCat: 3.510 notebooks e 583 desktops; TabiCat: 647 tablets), totalizando ${int(R.n_falhas)} registros. Os tratamentos específicos foram: conversão do marcador “????” em nulo; conversão da idade da MobiFix em ano de fabricação, padronizando com as demais; e, principalmente, a **harmonização dos rótulos de falha**. As quests usam listas diferentes (41 rótulos distintos no total), que foram agrupados em 14 categorias em português (Tabela 4). O agrupamento seguiu dois critérios: (a) componentes que o técnico testa com o mesmo procedimento de bancada ficam juntos (por exemplo, conector de carga, P2 e bandeja de chip em “Conectores e portas”); e (b) categorias finais devem corresponder a “ramos” do diagnóstico guiado do MVP.`),
  caption("Tabela", "Taxonomia unificada de falhas"),
  table(["Categoria unificada", "Rótulos originais agrupados"], [
    ["Energia / bateria", "Power/battery, Charger"], ["Tela", "Screen, Integrated screen"],
    ["Conectores e portas", "USB/charging port, Ports/slots/connectors, Headphone jack, Sim card slot, Memory card slot"],
    ["Botões / teclado / touchpad", "On/Off button, Volume buttons, Other buttons, Integrated keyboard, Integrated pointing device"],
    ["Áudio / câmera / mídia", "Speaker/amplifier, Microphone, Camera, Integrated media component, Bluetooth, Integrated optical drive"],
    ["Placa / dano interno", "System board, Internal damage"], ["Dano por líquido", "Liquid damage"], ["Carcaça / dano externo", "External damage, Case/chassis"],
    ["Armazenamento", "Storage problem, Internal storage"],
    ["Software / sistema", "Software update, Software issue, Operating system, Configuration, Help/configuration, Virus/malware, Access issue"],
    ["Inicialização (boot)", "Boot, Stuck booting"], ["Desempenho / superaquecimento", "Performance, Overheating"],
    ["Outros", "Other, Multiple"], ["Indeterminado", "Unknown, Poor data"],
  ], [3000, 6071], { size: 18 }),
  fonte("Elaborado pela equipe (2026), a partir das listas de falhas das quests MobiFix, CompCat e TabiCat."),
  P(`Para a próxima etapa (modelagem), foi criada a flag **apto_modelagem**, que exclui: registros sem descrição (não é possível aprender sintoma → causa sem texto), rótulos “Indeterminado” (não informam a causa) e duplicatas de mesmo texto + mesmo rótulo + mesmo aparelho (${int(1016)} casos), que fariam um futuro modelo “decorar” frases repetidas e superestimar sua acurácia. Resultaram **${int(R.n_modelagem)} registros aptos**. Os registros excluídos continuam na base descritiva, pois a proporção de rótulos indeterminados é, em si, um achado relevante (Seção 11.5).`),
  H2("9.3 Criação e validação do atributo idioma"),
  P("Como a coluna language não existe na versão atual e as bibliotecas de detecção de idioma não estavam disponíveis no ambiente, foi implementado um detector transparente baseado em palavras funcionais de oito idiomas (por exemplo, “the/not/won’t” no inglês; “niet/werkt” no holandês; “nicht/geht” no alemão). Quando o texto é curto ou ambíguo, usa-se o idioma oficial do país do evento, apenas para países monolíngues. A validação contra os 647 rótulos de idioma feitos manualmente na TabiCat resultou em **96,0% de cobertura e 96,5% de acurácia**, desempenho considerado suficiente para a análise exploratória."));

C.push(H1("10 DESCRIÇÃO DA BASE APÓS O TRATAMENTO"),
  caption("Tabela", "Bases tratadas disponibilizadas em data/processed"),
  table(["Arquivo", "Registros", "Colunas", "Conteúdo"], [
    ["reparos_limpo.csv", int(R.n_reparos), "19", "Base A: aparelho, marca, idade, desfecho, barreira, texto, idioma e flags de qualidade"],
    ["falhas_rotuladas_limpo.csv", int(R.n_falhas), "20", "Base B completa: sintoma, tipo de falha original e unificado, flags"],
    ["falhas_modelagem.csv", int(R.n_modelagem), "7", "Subconjunto da Base B apto para modelagem (TED 03)"],
    ["taxonomia_falhas.csv", "41", "3", "Mapeamento rótulo original → categoria unificada"],
  ], [2500, 1100, 900, 4571], { size: 18, align: [AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.CENTER] }),
  fonte("Elaborado pela equipe (2026)."),
  P(`A **Base A** passou de 23.612 para ${int(R.n_reparos)} registros (apenas a remoção dos itens fora de escopo reduziu linhas). Ela cobre eventos de ${R.periodo[0].split("-").reverse().join("/")} a ${R.periodo[1].split("-").reverse().join("/")}, em ${R.n_paises} países, com ${int(R.n_marcas)} marcas após a padronização. ${fmt(R.pct_com_texto)}% dos registros possuem descrição de sintoma e ${fmt(R.pct_com_idade)}% possuem idade válida. A Tabela 6 apresenta o dicionário de dados.`),
  caption("Tabela", "Dicionário de dados da Base A (reparos_limpo.csv)"),
  table(["Coluna", "Tipo", "Descrição"], [
    ["id", "texto", "Identificador ORDS original"], ["fonte / pais", "categórico", "Organização coletora / país (ISO-3)"],
    ["categoria", "categórico", "Notebook, Celular, Tablet, Desktop"], ["marca", "categórico", "Marca padronizada (“Desconhecida” quando não informada)"],
    ["ano_fabricacao / idade_anos", "inteiro (nulo permitido)", "Ano validado e idade no momento do reparo"], ["faixa_idade", "categórico ordinal", "0–2, 3–5, 6–8, 9–12, 13+ anos"],
    ["status_reparo", "categórico", "Consertado, Reparável (não concluído), Fim de vida, Desconhecido"], ["consertado", "binário (nulo permitido)", "1 = consertado; 0 = não; nulo = desconhecido"],
    ["barreira_reparo", "categórico", "Motivo do fim de vida, “Não informada” ou “Não se aplica”"], ["data_evento / ano_evento", "data / inteiro", "Data do evento de reparo"],
    ["problema", "texto", "Descrição do sintoma após limpeza"], ["n_palavras", "inteiro", "Tamanho da descrição"], ["idioma", "categórico", "Idioma detectado (Seção 9.3)"],
    ["tem_descricao, idade_atipica, registro_em_lote", "booleano", "Flags de qualidade"],
  ], [2900, 1900, 4271], { size: 17 }),
  fonte("Elaborado pela equipe (2026)."),
  P("A **Base B** mantém a mesma lógica e acrescenta as colunas quest, tipo_falha_original, categoria_falha, rotulo_informativo, duplicata_texto e apto_modelagem. O dicionário completo está em data/processed/README.md."));

// ---------------- EDA
C.push(H1("11 ANÁLISE EXPLORATÓRIA DE DADOS"),
  P("A análise foi planejada a partir da natureza das variáveis: quase todas são **categóricas** (aparelho, marca, desfecho, tipo de falha), há uma variável numérica discreta (idade), uma variável temporal (data do evento) e uma variável textual (sintoma). Por isso foram priorizadas tabelas de frequência, gráficos de barras, boxplots, séries temporais e o V de Cramér para medir associação entre categorias — a correlação de Pearson não seria adequada para esse tipo de dado."),
  H2("11.1 Estatísticas descritivas"),
  caption("Tabela", "Distribuição dos registros e do desfecho por categoria (Base A, %)"),
  table(["Categoria", "Registros", "% da base", "Consertado", "Reparável", "Fim de vida", "Desconhecido", "Taxa de conserto*"],
    cats.map((c) => { const f = R.freq_categoria.find((x) => x.categoria === c); const s = st[c];
      return [c, int(f.registros), fmt(f["%"]), fmt(s["Consertado"]), fmt(s["Reparável (não concluído)"]), fmt(s["Fim de vida"]), fmt(s["Desconhecido"]), `**${fmt(s["Taxa de conserto*"])}**`]; }),
    [1250, 1100, 1000, 1150, 1100, 1100, 1250, 1121], { size: 17, align: [AlignmentType.LEFT, ...Array(7).fill(AlignmentType.RIGHT)] }),
  fonte("Elaborado pela equipe (2026). *Taxa de conserto = consertados ÷ (total − desconhecidos)."),
  P(`Os notebooks formam mais da metade da base (53,9%), seguidos de celulares (23,0%). A taxa geral de conserto é de **${fmt(R.taxa_conserto_geral)}%**. Notebooks, celulares e desktops têm taxas semelhantes (56,7% a 58,3%), enquanto os **tablets destoam, com apenas 45,4%** e a maior proporção de fim de vida (22,0%).`),
  caption("Tabela", "Estatísticas descritivas da idade do aparelho no momento do reparo (anos)"),
  table(["Categoria", "n", "Média", "Mediana", "Moda", "Desvio-padrão", "Q1", "Q3", "IQR", "Máx.", "Assimetria"],
    [...cats.map((c) => { const d = idade[c]; return [c, int(d.count), fmt(d.mean), fmt(d["50%"], 0), fmt(d.moda, 0), fmt(d.std), fmt(d["25%"], 0), fmt(d["75%"], 0), fmt(d.IQR, 0), fmt(d.max, 0), fmt(d.assimetria, 2)]; }),
      ["**Geral**", int(R.idade_geral.count), fmt(R.idade_geral.mean), fmt(R.idade_geral["50%"], 0), "–", fmt(R.idade_geral.std), fmt(R.idade_geral["25%"], 0), fmt(R.idade_geral["75%"], 0), fmt(R.idade_geral["75%"] - R.idade_geral["25%"], 0), fmt(R.idade_geral.max, 0), "–"]],
    [1150, 700, 750, 850, 700, 1100, 600, 600, 600, 700, 1321], { size: 16, align: [AlignmentType.LEFT, ...Array(10).fill(AlignmentType.RIGHT)] }),
  fonte("Elaborado pela equipe (2026). Apenas registros com ano de fabricação válido."),
  P(`Os celulares chegam ao reparo mais novos (mediana de 4 anos, moda de 2 anos) e os desktops mais velhos (mediana de 8 anos, maior dispersão). Em todas as categorias a média supera a mediana e a assimetria é positiva: a maioria dos aparelhos tem poucos anos, com uma cauda de aparelhos antigos. Para a descrição do sintoma, a mediana é de apenas **${fmt(R.palavras_desc["50%"], 0)} palavras** (média ${fmt(R.palavras_desc.mean)}, desvio-padrão ${fmt(R.palavras_desc.std)}), e ${fmt(R.pct_ate_5_palavras)}% das descrições têm até 5 palavras.`),
  H2("11.2 Distribuição das principais variáveis"),
  ...figura("fig02_status_categoria", "Desfecho do reparo por categoria de aparelho", 580),
  P("**O que se observa.** Em todas as categorias, pouco mais da metade dos aparelhos sai consertada e cerca de um quarto fica como “reparável” — o defeito foi entendido, mas o reparo não foi concluído no evento (normalmente por falta de peça ou de tempo). **Por que importa.** A faixa “reparável” representa casos em que o diagnóstico foi feito, mas faltou continuidade — exatamente o tipo de acompanhamento que as Ordens de Serviço do MVP devem registrar. Os tablets, com mais fim de vida, indicam equipamentos de reparo mais difícil (colados, com poucas peças disponíveis)."),
  ...figura("fig03_idade", "Distribuição da idade dos aparelhos e boxplot por categoria", 600),
  P("**O que se observa.** A distribuição da idade é assimétrica à direita, com mediana de 5 anos e picos em valores “redondos” (5 e 10 anos), sinal de que muitas idades foram estimadas pelo dono do aparelho, e não informadas com precisão. O boxplot mostra outliers acima de 20 anos em todas as categorias. **Por que importa.** A imprecisão e a alta ausência da idade reduzem seu valor como variável de entrada do assistente; ela pode ser registrada na OS, mas não deve ser um fator central do diagnóstico."),
  ...figura("fig04_palavras", "Distribuição do número de palavras na descrição do sintoma", 560),
  P("**O que se observa.** As descrições são muito curtas: quase metade tem até cinco palavras (“broken screen”, “won’t charge”, “slow”). **Por que importa.** Isso confirma a premissa do MVP: o relato inicial do sintoma é insuficiente para um diagnóstico. O assistente precisa fazer **perguntas complementares e propor testes**, em vez de tentar classificar a falha apenas a partir do texto inicial."),
  caption("Tabela", "Distribuição dos idiomas das descrições (Base A)"),
  table(["Idioma", "Inglês", "Alemão", "Francês", "Holandês", "Outros", "Indefinido", "Sem texto"],
    [["% dos registros", fmt(R.idiomas.en), fmt(R.idiomas.de), fmt(R.idiomas.fr), fmt(R.idiomas.nl), fmt(R.idiomas.no + R.idiomas.it + R.idiomas.es + R.idiomas.da), fmt(R.idiomas.indefinido), fmt(R.idiomas["sem texto"])]],
    [1600, 1000, 1000, 1000, 1100, 1000, 1150, 1221], { size: 18, align: [AlignmentType.LEFT, ...Array(7).fill(AlignmentType.RIGHT)] }),
  fonte("Elaborado pela equipe (2026)."),
  P("Somente cerca de metade das descrições em texto está em inglês; nenhuma está em português. Para o MVP, que atenderá técnicos brasileiros, isso indica que os dados servem para **entender padrões e construir a árvore de diagnóstico**, mas qualquer modelo textual precisará de tradução ou de um modelo multilíngue — e, idealmente, de dados reais da Inforcell em português."));

C.push(H2("11.3 Tendências temporais"),
  ...figura("fig05_serie_temporal", "Registros por ano e taxa de conserto por ano (2013–2025)", 620),
  P(`**O que se observa.** O volume cresce até 2019 (${int(anual[2019])} registros), cai ${fmt(quedaPand, 0)}% em 2020, quando os eventos presenciais foram suspensos pela pandemia, e volta a crescer a partir de 2022, chegando a ${int(anual[2024])} registros em 2024 (2025 está incompleto, com dados até julho). Os notebooks lideram em todos os anos e têm o maior salto em 2022, período que coincide com um programa de recondicionamento em lote (463 registros de “SSD novo + Windows 10” só em 2022). A partir de 2014, a taxa de conserto oscila entre 46% e 61%, sem tendência clara de melhora ou piora (o valor de 2013, 34%, se apoia em apenas 126 registros).`),
  P("**Por que importa.** A série mostra que a base reflete o funcionamento dos eventos (pandemia, programas de recondicionamento) e não apenas a frequência natural dos defeitos. Por isso, nas próximas etapas, a comparação entre períodos deve ser feita com cautela e, para treinar modelos, é recomendável uma divisão treino/teste que respeite o tempo."));

C.push(H2("11.4 Relações entre variáveis"),
  H3("11.4.1 Idade do aparelho × sucesso do reparo"),
  ...figura("fig06_idade_conserto", "Taxa de conserto por faixa de idade e categoria", 560),
  P(`**O que se observa.** Contrariando a expectativa inicial da equipe, a idade do aparelho quase não altera a taxa de conserto: considerando todas as categorias, a taxa varia apenas entre ${fmt(R.taxa_faixa_geral["3–5 anos"])}% (3–5 anos) e ${fmt(R.taxa_faixa_geral["13+ anos"])}% (13 anos ou mais). A correlação de Spearman entre idade e conserto é praticamente nula (ρ = ${fmt(R.spearman_idade_conserto.rho, 3)}; n = ${int(R.spearman_idade_conserto.n)}). O teste qui-quadrado indica diferença estatisticamente significativa entre as faixas (χ² = ${fmt(R.chi2_faixa_conserto.chi2)}; gl = 4; p = ${fmt(R.chi2_faixa_conserto.p, 3)}), mas o efeito é pequeno: com amostras grandes, pequenas diferenças já se tornam “significativas”. A diferença consistente está na **categoria**: tablets ficam abaixo das demais em todas as faixas.`),
  P("**Por que importa.** A variável que mais orienta o resultado não é a idade, e sim o **tipo de aparelho** e, como se verá na Seção 11.5, o **tipo de falha**. Para o assistente, isso significa priorizar a identificação correta do equipamento e da falha, em vez de usar a idade como atalho para prever se “vale a pena consertar”."),
  H3("11.4.2 Marca × volume × sucesso"),
  ...figura("fig07_marcas", "Marcas mais frequentes e respectivas taxas de conserto (marcas com 150 ou mais registros)", 610),
  P(`**O que se observa.** Apple (${int(marcas.Apple.n)}), HP (${int(marcas.HP.n)}) e Samsung (${int(marcas.Samsung.n)}) lideram em volume. A HP tem a maior taxa de conserto (${fmt(marcas.HP.taxa)}%) e a Samsung, a menor (${fmt(marcas.Samsung.taxa)}%). Parte dessa diferença vem da composição: a HP está concentrada em notebooks e a Samsung em celulares e tablets. Além disso, ${fmt(R.pct_marca_desconhecida)}% dos registros não informam a marca.`),
  P("**Por que importa.** A marca é relevante, mas está misturada com a categoria do aparelho; ela deve entrar no assistente como **informação complementar** (por exemplo, para sugerir testes e peças específicas do modelo), e o cadastro de equipamentos do MVP deve tornar a marca e o modelo campos obrigatórios, para não repetir a lacuna de 34% observada na base."),
  H3("11.4.3 Barreiras ao reparo"),
  ...figura("fig08_barreiras", "Barreiras informadas para aparelhos em fim de vida", 540),
  P(`**O que se observa.** Dos ${int(R.n_fim_vida)} aparelhos em fim de vida, ${fmt(barr["Não informada"] / R.n_fim_vida * 100)}% não têm barreira registrada. Entre os ${int(nInf)} informados, as principais são **peças caras demais** (${fmt(barr["Peças caras demais"] / nInf * 100, 0)}%) e **falta de informação de reparo** (${fmt(barr["Falta de informação de reparo"] / nInf * 100, 0)}%).`),
  P("**Por que importa.** A “falta de informação de reparo” é justamente o problema que o assistente pretende atacar: um em cada cinco aparelhos descartados com motivo conhecido poderia, em tese, ter sido recuperado com um procedimento de diagnóstico e reparo bem documentado. Isso reforça a justificativa social do projeto (redução de lixo eletrônico)."));

C.push(H2("11.5 Sintoma × tipo de falha (Base B)"),
  ...figura("fig09_falhas_aparelho", "Distribuição dos tipos de falha por categoria de aparelho (sem rótulos indeterminados)", 520),
  P(`**O que se observa.** O perfil de falhas muda muito conforme o aparelho. Nos **celulares**, a tela domina (${fmt(fc["Celular"]["Tela"], 0)}% dos casos), seguida de energia/bateria (${fmt(fc["Celular"]["Energia / bateria"], 0)}%) e conectores (${fmt(fc["Celular"]["Conectores e portas"], 0)}%). Nos **notebooks**, predominam software/sistema (${fmt(fc["Notebook"]["Software / sistema"], 0)}%) e desempenho/superaquecimento (${fmt(fc["Notebook"]["Desempenho / superaquecimento"], 0)}%). Nos **tablets** e **desktops**, energia e software dividem a liderança.`),
  P("**Por que importa.** O assistente deve iniciar o diagnóstico com **hipóteses diferentes para cada tipo de aparelho**. Para um celular, as primeiras perguntas devem investigar tela, bateria e conector de carga; para um notebook, sistema operacional, desempenho e aquecimento. Essa distribuição pode ser usada diretamente como probabilidade inicial (a priori) das causas no MVP."),
  ...figura("fig10_conserto_falha", "Taxa de conserto por tipo de falha", 560),
  P(`**O que se observa.** Falhas de software, armazenamento e desempenho são resolvidas com mais frequência (65% a 68%), pois costumam exigir reinstalação, limpeza ou troca de SSD. Já **placa/dano interno (${fmt(tf["Placa / dano interno"].taxa, 0)}%)**, **energia/bateria (${fmt(tf["Energia / bateria"].taxa, 0)}%)** e **inicialização (${fmt(tf["Inicialização (boot)"].taxa, 0)}%)** têm as menores taxas.`),
  P("**Por que importa.** As falhas com menor taxa de conserto são exatamente as eletrônicas, que exigem medições e testes em sequência — o foco do diagnóstico guiado da Inforcell. É nesse grupo que o assistente tende a gerar mais valor."),
  ...figura("fig11_nao_liga", "Tipos de falha atribuídos a registros com o sintoma “não liga”", 540),
  P(`**O que se observa.** Foram encontrados ${int(nl.n)} registros cujo sintoma indica que o aparelho “não liga” (expressões como “won’t turn on”, “no power”, “geht nicht an”, “ne s’allume pas”). Esses registros foram atribuídos a **${nl.n_categorias} categorias de falha diferentes**: energia/bateria (${fmt(nl.distribuicao["Energia / bateria"], 0)}%), inicialização (${fmt(nl.distribuicao["Inicialização (boot)"], 0)}%), placa/dano interno (${fmt(nl.distribuicao["Placa / dano interno"], 0)}%), tela, armazenamento, software e outras.`),
  P("**Por que importa.** Este é o resultado mais importante da análise, pois **confirma com dados reais o problema identificado na Inforcell**: um mesmo sintoma leva a várias causas possíveis. Ao mesmo tempo, os dados mostram que as causas não são igualmente prováveis — cerca de 80% se concentram em energia e inicialização. O assistente pode, portanto, ordenar os testes pela probabilidade (primeiro alimentação/bateria, depois inicialização, depois placa), reduzindo o número de testes até a causa."),
  caption("Tabela", "Termos mais característicos de cada tipo de falha (descrições em inglês)"),
  table(["Tipo de falha", "Textos", "Termos característicos (log-odds)"],
    R.termos.map((t) => [t.categoria_falha, int(t.n_textos), t.termos]), [3000, 1000, 5071], { size: 17, align: [AlignmentType.LEFT, AlignmentType.RIGHT] }),
  fonte("Elaborado pela equipe (2026). Termos com ao menos 5 ocorrências, ordenados pela razão de log-odds em relação ao vocabulário geral."),
  P("**O que se observa.** Várias categorias possuem vocabulário próprio e coerente: “touchscreen, digitiser, cracked, lcd” para tela; “hinge, casing, screws” para carcaça; “keys, keyboard, trackpad” para teclado; “overheats, noisy, hot” para superaquecimento. Em energia/bateria e inicialização, porém, os termos se sobrepõem (“turn”, “start”, “won”). **Por que importa.** Há sinal textual suficiente para classificar automaticamente parte das falhas, mas os sintomas de energia e de inicialização são descritos com as mesmas palavras — é aí que o diagnóstico guiado por testes é insubstituível."),
  H3("11.5.1 Associação entre variáveis (V de Cramér)"),
  ...figura("fig12_cramer", "Matriz de associação entre variáveis categóricas (V de Cramér, Base B)", 440),
  P(`**O que se observa.** As associações mais fortes são entre **organização e idioma** (0,56), esperada, pois cada organização atua em países específicos; entre **aparelho e tipo de falha** (0,28), que confirma a Figura 9; e entre **organização e desfecho** (0,25), maior que a associação entre tipo de falha e desfecho (0,16). O desfecho está praticamente desassociado do aparelho (0,07) e da idade (0,09).`),
  P(`**Por que importa.** A relação entre organização e desfecho indica um **viés de coleta**: o resultado do reparo depende em parte de quem registra (critérios, equipe e estrutura de cada evento). Portanto, o repair_status é um indicador útil, mas imperfeito, de “diagnóstico correto”. Além disso, ${fmt(R.pct_indeterminado)}% dos registros rotulados têm falha indeterminada (${fmt(R.pct_indeterminado_cat.Desktop)}% nos desktops), mostrando que, mesmo para voluntários experientes, identificar a causa a partir do relato é difícil.`));

C.push(H1("12 INTERPRETAÇÃO DOS PADRÕES E TENDÊNCIAS"),
  P("Reunindo os resultados, os principais padrões encontrados são:"),
  N("**Um sintoma, várias causas — mas com probabilidades desiguais.** O sintoma “não liga” se distribuiu por 11 categorias de falha, porém cerca de 80% dos casos se concentram em energia e inicialização."),
  N("**O aparelho define o ponto de partida.** Celulares têm perfil dominado por tela; notebooks, por software e desempenho. A associação aparelho × falha (V = 0,28) é a mais forte entre as variáveis do problema."),
  N("**Falhas eletrônicas são as mais difíceis.** Placa, energia e inicialização têm taxas de conserto de 33% a 39%, contra 65% a 68% em software e desempenho."),
  N("**O relato inicial é pobre.** Metade das descrições tem até seis palavras, e 26% das falhas não puderam ser classificadas nem por avaliadores humanos."),
  N("**Idade quase não explica o resultado.** Além de ausente em 70% dos registros, a idade tem correlação praticamente nula com o conserto."),
  N("**Os dados refletem o processo de coleta.** A pandemia, os reparos em lote e o viés entre organizações influenciam volume e desfecho — cuidado necessário em qualquer modelagem."),
  N("**A falta de informação de reparo é uma barreira real.** É a segunda maior causa de descarte entre os aparelhos em fim de vida com motivo informado."));

C.push(H1("13 RELAÇÃO DOS RESULTADOS COM O PROBLEMA E O MVP"),
  P("**O que os dados revelam sobre o problema que o projeto pretende solucionar?** Revelam que o problema da Inforcell não é particular: em mais de 22 mil tentativas de reparo de celulares e computadores, em 28 países, (a) o mesmo sintoma aparece associado a causas diferentes, (b) o relato inicial raramente basta para chegar à causa e (c) as falhas que mais exigem investigação eletrônica são as que menos terminam consertadas. Isso sustenta a proposta de um assistente que **conduz o técnico por testes**, e não apenas responde uma causa a partir do texto."),
  H2("13.1 Implicações para o MVP"),
  caption("Tabela", "Resultados da EDA e implicações para o MVP"),
  table(["Resultado", "Implicação para o MVP"], [
    ["Perfil de falhas muda por aparelho (Fig. 9)", "O fluxo de diagnóstico começa pela identificação do equipamento; a lista inicial de causas prováveis é específica por categoria."],
    ["“Não liga” → 11 causas, ~80% em energia e boot (Fig. 11)", "Ordenar os testes do diagnóstico guiado por probabilidade: alimentação/bateria → inicialização → placa. Primeira árvore de diagnóstico a ser implementada."],
    ["Descrições curtas (Fig. 4)", "O assistente deve fazer perguntas complementares estruturadas (quando começou, houve queda/líquido, carrega, vibra, aquece)."],
    ["Falhas eletrônicas com menor taxa de conserto (Fig. 10)", "Priorizar no MVP os roteiros de energia, boot e placa, onde o apoio ao técnico gera mais valor."],
    ["26% de falhas indeterminadas; 34% sem marca", "A OS deve ter campos obrigatórios (marca, modelo, sintoma, causa final) e registrar o resultado de cada teste — os dados da Inforcell serão mais completos que a base pública."],
    ["Taxonomia de 14 categorias", "Adotar a taxonomia como lista de “causa final” no cadastro de diagnóstico da OS, permitindo comparar os dados da Inforcell com a base pública."],
    ["Textos em outros idiomas; nenhum em português", "Usar a base pública para construir e validar as árvores de diagnóstico; para classificação textual, usar modelo multilíngue ou tradução e, progressivamente, os registros da própria Inforcell."],
  ], [3700, 5371], { size: 18 }),
  fonte("Elaborado pela equipe (2026)."),
  H2("13.2 Decisões do projeto revisadas"),
  B("**Variáveis relevantes:** passam a ser consideradas centrais a categoria do aparelho, o sintoma (texto) e o tipo de falha unificado; a idade deixa de ser variável central (alta ausência e baixa associação) e a marca passa a ser complementar."),
  B("**Perguntas que precisam ser respondidas:** além de “qual a causa provável deste sintoma?”, o projeto passa a responder “em que ordem testar as causas possíveis?”, pergunta sustentada diretamente pelas distribuições da Seção 11.5."),
  B("**Uso do conjunto de dados:** a base pública deixa de ser vista como fonte de treino direto para o técnico brasileiro e passa a ter dois papéis: (1) fonte de probabilidades iniciais e de validação das árvores de diagnóstico; (2) base de treino e avaliação de um classificador de texto na TED 03, usando o subconjunto apto para modelagem (4.019 registros)."),
  B("**Métrica do Lean Canvas:** a métrica “redução do tempo/erro no diagnóstico” será acompanhada no MVP pelo **número de testes até a causa final** e pela proporção de OS com causa indeterminada, que tem na base pública um valor de referência de 26%."),
  B("**Canvas do MVP com diagnóstico por foto de placa:** a base analisada não contém imagens; portanto, a hipótese de visão computacional registrada no canvas não pode ser validada com ela e exigiria um conjunto próprio de fotos coletado pela equipe. Para esta disciplina, o foco permanece no diagnóstico guiado a partir do sintoma relatado."),
  H2("13.3 Próximos passos (TED 03)"),
  B("Treinar e comparar modelos de classificação sintoma → categoria de falha (por exemplo, TF-IDF + regressão logística como linha de base), com divisão treino/teste temporal e estratificada por aparelho;"),
  B("Tratar o desbalanceamento entre classes (dano por líquido e carcaça com poucos exemplos);"),
  B("Transformar as distribuições da Seção 11.5 nas probabilidades iniciais da primeira árvore de diagnóstico (“não liga”)."),
  sp(), nota("Atendimento na TED 03", "Os três passos acima foram executados nas Seções 16 a 21: divisão temporal e treinamento de seis modelos (a validação cruzada foi estratificada pelo rótulo, e não pelo aparelho, conforme a Seção 21.4), tratamento do desbalanceamento por pesos de classe e comparação entre as probabilidades a priori do “não liga” e o modelo (Seção 21.2)."));

C.push(H1("14 CONSIDERAÇÕES DA ETAPA 2"),
  P(`A TED 02 transformou o Open Repair Data, inicialmente bruto, multilíngue e heterogêneo, em duas bases consistentes e documentadas: ${int(R.n_reparos)} registros de reparos de celulares e computadores e ${int(R.n_falhas)} registros com o par sintoma → tipo de falha, dos quais ${int(R.n_modelagem)} estão aptos para modelagem. A inspeção levou à revisão justificada de premissas da TED 01, e cada tratamento foi escolhido de acordo com o uso futuro do dado — em especial, a decisão de não imputar a idade, de manter os reparos em lote e de unificar as taxonomias de falha.`),
  P("A análise exploratória mostrou que os dados confirmam o problema que motivou o projeto: um mesmo sintoma, como “não liga”, está associado a diversas causas, com probabilidades desiguais, e as falhas eletrônicas são as que menos terminam consertadas. Esses resultados dão base empírica para o assistente de diagnóstico guiado da Inforcell e definem o ponto de partida da próxima etapa: a modelagem."));

// ============================ PARTE III — TED 03 (Modelagem preditiva e Machine Learning) ============================
const RM = JSON.parse(fs.readFileSync(path.join(ROOT, "reports", "resultados_modelagem.json"), "utf8"));
const p1 = (x) => fmt(x * 100, 1);
const p0 = (x) => fmt(x * 100, 0);
const f2 = (x) => fmt(x, 2);
const f3 = (x) => fmt(x, 3);
const MF = RM.final.metricas, ICF = RM.final.ic95, PCL = RM.por_classe, TM = RM.tabela_modelos;
const BASE_AP = TM.find((t) => t.modelo === "Base · prior por aparelho"), BASE_MAJ = TM.find((t) => t.modelo === "Base · classe majoritária");
const FIN = TM.find((t) => t.modelo === RM.modelo_final);
const seg = (k, g) => RM.segmentos[k].find((x) => x.grupo === g);
const sel = (c) => RM.selecao_confianca.find((x) => Math.abs(x.corte - c) < 1e-6);
const TOP = RM.topk[RM.modelo_final];
const nlg = RM.nao_liga;
const classesOrd = Object.keys(RM.dist_treino);
const nome = (m) => m.replace("Base · ", "Base: ");

C.push(H1("15 DEFINIÇÃO DO PROBLEMA DE MACHINE LEARNING"),
  H2("15.1 Retomada da TED 01 e da TED 02"),
  P(`A TED 01 definiu o problema de dados: o técnico da Inforcell precisa chegar à causa de um defeito a partir de um sintoma relatado em poucas palavras, e a experiência individual é hoje o único guia. A TED 02 preparou o conjunto **Open Repair Data** e mostrou, com dados reais, que (a) um mesmo sintoma leva a várias causas, mas com probabilidades desiguais (o sintoma “não liga” se distribuiu por 11 categorias de falha, com cerca de 80% em energia e inicialização); (b) os relatos são curtos; (c) o tipo de aparelho está associado ao tipo de falha (V de Cramér = 0,28); e (d) cerca de um quarto dos registros rotulados tem falha indeterminada. A TED 03 usa o subconjunto **falhas_modelagem.csv** (${int(RM.n_total)} registros), construído na TED 02 justamente para esta etapa. A Tabela 12 mostra como os achados da TED 02 orientaram cada decisão de modelagem.`),
  caption("Tabela", "Como os achados da TED 02 orientaram as decisões da TED 03"),
  table(["Achado da TED 02", "Decisão na TED 03"], [
    ["Relatos curtos (mediana de 5 palavras) e em vários idiomas, nenhum em português", "Representar o texto por TF-IDF de palavras e de caracteres; avaliar o desempenho por idioma; não traduzir nesta etapa (Seção 20)."],
    ["O aparelho define o ponto de partida do diagnóstico (V = 0,28)", "Usar o tipo de aparelho como segunda variável de entrada e testar seu efeito em um experimento específico."],
    ["Idade ausente em 69% dos registros e sem associação com o resultado; marca desconhecida em 34% e misturada ao aparelho", "Idade e marca ficam fora do modelo."],
    ["O desfecho do reparo é indicador imperfeito e só é conhecido depois do reparo", "Não usar status_reparo como entrada (seria vazamento de informação)."],
    ["Os dados refletem o calendário dos eventos (pandemia, mutirões de recondicionamento)", "Dividir treino e teste por data (treino até 2018; teste em 2019–2020)."],
    ["Frases repetidas poderiam ser “decoradas” pelo modelo", "Duplicatas já removidas na TED 02 (flag apto_modelagem) e textos idênticos agrupados na validação cruzada."],
    ["Classes raras (dano por líquido, carcaça)", "Pesos de classe balanceados e F1-macro como métrica de seleção."],
  ], [4000, 5071], { size: 18 }),
  fonte("Elaborado pela equipe (2026)."),
  H2("15.2 O problema de Machine Learning e sua relação com o projeto"),
  P(`O problema tratado é: **dado o sintoma relatado (texto livre) e o tipo de aparelho, qual é a categoria de falha mais provável — e quais são as k mais prováveis?** Essa pergunta corresponde à primeira etapa do assistente de diagnóstico do MVP (“levantamento das possíveis causas”) e à nova pergunta do projeto registrada na TED 02: “em que ordem testar as causas possíveis?”. O modelo não substitui o diagnóstico guiado por testes; ele fornece ao assistente uma **lista ordenada de hipóteses iniciais**, aprendida com ${int(RM.n_total)} casos reais, que o técnico confirma ou descarta com testes de bancada.`),
  H2("15.3 Tipo de abordagem adotada"),
  P(`Adotou-se **classificação supervisionada multiclasse** (13 classes presentes nos dados aptos), com saída em probabilidades para permitir um ranking de causas. A escolha decorre das características do problema:`),
  B("**Por que classificação:** o alvo é uma categoria nomeada (tipo de falha) e os dados já possuem rótulos atribuídos por voluntários; o objetivo é prever uma causa conhecida, não descobrir estruturas novas."),
  B("**Por que não regressão:** não há variável numérica de interesse a estimar. A idade do aparelho poderia ser alvo, mas é ausente em 69% dos registros e não ajuda o diagnóstico."),
  B("**Por que não agrupamento (clustering):** sem rótulos finais não haveria nome de causa para apresentar ao técnico. O agrupamento poderia complementar o trabalho no futuro (por exemplo, para descobrir subcausas dentro de “Outros”), mas não resolve o problema central."),
  B("**Outras abordagens:** modelos de linguagem multilíngues e redes neurais exigiriam bem mais dados rotulados e infraestrutura do que o disponível (2.224 textos de treino) e reduziriam a interpretabilidade; ficam como evolução (Seção 21.5)."),
  H2("15.4 Variável-alvo"),
  P(`A variável-alvo é **categoria_falha**, a taxonomia unificada em português criada na TED 02 (Tabela 4). A categoria “Indeterminado” foi excluída na TED 02, pois não informa a causa e não pode ser prevista nem usada como referência. Restam 13 classes, muito desbalanceadas: a maior (software/sistema) tem ${int(RM.classes["Software / sistema"])} registros e a menor (dano por líquido), apenas ${int(RM.classes["Dano por líquido"])}. A Tabela 13 mostra a distribuição nos conjuntos de treino e teste.`),
  caption("Tabela", "Distribuição da variável-alvo nos conjuntos de treino e teste"),
  table(["Categoria de falha", "Treino (n)", "Treino (%)", "Teste (n)", "Teste (%)"],
    classesOrd.map((c) => [c, int(RM.dist_treino[c]), fmt(RM.dist_treino[c] / RM.n_treino * 100), int(RM.dist_teste[c]), fmt(RM.dist_teste[c] / RM.n_teste * 100)])
      .concat([["**Total**", `**${int(RM.n_treino)}**`, "100,0", `**${int(RM.n_teste)}**`, "100,0"]]),
    [3400, 1300, 1450, 1300, 1621], { size: 18, align: [AlignmentType.LEFT, ...Array(4).fill(AlignmentType.RIGHT)] }),
  fonte("Elaborado pela equipe (2026), a partir do notebook 01 da TED 03."));

C.push(H1("16 PREPARAÇÃO DOS DADOS PARA A MODELAGEM"),
  H2("16.1 Seleção de variáveis"),
  P("A Tabela 14 registra quais variáveis entram no modelo e por quê. O critério foi usar apenas informações que **estarão disponíveis no momento em que o técnico relata o defeito** e que a análise exploratória mostrou serem relevantes."),
  caption("Tabela", "Variáveis do modelo e justificativas"),
  table(["Variável", "Papel", "Justificativa"], [
    ["problema (texto)", "Entrada", "É o próprio sintoma relatado; contém o sinal textual identificado na Tabela 10 (por exemplo, “battery”, “screen”, “boot”)."],
    ["categoria (aparelho)", "Entrada", "Associada ao tipo de falha (V = 0,28) e sempre conhecida na abertura da OS."],
    ["categoria_falha", "Alvo", "Causa final em 13 categorias (Seção 15.4)."],
    ["idade_anos, ano_fabricacao", "Descartada", "69% ausente e correlação com o desfecho praticamente nula (Seção 11.4.1); imputar criaria dado artificial."],
    ["marca", "Descartada", "34% desconhecida e confundida com o tipo de aparelho (Seção 11.4.2); será complementar no MVP, não no modelo."],
    ["status_reparo / consertado", "Descartada", "Só é conhecido após o reparo; usá-lo seria vazamento de informação."],
    ["idioma, fonte, país", "Descartada", "O idioma é consequência do texto; fonte e país refletem o viés de coleta (V organização × desfecho = 0,25) e não existirão na Inforcell."],
    ["data_evento", "Apenas divisão", "Usada somente para separar treino e teste no tempo."],
  ], [2300, 1400, 5371], { size: 17 }),
  fonte("Elaborado pela equipe (2026)."),
  H2("16.2 Transformações aplicadas"),
  caption("Tabela", "Transformações aplicadas para a modelagem"),
  table(["Etapa", "O que foi feito", "Por quê"], [
    ["Normalização do texto", "Minúsculas; remoção de URLs, números e pontuação; acentos preservados", "Reduz variações sem perder palavras de idiomas diferentes."],
    ["Vetorização (palavras)", "TF-IDF de unigramas e bigramas, com frequência mínima de 2 documentos e escala logarítmica", "Transforma texto em números destacando termos que distinguem as categorias."],
    ["Vetorização (caracteres)", "TF-IDF de n-gramas de 2 a 5 caracteres dentro das palavras", "Tolera erros de digitação, flexões e palavras compostas do alemão e do holandês, relevantes em uma base multilíngue."],
    ["Codificação do aparelho", "One-hot", "Variável categórica nominal, sem ordem entre as categorias."],
    ["Normalização numérica", "Não aplicada", "Não há variável numérica entre as entradas (a idade foi descartada)."],
    ["Valores ausentes", "1 registro sem categoria de aparelho tratado como “Desconhecido”", "Evita perder o registro; o texto continua informativo."],
    ["Desbalanceamento", "Pesos de classe inversamente proporcionais à frequência", "Sem pesos, o modelo ignoraria classes raras como dano por líquido (7 exemplos de treino)."],
  ], [2000, 3700, 3371], { size: 17 }),
  fonte("Elaborado pela equipe (2026)."),
  P("Todas as transformações foram implementadas dentro de um ~~pipeline~~ do scikit-learn e são ajustadas **somente com os dados de treino de cada dobra**, de modo que nenhuma informação do conjunto de teste influencie o vocabulário, os pesos TF-IDF ou os hiperparâmetros."),
  H2("16.3 Divisão entre treino e teste"),
  P(`A TED 02 (Seção 11.3) mostrou que a base reflete o funcionamento dos eventos e recomendou uma divisão que respeite o tempo. Por isso, o **treino** reúne os eventos até ${RM.ano_corte} (${int(RM.n_treino)} registros, ${p1(RM.n_treino / RM.n_total)}%) e o **teste** reúne os eventos de 2019 e 2020 (${int(RM.n_teste)} registros, ${fmt(RM.pct_teste)}%). Assim, o modelo é avaliado em casos ~~posteriores~~ aos que viu, situação semelhante ao uso real. A escolha de modelo e de hiperparâmetros foi feita apenas por validação cruzada dentro do treino (Seção 17.3); o teste serve para medir o desempenho final. Todas as classes aparecem nos dois conjuntos.`),
  ...figura("fig13_divisao_dados", "Divisão temporal em treino e teste e distribuição das classes", 600),
  P(`**O que se observa.** O teste é grande (${fmt(RM.pct_teste, 0)}% dos dados) porque 2019 sozinho concentra ${int(1520)} registros. A distribuição das classes mudou entre os períodos: por exemplo, “tela” passou de ${fmt(RM.dist_treino["Tela"] / RM.n_treino * 100, 1)}% do treino para ${fmt(RM.dist_teste["Tela"] / RM.n_teste * 100, 1)}% do teste, e “software/sistema” de ${fmt(RM.dist_treino["Software / sistema"] / RM.n_treino * 100, 1)}% para ${fmt(RM.dist_teste["Software / sistema"] / RM.n_teste * 100, 1)}%. **Por que importa.** Há uma deriva moderada entre os períodos; o desempenho medido já inclui esse efeito (a Seção 19.3 compara com uma divisão aleatória).`),
  H2("16.4 Verificações de qualidade antes de modelar"),
  B(`**Rótulos conflitantes.** Dos ${int(RM.n_total)} textos, ${int(RM.textos_repetidos.textos)} aparecem mais de uma vez (${int(RM.textos_repetidos.registros)} registros) e ${int(RM.textos_repetidos.conflitantes)} deles receberam **rótulos diferentes** (${int(RM.textos_repetidos.registros_conflitantes)} registros). Textos curtos como “won’t turn on” são classificados de formas distintas por avaliadores distintos; isso impõe um **teto de desempenho** aos dados.`),
  B(`**Memorização.** Apenas ${fmt(RM.pct_teste_texto_visto)}% dos textos do teste aparecem idênticos no treino, de modo que o teste mede principalmente a capacidade de generalizar a textos novos.`),
  B(`**Tamanho dos textos.** A mediana é de ${fmt(RM.palavras_mediana, 0)} palavras e ${fmt(RM.pct_ate_3_palavras)}% dos relatos têm até 3 palavras, o que limita a informação disponível a qualquer modelo.`));

C.push(H1("17 TÉCNICAS, JUSTIFICATIVA E PROCESSO DE TREINAMENTO"),
  H2("17.1 Técnicas utilizadas"),
  P("Foram comparadas **duas linhas de base** e **seis modelos candidatos**, todos lineares sobre TF-IDF. As linhas de base servem de régua: sem elas, uma acurácia de 70% não teria significado."),
  caption("Tabela", "Linhas de base e modelos candidatos"),
  table(["Modelo", "Representação do texto", "Algoritmo", "Hiperparâmetros testados"], [
    ["Base: classe majoritária", "—", "Prevê sempre a classe mais frequente do treino", "—"],
    ["Base: prior por aparelho", "Apenas o tipo de aparelho", "Distribuição de falhas por aparelho (a priori da TED 02)", "—"],
    ["LR · palavras", "TF-IDF de palavras (1–2 gramas)", "Regressão logística multinomial", "C ∈ {1, 3, 10, 30}"],
    ["LR · caracteres", "TF-IDF de caracteres (2–5)", "Regressão logística multinomial", "C ∈ {1, 3, 10, 30}"],
    ["LR · palavras+caracteres", "Palavras e caracteres combinados", "Regressão logística multinomial", "C ∈ {1, 3, 10, 30}"],
    ["LR · palavras+caracteres+aparelho", "Palavras, caracteres e aparelho (one-hot)", "Regressão logística multinomial", "C ∈ {1, 3, 10, 30}"],
    ["SVM linear · palavras+caracteres", "Palavras e caracteres combinados", "Máquina de vetores de suporte linear", "C ∈ {0,03; 0,1; 0,3; 1}"],
    ["Naive Bayes (Complement) · palavras+caracteres", "Palavras e caracteres combinados", "Naive Bayes complementar", "alfa ∈ {0,03; 0,1; 0,3; 1}"],
  ], [2700, 2200, 2400, 1771], { size: 16 }),
  fonte("Elaborado pela equipe (2026). C: inverso da força de regularização; LR: regressão logística."),
  H2("17.2 Justificativa da escolha"),
  B("**Poucos dados e muitas variáveis esparsas.** Com 2.224 textos curtos de treino e milhares de termos, modelos lineares com regularização são a escolha mais robusta; modelos profundos tenderiam a sobreajustar."),
  B("**TF-IDF.** Dá mais peso a termos que distinguem uma categoria (“battery”, “hinge”) do que a termos comuns a todas, e é a linha de base recomendada na TED 02."),
  B("**Caracteres além de palavras.** A base mistura inglês, holandês, alemão e francês, com erros de digitação e palavras compostas (“Ladebuchse”, “waterschade”). N-gramas de caracteres compartilham sinal entre variações da mesma palavra."),
  B("**Regressão logística, SVM e Naive Bayes.** Representam três famílias clássicas de classificadores de texto (discriminativo probabilístico, margem máxima e generativo), permitindo checar se a conclusão depende do algoritmo. A regressão logística e o Naive Bayes fornecem **probabilidades**; o SVM fornece apenas margens."),
  B("**Aparelho como entrada.** O experimento com e sem o aparelho mede diretamente se a associação encontrada na EDA (V = 0,28) ajuda a prever."),
  H2("17.3 Treinamento, validação e avaliação"),
  B(`**Divisão:** treino até ${RM.ano_corte} e teste em 2019–2020 (Seção 16.3).`),
  B("**Validação cruzada:** 5 dobras dentro do treino, estratificadas pelo rótulo e **agrupadas por texto idêntico**, para que um mesmo texto não apareça ao mesmo tempo no treino e na validação de uma dobra."),
  B("**Seleção de hiperparâmetros:** busca em grade (Tabela 16) com F1-macro como critério; o melhor conjunto é reajustado em todo o treino."),
  B("**Avaliação final:** uma única avaliação no teste, com intervalos de confiança de 95% por ~~bootstrap~~ (1.000 reamostragens do teste)."),
  B("**Reprodutibilidade:** sementes aleatórias fixas (42); a segunda execução dos notebooks reproduziu todos os números deste documento."),
  H2("17.4 Métricas de avaliação"),
  P("A escolha das métricas considera dois fatos: as classes são muito desbalanceadas e o assistente precisa de uma **lista ordenada** de causas, não apenas de uma resposta."),
  caption("Tabela", "Métricas utilizadas e justificativas"),
  table(["Métrica", "O que mede", "Por que foi usada"], [
    ["F1-macro", "Média do F1 de cada classe, com o mesmo peso para todas", "Métrica principal: a acurácia esconderia o desempenho ruim em classes raras (por exemplo, com 14,5% um modelo que sempre prevê “software” já parece razoável)."],
    ["Acurácia", "% de previsões corretas", "Fácil de comunicar; usada junto com o F1-macro."],
    ["Precisão e revocação por classe", "Entre as previstas como X, quantas são X; entre as X reais, quantas foram encontradas", "Mostram onde o modelo erra por excesso ou por omissão."],
    ["Matriz de confusão", "Quais classes são confundidas entre si", "Permite relacionar os erros a categorias tecnicamente próximas."],
    ["Acerto top-k", "% de casos em que a causa correta está entre as k mais prováveis", "O assistente sugere uma lista ordenada de causas; é a métrica mais próxima do uso real."],
    ["Log-loss", "Qualidade das probabilidades previstas", "O MVP usará as probabilidades para ordenar testes e decidir quando pedir mais informações."],
    ["Calibração (ECE)", "Diferença entre a confiança prevista e o acerto observado", "Indica se a probabilidade pode ser lida como “chance de acerto” (GUO et al., 2017)."],
    ["IC 95% por bootstrap", "Incerteza da métrica por reamostragem do teste", "Evita comparar números que diferem apenas por acaso."],
  ], [2000, 3200, 3871], { size: 17 }),
  fonte("Elaborado pela equipe (2026), com base em Sokolova e Lapalme (2009)."));

C.push(H1("18 RESULTADOS OBTIDOS"),
  H2("18.1 Linhas de base e comparação dos modelos"),
  P(`A Tabela 18 apresenta, para cada modelo, o F1-macro médio da validação cruzada (± desvio-padrão entre as dobras) e as métricas no teste. A Figura 14 resume a comparação.`),
  caption("Tabela", "Comparação das linhas de base e dos modelos candidatos"),
  table(["Modelo", "F1-macro (validação cruzada)", "Acurácia (teste)", "F1-macro (teste)", "Top-3 (teste)"],
    TM.map((t) => [t.modelo === RM.modelo_final ? `**${nome(t.modelo)}** ★` : nome(t.modelo),
      t.cv_f1_macro == null ? "—" : `${f3(t.cv_f1_macro)} ± ${f3(t.cv_dp)}`, `${p1(t.acuracia)}%`, f3(t.f1_macro), `${p1(t.top3)}%`]),
    [3500, 1800, 1250, 1250, 1271], { size: 16, align: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  fonte("Elaborado pela equipe (2026), a partir do notebook 02 da TED 03. ★ modelo escolhido."),
  ...figura("fig14_comparacao_modelos", "Comparação das linhas de base e dos modelos candidatos", 560),
  P(`**O que se observa.** (1) Todos os modelos que usam o texto superam com folga as linhas de base: o prior por aparelho, que é a ideia da TED 02, obtém apenas ${p1(BASE_AP.acuracia)}% de acurácia e F1-macro de ${f3(BASE_AP.f1_macro)}, e a classe majoritária, ${p1(BASE_MAJ.acuracia)}%. (2) Entre os modelos com texto, as diferenças são pequenas: na validação cruzada, os quatro melhores ficam entre ${f3(0.583)} e ${f3(0.600)}, dentro de um desvio-padrão uns dos outros. (3) Usar **apenas palavras é pior** que usar caracteres (F1-macro de ${f3(TM.find((t) => t.modelo === "LR · palavras").f1_macro)} contra ${f3(TM.find((t) => t.modelo === "LR · caracteres").f1_macro)} no teste), o que sustenta a escolha de n-gramas de caracteres em uma base multilíngue. (4) O Naive Bayes foi o de menor desempenho entre os que usam caracteres. **Por que importa.** O ganho vem sobretudo de **usar o texto**; a escolha entre regressão logística e SVM é secundária.`),
  H2("18.2 Escolha do modelo"),
  P(`Pelo F1-macro de validação cruzada, o melhor candidato foi o **${nome(RM.melhor_cv)}** (${f3(TM.find((t) => t.modelo === RM.melhor_cv).cv_f1_macro)} ± ${f3(RM.criterio_empate_dp)}). Como os candidatos ficaram empatados dentro de um desvio-padrão, a decisão usou critérios derivados dos requisitos do MVP e da EDA: (i) saída em **probabilidades**, indispensável para ordenar causas e medir confiança (o SVM só fornece margens); e (ii) uso de **todas as informações disponíveis na abertura da OS** (sintoma e aparelho), já que a TED 02 associou o aparelho ao tipo de falha. Foi escolhido o modelo **${RM.modelo_final}** (C = ${RM.final.hiperparametros.clf__C}), que obteve F1-macro de ${f3(FIN.cv_f1_macro)} na validação cruzada, a ${f3(TM.find((t) => t.modelo === RM.melhor_cv).cv_f1_macro - FIN.cv_f1_macro)} do melhor, bem dentro de um desvio-padrão (${f3(RM.criterio_empate_dp)}). Como o intervalo de confiança do F1-macro no teste tem largura de cerca de ±0,03, **não se afirma que este modelo seja superior aos demais candidatos**; afirma-se que atende aos critérios do projeto com desempenho equivalente ao dos melhores.`),
  H2("18.3 Desempenho do modelo escolhido no teste"),
  caption("Tabela", "Desempenho do modelo escolhido no conjunto de teste (2019–2020, n = " + int(RM.n_teste) + ")"),
  table(["Métrica", "Valor", "IC 95%", "Linha de base (prior por aparelho)"], [
    ["Acurácia", `${p1(MF.acuracia)}%`, `${p1(ICF.acuracia[0])}% a ${p1(ICF.acuracia[1])}%`, `${p1(BASE_AP.acuracia)}%`],
    ["F1-macro", f3(MF.f1_macro), `${f3(ICF.f1_macro[0])} a ${f3(ICF.f1_macro[1])}`, f3(BASE_AP.f1_macro)],
    ["Precisão macro / revocação macro", `${f3(MF.precisao_macro)} / ${f3(MF.revocacao_macro)}`, "—", `${f3(BASE_AP.precisao_macro)} / ${f3(BASE_AP.revocacao_macro)}`],
    ["Acerto top-2", `${p1(MF.top2)}%`, "—", `${p1(BASE_AP.top2)}%`],
    ["Acerto top-3", `${p1(MF.top3)}%`, `${p1(ICF.top3[0])}% a ${p1(ICF.top3[1])}%`, `${p1(BASE_AP.top3)}%`],
    ["Acerto top-5", `${p1(MF.top5)}%`, "—", `${p1(BASE_AP.top5)}%`],
    ["Log-loss (menor é melhor)", f2(MF.log_loss), "—", f2(RM.final.log_loss_base)],
  ], [2600, 1700, 2300, 2471], { size: 17, align: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] }),
  fonte("Elaborado pela equipe (2026), a partir do notebook 02 da TED 03."),
  P(`O ganho de F1-macro sobre o prior por aparelho é de ${f2(ICF.dif_f1_macro_vs_base[0])} a ${f2(ICF.dif_f1_macro_vs_base[1])} (IC 95% da diferença pareada), isto é, **muito além do acaso**.`),
  H2("18.4 Desempenho por categoria de falha"),
  ...figura("fig15_f1_por_classe", "F1 por categoria de falha: modelo escolhido × prior por aparelho", 520),
  caption("Tabela", "Precisão, revocação e F1 por categoria de falha (teste)"),
  table(["Categoria de falha", "n treino", "n teste", "Precisão", "Revocação", "F1", "F1 da base"],
    Object.entries(PCL).sort((a, b) => b[1].f1 - a[1].f1).map(([c, v]) => [c, int(v.n_treino), int(v.n), f2(v.precisao), f2(v.revocacao), `**${f2(v.f1)}**`, f2(v.f1_base)]),
    [2850, 900, 900, 1000, 1100, 700, 1621], { size: 16, align: [AlignmentType.LEFT, ...Array(6).fill(AlignmentType.RIGHT)] }),
  fonte("Elaborado pela equipe (2026), a partir do notebook 02 da TED 03."),
  P(`**O que se observa.** O desempenho varia muito entre as categorias. As melhores são **tela** (F1 ${f2(PCL["Tela"].f1)}), **energia/bateria** (${f2(PCL["Energia / bateria"].f1)}), **conectores** (${f2(PCL["Conectores e portas"].f1)}), **áudio/câmera/mídia** (${f2(PCL["Áudio / câmera / mídia"].f1)}), **software** (${f2(PCL["Software / sistema"].f1)}) e **botões/teclado** (${f2(PCL["Botões / teclado / touchpad"].f1)}), categorias com vocabulário próprio (Tabela 10). As piores são **outros** (${f2(PCL["Outros"].f1)}), **placa/dano interno** (${f2(PCL["Placa / dano interno"].f1)}), **carcaça** (${f2(PCL["Carcaça / dano externo"].f1)}), **dano por líquido** (${f2(PCL["Dano por líquido"].f1)}) e **inicialização** (${f2(PCL["Inicialização (boot)"].f1)}). **Por que importa.** Duas das categorias mais difíceis (placa e inicialização) são justamente as falhas eletrônicas que a TED 02 identificou como as de menor taxa de conserto e como o foco do diagnóstico guiado; nelas o texto sozinho sustenta menos o diagnóstico, e a condução por testes continua indispensável. A classe “outros” é, por definição, heterogênea, e “dano por líquido” tem só ${int(PCL["Dano por líquido"].n)} exemplos no teste, portanto seu F1 é pouco confiável.`),
  H2("18.5 Matriz de confusão"),
  ...figura("fig16_matriz_confusao", "Matriz de confusão do modelo escolhido (cada linha soma 100%)", 540),
  caption("Tabela", "Principais confusões entre categorias (teste)"),
  table(["Categoria verdadeira", "Prevista como", "Casos", "% da categoria verdadeira"],
    RM.confusoes.map((c) => [c.verdadeiro, c.previsto, int(c.n), `${p1(c.pct_da_classe)}%`]),
    [2900, 3000, 900, 2271], { size: 17, align: [AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT] }),
  fonte("Elaborado pela equipe (2026), a partir do notebook 02 da TED 03."),
  P(`**O que se observa.** Os erros concentram-se em pares de categorias tecnicamente próximas: software ↔ desempenho/superaquecimento (um relato de “lento” pode ser sistema ou hardware), **energia/bateria → inicialização** (${int(RM.confusoes[1].n)} casos, ${p1(RM.confusoes[1].pct_da_classe)}% de energia), **placa → energia/bateria** e “outros” → software e desempenho. **Por que importa.** A confusão energia × inicialização reproduz a sobreposição de vocabulário já observada na EDA (“turn”, “start”, “won”): não é falha do algoritmo, e sim ambiguidade do próprio relato, que só os testes de bancada resolvem.`),
  H2("18.6 Acerto top-k"),
  ...figura("fig17_topk", "Acerto top-k do modelo escolhido e das linhas de base", 480),
  P(`**O que se observa.** A causa correta é a primeira sugestão em ${p1(TOP["1"])}% dos casos, está entre as duas primeiras em ${p1(TOP["2"])}% e entre as três primeiras em **${p1(TOP["3"])}%**; com cinco sugestões, chega a ${p1(TOP["5"])}%. As linhas de base alcançam apenas ${p1(RM.topk["Base · prior por aparelho"]["3"])}% no top-3. **Por que importa.** Para o assistente, que propõe testes em sequência, **apresentar as três causas mais prováveis já cobre cerca de nove em cada dez casos**, resultado muito mais útil do que exigir uma única resposta correta.`));

C.push(H1("19 INTERPRETAÇÃO DOS RESULTADOS"),
  H2("19.1 O que os números significam"),
  P(`Um F1-macro de ${f2(MF.f1_macro)} e uma acurácia de ${fmt(MF.acuracia * 100, 0)}% não são resultados de “alto desempenho” em termos absolutos, mas devem ser lidos junto com três fatos. **Primeiro**, o modelo faz o que as linhas de base não fazem: usa o relato para discriminar entre 13 causas, enquanto a base que apenas conhece o aparelho acerta ${fmt(BASE_AP.acuracia * 100, 0)}%. **Segundo**, há um **teto imposto pelos dados**: ${int(RM.textos_repetidos.conflitantes)} textos idênticos receberam rótulos diferentes e 26% dos casos da base original eram indeterminados até para voluntários experientes; nenhum modelo pode ser mais consistente do que os rótulos que aprende. **Terceiro**, o objetivo prático (sugerir causas a serem testadas) é bem atendido pelo top-3 de ${p0(TOP["3"])}%.`),
  P(`O F1-macro de teste (${f3(MF.f1_macro)}) é maior que o da validação cruzada (${f3(FIN.cv_f1_macro)}). Isso não indica vazamento: nas dobras, o modelo treina com 80% de um treino que já é pequeno, e classes raras (7 exemplos de dano por líquido) têm F1 instável por dobra; a curva de aprendizagem (Seção 19.3) mostra que o desempenho cresce com o número de exemplos, o que explica a diferença. Também por isso o intervalo de confiança do teste (${f3(ICF.f1_macro[0])} a ${f3(ICF.f1_macro[1])}) deve ser lido como a faixa plausível do desempenho, não como um número exato.`),
  H2("19.2 Onde o modelo acerta e onde erra"),
  P(`O modelo é confiável quando o relato contém **termos técnicos específicos** (tela, bateria, teclado, conector, sistema operacional) e frágil quando a causa depende de **medições** (placa, inicialização) ou quando a categoria é residual (“outros”). A Tabela 22 mostra erros de alta confiança que ilustram um ponto central: muitos “erros” são, na verdade, **rótulos discutíveis** (por exemplo, “Not booting” rotulado como energia/bateria e previsto como inicialização; “Touchpad and keyboard faulty” rotulado como “outros” e previsto como teclado/touchpad). O modelo não está “errado” nesses casos; a taxonomia e a rotulagem é que admitem mais de uma resposta.`),
  caption("Tabela", "Exemplos de erros de alta confiança (relatos curtos em inglês, teste)"),
  table(["Relato", "Rótulo verdadeiro", "Previsto", "Confiança"],
    RM.exemplos_erros.slice(0, 6).map((e) => [`“${e.texto}”`, e.verdadeiro, e.previsto, `${p0(e.confianca)}%`]),
    [3200, 2300, 2300, 1271], { size: 16, align: [AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.RIGHT] }),
  fonte("Elaborado pela equipe (2026), a partir do notebook 03 da TED 03. Selecionados automaticamente: os erros de maior confiança com até 12 palavras em inglês."),
  H2("19.3 Robustez e generalização"),
  ...figura("fig18_desempenho_segmentos", "Desempenho (top-1 e top-3) por idioma, aparelho e tamanho do relato", 600),
  P(`**Por segmento.** O desempenho é maior em **inglês** (${p0(seg("idioma", "en").acuracia)}% de acerto top-1) do que em **alemão** (${p0(seg("idioma", "de").acuracia)}%), com holandês em ${p0(seg("idioma", "nl").acuracia)}%; os idiomas menos frequentes (francês, italiano, norueguês, espanhol) somam apenas ${int(seg("idioma", "outros (fr, it, no, es)").n)} registros e ficam em ${p0(seg("idioma", "outros (fr, it, no, es)").acuracia)}%. Entre os aparelhos, **tablet** (${p0(seg("aparelho", "Tablet").acuracia)}%) e celular (${p0(seg("aparelho", "Celular").acuracia)}%) vão melhor que **desktop** (${p0(seg("aparelho", "Desktop").acuracia)}%), categoria com menos dados e mais rótulos indeterminados na TED 02. Surpreendentemente, relatos **curtos** (1–3 palavras) têm acerto maior (${p0(seg("tamanho", "1–3 palavras").acuracia)}%) que relatos **longos** (7 ou mais palavras: ${p0(seg("tamanho", "7+ palavras").acuracia)}%). A explicação mais plausível, **não testada**, é que relatos curtos usam termos padronizados (“screen”, “battery”) e relatos longos descrevem mais de um problema ou a história do reparo, o que torna o rótulo único mais ambíguo.`),
  caption("Tabela", "Divisão temporal × divisão aleatória (mesmo modelo e mesmo tamanho de teste)"),
  table(["Métrica", "Divisão temporal (usada)", "Divisão aleatória (média de 3 sementes)"],
    [["Acurácia", `${p1(RM.temporal_vs_aleatorio.acuracia.temporal)}%`, `${p1(RM.temporal_vs_aleatorio.acuracia.aleatorio)}% (± ${p1(RM.temporal_vs_aleatorio.acuracia.aleatorio_dp)})`],
     ["F1-macro", f3(RM.temporal_vs_aleatorio.f1_macro.temporal), `${f3(RM.temporal_vs_aleatorio.f1_macro.aleatorio)} (± ${f3(RM.temporal_vs_aleatorio.f1_macro.aleatorio_dp)})`],
     ["Top-3", `${p1(RM.temporal_vs_aleatorio.top3.temporal)}%`, `${p1(RM.temporal_vs_aleatorio.top3.aleatorio)}% (± ${p1(RM.temporal_vs_aleatorio.top3.aleatorio_dp)})`]],
    [2400, 3300, 3371], { size: 17, align: [AlignmentType.LEFT, AlignmentType.CENTER, AlignmentType.CENTER] }),
  fonte("Elaborado pela equipe (2026), a partir do notebook 03 da TED 03."),
  P(`**Temporal × aleatória.** A divisão aleatória rende resultados apenas um pouco melhores (acurácia ${p1(RM.temporal_vs_aleatorio.acuracia.aleatorio)}% contra ${p1(RM.temporal_vs_aleatorio.acuracia.temporal)}%), isto é, a deriva entre períodos existe, mas é pequena, e o resultado temporal não está inflado por misturar eventos do mesmo período. **Memorização.** O acerto em textos já vistos no treino é de ${p1(RM.visto_vs_inedito.visto.acuracia)}% (n = ${int(RM.visto_vs_inedito.visto.n)}), contra ${p1(RM.visto_vs_inedito.inedito.acuracia)}% nos textos inéditos (n = ${int(RM.visto_vs_inedito.inedito.n)}); como os textos inéditos são 95% do teste, o desempenho reflete generalização.`),
  P(`**Sobreajuste.** Há sobreajuste claro: o modelo reavaliado sobre os próprios dados de treino acerta 96,7% (F1-macro 0,969), contra 71,6% (F1-macro 0,654) no teste, uma diferença de cerca de 25 pontos de acurácia. É esperado para um modelo com milhares de atributos de texto (palavras e caracteres) e poucos exemplos por classe; a regularização escolhida na validação cruzada (C = 3) limita o problema, mas não o elimina. Por isso a nota do modelo é a do teste, e nunca a do treino, e a curva de aprendizagem acima indica que mais dados rotulados são o caminho mais seguro para reduzir essa diferença.`),
  ...figura("fig19_curva_aprendizagem", "Curva de aprendizagem: desempenho no teste conforme o tamanho do treino", 440),
  P(`**Curva de aprendizagem.** O F1-macro sobe de ${f2(RM.curva_aprendizagem[0].f1)} (com ${int(RM.curva_aprendizagem[0].n)} exemplos) para ${f2(RM.curva_aprendizagem[5].f1)} (com ${int(RM.curva_aprendizagem[5].n)}) e **ainda não estabilizou**: de 80% para 100% do treino houve ganho de ${f2(RM.curva_aprendizagem[5].f1 - RM.curva_aprendizagem[4].f1)}. Mais dados rotulados tendem a melhorar o modelo, o que reforça o valor de acumular as Ordens de Serviço reais da Inforcell.`),
  H2("19.4 Confiança e calibração"),
  ...figura("fig20_confianca", "Calibração das probabilidades e relação entre confiança, acerto e cobertura", 590),
  caption("Tabela", "Acerto e cobertura conforme o limiar de confiança (probabilidade da 1ª causa)"),
  table(["Limiar de confiança", "Casos em que o assistente sugere (cobertura)", "Acerto top-1 nesses casos"],
    [0, 0.25, 0.35, 0.5, 0.6, 0.8].map((c) => [c === 0 ? "sem limiar" : `≥ ${fmt(c, 2)}`, `${p1(sel(c).cobertura)}%`, `${p1(sel(c).acuracia)}%`]),
    [2700, 3600, 2771], { size: 17, align: [AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] }),
  fonte("Elaborado pela equipe (2026), a partir do notebook 03 da TED 03."),
  P(`**O que se observa.** O acerto aumenta de forma consistente com a confiança: nos casos em que a primeira causa tem probabilidade de pelo menos 0,50, o acerto top-1 é de ${p1(sel(0.5).acuracia)}%, cobrindo ${p0(sel(0.5).cobertura)}% dos casos. As probabilidades, porém, são **subestimadas** (erro de calibração esperado, ECE = ${f3(RM.calibracao.ece)}): quando o modelo diz 0,54, acerta cerca de 73% das vezes. **Por que importa.** As probabilidades servem bem para **ordenar** as causas e para **decidir quando pedir mais informações**, mas não devem ser exibidas ao técnico como “chance de acerto”. Com o limiar de 0,35, o assistente responderia em ${p0(sel(0.35).cobertura)}% dos casos com ${p0(sel(0.35).acuracia)}% de acerto top-1; abaixo dele, deveria fazer perguntas complementares (uma calibração posterior, como a isotônica, é um próximo passo).`),
  H2("19.5 O que o modelo aprendeu"),
  caption("Tabela", "Termos de maior peso por categoria de falha (bloco de palavras da regressão logística)"),
  table(["Categoria de falha", "Termos de maior peso"], Object.entries(RM.termos_modelo).map(([c, t]) => [c, t.join(", ")]), [3000, 6071], { size: 17 }),
  fonte("Elaborado pela equipe (2026), a partir do notebook 03 da TED 03."),
  P(`**O que se observa.** Em grande parte, os termos são tecnicamente plausíveis e multilíngues: “battery, power, charger, accu, akku” para energia; “screen, display, lcd, scherm, écran” para tela; “boot, bios” para inicialização; “virus, software, linux, password” para software; “waterschade, wasserschaden” para dano por líquido. Isso confirma que o modelo aprendeu sinais técnicos e que os n-gramas de caracteres, somados às palavras, transferem conhecimento entre idiomas. **Sinais de alerta:** em classes pequenas aparecem termos sem valor técnico (“das” em placa, “ed” em outros, “red” em conectores), indício de sobreajuste a poucos exemplos.`));

C.push(H1("20 LIMITAÇÕES IDENTIFICADAS"),
  P("As principais restrições do modelo, possíveis fontes de erro e aspectos que afetam sua generalização são:"),
  B(`**Idioma: nenhum texto em português.** O modelo foi treinado com inglês (${p0(RM.segmentos.idioma.find((x) => x.grupo === "en").n / RM.n_teste)}% do teste), holandês, alemão e francês. Em uma sondagem qualitativa com 8 frases em português, a primeira causa estava correta em apenas **1 de 8** frases e a causa esperada estava entre as três primeiras em 2 de 8; em compensação, o limiar de confiança marcou 7 dessas 8 respostas como “baixa confiança”, o que mostra que o filtro protege o técnico, mas que **o modelo não pode ser usado em português sem tradução ou dados novos**. É uma sondagem, não uma métrica: são poucas frases e a resposta esperada foi definida pela própria equipe.`),
  B("**Contexto dos dados.** Os registros vêm de eventos comunitários de reparo na Europa, com voluntários e outro perfil de aparelhos e defeitos que uma assistência técnica comercial em Balsas/MA."),
  B(`**Qualidade e ambiguidade dos rótulos.** ${int(RM.textos_repetidos.conflitantes)} textos idênticos têm rótulos conflitantes; a taxonomia de 14 categorias foi criada pela equipe a partir de duas listas diferentes (26 e 20 tipos). Parte do erro do modelo é erro (ou discordância) de rotulagem.`),
  B(`**Classes raras e residuais.** “Dano por líquido” tem ${int(PCL["Dano por líquido"].n_treino)} exemplos de treino e ${int(PCL["Dano por líquido"].n)} de teste; “outros” é heterogênea. Os resultados dessas classes são pouco confiáveis.`),
  B("**Granularidade.** O modelo prevê a categoria geral (por exemplo, “energia/bateria”), não o componente (conector de carga, CI de carga, bateria). O diagnóstico fino continua dependendo do roteiro de testes."),
  B("**Probabilidades mal calibradas.** Subestimam a chance de acerto (ECE = " + f3(RM.calibracao.ece) + "); servem para ordenar, não para informar porcentagens ao técnico."),
  B("**Deriva temporal e dados limitados.** Os dados vão de 2012 a 2020, enquanto os aparelhos atuais diferem; a curva de aprendizagem ainda sobe, ou seja, o modelo está limitado pelo volume de dados."),
  B("**Uso do teste na comparação.** O conjunto de teste não foi usado para ajustar hiperparâmetros nem para escolher o modelo, mas as métricas de teste de todos os candidatos foram examinadas durante o trabalho; comparações entre candidatos devem ser vistas como indicativas. A estimativa do desempenho absoluto do modelo escolhido não depende dessa comparação."),
  B("**Imagens.** A base não contém imagens; a hipótese de diagnóstico por foto de placa, registrada no canvas, continua sem validação com estes dados (Seção 13.2)."));

C.push(H1("21 RELAÇÃO COM O PROJETO DE INTELIGÊNCIA ARTIFICIAL E COM O MVP"),
  H2("21.1 Respostas às perguntas orientadoras"),
  caption("Tabela", "Perguntas orientadoras da TED 03 e respostas da equipe"),
  table(["Pergunta", "Resposta"], [
    ["O modelo contribui para resolver o problema do projeto?", `Sim, parcialmente. Ele entrega ao assistente uma lista ordenada de causas prováveis com ${p0(TOP["3"])}% de cobertura no top-3, contra ${p0(RM.topk["Base · prior por aparelho"]["3"])}% da linha de base. Não substitui os testes de bancada, necessários para separar causas com vocabulário sobreposto.`],
    ["Os resultados são adequados ao objetivo?", "Adequados como **ponto de partida** para sugerir hipóteses; insuficientes como diagnóstico final (acerto top-1 de " + p0(MF.acuracia) + "%) e ainda não validados para relatos em português."],
    ["Quais variáveis parecem relevantes?", "O **texto do sintoma** é a variável decisiva. O aparelho ajuda: o F1-macro de teste sobe de " + f3(TM.find((t) => t.modelo === "LR · palavras+caracteres").f1_macro) + " para " + f3(FIN.f1_macro) + " ao incluí-lo, embora a diferença na validação cruzada seja pequena (" + f3(TM.find((t) => t.modelo === "LR · palavras+caracteres").cv_f1_macro) + " contra " + f3(FIN.cv_f1_macro) + "). Idade, marca e desfecho ficaram de fora por falta de qualidade ou por vazamento."],
    ["Quais as principais limitações?", "Idioma, contexto comunitário, rótulos ambíguos, classes raras, granularidade e calibração (Seção 20)."],
    ["Em quais condições o modelo pode ser usado?", `Como **ordenador de hipóteses** no início do diagnóstico, para relatos em inglês, holandês, alemão ou francês, ou em português **depois de tradução e de validação**. Respostas com confiança abaixo de 0,35 devem acionar perguntas complementares em vez de sugerir causa.`],
    ["Como os resultados contribuem para o MVP?", "Fornecem a primeira camada do assistente (causas prováveis ordenadas), um critério objetivo para decidir quando pedir mais informações e uma meta de comparação para os dados reais da Inforcell (Seção 21.3)."],
  ], [3100, 5971], { size: 17 }),
  fonte("Elaborado pela equipe (2026)."),
  H2("21.2 O caso “não liga” revisitado"),
  P(`Na TED 02 propôs-se ordenar os testes do sintoma “não liga” pelas probabilidades a priori das causas (energia/bateria → inicialização → placa). Medimos se o modelo, que enxerga também o texto e o aparelho, faz melhor do que essa ordem fixa, usando os ${int(nlg.n_teste)} registros de teste com esse sintoma e a ordem calculada **apenas no treino** (${nlg.ordem_fixa_treino.slice(0, 3).join(" → ")}).`),
  ...figura("fig21_nao_liga", "Acerto da ordem fixa (a priori) e do modelo para o sintoma “não liga”", 440),
  P(`**O que se observa.** A causa correta é a primeira sugestão em ${p0(nlg.acerto["1"].modelo)}% dos casos com o modelo e em ${p0(nlg.acerto["1"].ordem_fixa)}% com a ordem fixa; até a segunda, ${p0(nlg.acerto["2"].modelo)}% e ${p0(nlg.acerto["2"].ordem_fixa)}%; até a terceira, ${p0(nlg.acerto["3"].modelo)}% e ${p0(nlg.acerto["3"].ordem_fixa)}%. **Por que importa.** O modelo melhora a ordenação (principalmente na primeira sugestão), mas a ordem fixa da TED 02 já é um bom ponto de partida (cobre ${p0(nlg.acerto["3"].ordem_fixa)}% no top-3). A recomendação é **manter a árvore de diagnóstico baseada nas probabilidades da TED 02 como roteiro básico** e usar o modelo para reordenar as causas quando o técnico fornecer mais detalhes no relato. Como o subconjunto é pequeno (${int(nlg.n_teste)} casos), a diferença deve ser confirmada com dados reais.`),
  H2("21.3 Como o modelo entra no MVP"),
  B("**Entrada:** texto do sintoma e tipo de aparelho cadastrados na OS. **Saída:** três causas mais prováveis com suas probabilidades e um indicador de confiança (src/inferencia.py, função sugerir_causas)."),
  B("**Regra de uso:** confiança ≥ 0,35 → o assistente inicia pelos testes da primeira causa e segue a ordem sugerida; confiança < 0,35 → faz primeiro as perguntas complementares estruturadas definidas na TED 02 (quando começou, houve queda ou líquido, carrega, aquece)."),
  B("**Papel junto ao assistente conversacional:** o modelo fornece as hipóteses iniciais e o assistente de IA do MVP conduz o diálogo e interpreta o resultado de cada teste; o técnico mantém a decisão final."),
  B("**Aprendizado contínuo:** o campo “causa final” da OS, com a mesma taxonomia de 14 categorias, passa a gerar dados rotulados em português, que permitem reajustar o modelo e medir a métrica “redução de tempo/erro no diagnóstico” definida no Lean Canvas."),
  H2("21.4 Decisões revisadas nesta etapa"),
  B("**Estratificação:** a TED 02 previa divisão estratificada por aparelho. Nesta etapa estratificou-se a validação cruzada pelo **rótulo** (mais crítico, por causa das classes raras) e o teste foi separado por **data**; o desempenho por aparelho é reportado na Seção 19.3."),
  B("**Papel da ordem a priori:** as probabilidades da TED 02 permanecem como roteiro básico do “não liga”; o modelo as complementa, e não as substitui (Seção 21.2)."),
  B("**Idioma:** a hipótese da TED 02 de que seria necessário tradução ou modelo multilíngue foi confirmada pela sondagem em português; um conjunto próprio em português passa a ser requisito para uso real."),
  B("**Variáveis:** idade, marca e desfecho permanecem fora do modelo; o aparelho entra como variável de entrada."),
  H2("21.5 Próximos passos"),
  B("Coletar e rotular, com a Inforcell, uma amostra de relatos em português e avaliar o modelo nela (ou treinar com tradução automática dos textos em inglês, holandês e alemão)."),
  B("Calibrar as probabilidades (calibração isotônica ou sigmoide) e reavaliar o limiar de confiança."),
  B("Testar modelos de linguagem multilíngues (embeddings de frases) como representação do texto, quando houver mais dados."),
  B("Investigar o rótulo “outros” e as confusões energia × inicialização para refinar a taxonomia."));

C.push(H1("22 REGISTRO DE DECISÕES, EXPERIMENTOS E REPRODUÇÃO"),
  H2("22.1 Decisões e experimentos"),
  caption("Tabela", "Registro das decisões e dos experimentos da TED 03"),
  table(["Nº", "Decisão / experimento", "Resultado", "Consequência"], [
    ["1", "Definir classificação multiclasse sintoma + aparelho → categoria de falha", "13 classes; 4.019 registros", "Abordagem justificada na Seção 15.3."],
    ["2", "Dividir por data (treino ≤ 2018; teste 2019–2020)", `${int(RM.n_treino)} / ${int(RM.n_teste)} registros`, "Avalia em dados futuros; o teste não é usado para escolher modelo."],
    ["3", "Validação cruzada agrupada por texto, estratificada por rótulo", "5 dobras", "Evita memorização de frases idênticas e preserva classes raras."],
    ["4", "Linhas de base (classe majoritária e prior por aparelho)", `Acurácia de ${p1(BASE_MAJ.acuracia)}% e ${p1(BASE_AP.acuracia)}%`, "Referência mínima que qualquer modelo deve superar."],
    ["5", "Comparar palavras × caracteres × palavras+caracteres", "Só palavras: F1-macro menor", "Caracteres incluídos por causa do multilinguismo."],
    ["6", "Comparar com e sem o aparelho", "Teste: " + f3(TM.find((t) => t.modelo === "LR · palavras+caracteres").f1_macro) + " → " + f3(FIN.f1_macro), "Aparelho mantido (coerente com a EDA)."],
    ["7", "Comparar LR, SVM e Naive Bayes", "Diferenças pequenas entre LR e SVM; NB inferior", "Escolha baseada em probabilidades e nos requisitos do MVP."],
    ["8", "Avaliar com IC 95% por bootstrap", `F1-macro ${f3(ICF.f1_macro[0])}–${f3(ICF.f1_macro[1])}`, "Resultados lidos como faixas, não como valores exatos."],
    ["9", "Comparar divisão temporal e aleatória", "Diferença pequena (≈ 1 p.p. de acurácia)", "Deriva temporal moderada."],
    ["10", "Curva de aprendizagem", "Ainda crescente", "Mais dados rotulados melhorariam o modelo."],
    ["11", "Análise de confiança e calibração", `ECE = ${f3(RM.calibracao.ece)}; limiar de 0,35`, "Probabilidades usadas para ordenar, não para informar percentuais."],
    ["12", "Caso “não liga” e sondagem em português", `Modelo > ordem fixa; português falha (1 de 8)`, "Registrado nas Seções 20 e 21."],
  ], [450, 3300, 2600, 2721], { size: 16, align: [AlignmentType.CENTER] }),
  fonte("Elaborado pela equipe (2026)."),
  H2("22.2 Código, dados e como reproduzir"),
  B("**src/modelagem.py** — normalização do texto, divisão temporal, linhas de base, catálogo de modelos, validação cruzada, métricas, bootstrap e curva de confiança."),
  B("**src/inferencia.py** — função sugerir_causas, para integração no protótipo do MVP."),
  B("**notebooks/TED03/01_definicao_preparacao.ipynb** — definição do problema, verificações de qualidade e divisão treino/teste."),
  B("**notebooks/TED03/02_treinamento_avaliacao.ipynb** — linhas de base, seis modelos, seleção, avaliação final, matriz de confusão e top-k."),
  B("**notebooks/TED03/03_analise_erros_aplicabilidade.ipynb** — erros, segmentos, curva de aprendizagem, calibração, caso “não liga”, termos e sondagem em português."),
  B("**reports/resultados_modelagem.json** e **reports/figures/fig13…fig21** — todos os números e figuras deste documento são gerados pelos notebooks."),
  B("**tests/** — 24 testes automatizados (pytest -q): dados e divisão temporal sem vazamento, linha de base, top-k e a função sugerir_causas (formato da saída, aparelho e texto inválidos, texto sem sentido com confiança baixa, determinismo)."),
  B("**models/** — modelo avaliado (treino até 2018) e modelo reajustado com todos os registros para o protótipo; **data/processed/** — divisão treino/teste e previsões do teste."),
  P("Para reproduzir: instalar as dependências (pip install -r requirements.txt) e executar, nesta ordem, os notebooks 01, 02 e 03 da pasta notebooks/TED03. Para conferir o ambiente, rodar pytest -q na raiz do repositório. Os dados utilizados são os arquivos da TED 02 em data/processed/ (falhas_modelagem.csv e falhas_rotuladas_limpo.csv), já versionados no repositório; as sementes aleatórias são fixas.", { noIndent: true }));

C.push(H1("23 CONSIDERAÇÕES FINAIS"),
  P(`A TED 03 transformou o problema de diagnóstico da Inforcell em um problema de Machine Learning bem definido: classificar o sintoma relatado, junto com o tipo de aparelho, em uma das 13 categorias de falha, produzindo uma lista ordenada de causas. Foram comparados seis modelos com validação cruzada adequada ao tempo e aos textos repetidos, e o modelo escolhido (regressão logística sobre TF-IDF de palavras, caracteres e aparelho) alcançou F1-macro de ${f2(MF.f1_macro)} e acurácia de ${p1(MF.acuracia)}% em dados futuros, contra ${p1(BASE_AP.acuracia)}% da linha de base, e acerto top-3 de ${p1(MF.top3)}%.`),
  P("A análise crítica mostrou três mensagens para o projeto. **Primeiro**, o texto é a variável que mais contribui, e o modelo é útil para **ordenar hipóteses**, não para dar o diagnóstico final: as falhas que mais exigem testes (placa, inicialização) são as que o modelo menos distingue, e parte dos erros é ambiguidade dos rótulos. **Segundo**, o modelo **não está pronto para relatos em português**; a sondagem e a ausência de português na base indicam que o próximo passo mais importante é reunir dados da própria Inforcell. **Terceiro**, as probabilidades podem ser usadas com segurança para decidir quando pedir mais informações, desde que o limiar seja confirmado com dados reais."),
  P("Assim, a etapa atende ao objetivo da atividade: há um modelo funcional, com decisões justificadas e resultados interpretados, além de limitações assumidas com clareza, e há uma contribuição concreta para o MVP (módulo de sugestão de causas e regra de confiança) e para a evolução do projeto de Inteligência Artificial."));

C.push(H1("REFERÊNCIAS"),
  ...[
    "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. **NBR 14724**: informação e documentação – trabalhos acadêmicos – apresentação. Rio de Janeiro: ABNT, 2024.",
    "BERGSMA, Wicher. A bias-correction for Cramér’s V and Tschuprow’s T. **Journal of the Korean Statistical Society**, v. 42, n. 3, p. 323-328, 2013.",
    "EFRON, Bradley; TIBSHIRANI, Robert J. **An introduction to the bootstrap**. New York: Chapman & Hall, 1993.",
    "GUO, Chuan et al. On calibration of modern neural networks. In: INTERNATIONAL CONFERENCE ON MACHINE LEARNING, 34., 2017, Sydney. **Proceedings** [...]. PMLR, v. 70, p. 1321-1330, 2017.",
    "HARRIS, Charles R. et al. Array programming with NumPy. **Nature**, v. 585, p. 357-362, 2020.",
    "HUNTER, John D. Matplotlib: a 2D graphics environment. **Computing in Science & Engineering**, v. 9, n. 3, p. 90-95, 2007.",
    "MAURYA, Ash. **Running Lean**: iterate from plan A to a plan that works. 3. ed. Sebastopol: O'Reilly Media, 2012.",
    "McKINNEY, Wes. Data structures for statistical computing in Python. In: PYTHON IN SCIENCE CONFERENCE, 9., 2010, Austin. **Proceedings** [...]. Austin: SciPy, 2010. p. 56-61.",
    "OPEN REPAIR ALLIANCE. **Open Repair Data**. Disponível em: https://openrepair.org/open-data/. Acesso em: 28 set. 2026.",
    "OPEN REPAIR ALLIANCE. **Repositório de dados (ORDS v0.3, exportação 202507, e quests MobiFix, CompCat e TabiCat)**. Disponível em: https://github.com/openrepair/data. Acesso em: 28 set. 2026.",
    "PEDREGOSA, Fabian et al. Scikit-learn: machine learning in Python. **Journal of Machine Learning Research**, v. 12, p. 2825-2830, 2011.",
    "SOKOLOVA, Marina; LAPALME, Guy. A systematic analysis of performance measures for classification tasks. **Information Processing & Management**, v. 45, n. 4, p. 427-437, 2009.",
    "THE RESTART PROJECT. **Restarters Workbench**: quests de classificação de falhas. Disponível em: https://restarters.net/workbench. Acesso em: 28 set. 2026.",
    "TUKEY, John W. **Exploratory data analysis**. Reading: Addison-Wesley, 1977.",
    "VIRTANEN, Pauli et al. SciPy 1.0: fundamental algorithms for scientific computing in Python. **Nature Methods**, v. 17, p. 261-272, 2020.",
  ].map((t) => new Paragraph({ children: runs(t), alignment: AlignmentType.LEFT, spacing: { line: 240, after: 200 } })));

C.push(brk(), H1("APÊNDICE A – ORGANIZAÇÃO DO REPOSITÓRIO"),
  P("O repositório utilizado na TED 01 foi mantido e atualizado, preservando os materiais anteriores:", { noIndent: true }),
  ...[
    "projeto-inforcell/",
    "├── README.md                      visão geral, histórico e como reproduzir",
    "├── requirements.txt",
    "├── docs/",
    "│   ├── TED01.pdf                  Documento Técnico v1.0",
    "│   ├── TED02.pdf / TED02.docx     Documento Técnico v2.0",
    "│   └── TED03.pdf / TED03.docx     Documento Técnico v3.0 (este documento)",
    "├── data/",
    "│   ├── raw/                       dados brutos sem alteração + script de download",
    "│   │   └── quests/                MobiFix, CompCat, TabiCat",
    "│   └── processed/                 bases tratadas + dicionário de dados",
    "├── notebooks/",
    "│   ├── TED01/",
    "│   ├── TED02/                     01_carga_inspecao, 02_limpeza_saneamento,",
    "│   │                              03_analise_exploratoria",
    "│   └── TED03/                     01_definicao_preparacao, 02_treinamento_avaliacao,",
    "│                                  03_analise_erros_aplicabilidade",
    "├── src/                           preparacao.py, graficos.py, modelagem.py, inferencia.py",
    "├── models/                        classificador_falhas(.joblib), classificador_falhas_treino2018(.joblib)",
    "└── reports/                       figures/ (fig01…fig21), resultados_eda.json e resultados_modelagem.json",
  ].map((l) => new Paragraph({ spacing: { line: 240, after: 0 }, children: [new TextRun({ text: l, font: "Courier New", size: 18 })] })));

// ---------------------------------------------------------------- documento
const doc = new Document({
  creator: "Equipe Inforcell – UNIBALSAS", title: "Documento Técnico v3.0 – TED 03", description: "Modelagem preditiva e Machine Learning – Inforcell",
  styles: {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, font: FONT }, paragraph: { spacing: { before: 360, after: 240 }, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: false, font: FONT }, paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1, keepNext: true } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, italics: true, font: FONT }, paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2, keepNext: true } },
    ],
  },
  numbering: { config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ] },
  features: { updateFields: true },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1701, left: 1701, bottom: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: [PageNumber.CURRENT], size: 20 })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(path.join(ROOT, "docs", "TED03.docx"), buf); console.log("docx ok", nFig, "figuras", nTab, "tabelas"); });
