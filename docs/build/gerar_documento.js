// Gera docs/TED02.docx — Documento Técnico v2.0 (preserva a TED 01 e acrescenta a TED 02)
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
  ...["Rafael R Pontara", "Matheus de Brito Araujo", "Pedro Junior Ubirajara", "Pedro Paulo T. Araujo"].map((n) => PC(n)),
  new Paragraph({ spacing: { before: 1800 }, children: [] }),
  PC("Seleção, Preparação e Análise Exploratória de Dados para o Assistente de Diagnóstico Inteligente da Inforcell", { run: { bold: true, size: 28 }, after: 200 }),
  PC("Documento Técnico do Projeto — Versão 2.0"), PC("TED 01 (Seleção do conjunto de dados) + TED 02 (Limpeza, saneamento e análise exploratória)", { run: { size: 22 } }),
  new Paragraph({ spacing: { before: 3000 }, children: [] }), PC("Balsas – MA"), PC("2026"), brk());

// Folha de rosto
C.push(...["Rafael R Pontara", "Matheus de Brito Araujo", "Pedro Junior Ubirajara", "Pedro Paulo T. Araujo"].map((n) => PC(n)),
  new Paragraph({ spacing: { before: 2200 }, children: [] }),
  PC("Seleção, Preparação e Análise Exploratória de Dados para o Assistente de Diagnóstico Inteligente da Inforcell", { run: { bold: true, size: 26 }, after: 600 }),
  new Paragraph({ indent: { left: 4535 }, alignment: AlignmentType.JUSTIFIED, spacing: { line: 240 }, children: runs("Documento técnico apresentado à disciplina de Big Data e Ciência de Dados do curso de Sistemas de Informação da UNIBALSAS – Balsas/MA, como requisito parcial de avaliação, referente às Etapas 1 (TED 01 – definição do problema de dados e seleção do conjunto de dados) e 2 (TED 02 – limpeza, saneamento e análise exploratória) do projeto desenvolvido para a empresa Inforcell.", { size: 20 }) }),
  new Paragraph({ spacing: { before: 4200 }, children: [] }), PC("Balsas – MA"), PC("2026"), brk());

// Histórico de versões
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: "HISTÓRICO DE VERSÕES", bold: true })] }),
  table(["Versão", "Etapa", "Data", "Conteúdo"], [
    ["1.0", "TED 01", "2026", "Definição do problema de dados, necessidades de dados e seleção do conjunto Open Repair Data."],
    ["2.0", "TED 02", "26/09/2026", "Inspeção da base, revisão da seleção (Seção 6), limpeza e saneamento com justificativas (Seções 7 a 10), análise exploratória (Seções 11 e 12) e relação com o problema e o MVP (Seção 13). Seções 1 a 5 preservadas da versão 1.0."],
  ], [900, 1100, 1300, 5771], { align: [AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER] }),
  brk(),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: "SUMÁRIO", bold: true })] }),
  new TableOfContents("Sumário", { hyperlink: true, headingStyleRange: "1-2" }), brk());

// ============================ PARTE I — TED 01 (preservada) ============================
C.push(H1("1 IDENTIFICAÇÃO DA EQUIPE"),
  P("Este documento apresenta as Etapas 1 (TED 01) e 2 (TED 02) do projeto desenvolvido na disciplina de Big Data e Ciência de Dados. A Etapa 1 tratou da seleção e avaliação de um conjunto de dados público relacionado ao problema e ao MVP definidos anteriormente na disciplina de Inteligência Artificial, para a empresa Inforcell, assistência técnica de celulares e computadores localizada em Balsas – MA. A Etapa 2 trata da limpeza, do saneamento e da análise exploratória desse conjunto de dados."),
  table(["Item", "Descrição"], [
    ["Instituição", "UNIBALSAS – Balsas/MA"], ["Curso", "Sistemas de Informação"], ["Disciplina", "Big Data e Ciência de Dados"],
    ["Organização / público-alvo atendido", "Inforcell — Assistência técnica de celulares e computadores, Balsas/MA"],
    ["Etapas", "TED 01 — Seleção e avaliação de um conjunto de dados; TED 02 — Limpeza, saneamento e análise exploratória"],
    ["Repositório", "Estrutura descrita no Apêndice A (docs/, data/raw, data/processed, notebooks/TED02, src/)"],
  ], [3000, 6071]), sp(),
  P("**Equipe:**", { noIndent: true }),
  B("Rafael R Pontara"), B("Matheus de Brito Araujo"), B("Pedro Junior Ubirajara"), B("Pedro Paulo T. Araujo"));

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
  B("Transformar as distribuições da Seção 11.5 nas probabilidades iniciais da primeira árvore de diagnóstico (“não liga”)."));

C.push(H1("14 CONSIDERAÇÕES FINAIS"),
  P(`A TED 02 transformou o Open Repair Data, inicialmente bruto, multilíngue e heterogêneo, em duas bases consistentes e documentadas: ${int(R.n_reparos)} registros de reparos de celulares e computadores e ${int(R.n_falhas)} registros com o par sintoma → tipo de falha, dos quais ${int(R.n_modelagem)} estão aptos para modelagem. A inspeção levou à revisão justificada de premissas da TED 01, e cada tratamento foi escolhido de acordo com o uso futuro do dado — em especial, a decisão de não imputar a idade, de manter os reparos em lote e de unificar as taxonomias de falha.`),
  P("A análise exploratória mostrou que os dados confirmam o problema que motivou o projeto: um mesmo sintoma, como “não liga”, está associado a diversas causas, com probabilidades desiguais, e as falhas eletrônicas são as que menos terminam consertadas. Esses resultados dão base empírica para o assistente de diagnóstico guiado da Inforcell e definem o ponto de partida da próxima etapa: a modelagem."));

C.push(H1("REFERÊNCIAS"),
  ...[
    "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. **NBR 14724**: informação e documentação – trabalhos acadêmicos – apresentação. Rio de Janeiro: ABNT, 2024.",
    "BERGSMA, Wicher. A bias-correction for Cramér’s V and Tschuprow’s T. **Journal of the Korean Statistical Society**, v. 42, n. 3, p. 323-328, 2013.",
    "HARRIS, Charles R. et al. Array programming with NumPy. **Nature**, v. 585, p. 357-362, 2020.",
    "HUNTER, John D. Matplotlib: a 2D graphics environment. **Computing in Science & Engineering**, v. 9, n. 3, p. 90-95, 2007.",
    "MAURYA, Ash. **Running Lean**: iterate from plan A to a plan that works. 3. ed. Sebastopol: O'Reilly Media, 2012.",
    "McKINNEY, Wes. Data structures for statistical computing in Python. In: PYTHON IN SCIENCE CONFERENCE, 9., 2010, Austin. **Proceedings** [...]. Austin: SciPy, 2010. p. 56-61.",
    "OPEN REPAIR ALLIANCE. **Open Repair Data**. Disponível em: https://openrepair.org/open-data/. Acesso em: 28 set. 2026.",
    "OPEN REPAIR ALLIANCE. **Repositório de dados (ORDS v0.3, exportação 202507, e quests MobiFix, CompCat e TabiCat)**. Disponível em: https://github.com/openrepair/data. Acesso em: 28 set. 2026.",
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
    "│   └── TED02.pdf / TED02.docx     Documento Técnico v2.0 (este documento)",
    "├── data/",
    "│   ├── raw/                       dados brutos sem alteração + script de download",
    "│   │   └── quests/                MobiFix, CompCat, TabiCat",
    "│   └── processed/                 bases tratadas + dicionário de dados",
    "├── notebooks/",
    "│   ├── TED01/",
    "│   └── TED02/                     01_carga_inspecao, 02_limpeza_saneamento,",
    "│                                  03_analise_exploratoria",
    "├── src/                           preparacao.py, graficos.py",
    "└── reports/                       figures/ e resultados_eda.json",
  ].map((l) => new Paragraph({ spacing: { line: 240, after: 0 }, children: [new TextRun({ text: l, font: "Courier New", size: 18 })] })));

// ---------------------------------------------------------------- documento
const doc = new Document({
  creator: "Equipe Inforcell – UNIBALSAS", title: "Documento Técnico v2.0 – TED 02", description: "Limpeza, saneamento e análise exploratória – Inforcell",
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
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(path.join(ROOT, "docs", "TED02.docx"), buf); console.log("docx ok", nFig, "figuras", nTab, "tabelas"); });
