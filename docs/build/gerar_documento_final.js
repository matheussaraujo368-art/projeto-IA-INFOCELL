// Gera docs/documento-final.docx — Documento Técnico, Versão Final (TED 04).
// Consolida as TEDs 01 a 03 nas 13 seções pedidas pela TED 04 e acrescenta a interpretação integrada,
// a análise ética, a privacidade e LGPD e o Data Storytelling.
// Uso: node docs/build/gerar_documento_final.js   (requer o pacote npm "docx")
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, AlignmentType,
  HeadingLevel, LevelFormat, WidthType, ShadingType, BorderStyle, PageBreak, TableOfContents,
  Header, PageNumber,
} = require("docx");

const ROOT = path.resolve(__dirname, "..", "..");
const FIG = path.join(ROOT, "reports", "figures");
const ler = (f) => JSON.parse(fs.readFileSync(path.join(ROOT, "reports", f), "utf8"));
const R = ler("resultados_eda.json");
const RM = ler("resultados_modelagem.json");
const R4 = ler("resultados_ted04.json");

const FONT = "Arial";
const W = 9071; // largura útil A4 com margens ABNT (3 cm esq., 2 cm dir.)
const fmt = (n, d = 1) => Number(n).toLocaleString("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d });
const int = (n) => Number(n).toLocaleString("pt-BR");
const p1 = (x) => fmt(x * 100, 1);
const p0 = (x) => fmt(x * 100, 0);
const f2 = (x) => fmt(x, 2);
const f3 = (x) => fmt(x, 3);

// ---------------------------------------------------------------- helpers (mesmo padrão da v3.0)
function runs(text, base = {}) {
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
  spacing: { line: 360, after: 120, ...(o.spacing || {}) }, indent: o.noIndent ? undefined : { firstLine: 709 }, keepNext: o.keepNext,
});
const PC = (text, o = {}) => new Paragraph({ children: runs(text, o.run || {}), alignment: AlignmentType.CENTER, spacing: { line: 360, after: o.after ?? 0, before: o.before ?? 0 } });
const H1 = (t, quebra = true) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)], pageBreakBefore: quebra });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const B = (text) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: runs(text), alignment: AlignmentType.JUSTIFIED, spacing: { line: 360, after: 60 } });
const N = (text) => new Paragraph({ numbering: { reference: "num", level: 0 }, children: runs(text), alignment: AlignmentType.JUSTIFIED, spacing: { line: 360, after: 60 } });
const brk = () => new Paragraph({ children: [new PageBreak()] });
const sp = () => new Paragraph({ spacing: { after: 120 }, children: [] });

let nFig = 0, nTab = 0;
const border = { style: BorderStyle.SINGLE, size: 4, color: "A6A6A6" };
const borders = { top: border, bottom: border, left: border, right: border };
function caption(kind, title) {
  const n = kind === "Figura" ? ++nFig : ++nTab;
  if (process.env.LISTAR) console.log(`${kind} ${n}: ${title}`);
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
const tab = (titulo, headers, rows, widths, o = {}, origem) => [caption("Tabela", titulo), table(headers, rows, widths, o), fonte(origem || "Elaborado pela equipe (2026).")];
function pngSize(file) { const b = fs.readFileSync(file); return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) }; }
function figura(nome, titulo, larguraPx = 600, origem) {
  const file = path.join(FIG, nome + ".png");
  const { w, h } = pngSize(file);
  return [caption("Figura", titulo),
    new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, children: [new ImageRun({ type: "png", data: fs.readFileSync(file), transformation: { width: larguraPx, height: Math.round((h / w) * larguraPx) }, altText: { title: titulo, description: titulo, name: nome } })] }),
    fonte(origem)];
}
function nota(titulo, texto) {
  const nada = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: W, type: WidthType.DXA }, borders: { top: nada, bottom: nada, right: nada, left: { style: BorderStyle.SINGLE, size: 24, color: "2A78D6" } },
      shading: { fill: "EEF4FB", type: ShadingType.CLEAR, color: "auto" }, margins: { top: 100, bottom: 100, left: 180, right: 140 },
      children: [new Paragraph({ spacing: { line: 276, after: 60 }, children: [new TextRun({ text: titulo, bold: true, size: 20, color: "1C5CAB" })] }),
        new Paragraph({ spacing: { line: 276 }, alignment: AlignmentType.JUSTIFIED, children: runs(texto, { size: 20 }) })],
    })] })],
  });
}
const L = AlignmentType.LEFT, RT = AlignmentType.RIGHT, CT = AlignmentType.CENTER;

// ---------------------------------------------------------------- dados derivados
const st = Object.fromEntries(R.status_categoria.map((r) => [r.categoria, r]));
const tf = Object.fromEntries(R.taxa_por_falha.map((r) => [r.categoria_falha, r]));
const fc = R.falha_por_cat;
const nl = R.nao_liga;
const barr = R.barreiras; const nInf = R.n_fim_vida - barr["Não informada"];
const MF = RM.final.metricas, ICF = RM.final.ic95, PCL = RM.por_classe, TM = RM.tabela_modelos;
const BASE_AP = TM.find((t) => t.modelo === "Base · prior por aparelho");
const BASE_MAJ = TM.find((t) => t.modelo === "Base · classe majoritária");
const FIN = TM.find((t) => t.modelo === RM.modelo_final);
const SEM_AP = TM.find((t) => t.modelo === "LR · palavras+caracteres");
const TOP = RM.topk[RM.modelo_final];
const nlg = RM.nao_liga;
const sel = (c) => RM.selecao_confianca.find((x) => Math.abs(x.corte - c) < 1e-6);
const seg = (k, g) => RM.segmentos[k].find((x) => x.grupo === g);
const nome = (m) => m.replace("Base · ", "Base: ");
const PRIV = R4.privacidade;
const EQUIPE = ["Rafael R Pontara", "Matheus de Brito Araujo", "Pedro Paulo T. Araujo"];
const TITULO = "Do sintoma à causa provável: dados, modelagem e uso responsável no Assistente de Diagnóstico Inteligente da Inforcell";

const C = [];

// ================================================================ elementos pré-textuais
C.push(PC("UNIBALSAS – CENTRO UNIVERSITÁRIO DE BALSAS", { run: { bold: true } }), PC("Curso de Sistemas de Informação"), PC("Disciplina de Big Data e Ciência de Dados"),
  new Paragraph({ spacing: { before: 1600 }, children: [] }), ...EQUIPE.map((n) => PC(n)),
  new Paragraph({ spacing: { before: 1800 }, children: [] }),
  PC(TITULO, { run: { bold: true, size: 28 }, after: 200 }),
  PC("Documento Técnico do Projeto — Versão Final (4.0)"),
  PC("TED 04 – Integração dos resultados, ética, LGPD e Data Storytelling", { run: { size: 22 } }),
  new Paragraph({ spacing: { before: 3000 }, children: [] }), PC("Balsas – MA"), PC("2026"), brk());

C.push(...EQUIPE.map((n) => PC(n)), new Paragraph({ spacing: { before: 2200 }, children: [] }),
  PC(TITULO, { run: { bold: true, size: 26 }, after: 600 }),
  new Paragraph({ indent: { left: 4535 }, alignment: AlignmentType.JUSTIFIED, spacing: { line: 240 }, children: runs("Documento técnico apresentado à disciplina de Big Data e Ciência de Dados do curso de Sistemas de Informação da UNIBALSAS – Balsas/MA, como requisito parcial de avaliação, referente à Etapa 4 (TED 04 – integração dos resultados, ética, LGPD e Data Storytelling). Consolida as Etapas 1 a 3 do projeto desenvolvido para a empresa Inforcell.", { size: 20 }) }),
  new Paragraph({ spacing: { before: 4200 }, children: [] }), PC("Balsas – MA"), PC("2026"), brk());

C.push(PC("HISTÓRICO DE VERSÕES", { run: { bold: true }, after: 240 }),
  table(["Versão", "Etapa", "Data", "Conteúdo"], [
    ["1.0", "TED 01", "19/09/2026", "Definição do problema de dados, necessidades de dados e seleção do conjunto Open Repair Data."],
    ["2.0", "TED 02", "26/09/2026", "Revisão da seleção, limpeza e saneamento com justificativas, análise exploratória e relação com o problema e o MVP."],
    ["3.0", "TED 03", "03/11/2026", "Problema de Machine Learning, preparação, seis modelos comparados a duas linhas de base, resultados, interpretação, limitações e relação com o MVP."],
    ["4.0 (final)", "TED 04", "01/12/2026", "Reorganização nas 13 seções da versão final; interpretação integrada; correção do estado de integração com o MVP; análise ética, vieses e mitigação; privacidade e LGPD (base pública e sistema em produção); Data Storytelling; conclusões com evidências; evolução. Novas Figuras 1 e 12 (funil dos dados e representatividade)."],
  ], [1200, 1000, 1300, 5571], { align: [CT, CT, CT] }), sp(),
  P("As versões 1.0, 2.0 e 3.0 continuam no repositório (docs/TED01.pdf, docs/TED02.pdf e docs/TED03.pdf), com todos os detalhes de cada etapa. Esta versão final não as substitui: reúne o que é necessário para entender o projeto de ponta a ponta e remete às versões anteriores quando um detalhe não é repetido (Apêndice B).", { noIndent: true }),
  brk());

C.push(PC("RESUMO", { run: { bold: true }, after: 240 }),
  P(`A Inforcell, assistência técnica de celulares e computadores em Balsas/MA, enfrenta um problema comum ao setor: um mesmo sintoma, como “o aparelho não liga”, pode ter várias causas, e a escolha dos testes depende da experiência de cada técnico. Este documento consolida as quatro etapas da disciplina de Big Data e Ciência de Dados realizadas para apoiar o assistente de diagnóstico do projeto de Inteligência Artificial. A partir da base pública Open Repair Data, foram tratados ${int(R.n_reparos)} registros de reparos e ${int(R.n_falhas)} registros com o tipo de falha atribuído por voluntários. A análise exploratória mostrou que o sintoma “não liga” se distribui por ${nl.n_categorias} categorias de falha, com cerca de 80% em energia/bateria e inicialização, e que as falhas eletrônicas são as menos consertadas. Um classificador de regressão logística sobre o texto do sintoma e o tipo de aparelho alcançou, em eventos posteriores aos do treino, acurácia de ${p1(MF.acuracia)}%, F1-macro de ${f3(MF.f1_macro)} e acerto de ${p1(MF.top3)}% entre as três causas mais prováveis, contra ${p1(BASE_AP.acuracia)}% de acurácia da melhor linha de base. Os resultados sustentam o uso do modelo para ordenar hipóteses, e não para dar o diagnóstico final. As principais limitações são a ausência de registros do Brasil e de relatos em português, a ambiguidade dos rótulos e as classes raras. A análise ética e de proteção de dados indica que a base pública tem risco baixo, mas não nulo, de conter dados pessoais em texto livre, e que o sistema em produção trata dados de clientes, o que exige finalidade, necessidade e segurança definidas. O modelo ainda não está integrado ao sistema web e sua utilidade para a Inforcell precisa ser confirmada com dados próprios.`, { noIndent: true }),
  P("**Palavras-chave:** ciência de dados; classificação de texto; diagnóstico de eletrônicos; LGPD; Data Storytelling.", { noIndent: true }),
  brk(),
  PC("SUMÁRIO", { run: { bold: true }, after: 240 }),
  new TableOfContents("Sumário", { hyperlink: true, headingStyleRange: "1-2" }));

// ================================================================ 1
C.push(H1("1 CONTEXTO E DEFINIÇÃO DO PROBLEMA"),
  H2("1.1 Organização atendida e equipe"),
  ...tab("Identificação do projeto", ["Item", "Descrição"], [
    ["Instituição / curso", "UNIBALSAS – Balsas/MA · Sistemas de Informação"],
    ["Disciplina", "Big Data e Ciência de Dados (integrada aos projetos de Inteligência Artificial e Programação para Web)"],
    ["Organização atendida", "Inforcell – assistência técnica de celulares e computadores, Balsas/MA"],
    ["Equipe", "Rafael R Pontara; Matheus de Brito Araujo; Pedro Paulo T. Araujo. Pedro Junior Ubirajara participou das Etapas 1 e 2"],
    ["Repositório", "https://github.com/matheussaraujo368-art/projeto-IA-INFOCELL"],
  ], [2700, 6371]),
  H2("1.2 O problema"),
  P("Na Inforcell, o técnico recebe um aparelho com um sintoma descrito em poucas palavras pelo cliente (“não liga”, “não carrega”, “tela piscando”) e precisa decidir por onde começar. O mesmo sintoma pode ter causas diferentes — bateria, conector, circuito de carga, placa, software — e a ordem dos testes depende hoje quase só da experiência individual. Além disso, as Ordens de Serviço eram acompanhadas pelo WhatsApp, o que dispersa o histórico que poderia orientar diagnósticos futuros."),
  P("**Por que é relevante.** Testes na ordem errada custam tempo de bancada, atrasam a entrega e podem levar à troca de peças sem necessidade. Do ponto de vista social, a TED 02 mostrou que a falta de informação de reparo é uma das principais razões para um aparelho ser descartado (Seção 4.4)."),
  H2("1.3 Objetivos da análise de dados"),
  B("Verificar com dados reais se um mesmo sintoma leva a várias causas e se essas causas têm probabilidades diferentes."),
  B("Identificar quais informações disponíveis na abertura da Ordem de Serviço ajudam a prever a causa."),
  B("Construir e avaliar um modelo que sugira as causas mais prováveis a partir do relato e do tipo de aparelho."),
  B("Avaliar limites, riscos éticos e cuidados de proteção de dados do uso dessas análises no assistente."),
  H2("1.4 A solução proposta no projeto de Inteligência Artificial"),
  P("A solução é uma plataforma web que une a gestão de Ordens de Serviço a um assistente de IA que conduz o técnico por um diagnóstico guiado: sintoma → causas possíveis → um teste por vez → interpretação do resultado → diagnóstico final, que é sempre decidido pelo técnico. No estado atual do MVP (novembro de 2026), o assistente é um modelo de linguagem da Anthropic acessado por API, com ferramentas de consulta somente leitura aos dados da própria assistência (casos anteriores, base de conhecimento, medições e preços)."),
  nota("Correção em relação à versão 3.0", "A Seção 21.3 da versão 3.0 descreveu no presente a entrada do classificador no MVP. Na verdade, **o sistema web ainda não chama a função sugerir_causas()** e o campo “causa final” da Ordem de Serviço ainda não existe. Nesta versão, a integração é apresentada como proposta (Seções 7.3 e 13). Além disso, até 26/11/2026 o assistente do sistema não havia sido testado com uma chave real da API: as evidências da disciplina de IA usaram respostas simuladas."),
  sp(),
  H2("1.5 O percurso das quatro etapas"),
  ...tab("Contribuição de cada etapa para a seguinte", ["Etapa", "Pergunta", "Principal resultado", "O que levou para a etapa seguinte"], [
    ["TED 01", "Que dados apoiam o problema?", "Seleção do Open Repair Data, a base pública mais próxima de “sintoma → causa → resultado”", "Lista de necessidades de dados e limitações a verificar (idioma, contexto comunitário)"],
    ["TED 02", "Os dados confirmam o problema?", "Duas bases tratadas; “não liga” → 11 causas, ~80% em energia e boot; falhas eletrônicas menos consertadas", "Variáveis relevantes (texto e aparelho), divisão por tempo, taxonomia de 14 categorias e 4.019 registros aptos"],
    ["TED 03", "É possível prever a causa?", `Modelo com top-3 de ${p1(MF.top3)}%, contra ${p1(BASE_AP.top3)}% da linha de base; limiar de confiança de 0,35`, "Papel do modelo (ordenar hipóteses), limites (português, classes raras) e regra de uso"],
    ["TED 04", "O que se pode afirmar e com que cuidados?", "Interpretação integrada, ética, LGPD e narrativa dos resultados", "Recomendações para o MVP e para a coleta de dados próprios"],
  ], [1000, 2000, 3171, 2900], { size: 17 }));

// ================================================================ 2
C.push(H1("2 CONJUNTO DE DADOS UTILIZADO"),
  H2("2.1 Origem, versão e licença"),
  ...tab("Identificação da base de dados", ["Item", "Descrição"], [
    ["Base", "Open Repair Data, padrão ORDS v0.3 (Open Repair Data Standard)"],
    ["Origem", "Open Repair Alliance: rede de organizações de eventos comunitários de reparo (The Restart Project, Repair Café International, anstiftung, entre outras)"],
    ["Versão", "Exportação 202507 (dados até 31/07/2025), repositório github.com/openrepair/data, commit 90eba80 (22/03/2026), mais as bases rotuladas das quests MobiFix, CompCat e TabiCat"],
    ["Licença", "Creative Commons BY-SA 4.0: exige atribuição e que derivações publicadas mantenham a mesma licença"],
    ["Período e alcance", `${R.periodo[0].split("-").reverse().join("/")} a ${R.periodo[1].split("-").reverse().join("/")}, ${R.n_paises} países`],
  ], [2300, 6771]),
  H2("2.2 Características e relação com o problema"),
  P("Cada registro descreve uma tentativa de reparo feita por voluntários: o tipo e a marca do aparelho, o ano de fabricação, o relato do problema em texto livre e o desfecho (consertado, reparável ou fim de vida). As bases das quests acrescentam o **tipo de falha** atribuído por voluntários que leram cada relato. Essa combinação permite estudar o par sintoma → causa, que é exatamente o que o assistente da Inforcell precisa sugerir. Nenhuma outra base pública avaliada na TED 01 oferecia esse par para eletrônicos de consumo."),
  P("A inspeção da TED 02 mudou o planejamento: a base agregada atual não tem o campo de tipo de falha nem o de idioma previstos na TED 01. Por isso, o projeto passou a trabalhar com duas bases complementares (Tabela 4)."),
  ...tab("As duas bases do projeto", ["Base", "Origem", "Registros tratados", "Uso"], [
    ["A – Reparos", "Base agregada ORDS, 4 categorias de interesse", int(R.n_reparos), "Contexto do atendimento: aparelho, marca, idade, desfecho, barreiras"],
    ["B – Falhas rotuladas", "Quests MobiFix (celulares), CompCat (notebooks e desktops) e TabiCat (tablets)", int(R.n_falhas), "Par sintoma → tipo de falha; base da modelagem"],
  ], [1700, 3100, 1500, 2771], { size: 18, align: [L, L, RT, L] }),
  H2("2.3 Restrições de uso"),
  B("**Licença:** os dados tratados publicados no repositório seguem a CC BY-SA 4.0, com atribuição à Open Repair Alliance (data/raw/README.md)."),
  B("**Tamanho:** o arquivo completo (58 MB) não é versionado; o repositório guarda o subconjunto bruto das 4 categorias, sem alteração, e um script que baixa o restante."),
  B("**Contexto:** são dados de eventos comunitários na Europa e na América do Norte, não de uma assistência comercial brasileira. Essa restrição atravessa todas as conclusões (Seção 10)."));

// ================================================================ 3
C.push(H1("3 PREPARAÇÃO E TRATAMENTO DOS DADOS"),
  P("A preparação foi feita na TED 02, em Python (pandas, NumPy e SciPy), com as regras concentradas em src/preparacao.py e aplicadas pelos notebooks de notebooks/TED02. A base bruta foi lida como texto, para que cada conversão de tipo fosse uma decisão explícita. A Figura 1 mostra quantos registros restaram em cada etapa, da base completa ao conjunto de teste da TED 03."),
  ...figura("fig22_funil_dados", "Do dado bruto ao modelo: registros em cada etapa", 600),
  P("**Como ler.** A grande redução inicial é de escopo, não de qualidade: das 305.649 tentativas de reparo, só 23.612 são de celulares, notebooks, desktops e tablets. A segunda redução importante acontece na Base B: dos 6.657 registros rotulados, 4.019 têm texto, uma causa informativa e não são duplicatas. É esse o volume que sustenta a modelagem, e ele é pequeno para um problema com 13 classes."),
  H2("3.1 Problemas de qualidade e decisões"),
  ...tab("Principais problemas de qualidade e tratamentos escolhidos", ["Problema", "Evidência", "Decisão", "Por quê"], [
    ["Campos previstos inexistentes", "Sem tipo de falha e idioma na base agregada", "Usar as quests (Base B) e criar o idioma por detecção", "Sem o tipo de falha não há par sintoma → causa"],
    ["Idade ausente", "69% sem ano de fabricação", "Não imputar", "Imputar criaria idade artificial em 2 de cada 3 registros"],
    ["Valores impossíveis", "Anos futuros ou anteriores à tecnologia", "Converter em ausente", "O restante do registro continua útil"],
    ["Itens fora de escopo", "“Tablet” incluía GPS, e-readers e estações meteorológicas", "Remover 1.079 registros", "Não são atendidos pela Inforcell"],
    ["Marcas com grafias diferentes", "942 grafias", "Padronizar para 685 marcas", "A Apple apareceria dividida em mais de dez marcas"],
    ["Duplicatas", "2.669 cópias com IDs distintos (reparos em lote)", "Manter e sinalizar na Base A; remover texto + rótulo repetidos na Base B", "Na Base A são aparelhos diferentes; na modelagem, frases repetidas seriam “decoradas”"],
    ["Rótulos pouco informativos", "“Unknown” e “Poor data” em 26% dos rotulados", "Agrupar como “Indeterminado” e excluir da modelagem", "Não informam a causa"],
    ["Taxonomias diferentes", "41 rótulos originais em listas distintas", "Taxonomia unificada de 14 categorias em português", "Corresponde aos “ramos” do diagnóstico de bancada"],
  ], [1900, 2200, 2300, 2671], { size: 16 }, "Elaborado pela equipe (2026). Detalhes nas Seções 8 a 10 da versão 2.0."),
  P(`**O que essas decisões têm em comum.** Cada tratamento foi escolhido pelo uso futuro do dado: preservar informação quando ela continua útil (não imputar, não apagar duplicatas da Base A) e remover quando ela distorceria a modelagem (duplicatas de texto e rótulos indeterminados). O detector de idioma criado pela equipe foi validado contra os 647 rótulos manuais da TabiCat, com 96,5% de acurácia.`),
  H2("3.2 Preparação específica para a modelagem (TED 03)"),
  B("**Variáveis de entrada:** apenas o que existe na abertura da Ordem de Serviço — texto do sintoma e tipo de aparelho. Idade (69% ausente), marca (34% desconhecida) e desfecho (só conhecido depois do reparo, o que seria vazamento de informação) ficaram de fora."),
  B("**Texto:** minúsculas, remoção de URLs, números e pontuação; TF-IDF de palavras (1 e 2 palavras) e de caracteres (2 a 5), para lidar com quatro idiomas e erros de digitação."),
  B(`**Divisão por tempo:** treino com eventos até ${RM.ano_corte} (${int(RM.n_treino)} registros) e teste com 2019 e 2020 (${int(RM.n_teste)}), simulando o uso em casos futuros.`),
  B("**Desbalanceamento:** pesos de classe inversamente proporcionais à frequência; dano por líquido tem só 7 exemplos de treino."));

// ================================================================ 4
C.push(H1("4 ANÁLISE EXPLORATÓRIA"),
  P("A análise priorizou tabelas de frequência, gráficos de barras, boxplots e o V de Cramér, porque quase todas as variáveis são categóricas. Esta seção reúne os achados que sustentam as conclusões do projeto; as demais análises estão na Seção 11 da versão 2.0."),
  H2("4.1 Estatísticas gerais"),
  ...tab("Indicadores gerais das bases tratadas", ["Indicador", "Valor", "Leitura"], [
    ["Registros de reparo (Base A)", int(R.n_reparos), "53,9% notebooks, 23,0% celulares"],
    ["Taxa geral de conserto", `${fmt(R.taxa_conserto_geral)}%`, "Pouco mais da metade dos aparelhos sai consertada"],
    ["Relatos com texto", `${fmt(R.pct_com_texto)}%`, `Mediana de ${fmt(R.palavras_desc["50%"], 0)} palavras por relato`],
    ["Registros com idade válida", `${fmt(R.pct_com_idade)}%`, "Limita qualquer conclusão sobre idade"],
    ["Registros com falha indeterminada (Base B)", `${fmt(R.pct_indeterminado)}%`, "Nem voluntários experientes identificam a causa em parte dos relatos"],
    ["Correlação idade × conserto (Spearman)", `ρ = ${fmt(R.spearman_idade_conserto.rho, 3)}`, "Praticamente nula"],
  ], [3500, 1600, 3971], { size: 18, align: [L, RT, L] }, "Elaborado pela equipe (2026), a partir do notebook 03 da TED 02."),
  H2("4.2 O tipo de aparelho define o ponto de partida"),
  ...figura("fig09_falhas_aparelho", "Distribuição dos tipos de falha por categoria de aparelho", 500),
  P(`**O que se observa.** Nos celulares, a tela domina (${fmt(fc["Celular"]["Tela"], 0)}% dos casos), seguida de energia/bateria (${fmt(fc["Celular"]["Energia / bateria"], 0)}%); nos notebooks, predominam software (${fmt(fc["Notebook"]["Software / sistema"], 0)}%) e desempenho/superaquecimento (${fmt(fc["Notebook"]["Desempenho / superaquecimento"], 0)}%). A associação entre aparelho e tipo de falha (V de Cramér = 0,28) é a mais forte entre as variáveis do problema. **O que significa.** O diagnóstico deve começar com hipóteses diferentes para cada tipo de aparelho.`),
  H2("4.3 Um sintoma, várias causas — com probabilidades desiguais"),
  ...figura("fig11_nao_liga", "Tipos de falha atribuídos a registros com o sintoma “não liga”", 520),
  P(`**O que se observa.** Os ${int(nl.n)} relatos de aparelhos que “não ligam” foram atribuídos a **${nl.n_categorias} categorias de falha**: energia/bateria (${fmt(nl.distribuicao["Energia / bateria"], 0)}%), inicialização (${fmt(nl.distribuicao["Inicialização (boot)"], 0)}%), placa/dano interno (${fmt(nl.distribuicao["Placa / dano interno"], 0)}%) e outras. **O que significa.** Este é o achado central do projeto: confirma com dados o problema relatado pela Inforcell e, ao mesmo tempo, mostra que os testes podem ser ordenados por probabilidade (alimentação → inicialização → placa).`),
  H2("4.4 As falhas eletrônicas são as mais difíceis"),
  ...figura("fig10_conserto_falha", "Taxa de conserto por tipo de falha", 540),
  P(`**O que se observa.** Software, armazenamento e desempenho são resolvidos em 65% a 68% dos casos; placa (${fmt(tf["Placa / dano interno"].taxa, 0)}%), energia/bateria (${fmt(tf["Energia / bateria"].taxa, 0)}%) e inicialização (${fmt(tf["Inicialização (boot)"].taxa, 0)}%) têm as menores taxas. Entre os aparelhos em fim de vida com motivo informado, ${fmt(barr["Falta de informação de reparo"] / nInf * 100, 0)}% foram descartados por **falta de informação de reparo**. **O que significa.** O apoio ao técnico tende a gerar mais valor justamente nas falhas que exigem medições em sequência.`),
  H2("4.5 O relato inicial é pobre"),
  ...figura("fig04_palavras", "Distribuição do número de palavras na descrição do sintoma", 520),
  P(`**O que se observa.** ${fmt(R.pct_ate_5_palavras)}% dos relatos têm até 5 palavras (“broken screen”, “won’t charge”). **O que significa.** O relato sozinho não basta para o diagnóstico: o assistente precisa fazer perguntas complementares e propor testes. Esse achado também limita o que qualquer modelo de texto pode alcançar.`));

// ================================================================ 5
C.push(H1("5 MODELAGEM DE MACHINE LEARNING"),
  H2("5.1 O problema de aprendizado"),
  P("**Pergunta:** dado o sintoma relatado e o tipo de aparelho, quais são as categorias de falha mais prováveis? Trata-se de **classificação supervisionada multiclasse** com 13 classes (as 14 da taxonomia menos “Indeterminado”) e saída em probabilidades, para que o assistente receba uma lista ordenada de hipóteses. Regressão não se aplica, porque não há valor numérico a prever; agrupamento não daria nomes de causa para mostrar ao técnico."),
  ...tab("Variáveis do modelo", ["Variável", "Papel", "Justificativa"], [
    ["Texto do sintoma", "Entrada", "É o próprio relato; contém termos que distinguem categorias (Seção 6.5)"],
    ["Tipo de aparelho", "Entrada", "Associado ao tipo de falha (V = 0,28) e sempre conhecido na abertura da OS"],
    ["Categoria de falha", "Alvo", "13 classes muito desbalanceadas: de 804 a 18 registros"],
    ["Idade, marca, desfecho, país, organização", "Descartadas", "Ausência alta, vazamento de informação ou viés de coleta que não existirá na Inforcell"],
  ], [2700, 1300, 5071], { size: 18 }),
  H2("5.2 Técnicas comparadas"),
  P("Foram comparadas **duas linhas de base** — prever sempre a classe mais frequente e usar só a distribuição de falhas por aparelho (a ideia da TED 02) — e **seis modelos lineares** sobre TF-IDF: regressão logística com palavras, com caracteres, com ambos e com ambos mais o aparelho; SVM linear; e Naive Bayes complementar. Modelos lineares regularizados foram escolhidos porque há poucos textos (2.224 de treino) e muitas variáveis esparsas, situação em que modelos profundos tendem a sobreajustar."),
  ...figura("fig13_divisao_dados", "Divisão temporal em treino e teste e distribuição das classes", 580),
  P("**Como ler.** O teste é grande (45% dos dados) porque 2019 concentra muitos registros, e a proporção das classes muda entre os períodos (por exemplo, “tela” cai e “software” sobe). O desempenho medido já inclui essa mudança, como aconteceria no uso real."),
  H2("5.3 Treinamento e avaliação"),
  B("**Seleção de hiperparâmetros:** validação cruzada de 5 dobras dentro do treino, estratificada pelo rótulo e agrupada por texto idêntico, com F1-macro como critério."),
  B("**Avaliação final:** uma única avaliação no teste, com intervalos de confiança de 95% por bootstrap (1.000 reamostragens)."),
  B("**Métricas:** F1-macro (principal, porque dá o mesmo peso às classes raras), acurácia, precisão e revocação por classe, matriz de confusão, acerto top-k (a causa certa entre as k primeiras, a métrica mais próxima do uso no assistente) e calibração (ECE)."),
  B("**Reprodutibilidade:** sementes fixas; os notebooks de notebooks/TED03 regeneram todos os números e figuras; 24 testes automatizados (pytest)."));

// ================================================================ 6
C.push(H1("6 RESULTADOS OBTIDOS"),
  H2("6.1 Comparação com as linhas de base"),
  ...tab("Linhas de base e modelos candidatos (teste: eventos de 2019 e 2020)", ["Modelo", "F1-macro (validação cruzada)", "Acurácia", "F1-macro", "Top-3"],
    TM.map((t) => [t.modelo === RM.modelo_final ? `**${nome(t.modelo)}** ★` : nome(t.modelo),
      t.cv_f1_macro == null ? "—" : `${f3(t.cv_f1_macro)} ± ${f3(t.cv_dp)}`, `${p1(t.acuracia)}%`, f3(t.f1_macro), `${p1(t.top3)}%`]),
    [3500, 1800, 1250, 1250, 1271], { size: 16, align: [L, CT, RT, RT, RT] }, "Elaborado pela equipe (2026), a partir do notebook 02 da TED 03. ★ modelo escolhido."),
  P(`**O que se observa.** Todos os modelos que usam o texto superam com folga as linhas de base. Entre eles, as diferenças são pequenas e ficam dentro de um desvio-padrão na validação cruzada. A regressão logística com palavras, caracteres e aparelho foi escolhida por fornecer probabilidades e usar todas as informações disponíveis na abertura da OS. **Não se afirma que ela seja superior aos demais candidatos**, e sim que atende aos requisitos com desempenho equivalente.`),
  H2("6.2 Desempenho do modelo escolhido"),
  ...tab(`Desempenho no teste (n = ${int(RM.n_teste)})`, ["Métrica", "Modelo", "IC 95%", "Linha de base (prior por aparelho)"], [
    ["Acurácia", `${p1(MF.acuracia)}%`, `${p1(ICF.acuracia[0])}% a ${p1(ICF.acuracia[1])}%`, `${p1(BASE_AP.acuracia)}%`],
    ["F1-macro", f3(MF.f1_macro), `${f3(ICF.f1_macro[0])} a ${f3(ICF.f1_macro[1])}`, f3(BASE_AP.f1_macro)],
    ["Acerto top-3", `${p1(MF.top3)}%`, `${p1(ICF.top3[0])}% a ${p1(ICF.top3[1])}%`, `${p1(BASE_AP.top3)}%`],
    ["Acerto top-5", `${p1(MF.top5)}%`, "—", `${p1(BASE_AP.top5)}%`],
  ], [2600, 1700, 2300, 2471], { size: 18, align: [L, CT, CT, CT] }, "Elaborado pela equipe (2026), a partir do notebook 02 da TED 03."),
  ...figura("fig17_topk", "Acerto top-k do modelo escolhido e das linhas de base", 460),
  P(`**O que se observa.** A causa correta é a primeira sugestão em ${p1(TOP["1"])}% dos casos e está entre as três primeiras em **${p1(TOP["3"])}%**. **O que significa.** Para um assistente que propõe testes em sequência, mostrar as três causas mais prováveis já cobre cerca de nove em cada dez casos — muito mais útil do que exigir uma única resposta correta.`),
  H2("6.3 Desempenho por categoria e principais confusões"),
  ...figura("fig15_f1_por_classe", "F1 por categoria de falha: modelo escolhido × prior por aparelho", 500),
  P(`**O que se observa.** O modelo vai bem nas categorias com vocabulário próprio — tela (F1 ${f2(PCL["Tela"].f1)}), energia/bateria (${f2(PCL["Energia / bateria"].f1)}), conectores (${f2(PCL["Conectores e portas"].f1)}) — e mal em placa (${f2(PCL["Placa / dano interno"].f1)}), inicialização (${f2(PCL["Inicialização (boot)"].f1)}), carcaça e dano por líquido (${f2(PCL["Dano por líquido"].f1)}, com só ${int(PCL["Dano por líquido"].n)} exemplos de teste).`),
  ...figura("fig16_matriz_confusao", "Matriz de confusão do modelo escolhido (cada linha soma 100%)", 520),
  P(`**O que se observa.** Os erros se concentram em pares tecnicamente próximos: software ↔ desempenho e **energia/bateria → inicialização** (${int(RM.confusoes[1].n)} casos). **O que significa.** É a mesma sobreposição de vocabulário vista na análise exploratória (“turn”, “start”, “won’t”): a ambiguidade está no relato, e só os testes de bancada a resolvem.`),
  H2("6.4 Confiança das previsões"),
  ...figura("fig20_confianca", "Calibração das probabilidades e relação entre confiança, acerto e cobertura", 570),
  P(`**O que se observa.** O acerto cresce com a confiança: com probabilidade da primeira causa de pelo menos 0,50, o acerto é de ${p1(sel(0.5).acuracia)}%, cobrindo ${p0(sel(0.5).cobertura)}% dos casos. As probabilidades são **subestimadas** (ECE = ${f3(RM.calibracao.ece)}). **O que significa.** Elas servem para ordenar as causas e para decidir quando pedir mais informações (limiar de 0,35: ${p0(sel(0.35).cobertura)}% dos casos com ${p0(sel(0.35).acuracia)}% de acerto), mas não devem ser exibidas ao técnico como “chance de acerto”.`),
  H2("6.5 O caso “não liga” e o que o modelo aprendeu"),
  ...figura("fig21_nao_liga", "Acerto da ordem fixa (a priori) e do modelo para o sintoma “não liga”", 430),
  P(`**O que se observa.** Nos ${int(nlg.n_teste)} relatos de teste com esse sintoma, o modelo acerta a primeira causa em ${p0(nlg.acerto["1"].modelo)}% dos casos, contra ${p0(nlg.acerto["1"].ordem_fixa)}% da ordem fixa da TED 02; nas três primeiras, ${p0(nlg.acerto["3"].modelo)}% contra ${p0(nlg.acerto["3"].ordem_fixa)}%. Os termos de maior peso do modelo são tecnicamente plausíveis e multilíngues (“battery, akku, accu” para energia; “screen, scherm, écran” para tela). **O que significa.** A ordem fixa já é um bom roteiro básico; o modelo melhora sobretudo a primeira sugestão. Com só ${int(nlg.n_teste)} casos, a diferença precisa ser confirmada.`));

// ================================================================ 7
C.push(H1("7 INTERPRETAÇÃO INTEGRADA DOS RESULTADOS"),
  P("Esta seção conecta a análise exploratória (TED 02) aos resultados da modelagem (TED 03) e ao problema da Inforcell. A pergunta é o que cada achado permite afirmar e o que ele não permite."),
  ...tab("Achados da análise exploratória e o que a modelagem mostrou", ["Achado da TED 02", "O que a TED 03 mostrou", "Conclusão integrada"], [
    ["“Não liga” → 11 causas; ~80% em energia e boot", `O modelo acerta a 1ª causa em ${p0(nlg.acerto["1"].modelo)}% desses casos; a ordem fixa, em ${p0(nlg.acerto["1"].ordem_fixa)}%`, "Ordenar testes por probabilidade funciona; o texto melhora a ordem, mas não elimina a necessidade de testar"],
    ["Aparelho associado ao tipo de falha (V = 0,28)", `Só o aparelho: acurácia de ${p1(BASE_AP.acuracia)}%; incluir o aparelho no modelo de texto eleva o F1-macro de teste de ${f3(SEM_AP.f1_macro)} para ${f3(FIN.f1_macro)}`, "O aparelho é útil como complemento; sozinho, não basta"],
    ["Relatos curtos (mediana de 5 palavras)", `Relatos de 1–3 palavras: ${p0(seg("tamanho", "1–3 palavras").acuracia)}% de acerto; 7+ palavras: ${p0(seg("tamanho", "7+ palavras").acuracia)}%`, "Relatos curtos usam termos padronizados; longos misturam problemas. O assistente deve perguntar de forma estruturada"],
    ["Falhas eletrônicas menos consertadas", "Placa e inicialização estão entre os piores F1", "Onde o técnico mais precisa de ajuda, o texto menos ajuda: o valor do assistente está na condução por testes"],
    ["Vocabulário sobreposto entre energia e boot", "Principal confusão do modelo é energia → inicialização", "A ambiguidade é do relato, não do algoritmo"],
    ["26% de falhas indeterminadas; textos idênticos com rótulos diferentes", `${int(RM.textos_repetidos.conflitantes)} textos repetidos têm rótulos conflitantes`, "Há um teto de desempenho imposto pelos próprios dados"],
    ["Nenhum relato em português", "Sondagem com 8 frases em português: 1 acerto na 1ª causa", "O modelo não pode ser usado na Inforcell sem dados em português"],
  ], [2700, 3200, 3171], { size: 16 }),
  H2("7.1 O que os dados permitem afirmar"),
  B("Que o problema da Inforcell é real e comum: em mais de 22 mil tentativas de reparo em 28 países, o mesmo sintoma leva a causas diferentes (Seção 4.3)."),
  B("Que essas causas não são igualmente prováveis e que o tipo de aparelho e o texto do relato ajudam a ordená-las (Seções 4.2 e 6.2)."),
  B(`Que um modelo simples e interpretável, treinado com ${int(RM.n_treino)} relatos, coloca a causa correta entre as três primeiras em ${p1(MF.top3)}% de casos futuros, em quatro idiomas europeus (Seção 6.2).`),
  H2("7.2 O que os dados não permitem afirmar"),
  B("Que o modelo funciona com relatos em português ou com o perfil de defeitos de Balsas/MA — não há dados para isso."),
  B("Que o assistente reduz o tempo de diagnóstico da Inforcell: essa é a hipótese do projeto de IA e ainda não foi medida."),
  B("Que a probabilidade mostrada pelo modelo é a chance real de acerto (as probabilidades são subestimadas)."),
  B("Que o modelo identifica o componente defeituoso: ele prevê categorias gerais, como “energia/bateria”, e não “conector de carga” ou “CI de carga”."),
  H2("7.3 Contribuição para o MVP"),
  P("As análises contribuem para o MVP de três formas, em graus diferentes de concretização:"),
  ...tab("Contribuições da ciência de dados ao MVP e situação em novembro de 2026", ["Contribuição", "Evidência", "Situação no sistema"], [
    ["Roteiro de testes do “não liga” ordenado por probabilidade (energia → inicialização → placa)", "Seções 4.3 e 6.5", "O assistente atual recebe a instrução de começar pelos testes mais simples (bateria, consumo na fonte) antes dos de placa, o que é coerente com o achado. As probabilidades da TED 02 não estão codificadas no sistema"],
    ["Perguntas complementares quando o relato é pobre", "Seções 4.5 e 6.4", "Proposta: usar o limiar de confiança de 0,35. Não implementado"],
    ["Classificador sugerir_causas() como primeira camada de hipóteses", "Seção 6", "Implementado no repositório de dados (src/inferencia.py, com testes). **Não integrado** ao sistema web"],
    ["Taxonomia de 14 categorias como “causa final” da OS", "Seção 3.1", "Proposta. O campo ainda não existe na OS"],
    ["Campos obrigatórios de marca e modelo no cadastro", "34% de marca desconhecida na base", "Implementado: marca e modelo são obrigatórios no cadastro do equipamento"],
  ], [3000, 1500, 4571], { size: 16 }),
  P("A distância entre o que foi demonstrado e o que está no sistema é o principal ponto em aberto da integração. A Seção 13 propõe os passos para fechá-la."));

// ================================================================ 8
C.push(H1("8 ANÁLISE ÉTICA"),
  P("A análise considera dois momentos: o uso dos dados públicos no projeto acadêmico e o uso das análises e de modelos de IA no atendimento real da Inforcell."),
  H2("8.1 Origem e uso responsável dos dados"),
  P("Os dados foram registrados por voluntários de eventos comunitários e publicados para estimular pesquisa e políticas de reparo. Usá-los para entender falhas e apoiar diagnósticos é compatível com essa finalidade. A licença CC BY-SA 4.0 foi respeitada: a fonte é citada e os dados tratados são publicados sob a mesma licença. Um cuidado ético adicional é não tratar os relatos como verdade técnica: eles refletem a percepção de donos de aparelhos e o julgamento de voluntários, e não um laudo."),
  H2("8.2 Transparência"),
  B("O modelo escolhido é linear: é possível listar os termos que mais pesam em cada categoria (Seção 6.5) e explicar por que uma causa foi sugerida."),
  B("O documento publica as métricas com intervalos de confiança, as linhas de base e as limitações, para que ninguém leia o resultado como “71% de acerto garantido”."),
  B("No uso real, o técnico deve ser informado de que as sugestões são hipóteses geradas por um sistema, de onde vieram e quando a confiança é baixa."),
  H2("8.3 Impactos de resultados incorretos e decisões automatizadas"),
  ...tab("Riscos de erro e cuidados", ["Erro possível", "Quem é afetado", "Consequência", "Cuidado proposto"], [
    ["Causa sugerida errada no topo da lista", "Técnico e cliente", "Testes desnecessários, mais tempo de bancada", "Mostrar três hipóteses, nunca uma só; o técnico confirma por teste"],
    ["Confiar demais na sugestão (viés de automação)", "Cliente", "Troca de peça sem necessidade e orçamento maior", "Exigir teste de confirmação antes do orçamento; registrar a causa final decidida pelo técnico"],
    ["Probabilidade lida como garantia", "Técnico e cliente", "Decisão com falsa segurança", "Não exibir porcentagens; mostrar “confiança baixa/alta”"],
    ["Sugestão sem alerta de risco físico", "Técnico", "Acidente com bateria estufada ou curto", "Alertas de segurança fixos, independentes do modelo"],
    ["Valor de medição ou preço inventado pelo assistente de linguagem", "Técnico e cliente", "Diagnóstico ou orçamento errado", "O assistente é instruído a não inventar valores e a usar só preços já registrados; precisa ser verificado com a IA real"],
  ], [2400, 1400, 2400, 2871], { size: 16 }),
  P("Nenhuma decisão do projeto é totalmente automatizada: o modelo e o assistente sugerem, e o técnico decide. Isso deve continuar assim, porque o desempenho medido não justifica decisões sem revisão humana."),
  H2("8.4 Vieses, riscos e mitigação"),
  ...figura("fig23_representatividade", "Países e idiomas representados na Base A", 600),
  P(`**O que se observa.** ${fmt(R4.pct_paises_top3)}% dos registros vêm de três países (Reino Unido, Alemanha e Bélgica), ${fmt(R4.pct_fonte_restart)}% de uma única organização (The Restart Project), e não há nenhum registro do Brasil nem relato em português. **O que significa.** O modelo aprendeu os aparelhos, as palavras e os hábitos de registro desses lugares. É a principal fonte de viés do projeto.`),
  ...tab("Vieses identificados", ["Viés", "Evidência", "Impacto", "Mitigação (feita / proposta)"], [
    ["Geográfico e linguístico", "0 registros do Brasil; 0% em português", "Desempenho desconhecido na Inforcell; sondagem em português falhou", "Feita: limiar de confiança. Proposta: coletar relatos da Inforcell e avaliar antes de usar"],
    ["Seleção (eventos comunitários)", "Voluntários, aparelhos de uso doméstico", "Perfil de defeitos diferente de uma assistência comercial", "Proposta: comparar a distribuição de causas da Inforcell com a da base"],
    ["Coleta por organização", "V de Cramér organização × desfecho = 0,25", "O desfecho depende de quem registra", "Feita: desfecho e organização fora do modelo"],
    ["Rotulagem", `${fmt(R.pct_indeterminado)}% indeterminados; ${int(RM.textos_repetidos.conflitantes)} textos com rótulos conflitantes`, "Teto de desempenho; erros que são discordância de rótulo", "Feita: exclusão dos indeterminados. Proposta: rotulagem dupla pelo técnico"],
    ["Taxonomia criada pela equipe", "41 rótulos agrupados em 14", "Fronteiras discutíveis (energia × boot)", "Proposta: revisar a taxonomia com o técnico da Inforcell"],
    ["Desbalanceamento", "804 a 18 registros por classe", "Classes raras com resultados instáveis", "Feita: pesos de classe e F1-macro. Proposta: mais dados"],
    ["Temporal", "Dados de 2012 a 2020", "Aparelhos atuais podem ter defeitos diferentes", "Feita: divisão por tempo. Proposta: reavaliar periodicamente"],
    ["Automação (no uso)", "Risco comportamental, não medido", "Técnico pode deixar de testar", "Proposta: três hipóteses, confirmação por teste, decisão do técnico registrada"],
  ], [1700, 2300, 2300, 2771], { size: 15 }),
  H2("8.5 Equidade e impacto social"),
  P("Aparelhos e defeitos sub-representados — desktops, dano por líquido, marcas menos comuns — recebem sugestões piores (Seção 6.3). Na prática, clientes com esses aparelhos poderiam ter diagnósticos mais demorados se o técnico confiasse no sistema. O limiar de confiança e a decisão humana reduzem esse efeito. No lado positivo, apoiar o diagnóstico pode aumentar a taxa de conserto e reduzir o descarte de aparelhos, mas esse benefício é uma hipótese ainda não demonstrada."));

// ================================================================ 9
C.push(H1("9 PRIVACIDADE E LGPD"),
  P("A Lei Geral de Proteção de Dados Pessoais (BRASIL, 2018) se aplica ao tratamento de dados pessoais, isto é, informações relacionadas a pessoa natural identificada ou identificável (art. 5º, I). Por isso, a análise separa os dados usados na disciplina dos dados que o sistema da Inforcell trata em produção. Esta seção não é um parecer jurídico; identifica as questões pertinentes e os cuidados compatíveis com o projeto."),
  H2("9.1 Os dados da análise (base pública)"),
  P("As colunas estruturadas da base não identificam pessoas: não há nome, contato ou documento do dono do aparelho, apenas organização, país, data, aparelho, marca e desfecho. O risco está no **texto livre** do relato, escrito por voluntários. A equipe fez uma varredura automática por padrões de dados pessoais (Tabela 14)."),
  ...tab("Varredura de padrões de dados pessoais nos relatos", ["Padrão procurado", "Base A (reparos)", "Base B (rotulados)", "Base da modelagem"],
    Object.keys(PRIV.reparos_limpo).filter((k) => k !== "textos").map((k) => [k, int(PRIV.reparos_limpo[k]), int(PRIV.falhas_rotuladas_limpo[k]), int(PRIV.falhas_modelagem[k])])
      .concat([["Relatos analisados", int(PRIV.reparos_limpo.textos), int(PRIV.falhas_rotuladas_limpo.textos), int(PRIV.falhas_modelagem.textos)]]),
    [3600, 1800, 1800, 1871], { size: 17, align: [L, RT, RT, RT] }, "Elaborado pela equipe (2026), com src/figuras_ted04.py."),
  P("**O que se observa.** Não foram encontrados e-mails, endereços web nem sequências de dígitos compatíveis com telefone ou documento. Foram encontrados poucos relatos com nomes de pessoas: dois com pronome de tratamento seguido de nome e, em leitura manual de uma amostra, relatos que citam prenomes de voluntários ou donos (por exemplo, “Joe had a mobile phone…”). **O que significa.** O risco é baixo, mas não nulo: a varredura por padrões não encontra todo nome próprio, e um prenome combinado com data, local do evento e aparelho pode, em tese, permitir reconhecer alguém da comunidade local. A anonimização na origem, portanto, não elimina todo risco."),
  P("**Enquadramento.** O tratamento foi feito com fins exclusivamente acadêmicos, hipótese em que a LGPD dispensa parte das obrigações, mas mantém a necessidade de base legal e de cuidados (art. 4º, II, “b”). Os dados foram tornados públicos pela própria organização com licença aberta; ainda assim, o tratamento deve respeitar a finalidade, a boa-fé e o interesse público que justificaram a publicação (art. 7º, § 3º). Os dados também foram coletados fora do Brasil, sob outras leis, como o regulamento europeu de proteção de dados."),
  P("**Cuidados adotados e propostos.** O modelo usa só o texto e o aparelho; país, organização e data não entram nas previsões. Os nomes não têm utilidade para o problema, portanto uma etapa de mascaramento de nomes antes da modelagem é recomendada como melhoria (princípio da necessidade, art. 6º, III). O repositório publica apenas dados que a licença permite e não contém credenciais."),
  H2("9.2 Os dados do sistema em produção"),
  P("Ao contrário da base pública, o sistema web da Inforcell trata dados pessoais de clientes e funcionários. Como as análises deste documento pretendem alimentar esse sistema, os cuidados precisam ser pensados agora."),
  ...tab("Dados pessoais no sistema e avaliação segundo os princípios da LGPD", ["Dado", "Finalidade", "Necessidade", "Situação e cuidado"], [
    ["Nome e telefone do cliente", "Identificar o dono do aparelho e avisar sobre o orçamento", "Necessários", "Obrigatórios no cadastro. Base legal provável: execução de contrato (art. 7º, V)"],
    ["CPF do cliente", "Comprovante e identificação", "Nem sempre necessário", "Opcional no cadastro, o que atende ao princípio da necessidade"],
    ["E-mail e endereço", "Contato e entrega", "Opcionais", "Opcionais no cadastro"],
    ["Relato do defeito e histórico da OS", "Diagnóstico e base para casos futuros", "Necessários", "Pode conter dados pessoais digitados pelo atendente; ver Seção 9.3"],
    ["Fotos enviadas ao assistente", "Análise visual do defeito", "Úteis", "Não são guardadas, mas podem mostrar dados na tela do aparelho (mensagens, fotos)"],
    ["Login e senha dos funcionários", "Acesso por perfil", "Necessários", "Senhas guardadas com hash (bcrypt); limite de tentativas de login"],
  ], [1900, 2300, 1400, 3471], { size: 15 }),
  H2("9.3 Envio de dados ao provedor de IA"),
  P("O assistente envia o relato, as medições e, se houver, a foto a um provedor de IA no exterior, o que configura transferência internacional de dados quando houver dado pessoal (art. 33). Para reduzir esse risco, o sistema remove do texto, antes do envio, padrões de e-mail, CPF e telefone, e insere o nome do cliente só no navegador, na mensagem final. Essas medidas foram verificadas no código, mas têm limites: **nomes próprios digitados no relato não são removidos**, e as fotos não passam por nenhum filtro. Além disso, essas proteções ainda não foram testadas com o provedor real."),
  H2("9.4 Segurança e direitos dos titulares"),
  B("**Medidas existentes:** isolamento dos dados por empresa em todas as consultas; senhas com hash; chave da IA de cada empresa criptografada (AES-256-GCM) e nunca devolvida inteira ao navegador; acesso por perfil."),
  B("**Lacunas:** não há prazo de retenção para Ordens de Serviço e conversas com o assistente, nem função para atender pedidos do titular, como acesso, correção e eliminação (art. 18). Também não há rotina definida para comunicar incidentes (art. 48)."),
  B("**Demonstrações:** apresentações e testes devem usar apenas clientes fictícios; o banco de demonstração não deve conter dados reais."),
  B("**Decisões automatizadas:** como o técnico decide, o sistema não toma decisões unicamente automatizadas; ainda assim, o cliente deve poder pedir explicação do diagnóstico (art. 20)."),
  nota("Síntese", "Para os dados públicos da análise, o risco de privacidade é baixo e o principal cuidado é não usar nem republicar nomes que aparecem no texto livre. Para o sistema em produção, o risco é real e as prioridades são: definir finalidade e retenção, ampliar a remoção de dados pessoais antes do envio à IA (incluindo nomes), informar o cliente sobre o uso de IA e criar o atendimento aos direitos do titular."));

// ================================================================ 10
C.push(H1("10 LIMITAÇÕES DOS DADOS, DAS ANÁLISES E DOS MODELOS"),
  H2("10.1 Dos dados"),
  B("**Representatividade:** nenhum registro do Brasil e nenhum relato em português; 55% dos registros vêm de três países (Seção 8.4)."),
  B("**Contexto:** eventos comunitários de voluntários, não assistência técnica comercial."),
  B("**Completude:** idade ausente em 69% dos registros; marca desconhecida em 34%; 26% das falhas indeterminadas."),
  B("**Granularidade:** a falha é uma categoria geral; não há o componente, as medições nem o tempo de diagnóstico, que seriam necessários para avaliar diretamente a hipótese do projeto."),
  B("**Sem imagens:** a hipótese de diagnóstico por foto não pode ser avaliada com esta base."),
  H2("10.2 Das análises"),
  B("Associações não provam causa: a relação entre aparelho e falha, por exemplo, pode refletir os hábitos de cada organização."),
  B("O detector de idioma e a taxonomia foram criados pela equipe; erros neles se propagam para as análises."),
  B("A busca de padrões de dados pessoais é indicativa e não exaustiva."),
  H2("10.3 Dos modelos"),
  B(`**Sobreajuste:** acerto de 96,7% nos dados de treino contra ${p1(MF.acuracia)}% no teste; a curva de aprendizagem ainda sobe, ou seja, faltam dados.`),
  B("**Classes raras e residuais:** dano por líquido, carcaça e “outros” têm resultados pouco confiáveis."),
  B("**Calibração:** probabilidades subestimadas (ECE = " + f3(RM.calibracao.ece) + ")."),
  B("**Idioma:** desempenho menor em alemão que em inglês e falha em português (sondagem de 8 frases, que não é métrica)."),
  B("**Uso do teste:** o teste não foi usado para escolher o modelo, mas as métricas de todos os candidatos foram examinadas; a comparação entre candidatos é indicativa."),
  B("**Validação no contexto real:** nenhum resultado foi medido com técnicos ou relatos da Inforcell."));

// ================================================================ 11
C.push(H1("11 DATA STORYTELLING"),
  P("Esta seção organiza os resultados como uma narrativa, seguindo a sequência Problema → Dados → Análise → Resultados → Interpretação → Responsabilidade → Solução → Próximos passos. A apresentação final (presentation/apresentacao-final.pdf) segue a mesma ordem, com um gráfico por mensagem."),
  P("**A história em uma frase:** “Um celular que não liga pode ter onze causas, mas não onze causas igualmente prováveis — e os dados de 22 mil reparos mostram em que ordem testar, desde que o técnico continue no comando e que os dados venham, um dia, da própria Inforcell.”", { noIndent: true }),
  ...tab("Roteiro da narrativa", ["Passo", "Mensagem", "Evidência visual", "Número que fica"], [
    ["1. Problema", "O mesmo sintoma tem várias causas e o técnico depende da própria experiência", "Relato da Inforcell", "—"],
    ["2. Dados", "22 mil reparos públicos em 28 países; 4 mil com a causa rotulada", "Figura 1 (funil)", `${int(R.n_reparos)} e ${int(RM.n_total)}`],
    ["3. Análise", "“Não liga” se divide em 11 causas, mas 80% são energia ou boot", "Figura 3", "80%"],
    ["4. Resultados", "O modelo põe a causa certa entre as três primeiras em quase 9 de 10 casos", "Figura 7 (top-k)", `${p1(MF.top3)}% contra ${p1(BASE_AP.top3)}%`],
    ["5. Interpretação", "Serve para ordenar hipóteses; onde o texto é ambíguo, o teste decide", "Figura 9 (confusões)", "energia → boot"],
    ["6. Responsabilidade", "Nenhum dado do Brasil; nomes no texto livre; risco de confiar demais", "Figura 12 (representatividade)", "0% em português"],
    ["7. Solução", "Três hipóteses + limiar de confiança + decisão do técnico", "Fluxo do assistente", "limiar 0,35"],
    ["8. Próximos passos", "Integrar ao sistema e coletar relatos em português da Inforcell", "Linha do tempo", "—"],
  ], [1500, 3700, 2100, 1771], { size: 16 }),
  H2("11.1 Escolhas de comunicação"),
  B("**Uma mensagem por gráfico:** cada figura da apresentação tem título que afirma o achado (por exemplo, “80% dos casos de ‘não liga’ são energia ou boot”), e não só o nome das variáveis."),
  B("**Comparação sempre com uma referência:** os resultados do modelo aparecem ao lado da linha de base, para que 71,6% tenha significado."),
  B("**Incerteza visível:** os intervalos de confiança e o tamanho das amostras são citados quando o número é pequeno (por exemplo, os 142 casos de “não liga” no teste)."),
  B("**Limitações na mesma narrativa:** a parte de responsabilidade não fica no fim como apêndice; ela muda a recomendação final (usar como ordenador, não como diagnóstico)."));

// ================================================================ 12
C.push(H1("12 CONCLUSÕES"),
  P("O projeto partiu de uma dificuldade relatada pela Inforcell e a transformou em perguntas que os dados podiam responder. A Tabela 17 resume as respostas e as evidências que sustentam cada uma."),
  ...tab("Conclusões e evidências", ["Pergunta inicial", "Resposta", "Evidência", "Grau de confiança"], [
    ["O mesmo sintoma leva a várias causas?", "Sim", "“Não liga” → 11 categorias (Seção 4.3)", "Alto, nos dados públicos"],
    ["As causas têm probabilidades diferentes?", "Sim; energia e boot somam ~80% no “não liga”", "Seções 4.3 e 6.5", "Alto, nos dados públicos"],
    ["O que ajuda a prever a causa?", "O texto do relato, complementado pelo tipo de aparelho", "Seção 6.1 e Tabela 8", "Alto"],
    ["É possível sugerir causas com utilidade?", `Sim, como lista: top-3 de ${p1(MF.top3)}%`, `IC 95% de ${p1(ICF.top3[0])}% a ${p1(ICF.top3[1])}%`, "Alto para os idiomas da base; desconhecido para o português"],
    ["O modelo pode dar o diagnóstico final?", "Não", `Acerto da 1ª causa de ${p1(MF.acuracia)}%; erros em placa e boot`, "Alto"],
    ["Funciona na Inforcell?", "Ainda não se sabe", "Nenhum dado brasileiro; sondagem em português falhou", "Questão em aberto"],
    ["O assistente reduz o tempo de diagnóstico?", "Ainda não se sabe", "Depende do piloto do projeto de IA", "Questão em aberto"],
  ], [2400, 2300, 2500, 1871], { size: 16 }),
  P("Em síntese, a ciência de dados confirmou o problema, indicou a ordem dos testes e produziu um modelo útil para sugerir hipóteses, com limitações claras de idioma, contexto e granularidade. A principal lição do semestre é que a contribuição dos dados para o MVP não é substituir o técnico, e sim organizar a experiência acumulada em muitos reparos para que o técnico comece pelos testes mais prováveis. Para que essa contribuição chegue de fato à Inforcell, falta integrar o modelo ao sistema e, sobretudo, reunir dados próprios em português."),
  P("**Questões que permanecem em aberto:** o desempenho com relatos reais da Inforcell; o efeito do assistente no tempo e no número de testes; a calibração das probabilidades; e a eficácia das proteções de privacidade com o provedor de IA real."));

// ================================================================ 13
C.push(H1("13 POSSIBILIDADES DE EVOLUÇÃO DA SOLUÇÃO"),
  ...tab("Evolução proposta, em ordem de prioridade", ["Prioridade", "Ação", "Por quê", "Depende de"], [
    ["1", "Criar na OS o campo “causa final” com a taxonomia de 14 categorias e registrar o resultado de cada teste", "Gera dados rotulados em português, a principal lacuna do projeto", "Sistema web"],
    ["2", "Coletar e rotular uma amostra de relatos reais da Inforcell e avaliar o modelo nela", "Responde à questão em aberto mais importante", "Item 1 e consentimento da Inforcell"],
    ["3", "Integrar sugerir_causas() ao assistente como primeira camada de hipóteses, com o limiar de 0,35", "Transforma o resultado da TED 03 em funcionalidade", "Serviço Python ao lado da API Node.js, ou reimplementação do modelo"],
    ["4", "Mascarar nomes no texto antes da modelagem e do envio à IA; definir retenção e direitos do titular", "Reduz riscos identificados na Seção 9", "Sistema web"],
    ["5", "Calibrar as probabilidades (isotônica ou sigmoide)", "Permite ler a confiança de forma mais fiel", "Mais dados de validação"],
    ["6", "Testar representações multilíngues (embeddings de frases) e tradução automática", "Pode melhorar alemão, holandês e português", "Mais dados e infraestrutura"],
    ["7", "Revisar a taxonomia com o técnico (separar energia e boot com critérios de bancada)", "Reduz a principal confusão do modelo", "Item 2"],
    ["8", "Medir no piloto o número de testes até a causa e o tempo de diagnóstico", "Valida a hipótese do Lean Canvas", "Projeto de IA"],
  ], [1000, 3500, 2700, 1871], { size: 16 }),
  P("Novos dados também permitiriam avaliar a hipótese de diagnóstico por imagem, registrada no canvas do projeto, que esta base não permitiu testar."));

// ================================================================ referências e apêndices
C.push(H1("REFERÊNCIAS"),
  ...[
    "ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. **NBR 14724**: informação e documentação – trabalhos acadêmicos – apresentação. Rio de Janeiro: ABNT, 2024.",
    "BRASIL. Lei nº 13.709, de 14 de agosto de 2018. Lei Geral de Proteção de Dados Pessoais (LGPD). **Diário Oficial da União**, Brasília, DF, 15 ago. 2018.",
    "EFRON, Bradley; TIBSHIRANI, Robert J. **An introduction to the bootstrap**. New York: Chapman & Hall, 1993.",
    "GEBRU, Timnit et al. Datasheets for datasets. **Communications of the ACM**, v. 64, n. 12, p. 86-92, 2021.",
    "GUO, Chuan et al. On calibration of modern neural networks. In: INTERNATIONAL CONFERENCE ON MACHINE LEARNING, 34., 2017, Sydney. **Proceedings** [...]. PMLR, v. 70, p. 1321-1330, 2017.",
    "KNAFLIC, Cole Nussbaumer. **Storytelling with data**: a data visualization guide for business professionals. Hoboken: Wiley, 2015.",
    "MEHRABI, Ninareh et al. A survey on bias and fairness in machine learning. **ACM Computing Surveys**, v. 54, n. 6, p. 1-35, 2021.",
    "MITCHELL, Margaret et al. Model cards for model reporting. In: CONFERENCE ON FAIRNESS, ACCOUNTABILITY, AND TRANSPARENCY, 2019, Atlanta. **Proceedings** [...]. New York: ACM, 2019. p. 220-229.",
    "OPEN REPAIR ALLIANCE. **Open Repair Data**. Disponível em: https://openrepair.org/open-data/. Acesso em: 28 set. 2026.",
    "OPEN REPAIR ALLIANCE. **Repositório de dados (ORDS v0.3, exportação 202507, e quests MobiFix, CompCat e TabiCat)**. Disponível em: https://github.com/openrepair/data. Acesso em: 28 set. 2026.",
    "PEDREGOSA, Fabian et al. Scikit-learn: machine learning in Python. **Journal of Machine Learning Research**, v. 12, p. 2825-2830, 2011.",
    "SOKOLOVA, Marina; LAPALME, Guy. A systematic analysis of performance measures for classification tasks. **Information Processing & Management**, v. 45, n. 4, p. 427-437, 2009.",
    "TUKEY, John W. **Exploratory data analysis**. Reading: Addison-Wesley, 1977.",
  ].map((t) => new Paragraph({ children: runs(t), alignment: AlignmentType.LEFT, spacing: { line: 240, after: 200 } })));

C.push(H1("APÊNDICE A – ORGANIZAÇÃO DO REPOSITÓRIO"),
  P("O repositório das TEDs anteriores foi mantido; a TED 04 acrescentou o documento final, a apresentação e o script das figuras novas.", { noIndent: true }),
  ...[
    "projeto-inforcell/",
    "├── README.md                      versão final: visão geral, resultados e como reproduzir",
    "├── docs/",
    "│   ├── TED01.pdf, TED02.pdf, TED03.pdf   versões 1.0 a 3.0 (preservadas)",
    "│   ├── documento-final.pdf / .docx       este documento (versão final)",
    "│   └── build/                     scripts que geram os documentos",
    "├── presentation/",
    "│   └── apresentacao-final.pdf / .pptx    apresentação de Data Storytelling",
    "├── data/raw/ e data/processed/    dados brutos sem alteração e bases tratadas",
    "├── notebooks/TED01, TED02, TED03",
    "├── src/                           preparacao, graficos, modelagem, inferencia, figuras_ted04",
    "├── models/                        classificadores treinados",
    "├── tests/                         24 testes automatizados",
    "└── reports/                       figures/ (fig01…fig23) e resultados em JSON",
  ].map((l) => new Paragraph({ spacing: { line: 240, after: 0 }, children: [new TextRun({ text: l, font: "Courier New", size: 18 })] })),
  sp(), P("**Para reproduzir:** pip install -r requirements.txt; executar os notebooks de notebooks/TED02 e notebooks/TED03; python src/figuras_ted04.py; pytest -q. Para gerar este documento: node docs/build/gerar_documento_final.js (pacote npm docx).", { noIndent: true }));

C.push(H1("APÊNDICE B – ONDE ESTÁ CADA CONTEÚDO DAS VERSÕES ANTERIORES"),
  ...tab("Rastreabilidade entre as versões", ["Conteúdo", "Versão final", "Detalhe completo"], [
    ["Necessidades de dados e seleção da base", "Seção 2", "v1.0, Seções 3 e 4"],
    ["Revisão da seleção e duas bases", "Seção 2.2", "v2.0, Seção 6"],
    ["Problemas de qualidade e tratamentos", "Seção 3", "v2.0, Seções 7 a 10"],
    ["Dicionário de dados", "—", "v2.0, Seção 10, e data/processed/README.md"],
    ["Análise exploratória completa (12 figuras)", "Seção 4 (4 figuras)", "v2.0, Seções 11 a 13"],
    ["Problema de ML, variáveis e transformações", "Seções 3.2 e 5", "v3.0, Seções 15 e 16"],
    ["Técnicas, treinamento e métricas", "Seção 5", "v3.0, Seção 17"],
    ["Resultados e interpretação do modelo", "Seções 6 e 7", "v3.0, Seções 18 e 19"],
    ["Limitações", "Seção 10", "v2.0, Seção 4.7, e v3.0, Seção 20"],
    ["Relação com o MVP", "Seção 7.3 (corrigida)", "v3.0, Seção 21"],
    ["Registro de decisões e experimentos", "—", "v3.0, Seção 22"],
  ], [3700, 2400, 2971], { size: 17 }));

// ---------------------------------------------------------------- documento
const doc = new Document({
  creator: "Equipe Inforcell – UNIBALSAS", title: "Documento Técnico – Versão Final (TED 04)", description: "Integração dos resultados, ética, LGPD e Data Storytelling – Inforcell",
  styles: {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, font: FONT }, paragraph: { spacing: { before: 360, after: 240 }, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: false, font: FONT }, paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1, keepNext: true } },
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
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(path.join(ROOT, "docs", "documento-final.docx"), buf); console.log("docx ok", nFig, "figuras", nTab, "tabelas"); });
