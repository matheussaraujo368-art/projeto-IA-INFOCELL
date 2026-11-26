// Gera presentation/apresentacao-final.pptx — apresentação de Data Storytelling da TED 04.
// Sequência: Problema → Dados → Análise → Resultados → Interpretação → Responsabilidade → Solução → Próximos passos.
// Uso: node docs/build/gerar_apresentacao.js   (requer o pacote npm "pptxgenjs")
// PDF: soffice --headless --convert-to pdf --outdir presentation presentation/apresentacao-final.pptx
const fs = require("fs");
const path = require("path");
const PptxGenJS = require("pptxgenjs");

const ROOT = path.resolve(__dirname, "..", "..");
const FIG = path.join(ROOT, "reports", "figures");
const RM = JSON.parse(fs.readFileSync(path.join(ROOT, "reports", "resultados_modelagem.json"), "utf8"));
const R = JSON.parse(fs.readFileSync(path.join(ROOT, "reports", "resultados_eda.json"), "utf8"));
const pct = (x, d = 1) => (x * 100).toLocaleString("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d }) + "%";
const MF = RM.final.metricas;
const BASE = RM.tabela_modelos.find((t) => t.modelo === "Base · prior por aparelho");

const AZUL = "2A78D6", ESCURO = "0B0B0B", CINZA = "52514E", CLARO = "EEF4FB", LARANJA = "EB6834", VERDE = "1BAF7A";
const FONTE = "Arial";
const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE"; // 13,33 × 7,5 pol
pptx.title = "Do sintoma à causa provável – TED 04 Big Data";

const PASSOS = ["Problema", "Dados", "Análise", "Resultados", "Interpretação", "Responsabilidade", "Solução", "Próximos passos"];
let nSlide = 0;

function base(passo, titulo) {
  const s = pptx.addSlide();
  s.background = { color: "FFFFFF" };
  nSlide++;
  if (passo !== null) {
    // trilha da narrativa
    PASSOS.forEach((p, i) => {
      const ativo = i === passo;
      s.addText(p, { x: 0.45 + i * 1.55, y: 0.18, w: 1.5, h: 0.3, fontFace: FONTE, fontSize: 9, bold: ativo,
        color: ativo ? "FFFFFF" : CINZA, fill: { color: ativo ? AZUL : "F2F2F0" }, align: "center", valign: "middle" });
    });
  }
  if (titulo) s.addText(titulo, { x: 0.45, y: 0.65, w: 12.4, h: 0.85, fontFace: FONTE, fontSize: 22, bold: true, color: ESCURO, valign: "top" });
  s.addText(`Inforcell · Big Data e Ciência de Dados · TED 04`, { x: 0.45, y: 7.05, w: 8, h: 0.3, fontFace: FONTE, fontSize: 9, color: CINZA });
  s.addText(String(nSlide), { x: 12.3, y: 7.05, w: 0.6, h: 0.3, fontFace: FONTE, fontSize: 9, color: CINZA, align: "right" });
  return s;
}
function imagem(s, nome, x, y, wMax, hMax) {
  const b = fs.readFileSync(path.join(FIG, nome + ".png"));
  const ar = b.readUInt32BE(16) / b.readUInt32BE(20);
  let w = wMax, h = w / ar;
  if (h > hMax) { h = hMax; w = h * ar; }
  s.addImage({ path: path.join(FIG, nome + ".png"), x: x + (wMax - w) / 2, y, w, h });
}
function texto(s, linhas, x, y, w, h, o = {}) {
  s.addText(linhas.map((l) => ({ text: l, options: { bullet: o.bullet === false ? false : { indent: 14 }, breakLine: true, paraSpaceAfter: 8 } })),
    { x, y, w, h, fontFace: FONTE, fontSize: o.size || 16, color: ESCURO, valign: "top", ...o.extra });
}
function cartao(s, valor, rotulo, x, y, cor = AZUL, w = 2.9) {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h: 1.35, fill: { color: CLARO }, line: { color: CLARO } });
  s.addShape(pptx.ShapeType.rect, { x, y, w: 0.08, h: 1.35, fill: { color: cor }, line: { color: cor } });
  s.addText(valor, { x: x + 0.2, y: y + 0.1, w: w - 0.3, h: 0.7, fontFace: FONTE, fontSize: 30, bold: true, color: cor });
  s.addText(rotulo, { x: x + 0.2, y: y + 0.78, w: w - 0.3, h: 0.5, fontFace: FONTE, fontSize: 12, color: CINZA });
}
function conclusao(s, txt, y = 6.35) {
  s.addText(txt, { x: 0.45, y, w: 12.4, h: 0.55, fontFace: FONTE, fontSize: 15, bold: true, color: AZUL, fill: { color: CLARO }, margin: 8 });
}

// 1 — título
{
  const s = base(null, null);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.25, h: 7.5, fill: { color: AZUL }, line: { color: AZUL } });
  s.addText("Do sintoma à causa provável", { x: 0.8, y: 1.6, w: 11.5, h: 1, fontFace: FONTE, fontSize: 40, bold: true, color: ESCURO });
  s.addText("O que 22 mil reparos ensinam ao assistente de diagnóstico da Inforcell — e com que cuidados", { x: 0.8, y: 2.6, w: 11.5, h: 0.9, fontFace: FONTE, fontSize: 20, color: CINZA });
  s.addText("Big Data e Ciência de Dados · TED 04 – Integração, ética, LGPD e Data Storytelling", { x: 0.8, y: 4.1, w: 11.5, h: 0.4, fontFace: FONTE, fontSize: 14, color: AZUL, bold: true });
  s.addText("Rafael R Pontara · Matheus de Brito Araujo · Pedro Paulo T. Araujo\nUNIBALSAS – Sistemas de Informação · Balsas/MA · 2026", { x: 0.8, y: 4.7, w: 11.5, h: 0.9, fontFace: FONTE, fontSize: 14, color: ESCURO });
}

// 2 — problema
{
  const s = base(0, "“Não liga” não é um diagnóstico: é o começo de uma investigação");
  texto(s, [
    "Na Inforcell, o cliente descreve o defeito em poucas palavras.",
    "O mesmo sintoma pode ser bateria, conector, circuito de carga, placa ou software.",
    "Hoje, a ordem dos testes depende da experiência de cada técnico.",
    "Testar na ordem errada custa tempo de bancada — e pode levar à troca de peça sem necessidade.",
  ], 0.45, 1.7, 7.2, 4.3, { size: 18 });
  s.addShape(pptx.ShapeType.rect, { x: 8.1, y: 1.8, w: 4.75, h: 3.9, fill: { color: CLARO }, line: { color: CLARO } });
  s.addText([
    { text: "A pergunta deste trabalho", options: { bold: true, color: AZUL, breakLine: true, fontSize: 16 } },
    { text: "Os dados conseguem dizer ao técnico por onde começar?", options: { fontSize: 22, bold: true, color: ESCURO } },
  ], { x: 8.35, y: 2.1, w: 4.3, h: 3.3, fontFace: FONTE, valign: "top" });
}

// 3 — dados
{
  const s = base(1, `${R.n_reparos.toLocaleString("pt-BR")} reparos públicos; ${RM.n_total.toLocaleString("pt-BR")} com a causa rotulada`);
  imagem(s, "fig22_funil_dados", 0.45, 1.55, 8.3, 4.7);
  texto(s, [
    "Open Repair Data: eventos comunitários de reparo, 28 países, 2012–2025.",
    "Base A: contexto do reparo (aparelho, marca, desfecho).",
    "Base B: o relato e a causa atribuída por voluntários.",
    "Licença aberta CC BY-SA 4.0.",
  ], 9.0, 1.7, 3.9, 4.5, { size: 14 });
  conclusao(s, "O par “relato → causa” existe, mas só 4 mil registros servem para modelar.");
}

// 4 — análise: não liga
{
  const s = base(2, "Onze causas possíveis para “não liga” — mas 80% são energia ou boot");
  imagem(s, "fig11_nao_liga", 0.45, 1.55, 7.6, 4.7);
  texto(s, [
    `${R.nao_liga.n} relatos de aparelhos que não ligam.`,
    `Energia/bateria: ${Math.round(R.nao_liga.distribuicao["Energia / bateria"])}%.`,
    `Inicialização (boot): ${Math.round(R.nao_liga.distribuicao["Inicialização (boot)"])}%.`,
    "O problema da Inforcell aparece nos dados — e já sugere uma ordem de testes.",
  ], 8.4, 1.8, 4.5, 4.3, { size: 16 });
  conclusao(s, "Testar primeiro alimentação, depois inicialização, depois placa.");
}

// 5 — análise: aparelho e falhas difíceis
{
  const s = base(2, "O aparelho define o ponto de partida; as falhas eletrônicas são as mais difíceis");
  imagem(s, "fig09_falhas_aparelho", 0.45, 1.55, 6.1, 4.6);
  imagem(s, "fig10_conserto_falha", 6.75, 1.55, 6.1, 4.6);
  conclusao(s, "Celular começa por tela e bateria; placa, energia e boot são consertados em só 33–39% dos casos.");
}

// 6 — resultados
{
  const s = base(3, "A causa certa está entre as três primeiras sugestões em 9 de 10 casos");
  imagem(s, "fig17_topk", 0.45, 1.55, 6.6, 4.6);
  cartao(s, pct(MF.top3), "causa certa entre as 3 primeiras (teste 2019–2020)", 7.4, 1.7, AZUL, 5.4);
  cartao(s, pct(BASE.top3), "o mesmo indicador usando só o tipo de aparelho", 7.4, 3.2, CINZA, 5.4);
  cartao(s, pct(MF.acuracia), "acerto da 1ª sugestão · F1-macro " + MF.f1_macro.toLocaleString("pt-BR", { maximumFractionDigits: 3 }), 7.4, 4.7, LARANJA, 5.4);
  conclusao(s, "Regressão logística sobre o texto e o aparelho: simples, interpretável e muito acima da linha de base.");
}

// 7 — interpretação: confusões
{
  const s = base(4, "Onde o relato é ambíguo, o teste de bancada decide");
  imagem(s, "fig16_matriz_confusao", 0.45, 1.5, 6.6, 4.75);
  texto(s, [
    "Erros entre causas parecidas: software ↔ desempenho, energia → boot.",
    "“Won’t turn on” descreve tanto bateria quanto inicialização.",
    "Placa e boot — as falhas mais difíceis — são onde o texto menos ajuda.",
    `${RM.textos_repetidos.conflitantes} textos idênticos têm rótulos diferentes: há um teto nos próprios dados.`,
  ], 7.4, 1.7, 5.5, 4.5, { size: 15 });
  conclusao(s, "O modelo serve para ordenar hipóteses, não para dar o diagnóstico final.");
}

// 8 — interpretação: confiança
{
  const s = base(4, "Quando a confiança é baixa, o assistente deve perguntar antes de sugerir");
  imagem(s, "fig20_confianca", 0.45, 1.55, 8.2, 4.6);
  const sel = (c) => RM.selecao_confianca.find((x) => Math.abs(x.corte - c) < 1e-6);
  texto(s, [
    `Confiança ≥ 0,35: sugere em ${pct(sel(0.35).cobertura, 0)} dos casos, com ${pct(sel(0.35).acuracia, 0)} de acerto.`,
    "Abaixo disso: perguntas complementares (quando começou, caiu, molhou, carrega?).",
    "As probabilidades são subestimadas: não mostrar como “chance de acerto”.",
  ], 8.9, 1.7, 4.0, 4.5, { size: 14 });
  conclusao(s, "A confiança é útil para decidir o próximo passo, não para impressionar o cliente.");
}

// 9 — responsabilidade: representatividade
{
  const s = base(5, "Nenhum registro do Brasil, nenhum relato em português");
  imagem(s, "fig23_representatividade", 0.45, 1.55, 12.4, 4.55);
  conclusao(s, "Em 8 frases em português, o modelo acertou a 1ª causa só uma vez. Antes de usar na Inforcell, é preciso dado próprio.");
}

// 10 — responsabilidade: ética e LGPD
{
  const s = base(5, "Ética e LGPD: o risco muda quando o dado sai da base pública e entra na loja");
  s.addText("Base pública (análise)", { x: 0.45, y: 1.6, w: 6, h: 0.4, fontFace: FONTE, fontSize: 17, bold: true, color: AZUL });
  texto(s, [
    "Colunas sem nome, contato ou documento.",
    "Varredura: 0 e-mails, URLs ou números de documento; poucos nomes no texto livre.",
    "Risco baixo, não nulo: nome + data + local podem identificar alguém.",
    "Vieses: geográfico, de seleção, de rotulagem e de coleta.",
  ], 0.45, 2.1, 6, 3.9, { size: 14 });
  s.addText("Sistema da Inforcell (produção)", { x: 6.85, y: 1.6, w: 6, h: 0.4, fontFace: FONTE, fontSize: 17, bold: true, color: LARANJA });
  texto(s, [
    "Nome, telefone e CPF (opcional) de clientes: finalidade e necessidade definidas.",
    "Texto enviado à IA sem e-mail, CPF e telefone — mas nomes e fotos não são filtrados.",
    "Faltam: prazo de retenção, atendimento aos direitos do titular e aviso sobre o uso de IA.",
    "O técnico decide; a IA só sugere (viés de automação).",
  ], 6.85, 2.1, 6, 3.9, { size: 14 });
  conclusao(s, "Anonimizar não basta: é preciso avaliar finalidade, necessidade, segurança e quem decide.");
}

// 11 — solução
{
  const s = base(6, "Como os dados entram no assistente: três hipóteses, um limiar e o técnico no comando");
  const etapas = [
    ["Relato + aparelho", "abertos na OS"],
    ["Modelo", "3 causas mais prováveis"],
    ["Confiança ≥ 0,35?", "sim: testar na ordem\nnão: perguntar antes"],
    ["Teste de bancada", "um por vez, o técnico interpreta"],
    ["Causa final", "decidida e registrada pelo técnico"],
  ];
  etapas.forEach(([t, d], i) => {
    const x = 0.45 + i * 2.5;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 2.0, w: 2.2, h: 1.9, fill: { color: i === 4 ? AZUL : CLARO }, line: { color: AZUL, width: 1 }, rectRadius: 0.12 });
    s.addText([{ text: t, options: { bold: true, fontSize: 15, breakLine: true, color: i === 4 ? "FFFFFF" : ESCURO } }, { text: d, options: { fontSize: 12, color: i === 4 ? "FFFFFF" : CINZA } }],
      { x: x + 0.1, y: 2.1, w: 2.0, h: 1.7, fontFace: FONTE, align: "center", valign: "middle" });
    if (i < etapas.length - 1) s.addShape(pptx.ShapeType.rightArrow, { x: x + 2.23, y: 2.8, w: 0.25, h: 0.3, fill: { color: AZUL }, line: { color: AZUL } });
  });
  s.addText("A causa final volta como dado rotulado em português e melhora o modelo", { x: 0.45, y: 4.2, w: 12.4, h: 0.4, fontFace: FONTE, fontSize: 14, italic: true, color: CINZA, align: "center" });
  s.addShape(pptx.ShapeType.rect, { x: 0.45, y: 4.85, w: 12.4, h: 1.25, fill: { color: "FFF4E5" }, line: { color: "EDA100" } });
  s.addText([
    { text: "Situação hoje: ", options: { bold: true } },
    { text: "o modelo e a função sugerir_causas() estão prontos e testados no repositório de dados, mas ainda não estão ligados ao sistema web, e o campo “causa final” ainda não existe na OS. O assistente atual usa um modelo de linguagem com consulta aos dados da loja." },
  ], { x: 0.65, y: 4.95, w: 12.0, h: 1.05, fontFace: FONTE, fontSize: 13, color: ESCURO, valign: "middle" });
  conclusao(s, "Os dados não substituem o técnico: organizam a experiência de milhares de reparos para ele começar pelo mais provável.");
}

// 12 — próximos passos
{
  const s = base(7, "Próximos passos: do dado público ao dado da Inforcell");
  const itens = [
    ["1", "Campo “causa final” na OS", "gera dados rotulados em português"],
    ["2", "Avaliar com relatos reais da Inforcell", "responde se o modelo serve na loja"],
    ["3", "Integrar sugerir_causas() ao assistente", "com o limiar de 0,35"],
    ["4", "Mascarar nomes e definir retenção", "privacidade antes de escalar"],
    ["5", "Medir no piloto: testes até a causa e tempo", "valida a hipótese do projeto"],
  ];
  itens.forEach(([n, t, d], i) => {
    const y = 1.65 + i * 0.92;
    s.addShape(pptx.ShapeType.ellipse, { x: 0.5, y, w: 0.6, h: 0.6, fill: { color: AZUL }, line: { color: AZUL } });
    s.addText(n, { x: 0.5, y, w: 0.6, h: 0.6, fontFace: FONTE, fontSize: 18, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    s.addText([{ text: t, options: { bold: true, color: ESCURO } }, { text: `  —  ${d}`, options: { color: CINZA } }], { x: 1.3, y, w: 11.5, h: 0.6, fontFace: FONTE, fontSize: 18, valign: "middle" });
  });
}

// 13 — o que fica
{
  const s = base(null, "O que fica");
  const blocos = [
    ["Confirmado pelos dados", VERDE, ["O mesmo sintoma tem várias causas, com probabilidades desiguais.", "Texto + aparelho ordenam bem as hipóteses (top-3 de " + pct(MF.top3) + ")."]],
    ["Não demonstrado", LARANJA, ["Que funciona com relatos em português.", "Que reduz o tempo de diagnóstico na Inforcell."]],
    ["Nossa responsabilidade", AZUL, ["Sugerir, não decidir.", "Coletar dados próprios com finalidade, necessidade e segurança."]],
  ];
  blocos.forEach(([t, cor, linhas], i) => {
    const x = 0.45 + i * 4.2;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.7, w: 3.95, h: 4.3, fill: { color: "FAFAF8" }, line: { color: cor, width: 2 } });
    s.addText(t, { x: x + 0.2, y: 1.85, w: 3.6, h: 0.5, fontFace: FONTE, fontSize: 18, bold: true, color: cor });
    texto(s, linhas, x + 0.2, 2.5, 3.6, 3.4, { size: 16 });
  });
  s.addText("Documento completo: docs/documento-final.pdf · github.com/matheussaraujo368-art/projeto-IA-INFOCELL", { x: 0.45, y: 6.35, w: 12.4, h: 0.4, fontFace: FONTE, fontSize: 12, color: CINZA, align: "center" });
}

fs.mkdirSync(path.join(ROOT, "presentation"), { recursive: true });
pptx.writeFile({ fileName: path.join(ROOT, "presentation", "apresentacao-final.pptx") }).then((f) => console.log("pptx ok", nSlide, "slides", f));
