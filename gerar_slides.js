// Gera a apresentação do Trabalho 07 — NBR 16636-1 (Etapas Iniciais do Projeto)
// Conceito: o projeto como uma corrida de revezamento; estudo de caso: TCC Arena Poty
// Uso: npm install --no-save pptxgenjs react-icons react react-dom sharp && node gerar_slides.js
const pptxgen = require("pptxgenjs");
const path = require("path");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const IMG = (f) => path.join(__dirname, "imagens", f);
const OUT = path.join(__dirname, "Claudio_Castro_de_Oliveira_Junior.pptx");

// Paleta do moodboard do TCC
const C = {
  navy: "1F2A44",
  terra: "B0573F",
  track: "B85A40",
  sand: "BFA58E",
  olive: "6B7751",
  forest: "2F5438",
  peach: "E9CDB8",
  blue: "2F6592",
  bg: "FBF9F1",
  card: "FFFFFF",
  ink: "2B2B2B",
  muted: "6E6A60",
  line: "DDD6C6",
};
const DF = "Arial"; // display: títulos em caixa alta, negrito itálico, como no TCC
const HF = "Arial";
const BF = "Calibri";
const W = 13.333, H = 7.5;

const LEGS = [
  { k: "LV", q: "Onde estou?", n: "Levantamento" },
  { k: "PN", q: "O que preciso?", n: "Programa" },
  { k: "EP", q: "Qual é a ideia?", n: "Estudo preliminar" },
  { k: "AP", q: "Como fica de verdade?", n: "Anteprojeto" },
];

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}

const shadow = () => ({ type: "outer", color: "000000", opacity: 0.13, blur: 8, offset: 3, angle: 90 });

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title = "NBR 16636-1 — Etapas Iniciais do Projeto";
  pres.author = "Cláudio Castro de Oliveira Júnior";

  const T = (s, text, o) => s.addText(text, { margin: 0, isTextBox: true, fontFace: BF, color: C.ink, ...o });

  // Pista de atletismo: faixa terracota com raias brancas
  function track(s, y, h, lanes = 3) {
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y, w: W, h, fill: { color: C.track }, line: { color: C.track } });
    for (let i = 1; i < lanes; i++) {
      s.addShape(pres.shapes.LINE, { x: 0, y: y + (h * i) / lanes, w: W, h: 0, line: { color: "FFFFFF", width: 1.25, transparency: 30 } });
    }
  }

  // Mini-pista de navegação no rodapé dos slides de etapa
  function progress(s, cur) {
    const y = 6.85, h = 0.65;
    track(s, y, h, 1);
    const seg = (W - 1.2) / LEGS.length;
    LEGS.forEach((l, i) => {
      const x = 0.6 + i * seg;
      const on = i === cur, done = i < cur;
      s.addShape(pres.shapes.OVAL, {
        x, y: y + 0.13, w: 0.4, h: 0.4,
        fill: { color: on ? C.navy : done ? C.peach : C.track }, line: { color: "FFFFFF", width: 1.5 },
      });
      T(s, String(i + 1), { x, y: y + 0.13, w: 0.4, h: 0.4, align: "center", valign: "middle", fontFace: HF, fontSize: 12, bold: true, color: on ? "FFFFFF" : done ? C.navy : "FFFFFF" });
      T(s, `${l.k}  ·  ${l.q}`, { x: x + 0.5, y: y + 0.13, w: seg - 0.6, h: 0.4, valign: "middle", fontFace: HF, fontSize: 12, bold: on, color: "FFFFFF", transparency: on || done ? 0 : 25 });
    });
  }

  function legHeader(s, i, title) {
    s.background = { color: C.bg };
    T(s, `${i + 1}ª PERNA  ·  ${LEGS[i].k}-ARQ`, { x: 0.6, y: 0.4, w: 6, h: 0.35, fontFace: HF, fontSize: 13, bold: true, color: C.terra, charSpacing: 3 });
    T(s, title, { x: 0.6, y: 0.78, w: 9.6, h: 0.85, fontFace: DF, fontSize: 38, bold: true, italic: true, color: C.navy });
  }

  function card(s, x, y, w, h, fill = C.card) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: C.line, width: 0.75 }, shadow: shadow() });
  }

  function batonNote(s, text, x = 0.6, y = 6.2, w = 12.1) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.48, rectRadius: 0.24, fill: { color: C.navy }, line: { color: C.navy } });
    T(s, [
      { text: "BASTÃO PASSADO  ▸  ", options: { bold: true, color: C.peach, fontFace: HF, fontSize: 12, charSpacing: 1 } },
      { text, options: { color: "FFFFFF", fontSize: 14 } },
    ], { x: x + 0.3, y, w: w - 0.6, h: 0.48, valign: "middle" });
  }

  // ============================================================ 1. CAPA
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    track(s, 5.55, 1.95, 4);
    // números de largada nas raias
    ["LV", "PN", "EP", "AP"].forEach((k, i) => {
      T(s, k, { x: 0.35, y: 5.55 + i * 0.4875, w: 0.8, h: 0.4875, valign: "middle", fontFace: DF, fontSize: 18, bold: true, italic: true, color: "FFFFFF", transparency: 15 });
    });
    s.addImage({ path: IMG("atleta_chute.png"), x: 0.45, y: 2.75, w: 3.0, h: 3.2 * 868 / 772, altText: "Atleta chutando bola" });
    s.addImage({ path: IMG("atleta_saque.png"), x: 10.35, y: 2.3, w: 2.45, h: 2.45 * 738 / 515, altText: "Atleta sacando" });
    T(s, "TRABALHO 07  ·  NBR 16636-1", { x: 3.6, y: 0.55, w: 6.2, h: 0.4, align: "center", fontFace: HF, fontSize: 14, bold: true, color: C.terra, charSpacing: 4 });
    T(s, "DO TERRENO\nAO ANTEPROJETO", { x: 1.6, y: 1.0, w: 10.1, h: 1.9, align: "center", fontFace: DF, fontSize: 54, lineSpacingMultiple: 0.95, bold: true, italic: true, color: C.navy });
    T(s, "As etapas iniciais do projeto explicadas como uma corrida de revezamento, com o meu TCC, o Centro Esportivo Arena Poty, na pista.", {
      x: 3.75, y: 3.05, w: 5.8, h: 0.95, align: "center", fontSize: 17, color: C.ink,
    });
    T(s, [
      { text: "Cláudio Castro de Oliveira Júnior", options: { bold: true, breakLine: true } },
      { text: "Arquitetura e Urbanismo  ·  Prof. Danilo Sérvio  ·  2026" },
    ], { x: 3.75, y: 4.2, w: 5.8, h: 0.75, align: "center", fontSize: 14, color: C.muted });
    s.addNotes("Meu trabalho é o 07, sobre a NBR 16636-1 e as etapas iniciais do projeto. Para ficar fácil de acompanhar, vou tratar o projeto como uma corrida de revezamento, e o atleta dessa corrida é o meu TCC, o Centro Esportivo Arena Poty.");
  }

  // ============================================================ 2. O QUE DIZ A NORMA
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    T(s, "A NORMA", { x: 0.6, y: 0.4, w: 6, h: 0.35, fontFace: HF, fontSize: 13, bold: true, color: C.terra, charSpacing: 3 });
    T(s, "O QUE DIZ A NBR 16636-1", { x: 0.6, y: 0.78, w: 12, h: 0.85, fontFace: DF, fontSize: 38, bold: true, italic: true, color: C.navy });
    const defs = [
      [fa.FaBullseye, "Objetivo", "Estabelecer as diretrizes e a terminologia para elaborar e desenvolver os serviços técnicos de projeto de arquitetura e urbanismo."],
      [fa.FaLayerGroup, "Campo de abrangência", "Projetos de arquitetura e urbanismo e as engenharias ligadas a eles, em obras novas, reformas e ampliações. A Parte 2 detalha o projeto arquitetônico."],
      [fa.FaTools, "Serviço técnico de projeto", "Conjunto de atividades feitas em etapas sucessivas, da coleta de dados até a documentação da obra. Cada etapa tem produtos definidos."],
    ];
    for (let i = 0; i < defs.length; i++) {
      const [ic, t, d] = defs[i];
      const y = 1.85 + i * 1.6;
      s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.8, h: 0.8, fill: { color: i === 2 ? C.terra : C.navy }, line: { color: "FFFFFF" } });
      s.addImage({ data: await icon(ic, "FFFFFF"), x: 0.8, y: y + 0.2, w: 0.4, h: 0.4 });
      T(s, t, { x: 1.6, y: y - 0.02, w: 4.5, h: 0.4, fontFace: HF, fontSize: 16, bold: true, color: C.navy });
      T(s, d, { x: 1.6, y: y + 0.4, w: 4.5, h: 1.05, valign: "top", fontSize: 13 });
    }
    // tabela de parâmetros normativos: etapas
    T(s, "Etapas do projeto de arquitetura previstas na norma", { x: 6.55, y: 1.8, w: 6.2, h: 0.35, fontFace: HF, fontSize: 14, bold: true, color: C.navy });
    const hdr = { bold: true, color: "FFFFFF", fill: { color: C.navy }, fontFace: HF, fontSize: 11, valign: "middle" };
    const rowsData = [
      ["LV-ARQ", "Levantamento de dados", "Dados do terreno, do entorno e da legislação"],
      ["PN-ARQ", "Programa de necessidades", "Lista de ambientes, áreas e relações"],
      ["EV-ARQ", "Estudo de viabilidade", "Análise técnica, legal e econômica da proposta"],
      ["EP-ARQ", "Estudo preliminar", "Partido, soluções conceituais e pré-dimensionamento"],
      ["AP-ARQ", "Anteprojeto", "Volumetria e solução funcional definidas, plantas cotadas"],
      ["PL-ARQ", "Projeto legal", "Peças para aprovação na prefeitura e nos órgãos"],
      ["PB-ARQ", "Projeto básico (opcional)", "Detalhamento para orçamento e licitação"],
      ["PE-ARQ", "Projeto executivo", "Detalhamento completo para a obra"],
    ];
    const rows = [[{ text: "Sigla", options: hdr }, { text: "Etapa", options: hdr }, { text: "Produto principal", options: hdr }]];
    rowsData.forEach(([a, b, c], i) => {
      const on = i < 5;
      const base = { fontFace: BF, fontSize: 11, valign: "middle", color: on ? C.ink : C.muted, fill: { color: on ? "FFFFFF" : "F1EEE6" } };
      rows.push([
        { text: a, options: { ...base, bold: true, color: on ? "FFFFFF" : C.muted, fill: { color: on ? C.terra : "E4DFD3" } } },
        { text: b, options: { ...base, bold: on } },
        { text: c, options: base },
      ]);
    });
    s.addTable(rows, { x: 6.55, y: 2.2, w: 6.2, colW: [0.95, 2.05, 3.2], rowH: 0.44, border: { type: "solid", pt: 0.75, color: C.line } });
    T(s, [
      { text: "■ ", options: { color: C.terra } },
      { text: "etapas iniciais: tema deste trabalho   " },
      { text: "■ ", options: { color: "E4DFD3" } },
      { text: "etapas finais: Trabalho 08" },
    ], { x: 6.55, y: 6.3, w: 6.2, h: 0.35, fontSize: 11, color: C.muted });
    s.addNotes("A NBR 16636-1 tem como objetivo estabelecer as diretrizes e a terminologia dos serviços técnicos de projeto de arquitetura e urbanismo. Vale para esses projetos e para as engenharias ligadas a eles, em obra nova, reforma ou ampliação. O serviço técnico de projeto é organizado em etapas sucessivas, cada uma com seus produtos, como mostra a tabela. As cinco primeiras são as etapas iniciais, tema deste trabalho; as finais ficam para o Trabalho 08.");
  }

  // ============================================================ 3. REVEZAMENTO + ESTUDO DE CASO
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    T(s, "A NORMA EM UMA IMAGEM", { x: 0.6, y: 0.4, w: 6, h: 0.35, fontFace: HF, fontSize: 13, bold: true, color: C.terra, charSpacing: 3 });
    T(s, "PROJETAR É UM REVEZAMENTO", { x: 0.6, y: 0.78, w: 9.7, h: 0.85, fontFace: DF, fontSize: 38, bold: true, italic: true, color: C.navy });
    T(s, [
      { text: "Cada etapa responde uma pergunta e entrega um produto, o " },
      { text: "bastão", options: { bold: true, color: C.terra } },
      { text: ", que a próxima usa como ponto de partida. Pular etapa vira retrabalho lá na frente." },
    ], { x: 0.6, y: 1.72, w: 9.3, h: 0.7, fontSize: 16 });

    const all = [
      ["LV", "Levantamento", "Onde estou?"], ["PN", "Programa", "O que preciso?"], ["EV", "Viabilidade", "Dá para fazer?"],
      ["EP", "Estudo preliminar", "Qual é a ideia?"], ["AP", "Anteprojeto", "Como fica de verdade?"],
      ["PL", "Projeto legal", "A prefeitura aprova?"], ["PB", "Projeto básico", "Quanto custa?"], ["PE", "Projeto executivo", "Como se constrói?"],
    ];
    track(s, 3.15, 1.3, 2);
    const x0 = 0.55, step = 1.555;
    all.forEach(([k, n, q], i) => {
      const x = x0 + i * step, on = i < 5;
      s.addShape(pres.shapes.OVAL, { x: x + 0.27, y: 3.25, w: 1.1, h: 1.1, fill: { color: on ? C.navy : C.bg }, line: { color: "FFFFFF", width: 2.5 } });
      T(s, k, { x: x + 0.27, y: 3.25, w: 1.1, h: 1.1, align: "center", valign: "middle", fontFace: DF, fontSize: 26, bold: true, italic: true, color: on ? "FFFFFF" : C.muted });
      T(s, n, { x, y: 4.6, w: 1.65, h: 0.35, align: "center", fontFace: HF, fontSize: 12, bold: true, color: on ? C.navy : C.muted });
      T(s, q, { x, y: 4.95, w: 1.65, h: 0.55, align: "center", fontSize: 13, italic: true, color: on ? C.terra : C.muted });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 2.6, w: 7.7, h: 0.4, rectRadius: 0.2, fill: { color: C.peach }, line: { color: C.peach } });
    T(s, "ETAPAS INICIAIS: TEMA DE HOJE", { x: 0.6, y: 2.6, w: 7.7, h: 0.4, align: "center", valign: "middle", fontFace: HF, fontSize: 12, bold: true, color: C.navy, charSpacing: 2 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.45, y: 2.6, w: 4.3, h: 0.4, rectRadius: 0.2, fill: { color: C.bg }, line: { color: C.line, width: 1 } });
    T(s, "ETAPAS FINAIS: TRABALHO 08", { x: 8.45, y: 2.6, w: 4.3, h: 0.4, align: "center", valign: "middle", fontFace: HF, fontSize: 12, bold: true, color: C.muted, charSpacing: 2 });
    s.addImage({ path: IMG("atleta_passe.png"), x: 10.6, y: 0.3, w: 2.15, h: 2.15 * 710 / 723, altText: "Atleta" });

    // o atleta: Arena Poty
    card(s, 0.6, 5.75, 12.15, 1.3, C.navy);
    T(s, "NOSSO ATLETA", { x: 0.9, y: 5.9, w: 3, h: 0.3, fontFace: HF, fontSize: 11, bold: true, color: C.sand, charSpacing: 2 });
    T(s, "Centro Esportivo Arena Poty", { x: 0.9, y: 6.2, w: 4.2, h: 0.45, fontFace: DF, fontSize: 22, bold: true, italic: true, color: "FFFFFF" });
    T(s, "TCC · Teresina – PI", { x: 0.9, y: 6.62, w: 4.8, h: 0.3, fontSize: 12, color: C.sand });
    [["14.762 m²", "de terreno"], ["10.936 m²", "de programa"], ["9", "setores"]].forEach(([v, l], i) => {
      const x = 6.0 + i * 2.25;
      T(s, v, { x, y: 5.9, w: 2.1, h: 0.6, valign: "bottom", fontFace: DF, fontSize: 26, bold: true, italic: true, color: C.peach });
      T(s, l, { x, y: 6.52, w: 2.1, h: 0.3, fontSize: 12, color: C.sand });
    });
    s.addNotes("Para ficar fácil de acompanhar, trato o projeto como um revezamento: cada etapa responde uma pergunta simples e passa um produto, o bastão, para a próxima. O atleta dessa corrida é o meu TCC, o Centro Esportivo Arena Poty, em Teresina: 14.762 m² de terreno e 10.936 m² de programa em 9 setores. Nos próximos slides ele corre as etapas iniciais.");
  }

  // ============================================================ 4. LV: ONDE ESTOU?
  {
    const s = pres.addSlide();
    legHeader(s, 0, "ONDE ESTOU?");
    T(s, "Antes de desenhar, conhecer o lugar. A norma pede estes levantamentos:", { x: 0.6, y: 1.72, w: 6.9, h: 0.4, fontSize: 16 });
    const items = [
      [fa.FaDraftingCompass, "Topográfico cadastral", "limites, medidas, níveis, vias"],
      [fa.FaSun, "Físico / ambiental", "sol, ventos, solo, vegetação"],
      [fa.FaBuilding, "Arquitetônico", "o que já está construído"],
      [fa.FaCamera, "Fotográfico", "registro do lote e das visadas"],
      [fa.FaCity, "Entorno imediato", "vizinhos, fluxos, equipamentos"],
      [fa.FaBalanceScale, "Legislação", "zona, índices, recuos"],
    ];
    for (let i = 0; i < items.length; i++) {
      const [ic, t, d] = items[i];
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.6 + col * 3.5, y = 2.3 + row * 1.25;
      s.addShape(pres.shapes.OVAL, { x, y, w: 0.85, h: 0.85, fill: { color: C.terra }, line: { color: C.terra } });
      s.addImage({ data: await icon(ic, "FFFFFF"), x: x + 0.22, y: y + 0.22, w: 0.41, h: 0.41 });
      T(s, t, { x: x + 1.0, y: y + 0.03, w: 2.4, h: 0.4, fontFace: HF, fontSize: 14, bold: true, color: C.navy });
      T(s, d, { x: x + 1.0, y: y + 0.43, w: 2.4, h: 0.4, fontSize: 13, color: C.muted });
    }
    // mapa do TCC
    card(s, 7.75, 0.45, 5.0, 4.6);
    s.addImage({ path: IMG("site.png"), x: 7.95, y: 0.58, w: 4.6, h: 4.35, sizing: { type: "contain", w: 4.6, h: 4.35 }, altText: "Mapa de condicionantes do TCC" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.75, y: 5.2, w: 5.0, h: 1.48, rectRadius: 0.1, fill: { color: C.navy }, line: { color: C.navy } });
    T(s, "No TCC: lote cercado por vias; sol nasce a leste e se põe na Av. Raul Lopes (oeste); vento de leste", { x: 7.95, y: 5.2, w: 4.6, h: 1.48, valign: "middle", fontSize: 13, color: "FFFFFF" });
    batonNote(s, "o mapa de condicionantes do terreno", 0.6, 6.2, 6.9);
    progress(s, 0);
    s.addNotes("Primeira perna: onde estou? É o levantamento de dados. A norma pede o topográfico cadastral, o físico, o arquitetônico, o fotográfico e a análise do entorno; na prática também entra a legislação. No TCC isso virou este mapa: as vias do entorno, a trajetória do sol, que se põe do lado da Avenida Raul Lopes, e os ventos predominantes de leste.");
  }

  // ============================================================ 5. LV: O QUE A LEI DEIXA
  {
    const s = pres.addSlide();
    legHeader(s, 0, "O QUE A LEI DEIXA FAZER?");
    T(s, "A lei da zona ZOM-4 também é levantamento:", { x: 0.6, y: 1.72, w: 7.7, h: 0.4, fontSize: 16 });

    // Barra do terreno: 100% com TO 80% e TP 7,5%
    const bx = 0.6, bw = 7.6, by = 2.6, bh = 1.0;
    T(s, "Terreno: 14.762 m²", { x: bx, y: by - 0.42, w: 4, h: 0.35, fontFace: HF, fontSize: 13, bold: true, color: C.navy });
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: bw * 0.8, h: bh, fill: { color: C.terra }, line: { color: "FFFFFF", width: 1 } });
    s.addShape(pres.shapes.RECTANGLE, { x: bx + bw * 0.8, y: by, w: bw * 0.125, h: bh, fill: { color: C.sand }, line: { color: "FFFFFF", width: 1 } });
    s.addShape(pres.shapes.RECTANGLE, { x: bx + bw * 0.925, y: by, w: bw * 0.075, h: bh, fill: { color: C.olive }, line: { color: "FFFFFF", width: 1 } });
    T(s, "até 80% pode ser ocupado  ·  11.809 m²", { x: bx + 0.2, y: by, w: bw * 0.8 - 0.3, h: bh, valign: "middle", fontFace: HF, fontSize: 15, bold: true, color: "FFFFFF" });
    T(s, [
      { text: "7,5% ", options: { bold: true, color: C.olive } },
      { text: "no mínimo tem que ser permeável: 1.107 m²" },
    ], { x: bx + bw * 0.52, y: by + bh + 0.12, w: bw * 0.48, h: 0.55, align: "right", fontSize: 13, color: C.ink });

    // IA = 6 terrenos
    T(s, "Índice de aproveitamento 6,0", { x: bx, y: 4.35, w: 5, h: 0.35, fontFace: HF, fontSize: 13, bold: true, color: C.navy });
    for (let i = 0; i < 6; i++) {
      s.addShape(pres.shapes.RECTANGLE, { x: bx + i * 0.85, y: 4.8, w: 0.72, h: 0.72, fill: { color: i === 0 ? C.navy : C.peach }, line: { color: C.navy, width: 1 } });
    }
    T(s, "= até 6 vezes a área do terreno em área construída: 88.572 m²", { x: bx + 5.2, y: 4.8, w: 2.6, h: 0.72, valign: "middle", fontSize: 13 });

    // corte esquemático: resposta ao clima (esquema ilustrativo do partido)
    card(s, 8.55, 1.75, 4.2, 4.25);
    T(s, "Corte esquemático: resposta ao clima", { x: 8.75, y: 1.88, w: 3.85, h: 0.3, fontFace: HF, fontSize: 12, bold: true, color: C.navy });
    T(s, "esquema ilustrativo do partido do TCC", { x: 8.75, y: 2.16, w: 3.85, h: 0.25, fontSize: 10, italic: true, color: C.muted });
    const G = 5.2;
    // sol da tarde (oeste) e raios
    s.addShape(pres.shapes.OVAL, { x: 8.8, y: 2.55, w: 0.42, h: 0.42, fill: { color: "E3A857" }, line: { color: "E3A857" } });
    [0, 0.25, 0.5].forEach((d) => s.addShape(pres.shapes.LINE, { x: 9.15 + d * 0.1, y: 2.95 + d * 0.2, w: 0.3, h: 0.8, line: { color: "E3A857", width: 1.25, dashType: "dash" } }));
    T(s, "sol da tarde (oeste)", { x: 9.3, y: 2.6, w: 1.6, h: 0.3, fontSize: 10, color: C.terra, bold: true });
    // edifício e cobertura com beiral
    s.addShape(pres.shapes.RECTANGLE, { x: 9.8, y: 3.75, w: 1.75, h: G - 3.75, fill: { color: "EFEBE0" }, line: { color: C.navy, width: 1.25 } });
    s.addShape(pres.shapes.RECTANGLE, { x: 9.25, y: 3.6, w: 2.85, h: 0.15, fill: { color: C.navy }, line: { color: C.navy } });
    T(s, "grande cobertura com beiral", { x: 10.2, y: 3.28, w: 2.4, h: 0.28, fontSize: 10, color: C.navy });
    // brise vertical na fachada oeste
    for (let i = 0; i < 4; i++) s.addShape(pres.shapes.RECTANGLE, { x: 9.4 + i * 0.1, y: 3.85, w: 0.045, h: 1.2, fill: { color: C.terra }, line: { color: C.terra } });
    // cobogó na fachada leste
    for (let r = 0; r < 6; r++) for (let c = 0; c < 2; c++)
      s.addShape(pres.shapes.RECTANGLE, { x: 11.4 + c * 0.1, y: 3.95 + r * 0.19, w: 0.07, h: 0.13, fill: { color: C.sand }, line: { color: C.terra, width: 0.5 } });
    // vento leste atravessando o edifício
    s.addShape(pres.shapes.LINE, { x: 9.25, y: 4.45, w: 3.25, h: 0, line: { color: C.blue, width: 2, beginArrowType: "triangle" } });
    T(s, "vento leste", { x: 11.75, y: 4.1, w: 0.9, h: 0.28, fontSize: 10, bold: true, color: C.blue });
    // árvore e solo
    s.addShape(pres.shapes.RECTANGLE, { x: 12.3, y: 4.85, w: 0.05, h: G - 4.85, fill: { color: "7A5C3E" }, line: { color: "7A5C3E" } });
    s.addShape(pres.shapes.OVAL, { x: 12.1, y: 4.6, w: 0.45, h: 0.4, fill: { color: C.olive }, line: { color: C.olive } });
    s.addShape(pres.shapes.LINE, { x: 8.75, y: G, w: 3.85, h: 0, line: { color: C.navy, width: 1.5 } });
    T(s, [
      { text: "brise ", options: { bold: true, color: C.terra } }, { text: "barra o sol da tarde  ·  " },
      { text: "cobogó ", options: { bold: true, color: C.terra } }, { text: "deixa o vento passar  ·  " },
      { text: "beiral ", options: { bold: true, color: C.terra } }, { text: "sombreia as fachadas" },
    ], { x: 8.75, y: 5.3, w: 3.85, h: 0.6, fontSize: 10, color: C.ink });
    batonNote(s, "quanto ocupar, quanto construir e quanto deixar permeável");
    progress(s, 0);
    s.addNotes("Ainda no levantamento: a legislação. Na ZOM-4, até 80% do terreno pode ser ocupado, pelo menos 7,5% tem que ficar permeável e o aproveitamento é 6, ou seja, posso construir até seis vezes a área do terreno. Conclusão: área sobra; o que pesa no desenho é o clima quente. O corte esquemático à direita mostra a resposta do partido: brise contra o sol da tarde, cobogó para o vento leste passar e uma grande cobertura com beiral.");
  }

  // ============================================================ 6. PN: O QUE PRECISO?
  {
    const s = pres.addSlide();
    legHeader(s, 1, "O QUE PRECISO?");
    T(s, "O programa de necessidades lista cada ambiente e sua área. No Arena Poty, 10.936 m² divididos assim:", { x: 0.6, y: 1.72, w: 7.2, h: 0.7, fontSize: 16 });
    // waffle 10x10: 69 externas / 31 edificadas
    const wx = 0.6, wy = 2.6, cs = 0.3, gap = 0.05;
    for (let i = 0; i < 100; i++) {
      const r = Math.floor(i / 10), c = i % 10;
      const ext = i < 69;
      s.addShape(pres.shapes.RECTANGLE, { x: wx + c * (cs + gap), y: wy + r * (cs + gap), w: cs, h: cs, fill: { color: ext ? C.olive : C.terra }, line: { color: ext ? C.olive : C.terra, width: 0.5 } });
    }
    T(s, "cada quadrado ≈ 1% do programa", { x: wx, y: wy + 3.55, w: 3.5, h: 0.3, fontSize: 11, italic: true, color: C.muted });
    // legendas grandes
    T(s, "69%", { x: 4.3, y: 2.55, w: 1.7, h: 0.75, fontFace: DF, fontSize: 40, bold: true, italic: true, color: C.olive });
    T(s, [
      { text: "ao ar livre  ·  7.560 m²", options: { bold: true, breakLine: true } },
      { text: "campo society, pista de corrida, quadra de tênis, praças, estacionamento, paisagismo" },
    ], { x: 4.3, y: 3.3, w: 3.9, h: 0.8, fontSize: 12, color: C.ink });
    T(s, "31%", { x: 4.3, y: 4.2, w: 1.7, h: 0.75, fontFace: DF, fontSize: 40, bold: true, italic: true, color: C.terra });
    T(s, [
      { text: "edificado  ·  3.376 m²", options: { bold: true, breakLine: true } },
      { text: "piscina, quadra coberta, academia, cultural, saúde, administração, apoio" },
    ], { x: 4.3, y: 4.95, w: 3.9, h: 0.8, fontSize: 12, color: C.ink });

    // setores edificados (barras horizontais simples)
    card(s, 8.6, 1.75, 4.15, 4.25);
    T(s, "Setores edificados (m²)", { x: 8.85, y: 1.9, w: 3.7, h: 0.35, fontFace: HF, fontSize: 13, bold: true, color: C.navy });
    const set = [["Esportivo interno", 1510], ["Convivência / alimentação", 830], ["Cultural", 265], ["Técnico / serviços", 245], ["Apoio esportivo", 207], ["Administrativo", 178], ["Saúde", 80], ["Sanitários públicos", 61]];
    set.forEach(([n, v], i) => {
      const y = 2.35 + i * 0.45;
      T(s, n, { x: 8.85, y, w: 1.95, h: 0.38, valign: "middle", fontSize: 11 });
      const bw = Math.max(0.05, (v / 1510) * 1.2);
      s.addShape(pres.shapes.RECTANGLE, { x: 10.8, y: y + 0.08, w: bw, h: 0.22, fill: { color: C.terra }, line: { color: C.terra } });
      T(s, String(v), { x: 10.85 + bw, y, w: 0.6, h: 0.38, valign: "middle", fontSize: 11, bold: true, color: C.navy });
    });
    batonNote(s, "a lista de ambientes com áreas: é ela que vai ser “desenhada” no estudo preliminar");
    progress(s, 1);
    s.addNotes("Segunda perna: o que preciso? É o programa de necessidades. O dado que mais chama atenção é que quase 70% do programa é ao ar livre: campo, pista, praças. Só 31% é edificado, cerca de 3.376 m², com o setor esportivo como o maior. Isso já diz que o projeto é, antes de tudo, um projeto de espaços abertos.");
  }

  // ============================================================ 7. EP: QUAL É A IDEIA?
  {
    const s = pres.addSlide();
    legHeader(s, 2, "QUAL É A IDEIA?");
    T(s, "O estudo preliminar transforma o programa em uma ideia de projeto, em três passos:", { x: 0.6, y: 1.72, w: 12, h: 0.4, fontSize: 16 });
    const steps = [
      [fa.FaLightbulb, "CONCEITO", "a ideia-força", "“Esporte como elemento de conexão” entre pessoas, cidade e paisagem."],
      [fa.FaCompass, "PARTIDO", "a ideia vira decisão", "Praça de acesso como transição · pátios que ligam os setores · conexão com o Rio Poti · brises, vegetação, grandes coberturas e ventilação natural."],
      [fa.FaPencilRuler, "DIMENSIONAMENTO PRÉVIO", "cabe no terreno?", "3.376 m² edificados, bem abaixo dos 88.572 m² permitidos · 1.800 m² de paisagismo, acima dos 1.107 m² exigidos."],
    ];
    for (let i = 0; i < steps.length; i++) {
      const [ic, t, sub, d] = steps[i];
      const x = 0.6 + i * 4.15;
      card(s, x, 2.35, 3.8, 3.6);
      s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: 2.6, w: 0.95, h: 0.95, fill: { color: i === 1 ? C.terra : C.navy }, line: { color: "FFFFFF" } });
      s.addImage({ data: await icon(ic, "FFFFFF"), x: x + 0.54, y: 2.84, w: 0.47, h: 0.47 });
      T(s, `0${i + 1}`, { x: x + 2.6, y: 2.55, w: 1.0, h: 0.8, align: "right", fontFace: DF, fontSize: 40, bold: true, italic: true, color: C.peach });
      T(s, t, { x: x + 0.3, y: 3.7, w: 3.3, h: 0.4, fontFace: HF, fontSize: 15, bold: true, color: C.navy });
      T(s, sub, { x: x + 0.3, y: 4.08, w: 3.3, h: 0.35, fontSize: 13, italic: true, color: C.terra });
      T(s, d, { x: x + 0.3, y: 4.5, w: 3.25, h: 1.35, valign: "top", fontSize: 13 });
      if (i < 2) T(s, "▶", { x: x + 3.8, y: 3.8, w: 0.35, h: 0.5, align: "center", valign: "middle", fontSize: 18, color: C.terra });
    }
    batonNote(s, "conceito + partido + números conferidos. O desenho vem no próximo slide");
    progress(s, 2);
    s.addNotes("Terceira perna: qual é a ideia? O estudo preliminar tem três passos. O conceito: esporte como elemento de conexão. O partido: praça de acesso, pátios, conexão com o Rio Poti e conforto passivo, que responde ao problema de clima visto no levantamento. E o dimensionamento prévio, que confere se tudo cabe: cabe com folga.");
  }

  // ============================================================ 8. EP: O DESENHO
  {
    const s = pres.addSlide();
    legHeader(s, 2, "A IDEIA NO PAPEL");
    // fluxograma
    card(s, 0.6, 1.8, 4.6, 3.55);
    s.addImage({ path: IMG("fluxos.png"), x: 0.75, y: 1.95, w: 4.3, h: 3.25, sizing: { type: "contain", w: 4.3, h: 3.25 }, altText: "Fluxograma" });
    T(s, [
      { text: "Fluxograma  ", options: { bold: true, color: C.navy } },
      { text: "a praça central é o “coração”: todos os setores passam por ela" },
    ], { x: 0.6, y: 5.45, w: 4.6, h: 0.6, fontSize: 12 });
    // croqui com marcadores
    const cx = 5.55, cy = 1.8, cw = 7.2, ch = cw * 1411 / 2923;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx - 0.1, y: cy - 0.1, w: cw + 0.2, h: ch + 0.2, rectRadius: 0.1, fill: { color: "F3E6DC" }, line: { color: C.line }, shadow: shadow() });
    s.addImage({ path: IMG("croqui.png"), x: cx, y: cy, w: cw, h: ch, altText: "Croqui de implantação" });
    const pins = [
      [0.18, 0.51], [0.27, 0.37], [0.28, 0.5], [0.38, 0.37], [0.48, 0.37], [0.57, 0.35], [0.45, 0.53], [0.73, 0.36], [0.61, 0.55],
    ];
    pins.forEach(([px, py], i) => {
      s.addShape(pres.shapes.OVAL, { x: cx + px * cw - 0.17, y: cy + py * ch - 0.17, w: 0.34, h: 0.34, fill: { color: C.terra }, line: { color: "FFFFFF", width: 1.5 } });
      T(s, String(i + 1), { x: cx + px * cw - 0.17, y: cy + py * ch - 0.17, w: 0.34, h: 0.34, align: "center", valign: "middle", fontFace: HF, fontSize: 11, bold: true, color: "FFFFFF" });
    });
    const legend = ["Praça de acesso", "Cultural (2 pisos)", "Atendimento + adm", "Piscina", "Apoio esportivo", "Quadra poliesportiva", "Comedoria", "Campo society", "Quadra de tênis"];
    legend.forEach((l, i) => {
      const x = cx + (i % 3) * 2.4, y = cy + ch + 0.15 + Math.floor(i / 3) * 0.3;
      T(s, [{ text: `${i + 1}  `, options: { bold: true, color: C.terra } }, { text: l }], { x, y, w: 2.3, h: 0.28, fontSize: 11 });
    });
    T(s, [{ text: "━  ", options: { bold: true, color: "C0392B" } }, { text: "linha vermelha: pista de corrida contornando o conjunto" }], { x: 0.6, y: 6.1, w: 4.6, h: 0.3, fontSize: 11 });
    progress(s, 2);
    s.addNotes("Os produtos gráficos do estudo preliminar no TCC são dois. O fluxograma mostra a praça central como o coração do projeto, distribuindo todos os setores. E o croqui de implantação mostra a ideia no terreno: a praça de acesso na Avenida Raul Lopes, os blocos cultural, de atendimento, piscina, apoio e quadra ao longo da rua, a comedoria no centro, a pista de corrida contornando tudo e o campo e a quadra de tênis ao fundo.");
  }

  // ============================================================ 9. AP: COMO FICA DE VERDADE?
  {
    const s = pres.addSlide();
    legHeader(s, 3, "COMO FICA DE VERDADE?");
    T(s, "O anteprojeto pega a ideia aprovada e a deixa definida e medida. É a próxima etapa do Arena Poty:", { x: 0.6, y: 1.72, w: 9.3, h: 0.45, fontSize: 16 });
    s.addImage({ path: IMG("atleta_goleira.png"), x: 10.2, y: 0.35, w: 2.6, h: 2.6 * 609 / 898, altText: "Goleira" });
    // de → para
    T(s, "NO ESTUDO PRELIMINAR", { x: 0.6, y: 2.35, w: 5.2, h: 0.35, fontFace: HF, fontSize: 12, bold: true, color: C.muted, charSpacing: 2 });
    T(s, "NO ANTEPROJETO", { x: 7.1, y: 2.35, w: 5.6, h: 0.35, fontFace: HF, fontSize: 12, bold: true, color: C.terra, charSpacing: 2 });
    const rows = [
      ["Volumes esboçados no croqui", "Volumetria definida: alturas, cobertura, fachadas"],
      ["Setores ligados no fluxograma", "Solução funcional final, com rotas acessíveis (NBR 9050)"],
      ["Áreas estimadas do programa", "Plantas, cortes e elevações cotados, em escala"],
      ["Moodboard de referências", "Materiais especificados, com brises e cobogós posicionados"],
    ];
    rows.forEach(([a, b], i) => {
      const y = 2.8 + i * 0.72;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 5.6, h: 0.6, rectRadius: 0.1, fill: { color: "EFEBE0" }, line: { color: "EFEBE0" } });
      T(s, a, { x: 0.8, y, w: 5.3, h: 0.6, valign: "middle", fontSize: 14, color: C.muted });
      T(s, "➜", { x: 6.25, y, w: 0.8, h: 0.6, align: "center", valign: "middle", fontSize: 20, bold: true, color: C.terra });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.1, y, w: 5.65, h: 0.6, rectRadius: 0.1, fill: { color: C.card }, line: { color: C.line }, shadow: shadow() });
      T(s, b, { x: 7.3, y, w: 5.35, h: 0.6, valign: "middle", fontSize: 14, bold: true, color: C.navy });
    });
    s.addImage({ path: IMG("materiais.png"), x: 0.6, y: 5.72, w: 5.6, h: 5.6 * 520 / 2918, altText: "Materiais do moodboard" });
    batonNote(s, "segue para o projeto legal", 7.1, 5.97, 5.65);
    progress(s, 3);
    s.addNotes("Quarta perna: como fica de verdade? O anteprojeto pega a ideia e define tudo com medida. O volume esboçado vira volumetria definida, o fluxograma vira solução funcional com rotas acessíveis, as áreas estimadas viram plantas cotadas e o moodboard vira especificação de materiais. Esse é o bastão para o projeto legal, que é tema do Trabalho 08.");
  }

  // ============================================================ 10. CHEGADA
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };
    T(s, "LINHA DE CHEGADA", { x: 0.6, y: 0.55, w: 12.1, h: 0.9, fontFace: DF, fontSize: 44, bold: true, italic: true, color: "FFFFFF" });
    T(s, "Quatro perguntas para lembrar das etapas iniciais", { x: 0.6, y: 1.45, w: 12.1, h: 0.4, fontSize: 17, color: C.sand });
    const out = ["mapa de condicionantes + índices", "9 setores, 10.936 m²", "conceito, partido, fluxograma e croqui", "plantas cotadas e volumetria definida"];
    LEGS.forEach((l, i) => {
      const x = 0.6 + i * 3.1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.15, w: 2.85, h: 2.85, rectRadius: 0.15, fill: { color: i === 3 ? C.terra : "2C3A5C" }, line: { color: i === 3 ? C.terra : "3A4A70" } });
      T(s, l.k, { x: x + 0.25, y: 2.3, w: 2.4, h: 0.8, fontFace: DF, fontSize: 40, bold: true, italic: true, color: i === 3 ? "FFFFFF" : C.peach });
      T(s, l.q, { x: x + 0.25, y: 3.1, w: 2.4, h: 0.8, fontFace: HF, fontSize: 18, bold: true, color: "FFFFFF", valign: "top" });
      T(s, out[i], { x: x + 0.25, y: 4.0, w: 2.4, h: 0.85, fontSize: 13, color: "DCDCDC", valign: "top" });
    });
    track(s, 5.35, 0.75, 1);
    T(s, "Seguir a ordem da norma = menos retrabalho, mais clareza para quem projeta, aprova e constrói.", { x: 0.6, y: 5.35, w: 12.1, h: 0.75, valign: "middle", align: "center", fontFace: HF, fontSize: 16, bold: true, color: "FFFFFF" });
    T(s, "Referências: ABNT NBR 16636-1:2017. Elaboração e desenvolvimento de serviços técnicos especializados de projetos arquitetônicos e urbanísticos, Parte 1: Diretrizes e terminologia.  ·  OLIVEIRA JÚNIOR, C. C. Centro Esportivo Arena Poty. TCC II, Arquitetura e Urbanismo, Teresina, 2026.", {
      x: 0.6, y: 6.45, w: 12.1, h: 0.6, fontSize: 10, color: "AEB4C4",
    });
    s.addNotes("Para fechar, quatro perguntas: onde estou, o que preciso, qual é a ideia e como fica de verdade. São as quatro etapas iniciais da NBR 16636-1, e cada uma entregou um produto concreto no Arena Poty. Seguir essa ordem evita retrabalho e deixa o projeto claro para quem projeta, aprova e constrói. Obrigado.");
  }

  await pres.writeFile({ fileName: OUT });
  console.log("OK:", OUT);
})();
