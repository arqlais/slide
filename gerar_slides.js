// Gera a apresentação do Trabalho 07 — NBR 16636-1 (Etapas Iniciais do Projeto)
// Apresentação direta, etapa por etapa; estudo de caso: TCC Arena Poty
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
  { k: "LV", n: "Levantamento", e: "diagnóstico do lugar" },
  { k: "PN", n: "Programa", e: "ambientes e áreas" },
  { k: "EP", n: "Estudo preliminar", e: "conceito, partido e croqui" },
  { k: "AP", n: "Anteprojeto", e: "volumetria e plantas cotadas" },
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
      T(s, `${l.k}  ·  ${l.n}`, { x: x + 0.5, y: y + 0.13, w: seg - 0.6, h: 0.4, valign: "middle", fontFace: HF, fontSize: 12, bold: on, color: "FFFFFF", transparency: on || done ? 0 : 25 });
    });
  }

  function legHeader(s, i, title) {
    s.background = { color: C.bg };
    T(s, `ETAPA ${i + 1}  ·  ${LEGS[i].k}-ARQ`, { x: 0.6, y: 0.4, w: 6, h: 0.35, fontFace: HF, fontSize: 13, bold: true, color: C.terra, charSpacing: 3 });
    T(s, title, { x: 0.6, y: 0.78, w: 9.6, h: 0.85, fontFace: DF, fontSize: 38, bold: true, italic: true, color: C.navy });
  }

  function card(s, x, y, w, h, fill = C.card) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: C.line, width: 0.75 }, shadow: shadow() });
  }

  function batonNote(s, text, x = 0.6, y = 6.2, w = 12.1) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.48, rectRadius: 0.24, fill: { color: C.navy }, line: { color: C.navy } });
    T(s, [
      { text: "ENTREGA  ▸  ", options: { bold: true, color: C.peach, fontFace: HF, fontSize: 12, charSpacing: 1 } },
      { text, options: { color: "FFFFFF", fontSize: 14 } },
    ], { x: x + 0.3, y, w: w - 0.6, h: 0.48, valign: "middle" });
  }

  // ============================================================ 1. CAPA
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    track(s, 5.55, 1.95, 4);
    // as 4 etapas iniciais e suas entregas, nas raias
    LEGS.forEach((l, i) => {
      T(s, [
        { text: `${l.k}  `, options: { fontFace: DF, fontSize: 18, bold: true, italic: true } },
        { text: `${l.n}  →  ${l.e}`, options: { fontSize: 14 } },
      ], { x: 3.9, y: 5.55 + i * 0.4875, w: 8.5, h: 0.4875, valign: "middle", color: "FFFFFF" });
    });
    s.addImage({ path: IMG("atleta_chute.png"), x: 0.45, y: 2.75, w: 3.0, h: 3.2 * 868 / 772, altText: "Atleta chutando bola" });
    s.addImage({ path: IMG("atleta_saque.png"), x: 10.35, y: 2.3, w: 2.45, h: 2.45 * 738 / 515, altText: "Atleta sacando" });
    T(s, "TRABALHO 07  ·  NBR 16636-1", { x: 3.6, y: 0.55, w: 6.2, h: 0.4, align: "center", fontFace: HF, fontSize: 14, bold: true, color: C.terra, charSpacing: 4 });
    T(s, "DO TERRENO\nAO ANTEPROJETO", { x: 1.6, y: 1.0, w: 10.1, h: 1.9, align: "center", fontFace: DF, fontSize: 54, lineSpacingMultiple: 0.95, bold: true, italic: true, color: C.navy });
    T(s, "As etapas iniciais do projeto de arquitetura, aplicadas ao meu TCC: o Centro Esportivo Arena Poty.", {
      x: 3.75, y: 3.05, w: 5.8, h: 0.95, align: "center", fontSize: 17, color: C.ink,
    });
    T(s, [
      { text: "Cláudio Castro de Oliveira Júnior", options: { bold: true, breakLine: true } },
      { text: "Arquitetura e Urbanismo  ·  Prof. Danilo Sérvio  ·  2026" },
    ], { x: 3.75, y: 4.2, w: 5.8, h: 0.75, align: "center", fontSize: 14, color: C.muted });
    s.addNotes("Meu trabalho é o 07: a NBR 16636-1 e as etapas iniciais do projeto de arquitetura. Vou mostrar cada etapa com o que a norma diz e o que ela entregou no meu TCC, o Centro Esportivo Arena Poty.");
  }

  // ============================================================ 2. O QUE DIZ A NORMA
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    T(s, "A NORMA", { x: 0.6, y: 0.4, w: 6, h: 0.35, fontFace: HF, fontSize: 13, bold: true, color: C.terra, charSpacing: 3 });
    T(s, "O QUE DIZ A NBR 16636-1", { x: 0.6, y: 0.78, w: 12, h: 0.85, fontFace: DF, fontSize: 38, bold: true, italic: true, color: C.navy });
    const defs = [
      [fa.FaBullseye, "Objetivo", "Estabelecer os procedimentos gerais e as diretrizes para produzir as principais etapas dos serviços de projeto arquitetônico e urbanístico (item 1, Escopo)."],
      [fa.FaLayerGroup, "Campo de abrangência", "Projetos de edificações ou espaços abertos, públicos ou privados. Parte 1: diretrizes e terminologia · Parte 2: projeto arquitetônico · Parte 3: urbanístico. Substitui as NBR 13531 e 13532."],
      [fa.FaTools, "Serviço técnico de projeto", "Atividade técnica de profissional habilitado, física ou intelectual (item 3.113), feita em etapas: períodos de trabalho em sequência (item 3.56)."],
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
    T(s, "Etapas do projeto e o que a Parte 1 diz de cada uma", { x: 6.55, y: 1.8, w: 6.2, h: 0.35, fontFace: HF, fontSize: 14, bold: true, color: C.navy });
    const hdr = { bold: true, color: "FFFFFF", fill: { color: C.navy }, fontFace: HF, fontSize: 11, valign: "middle" };
    const rowsData = [
      ["LV-ARQ", "Levantamento", "Condições preexistentes (LV-PRE, 3.67)"],
      ["PN-ARQ", "Programa", "Necessidades do contratante (PGN, 3.91)"],
      ["EV-ARQ", "Viabilidade", "Seleção de alternativas (3.49)"],
      ["EP-ARQ", "Estudo preliminar", "Dimensiona os conceitos (3.54)"],
      ["AP-ARQ", "Anteprojeto", "Detalhamento inicial (3.4)"],
      ["PL-ARQ", "Projeto legal", "Aprovação e alvará (3.106)"],
      ["PB-ARQ", "Projeto básico", "Não definido na Parte 1"],
      ["PE-ARQ", "Projeto executivo", "Informações finais para a obra (3.96)"],
    ];
    const rows = [[{ text: "Sigla", options: hdr }, { text: "Etapa", options: hdr }, { text: "O que a Parte 1 diz (item)", options: hdr }]];
    rowsData.forEach(([a, b, c], i) => {
      const on = i < 5;
      const base = { fontFace: BF, fontSize: 11, valign: "middle", color: on ? C.ink : C.muted, fill: { color: on ? "FFFFFF" : "F1EEE6" } };
      rows.push([
        { text: a, options: { ...base, bold: true, color: on ? "FFFFFF" : C.muted, fill: { color: on ? C.terra : "E4DFD3" } } },
        { text: b, options: { ...base, bold: on } },
        { text: c, options: base },
      ]);
    });
    s.addTable(rows, { x: 6.55, y: 2.2, w: 6.2, colW: [0.95, 1.85, 3.4], rowH: 0.42, border: { type: "solid", pt: 0.75, color: C.line } });
    T(s, [
      { text: "■ ", options: { color: C.terra } },
      { text: "etapas iniciais: tema deste trabalho   " },
      { text: "■ ", options: { color: "E4DFD3" } },
      { text: "etapas finais (Trab. 08)" },
    ], { x: 6.55, y: 6.15, w: 6.2, h: 0.3, fontSize: 11, color: C.muted });
    T(s, "Siglas -ARQ conforme o enunciado. Na Parte 1, o PL se chama “projeto para licenciamentos”.", { x: 6.55, y: 6.45, w: 6.2, h: 0.3, fontSize: 10, italic: true, color: C.muted });
    s.addNotes("Pelo escopo, a NBR 16636-1 estabelece os procedimentos gerais e as diretrizes para produzir as principais etapas dos projetos arquitetônicos e urbanísticos. Vale para edificações e espaços abertos, públicos ou privados, e substitui as antigas NBR 13531 e 13532. Serviço técnico é a atividade de um profissional habilitado, e ela acontece em etapas, que a norma define como períodos de trabalho em sequência. A tabela resume o que a Parte 1 diz de cada etapa; as siglas terminadas em ARQ seguem o enunciado. As cinco primeiras são as etapas iniciais, tema deste trabalho; as finais ficam para o Trabalho 08.");
  }

  // ============================================================ 3. ETAPAS EM SEQUÊNCIA + ESTUDO DE CASO
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    T(s, "A NORMA NA PRÁTICA", { x: 0.6, y: 0.4, w: 6, h: 0.35, fontFace: HF, fontSize: 13, bold: true, color: C.terra, charSpacing: 3 });
    T(s, "AS ETAPAS EM SEQUÊNCIA", { x: 0.6, y: 0.78, w: 9.7, h: 0.85, fontFace: DF, fontSize: 38, bold: true, italic: true, color: C.navy });
    T(s, [
      { text: "Cada etapa " },
      { text: "entrega um produto", options: { bold: true, color: C.terra } },
      { text: " que é a base da seguinte, e a próxima só começa depois que a anterior é aceita pelo contratante (item 5.2.2)." },
    ], { x: 0.6, y: 1.72, w: 9.3, h: 0.7, fontSize: 16 });

    const all = [
      ["LV", "Levantamento", "diagnóstico do lugar"], ["PN", "Programa", "ambientes e áreas"], ["EV", "Viabilidade", "alternativas avaliadas"],
      ["EP", "Estudo preliminar", "conceito e\npartido"], ["AP", "Anteprojeto", "plantas\ncotadas"],
      ["PL", "Projeto legal", "peças para aprovação"], ["PB", "Projeto básico", "base do orçamento"], ["PE", "Projeto executivo", "detalhes\nda obra"],
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

    // estudo de caso: Arena Poty
    card(s, 0.6, 5.75, 12.15, 1.3, C.navy);
    T(s, "ESTUDO DE CASO", { x: 0.9, y: 5.9, w: 3, h: 0.3, fontFace: HF, fontSize: 11, bold: true, color: C.sand, charSpacing: 2 });
    T(s, "Centro Esportivo Arena Poty", { x: 0.9, y: 6.2, w: 4.2, h: 0.45, fontFace: DF, fontSize: 22, bold: true, italic: true, color: "FFFFFF" });
    T(s, "TCC · Teresina – PI", { x: 0.9, y: 6.62, w: 4.8, h: 0.3, fontSize: 12, color: C.sand });
    [["14.762 m²", "de terreno"], ["10.936 m²", "de programa"], ["9", "setores"]].forEach(([v, l], i) => {
      const x = 6.0 + i * 2.25;
      T(s, v, { x, y: 5.9, w: 2.1, h: 0.6, valign: "bottom", fontFace: DF, fontSize: 26, bold: true, italic: true, color: C.peach });
      T(s, l, { x, y: 6.52, w: 2.1, h: 0.3, fontSize: 12, color: C.sand });
    });
    s.addNotes("As etapas vêm em sequência, e cada uma entrega um produto que é a base da seguinte. A norma exige que o contratante aceite os documentos de uma etapa antes de a próxima começar, item 5.2.2. Hoje o foco são as cinco iniciais. O estudo de caso é o Arena Poty: 14.762 m² de terreno e 10.936 m² de programa.");
  }

  // ============================================================ 4. LV: ONDE ESTOU?
  {
    const s = pres.addSlide();
    legHeader(s, 0, "LEVANTAMENTO DE DADOS");
    T(s, [
      { text: "Na norma (LV-PRE, 3.67): ", options: { bold: true, color: C.terra } },
      { text: "coleta das informações das condições preexistentes do lugar." },
    ], { x: 0.6, y: 1.68, w: 12.1, h: 0.4, fontSize: 15 });
    // o que o enunciado pede x o que o TCC produziu
    const items = [
      [fa.FaDraftingCompass, "Topográfico cadastral", "limites, medidas e área: 14.762,06 m²"],
      [fa.FaSun, "Físico", "sol poente a oeste, vento leste"],
      [fa.FaBuilding, "Arquitetônico", "lote vazio: nada construído a levantar"],
      [fa.FaCamera, "Fotográfico", "registro do lote e das visadas"],
      [fa.FaCity, "Entorno imediato", "vias, usos, gabaritos e equipamentos"],
    ];
    for (let i = 0; i < items.length; i++) {
      const [ic, t, d] = items[i];
      const y = 2.25 + i * 0.78;
      s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.6, h: 0.6, fill: { color: C.terra }, line: { color: C.terra } });
      s.addImage({ data: await icon(ic, "FFFFFF"), x: 0.76, y: y + 0.16, w: 0.28, h: 0.28 });
      T(s, t, { x: 1.35, y: y - 0.02, w: 3.3, h: 0.32, fontFace: HF, fontSize: 13, bold: true, color: C.navy });
      T(s, d, { x: 1.35, y: y + 0.29, w: 3.4, h: 0.32, fontSize: 12, color: C.muted });
    }
    // mapas do TCC
    const maps = [
      ["mapa_territ.png", "Terreno", "medidas e área"],
      ["mapa_clima.png", "Clima", "sol e ventos"],
      ["mapa_viario.png", "Sistema viário", "Raul Lopes: via rápida"],
      ["mapa_uso.png", "Uso do solo", "residência, serviço, comércio"],
      ["mapa_gabarito.png", "Gabarito", "térreo a 5+ pavimentos"],
      ["mapa_equip.png", "Equipamentos", "escolas e esporte perto"],
    ];
    maps.forEach(([f, t, d], i) => {
      const x = 5.05 + (i % 3) * 2.6, y = 2.2 + Math.floor(i / 3) * 1.95;
      s.addImage({ path: IMG(f), x: x + 0.52, y, w: 1.4, h: 1.4, altText: t });
      T(s, t, { x, y: y + 1.43, w: 2.45, h: 0.25, align: "center", fontFace: HF, fontSize: 11, bold: true, color: C.navy });
      T(s, d, { x, y: y + 1.67, w: 2.45, h: 0.25, align: "center", fontSize: 10, color: C.muted });
    });
    batonNote(s, "o diagnóstico do lugar: terreno, clima e entorno mapeados");
    progress(s, 0);
    s.addNotes("Etapa 1, levantamento de dados. A norma define como a coleta das informações das condições que já existem no lugar. No TCC isso virou seis mapas: terreno, clima, sistema viário, uso do solo, gabarito e equipamentos. O lote está vazio, então não há edificação existente para levantar.");
  }

  // ============================================================ 4b. LV: LEVANTAMENTO FOTOGRÁFICO
  {
    const s = pres.addSlide();
    legHeader(s, 0, "LEVANTAMENTO FOTOGRÁFICO");
    T(s, "Levantamento fotográfico: o registro do lote e do entorno, pedido no enunciado.", { x: 0.6, y: 1.68, w: 12.1, h: 0.4, fontSize: 15 });
    const camera = await icon(fa.FaCamera, C.sand);
    const caps = ["Lote vazio visto da avenida", "Ciclofaixa e iluminação pública", "Pedestres e corredores à noite", "A via já é usada para esporte"];
    caps.forEach((c, i) => {
      const x = 0.6 + i * 3.08, y = 2.25, w = 2.85, h = 2.14;
      // moldura: arraste a foto por cima deste retângulo
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: "EFEBE0" }, line: { color: C.sand, width: 1.25, dashType: "dash" } });
      s.addImage({ data: camera, x: x + w / 2 - 0.25, y: y + 0.6, w: 0.5, h: 0.5 });
      T(s, `FOTO ${i + 1}`, { x, y: y + 1.2, w, h: 0.3, align: "center", fontFace: HF, fontSize: 11, bold: true, color: C.sand, charSpacing: 2 });
      T(s, [{ text: `${i + 1}  `, options: { bold: true, color: C.terra } }, { text: c }], { x, y: y + h + 0.08, w, h: 0.3, fontSize: 12, color: C.ink });
    });
    card(s, 0.6, 4.95, 12.1, 1.05);
    T(s, [
      { text: "O que as fotos mostram:  ", options: { bold: true, color: C.terra } },
      { text: "lote vazio e cercado, carros estacionados ao longo da calçada, ciclofaixa e iluminação pública e, à noite, gente caminhando, correndo e pedalando. " },
      { text: "O lugar já tem vocação esportiva.", options: { bold: true, color: C.navy } },
    ], { x: 0.9, y: 4.95, w: 11.5, h: 1.05, valign: "middle", fontSize: 14 });
    batonNote(s, "registro do local: o lugar já é usado para esporte");
    progress(s, 0);
    s.addNotes("Ainda no levantamento, o registro fotográfico. As fotos mostram o lote vazio e cercado, a avenida com ciclofaixa e, à noite, muita gente caminhando e correndo. O lugar já é usado para esporte.");
  }

  // ============================================================ 5. LV: O QUE A LEI DEIXA
  {
    const s = pres.addSlide();
    legHeader(s, 0, "LEGISLAÇÃO DO TERRENO");
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

    // IA 6,0: pilha de 6 "andares", cada um do tamanho do terreno
    T(s, "Índice de aproveitamento 6,0", { x: bx, y: 4.3, w: 5, h: 0.35, fontFace: HF, fontSize: 13, bold: true, color: C.navy });
    for (let i = 0; i < 6; i++) {
      const y = 5.8 - i * 0.2;
      s.addShape(pres.shapes.RECTANGLE, { x: bx, y, w: 2.4, h: 0.15, fill: { color: i === 0 ? C.navy : C.peach }, line: { color: C.navy, width: 0.75 } });
    }
    T(s, [
      { text: "Somando todos os andares, posso construir até 6 vezes a área do terreno:", options: { breakLine: true } },
      { text: "6 × 14.762 m² = 88.572 m²", options: { bold: true, color: C.navy, breakLine: true } },
      { text: "O Arena Poty usa 3.376 m², cerca de 4% disso.", options: { color: C.terra, bold: true } },
    ], { x: bx + 2.75, y: 4.7, w: 4.9, h: 1.3, valign: "middle", fontSize: 12.5 });

    // corte esquemático: resposta ao clima (esquema ilustrativo do partido)
    card(s, 8.55, 1.75, 4.2, 4.25);
    T(s, "Corte esquemático: resposta ao clima", { x: 8.75, y: 1.88, w: 3.85, h: 0.3, fontFace: HF, fontSize: 12, bold: true, color: C.navy });
    T(s, "esquema elaborado para esta apresentação, a partir do partido do TCC", { x: 8.75, y: 2.16, w: 3.85, h: 0.25, fontSize: 10, italic: true, color: C.muted });
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
    batonNote(s, "limites de ocupação, construção e permeabilidade");
    progress(s, 0);
    s.addNotes("A legislação também faz parte do levantamento. Na ZOM-4, até 80% do terreno pode ser ocupado, no mínimo 7,5% fica permeável e, somando todos os andares, posso construir até 6 vezes a área do terreno: 88.572 m². O projeto usa só 3.376 m². Área sobra; o desafio é o clima. O corte, que eu fiz a partir do partido do TCC, mostra a resposta: brise, cobogó e beiral.");
  }

  // ============================================================ 6. PN: O QUE PRECISO?
  {
    const s = pres.addSlide();
    legHeader(s, 1, "PROGRAMA DE NECESSIDADES");
    T(s, "Na norma (PGN, item 3.91): as necessidades do contratante, organizadas.", { x: 0.6, y: 1.72, w: 7.7, h: 0.4, fontSize: 14 });
    // barra única do programa: 69% ao ar livre / 31% edificado
    const bx = 0.6, bw = 7.6, by = 2.75, bh = 1.0;
    T(s, "Programa total: 10.936 m²", { x: bx, y: by - 0.38, w: 5, h: 0.32, fontFace: HF, fontSize: 13, bold: true, color: C.navy });
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: bw * 0.69, h: bh, fill: { color: C.olive }, line: { color: "FFFFFF", width: 1 } });
    s.addShape(pres.shapes.RECTANGLE, { x: bx + bw * 0.69, y: by, w: bw * 0.31, h: bh, fill: { color: C.terra }, line: { color: "FFFFFF", width: 1 } });
    T(s, [{ text: "69%  ", options: { fontSize: 24, italic: true } }, { text: "ao ar livre · 7.560 m²" }], { x: bx + 0.2, y: by, w: bw * 0.69 - 0.3, h: bh, valign: "middle", fontFace: HF, fontSize: 14, bold: true, color: "FFFFFF" });
    T(s, [{ text: "31%  ", options: { fontSize: 24, italic: true } }, { text: "edificado" }], { x: bx + bw * 0.69 + 0.15, y: by, w: bw * 0.31 - 0.2, h: bh, valign: "middle", fontFace: HF, fontSize: 14, bold: true, color: "FFFFFF" });
    // o que entra em cada parte
    T(s, [
      { text: "Ao ar livre (m²)", options: { bold: true, color: C.olive, breakLine: true } },
      { text: "Campo society  1.800", options: { breakLine: true } },
      { text: "Paisagismo  1.800", options: { breakLine: true } },
      { text: "Estacionamento  1.500", options: { breakLine: true } },
      { text: "Pista de corrida  1.100", options: { breakLine: true } },
      { text: "Quadra de tênis  670", options: { breakLine: true } },
      { text: "Praças  650  ·  Bicicletário  40" },
    ], { x: bx, y: 4.0, w: 4.2, h: 1.95, valign: "top", fontSize: 12, color: C.ink });
    T(s, [
      { text: "Edificado (m²)", options: { bold: true, color: C.terra, breakLine: true } },
      { text: "3.376 m² em 8 setores,", options: { breakLine: true } },
      { text: "detalhados no quadro ao lado", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 6 } },
      { text: "Conclusão: é, antes de tudo, um projeto de espaços abertos.", options: { bold: true, color: C.navy } },
    ], { x: bx + 4.35, y: 4.0, w: 3.3, h: 1.95, valign: "top", fontSize: 12, color: C.ink });

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
    batonNote(s, "lista de ambientes com áreas, base do estudo preliminar");
    progress(s, 1);
    s.addNotes("Etapa 2, programa de necessidades: as necessidades do contratante organizadas em ambientes e áreas. São 10.936 m², e 69% é ao ar livre: campo, pista, praças e estacionamento. É, antes de tudo, um projeto de espaços abertos.");
  }

  // ============================================================ 7. EP: QUAL É A IDEIA?
  {
    const s = pres.addSlide();
    legHeader(s, 2, "ESTUDO PRELIMINAR");
    T(s, "Na norma (item 3.54): dimensionamento preliminar dos conceitos do projeto, podendo incluir alternativas. Em três passos:", { x: 0.6, y: 1.72, w: 12, h: 0.4, fontSize: 16 });
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
    batonNote(s, "conceito, partido e pré-dimensionamento conferido");
    progress(s, 2);
    s.addNotes("Etapa 3, estudo preliminar. A norma define como o dimensionamento preliminar dos conceitos do projeto. No TCC: o conceito, esporte como elemento de conexão; o partido, com praça de acesso, pátios, conexão com o Rio Poti e conforto passivo; e o pré-dimensionamento, que confirma que tudo cabe no terreno.");
  }

  // ============================================================ 8. EP: O DESENHO
  {
    const s = pres.addSlide();
    legHeader(s, 2, "FLUXOGRAMA E CROQUI");
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
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx + 0.12, y: cy + 0.12, w: 4.0, h: 0.36, rectRadius: 0.18, fill: { color: C.navy }, line: { color: C.navy } });
    T(s, "PLANO DE MASSAS  ·  item 3.87 da norma", { x: cx + 0.12, y: cy + 0.12, w: 4.0, h: 0.36, align: "center", valign: "middle", fontFace: HF, fontSize: 10, bold: true, color: "FFFFFF", charSpacing: 1 });
    progress(s, 2);
    s.addNotes("Os produtos gráficos do estudo preliminar: o fluxograma, com a praça central ligando todos os setores, e o croqui, que na norma é um plano de massas, item 3.87. Os números mostram onde fica cada setor no terreno.");
  }

  // ============================================================ 9. AP: COMO FICA DE VERDADE?
  {
    const s = pres.addSlide();
    legHeader(s, 3, "ANTEPROJETO");
    T(s, "Na norma (item 3.4): concepção e representação das informações técnicas iniciais de detalhamento. É a próxima etapa do TCC:", { x: 0.6, y: 1.72, w: 9.3, h: 0.45, fontSize: 16 });
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
    batonNote(s, "base do projeto legal (PL)", 7.1, 5.97, 5.65);
    progress(s, 3);
    s.addNotes("Etapa 4, anteprojeto. A norma define como as informações técnicas iniciais de detalhamento. É a próxima etapa do TCC: o croqui vira volumetria definida, o fluxograma vira solução funcional com rotas acessíveis, as áreas viram plantas cotadas e o moodboard vira especificação de materiais. Ele é a base do projeto legal.");
  }

  // ============================================================ 11. RESUMO
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };
    T(s, "RESUMO DAS ETAPAS INICIAIS", { x: 0.6, y: 0.55, w: 12.1, h: 0.9, fontFace: DF, fontSize: 40, bold: true, italic: true, color: "FFFFFF" });
    T(s, "O que cada etapa entregou no Arena Poty", { x: 0.6, y: 1.45, w: 12.1, h: 0.4, fontSize: 17, color: C.sand });
    const out = ["mapas do lugar + índices", "9 setores, 10.936 m²", "conceito, partido, fluxograma e croqui", "plantas cotadas e volumetria definida"];
    LEGS.forEach((l, i) => {
      const x = 0.6 + i * 3.1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.15, w: 2.85, h: 2.85, rectRadius: 0.15, fill: { color: i === 3 ? C.terra : "2C3A5C" }, line: { color: i === 3 ? C.terra : "3A4A70" } });
      T(s, l.k, { x: x + 0.25, y: 2.3, w: 2.4, h: 0.8, fontFace: DF, fontSize: 40, bold: true, italic: true, color: i === 3 ? "FFFFFF" : C.peach });
      T(s, l.n, { x: x + 0.25, y: 3.1, w: 2.4, h: 0.8, fontFace: HF, fontSize: 18, bold: true, color: "FFFFFF", valign: "top" });
      T(s, out[i], { x: x + 0.25, y: 4.0, w: 2.4, h: 0.85, fontSize: 13, color: "DCDCDC", valign: "top" });
    });
    track(s, 5.35, 0.75, 1);
    T(s, "Seguir a ordem da norma = menos retrabalho, mais clareza para quem projeta, aprova e constrói.", { x: 0.6, y: 5.35, w: 12.1, h: 0.75, valign: "middle", align: "center", fontFace: HF, fontSize: 16, bold: true, color: "FFFFFF" });
    T(s, "Referências: ABNT NBR 16636-1:2017. Elaboração e desenvolvimento de serviços técnicos especializados de projetos arquitetônicos e urbanísticos, Parte 1: Diretrizes e terminologia (itens 1, 3.4, 3.54, 3.67, 3.87, 3.91, 5.2.2).  ·  OLIVEIRA JÚNIOR, C. C. Centro Esportivo Arena Poty. TCC II, Arquitetura e Urbanismo, Teresina, 2026.", {
      x: 0.6, y: 6.45, w: 12.1, h: 0.6, fontSize: 10, color: "AEB4C4",
    });
    s.addNotes("Resumindo: levantamento, programa, estudo preliminar e anteprojeto, cada um com a sua entrega. Seguir essa ordem evita retrabalho e deixa o projeto claro para quem projeta, aprova e constrói. Obrigado.");
  }

  await pres.writeFile({ fileName: OUT });
  console.log("OK:", OUT);
})();
