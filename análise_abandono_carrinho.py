const pptxgen = require("pptxgenjs");

// ---------- DADOS (extraídos do gráfico fornecido) ----------
const months   = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Nov","Dez"];
const vendas   = [102,106,101,100, 80, 75, 70, 62, 60, 62, 65];
const abandono = [ 13, 13, 10,  9, 15, 20, 27, 32, 35, 33, 34];
const frete    = [ 11, 12, 12, 11, 25, 26, 25, 26, 25, 25, 26];
const ticket   = [ 55, 56, 54, 55, 56, 55, 56, 54, 55, 54, 55];
const freteTicketPct = frete.map((f,i)=> Math.round((f/ticket[i])*1000)/10);

// ---------- PALETA ----------
const DARK   = "0B3D3E";  // teal profundo
const TEAL   = "00A896";  // vendas / positivo
const RED    = "E63946";  // abandono / alerta
const AMBER  = "E8A94B";  // frete / custo
const GRAY   = "6B7280";  // ticket médio / neutro
const INK    = "1B2426";  // texto principal
const MUTE   = "5B6B6C";  // texto secundário
const LIGHT  = "FFFFFF";
const PANEL  = "F2F7F6";  // painel levemente esverdeado

const FONT_HEAD = "Cambria";
const FONT_BODY = "Calibri";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5
const PW = 13.33, PH = 7.5;

function footer(slide, n, light) {
  slide.addText(`${n}`, { x: PW-0.7, y: PH-0.5, w: 0.5, h: 0.3, fontFace: FONT_BODY,
    fontSize: 10, color: light ? "9FD8D2" : "A9B4B4", align: "right" });
}

function kicker(slide, text, color) {
  slide.addText(text.toUpperCase(), { x: 0.6, y: 0.5, w: 8, h: 0.35, fontFace: FONT_BODY,
    bold: true, fontSize: 12, color: color, charSpacing: 2 });
}

// ============================================================
// SLIDE 1 — TÍTULO
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };

  // motivo visual: carrinho estilizado com círculos concêntricos
  s.addShape("ellipse", { x: 10.3, y: 1.1, w: 2.5, h: 2.5, fill: { color: "0F4A4B" }, line: { type: "none" } });
  s.addShape("ellipse", { x: 10.75, y: 1.55, w: 1.6, h: 1.6, fill: { color: "12595A" }, line: { type: "none" } });

  s.addText("TIME DE CIÊNCIA DE DADOS  →  TIME DE MARKETING", { x: 0.7, y: 1.5, w: 9.5, h: 0.4,
    fontFace: FONT_BODY, bold: true, fontSize: 13, color: TEAL, charSpacing: 2 });

  s.addText("Abandono de carrinho:\nda causa à ação", { x: 0.65, y: 2.0, w: 10.5, h: 2.2,
    fontFace: FONT_HEAD, bold: true, fontSize: 44, color: LIGHT, lineSpacing: 48, isTextBox: true });

  s.addText("O que os dados revelam sobre a queda em vendas — e como um modelo preditivo vai nos ajudar a agir, cliente a cliente.",
    { x: 0.7, y: 4.3, w: 8.7, h: 1.0, fontFace: FONT_BODY, fontSize: 16, color: "CFE8E4", lineSpacing: 22, isTextBox: true });

  s.addShape("line", { x: 0.7, y: 5.55, w: 1.4, h: 0, line: { color: TEAL, width: 3 } });
  s.addText("Análise de storytelling de dados  ·  2026", { x: 0.7, y: 5.7, w: 6, h: 0.35,
    fontFace: FONT_BODY, fontSize: 11, color: "8FB8B4" });
}

// ============================================================
// SLIDE 2 — CONTEXTO DO DESAFIO
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  kicker(s, "O contexto", TEAL);
  s.addText("As vendas finalizadas vêm caindo mês após mês", { x: 0.6, y: 0.85, w: 11.8, h: 0.9,
    fontFace: FONT_HEAD, bold: true, fontSize: 30, color: INK, isTextBox: true });

  s.addText([
    { text: "O time de marketing identificou um aumento consistente de clientes que chegam ao carrinho, mas não concluem a compra. ", options:{} },
    { text: "Antes de qualquer ação, precisávamos entender: isso é ruído do dia a dia — ou existe uma causa real por trás da queda?", options: { bold: true, color: INK } },
  ], { x: 0.6, y: 1.9, w: 7.1, h: 2.0, fontFace: FONT_BODY, fontSize: 15.5, color: MUTE, lineSpacing: 24, isTextBox: true });

  s.addText("Este material responde a essa pergunta em duas partes:", { x: 0.6, y: 4.0, w: 7.1, h: 0.4,
    fontFace: FONT_BODY, bold: true, fontSize: 13.5, color: INK, isTextBox: true });

  const parts = [
    ["1", "O que os dados do gráfico de abandono revelam sobre a causa"],
    ["2", "Como o novo modelo preditivo vai nos ajudar a agir a tempo, cliente a cliente"],
  ];
  let py = 4.55;
  parts.forEach(([n, t]) => {
    s.addShape("ellipse", { x: 0.6, y: py, w: 0.42, h: 0.42, fill: { color: DARK }, line: { type: "none" } });
    s.addText(n, { x: 0.6, y: py, w: 0.42, h: 0.42, align: "center", valign: "middle",
      fontFace: FONT_BODY, bold: true, fontSize: 15, color: LIGHT });
    s.addText(t, { x: 1.2, y: py-0.03, w: 6.5, h: 0.5, valign:"middle", fontFace: FONT_BODY, fontSize: 14, color: INK, isTextBox: true });
    py += 0.68;
  });

  // painel lateral com o "sintoma" central em número grande
  s.addShape("roundRect", { x: 8.3, y: 1.75, w: 4.4, h: 4.3, rectRadius: 0.12, fill: { color: PANEL }, line: { type: "none" } });
  s.addText("TAXA DE ABANDONO", { x: 8.6, y: 2.05, w: 3.8, h: 0.35, fontFace: FONT_BODY, bold: true,
    fontSize: 11.5, color: MUTE, charSpacing: 1.5 });
  s.addText("9,9%", { x: 8.6, y: 2.4, w: 1.7, h: 0.9, fontFace: FONT_HEAD, bold: true, fontSize: 34, color: TEAL });
  s.addText("em Jan–Abr", { x: 8.6, y: 3.15, w: 3.6, h: 0.35, fontFace: FONT_BODY, fontSize: 12, color: MUTE });

  s.addShape("line", { x: 8.6, y: 3.65, w: 3.8, h: 0, line: { color: "D8E4E2", width: 1 } });

  s.addText("29,3%", { x: 8.6, y: 3.85, w: 2.2, h: 0.9, fontFace: FONT_HEAD, bold: true, fontSize: 34, color: RED });
  s.addText("em Mai–Dez", { x: 8.6, y: 4.6, w: 3.6, h: 0.35, fontFace: FONT_BODY, fontSize: 12, color: MUTE });

  s.addText("quase 3x mais abandono no 2º semestre", { x: 8.6, y: 5.15, w: 3.8, h: 0.7,
    fontFace: FONT_BODY, italic: true, fontSize: 12, color: INK, lineSpacing: 16 });

  footer(s, 2, false);
}

// ============================================================
// SLIDE 3 — O GRÁFICO (evidência bruta)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  kicker(s, "Cenário 1 · O gráfico", TEAL);
  s.addText("O que os dados de 2025 mostram, mês a mês", { x: 0.6, y: 0.85, w: 11.8, h: 0.7,
    fontFace: FONT_HEAD, bold: true, fontSize: 26, color: INK });

  s.addChart(
    [
      {
        type: pres.charts.BAR,
        data: [
          { name: "Vendas finalizadas", labels: months, values: vendas },
          { name: "Abandono de carrinho", labels: months, values: abandono },
        ],
        options: { chartColors: [TEAL, RED], barGrouping: "clustered" },
      },
      {
        type: pres.charts.LINE,
        data: [ { name: "Frete médio (R$)", labels: months, values: frete } ],
        options: { chartColors: [AMBER], secondaryValAxis: true, secondaryCatAxis: true, lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 6 },
      },
    ],
    {
      x: 0.5, y: 1.65, w: 12.3, h: 5.1,
      barGapWidthPct: 35,
      showTitle: false,
      showLegend: true, legendPos: "b", legendFontSize: 12, legendColor: INK,
      showValue: false,
      catAxisLabelColor: MUTE, catAxisLabelFontSize: 11,
      valAxes: [
        { showValAxisTitle: true, valAxisTitle: "Clientes", valAxisTitleFontSize: 11, valAxisTitleColor: MUTE,
          valAxisLabelColor: MUTE, valAxisLabelFontSize: 10, valGridLine: { color: "EAEFEE", size: 1 }, valAxisMinVal: 0, valAxisMaxVal: 120 },
        { showValAxisTitle: true, valAxisTitle: "Frete médio (R$)", valAxisTitleFontSize: 11, valAxisTitleColor: MUTE,
          valAxisLabelColor: MUTE, valAxisLabelFontSize: 10, valGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: 35 },
      ],
      catAxes: [
        { catAxisLabelColor: MUTE, catAxisLabelFontSize: 11 },
        { catAxisHidden: true },
      ],
    }
  );

  footer(s, 3, false);
}

// ============================================================
// SLIDE 4 — A DESCOBERTA CENTRAL (manchete)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  kicker(s, "Cenário 1 · A descoberta", TEAL);
  s.addText("O frete médio dobrou em maio — e o abandono nunca mais foi o mesmo", { x: 0.6, y: 0.85, w: 11.8, h: 1.3,
    fontFace: FONT_HEAD, bold: true, fontSize: 27, color: INK, lineSpacing: 32, isTextBox: true });

  s.addText("As duas curvas se movem juntas: assim que o frete médio sobe de forma abrupta, o abandono de carrinho começa a crescer mês a mês — e as vendas finalizadas caem na mesma proporção.",
    { x: 0.6, y: 2.15, w: 11.8, h: 0.75, fontFace: FONT_BODY, fontSize: 14.5, color: MUTE, lineSpacing: 20, isTextBox: true });

  const cards = [
    ["Frete médio", "R$ 11,50 → R$ 25,40", "+121%", AMBER],
    ["Taxa de abandono", "9,9% → 29,3%", "quase 3x", RED],
    ["Vendas finalizadas", "102 → 68 / mês", "-33%", TEAL],
  ];
  let cx = 0.6;
  const cw = 3.95, gap = 0.28;
  cards.forEach(([label, val, delta, color]) => {
    s.addShape("roundRect", { x: cx, y: 3.15, w: cw, h: 3.1, rectRadius: 0.1, fill: { color: PANEL }, line: { type: "none" } });
    s.addShape("roundRect", { x: cx+0.35, y: 3.5, w: 1.5, h: 0.42, rectRadius: 0.21, fill: { color: color }, line: { type: "none" } });
    s.addText(delta, { x: cx+0.35, y: 3.5, w: 1.5, h: 0.42, align: "center", valign: "middle",
      fontFace: FONT_BODY, bold: true, fontSize: 14, color: LIGHT });
    s.addText(label.toUpperCase(), { x: cx+0.35, y: 4.15, w: cw-0.7, h: 0.4, fontFace: FONT_BODY,
      bold: true, fontSize: 11.5, color: MUTE, charSpacing: 1 });
    s.addText(val, { x: cx+0.35, y: 4.55, w: cw-0.7, h: 1.3, fontFace: FONT_HEAD, bold: true,
      fontSize: 21, color: INK, lineSpacing: 26, isTextBox: true });
    cx += cw + gap;
  });

  s.addText("Antes (Jan–Abr)  →  Depois (Mai–Dez)", { x: 0.6, y: 6.45, w: 8, h: 0.35,
    fontFace: FONT_BODY, italic: true, fontSize: 12, color: MUTE });

  footer(s, 4, false);
}

// ============================================================
// SLIDE 5 — O "SMOKING GUN": FRETE / TICKET MÉDIO
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  kicker(s, "Cenário 1 · A causa mais provável", TEAL);
  s.addText("O ticket médio não mudou. O peso do frete no bolso do cliente, sim.", { x: 0.6, y: 0.85, w: 11.8, h: 1.0,
    fontFace: FONT_HEAD, bold: true, fontSize: 25, color: INK, lineSpacing: 30, isTextBox: true });

  s.addText("O valor médio de cada pedido ficou estável o ano inteiro (~R$ 55). Isso descarta a ideia de que o cliente passou a comprar mais e por isso \"sentiu\" mais o frete — o custo do frete é que aumentou sozinho, sem nenhum ganho percebido em troca.",
    { x: 0.6, y: 1.85, w: 6.2, h: 2.3, fontFace: FONT_BODY, fontSize: 14, color: MUTE, lineSpacing: 21, isTextBox: true });

  s.addShape("roundRect", { x: 0.6, y: 4.25, w: 6.2, h: 1.95, rectRadius: 0.1, fill: { color: DARK }, line: { type: "none" } });
  s.addText("Em Jan–Abr, o frete pesava ~21% do valor do pedido.\nEm Mai–Dez, passou a representar ~46% — quase metade da compra.",
    { x: 0.95, y: 4.45, w: 5.6, h: 1.55, fontFace: FONT_BODY, fontSize: 14.5, color: LIGHT, lineSpacing: 21, valign: "middle", isTextBox: true });

  s.addChart(
    pres.charts.BAR,
    [ { name: "Frete como % do ticket médio", labels: months, values: freteTicketPct } ],
    {
      x: 7.1, y: 1.7, w: 5.65, h: 5.0,
      chartColors: months.map((_,i)=> i<4 ? "B9D9D3" : AMBER),
      showTitle: true, title: "Frete ÷ ticket médio, por mês (%)", titleFontFace: FONT_BODY, titleFontSize: 13, titleColor: INK,
      showLegend: false,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 10, dataLabelColor: INK, dataLabelFormatCode: '0"%"',
      catAxisLabelColor: MUTE, catAxisLabelFontSize: 10,
      valAxisHidden: true, valGridLine: { style: "none" }, valAxisMaxVal: 55,
    }
  );

  footer(s, 5, false);
}

// ============================================================
// SLIDE 6 — RECOMENDAÇÃO CENÁRIO 1
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  kicker(s, "Cenário 1 · O que fazer", TEAL);
  s.addText("Recomendação: tratar o frete como causa raiz do abandono", { x: 0.6, y: 0.85, w: 11.8, h: 0.8,
    fontFace: FONT_HEAD, bold: true, fontSize: 27, color: INK, isTextBox: true });

  const recs = [
    ["Revisar a tabela de frete", "Entender por que o custo médio saltou de ~R$12 para ~R$26 a partir de maio, e se há espaço de negociação com transportadoras."],
    ["Testar frete grátis / com desconto acima de um valor mínimo", "Reduzir o choque no checkout para pedidos de ticket médio, sem abrir mão de margem em todos os pedidos."],
    ["Comunicar o frete mais cedo na jornada", "Mostrar o custo estimado antes do checkout final reduz a sensação de \"surpresa\" que costuma gerar abandono."],
    ["Acompanhar o frete como métrica de negócio", "Colocar o frete médio e sua relação com o ticket lado a lado com a taxa de abandono nos relatórios recorrentes."],
  ];
  let ry = 1.9;
  recs.forEach(([t, d], i) => {
    const rowH = 1.05;
    s.addShape("ellipse", { x: 0.6, y: ry+0.05, w: 0.5, h: 0.5, fill: { color: TEAL }, line: { type: "none" } });
    s.addText(String(i+1), { x: 0.6, y: ry+0.05, w: 0.5, h: 0.5, align: "center", valign: "middle",
      fontFace: FONT_BODY, bold: true, fontSize: 16, color: LIGHT });
    s.addText(t, { x: 1.35, y: ry, w: 10.8, h: 0.4, fontFace: FONT_BODY, bold: true, fontSize: 15, color: INK, isTextBox: true });
    s.addText(d, { x: 1.35, y: ry+0.4, w: 10.8, h: 0.55, fontFace: FONT_BODY, fontSize: 12.5, color: MUTE, lineSpacing: 17, isTextBox: true });
    ry += rowH;
  });

  footer(s, 6, false);
}

// ============================================================
// SLIDE 7 — TRANSIÇÃO / PONTE PARA O MODELO
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };
  s.addShape("ellipse", { x: -1.2, y: 4.2, w: 4, h: 4, fill: { color: "0F4A4B" }, line: { type: "none" } });

  s.addText("MAS A REVISÃO DE FRETE LEVA TEMPO.", { x: 0.7, y: 2.5, w: 10, h: 0.5,
    fontFace: FONT_BODY, bold: true, fontSize: 15, color: TEAL, charSpacing: 1.5 });
  s.addText("E os clientes de hoje? Como agir agora,\ncliente a cliente, enquanto isso?", { x: 0.65, y: 3.0, w: 11.5, h: 2.0,
    fontFace: FONT_HEAD, bold: true, fontSize: 34, color: LIGHT, lineSpacing: 42, isTextBox: true });

  s.addText("É aqui que entra o modelo preditivo desenvolvido pelo time de dados.", { x: 0.7, y: 5.1, w: 9, h: 0.5,
    fontFace: FONT_BODY, fontSize: 16, color: "CFE8E4" });

  footer(s, 7, true);
}

// ============================================================
// SLIDE 8 — O QUE O MODELO FAZ (linguagem simples)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  kicker(s, "Cenário 2 · O modelo", TEAL);
  s.addText("Um \"radar\" que avisa, em tempo real, quem está prestes a desistir da compra", { x: 0.6, y: 0.85, w: 11.8, h: 1.1,
    fontFace: FONT_HEAD, bold: true, fontSize: 25, color: INK, lineSpacing: 30, isTextBox: true });

  s.addShape("roundRect", { x: 0.6, y: 2.15, w: 7.1, h: 4.55, rectRadius: 0.12, fill: { color: PANEL }, line: { type: "none" } });
  s.addText("A ANALOGIA", { x: 0.95, y: 2.4, w: 6, h: 0.35, fontFace: FONT_BODY, bold: true, fontSize: 12, color: TEAL, charSpacing: 1.5 });
  s.addText("É como um vendedor experiente de loja física, que percebe quando um cliente está hesitando na fila do caixa e se antecipa antes que ele desista da compra.",
    { x: 0.95, y: 2.8, w: 6.4, h: 1.3, fontFace: FONT_BODY, italic: true, fontSize: 15.5, color: INK, lineSpacing: 22, isTextBox: true });
  s.addText("A diferença é a escala: o modelo faz isso para milhares de clientes ao mesmo tempo, 24 horas por dia, observando o comportamento de cada um no site (itens no carrinho, tempo de navegação, histórico de compras) para estimar a probabilidade de abandono antes que ele aconteça.",
    { x: 0.95, y: 4.25, w: 6.4, h: 2.3, fontFace: FONT_BODY, fontSize: 14, color: MUTE, lineSpacing: 21, isTextBox: true });

  s.addShape("roundRect", { x: 8.0, y: 2.15, w: 4.75, h: 2.15, rectRadius: 0.12, fill: { color: DARK }, line: { type: "none" } });
  s.addText("O QUE ELE ESTIMA", { x: 8.3, y: 2.4, w: 4, h: 0.35, fontFace: FONT_BODY, bold: true, fontSize: 11.5, color: TEAL, charSpacing: 1.5 });
  s.addText("A probabilidade de cada cliente abandonar o carrinho antes de finalizar a compra.",
    { x: 8.3, y: 2.8, w: 4.15, h: 1.4, fontFace: FONT_BODY, fontSize: 14.5, color: LIGHT, lineSpacing: 20, isTextBox: true });

  s.addShape("roundRect", { x: 8.0, y: 4.5, w: 4.75, h: 2.2, rectRadius: 0.12, fill: { color: TEAL }, line: { type: "none" } });
  s.addText("POR QUE ISSO IMPORTA", { x: 8.3, y: 4.75, w: 4, h: 0.35, fontFace: FONT_BODY, bold: true, fontSize: 11.5, color: DARK, charSpacing: 1.5 });
  s.addText("Em vez de agir depois que o cliente já saiu, conseguimos agir no momento exato da decisão.",
    { x: 8.3, y: 5.15, w: 4.15, h: 1.4, fontFace: FONT_BODY, fontSize: 14.5, color: LIGHT, lineSpacing: 20, isTextBox: true });

  footer(s, 8, false);
}

// ============================================================
// SLIDE 9 — FLUXO DA ESTRATÉGIA
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  kicker(s, "Cenário 2 · A estratégia", TEAL);
  s.addText("Como o cupom entra em ação — só onde faz diferença", { x: 0.6, y: 0.85, w: 11.8, h: 0.7,
    fontFace: FONT_HEAD, bold: true, fontSize: 26, color: INK, isTextBox: true });

  s.addText("O cupom não é enviado para todo mundo — isso encareceria a operação sem necessidade. Ele é acionado apenas no caso mais valioso: quando o modelo não esperava aquele abandono.",
    { x: 0.6, y: 1.65, w: 11.8, h: 0.6, fontFace: FONT_BODY, fontSize: 13.5, color: MUTE, lineSpacing: 18, isTextBox: true });

  const steps = [
    ["1", "Cliente navega e chega ao carrinho", TEAL],
    ["2", "O modelo classifica o cliente como baixa propensão ao abandono", TEAL],
    ["3", "Mesmo assim, o cliente sinaliza que vai abandonar a compra", RED],
    ["4", "O sistema aciona um alerta automático de \"caso inesperado\"", AMBER],
    ["5", "Marketing envia o cupom de desconto na hora", TEAL],
  ];
  const sw = 2.25, sh = 2.35, gap = 0.14;
  let sx = 0.6;
  steps.forEach(([n, t, color], i) => {
    s.addShape("roundRect", { x: sx, y: 2.55, w: sw, h: sh, rectRadius: 0.1, fill: { color: PANEL }, line: { type: "none" } });
    s.addShape("ellipse", { x: sx+0.22, y: 2.77, w: 0.5, h: 0.5, fill: { color: color }, line: { type: "none" } });
    s.addText(n, { x: sx+0.22, y: 2.77, w: 0.5, h: 0.5, align: "center", valign: "middle",
      fontFace: FONT_BODY, bold: true, fontSize: 16, color: LIGHT });
    s.addText(t, { x: sx+0.22, y: 3.42, w: sw-0.44, h: 1.35, fontFace: FONT_BODY, fontSize: 11.5,
      color: INK, lineSpacing: 15, isTextBox: true });
    if (i < steps.length - 1) {
      s.addText("→", { x: sx+sw, y: 3.35, w: gap+0.12, h: 0.5, align: "center", fontFace: FONT_BODY,
        bold: true, fontSize: 16, color: MUTE });
    }
    sx += sw + gap + 0.12;
  });

  s.addText("Resultado: o desconto é usado como uma ferramenta cirúrgica — para recuperar vendas que, sem o alerta, estariam perdidas — e não como um desconto genérico para toda a base.",
    { x: 0.6, y: 5.3, w: 11.8, h: 0.9, fontFace: FONT_BODY, italic: true, fontSize: 13.5, color: MUTE, lineSpacing: 19, isTextBox: true });

  footer(s, 9, false);
}

// ============================================================
// SLIDE 10 — TRADUZINDO ACURÁCIA E RECALL
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  kicker(s, "Cenário 2 · Os números, sem jargão", TEAL);
  s.addText("O que \"90% de acurácia\" e \"85% de recall\" significam na prática", { x: 0.6, y: 0.85, w: 11.8, h: 1.0,
    fontFace: FONT_HEAD, bold: true, fontSize: 24.5, color: INK, lineSpacing: 29, isTextBox: true });

  // card acuracia
  s.addShape("roundRect", { x: 0.6, y: 2.1, w: 5.75, h: 4.5, rectRadius: 0.12, fill: { color: PANEL }, line: { type: "none" } });
  s.addText("ACURÁCIA", { x: 0.95, y: 2.35, w: 5, h: 0.35, fontFace: FONT_BODY, bold: true, fontSize: 12, color: TEAL, charSpacing: 1.5 });
  s.addText("90%", { x: 0.95, y: 2.65, w: 3, h: 1.0, fontFace: FONT_HEAD, bold: true, fontSize: 46, color: DARK });
  s.addText("A cada 100 previsões que o modelo faz sobre o comportamento dos clientes, 90 estão corretas.",
    { x: 0.95, y: 3.75, w: 5.1, h: 1.0, fontFace: FONT_BODY, fontSize: 14, color: INK, lineSpacing: 19, isTextBox: true });
  s.addText("É a nota geral de confiabilidade do modelo — o quão bem ele acerta, olhando para todas as previsões, de abandono ou não.",
    { x: 0.95, y: 4.85, w: 5.1, h: 1.5, fontFace: FONT_BODY, fontSize: 12.5, color: MUTE, lineSpacing: 18, isTextBox: true });

  // card recall
  s.addShape("roundRect", { x: 6.65, y: 2.1, w: 6.05, h: 4.5, rectRadius: 0.12, fill: { color: DARK }, line: { type: "none" } });
  s.addText("RECALL", { x: 7.0, y: 2.35, w: 5, h: 0.35, fontFace: FONT_BODY, bold: true, fontSize: 12, color: TEAL, charSpacing: 1.5 });
  s.addText("85%", { x: 7.0, y: 2.65, w: 3, h: 1.0, fontFace: FONT_HEAD, bold: true, fontSize: 46, color: LIGHT });
  s.addText("A cada 100 clientes que realmente vão abandonar o carrinho, o modelo consegue identificar 85 deles antes que isso aconteça.",
    { x: 7.0, y: 3.75, w: 5.4, h: 1.1, fontFace: FONT_BODY, fontSize: 14, color: LIGHT, lineSpacing: 19, isTextBox: true });
  s.addText("Para a nossa estratégia, esse é o número mais importante: quanto maior o recall, menos clientes em risco passam despercebidos — e mais cupons chegam a quem precisa, a tempo.",
    { x: 7.0, y: 4.95, w: 5.4, h: 1.5, fontFace: FONT_BODY, fontSize: 12.5, color: "CFE8E4", lineSpacing: 18, isTextBox: true });

  footer(s, 10, false);
}

// ============================================================
// SLIDE 11 — BENEFÍCIOS ESPERADOS
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  kicker(s, "Cenário 2 · O que esperamos ganhar", TEAL);
  s.addText("Benefícios esperados para o negócio", { x: 0.6, y: 0.85, w: 11.8, h: 0.7,
    fontFace: FONT_HEAD, bold: true, fontSize: 28, color: INK });

  const bens = [
    ["Recuperação de vendas", "Parte das vendas que seriam perdidas passa a ser recuperada no exato momento da hesitação do cliente.", TEAL],
    ["Uso eficiente do orçamento", "O cupom só é usado onde há real necessidade — não como desconto generalizado para toda a base.", AMBER],
    ["Ação no momento certo", "A intervenção acontece durante a jornada de compra, não depois que o cliente já foi embora.", RED],
    ["Mais dados para decidir", "Cada ciclo do modelo gera novas evidências para calibrar a estratégia e apoiar decisões como a do frete.", GRAY],
  ];
  const cw = 2.87, gap = 0.22;
  let bx = 0.6;
  bens.forEach(([t, d, color]) => {
    s.addShape("roundRect", { x: bx, y: 1.85, w: cw, h: 4.7, rectRadius: 0.1, fill: { color: PANEL }, line: { type: "none" } });
    s.addShape("roundRect", { x: bx+0.3, y: 2.15, w: 0.55, h: 0.1, fill: { color: color }, line: { type: "none" } });
    s.addText(t, { x: bx+0.3, y: 2.4, w: cw-0.6, h: 0.9, fontFace: FONT_BODY, bold: true, fontSize: 15, color: INK, lineSpacing: 19, isTextBox: true });
    s.addText(d, { x: bx+0.3, y: 3.35, w: cw-0.6, h: 3.0, fontFace: FONT_BODY, fontSize: 12.5, color: MUTE, lineSpacing: 18, isTextBox: true });
    bx += cw + gap;
  });

  footer(s, 11, false);
}

// ============================================================
// SLIDE 12 — CONCLUSÃO / PRÓXIMOS PASSOS
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };
  s.addShape("ellipse", { x: 10.6, y: -1.2, w: 4, h: 4, fill: { color: "0F4A4B" }, line: { type: "none" } });

  s.addText("EM RESUMO", { x: 0.7, y: 0.75, w: 6, h: 0.4, fontFace: FONT_BODY, bold: true, fontSize: 13, color: TEAL, charSpacing: 2 });
  s.addText("Duas frentes, um mesmo objetivo: recuperar a compra", { x: 0.65, y: 1.2, w: 11.5, h: 1.1,
    fontFace: FONT_HEAD, bold: true, fontSize: 28, color: LIGHT, lineSpacing: 34, isTextBox: true });

  s.addShape("roundRect", { x: 0.7, y: 2.6, w: 5.7, h: 2.7, rectRadius: 0.12, fill: { color: "0F4A4B" }, line: { type: "none" } });
  s.addText("ESTRUTURAL", { x: 1.0, y: 2.85, w: 5, h: 0.35, fontFace: FONT_BODY, bold: true, fontSize: 11.5, color: TEAL, charSpacing: 1.5 });
  s.addText("Revisar a política de frete", { x: 1.0, y: 3.2, w: 5.1, h: 0.5, fontFace: FONT_BODY, bold: true, fontSize: 16, color: LIGHT });
  s.addText("Atacar a causa raiz identificada no gráfico: o frete que dobrou e passou a pesar quase metade do valor da compra.",
    { x: 1.0, y: 3.7, w: 5.1, h: 1.4, fontFace: FONT_BODY, fontSize: 13, color: "CFE8E4", lineSpacing: 18, isTextBox: true });

  s.addShape("roundRect", { x: 6.7, y: 2.6, w: 5.7, h: 2.7, rectRadius: 0.12, fill: { color: TEAL }, line: { type: "none" } });
  s.addText("TÁTICA", { x: 7.0, y: 2.85, w: 5, h: 0.35, fontFace: FONT_BODY, bold: true, fontSize: 11.5, color: DARK, charSpacing: 1.5 });
  s.addText("Ativar o modelo preditivo + cupom", { x: 7.0, y: 3.2, w: 5.1, h: 0.5, fontFace: FONT_BODY, bold: true, fontSize: 16, color: DARK });
  s.addText("Agir agora, cliente a cliente, enquanto a mudança estrutural do frete é avaliada e implementada.",
    { x: 7.0, y: 3.7, w: 5.1, h: 1.4, fontFace: FONT_BODY, fontSize: 13, color: "0B3D3E", lineSpacing: 18, isTextBox: true });

  s.addText("Próximo passo: iniciar um piloto do modelo com envio automático de cupom, em paralelo à revisão da tabela de frete.",
    { x: 0.7, y: 5.65, w: 11.4, h: 0.7, fontFace: FONT_BODY, italic: true, fontSize: 14, color: "CFE8E4", lineSpacing: 19, isTextBox: true });

  footer(s, 12, true);
}

pres.writeFile({ fileName: "/home/claude/abandono-carrinho.pptx" }).then(() => {
  console.log("done");
});
