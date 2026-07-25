/* Efes Nexus — business case deck (PPTX)
   Faithful to presentation.html, minus live-demo & risk slides,
   money -> working-hours efficiency, with volume-driven automation growth. */
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa6");
const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, "..", "..");   // repository root
const BADGE = path.join(ROOT, "assets/efes-badge.png");
const OUT = path.join(ROOT, "deck", "base", "Efes_Nexus_RU_base.pptx");
const TMP = path.join(__dirname, "gen");        // scratch for rendered icons (gitignored)
if (!fs.existsSync(TMP)) fs.mkdirSync(TMP, { recursive: true });

const FONT = "Calibri", MONO = "Consolas";

const C = {
  brand:'004C8D', brand700:'005BA8', brand600:'1166A8', brand500:'2D83C4',
  brand400:'5CA3DB', brand300:'9CC6E8', brand100:'E3EEF8', brand50:'F1F6FB',
  ink:'0A1018', ink1:'0E1726', ink2:'13203A', inkCard:'16223B', inkLine:'28364E',
  paper:'EEF1F5', paper2:'F6F8FA', cardW:'FFFFFF', line:'E2E7EE', line2:'EEF1F4',
  t1:'0C1726', t2:'56616F', t3:'8B94A2',
  tw1:'EEF2F7', tw2:'9AA9BF', tw3:'73839B',
  good:'15924F', goodBg:'E7F5EC', goodLine:'BBE3C8', goodDk:'2FB46E',
  warn:'BD7D12', warnBg:'FBF1DD', warnLine:'EED9A7',
  bad:'C33A30', badBg:'FBE9E7', badLine:'F0C5C0',
  accent:'D97757',
};

// ---------- icons ----------
const COMP = {
  clock: FA.FaRegClock, users: FA.FaUsers, bolt: FA.FaBolt, check: FA.FaCheck,
  table: FA.FaTableCells, envelope: FA.FaEnvelope, calc: FA.FaCalculator,
  userclock: FA.FaUserClock, box: FA.FaBoxOpen, gift: FA.FaGift, coins: FA.FaCoins,
  layers: FA.FaLayerGroup, diagram: FA.FaDiagramProject, robot: FA.FaRobot,
  bell: FA.FaBell, wand: FA.FaWandMagicSparkles, chart: FA.FaChartColumn,
  trend: FA.FaArrowTrendUp, circleCheck: FA.FaCircleCheck, gauge: FA.FaGaugeHigh,
  sitemap: FA.FaSitemap, scale: FA.FaScaleBalanced, boxes: FA.FaBoxesStacked,
  store: FA.FaStore, handshake: FA.FaHandshake,
};
const _cache = {};
async function ic(name, color) {
  const key = name + "_" + color;
  if (_cache[key]) return _cache[key];
  if (!COMP[name]) throw new Error("missing icon " + name);
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(COMP[name], { color: "#" + color, size: "256" }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  const d = "image/png;base64," + png.toString("base64");
  _cache[key] = d; return d;
}

// ---------- background generation ----------
async function bg(name, svg) {
  const p = path.join(TMP, name + ".png");
  await sharp(Buffer.from(svg)).png().toFile(p);
  return p;
}
const W = 2000, H = 1125;
async function makeBackgrounds() {
  const title = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <defs>
      <radialGradient id="b" cx="82%" cy="8%" r="55%">
        <stop offset="0%" stop-color="#1166a8" stop-opacity="0.42"/>
        <stop offset="55%" stop-color="#0e3a63" stop-opacity="0.12"/>
        <stop offset="100%" stop-color="#0A1018" stop-opacity="0"/>
      </radialGradient>
      <radialGradient id="c" cx="10%" cy="96%" r="48%">
        <stop offset="0%" stop-color="#d97757" stop-opacity="0.22"/>
        <stop offset="100%" stop-color="#0A1018" stop-opacity="0"/>
      </radialGradient>
    </defs>
    <rect width="100%" height="100%" fill="#0A1018"/>
    <rect width="100%" height="100%" fill="url(#b)"/>
    <rect width="100%" height="100%" fill="url(#c)"/></svg>`;
  const dark = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <defs><radialGradient id="b" cx="88%" cy="6%" r="60%">
      <stop offset="0%" stop-color="#13427a" stop-opacity="0.40"/>
      <stop offset="60%" stop-color="#0d2440" stop-opacity="0.10"/>
      <stop offset="100%" stop-color="#0A1018" stop-opacity="0"/>
    </radialGradient></defs>
    <rect width="100%" height="100%" fill="#0A1018"/>
    <rect width="100%" height="100%" fill="url(#b)"/></svg>`;
  const ink = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <defs><radialGradient id="g" cx="14%" cy="90%" r="60%">
      <stop offset="0%" stop-color="#0f5a39" stop-opacity="0.26"/>
      <stop offset="100%" stop-color="#0B1220" stop-opacity="0"/>
    </radialGradient>
    <radialGradient id="b" cx="92%" cy="10%" r="55%">
      <stop offset="0%" stop-color="#123f6e" stop-opacity="0.30"/>
      <stop offset="100%" stop-color="#0B1220" stop-opacity="0"/>
    </radialGradient></defs>
    <rect width="100%" height="100%" fill="#0B1220"/>
    <rect width="100%" height="100%" fill="url(#b)"/>
    <rect width="100%" height="100%" fill="url(#g)"/></svg>`;
  const brand = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <defs><linearGradient id="g" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#013a73"/>
      <stop offset="55%" stop-color="#0a4f96"/>
      <stop offset="100%" stop-color="#1166a8"/>
    </linearGradient>
    <radialGradient id="h" cx="85%" cy="12%" r="50%">
      <stop offset="0%" stop-color="#2d83c4" stop-opacity="0.35"/>
      <stop offset="100%" stop-color="#0a4f96" stop-opacity="0"/>
    </radialGradient></defs>
    <rect width="100%" height="100%" fill="url(#g)"/>
    <rect width="100%" height="100%" fill="url(#h)"/></svg>`;
  return {
    title: await bg("bg_title", title),
    dark: await bg("bg_dark", dark),
    ink: await bg("bg_ink", ink),
    brand: await bg("bg_brand", brand),
  };
}

// ================= main =================
(async () => {
  const BG = await makeBackgrounds();
  const pres = new pptxgen();
  pres.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
  pres.layout = "WIDE";
  pres.author = "Efes Kazakhstan";
  pres.title = "Efes Nexus — бизнес-кейс";

  const MX = 0.62, CW = 13.333 - MX * 2; // content width 12.09
  const shadow = () => ({ type: "outer", color: "0C1726", blur: 7, offset: 2, angle: 135, opacity: 0.08 });

  // ---- helpers ----
  function footer(s, dark, right) {
    s.addShape(pres.shapes.LINE, { x: MX, y: 6.96, w: CW, h: 0, line: { color: dark ? C.inkLine : C.line, width: 1 } });
    s.addImage({ path: BADGE, x: MX, y: 7.04, w: 0.24, h: 0.24, transparency: dark ? 35 : 0 });
    s.addText("Efes Kazakhstan", { x: MX + 0.32, y: 7.0, w: 5, h: 0.32, fontFace: FONT, fontSize: 8.5, color: dark ? C.tw3 : C.t3, valign: "middle" });
    s.addText(right, { x: 7.5, y: 7.0, w: CW - 6.88, h: 0.32, fontFace: FONT, fontSize: 8.5, color: dark ? C.tw3 : C.t3, align: "right", valign: "middle" });
  }
  function kicker(s, x, y, text, dark) {
    s.addShape(pres.shapes.RECTANGLE, { x, y: y + 0.085, w: 0.17, h: 0.032, fill: { color: dark ? C.brand400 : C.brand } });
    s.addText(text.toUpperCase(), { x: x + 0.25, y: y - 0.05, w: 9, h: 0.3, fontFace: FONT, fontSize: 10.5, bold: true, charSpacing: 2, color: dark ? C.brand300 : C.brand600, valign: "middle" });
  }
  function h2(s, x, y, text, dark, size) {
    s.addText(text, { x, y, w: CW, h: 0.66, fontFace: FONT, fontSize: size || 29, bold: true, color: dark ? C.tw1 : C.t1, valign: "middle" });
  }
  function sub(s, x, y, w, text, dark) {
    s.addText(text, { x, y, w, h: 0.8, fontFace: FONT, fontSize: 12.5, color: dark ? C.tw2 : C.t2, valign: "top", lineSpacingMultiple: 1.15 });
  }
  function card(s, x, y, w, h, dark, o = {}) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w, h, rectRadius: 0.09,
      fill: { color: o.fill || (dark ? C.inkCard : C.cardW) },
      line: { color: o.line || (dark ? C.inkLine : C.line), width: 1 },
      shadow: (dark || o.noShadow) ? undefined : shadow(),
    });
  }
  const TINT = {
    blue:  { l: [C.brand100, C.brand500], d: ["14304F", C.brand400] },
    red:   { l: [C.badBg, C.bad],         d: ["2C1A1A", "E07A72"] },
    amber: { l: [C.warnBg, C.warn],       d: ["2C2310", "E0A64B"] },
    green: { l: [C.goodBg, C.good],       d: ["10301F", "3FB572"] },
  };
  async function iconChip(s, x, y, d, name, variant, dark, brandSlide) {
    let bgc, fgc;
    if (brandSlide) { bgc = "1E5C99"; fgc = "FFFFFF"; }
    else { const t = TINT[variant][dark ? "d" : "l"]; bgc = t[0]; fgc = t[1]; }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: d, h: d, rectRadius: d * 0.27, fill: { color: bgc } });
    s.addImage({ data: await ic(name, fgc), x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52 });
  }
  function pill(s, x, y, text, variant, w, dark) {
    const map = { brand: [C.brand100, C.brand, C.brand300], good: [C.goodBg, C.good, C.goodLine], warn: [C.warnBg, C.warn, C.warnLine], bad: [C.badBg, C.bad, C.badLine] };
    let [bg, fg, ln] = map[variant] || map.brand;
    if (dark && variant === "good") { bg = "10301F"; fg = C.goodDk; ln = "1D4A32"; }
    if (dark && variant === "brand") { bg = "14304F"; fg = C.brand300; ln = "1E4974"; }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.3, rectRadius: 0.15, fill: { color: bg }, line: { color: ln, width: 1 } });
    s.addText(text, { x: x, y: y, w: w, h: 0.3, fontFace: FONT, fontSize: 10, bold: true, color: fg, align: "center", valign: "middle", margin: 0 });
  }

  // ============ SLIDE 1 — TITLE ============
  {
    const s = pres.addSlide(); s.background = { path: BG.title };
    s.addImage({ path: BADGE, x: MX, y: 0.62, w: 0.6, h: 0.6 });
    s.addShape(pres.shapes.LINE, { x: MX + 0.78, y: 0.66, w: 0, h: 0.52, line: { color: "33445E", width: 1 } });
    s.addText([
      { text: "EFES KAZAKHSTAN", options: { bold: true, color: C.tw1, fontSize: 11, charSpacing: 1, breakLine: true } },
      { text: "Отдел продаж", options: { color: C.tw3, fontSize: 10 } },
    ], { x: MX + 0.94, y: 0.6, w: 6, h: 0.66, fontFace: FONT, valign: "middle", lineSpacingMultiple: 1.05 });

    kicker(s, MX, 2.32, "Бизнес-кейс · автоматизация", true);
    s.addText("Efes Nexus", { x: MX - 0.02, y: 2.62, w: 11, h: 1.1, fontFace: FONT, fontSize: 60, bold: true, color: "FFFFFF" });
    s.addText("Менеджер промо-активностей", { x: MX, y: 3.78, w: 11, h: 0.6, fontFace: FONT, fontSize: 27, bold: true, color: C.brand300 });
    s.addText("Единая платформа, которая заменяет ручную работу в Excel и Outlook сквозными цифровыми процессами, авто-уведомлениями и AI-проверкой запросов партнёров.",
      { x: MX, y: 4.62, w: 8.2, h: 1.0, fontFace: FONT, fontSize: 14.5, color: C.tw2, lineSpacingMultiple: 1.3 });

    const tags = ["Промо-активности", "3 бизнес-процесса"];
    let tx = MX;
    for (const t of tags) {
      const w = 0.42 + t.length * 0.092;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: tx, y: 5.78, w, h: 0.42, rectRadius: 0.21, fill: { color: "13203A" }, line: { color: "2A3A56", width: 1 } });
      s.addShape(pres.shapes.OVAL, { x: tx + 0.18, y: 5.96, w: 0.07, h: 0.07, fill: { color: C.brand400 } });
      s.addText(t, { x: tx + 0.32, y: 5.78, w: w - 0.36, h: 0.42, fontFace: FONT, fontSize: 11, color: C.tw1, valign: "middle", margin: 0 });
      tx += w + 0.18;
    }
    footer(s, true, "Efes Nexus · бизнес-кейс для руководителя продаж");
  }

  // ============ SLIDE 2 — SCOPE ============
  {
    const s = pres.addSlide(); s.background = { color: C.paper };
    kicker(s, MX, 0.56, "Охват", false);
    h2(s, MX, 0.86, "Что охватывает система", false);
    sub(s, MX, 1.56, 11.4, "Один продукт закрывает полный жизненный цикл промо-активностей — от создания и активации до закрытия месяца с AI-проверкой запросов.", false);

    // two promo-type cards
    const cy = 2.5, ch = 2.35, cw = (CW - 0.4) / 2;
    const types = [
      { x: MX, code: "N + 1", icon: "gift", title: "Подарочные литры",
        body: "За каждые N литров — 1 литр в подарок.", chan: "On-trade + Off-trade" },
      { x: MX + cw + 0.4, code: "N + KZT", icon: "coins", title: "Денежная компенсация",
        body: "Компенсация на достигнутый объём продаж.", chan: "Только Off-trade" },
    ];
    for (const t of types) {
      card(s, t.x, cy, cw, ch, false);
      await iconChip(s, t.x + 0.34, cy + 0.34, 0.62, t.icon, "blue", false);
      s.addText(t.title, { x: t.x + 1.12, y: cy + 0.36, w: cw - 2.6, h: 0.6, fontFace: FONT, fontSize: 16, bold: true, color: C.t1, valign: "middle" });
      pill(s, t.x + cw - 1.36, cy + 0.42, t.code, "brand", 1.0, false);
      s.addText(t.body, { x: t.x + 0.36, y: cy + 1.18, w: cw - 0.72, h: 0.5, fontFace: FONT, fontSize: 12.5, color: C.t2, valign: "top" });
      s.addShape(pres.shapes.LINE, { x: t.x + 0.36, y: cy + 1.78, w: cw - 0.72, h: 0, line: { color: C.line2, width: 1 } });
      s.addImage({ data: await ic("store", C.brand500), x: t.x + 0.36, y: cy + 1.92, w: 0.18, h: 0.18 });
      s.addText(t.chan, { x: t.x + 0.62, y: cy + 1.86, w: cw - 1, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: C.brand600, valign: "middle" });
    }

    // lifecycle strip
    const ly = 5.3;
    card(s, MX, ly, CW, 1.18, false, { fill: C.brand50, line: C.brand100 });
    s.addText("Полный жизненный цикл — с AI-проверкой запросов партнёров", { x: MX + 0.36, y: ly + 0.16, w: CW - 0.72, h: 0.34, fontFace: FONT, fontSize: 12, bold: true, color: C.brand700 });
    const steps = ["Создание", "Активация", "Изменение", "Закрытие месяца"];
    const sw = 2.2, gap = ((CW - 0.72) - sw * steps.length) / (steps.length - 1);
    let sx = MX + 0.36;
    for (let i = 0; i < steps.length; i++) {
      const last = i === steps.length - 1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: sx, y: ly + 0.62, w: sw, h: 0.42, rectRadius: 0.08, fill: { color: last ? C.brand : "FFFFFF" }, line: { color: last ? C.brand : C.brand100, width: 1 } });
      s.addText(steps[i], { x: sx, y: ly + 0.62, w: sw, h: 0.42, fontFace: FONT, fontSize: 11.5, bold: true, color: last ? "FFFFFF" : C.t1, align: "center", valign: "middle", margin: 0 });
      if (!last) s.addText("→", { x: sx + sw, y: ly + 0.62, w: gap, h: 0.42, fontFace: FONT, fontSize: 15, bold: true, color: C.brand400, align: "center", valign: "middle", margin: 0 });
      sx += sw + gap;
    }
    footer(s, false, "Efes Nexus");
  }

  // ============ SLIDE 3 — PROBLEM ============
  {
    const s = pres.addSlide(); s.background = { path: BG.dark };
    kicker(s, MX, 0.56, "Проблема", true);
    h2(s, MX, 0.86, "Сегодня всё делается вручную", true);
    const items = [
      { v: "red", icon: "table", t: "Десятки несвязанных файлов Excel", b: "Нет единого источника данных — версии конфликтуют." },
      { v: "amber", icon: "envelope", t: "Коммуникация в Outlook и мессенджерах", b: "Письма теряются, согласования задерживаются." },
      { v: "blue", icon: "calc", t: "Ручное кодирование и расчёты", b: "Риск дублей, переплат и искажений в отчётах." },
      { v: "red", icon: "userclock", t: "Зависимость от человеческого фактора", b: "300+ промо в месяц — одна ошибка идёт каскадом." },
    ];
    const gx = 0.4, gy = 0.36, cw = (CW - gx) / 2, ch = 2.18, top = 1.9;
    for (let i = 0; i < 4; i++) {
      const x = MX + (i % 2) * (cw + gx), y = top + Math.floor(i / 2) * (ch + gy);
      card(s, x, y, cw, ch, true);
      await iconChip(s, x + 0.4, y + 0.4, 0.66, items[i].icon, items[i].v, true);
      s.addText(items[i].t, { x: x + 0.4, y: y + 1.18, w: cw - 0.8, h: 0.56, fontFace: FONT, fontSize: 16.5, bold: true, color: C.tw1, valign: "middle" });
      s.addText(items[i].b, { x: x + 0.4, y: y + 1.66, w: cw - 0.8, h: 0.42, fontFace: FONT, fontSize: 12, color: C.tw2, valign: "top" });
    }
    footer(s, true, "Efes Nexus");
  }

  // ============ SLIDE 4 — SCALE GROWTH ============
  {
    const s = pres.addSlide(); s.background = { color: C.paper2 };
    kicker(s, MX, 0.52, "Масштаб", false);
    h2(s, MX, 0.82, "Объём растёт быстрее, чем успевают руки", false);
    sub(s, MX, 1.52, 11.9, "За год число промо-активностей выросло в 4,6 раза при почти неизменном числе партнёров — нагрузка на каждого специалиста выросла кратно. Ручные процессы такого темпа не выдерживают.", false);

    // left chart card
    const lx = MX, ly = 2.55, lw = 6.0, lh = 3.92;
    card(s, lx, ly, lw, lh, false);
    s.addText("Промо-активности в год", { x: lx + 0.34, y: ly + 0.28, w: 3.6, h: 0.32, fontFace: FONT, fontSize: 14, bold: true, color: C.t1 });
    s.addText("Количество заведённых акций за год", { x: lx + 0.34, y: ly + 0.62, w: 3.6, h: 0.3, fontFace: FONT, fontSize: 10, color: C.t3 });
    pill(s, lx + lw - 1.78, ly + 0.32, "2024 → 2025", "brand", 1.42, false);
    // bars
    const baseY = ly + lh - 0.62, maxBarH = 2.35;
    const b24 = maxBarH * 2187 / 10112, b25 = maxBarH;
    const bw = 1.0;
    const x24 = lx + 0.95, x25 = lx + lw - 0.95 - bw;
    // 2024 bar
    s.addShape(pres.shapes.RECTANGLE, { x: x24, y: baseY - b24, w: bw, h: b24, fill: { color: "B7C0CC" } });
    s.addText("2 187", { x: x24 - 0.4, y: baseY - b24 - 0.34, w: bw + 0.8, h: 0.3, fontFace: MONO, fontSize: 13, bold: true, color: C.t2, align: "center" });
    s.addText("2024", { x: x24 - 0.4, y: baseY + 0.06, w: bw + 0.8, h: 0.28, fontFace: FONT, fontSize: 12, bold: true, color: C.t2, align: "center" });
    // 2025 bar
    s.addShape(pres.shapes.RECTANGLE, { x: x25, y: baseY - b25, w: bw, h: b25, fill: { color: C.brand } });
    s.addShape(pres.shapes.RECTANGLE, { x: x25, y: baseY - b25, w: bw, h: 0.5, fill: { color: C.brand500 } });
    s.addText("10 112", { x: x25 - 0.4, y: baseY - b25 - 0.34, w: bw + 0.8, h: 0.3, fontFace: MONO, fontSize: 14, bold: true, color: C.brand, align: "center" });
    s.addText("2025", { x: x25 - 0.4, y: baseY + 0.06, w: bw + 0.8, h: 0.28, fontFace: FONT, fontSize: 12, bold: true, color: C.brand, align: "center" });
    // center multiplier
    const mcx = (x24 + bw + x25) / 2;
    s.addText("×4,6", { x: mcx - 0.85, y: baseY - 1.78, w: 1.7, h: 0.6, fontFace: FONT, fontSize: 34, bold: true, color: C.brand, align: "center" });
    pill(s, mcx - 0.62, baseY - 1.16, "+362 %", "good", 1.24, false);

    // right metric cards
    const rx = MX + lw + 0.34, rw = CW - lw - 0.34;
    const metrics = [
      { lab: "Партнёры", rng: "за 6 мес", a: "31", b: "33", d: "+6,5 %" },
      { lab: "Продукты · SKU", rng: "за 6 мес", a: "78", b: "105", d: "+35 %" },
      { lab: "Торговые точки", rng: "2025 → 2026", a: "3 857", b: "5 410", d: "+40 %" },
      { lab: "Промо на специалиста / мес", rng: "2024 → 2025", a: "36", b: "169", d: "×4,6" },
    ];
    const mh = 0.86, mg = 0.155;
    for (let i = 0; i < 4; i++) {
      const y = 2.55 + i * (mh + mg);
      card(s, rx, y, rw, mh, false);
      s.addText(metrics[i].lab.toUpperCase(), { x: rx + 0.28, y: y + 0.12, w: rw - 1.6, h: 0.26, fontFace: FONT, fontSize: 9.5, bold: true, charSpacing: 1, color: C.t2, valign: "middle" });
      s.addText(metrics[i].rng, { x: rx + rw - 1.5, y: y + 0.12, w: 1.24, h: 0.26, fontFace: MONO, fontSize: 9, color: C.t3, align: "right", valign: "middle" });
      s.addText(metrics[i].a, { x: rx + 0.28, y: y + 0.4, w: 1.2, h: 0.36, fontFace: MONO, fontSize: 15, bold: true, color: C.t3, valign: "middle" });
      s.addText("→", { x: rx + 1.5, y: y + 0.4, w: 0.5, h: 0.36, fontFace: FONT, fontSize: 14, bold: true, color: C.brand400, align: "center", valign: "middle" });
      s.addText(metrics[i].b, { x: rx + 2.05, y: y + 0.36, w: 1.5, h: 0.4, fontFace: MONO, fontSize: 20, bold: true, color: C.brand, valign: "middle" });
      pill(s, rx + rw - 1.18, y + 0.42, metrics[i].d, "good", 0.92, false);
    }
    footer(s, false, "Efes Nexus");
  }

  // ============ SLIDE 5 — SOLUTION ============
  {
    const s = pres.addSlide(); s.background = { path: BG.brand };
    kicker(s, MX, 0.56, "Решение", true);
    s.addText("Единая платформа вместо Excel и почты", { x: MX, y: 0.86, w: CW, h: 0.66, fontFace: FONT, fontSize: 29, bold: true, color: "FFFFFF", valign: "middle" });
    s.addText("Веб-система превращает разрозненные таблицы и цепочки писем в структурированные процессы с автоматизацией и проверкой.",
      { x: MX, y: 1.56, w: 11.4, h: 0.5, fontFace: FONT, fontSize: 12.5, color: "DCE8F4", lineSpacingMultiple: 1.15 });

    const cards = [
      { icon: "layers", t: "Единый реестр", b: "Все промо-активности и компенсации — в одном источнике данных. Никаких конфликтов версий." },
      { icon: "sitemap", t: "Сквозные процессы", b: "3 бизнес-процесса оцифрованы целиком: создание, изменение, закрытие." },
      { icon: "wand", t: "AI-верификация", b: "Автоматическая сверка запросов партнёров с данными Panorama — мгновенно и без человека." },
      { icon: "bell", t: "Авто-уведомления", b: "Письма партнёрам и напоминания по срокам формируются и отправляются системой." },
    ];
    const gx = 0.34, cw = (CW - gx * 3) / 4, cy = 2.42, ch = 3.0;
    for (let i = 0; i < 4; i++) {
      const x = MX + i * (cw + gx);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy, w: cw, h: ch, rectRadius: 0.09, fill: { color: "0E5093" }, line: { color: "2E72AE", width: 1 } });
      await iconChip(s, x + 0.3, cy + 0.34, 0.64, cards[i].icon, "blue", false, true);
      s.addText(cards[i].t, { x: x + 0.3, y: cy + 1.18, w: cw - 0.6, h: 0.4, fontFace: FONT, fontSize: 15.5, bold: true, color: "FFFFFF", valign: "middle" });
      s.addText(cards[i].b, { x: x + 0.3, y: cy + 1.62, w: cw - 0.6, h: 1.2, fontFace: FONT, fontSize: 11.5, color: "CFE0F1", valign: "top", lineSpacingMultiple: 1.16 });
    }
    // result band
    const by = 5.62;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: by, w: CW, h: 0.78, rectRadius: 0.1, fill: { color: "0C3F77" }, line: { color: "2E72AE", width: 1 } });
    s.addImage({ data: await ic("circleCheck", "FFFFFF"), x: MX + 0.34, y: by + 0.24, w: 0.3, h: 0.3 });
    s.addText("Результат — прозрачность, полный аудит-трейл и устранение человеческого фактора в рутинных операциях.",
      { x: MX + 0.82, y: by, w: CW - 1.1, h: 0.78, fontFace: FONT, fontSize: 13.5, bold: true, color: "FFFFFF", valign: "middle" });
    footer(s, true, "Efes Nexus");
  }

  // ============ SLIDE 6 — THREE PROCESSES ============
  {
    const s = pres.addSlide(); s.background = { color: C.paper };
    kicker(s, MX, 0.56, "Архитектура процессов", false);
    h2(s, MX, 0.86, "Три бизнес-процесса — в одной системе", false);
    sub(s, MX, 1.56, 11.8, "Каждый процесс описан в нотации BPMN. Следующий слайд показывает, какие шаги берёт на себя система.", false);

    const procs = [
      { n: "01", tag: "Промо", t: "Создание промо-активности", b: "Запрос от Trade-маркетинга или On-trade → присвоение кода → запись в Panorama → уведомление партнёров." },
      { n: "02", tag: "Промо", t: "Изменение промо-активности", b: "Смена квоты, периода, добавление продукта или точки → правка в Panorama → синхронизация и уведомление." },
      { n: "03", tag: "Закрытие", t: "Закрытие месяца и расчёт", b: "Партнёр присылает «Запрос» → расчёт ∑(литры × цена периода) → AI-проверка → передача в Контроллинг." },
    ];
    const gx = 0.4, cw = (CW - gx * 2) / 3, cy = 2.45, ch = 2.78;
    for (let i = 0; i < 3; i++) {
      const x = MX + i * (cw + gx);
      card(s, x, cy, cw, ch, false);
      s.addText(procs[i].n, { x: x + 0.32, y: cy + 0.28, w: 1.2, h: 0.4, fontFace: MONO, fontSize: 17, bold: true, color: C.brand });
      pill(s, x + cw - 1.34, cy + 0.32, procs[i].tag, "brand", 1.02, false);
      s.addText(procs[i].t, { x: x + 0.32, y: cy + 0.82, w: cw - 0.64, h: 0.66, fontFace: FONT, fontSize: 15.5, bold: true, color: C.t1, valign: "top", lineSpacingMultiple: 1.05 });
      s.addText(procs[i].b, { x: x + 0.32, y: cy + 1.5, w: cw - 0.64, h: 1.1, fontFace: FONT, fontSize: 11.5, color: C.t2, valign: "top", lineSpacingMultiple: 1.18 });
    }
    // flow row
    const fy = 5.58;
    card(s, MX, fy, CW, 0.82, false, { fill: C.brand50, line: C.brand100 });
    const fsteps = [{ t: "Создание", b: false }, { t: "Активно & изменения", b: false }, { t: "Закрытие месяца", b: true }];
    const fw = 3.0, fgap = ((CW - 0.72) - fw * 3) / 2;
    let fx = MX + 0.36;
    for (let i = 0; i < 3; i++) {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: fx, y: fy + 0.2, w: fw, h: 0.42, rectRadius: 0.08, fill: { color: fsteps[i].b ? C.brand : "FFFFFF" }, line: { color: fsteps[i].b ? C.brand : C.brand100, width: 1 } });
      s.addText(fsteps[i].t, { x: fx, y: fy + 0.2, w: fw, h: 0.42, fontFace: FONT, fontSize: 12, bold: true, color: fsteps[i].b ? "FFFFFF" : C.t1, align: "center", valign: "middle", margin: 0 });
      if (i < 2) s.addText("→", { x: fx + fw, y: fy + 0.2, w: fgap, h: 0.42, fontFace: FONT, fontSize: 16, bold: true, color: C.brand400, align: "center", valign: "middle", margin: 0 });
      fx += fw + fgap;
    }
    footer(s, false, "Efes Nexus");
  }

  // ============ SLIDE 7 — BPMN LIFECYCLE ============
  {
    const s = pres.addSlide(); s.background = { color: C.paper2 };
    kicker(s, MX, 0.5, "BPMN · процессы 01–03", false);
    h2(s, MX, 0.8, "Жизненный цикл промо-активности", false, 27);
    s.addText([
      { text: "Шаги BPMN-диаграммы. ", options: {} },
      { text: "Синим", options: { color: C.brand, bold: true } },
      { text: " выделено то, что выполняет система; ", options: {} },
      { text: "серым", options: { color: C.t2, bold: true } },
      { text: " — действия и решения человека.", options: {} },
    ], { x: MX, y: 1.44, w: 11.8, h: 0.34, fontFace: FONT, fontSize: 12, color: C.t2, valign: "middle" });

    const arrowW = 0.34;
    const FNW = (CW - arrowW * 5) / 6; // fixed node width (6-up); shorter rows centered
    async function flow(label, y, nodes) {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: y + 0.02, w: 0.42, h: 0.26, rectRadius: 0.05, fill: { color: C.brand } });
      s.addText(label.n, { x: MX, y: y + 0.02, w: 0.42, h: 0.26, fontFace: MONO, fontSize: 11, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
      s.addText(label.t, { x: MX + 0.54, y: y, w: 8, h: 0.3, fontFace: FONT, fontSize: 12, bold: true, color: C.t1, valign: "middle" });
      const ny = y + 0.34, nh = 1.0, n = nodes.length;
      const rowW = n * FNW + (n - 1) * arrowW;
      let nx = MX + (CW - rowW) / 2;
      for (let i = 0; i < n; i++) {
        const nd = nodes[i], auto = nd.a;
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: FNW, h: nh, rectRadius: 0.07, fill: { color: auto ? C.brand100 : "FFFFFF" }, line: { color: auto ? C.brand300 : C.line, width: 1 }, shadow: shadow() });
        s.addShape(pres.shapes.RECTANGLE, { x: nx, y: ny + 0.1, w: 0.055, h: nh - 0.2, fill: { color: auto ? C.brand : "C2CAD4" } });
        s.addShape(pres.shapes.OVAL, { x: nx + 0.16, y: ny + 0.14, w: 0.25, h: 0.25, fill: { color: auto ? C.brand : "E5E9EF" } });
        s.addText(String(i + 1), { x: nx + 0.16, y: ny + 0.14, w: 0.25, h: 0.25, fontFace: MONO, fontSize: 9.5, bold: true, color: auto ? "FFFFFF" : C.t2, align: "center", valign: "middle", margin: 0 });
        s.addText(nd.role, { x: nx + 0.48, y: ny + 0.14, w: FNW - 0.6, h: 0.24, fontFace: FONT, fontSize: 8.5, bold: true, charSpacing: 0.3, color: auto ? C.brand600 : C.t2, valign: "middle", margin: 0 });
        s.addText(nd.name, { x: nx + 0.2, y: ny + 0.40, w: FNW - 0.36, h: 0.36, fontFace: FONT, fontSize: 8.5, color: C.t1, valign: "top", lineSpacingMultiple: 1.02, margin: 0 });
        if (auto) {
          s.addImage({ data: await ic("check", C.brand), x: nx + 0.2, y: ny + 0.8, w: 0.12, h: 0.12 });
          s.addText("авто", { x: nx + 0.36, y: ny + 0.76, w: 1, h: 0.2, fontFace: FONT, fontSize: 8.5, bold: true, color: C.brand, valign: "middle", margin: 0 });
        } else {
          s.addShape(pres.shapes.OVAL, { x: nx + 0.21, y: ny + 0.82, w: 0.085, h: 0.085, fill: { color: C.t3 } });
          s.addText("человек", { x: nx + 0.35, y: ny + 0.76, w: 1.2, h: 0.2, fontFace: FONT, fontSize: 8.5, bold: true, color: C.t3, valign: "middle", margin: 0 });
        }
        if (i < n - 1) s.addText("→", { x: nx + FNW, y: ny, w: arrowW, h: nh, fontFace: FONT, fontSize: 13, bold: true, color: auto ? C.brand400 : "B9C1CC", align: "center", valign: "middle", margin: 0 });
        nx += FNW + arrowW;
      }
    }

    await flow({ n: "01", t: "Создание промо-активности" }, 1.88, [
      { a: false, role: "TRADE-МАРКЕТИНГ", name: "Канал, партнёр, продукты, квоты" },
      { a: true, role: "СИСТЕМА", name: "Проверка наличия продукта на складе" },
      { a: true, role: "СИСТЕМА", name: "Присвоение кода PR-#####" },
      { a: true, role: "СИСТЕМА", name: "Составление Excel-файла промо" },
      { a: false, role: "CRM-СПЕЦИАЛИСТ", name: "Создание промо-активности в Panorama" },
      { a: true, role: "СИСТЕМА", name: "Авто-уведомление партнёров письмом" },
    ]);
    await flow({ n: "02", t: "Изменение промо-активности" }, 3.40, [
      { a: false, role: "TM / ON-TRADE", name: "Запрос: квота, период или продукт" },
      { a: true, role: "СИСТЕМА", name: "Обновление Excel-файла с изменениями" },
      { a: false, role: "CRM-СПЕЦИАЛИСТ", name: "Правка промо-активности в Panorama" },
      { a: true, role: "СИСТЕМА", name: "Синхронизация и уведомление партнёров" },
    ]);
    await flow({ n: "03", t: "Закрытие месяца и расчёт" }, 4.92, [
      { a: false, role: "ПАРТНЁР", name: "«Запрос»: проданные литры по промо" },
      { a: true, role: "СИСТЕМА", name: "Расчёт ∑(литры × цена периода)" },
      { a: true, role: "AI", name: "Верификация «Запроса» и сводный файл" },
    ]);

    // legend + callout
    const ly = 6.46;
    s.addShape(pres.shapes.RECTANGLE, { x: MX, y: ly + 0.05, w: 0.22, h: 0.16, fill: { color: C.brand100 }, line: { color: C.brand300, width: 1 } });
    s.addText("Автоматизировано системой", { x: MX + 0.3, y: ly, w: 2.6, h: 0.28, fontFace: FONT, fontSize: 9.5, color: C.t2, valign: "middle" });
    s.addShape(pres.shapes.RECTANGLE, { x: MX + 2.9, y: ly + 0.05, w: 0.22, h: 0.16, fill: { color: "FFFFFF" }, line: { color: C.line, width: 1 } });
    s.addText("Действие / решение человека", { x: MX + 3.2, y: ly, w: 2.8, h: 0.28, fontFace: FONT, fontSize: 9.5, color: C.t2, valign: "middle" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.6, y: ly - 0.07, w: CW - 6.98, h: 0.42, rectRadius: 0.08, fill: { color: C.brand100 }, line: { color: C.brand300, width: 1 } });
    s.addText([
      { text: "8 из 13", options: { fontFace: MONO, bold: true, color: C.brand } },
      { text: "  шагов выполняет система — людям остаются ввод промо и его создание в Panorama", options: { color: C.brand700 } },
    ], { x: 7.75, y: ly - 0.07, w: CW - 7.2, h: 0.42, fontFace: FONT, fontSize: 9.5, valign: "middle", margin: 0 });
    footer(s, false, "Efes Nexus · BPMN");
  }

  // ============ SLIDE 8 — EFFICIENCY EFFECT (hours) ============
  {
    const s = pres.addSlide(); s.background = { path: BG.ink };
    kicker(s, MX, 0.56, "Результат", true);
    h2(s, MX, 0.86, "Что даёт автоматизация — в рабочих часах", true);
    sub(s, MX, 1.56, 11.9, "Эффект считается без денежных оценок: только возвращённое рабочее время. Общая нагрузка на промо — 700 ч/мес (5 специалистов × 140 ч).", true);

    // big left card
    const lx = MX, ly = 2.45, lw = 4.5, lh = 3.95;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: lx, y: ly, w: lw, h: lh, rectRadius: 0.1, fill: { color: "0F2A1E" }, line: { color: "1F5A3B", width: 1 } });
    s.addText("Возвращённое рабочее время · Год 1", { x: lx + 0.4, y: ly + 0.4, w: lw - 0.8, h: 0.4, fontFace: FONT, fontSize: 13, bold: true, color: C.tw2 });
    s.addText([
      { text: "245", options: { fontSize: 72, bold: true } },
      { text: "  ч/мес", options: { fontSize: 24, bold: true } },
    ], { x: lx + 0.36, y: ly + 0.86, w: lw - 0.7, h: 1.3, fontFace: FONT, color: C.goodDk, valign: "middle" });
    s.addText("≈ 1,75 штатной ставки, освобождённой от ручной рутины промо", { x: lx + 0.4, y: ly + 2.2, w: lw - 0.8, h: 0.66, fontFace: FONT, fontSize: 13, color: C.tw1, lineSpacingMultiple: 1.2 });
    pill(s, lx + 0.4, ly + 3.06, "−35 % ручного труда", "good", 2.4, true);
    s.addText("при 35 % автоматизируемого времени", { x: lx + 0.4, y: ly + 3.48, w: lw - 0.8, h: 0.32, fontFace: FONT, fontSize: 10.5, color: C.tw3 });

    // right stat cards
    const rx = MX + lw + 0.34, rw = CW - lw - 0.34;
    const stats = [
      { icon: "clock", v: "245 ч/мес", l: "возвращённого рабочего времени — расчёт: 5 спец. × 140 ч × 35 % автоматизации" },
      { icon: "users", v: "≈ 1,75 ставки", l: "штатной ёмкости освобождается от ручной обработки промо (Год 1)" },
      { icon: "trend", v: "Растёт с объёмом", l: "чем больше промо-активностей, тем выше доля автоматизации — эффект усиливается без доп. затрат" },
    ];
    const sh = 1.18, sg = 0.2;
    for (let i = 0; i < 3; i++) {
      const y = ly + i * (sh + sg);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: rx, y, w: rw, h: sh, rectRadius: 0.09, fill: { color: C.inkCard }, line: { color: C.inkLine, width: 1 } });
      await iconChip(s, rx + 0.34, y + 0.28, 0.62, stats[i].icon, "green", true);
      s.addText(stats[i].v, { x: rx + 1.18, y: y + 0.2, w: rw - 1.5, h: 0.44, fontFace: FONT, fontSize: 21, bold: true, color: C.goodDk, valign: "middle" });
      s.addText(stats[i].l, { x: rx + 1.18, y: y + 0.62, w: rw - 1.5, h: 0.5, fontFace: FONT, fontSize: 11, color: C.tw2, valign: "top", lineSpacingMultiple: 1.1 });
    }
    // speed band
    const by = 6.46; // overlaps footer? footer at 6.96 -> keep band above
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 6.46, w: CW, h: 0.4, rectRadius: 0.08, fill: { color: "12203A" }, line: { color: C.inkLine, width: 1 } });
    s.addText("Скорость обработки запроса партнёра", { x: MX + 0.34, y: 6.46, w: 6, h: 0.4, fontFace: FONT, fontSize: 11.5, bold: true, color: C.tw1, valign: "middle" });
    const sbR = MX + CW; // band right edge
    pill(s, sbR - 0.3 - 1.0, 6.51, "< 1 дня", "good", 1.0, true);
    s.addText("→", { x: sbR - 0.3 - 1.0 - 0.5, y: 6.46, w: 0.42, h: 0.4, fontFace: FONT, fontSize: 15, bold: true, color: C.goodDk, align: "center", valign: "middle", margin: 0 });
    pill(s, sbR - 0.3 - 1.0 - 0.5 - 0.08 - 1.12, 6.51, "3–5 дней", "bad", 1.12, true);
    footer(s, true, "Efes Nexus");
  }

  // ============ SLIDE 9 — 3-YEAR OUTLOOK / automation grows with volume ============
  {
    const s = pres.addSlide(); s.background = { path: BG.dark };
    kicker(s, MX, 0.52, "Перспектива · 3 года", true);
    h2(s, MX, 0.82, "Чем больше объём — тем выше автоматизация", true);
    sub(s, MX, 1.5, 12.0, "Автоматизированные шаги — уведомления, расчёты, Excel, AI-проверка, синхронизация — выполняются почти без затрат времени на каждую новую промо. Ручные шаги фиксированы. Поэтому с ростом объёма доля автоматизации растёт к структурному потолку 8 из 13 шагов (≈ 62 %).", true);

    const years = [
      { y: "Год 1", pct: "35 %", h: "245 ч/мес", fte: "≈ 1,75 ставки", b: "Запуск платформы. Единый источник данных вместо разрозненных Excel-файлов.", hot: false },
      { y: "Год 2", pct: "40 %", h: "280 ч/мес", fte: "≈ 2,0 ставки", b: "Объём промо растёт — система берёт на себя бóльшую долю рутины.", hot: false },
      { y: "Год 3", pct: "45 %", h: "315 ч/мес", fte: "≈ 2,25 ставки", b: "Автоматизация приближается к потолку 8/13 шагов. Команда — на стратегии бренда.", hot: true },
    ];
    const gx = 0.36, cw = (CW - gx * 2) / 3, cy = 2.46, ch = 2.42;
    for (let i = 0; i < 3; i++) {
      const x = MX + i * (cw + gx);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy, w: cw, h: ch, rectRadius: 0.09, fill: { color: C.inkCard }, line: { color: years[i].hot ? C.brand400 : C.inkLine, width: years[i].hot ? 1.5 : 1 } });
      pill(s, x + 0.3, cy + 0.3, years[i].y, "brand", 1.0, true);
      s.addText(years[i].pct, { x: x + cw - 1.7, y: cy + 0.24, w: 1.4, h: 0.46, fontFace: FONT, fontSize: 22, bold: true, color: C.brand300, align: "right", valign: "middle" });
      s.addText("автоматизации", { x: x + cw - 1.7, y: cy + 0.66, w: 1.4, h: 0.2, fontFace: FONT, fontSize: 8.5, color: C.tw3, align: "right" });
      s.addText([
        { text: years[i].h + "   ", options: { fontSize: 19, bold: true, color: C.goodDk } },
        { text: years[i].fte, options: { fontSize: 12, color: C.tw2 } },
      ], { x: x + 0.3, y: cy + 1.0, w: cw - 0.6, h: 0.4, fontFace: FONT, valign: "middle" });
      s.addShape(pres.shapes.LINE, { x: x + 0.3, y: cy + 1.5, w: cw - 0.6, h: 0, line: { color: C.inkLine, width: 1 } });
      s.addText(years[i].b, { x: x + 0.3, y: cy + 1.6, w: cw - 0.6, h: 0.74, fontFace: FONT, fontSize: 11, color: C.tw2, valign: "top", lineSpacingMultiple: 1.16 });
    }

    // bottom: cumulative + checklist
    const by = 5.18, bh = 1.62;
    const cumW = 3.5;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: by, w: cumW, h: bh, rectRadius: 0.1, fill: { color: "0F2A1E" }, line: { color: "1F5A3B", width: 1 } });
    s.addText("Накопленный эффект за 3 года", { x: MX + 0.26, y: by + 0.2, w: cumW - 0.5, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: C.tw2, align: "center" });
    s.addText("≈ 10 000 ч", { x: MX, y: by + 0.5, w: cumW, h: 0.6, fontFace: FONT, fontSize: 38, bold: true, color: C.goodDk, align: "center", valign: "middle" });
    s.addText("≈ 6 человеко-лет рабочего времени", { x: MX + 0.2, y: by + 1.14, w: cumW - 0.4, h: 0.3, fontFace: FONT, fontSize: 10.5, color: C.tw3, align: "center" });

    const klx = MX + cumW + 0.34, klw = CW - cumW - 0.34;
    const checks = [
      "Efes Nexus автоматизирует процессы всей компании, а не только промо",
      "Команда сфокусирована на стратегии и развитии бренда вместо ручной рутины",
      "Технологичный бренд — сильный PR и аргумент для привлечения инвестиций",
      "Отрыв от конкурентов по скорости, точности и масштабу процессов",
    ];
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: klx, y: by, w: klw, h: bh, rectRadius: 0.1, fill: { color: C.inkCard }, line: { color: C.inkLine, width: 1 } });
    const chk = await ic("check", C.goodDk);
    for (let i = 0; i < 4; i++) {
      const yy = by + 0.22 + i * 0.34;
      s.addImage({ data: chk, x: klx + 0.32, y: yy + 0.03, w: 0.16, h: 0.16 });
      s.addText(checks[i], { x: klx + 0.6, y: yy - 0.04, w: klw - 0.9, h: 0.3, fontFace: FONT, fontSize: 11.5, bold: true, color: C.tw1, valign: "middle" });
    }
    footer(s, true, "Efes Nexus");
  }

  await pres.writeFile({ fileName: OUT });
  console.log("WROTE", OUT);
})().catch(e => { console.error(e); process.exit(1); });
