#!/usr/bin/env node
/**
 * 產出「黃家朗新人報告.pptx」
 *
 * 內容來源：report/outline.md（改內容請改那份，然後重跑這支腳本）
 * 配色沿用參考範本 江宜恩新人報告.pdf：深藍 / 薰衣草紫 / 奶油黃 / 橘紅重點
 *
 *   node report/build_pptx.js
 */

const fs = require("fs");
const path = require("path");
const PptxGenJS = require("pptxgenjs");

const ROOT = path.resolve(__dirname, "..");
const ASSETS = path.join(__dirname, "assets");
const OUT = path.join(ROOT, "黃家朗新人報告.pptx");

// ---------------------------------------------------------------- design tokens

const NAVY = "2B2D5C";
const NAVY_DEEP = "1E2048";
const LAV = "D6D9F0";
const LAV_MID = "AEB4E4";
const CREAM = "F7F1CE";
const ACCENT = "D95F2B";
const TEAL = "17A0A0";
const MINT = "6DCFA5";
const SLATE = "35697F";
const BLUSH = "F2E2DA";
const WHITE = "FFFFFF";
const MUTED = "6E7191";
const SOFT_BG = "ECEEF8";
const PLACEHOLDER = "E4E6F0";

const FONT = "微軟正黑體";

const W = 13.333;
const H = 7.5;
const M = 0.7; // 版面邊界
const CW = W - M * 2; // 內容寬度 11.933

// 每次都回傳新物件——pptxgenjs 會就地改寫傳入的 options
const shadow = () => ({ type: "outer", color: "9AA0BF", blur: 8, offset: 2, angle: 90, opacity: 0.22 });

let pageNo = 0;

// ---------------------------------------------------------------- helpers

function newSlide(pres, bg = WHITE) {
  const s = pres.addSlide();
  s.background = { color: bg };
  pageNo += 1;
  return s;
}

/** 右下角頁碼 */
function stampPage(slide, color = MUTED) {
  slide.addText(String(pageNo), {
    x: W - 1.1, y: H - 0.62, w: 0.6, h: 0.36,
    fontFace: FONT, fontSize: 11, color, align: "right", margin: 0,
  });
}

/** 內容頁標準頁首：章節 chip + 標題 */
function header(slide, chapter, title) {
  if (chapter) {
    slide.addText(chapter, {
      x: M, y: 0.52, w: 2.5, h: 0.36,
      shape: "roundRect", rectRadius: 0.16, fill: { color: LAV },
      fontFace: FONT, fontSize: 12, bold: true, color: NAVY,
      align: "center", valign: "middle", margin: 0,
    });
  }
  slide.addText(title, {
    x: M, y: chapter ? 1.0 : 0.7, w: CW, h: 0.8,
    fontFace: FONT, fontSize: 32, bold: true, color: NAVY,
    align: "left", valign: "middle", margin: 0,
  });
}

/** 找圖；找不到回傳 null */
function findAsset(base) {
  for (const ext of [".png", ".jpg", ".jpeg"]) {
    const p = path.join(ASSETS, base + ext);
    if (fs.existsSync(p)) return p;
  }
  return null;
}

/** 放圖，沒圖就放灰底佔位方塊（標示該放什麼） */
function image(slide, base, label, { x, y, w, h }) {
  const p = findAsset(base);
  if (p) {
    slide.addImage({ path: p, x, y, w, h, sizing: { type: "contain", w, h } });
    return true;
  }
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: PLACEHOLDER },
    line: { color: LAV_MID, width: 1, dashType: "dash" },
  });
  slide.addText(`【待放圖】${label}\n${base}.png`, {
    x: x + 0.15, y, w: w - 0.3, h,
    fontFace: FONT, fontSize: 11, color: MUTED,
    align: "center", valign: "middle", margin: 0,
  });
  return false;
}

/** 卡片：標題 + 內文 */
function card(slide, { x, y, w, h, title, body, tint = SOFT_BG, titleColor = NAVY, titleSize = 17 }) {
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: tint }, line: { color: "FFFFFF", width: 0 },
    shadow: shadow(),
  });
  slide.addText(title, {
    x: x + 0.32, y: y + 0.24, w: w - 0.64, h: 0.42,
    fontFace: FONT, fontSize: titleSize, bold: true, color: titleColor, margin: 0, valign: "middle",
  });
  slide.addText(body, {
    x: x + 0.32, y: y + 0.72, w: w - 0.64, h: h - 0.98,
    fontFace: FONT, fontSize: 13.5, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });
}

/** 編號圓圈 */
function numCircle(slide, n, { x, y, d = 0.52, fill = NAVY, color = WHITE, size = 16 }) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: fill } });
  slide.addText(String(n), {
    x, y, w: d, h: d,
    fontFace: FONT, fontSize: size, bold: true, color,
    align: "center", valign: "middle", margin: 0,
  });
}

/** 表格 */
function table(slide, rows, { x, y, w, colW, rowH, headFill = LAV, fontSize = 12.5 }) {
  slide.addTable(rows, {
    x, y, w, colW, rowH,
    fontFace: FONT, fontSize, color: NAVY, valign: "middle",
    border: { type: "solid", color: "DDE0EE", pt: 1 },
    fill: { color: WHITE },
    autoPage: false,
  });
  return headFill;
}

const th = (t) => ({ text: t, options: { bold: true, fill: { color: LAV }, color: NAVY, align: "center" } });
const td = (t, opts = {}) => ({ text: t, options: { margin: [4, 8, 4, 8], ...opts } });

/** 標籤／內容成對的規格清單（避免長段落擠成一塊被截斷） */
function specList(slide, { x, y, w, pairs, gap = 0.8 }) {
  pairs.forEach((p, i) => {
    const yy = y + i * gap;
    slide.addText(p[0], {
      x, y: yy, w, h: 0.3,
      fontFace: FONT, fontSize: 12.5, bold: true, color: ACCENT, margin: 0, valign: "middle",
    });
    slide.addText(p[1], {
      x, y: yy + 0.3, w, h: 0.44,
      fontFace: FONT, fontSize: 14.5, color: p[2] ? ACCENT : NAVY, margin: 0, valign: "top",
    });
  });
}

/** 圓餅圖 + 右側圖例註解 */
function pieWithNotes(pres, slide, { labels, values, colors, notes, noteTitle }) {
  slide.addChart(pres.ChartType.pie, [{ name: "分佈", labels, values }], {
    x: 0.45, y: 1.75, w: 5.5, h: 4.9,
    chartColors: colors,
    showLegend: false,
    showPercent: true,
    dataLabelColor: WHITE,
    dataLabelFontFace: FONT,
    dataLabelFontSize: 13,
    dataLabelFontBold: true,
    holeSize: 0,
  });

  const bx = 6.35;
  let by = 1.95;
  if (noteTitle) {
    slide.addText(noteTitle, {
      x: bx, y: by, w: CW - (bx - M), h: 0.4,
      fontFace: FONT, fontSize: 14, bold: true, color: MUTED, margin: 0,
    });
    by += 0.5;
  }
  const step = 0.86;
  labels.forEach((lab, i) => {
    const yy = by + i * step;
    slide.addShape("ellipse", { x: bx, y: yy + 0.1, w: 0.26, h: 0.26, fill: { color: colors[i] } });
    slide.addText(
      [
        { text: `${lab}　`, options: { bold: true, fontSize: 16, color: NAVY } },
        { text: `${values[i]}%`, options: { fontSize: 13, color: MUTED } },
      ],
      { x: bx + 0.42, y: yy, w: CW - (bx - M) - 0.42, h: 0.4, fontFace: FONT, margin: 0, valign: "middle" }
    );
    if (notes && notes[i]) {
      slide.addText(notes[i], {
        x: bx + 0.42, y: yy + 0.38, w: CW - (bx - M) - 0.42, h: 0.36,
        fontFace: FONT, fontSize: 12, color: MUTED, margin: 0, valign: "top",
      });
    }
  });
}

// ================================================================ deck

const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "黃家朗";
pres.title = "新人報告";

// ---------------------------------------------------------------- P1 封面
{
  const s = newSlide(pres, NAVY_DEEP);
  s.addText("新人報告", {
    x: M, y: 2.25, w: CW, h: 1.5,
    fontFace: FONT, fontSize: 62, bold: true, color: WHITE, margin: 0, valign: "middle",
  });
  s.addText("產品工程師－黃家朗", {
    x: M, y: 3.95, w: 5.2, h: 0.62,
    shape: "roundRect", rectRadius: 0.1, fill: { color: LAV },
    fontFace: FONT, fontSize: 19, bold: true, color: NAVY,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText("創新研發部　｜　到職日 2026.04.10", {
    x: M, y: 4.85, w: CW, h: 0.44,
    fontFace: FONT, fontSize: 15, color: LAV_MID, margin: 0,
  });
  s.addText("程曦資訊整合股份有限公司", {
    x: M, y: H - 1.05, w: CW, h: 0.4,
    fontFace: FONT, fontSize: 12.5, color: "7A80AE", margin: 0,
  });
  stampPage(s, "7A80AE");
}

// ---------------------------------------------------------------- P2 目錄
{
  const s = newSlide(pres);
  header(s, null, "目錄");
  const items = [
    ["自我介紹", "我是誰、從哪裡來"],
    ["公司精神", "愛、新、勤、誠與我的實踐"],
    ["我的三階段工作主軸", "這三個月實際在做什麼"],
    ["學習與工作成果", "對外四門課｜對內神秘客系統"],
    ["學習心得", "收穫、成就感與挑戰"],
    ["職涯規劃", "短、中、長期目標"],
    ["對公司的建議", "新人入職與知識傳承"],
  ];
  const colX = [M, M + CW / 2 + 0.25];
  items.forEach((it, i) => {
    const col = i < 4 ? 0 : 1;
    const row = i < 4 ? i : i - 4;
    const x = colX[col];
    const y = 1.85 + row * 1.18;
    numCircle(s, i + 1, { x, y: y + 0.08, d: 0.52, fill: i === 3 ? ACCENT : LAV, color: i === 3 ? WHITE : NAVY });
    s.addText(it[0], {
      x: x + 0.72, y, w: CW / 2 - 1.0, h: 0.42,
      fontFace: FONT, fontSize: 18, bold: true, color: NAVY, margin: 0, valign: "middle",
    });
    s.addText(it[1], {
      x: x + 0.72, y: y + 0.42, w: CW / 2 - 1.0, h: 0.34,
      fontFace: FONT, fontSize: 12, color: MUTED, margin: 0, valign: "middle",
    });
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P3 自我介紹
{
  const s = newSlide(pres);
  header(s, "一、自我介紹", "自我介紹");
  const cw = (CW - 0.4) / 2;
  card(s, { x: M, y: 2.05, w: cw, h: 2.05, title: "學歷", body: "【待補：學校／科系／畢業年】", tint: LAV });
  card(s, { x: M + cw + 0.4, y: 2.05, w: cw, h: 2.05, title: "到職前經歷", body: "【待補：公司／職稱】", tint: SOFT_BG });
  card(s, { x: M, y: 4.35, w: cw, h: 2.05, title: "現職", body: "程曦資訊整合股份有限公司\n創新研發部｜產品工程師", tint: SOFT_BG });
  card(s, { x: M + cw + 0.4, y: 4.35, w: cw, h: 2.05, title: "興趣", body: "【待補：興趣】", tint: CREAM });
  stampPage(s);
}

// ---------------------------------------------------------------- P4 工作歷程
{
  const s = newSlide(pres);
  header(s, "一、自我介紹", "工作歷程");

  // 時間軸底線
  s.addShape("rect", { x: M + 0.3, y: 3.35, w: CW - 0.6, h: 0.045, fill: { color: LAV } });

  const nodes = [
    { t: "【待補：年月】", n: "【待補：前職公司】", r: "【待補：職稱】", cx: 3.2, color: LAV_MID },
    { t: "2026.04", n: "程曦資訊整合股份有限公司", r: "創新研發部｜產品工程師", cx: 9.6, color: ACCENT },
  ];
  nodes.forEach((nd) => {
    s.addShape("ellipse", { x: nd.cx - 0.16, y: 3.2, w: 0.34, h: 0.34, fill: { color: nd.color } });
    s.addText(nd.t, {
      x: nd.cx - 2.0, y: 2.45, w: 4.0, h: 0.4,
      fontFace: FONT, fontSize: 17, bold: true, color: NAVY, align: "center", margin: 0,
    });
    s.addText(nd.n, {
      x: nd.cx - 2.4, y: 3.85, w: 4.8, h: 0.44,
      fontFace: FONT, fontSize: 15, bold: true, color: NAVY, align: "center", margin: 0,
    });
    s.addText(nd.r, {
      x: nd.cx - 2.4, y: 4.28, w: 4.8, h: 0.4,
      fontFace: FONT, fontSize: 12.5, color: MUTED, align: "center", margin: 0,
    });
  });

  s.addText("轉職動機：從【待補：前職領域】轉進 AI／教育訓練領域", {
    x: M, y: 5.45, w: CW, h: 0.68,
    shape: "roundRect", rectRadius: 0.08, fill: { color: CREAM },
    fontFace: FONT, fontSize: 15, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P5 個人特質
{
  const s = newSlide(pres);
  header(s, "一、自我介紹", "個人特質");
  pieWithNotes(pres, s, {
    labels: ["解決力", "自學力", "系統思維", "換位思考", "交付到底"],
    values: [25, 25, 20, 15, 15],
    colors: [NAVY, TEAL, MINT, SLATE, LAV_MID],
    notes: [
      "不只回報問題，直接把工具做出來。",
      "沒學過的東西，用 AI 工具把它做完。",
      "看見五份 Excel 背後其實是同一件事。",
      "依三種層級，設計三種權限。",
      "做到實際上線在用，不是 demo。",
    ],
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P6 公司精神
{
  const s = newSlide(pres);
  header(s, "二、公司精神", "愛、新、勤、誠");
  const items = [
    { ch: "愛", body: "愛自己、愛同事、愛家人\n愛客戶、愛公司、愛競爭對手", tint: BLUSH },
    { ch: "新", body: "求新、創新", tint: LAV },
    { ch: "勤", body: "勤勞、勤儉", tint: CREAM },
    { ch: "誠", body: "誠實、忠誠", tint: SOFT_BG },
  ];
  const cw = (CW - 0.36 * 3) / 4;
  items.forEach((it, i) => {
    const x = M + i * (cw + 0.36);
    s.addShape("roundRect", {
      x, y: 2.05, w: cw, h: 3.2, rectRadius: 0.06,
      fill: { color: it.tint }, shadow: shadow(),
    });
    s.addText(it.ch, {
      x, y: 2.35, w: cw, h: 1.3,
      fontFace: FONT, fontSize: 60, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0,
    });
    s.addText(it.body, {
      x: x + 0.22, y: 3.7, w: cw - 0.44, h: 1.35,
      fontFace: FONT, fontSize: 12.5, color: NAVY, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.3,
    });
  });
  s.addText("勤勞地用創新的方式，誠懇地愛人。", {
    x: M, y: 5.62, w: CW, h: 0.7,
    fontFace: FONT, fontSize: 21, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P7 我的實踐
{
  const s = newSlide(pres);
  header(s, "二、公司精神", "我的實踐");
  const items = [
    ["新", "全程用 replit 等 AI 工具開發神秘客系統——用新方法解決舊問題。", LAV],
    ["勤", "把每月要重做一次的流程，改成設定一次就長期適用。", CREAM],
    ["愛", "站在三種角色的立場設計三種權限，讓每個人只看到他該看的。", BLUSH],
    ["誠", "效益數字標示為「推估」，沒有實測就不寫成實測。", SOFT_BG],
  ];
  const cw = (CW - 0.4) / 2;
  items.forEach((it, i) => {
    const x = M + (i % 2) * (cw + 0.4);
    const y = 2.05 + Math.floor(i / 2) * 2.35;
    s.addShape("roundRect", { x, y, w: cw, h: 2.1, rectRadius: 0.06, fill: { color: it[2] }, shadow: shadow() });
    s.addShape("ellipse", { x: x + 0.34, y: y + 0.55, w: 0.95, h: 0.95, fill: { color: NAVY } });
    s.addText(it[0], {
      x: x + 0.34, y: y + 0.55, w: 0.95, h: 0.95,
      fontFace: FONT, fontSize: 32, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0,
    });
    s.addText(it[1], {
      x: x + 1.55, y: y + 0.42, w: cw - 1.9, h: 1.25,
      fontFace: FONT, fontSize: 14.5, color: NAVY, margin: 0, valign: "middle", lineSpacingMultiple: 1.3,
    });
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P8 三階段工作主軸
{
  const s = newSlide(pres);
  header(s, "三、工作主軸", "我的三階段工作主軸");
  const rows = [
    [th("階段"), th("期間"), th("主軸"), th("具體工作")],
    [
      td("一、熟悉與投入", { bold: true, fill: { color: SOFT_BG } }),
      td("到職～第 4 週", { align: "center" }),
      td("熟悉公司業務與 AI 教育訓練產品線"),
      td("認識部門定位與產品線；投入對外課程教材製作"),
    ],
    [
      td("二、產出", { bold: true, fill: { color: SOFT_BG } }),
      td("第 5～8 週", { align: "center" }),
      td("對外課程教材產出與授課"),
      td("完成 2 門課教材撰寫；擔任 2 門課講師"),
    ],
    [
      td("三、落地", { bold: true, fill: { color: SOFT_BG } }),
      td("第 9～12 週", { align: "center" }),
      td("對內流程數位化，從需求訪談到系統上線"),
      td("神秘客訪問電話系統開發並實際上線"),
    ],
  ];
  table(s, rows, { x: M, y: 2.1, w: CW, colW: [2.0, 1.6, 3.5, 4.833], rowH: 0.95 });
  s.addText("三段的共同點：都是在做「把 AI 變成別人真的用得到的東西」。", {
    x: M, y: 6.05, w: CW, h: 0.62,
    shape: "roundRect", rectRadius: 0.08, fill: { color: CREAM },
    fontFace: FONT, fontSize: 15, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P9 工作定位
{
  const s = newSlide(pres);
  header(s, "三、工作主軸", "工作定位：對外 7 ： 對內 3");
  const cw = (CW - 0.4) / 2;

  const panes = [
    {
      x: M, tint: LAV, pct: "70%", head: "對外｜教育訓練產品線",
      lines: ["四門 AI 應用課程", "教材撰寫 2 門、擔任講師 2 門", "目標：成為可對外銷售的課程商品"],
    },
    {
      x: M + cw + 0.4, tint: SOFT_BG, pct: "30%", head: "對內｜內部流程數位化",
      lines: ["神秘客訪問電話系統", "從需求訪談到開發上線", "目標：讓內部作業真的變快"],
    },
  ];
  panes.forEach((p) => {
    s.addShape("roundRect", { x: p.x, y: 2.05, w: cw, h: 3.55, rectRadius: 0.06, fill: { color: p.tint }, shadow: shadow() });
    s.addText(p.pct, {
      x: p.x + 0.35, y: 2.25, w: cw - 0.7, h: 0.95,
      fontFace: FONT, fontSize: 46, bold: true, color: NAVY, margin: 0, valign: "middle",
    });
    s.addText(p.head, {
      x: p.x + 0.35, y: 3.2, w: cw - 0.7, h: 0.45,
      fontFace: FONT, fontSize: 17, bold: true, color: NAVY, margin: 0, valign: "middle",
    });
    s.addText(
      p.lines.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i !== p.lines.length - 1 } })),
      {
        x: p.x + 0.35, y: 3.75, w: cw - 0.7, h: 1.6,
        fontFace: FONT, fontSize: 13.5, color: NAVY, margin: 0, valign: "top", paraSpaceAfter: 6,
      }
    );
  });
  s.addText("對外，把 AI 變成可以賣的課；對內，把 AI 變成可以用的系統。", {
    x: M, y: 5.9, w: CW, h: 0.72,
    shape: "roundRect", rectRadius: 0.08, fill: { color: NAVY },
    fontFace: FONT, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P10 扉頁：成果
{
  const s = newSlide(pres, NAVY_DEEP);
  s.addText("四", {
    x: M, y: 2.1, w: 2.0, h: 0.9,
    fontFace: FONT, fontSize: 22, bold: true, color: LAV_MID, margin: 0, valign: "middle",
  });
  s.addText("學習與工作成果", {
    x: M, y: 2.85, w: CW, h: 1.25,
    fontFace: FONT, fontSize: 52, bold: true, color: WHITE, margin: 0, valign: "middle",
  });
  s.addText("對外｜四門 AI 應用課程　　　對內｜神秘客訪問電話系統", {
    x: M, y: 4.3, w: CW, h: 0.5,
    fontFace: FONT, fontSize: 17, color: LAV_MID, margin: 0,
  });
  stampPage(s, "7A80AE");
}

// ---------------------------------------------------------------- P11 四門課總覽
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對外", "對外線：四門 AI 應用課");
  const rows = [
    [th("課程"), th("我的角色"), th("狀態")],
    [td("① AI 時代的商業企劃思維與簡報製作工具"), td("教材撰寫", { align: "center", bold: true, color: ACCENT }), td("教材已完成", { align: "center" })],
    [td("② AI 圖片生成課"), td("教材撰寫", { align: "center", bold: true, color: ACCENT }), td("教材已完成", { align: "center" })],
    [td("③ AI 輔助 UI/UX 與原型設計"), td("擔任講師", { align: "center", bold: true, color: TEAL }), td("已授課", { align: "center" })],
    [td("④ AI 輔助 vibe-coding 程式開發"), td("擔任講師", { align: "center", bold: true, color: TEAL }), td("已授課", { align: "center" })],
  ];
  table(s, rows, { x: M, y: 2.1, w: CW, colW: [6.6, 2.6, 2.733], rowH: 0.72, fontSize: 13.5 });
  s.addText("【待補：各課時數與面向對象】", {
    x: M, y: 5.85, w: CW, h: 0.4,
    fontFace: FONT, fontSize: 12, color: ACCENT, margin: 0,
  });
  s.addText("寫教材要把主題徹底搞懂才寫得出來；上台授課才知道學員會卡在哪裡——這兩件事互相回饋。", {
    x: M, y: 6.3, w: CW, h: 0.6,
    fontFace: FONT, fontSize: 14, color: MUTED, margin: 0, valign: "middle",
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P12 教材①
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對外", "教材撰寫①　AI 時代的商業企劃思維與簡報製作工具");
  const tw = 6.1;
  specList(s, {
    x: M, y: 2.15, w: tw,
    pairs: [
      ["我負責", "課程架構、教材內容、實作案例"],
      ["課程主題", "用 AI 輔助商業企劃的思考流程，並產出簡報"],
      ["對象", "【待補：面向對象】", true],
      ["規模", "【待補：時數／教材頁數／案例數】", true],
    ],
  });
  s.addText("重點不在工具操作，而在思維：先拆解企劃從無到有的流程，再看每一段可以怎麼用 AI，而不是從工具功能表開始教。", {
    x: M, y: 5.5, w: tw, h: 1.15,
    shape: "roundRect", rectRadius: 0.08, fill: { color: CREAM },
    fontFace: FONT, fontSize: 13.5, color: NAVY, margin: [10, 14, 10, 14], valign: "middle", lineSpacingMultiple: 1.25,
  });
  image(s, "course-01-slide", "課程①教材頁", { x: M + tw + 0.45, y: 2.15, w: CW - tw - 0.45, h: 4.5 });
  stampPage(s);
}

// ---------------------------------------------------------------- P13 教材②
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對外", "教材撰寫②　AI 圖片生成課");
  const iw = 6.0;
  image(s, "course-02-slide", "課程②教材頁", { x: M, y: 2.15, w: iw, h: 4.5 });
  const tx = M + iw + 0.45;
  const tw = CW - iw - 0.45;
  specList(s, {
    x: tx, y: 2.15, w: tw,
    pairs: [
      ["我負責", "課程架構、教材內容、實作案例"],
      ["課程主題", "AI 圖片生成的原理、提示詞撰寫與實務應用"],
      ["對象", "【待補：面向對象】", true],
      ["規模", "【待補：時數／教材頁數／案例數】", true],
    ],
  });
  s.addText("圖片生成結果不穩定，所以教材重點放在「怎麼描述」與「怎麼修正」，而不是給一堆可複製的提示詞——範本會過期，方法不會。", {
    x: tx, y: 5.5, w: tw, h: 1.15,
    shape: "roundRect", rectRadius: 0.08, fill: { color: LAV },
    fontFace: FONT, fontSize: 13.5, color: NAVY, margin: [10, 14, 10, 14], valign: "middle", lineSpacingMultiple: 1.25,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P14 擔任講師
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對外", "擔任講師的兩門課");
  const cw = (CW - 0.4) / 2;
  const items = [
    { x: M, n: "③", name: "AI 輔助 UI/UX 與原型設計", asset: "course-03-slide", tint: SOFT_BG },
    { x: M + cw + 0.4, n: "④", name: "AI 輔助 vibe-coding 程式開發", asset: "course-04-slide", tint: LAV },
  ];
  items.forEach((it) => {
    s.addShape("roundRect", { x: it.x, y: 2.05, w: cw, h: 3.05, rectRadius: 0.06, fill: { color: it.tint }, shadow: shadow() });
    s.addText(`${it.n}　${it.name}`, {
      x: it.x + 0.3, y: 2.25, w: cw - 0.6, h: 0.75,
      fontFace: FONT, fontSize: 18, bold: true, color: NAVY, margin: 0, valign: "middle",
    });
    s.addText(
      [
        { text: "教材：同事撰寫　／　我負責：授課", options: { breakLine: true } },
        { text: "【待補：時數／對象／場次】", options: { color: ACCENT } },
      ],
      { x: it.x + 0.3, y: 3.1, w: cw - 0.6, h: 0.9, fontFace: FONT, fontSize: 13.5, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.3 }
    );
    image(s, it.asset, `${it.n} 教材頁`, { x: it.x + 0.3, y: 4.05, w: cw - 0.6, h: 0.92 });
  });
  s.addText("vibe-coding 這門課教的方法，就是我後來開發神秘客系統實際用的方法——先教別人怎麼用 AI 寫程式，然後自己拿這套方法做出一個真的上線的系統。", {
    x: M, y: 5.4, w: CW, h: 1.05,
    shape: "roundRect", rectRadius: 0.08, fill: { color: CREAM },
    fontFace: FONT, fontSize: 14.5, bold: true, color: NAVY, margin: [10, 16, 10, 16], valign: "middle", lineSpacingMultiple: 1.25,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P15 教材展示
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對外", "教材與課堂實況");
  const gw = (CW - 0.5 * 2) / 3;
  ["course-showcase-01", "course-showcase-02", "course-showcase-03"].forEach((a, i) => {
    image(s, a, `教材展示 ${i + 1}`, { x: M + i * (gw + 0.5), y: 2.15, w: gw, h: 3.9 });
  });
  s.addText("完整教材可於課後單獨提供。", {
    x: M, y: 6.25, w: CW, h: 0.42,
    fontFace: FONT, fontSize: 13, color: MUTED, margin: 0,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P16 對外線現況
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對外", "對外線現況與後續");
  const cw = (CW - 0.4 * 2) / 3;
  const items = [
    { title: "現況", body: "四門課教材皆已完成\n具備對外開課條件", tint: MINT, tc: NAVY },
    { title: "下一步", body: "待對外開賣\n【待補：預計開賣時間】", tint: CREAM, tc: NAVY },
    { title: "我可以再做的", body: "依招生回饋修訂教材\n擴充案例、開發進階課程", tint: LAV, tc: NAVY },
  ];
  items.forEach((it, i) => {
    card(s, {
      x: M + i * (cw + 0.4), y: 2.15, w: cw, h: 2.0,
      title: it.title, body: it.body, tint: it.tint, titleColor: it.tc, titleSize: 19,
    });
  });
  s.addText("教材完成只是第一步。真正會讓課變好的是實際開課後的學員回饋，後續若有招生與開課，我希望能持續參與。", {
    x: M, y: 4.6, w: CW, h: 1.1,
    shape: "roundRect", rectRadius: 0.08, fill: { color: SOFT_BG },
    fontFace: FONT, fontSize: 15, color: NAVY, margin: [12, 18, 12, 18], valign: "middle", lineSpacingMultiple: 1.3,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P17 扉頁：神秘客
{
  const s = newSlide(pres, NAVY_DEEP);
  s.addText("四、工作成果｜對內", {
    x: M, y: 2.1, w: 5.0, h: 0.9,
    fontFace: FONT, fontSize: 22, bold: true, color: LAV_MID, margin: 0, valign: "middle",
  });
  s.addText("神秘客訪問電話系統", {
    x: M, y: 2.85, w: CW, h: 1.25,
    fontFace: FONT, fontSize: 52, bold: true, color: WHITE, margin: 0, valign: "middle",
  });
  s.addText("從需求訪談到實際上線　｜　目前已在使用中", {
    x: M, y: 4.3, w: CW, h: 0.5,
    fontFace: FONT, fontSize: 17, color: LAV_MID, margin: 0,
  });
  stampPage(s, "7A80AE");
}

// ---------------------------------------------------------------- P18 業務流程
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對內", "神秘客訪問電話是什麼");
  const steps = [
    ["互打", "各專案組每月隨機撥打電話給其他專案組"],
    ["檢核", "以腳本上的常見 FAQ，檢驗接聽同仁是否合乎禮儀、使用標準起手台詞、答案正確"],
    ["彙整", "各組寫抽查成果與分數給主管，主管彙整報表與未達滿分的經緯報告"],
    ["發布", "分數向下發給各專案組、向上呈報高級主管"],
  ];
  const cw = (CW - 0.34 * 3) / 4;
  steps.forEach((st, i) => {
    const x = M + i * (cw + 0.34);
    s.addShape("roundRect", { x, y: 2.2, w: cw, h: 3.0, rectRadius: 0.06, fill: { color: i % 2 ? LAV : SOFT_BG }, shadow: shadow() });
    numCircle(s, i + 1, { x: x + cw / 2 - 0.29, y: 2.5, d: 0.58, fill: NAVY, size: 17 });
    s.addText(st[0], {
      x: x + 0.2, y: 3.25, w: cw - 0.4, h: 0.55,
      fontFace: FONT, fontSize: 21, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0,
    });
    s.addText(st[1], {
      x: x + 0.25, y: 3.85, w: cw - 0.5, h: 1.55,
      fontFace: FONT, fontSize: 12.5, color: NAVY, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.25,
    });
    if (i < 3) {
      // 用圖形而非字元，避免不同電腦字型 fallback 把箭頭畫成方塊
      s.addShape("triangle", {
        x: x + cw + 0.075, y: 3.55, w: 0.19, h: 0.24,
        fill: { color: LAV_MID }, rotate: 90,
      });
    }
  });
  s.addText("這整套流程每個月都要跑一次。", {
    x: M, y: 5.5, w: CW, h: 0.6,
    fontFace: FONT, fontSize: 15, bold: true, color: ACCENT, align: "center", valign: "middle", margin: 0,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P19 舊流程痛點
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對內", "舊流程的痛點");
  const pains = [
    { h1: "五份以上", h2: "的 Excel", body: "同一件事的資料散在五份以上的表單裡，一份工作要開一堆檔案。", asset: "legacy-excel-01" },
    { h1: "每個月", h2: "都要重來", body: "每到月底就要重開報表、重設一次 Google 表單給人填。", asset: "legacy-form-01" },
    { h1: "手動算分", h2: "反覆貼上", body: "收回來的資料要手動計算分數，再一直複製貼上到給不同對象的 Excel。", asset: "legacy-excel-02" },
  ];
  const cw = (CW - 0.4 * 2) / 3;
  pains.forEach((p, i) => {
    const x = M + i * (cw + 0.4);
    s.addShape("roundRect", { x, y: 2.1, w: cw, h: 4.05, rectRadius: 0.06, fill: { color: BLUSH }, shadow: shadow() });
    s.addText(p.h1, {
      x: x + 0.25, y: 2.3, w: cw - 0.5, h: 0.5,
      fontFace: FONT, fontSize: 23, bold: true, color: ACCENT, align: "center", valign: "middle", margin: 0,
    });
    s.addText(p.h2, {
      x: x + 0.25, y: 2.78, w: cw - 0.5, h: 0.5,
      fontFace: FONT, fontSize: 23, bold: true, color: ACCENT, align: "center", valign: "middle", margin: 0,
    });
    s.addText(p.body, {
      x: x + 0.3, y: 3.35, w: cw - 0.6, h: 1.15,
      fontFace: FONT, fontSize: 13, color: NAVY, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.25,
    });
    image(s, p.asset, "舊流程截圖", { x: x + 0.3, y: 4.5, w: cw - 0.6, h: 1.4 });
  });
  s.addText("最花時間的不是判斷，是搬資料。", {
    x: M, y: 6.35, w: CW, h: 0.45,
    fontFace: FONT, fontSize: 15, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P20 我的解法
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對內", "我的解法：單一後台 × 三層權限");
  const tw = 6.3;
  const points = [
    ["整併", "把散在五份以上 Excel 的流程，整併進同一個網站、同一個後台。"],
    ["不搬資料", "資料不需要再搬移，只針對不同對象切換顯示內容。"],
    ["設定一次", "設定一次就長期適用，不必每月重做，直到當事人轉職或離職。"],
  ];
  points.forEach((p, i) => {
    const y = 2.15 + i * 1.28;
    s.addShape("roundRect", { x: M, y, w: tw, h: 1.12, rectRadius: 0.06, fill: { color: i === 2 ? CREAM : SOFT_BG } });
    numCircle(s, i + 1, { x: M + 0.28, y: y + 0.3, d: 0.52, fill: NAVY });
    s.addText(p[0], {
      x: M + 0.95, y: y + 0.14, w: tw - 1.2, h: 0.42,
      fontFace: FONT, fontSize: 16, bold: true, color: ACCENT, margin: 0, valign: "middle",
    });
    s.addText(p[1], {
      x: M + 0.95, y: y + 0.55, w: tw - 1.25, h: 0.5,
      fontFace: FONT, fontSize: 12.5, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
    });
  });
  s.addText("開發方式：全程使用 replit 等 AI 工具開發", {
    x: M, y: 6.1, w: tw, h: 0.55,
    fontFace: FONT, fontSize: 13.5, bold: true, color: NAVY, margin: 0, valign: "middle",
  });
  image(s, "mystery-01-login", "系統首頁／登入頁", { x: M + tw + 0.45, y: 2.15, w: CW - tw - 0.45, h: 4.5 });
  stampPage(s);
}

// ---------------------------------------------------------------- P21 權限設計
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對內", "權限設計：三種角色，三種畫面");
  s.addText("依涉入神秘客電訪的人員層級分為三種權限等級，不同帳號開放不同功能。", {
    x: M, y: 1.95, w: CW, h: 0.42,
    fontFace: FONT, fontSize: 14, color: MUTED, margin: 0,
  });
  const rows = [
    [th("層級"), th("大致角色"), th("開放範圍")],
    [td("第一層", { align: "center", bold: true, fill: { color: SOFT_BG } }), td("執行電訪與被檢核的專案組同仁"), td("執行與提交相關功能")],
    [td("第二層", { align: "center", bold: true, fill: { color: SOFT_BG } }), td("管理神秘客電訪的管理者"), td("出題、派發、評分、報表")],
    [td("第三層", { align: "center", bold: true, fill: { color: SOFT_BG } }), td("需要看結果的高級主管"), td("彙總報表與結果檢視")],
  ];
  table(s, rows, { x: M, y: 2.5, w: CW, colW: [1.7, 5.0, 5.233], rowH: 0.72, fontSize: 13.5 });
  image(s, "mystery-02-permission", "權限分層畫面", { x: M, y: 5.42, w: 5.6, h: 1.25 });
  s.addText("先有權限分層，才有辦法把資料集中——過去要拆成好幾份 Excel，本來就是為了讓不同的人看到不同的東西。", {
    x: M + 6.0, y: 5.42, w: CW - 6.0, h: 1.25,
    shape: "roundRect", rectRadius: 0.08, fill: { color: CREAM },
    fontFace: FONT, fontSize: 13.5, color: NAVY, margin: [10, 16, 10, 16], valign: "middle", lineSpacingMultiple: 1.3,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P22 / P23 / P24 功能
function featureSlide(title, feats, tail) {
  const s = newSlide(pres);
  header(s, "四、工作成果｜對內", title);
  const n = feats.length;
  const gap = 0.4;
  const cw = (CW - gap * (n - 1)) / n;
  feats.forEach((f, i) => {
    const x = M + i * (cw + gap);
    s.addShape("roundRect", { x, y: 2.1, w: cw, h: 3.9, rectRadius: 0.06, fill: { color: i % 2 ? SOFT_BG : LAV }, shadow: shadow() });
    numCircle(s, f.n, { x: x + 0.28, y: 2.32, d: 0.56, fill: NAVY, size: 17 });
    s.addText(f.title, {
      x: x + 0.95, y: 2.3, w: cw - 1.2, h: 0.6,
      fontFace: FONT, fontSize: 16, bold: true, color: NAVY, margin: 0, valign: "middle",
    });
    s.addText(f.body, {
      x: x + 0.3, y: 3.0, w: cw - 0.6, h: 0.85,
      fontFace: FONT, fontSize: 12.5, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
    });
    image(s, f.asset, f.title, { x: x + 0.3, y: 3.9, w: cw - 0.6, h: 1.85 });
  });
  if (tail) {
    s.addText(tail, {
      x: M, y: 6.15, w: CW, h: 0.72,
      shape: "roundRect", rectRadius: 0.08, fill: { color: CREAM },
      fontFace: FONT, fontSize: 14, color: NAVY, align: "center", valign: "middle", margin: 0,
    });
  }
  stampPage(s);
}

featureSlide(
  "功能 ①②③：出題 → 抽選派發 → 評分",
  [
    { n: 1, title: "設計新問題", body: "建立檢核用的題目。", asset: "mystery-03-question" },
    { n: 2, title: "抽選並分配專案組", body: "取代每月重設一次 Google 表單的工作。", asset: "mystery-04-assign" },
    { n: 3, title: "更簡單的分數評鑑介面", body: "取代手動計算與貼上分數。", asset: "mystery-05-scoring" },
  ],
  "這三項合起來，解決了「每個月都要重來一次」這個痛點。"
);

featureSlide(
  "功能 ④⑤⑥：報表匯出 → 檢討報告 → 申覆",
  [
    { n: 4, title: "報表匯出", body: "各對象所需的報表直接從系統產生，不用再複製貼上。", asset: "mystery-06-report" },
    { n: 5, title: "提交檢討報告", body: "未達滿分的經緯報告直接在系統內提交。", asset: "mystery-07-review" },
    { n: 6, title: "提交申覆系統", body: "對評分結果有疑義可線上申覆。", asset: "mystery-08-appeal" },
  ],
  "申覆功能不在任何需求文件上——是理解流程之後才發現需要的。"
);

featureSlide(
  "功能 ⑦⑧：帳號與身分組管理、通知系統",
  [
    { n: 7, title: "帳號與身分組管理", body: "創建帳號、創建身分組，以及帳號／身分組的移轉機制。人事異動時管理者自己就能處理，不必找工程師。", asset: "mystery-09-account" },
    { n: 8, title: "通知系統（全服／個人）", body: "分成全服公告與個人通知兩種，該知道的人在系統裡就會知道，不用另外發訊息提醒。", asset: "mystery-10-notify" },
  ],
  "這兩項是讓系統可以長期活下去的功能。"
);

// ---------------------------------------------------------------- P25 前後對比
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對內", "前後對比");
  const rows = [
    [th(""), th("舊流程"), { text: "神秘客系統", options: { bold: true, fill: { color: NAVY }, color: WHITE, align: "center" } }],
    [td("資料位置", { bold: true, fill: { color: SOFT_BG } }), td("分散在 5 份以上 Excel"), td("集中在單一後台", { fill: { color: "EFF7F3" } })],
    [td("每月作業", { bold: true, fill: { color: SOFT_BG } }), td("重開報表、重設 Google 表單"), td("設定一次長期適用", { fill: { color: "EFF7F3" } })],
    [td("分數計算", { bold: true, fill: { color: SOFT_BG } }), td("手動計算"), td("系統評鑑介面", { fill: { color: "EFF7F3" } })],
    [td("給不同對象", { bold: true, fill: { color: SOFT_BG } }), td("複製貼上到不同 Excel"), td("依權限切換顯示內容", { fill: { color: "EFF7F3" } })],
    [td("檢討與申覆", { bold: true, fill: { color: SOFT_BG } }), td("另外處理"), td("系統內提交", { fill: { color: "EFF7F3" } })],
    [td("人事異動", { bold: true, fill: { color: SOFT_BG } }), td("需重新整理表單"), td("帳號／身分組移轉機制", { fill: { color: "EFF7F3" } })],
  ];
  table(s, rows, { x: M, y: 2.1, w: CW, colW: [2.6, 4.6, 4.733], rowH: 0.58, fontSize: 13 });
  s.addText("改變的重點不是把工作做得更快，而是把「每個月都要重做一次」變成「設定一次就好」。", {
    x: M, y: 6.25, w: CW, h: 0.62,
    fontFace: FONT, fontSize: 15, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P26 成效與現況
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對內", "成效與現況");

  s.addShape("roundRect", { x: M, y: 2.1, w: 5.1, h: 2.65, rectRadius: 0.06, fill: { color: NAVY } });
  s.addText("約 90%", {
    x: M + 0.3, y: 2.35, w: 4.5, h: 1.15,
    fontFace: FONT, fontSize: 54, bold: true, color: WHITE, margin: 0, valign: "middle",
  });
  s.addText("推估節省的作業時間", {
    x: M + 0.3, y: 3.5, w: 4.5, h: 0.42,
    fontFace: FONT, fontSize: 16, bold: true, color: LAV, margin: 0,
  });
  s.addText("＊本數字為個人推估，尚未經實測驗證", {
    x: M + 0.3, y: 3.95, w: 4.5, h: 0.55,
    fontFace: FONT, fontSize: 12, italic: true, color: LAV_MID, margin: 0, lineSpacingMultiple: 1.2,
  });

  const info = [
    ["狀態", "已實際上線，目前使用中"],
    ["使用單位", "所有需執行神秘客訪問電話的專案組，以及神秘客電訪管理者"],
    ["開發方式", "全程使用 replit 等 AI 工具開發"],
  ];
  info.forEach((it, i) => {
    const y = 2.1 + i * 0.92;
    s.addShape("roundRect", { x: M + 5.5, y, w: CW - 5.5, h: 0.8, rectRadius: 0.06, fill: { color: i === 0 ? MINT : SOFT_BG } });
    s.addText(it[0], {
      x: M + 5.75, y, w: 1.4, h: 0.8,
      fontFace: FONT, fontSize: 13.5, bold: true, color: NAVY, margin: 0, valign: "middle",
    });
    s.addText(it[1], {
      x: M + 7.15, y, w: CW - 5.5 - 1.9, h: 0.8,
      fontFace: FONT, fontSize: 13, color: NAVY, margin: 0, valign: "middle", lineSpacingMultiple: 1.15,
    });
  });
  s.addText("【待補：上線日期、使用單位數、使用人數】", {
    x: M + 5.5, y: 4.95, w: CW - 5.5, h: 0.4,
    fontFace: FONT, fontSize: 12, color: ACCENT, margin: 0,
  });

  s.addText("推估依據：過去每月要開所有 Excel 報表、設一次新的 Google 表單、手動算分、再複製貼上給不同對象；現在全部在同一個後台完成，資料不用搬，也不需要每月重改。若主管認為需要，我可以實際量測作業時間，把真實數字補上。", {
    x: M, y: 5.4, w: CW, h: 1.15,
    shape: "roundRect", rectRadius: 0.08, fill: { color: CREAM },
    fontFace: FONT, fontSize: 13, color: NAVY, margin: [10, 18, 10, 18], valign: "middle", lineSpacingMultiple: 1.3,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P27 工時分佈
{
  const s = newSlide(pres);
  header(s, "四、工作成果", "三個月工時分佈");
  pieWithNotes(pres, s, {
    labels: ["對外課程教材撰寫", "對外課程授課與備課", "神秘客系統開發", "其他（會議、行政、學習）"],
    values: [40, 30, 25, 5],
    colors: [NAVY, TEAL, MINT, LAV_MID],
    noteTitle: "＊以上為推估值　【待補：實際百分比】",
    notes: [
      "寫教材要把主題徹底弄懂，花的時間比上課本身多。",
      "含備課、講義調整與課後回饋整理。",
      "從需求訪談、開發到上線。",
      "部門會議、行政作業與自主學習。",
    ],
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P28 學習心得
{
  const s = newSlide(pres);
  header(s, "五、學習心得", "學習心得");
  const cols = [
    { h: "收穫最多", t: "從使用者身上學需求", b: "不管是寫教材還是做系統，難的都不是技術，是搞懂對方到底要什麼。", tint: LAV },
    { h: "最有成就感", t: "系統真的被用起來", b: "不是 demo，是每個月都有人真的登入在用。做出能動的東西不難，做出有人用的東西完全不一樣。", tint: MINT },
    { h: "最挑戰", t: "判斷什麼不該做", b: "用 AI 開發很快，快到想到什麼功能都做得出來，反而要克制——這是使用者需要的，還是我想做的？", tint: BLUSH },
  ];
  const cw = (CW - 0.4 * 2) / 3;
  cols.forEach((c, i) => {
    const x = M + i * (cw + 0.4);
    s.addShape("roundRect", { x, y: 2.1, w: cw, h: 3.55, rectRadius: 0.06, fill: { color: c.tint }, shadow: shadow() });
    s.addText(c.h, {
      x: x + 0.3, y: 2.35, w: cw - 0.6, h: 0.45,
      fontFace: FONT, fontSize: 14, bold: true, color: "3F4468", align: "center", valign: "middle", margin: 0,
    });
    s.addText(c.t, {
      x: x + 0.28, y: 2.9, w: cw - 0.56, h: 1.0,
      fontFace: FONT, fontSize: 23, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.15,
    });
    s.addText(c.b, {
      x: x + 0.32, y: 4.0, w: cw - 0.64, h: 1.5,
      fontFace: FONT, fontSize: 13, color: NAVY, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.35,
    });
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P29 最重要的一課
{
  const s = newSlide(pres, NAVY_DEEP);
  s.addText("五、學習心得", {
    x: M, y: 1.15, w: 5.0, h: 0.5,
    fontFace: FONT, fontSize: 15, bold: true, color: LAV_MID, margin: 0, valign: "middle",
  });
  s.addText("三個月最重要的一課", {
    x: M, y: 1.75, w: CW, h: 0.7,
    fontFace: FONT, fontSize: 24, color: LAV, margin: 0, valign: "middle",
  });
  s.addText("從「做出功能」\n到「交付一個會被用的系統」", {
    x: M, y: 2.6, w: CW, h: 1.9,
    fontFace: FONT, fontSize: 40, bold: true, color: WHITE, margin: 0, valign: "middle", lineSpacingMultiple: 1.25,
  });
  const pts = [
    "功能做完 ≠ 交付完成",
    "要考慮：誰在用、人事異動了怎麼辦、被評分的人不服氣怎麼辦",
    "這些都不是技術問題，但不解決，系統就不會有人用",
  ];
  const cw = (CW - 0.35 * 2) / 3;
  pts.forEach((p, i) => {
    const x = M + i * (cw + 0.35);
    s.addShape("roundRect", { x, y: 4.85, w: cw, h: 1.5, rectRadius: 0.06, fill: { color: "2E3266" } });
    s.addText(p, {
      x: x + 0.28, y: 4.85, w: cw - 0.56, h: 1.5,
      fontFace: FONT, fontSize: 13.5, color: LAV, margin: 0, valign: "middle", lineSpacingMultiple: 1.3,
    });
  });
  stampPage(s, "7A80AE");
}

// ---------------------------------------------------------------- P30 職涯規劃
{
  const s = newSlide(pres);
  header(s, "六、職涯規劃", "職涯規劃");
  const cols = [
    { tag: "短期", sub: "半年", tint: LAV, items: ["把神秘客系統的實際效益量測出來，並依回饋持續優化", "參與四門課的實際開課，依學員回饋修訂教材"] },
    { tag: "中期", sub: "1 年", tint: CREAM, items: ["具備獨立負責一條產品線的能力，從需求訪談到上線", "累積 AI 應用在不同部門的落地經驗，形成可複製的方法"] },
    { tag: "長期", sub: "3～5 年", tint: MINT, items: ["成為能同時掌握對外課程產品與對內系統落地的產品負責人", "讓 AI 導入在公司內部成為標準做法，而不是個案"] },
  ];
  const cw = (CW - 0.4 * 2) / 3;
  cols.forEach((c, i) => {
    const x = M + i * (cw + 0.4);
    s.addShape("roundRect", { x, y: 2.15, w: cw, h: 3.4, rectRadius: 0.06, fill: { color: c.tint }, shadow: shadow() });
    s.addText(c.tag, {
      x: x + 0.3, y: 2.35, w: cw - 0.6, h: 0.62,
      fontFace: FONT, fontSize: 26, bold: true, color: NAVY, margin: 0, valign: "middle",
    });
    s.addText(c.sub, {
      x: x + 0.3, y: 2.95, w: cw - 0.6, h: 0.4,
      fontFace: FONT, fontSize: 14, color: "4A4F7A", margin: 0, valign: "middle",
    });
    s.addText(
      c.items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j !== c.items.length - 1 } })),
      {
        x: x + 0.32, y: 3.45, w: cw - 0.64, h: 1.95,
        fontFace: FONT, fontSize: 13, color: NAVY, margin: 0, valign: "top", paraSpaceAfter: 10, lineSpacingMultiple: 1.3,
      }
    );
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P31 對公司的建議
{
  const s = newSlide(pres);
  header(s, "七、對公司的建議", "對公司的建議｜新人入職與知識傳承");
  const sugs = [
    ["新人報告的目標，建議由主管與新人共同確認", "目前的十二週目標表是所有新人共用版本，與實際派工內容常有落差。建議到職第一週由主管與新人一起確認個人化目標，中間再對一次。"],
    ["把各部門的重複性作業做成一份「待數位化清單」", "神秘客這套流程原本不在任何待辦清單上，是碰巧發現的。若各部門能列出自己每月重複的作業，創新研發部就能有系統地評估導入順序。"],
    ["內部工具做完後，建立交接與維護機制", "系統上線只是開始。建議每套內部工具都要有指定的維護窗口與使用說明，避免變成只有開發者才懂的黑箱。"],
  ];
  sugs.forEach((sg, i) => {
    const y = 2.05 + i * 1.5;
    s.addShape("roundRect", { x: M, y, w: CW, h: 1.32, rectRadius: 0.06, fill: { color: [LAV, CREAM, SOFT_BG][i] } });
    numCircle(s, i + 1, { x: M + 0.32, y: y + 0.4, d: 0.55, fill: NAVY, size: 17 });
    s.addText(sg[0], {
      x: M + 1.05, y: y + 0.17, w: CW - 1.4, h: 0.45,
      fontFace: FONT, fontSize: 16.5, bold: true, color: NAVY, margin: 0, valign: "middle",
    });
    s.addText(sg[1], {
      x: M + 1.05, y: y + 0.63, w: CW - 1.4, h: 0.6,
      fontFace: FONT, fontSize: 12.5, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
    });
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P32 結尾
{
  const s = newSlide(pres, NAVY_DEEP);
  s.addText("謝謝您的聆聽！", {
    x: M, y: 2.9, w: CW, h: 1.5,
    fontFace: FONT, fontSize: 54, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0,
  });
  s.addText("產品工程師－黃家朗　｜　創新研發部", {
    x: M, y: 4.5, w: CW, h: 0.5,
    fontFace: FONT, fontSize: 16, color: LAV_MID, align: "center", margin: 0,
  });
  stampPage(s, "7A80AE");
}

// ---------------------------------------------------------------- 講稿（備註欄）

const NOTES = [
  // P1
  `各位主管、各位同事早安，我是創新研發部的產品工程師黃家朗。\n今天要跟大家報告我到職這三個月的工作內容與成果。\n報告大概三十分鐘，中間如果有問題可以隨時打斷我，最後也會留 Q&A 時間。`,
  // P2
  `今天分七個部分。\n前面三段講我是誰、以及這三個月我實際在做什麼；\n第四段是重點，會花最多時間講工作成果，分成對外和對內兩條線；\n最後是心得、職涯規劃，還有幾個給公司的建議。`,
  // P3
  `先簡單介紹我自己。我是【待補：學校科系】畢業，\n進程曦之前有一段工作經驗，在【待補：前職】。\n我不是一畢業就做 AI 的，是工作一段時間之後才轉進 AI 跟教育訓練這個領域。\n這個轉職經驗其實對我現在的工作滿有幫助的，\n因為我很知道一個沒有技術背景的上班族，第一次接觸 AI 工具的時候會卡在哪裡——\n這剛好就是我現在在做的事：把 AI 工具教給不懂技術的人。`,
  // P4
  `這是我的工作歷程。\n【待補：這裡補一句前職在做什麼】。\n今年四月十號到職程曦，進創新研發部擔任產品工程師，到今天大概三個多月。`,
  // P5
  `這是我給自己的五個特質，每一個後面我都有對應的實際例子，等一下都會講到。\n佔比最高的是解決力跟自學力。\n解決力的意思是，我看到問題的反應不是回報給主管就結束，而是直接把工具做出來——\n等一下講的神秘客系統就是這樣來的。\n自學力是說，那套系統用到的東西有很多我原本不會，我是靠 AI 工具邊做邊學把它做完的。\n系統思維、換位思考、交付到底這三個，也都會在後面的內容裡出現，\n這裡先不展開，各位看後面的成果就會知道我為什麼這樣寫。`,
  // P6
  `公司精神是「愛、新、勤、誠」，\n合起來就是那句話：勤勞地用創新的方式，誠懇地愛人。\n我不想只是把這四個字唸過去，所以下一頁我整理了這三個月裡，\n我自己覺得有實際對應到這四個字的四件事。`,
  // P7
  `「新」的部分，神秘客系統我是全程用 replit 這類 AI 工具開發的，\n這在部門裡算是比較新的做法。\n「勤」的部分，我認為真正的勤勞不是每個月都認真重做一次，\n而是想辦法讓它只需要做一次——這也是那套系統的核心。\n「愛」的部分，那套系統分三種權限，\n是因為我去理解了打電話的人、管理的人、看報表的主管，各自需要什麼、不需要看到什麼。\n最後「誠」，等一下我會報一個效益數字，\n我要先說那是我自己的推估，不是實測出來的，我會誠實標註清楚。`,
  // P8
  `這三個月我把它分成三個階段。\n前四週是熟悉期，主要是搞懂公司在做什麼、我們部門的產品線長什麼樣子，\n同時就開始投入對外課程的教材製作。\n第五到第八週是產出期，教材開始有成品，我也開始上台當講師。\n第九到十二週是落地期，重心轉到對內，\n從需求訪談一路做到神秘客系統實際上線。\n這三段其實有一個共同點：都是在做「把 AI 變成別人真的用得到的東西」。`,
  // P9
  `用比重來看的話，我的工作大概是對外七、對內三。\n對外這條線佔大部分，就是我們部門的教育訓練產品線，四門 AI 應用課；\n對內這條線雖然只佔三成，但它是一個完整的系統，\n從訪談需求到寫完上線都是我做的，所以等一下我會花比較多時間講它。\n一句話總結我這三個月：對外是把 AI 變成可以賣的課，對內是把 AI 變成可以用的系統。`,
  // P10
  `接下來進入重點，工作成果。先講對外的四門課。`,
  // P11
  `對外這條線總共四門課，我在裡面有兩種角色。\n第一、二門是我負責寫教材，從課程架構、內容到範例都是我做的。\n第三、四門教材是同事寫的，我負責上台授課。\n我覺得這個組合對新人來說滿好的——\n寫教材讓我必須把一個主題徹底搞懂才寫得出來，\n上台授課則讓我馬上知道學員會在哪裡卡住、哪裡聽不懂。\n這兩件事其實是互相回饋的。`,
  // P12
  `第一門是「AI 時代的商業企劃思維與簡報製作工具」。\n這門課的重點其實不在工具操作，而在思維——\n很多人以為 AI 簡報工具就是把字丟進去讓它排版，\n但真正的價值是在前面的企劃思考階段就讓 AI 參與。\n所以我在寫教材的時候，是先把一個企劃案從無到有的流程拆開，\n再去看每一段可以怎麼用 AI，而不是反過來從工具功能表開始教。\n\n（時間不夠時本頁可略過）`,
  // P13
  `第二門是 AI 圖片生成課，一樣是我寫教材。\n這門課最難處理的地方是，圖片生成的結果不穩定，\n同樣一句提示詞，這次跟下次出來的東西可能差很多。\n所以我在教材裡花了不少篇幅在講「怎麼描述」跟「怎麼修正」，\n而不是只給一堆可以複製貼上的提示詞範本——\n因為範本會過期，方法不會。`,
  // P14
  `第三、四門是我上台講的。\n這兩門課對我來說有一個額外的意義：\nvibe-coding 這門課教的方法，就是我後來開發神秘客系統實際用的方法。\n我等於是先教了別人怎麼用 AI 寫程式，然後自己拿這套方法去做了一個真的上線的系統。\n我覺得這件事讓我在台上講的時候比較有底氣，\n因為我講的不是從教材上讀來的，是我真的做過的。`,
  // P15
  `這幾張是教材的實際畫面，給大家看一下我們課程的呈現方式。\n如果主管想看完整版本，我課後可以再單獨提供。\n\n（時間不夠時本頁可略過）`,
  // P16
  `對外這條線目前的狀態是：四門課的教材都已經完成，具備對外開課的條件，\n接下來就是等對外開賣。\n我自己覺得教材完成只是第一步，真正會讓課變好的是實際開課之後的學員回饋，\n所以如果後續有招生跟開課，我希望能持續參與，\n依照回饋去修教材、補案例，甚至往下做進階課。`,
  // P17
  `接下來是對內這條線，也是我這三個月最想跟大家報告的一件事。`,
  // P18
  `先說明業務背景，因為不是每個部門都熟悉這件事。\n程曦是承接客服外包的公司，所以服務品質必須被檢核。\n做法是各個專案組之間每個月隨機互相打電話，\n用腳本上的常見問題去問對方，\n檢驗接聽的同仁有沒有用符合禮儀、符合標準起手台詞的方式，說出正確的答案。\n這種電話我們叫做神秘客訪問電話。\n打完之後，各組要寫抽查成果跟分數交給主管，\n主管再彙整出報表、還有未達滿分的經緯報告，\n然後把分數向下發給各專案組、向上呈報給高級主管。\n這整套流程是每個月都要跑一次的。`,
  // P19
  `這套流程原本是這樣跑的。\n第一，資料是散的。做同一件事，要開五份甚至更多份的 Excel 表單。\n第二，每個月都要重來一次。每到月底就要重開所有報表，\n還要重新設定一份 Google 表單發給大家填。\n第三，收回來之後要手動算分數，\n然後一直複製貼上到要給不同對象的 Excel 裡面——\n給專案組一份、給主管一份、給高級主管又一份。\n這裡面最花時間的其實不是判斷，是搬資料。`,
  // P20
  `我的解法核心只有一句話：東西都放在同一個後台，用權限決定誰看得到什麼。\n這樣做的好處是，資料不需要再搬移了。\n過去要複製貼上到三份不同的 Excel，是因為三個對象要看不同的東西；\n現在資料只有一份，我只要針對不同對象改變「顯示的內容」就好。\n而且這個設定不是每個月都要改，是設定一次就一直適用，\n直到當事人轉職或離職才需要動。\n開發的部分，我是全程用 replit 這類 AI 工具做的。`,
  // P21
  `權限的部分，我依照會涉及神秘客電訪的人的層級，分成三種等級，\n不同的登入帳號會開放不同的功能。\n這件事聽起來只是技術設定，但它其實是整套系統能成立的關鍵。\n因為過去要拆成好幾份 Excel，本來就是為了讓不同的人看到不同的東西；\n如果我只是把五份表單合成一份而沒有權限控制，那反而會出事。\n所以是先有權限分層，才有辦法把資料集中。`,
  // P22
  `系統裡總共有八項功能，我分三頁講。\n前三項是流程的前半段。\n第一是設計新問題，把檢核用的題目建進系統。\n第二是抽選問題並分配到指定的專案組——這一項直接取代了過去每個月重設 Google 表單的工作。\n第三是評分介面，讓評分的人在同一個地方打分數，不用再自己在 Excel 上算。\n這三項合起來，就把「每月重來一次」這個痛點解決掉了。`,
  // P23
  `中間三項是流程的後半段。\n第四是報表匯出，需要給不同對象的報表直接從系統產生，不用再複製貼上。\n第五是檢討報告，過去未達滿分要另外寫的經緯報告，現在直接在系統裡提交。\n第六是申覆。這一項是我做完前面之後才發現需要的——\n因為分數會影響到專案組，如果被評分的人覺得判定有問題，\n應該要有一個正式的管道處理，而不是私下用訊息吵。\n所以我加了申覆流程進去。`,
  // P24
  `最後兩項是讓系統可以長期活下去的功能。\n第七是帳號跟身分組的建立，還有移轉機制。\n這一項很重要，因為公司會有人轉職、有人離職，\n如果每次人事異動都要工程師進資料庫改，那這套系統遲早會沒人維護。\n有了移轉機制，管理者自己就能處理。\n第八是通知系統，分成全服公告跟個人通知兩種，\n讓該知道的人在系統裡就會知道，不用另外再發訊息提醒。\n\n（時間不夠時本頁可略過）`,
  // P25
  `這頁是前後對比，我覺得最能說明這套系統的價值。\n左邊是舊流程，右邊是現在。\n各位可以看到，改變的重點不是把工作做得更快，\n而是把「每個月都要重做一次」的工作，變成「設定一次就好」。\n尤其最下面那一列，人事異動——\n過去有人離職，那些表單就要重新整理一次；\n現在管理者自己在後台移轉就結束了。`,
  // P26
  `這套系統目前的狀態是已經上線，實際在使用中。\n使用的單位是所有需要打神秘客訪問電話的單位，以及負責管理的管理者。\n效益的部分我要誠實說明：\n這個「節省近 90%」是我個人的推估，不是實測出來的數字。\n我的推估依據是，過去每個月月底要開所有 Excel 報表、\n設一次新的 Google 表單、手動算分數、再一直複製貼上給不同對象；\n現在這些全部在同一個後台完成，資料不用搬，而且不需要每個月都改。\n如果主管覺得需要，我可以後續實際去量測作業時間，把真實數字補上。`,
  // P27
  `這是我三個月的工時分佈，先說明這是推估值。\n對外課程加起來大概七成，其中教材撰寫佔比最高，\n因為寫教材要把一個主題徹底弄懂，花的時間比上課本身多很多。\n神秘客系統大概兩成五。\n這個比例跟前面講的對外七、對內三是一致的。`,
  // P28
  `講一下心得，分三個部分。\n收穫最多的是「從使用者身上學需求」。\n不管是寫教材還是做系統，我發現難的都不是技術，\n是搞懂對方到底要什麼。\n申覆那個功能就是例子——它不在任何一份需求文件上，\n是我去理解這個流程之後才發現需要的。\n最有成就感的是系統真的被用起來了。\n做出一個能動的東西不難，但做出一個每個月都有人真的登入去用的東西，感覺完全不一樣。\n最挑戰的是判斷什麼不該做。\n用 AI 開發真的很快，快到你想到什麼功能都做得出來，\n所以反而要克制，要一直問自己這個功能是使用者需要的，還是我想做的。`,
  // P29
  `如果要我只講一件這三個月最重要的收穫，是這個：\n從「做出功能」，變成「交付一個會被用的系統」。\n我剛到職的時候會覺得，功能寫完就是做完了。\n但實際做下來才知道，功能做完只是一半。\n你還要想：誰在用、有人離職了怎麼辦、被評分的人不服氣怎麼辦。\n這些都不是技術問題，可是不解決，系統做得再好也不會有人用。\n我覺得這是我從工程師往產品的方向走，很關鍵的一步。`,
  // P30
  `職涯規劃分短中長期。\n短期半年，我想把神秘客系統的實際效益量測出來——\n就是前面說的那個 90%，我想把它變成真實數字。\n同時參與四門課的實際開課，依學員回饋修教材。\n中期一年，我希望能獨立負責一條產品線，從需求訪談一路做到上線。\n長期三到五年，我的目標是成為能同時掌握對外課程產品跟對內系統落地的產品負責人，\n讓 AI 導入在公司裡變成一件標準的事，而不是每次都要重新來過的個案。`,
  // P31
  `最後是給公司的三個建議，主題是新人入職跟知識傳承。\n第一個，新人報告的目標建議由主管跟新人一起確認。\n我理解目前的十二週目標表是為了讓新人有依循，\n但它是所有新人共用的版本，跟實際派的工作常常會有落差。\n我的建議是到職第一週由主管跟新人一起確認個人化的目標，中間再對一次，\n這樣新人會更清楚自己被期待什麼。\n第二個，建議各部門列一份自己每月重複作業的清單。\n神秘客這套流程本來不在任何待辦清單上，\n老實說是碰巧發現的。如果有這樣一份清單，\n我們部門就可以有系統地評估哪些該先做，而不是等人來說。\n第三個，內部工具做完之後要有交接跟維護機制。\n我自己做完神秘客系統之後最擔心的就是這件事——\n系統上線只是開始，如果沒有指定的維護窗口跟使用說明，\n它遲早會變成只有我看得懂的黑箱。`,
  // P32
  `以上就是我這三個月的報告，謝謝各位的聆聽。\n有任何問題都歡迎提出。`,
];

pres.slides.forEach((s, i) => {
  if (NOTES[i]) s.addNotes(NOTES[i]);
});

if (pres.slides.length !== NOTES.length) {
  console.warn(`⚠️  投影片 ${pres.slides.length} 頁，但講稿有 ${NOTES.length} 則——請檢查對應關係。`);
}

pres.writeFile({ fileName: OUT }).then(() => {
  console.log(`✅ 已產出：${OUT}`);
  console.log(`   共 ${pres.slides.length} 頁`);
  const missing = [];
  const check = (b) => { if (!findAsset(b)) missing.push(b); };
  [
    "mystery-01-login", "mystery-02-permission", "mystery-03-question", "mystery-04-assign",
    "mystery-05-scoring", "mystery-06-report", "mystery-07-review", "mystery-08-appeal",
    "mystery-09-account", "mystery-10-notify",
    "legacy-excel-01", "legacy-excel-02", "legacy-form-01",
    "course-01-slide", "course-02-slide", "course-03-slide", "course-04-slide",
    "course-showcase-01", "course-showcase-02", "course-showcase-03",
  ].forEach(check);
  if (missing.length) {
    console.log(`\n📷 尚缺 ${missing.length} 張圖（目前顯示為灰底佔位方塊）：`);
    missing.forEach((m) => console.log(`   - report/assets/${m}.png`));
    console.log(`   放圖說明見 report/assets/README.md`);
  }
});
