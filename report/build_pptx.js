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
function card(slide, { x, y, w, h, title, body, tint = SOFT_BG, titleColor = NAVY, titleSize = 17, bodySize = 13.5 }) {
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
    fontFace: FONT, fontSize: bodySize, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
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
  // 項目多的時候壓縮行距，避免撞到頁碼
  const step = labels.length >= 5 ? 0.8 : 0.86;
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
    ["我的三階段工作主軸", "這四個多月實際在做什麼"],
    ["學習與工作成果", "對外四門課｜對內兩套系統"],
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
  // 右側人像，3:4 直式；左側 2×2 資料卡
  const pw = 3.26;
  const ph = 4.35;
  const px = M + CW - pw;
  image(s, "profile", "個人照（直式 3:4）", { x: px, y: 2.05, w: pw, h: ph });

  const cw = (CW - pw - 0.45 - 0.4) / 2;
  const ch = (ph - 0.3) / 2;
  card(s, {
    x: M, y: 2.05, w: cw, h: ch, tint: LAV, bodySize: 12.5,
    title: "學歷",
    body: "國立清華大學 科技管理學院學士班\n管理＋經濟雙專長\n2018.09 – 2023.06",
  });
  card(s, {
    x: M + cw + 0.4, y: 2.05, w: cw, h: ch, tint: SOFT_BG, bodySize: 12.5,
    title: "到職前經歷",
    body: "台北數位廣告股份有限公司\n產品企劃部｜產品企劃專員\n2024.10 – 2026.03",
  });
  card(s, {
    x: M, y: 2.05 + ch + 0.3, w: cw, h: ch, tint: SOFT_BG, bodySize: 12.5,
    title: "現職",
    body: "程曦資訊整合股份有限公司\n創新研發部｜產品工程師\n2026.04 –",
  });
  card(s, {
    x: M + cw + 0.4, y: 2.05 + ch + 0.3, w: cw, h: ch, tint: CREAM, bodySize: 12.5,
    title: "興趣",
    body: "音樂（龐克搖滾／重金屬）、攝影、旅行\nTRPG 主持與劇本創作\n工作之外也用 AI 做數據分析與翻譯",
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P4 工作歷程
{
  const s = newSlide(pres);
  header(s, "一、自我介紹", "工作歷程");

  // 時間軸底線
  s.addShape("rect", { x: M + 0.35, y: 3.08, w: CW - 0.7, h: 0.045, fill: { color: LAV } });

  const nodes = [
    { t: "2020.05", n: "攸你資訊", r: "共同創辦人／PM", d: "APP 產品從 0 到 1\n領導 5 人團隊" },
    { t: "2024.05", n: "攸你資訊", r: "PM／AI 提示工程師", d: "用 Custom GPTs\n做內部工具" },
    { t: "2024.10", n: "台北數位廣告", r: "產品企劃專員", d: "AI 工具調研\n與內部培訓" },
    { t: "2026.04", n: "程曦資訊", r: "創新研發部 產品工程師", d: "對外課程\n＋ 對內系統", now: true },
  ];
  const colW = CW / nodes.length;
  nodes.forEach((nd, i) => {
    const cx = M + colW * i + colW / 2;
    const color = nd.now ? ACCENT : LAV_MID;
    s.addShape("ellipse", { x: cx - 0.16, y: 2.94, w: 0.33, h: 0.33, fill: { color } });
    s.addText(nd.t, {
      x: cx - colW / 2, y: 2.32, w: colW, h: 0.42,
      fontFace: FONT, fontSize: 16, bold: true, color: nd.now ? ACCENT : NAVY, align: "center", valign: "middle", margin: 0,
    });
    s.addText(nd.n, {
      x: cx - colW / 2, y: 3.4, w: colW, h: 0.44,
      fontFace: FONT, fontSize: 15, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0,
    });
    s.addText(nd.r, {
      x: cx - colW / 2 + 0.1, y: 3.84, w: colW - 0.2, h: 0.42,
      fontFace: FONT, fontSize: 12, color: "4A4F7A", align: "center", valign: "top", margin: 0,
    });
    s.addText(nd.d, {
      x: cx - colW / 2 + 0.1, y: 4.32, w: colW - 0.2, h: 0.8,
      fontFace: FONT, fontSize: 11.5, color: MUTED, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.25,
    });
  });

  s.addText("轉職動機：成為連結業務需求與工程開發的 AI 實踐者，為下一個十年的「辦公室 AI 化」做準備。", {
    x: M, y: 5.4, w: CW, h: 0.78,
    shape: "roundRect", rectRadius: 0.08, fill: { color: CREAM },
    fontFace: FONT, fontSize: 15, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P5 個人特質
{
  const s = newSlide(pres);
  header(s, "一、自我介紹", "個人特質");
  const traits = [
    ["解決力", "不只回報問題，直接把工具做出來。", "神秘客系統：把五份以上 Excel 整併成單一後台。", LAV],
    ["自學力", "沒學過的東西，用 AI 工具邊做邊學做完。", "以 replit 開發整套神秘客系統。", SOFT_BG],
    ["換位思考", "先搞懂對方要什麼，再動手。", "PRD 機器人：用大量反問把抽象概念逼成規格。", CREAM],
    ["系統思維", "看見五份 Excel 背後其實是同一件事。", "先做權限分層，資料才有辦法集中。", SOFT_BG],
    ["交付到底", "做到有人真的在用，不是 demo。", "神秘客已上線；PRD 機器人 10 位跨部門同事測試中。", LAV],
  ];
  traits.forEach((t, i) => {
    const y = 2.0 + i * 0.92;
    s.addShape("roundRect", { x: M, y, w: CW, h: 0.8, rectRadius: 0.06, fill: { color: t[3] } });
    s.addText(t[0], {
      x: M + 0.3, y, w: 1.75, h: 0.8,
      fontFace: FONT, fontSize: 18, bold: true, color: NAVY, margin: 0, valign: "middle",
    });
    s.addText(t[1], {
      x: M + 2.15, y, w: 4.6, h: 0.8,
      fontFace: FONT, fontSize: 13.5, color: NAVY, margin: 0, valign: "middle",
    });
    s.addText(t[2], {
      x: M + 7.0, y, w: CW - 7.3, h: 0.8,
      fontFace: FONT, fontSize: 12.5, color: "4A4F7A", margin: 0, valign: "middle",
    });
  });
  s.addText("每一個特質後面都有對應的實際例子，後面會逐一講到。", {
    x: M, y: 6.65, w: CW, h: 0.4,
    fontFace: FONT, fontSize: 12.5, color: MUTED, margin: 0,
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
    ["愛", "PRD 機器人不是替人寫，是用反問幫人把自己的想法講清楚。", BLUSH],
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
      td("第 1 個月", { align: "center" }),
      td("熟悉公司業務與 AI 教育訓練產品線"),
      td("認識部門定位與產品線；投入對外課程教材製作"),
    ],
    [
      td("二、產出", { bold: true, fill: { color: SOFT_BG } }),
      td("第 2～3 個月", { align: "center" }),
      td("對外課程教材產出與內部試教"),
      td("完成 2 門課教材撰寫；擔任 2 門課講師並完成內部試教"),
    ],
    [
      td("三、落地", { bold: true, fill: { color: SOFT_BG } }),
      td("第 4 個月至今", { align: "center", bold: true, color: ACCENT }),
      td("對內流程數位化，從需求訪談到系統上線"),
      td("神秘客系統開發並上線；PRD 機器人開發並進入跨部門測試"),
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
      lines: ["神秘客訪問電話系統（已上線）", "PRD 機器人（跨部門測試中）", "其他零星支援與小工具"],
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
    [th("課程"), th("我的角色"), th("目前狀態")],
    [td("① AI 時代的商業企劃思維與簡報製作工具"), td("教材撰寫", { align: "center", bold: true, color: ACCENT }), td("已內部試教｜教材修正中", { align: "center" })],
    [td("② AI 圖片生成課"), td("教材撰寫", { align: "center", bold: true, color: ACCENT }), td("已內部試教｜教材修正中", { align: "center" })],
    [td("③ AI 輔助 UI/UX 與原型設計"), td("擔任講師", { align: "center", bold: true, color: TEAL }), td("已內部試教｜教材修正中", { align: "center" })],
    [td("④ AI 輔助 vibe-coding 程式開發"), td("擔任講師", { align: "center", bold: true, color: TEAL }), td("已內部試教｜教材修正中", { align: "center" })],
  ];
  table(s, rows, { x: M, y: 2.1, w: CW, colW: [6.0, 2.2, 3.733], rowH: 0.72, fontSize: 13.5 });
  s.addText("四門課皆已完成內部試教，尚未對外開課；教材依講師風格持續修正中，尚未定稿。", {
    x: M, y: 5.85, w: CW, h: 0.5,
    fontFace: FONT, fontSize: 13.5, bold: true, color: ACCENT, margin: 0, valign: "middle",
  });
  s.addText("寫教材要把主題徹底搞懂才寫得出來；上台試教才知道學員會卡在哪裡——這兩件事互相回饋。　【待補：各課時數與面向對象】", {
    x: M, y: 6.35, w: CW, h: 0.55,
    fontFace: FONT, fontSize: 13, color: MUTED, margin: 0, valign: "middle",
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
        { text: "已完成內部試教，尚未對外開課", options: { bold: true, color: ACCENT, breakLine: true } },
        { text: "【待補：時數／對象／場次】", options: { color: ACCENT } },
      ],
      { x: it.x + 0.3, y: 3.05, w: cw - 0.6, h: 1.2, fontFace: FONT, fontSize: 13, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.3 }
    );
    image(s, it.asset, `${it.n} 教材頁`, { x: it.x + 0.3, y: 4.3, w: cw - 0.6, h: 0.68 });
  });
  s.addText("vibe-coding 這門課教的方法，就是我後來開發神秘客系統與 PRD 機器人實際用的方法——先教別人怎麼用 AI 寫程式，然後自己拿這套方法做出兩個真的有人在用的系統。", {
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
    { title: "已完成", body: "四門課皆完成內部試教\n教材主體已具備", tint: MINT, tc: NAVY },
    { title: "進行中", body: "教材依講師的風格與側重點\n持續修正，尚未定稿", tint: CREAM, tc: NAVY },
    { title: "尚未開始", body: "對外開課與招生\n【待補：預計開賣時間】", tint: LAV, tc: NAVY },
  ];
  items.forEach((it, i) => {
    card(s, {
      x: M + i * (cw + 0.4), y: 2.15, w: cw, h: 2.1,
      title: it.title, body: it.body, tint: it.tint, titleColor: it.tc, titleSize: 19,
    });
  });
  s.addText("教材為什麼還沒定稿：同一份教材，不同講師講起來不一樣。有人擅長帶實作、有人擅長講原理，側重點不同，教材就要跟著調整。與其先定稿再讓講師自己想辦法，不如讓教材配合講師——所以「還在修」不是進度落後，是這門課的必經過程。", {
    x: M, y: 4.7, w: CW, h: 1.35,
    shape: "roundRect", rectRadius: 0.08, fill: { color: SOFT_BG },
    fontFace: FONT, fontSize: 14, color: NAVY, margin: [12, 18, 12, 18], valign: "middle", lineSpacingMultiple: 1.3,
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
  s.addText("內部工具與系統", {
    x: M, y: 2.85, w: CW, h: 1.25,
    fontFace: FONT, fontSize: 52, bold: true, color: WHITE, margin: 0, valign: "middle",
  });
  s.addText("神秘客訪問電話系統（已上線）　｜　PRD 機器人（跨部門測試中）　｜　其他零星支援", {
    x: M, y: 4.3, w: CW, h: 0.5,
    fontFace: FONT, fontSize: 15, color: LAV_MID, margin: 0,
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
  const gap = n >= 4 ? 0.3 : 0.4;
  const cw = (CW - gap * (n - 1)) / n;
  // 欄數越多，字級與圖框要跟著縮，避免擠爆
  const wide = n <= 3;
  const titleSize = wide ? 16 : 14;
  const bodySize = wide ? 12.5 : 11.5;
  feats.forEach((f, i) => {
    const x = M + i * (cw + gap);
    s.addShape("roundRect", { x, y: 2.1, w: cw, h: 3.9, rectRadius: 0.06, fill: { color: i % 2 ? SOFT_BG : LAV }, shadow: shadow() });
    numCircle(s, f.n, { x: x + 0.24, y: 2.3, d: 0.5, fill: NAVY, size: 15 });
    s.addText(f.title, {
      x: x + 0.84, y: 2.26, w: cw - 1.05, h: 0.6,
      fontFace: FONT, fontSize: titleSize, bold: true, color: NAVY, margin: 0, valign: "middle",
    });
    s.addText(f.body, {
      x: x + 0.26, y: 2.95, w: cw - 0.52, h: 1.0,
      fontFace: FONT, fontSize: bodySize, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
    });
    image(s, f.asset, f.title, { x: x + 0.26, y: 4.0, w: cw - 0.52, h: 1.75 });
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
  "八大功能 ①–④：出題 → 派發 → 評分 → 報表",
  [
    { n: 1, title: "設計新問題", body: "建立檢核用的題目。", asset: "mystery-03-question" },
    { n: 2, title: "抽選並分配專案組", body: "取代每月重設一次 Google 表單的工作。", asset: "mystery-04-assign" },
    { n: 3, title: "分數評鑑介面", body: "取代手動計算與貼上分數。", asset: "mystery-05-scoring" },
    { n: 4, title: "報表匯出", body: "各對象所需的報表直接產生，不用再複製貼上。", asset: "mystery-06-report" },
  ],
  "這四項合起來，解決了「每個月都要重來一次」這個痛點。"
);

featureSlide(
  "八大功能 ⑤–⑧：檢討 → 申覆 → 帳號管理 → 通知",
  [
    { n: 5, title: "提交檢討報告", body: "未達滿分的經緯報告直接在系統內提交。", asset: "mystery-07-review" },
    { n: 6, title: "提交申覆系統", body: "對評分結果有疑義可線上申覆，不必私下爭執。", asset: "mystery-08-appeal" },
    { n: 7, title: "帳號與身分組管理", body: "含移轉機制。人事異動時管理者自己就能處理，不必找工程師。", asset: "mystery-09-account" },
    { n: 8, title: "通知系統", body: "分全服公告與個人通知，該知道的人在系統裡就會知道。", asset: "mystery-10-notify" },
  ],
  "申覆功能不在任何需求文件上——是理解流程之後才發現需要的；⑦⑧ 則是讓系統長期活下去的關鍵。"
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

// ---------------------------------------------------------------- P26 PRD 機器人：問題
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對內", "PRD 機器人：解決什麼問題");

  s.addText("想法很抽象，但開發需要規格。從「我想做一個 X」到一份能動工的 PRD，中間那段最卡。", {
    x: M, y: 1.95, w: CW, h: 0.5,
    fontFace: FONT, fontSize: 15, bold: true, color: NAVY, margin: 0, valign: "middle",
  });

  // 過去 vs 現在
  const cw = (CW - 0.45) / 2;
  s.addShape("roundRect", { x: M, y: 2.6, w: cw, h: 1.75, rectRadius: 0.06, fill: { color: BLUSH } });
  s.addText("過去", {
    x: M + 0.3, y: 2.78, w: cw - 0.6, h: 0.42,
    fontFace: FONT, fontSize: 16, bold: true, color: ACCENT, margin: 0, valign: "middle",
  });
  s.addText("靠有經驗的人一次次開會追問，或是自己對著空白文件硬寫——品質看人，也很花時間。", {
    x: M + 0.3, y: 3.22, w: cw - 0.6, h: 0.95,
    fontFace: FONT, fontSize: 13, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  s.addShape("roundRect", { x: M + cw + 0.45, y: 2.6, w: cw, h: 1.75, rectRadius: 0.06, fill: { color: MINT } });
  s.addText("現在", {
    x: M + cw + 0.75, y: 2.78, w: cw - 0.6, h: 0.42,
    fontFace: FONT, fontSize: 16, bold: true, color: NAVY, margin: 0, valign: "middle",
  });
  s.addText("一個會反問的機器人，把抽象的概念一路問到能寫成規格為止。", {
    x: M + cw + 0.75, y: 3.22, w: cw - 0.6, h: 0.95,
    fontFace: FONT, fontSize: 13, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
  });

  // 輸入 → 輸出
  s.addShape("roundRect", { x: M, y: 4.65, w: 4.0, h: 1.5, rectRadius: 0.06, fill: { color: SOFT_BG } });
  s.addText("輸入", {
    x: M + 0.3, y: 4.82, w: 3.4, h: 0.36,
    fontFace: FONT, fontSize: 13, bold: true, color: ACCENT, margin: 0, valign: "middle",
  });
  s.addText("一個抽象的概念或想法", {
    x: M + 0.3, y: 5.2, w: 3.4, h: 0.75,
    fontFace: FONT, fontSize: 15, bold: true, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  s.addShape("triangle", { x: 4.95, y: 5.27, w: 0.24, h: 0.3, fill: { color: LAV_MID }, rotate: 90 });

  s.addShape("roundRect", { x: 5.45, y: 4.65, w: CW - 4.75, h: 1.5, rectRadius: 0.06, fill: { color: LAV } });
  s.addText("輸出", {
    x: 5.75, y: 4.82, w: CW - 5.35, h: 0.36,
    fontFace: FONT, fontSize: 13, bold: true, color: ACCENT, margin: 0, valign: "middle",
  });
  s.addText("一份 PRD 初稿與技術文件", {
    x: 5.75, y: 5.2, w: CW - 5.35, h: 0.75,
    fontFace: FONT, fontSize: 15, bold: true, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });

  s.addText("這件事跟神秘客系統其實是同一個主題：把「靠人」的流程，變成「靠系統」也做得到。", {
    x: M, y: 6.35, w: CW, h: 0.45,
    fontFace: FONT, fontSize: 13.5, color: MUTED, margin: 0, valign: "middle",
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P27 PRD 機器人：機制
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對內", "怎麼運作：用大量反問，把抽象逼成具體");

  const tw = 6.3;
  const steps = [
    ["不等你講清楚", "一般工具是你把需求寫好，它幫你排版；這個是反過來——你只要有個模糊的想法就可以開始。"],
    ["大量反問", "反問 → 你補完 → 再反問，一路把沒想到的地方問出來。"],
    ["收斂成文件", "問到夠了，直接產出 PRD 初稿與技術文件。"],
  ];
  steps.forEach((p, i) => {
    const y = 2.1 + i * 1.42;
    s.addShape("roundRect", { x: M, y, w: tw, h: 1.26, rectRadius: 0.06, fill: { color: i === 1 ? CREAM : SOFT_BG } });
    numCircle(s, i + 1, { x: M + 0.28, y: y + 0.37, d: 0.52, fill: NAVY });
    s.addText(p[0], {
      x: M + 0.95, y: y + 0.14, w: tw - 1.2, h: 0.42,
      fontFace: FONT, fontSize: 16, bold: true, color: ACCENT, margin: 0, valign: "middle",
    });
    s.addText(p[1], {
      x: M + 0.95, y: y + 0.55, w: tw - 1.25, h: 0.65,
      fontFace: FONT, fontSize: 12.5, color: NAVY, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
    });
  });

  s.addText("為什麼有效", {
    x: M, y: 6.4, w: 1.6, h: 0.4,
    fontFace: FONT, fontSize: 13, bold: true, color: ACCENT, margin: 0, valign: "middle",
  });
  s.addText("使用者常常不是不願意寫，是不知道該寫什麼——回答問題永遠比面對空白頁容易。", {
    x: M + 1.7, y: 6.4, w: CW - 1.7, h: 0.4,
    fontFace: FONT, fontSize: 13, color: NAVY, margin: 0, valign: "middle",
  });

  image(s, "prd-02-dialogue", "反問過程的對話畫面", { x: M + tw + 0.45, y: 2.1, w: CW - tw - 0.45, h: 4.16 });
  stampPage(s);
}

// ---------------------------------------------------------------- P28 PRD 機器人：現況
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對內", "PRD 機器人：現況與測試回饋");

  s.addShape("roundRect", { x: M, y: 2.1, w: 5.1, h: 2.5, rectRadius: 0.06, fill: { color: NAVY } });
  s.addText("約 10 位", {
    x: M + 0.3, y: 2.35, w: 4.5, h: 1.1,
    fontFace: FONT, fontSize: 48, bold: true, color: WHITE, margin: 0, valign: "middle",
  });
  s.addText("其他部門的同事共同測試中", {
    x: M + 0.3, y: 3.45, w: 4.5, h: 0.42,
    fontFace: FONT, fontSize: 15, bold: true, color: LAV, margin: 0,
  });
  s.addText("＊尚未全面對公司內其他人開放", {
    x: M + 0.3, y: 3.9, w: 4.5, h: 0.5,
    fontFace: FONT, fontSize: 12, italic: true, color: LAV_MID, margin: 0, lineSpacingMultiple: 1.2,
  });

  const info = [
    ["狀態", "已完成開發，仍在持續更新版本", MINT],
    ["測試範圍", "公司內部跨部門同事共同試用", SOFT_BG],
    ["我的做法", "收到 bug 或改善建議就持續修", SOFT_BG],
  ];
  info.forEach((it, i) => {
    const y = 2.1 + i * 0.88;
    s.addShape("roundRect", { x: M + 5.5, y, w: CW - 5.5, h: 0.76, rectRadius: 0.06, fill: { color: it[2] } });
    s.addText(it[0], {
      x: M + 5.75, y, w: 1.5, h: 0.76,
      fontFace: FONT, fontSize: 13.5, bold: true, color: NAVY, margin: 0, valign: "middle",
    });
    s.addText(it[1], {
      x: M + 7.25, y, w: CW - 5.5 - 2.0, h: 0.76,
      fontFace: FONT, fontSize: 13, color: NAVY, margin: 0, valign: "middle",
    });
  });
  image(s, "prd-03-output", "產出的 PRD 初稿", { x: M + 5.5, y: 4.78, w: CW - 5.5, h: 1.35 });

  s.addText("這是我第一次做「跨部門一起測」的工具。收回饋比寫功能難——十個人給的建議常常互相衝突，要判斷哪些是真的需求、哪些只是個人習慣。這件事還沒結束，我還在持續修。", {
    x: M, y: 4.78, w: 5.1, h: 1.35,
    shape: "roundRect", rectRadius: 0.08, fill: { color: CREAM },
    fontFace: FONT, fontSize: 12.5, color: NAVY, margin: [10, 14, 10, 14], valign: "middle", lineSpacingMultiple: 1.3,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P29 其他對內支援
{
  const s = newSlide(pres);
  header(s, "四、工作成果｜對內", "其他對內支援與小工具");
  s.addText("除了上面兩套系統，這段期間也處理了一些零星的支援與小工具需求。", {
    x: M, y: 1.95, w: CW, h: 0.42,
    fontFace: FONT, fontSize: 14, color: MUTED, margin: 0,
  });
  const gw = (CW - 0.4 * 2) / 3;
  for (let i = 0; i < 6; i += 1) {
    const x = M + (i % 3) * (gw + 0.4);
    const y = 2.55 + Math.floor(i / 3) * 1.95;
    s.addShape("roundRect", {
      x, y, w: gw, h: 1.7, rectRadius: 0.06,
      fill: { color: PLACEHOLDER }, line: { color: LAV_MID, width: 1, dashType: "dash" },
    });
    s.addText(`【待補：支援項目 ${i + 1}】\n做了什麼／給誰用／目前狀態`, {
      x: x + 0.2, y, w: gw - 0.4, h: 1.7,
      fontFace: FONT, fontSize: 12, color: MUTED, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.3,
    });
  }
  s.addText("＊本頁待填；若無合適內容可整頁刪除。", {
    x: M, y: 6.45, w: CW, h: 0.4,
    fontFace: FONT, fontSize: 12, color: ACCENT, margin: 0,
  });
  stampPage(s);
}

// ---------------------------------------------------------------- P30 工時分佈
{
  const s = newSlide(pres);
  header(s, "四、工作成果", "四個月工時分佈");
  pieWithNotes(pres, s, {
    labels: ["對外課程教材撰寫", "對外課程授課與備課", "神秘客系統開發", "PRD 機器人開發", "其他（會議、行政、學習）"],
    values: [38, 30, 18, 10, 4],
    colors: [NAVY, TEAL, MINT, SLATE, LAV_MID],
    noteTitle: "＊以上為推估值　【待補：實際百分比】",
    notes: [
      "寫教材要把主題徹底弄懂，花的時間比上課本身多。",
      "含備課、講義調整與試教後的回饋整理。",
      "從需求訪談、開發到上線。",
      "從開發到跨部門測試與持續修正。",
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
    { h: "收穫最多", t: "從使用者身上學需求", b: "難的從來不是技術，是搞懂對方到底要什麼。到後來我乾脆把這件事做成工具——PRD 機器人就是在幫別人做需求釐清。", tint: LAV },
    { h: "最有成就感", t: "東西真的被用起來", b: "神秘客系統每個月都有人真的登入在用，PRD 機器人有十個跨部門同事在測。做出能動的東西不難，做出有人用的東西完全不一樣。", tint: MINT },
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
    {
      tag: "短期", sub: "半年", tint: LAV, size: 12,
      items: [
        "主動探索 AI 新工具與新潮流，維持技術敏感度",
        "把兩套工具從「能用」做到「穩定」：權限、資料、錯誤處理與維運",
        "對外課程正式開賣後擔任主力講師，並推動教材定稿",
        "開始把兩個專案的經驗，沉澱成可複製的開發與導入方法",
      ],
    },
    {
      tag: "中期", sub: "1 年", tint: CREAM,
      items: ["獨立負責一條完整產品線——從題目發想、需求訪談，到上線與維運", "把可複製的方法真正用在第三、第四個專案上，驗證它站得住"],
    },
    {
      tag: "長期", sub: "3～5 年", tint: MINT,
      items: ["成為能同時掌握「對外課程產品」與「對內系統落地」的產品負責人", "讓 AI 導入在公司內成為標準做法，而不是每次都從零開始的個案"],
    },
  ];
  const cw = (CW - 0.4 * 2) / 3;
  cols.forEach((c, i) => {
    const x = M + i * (cw + 0.4);
    s.addShape("roundRect", { x, y: 2.15, w: cw, h: 4.3, rectRadius: 0.06, fill: { color: c.tint }, shadow: shadow() });
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
        x: x + 0.32, y: 3.45, w: cw - 0.64, h: 2.85,
        fontFace: FONT, fontSize: c.size || 13, color: NAVY, margin: 0, valign: "top",
        paraSpaceAfter: c.size ? 7 : 10, lineSpacingMultiple: 1.3,
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
  // P1 封面
  `各位主管、各位同事早安，我是創新研發部的產品工程師黃家朗。
今天要跟大家報告我到職這四個多月的工作內容與成果。
報告大概三十分鐘，中間如果有問題可以隨時打斷我，最後也會留 Q&A 時間。`,
  // P2 目錄
  `今天分七個部分。
前面三段講我是誰、以及這四個多月我實際在做什麼；
第四段是重點，會花最多時間講工作成果，分成對外跟對內兩條線；
最後是心得、職涯規劃，還有幾個給公司的建議。`,
  // P3 自我介紹
  `先簡單介紹我自己。
我是清華大學科技管理學院學士班畢業，主修管理跟經濟雙專長。
進程曦之前在台北數位廣告的產品企劃部擔任產品企劃專員。
我不是一畢業就做 AI 的，是工作幾年之後才轉進 AI 跟教育訓練這個領域。
這個轉職經驗對我現在的工作很有幫助，
因為我很知道一個沒有技術背景的上班族，第一次接觸 AI 工具的時候會卡在哪裡——
這剛好就是我現在在做的事：把 AI 工具教給不懂技術的人。
興趣的部分，我喜歡音樂、攝影、旅行，也會主持 TRPG 跟寫遊戲劇本。
最後一行想特別提一下：我在工作以外參與一個粉絲團體的營運，
在那裡我一樣是用 AI 在解決問題——用 AI 輔助數據分析寫市場洞察報告，
也建了客製化的術語庫來做字幕翻譯校對。
我講這個是想說明，用 AI 解決問題對我來說不是工作要求，是習慣。`,
  // P4 工作歷程
  `這是我的工作歷程，四個節點。
2020 年我在攸你資訊當共同創辦人，負責一個叫 UniLife 的 APP，
從功能設計到使用者旅程都是我做的，帶五個人的團隊，
產品拿過 Appworks 加速器的投資，也入選教育部的 U-start。
2024 年五月我回攸你資訊，職稱變成 AI 提示工程師——
這是我第一次專職做 AI，用 Custom GPTs 幫同事做內部工具。
2024 年十月到台北數位廣告當產品企劃專員，
主要做前沿 AI 工具的調研跟內部培訓，測評過十幾款工具。
今年四月十號到職程曦，到今天大概四個多月。
各位可以看到這條線的方向：做產品、做 AI 工具、做 AI 培訓，
到現在對外做課程、對內做系統，兩件事同時做。
我現在的工作不是轉行，是這條線的延續。
下面這句是我當初應徵程曦時寫在履歷上的：
我想成為連結業務需求與工程開發的 AI 實踐者，
為下一個十年的「辦公室 AI 化」做準備。
這也是我這四個多月一直在做的事。`,
  // P5 個人特質
  `這是我給自己的五個特質。
我沒有用比例圖，因為我覺得給個性打百分比沒什麼意義，
所以每一個特質我都直接配一件我真的做過的事。
解決力：看到問題我的反應不是回報給主管就結束，是直接把工具做出來。
自學力：那些系統用到的東西有很多我原本不會，我是靠 AI 工具邊做邊學做完的。
到目前為止我摸過的生成式 AI 工具超過十款，現在還在自學 n8n 跟 Claude Skills。
換位思考：等一下講的 PRD 機器人，核心就是先搞懂對方要什麼。
系統思維跟交付到底，後面的成果會自己說明，這裡先不展開。

（如果被問到「這些是不是第一次做」，可以補：
在攸你資訊我就在訪談每位同事的工作流程、畫流程圖，
然後為他們客製指令集——PRD 機器人其實是把這件事自動化。
在台北數位我導入過 Notion 建自動化工作流，
把跨部門溝通耗時降了 35%，那跟神秘客系統是同一種思路。）`,
  // P6 公司精神
  `公司精神是「愛、新、勤、誠」，
合起來就是那句話：勤勞地用創新的方式，誠懇地愛人。
我不想只是把這四個字唸過去，所以下一頁我整理了這幾個月裡，
我自己覺得有實際對應到這四個字的四件事。`,
  // P7 我的實踐
  `「新」的部分，神秘客系統我是全程用 replit 這類 AI 工具開發的，
這在部門裡算是比較新的做法。
「勤」的部分，我認為真正的勤勞不是每個月都認真重做一次，
而是想辦法讓它只需要做一次——這也是那套系統的核心。
「愛」的部分，我想講 PRD 機器人。
它不是替人把文件寫掉，是用反問幫人把自己的想法講清楚。
我覺得這個差別滿重要的，替人做跟幫人想是兩件事。
最後「誠」，等一下我會報一個效益數字，
我要先說那是我自己的推估，不是實測出來的，我會誠實標註清楚。`,
  // P8 三階段
  `這四個多月我把它分成三個階段。
第一個月是熟悉期，主要是搞懂公司在做什麼、我們部門的產品線長什麼樣子，
同時就開始投入對外課程的教材製作。
第二到第三個月是產出期，教材開始有成品，我也開始上台試教。
第四個月到現在是落地期，重心轉到對內，
神秘客系統從需求訪談一路做到上線，PRD 機器人也在這段時間開發完成、
進入跨部門測試。
這三段有一個共同點：都是在做「把 AI 變成別人真的用得到的東西」。`,
  // P9 工作定位
  `用比重來看，我的工作大概是對外七、對內三。
對外這條線佔大部分，就是我們部門的教育訓練產品線，四門 AI 應用課。
對內雖然只佔三成，但有兩套完整的系統：
神秘客訪問電話系統已經上線在用，PRD 機器人正在跨部門測試，
另外還有一些零星的支援跟小工具。
一句話總結：對外是把 AI 變成可以賣的課，對內是把 AI 變成可以用的系統。`,
  // P10 扉頁
  `接下來進入重點，工作成果。先講對外的四門課。`,
  // P11 四門課總覽
  `對外這條線總共四門課，我在裡面有兩種角色。
第一、二門是我負責寫教材，從課程架構、內容到範例都是我做的。
第三、四門教材是同事寫的，我負責上台講。
這裡我要先把狀態講清楚，避免誤會：
這四門課目前都只完成了內部試教，還沒有對外開課，
教材也還在持續修正，還沒有定稿。
下一頁之後講的內容，都是在這個前提下。
我覺得這個組合對新人來說滿好的——
寫教材讓我必須把一個主題徹底搞懂才寫得出來，
上台試教則讓我馬上知道學員會在哪裡卡住。這兩件事是互相回饋的。
另外補充一下，這不是我第一次做 AI 培訓。
我在台北數位廣告的時候就在策劃 AI 工具選用的內部課程，
做過主流 AGI 工具的對比分析，也教過 Lovable 這類 prompt-driven 的開發工具。
所以四門課的教材我不是從零摸索，是把做過的方法帶過來用。`,
  // P12 教材①
  `第一門是「AI 時代的商業企劃思維與簡報製作工具」。
這門課的重點其實不在工具操作，而在思維——
很多人以為 AI 簡報工具就是把字丟進去讓它排版，
但真正的價值是在前面的企劃思考階段就讓 AI 參與。
所以我在寫教材的時候，是先把一個企劃案從無到有的流程拆開，
再去看每一段可以怎麼用 AI，而不是反過來從工具功能表開始教。

（時間不夠時本頁可略過）`,
  // P13 教材②
  `第二門是 AI 圖片生成課，一樣是我寫教材。
這門課最難處理的地方是，圖片生成的結果不穩定，
同樣一句提示詞，這次跟下次出來的東西可能差很多。
所以我在教材裡花了不少篇幅在講「怎麼描述」跟「怎麼修正」，
而不是只給一堆可以複製貼上的提示詞範本——
因為範本會過期，方法不會。`,
  // P14 擔任講師
  `第三、四門是我上台講的，目前都完成了內部試教。
這兩門課對我來說有一個額外的意義：
vibe-coding 這門課教的方法，就是我後來開發神秘客系統跟 PRD 機器人實際用的方法。
我等於是先教了別人怎麼用 AI 寫程式，然後自己拿這套方法去做了兩個真的有人在用的系統。
這件事讓我在台上講的時候比較有底氣，
因為我講的不是從教材上讀來的，是我真的做過的。`,
  // P15 教材展示
  `這幾張是教材的實際畫面，給大家看一下我們課程的呈現方式。
如果主管想看完整版本，我課後可以再單獨提供。

（時間不夠時本頁可略過）`,
  // P16 對外現況
  `對外這條線目前的狀態，我分成三塊講。
已完成的是：四門課都完成了內部試教，教材主體也都有了。
進行中的是教材修正，這也是我想多說一句的地方。
同一份教材，不同的講師講起來是不一樣的。
有人擅長帶實作、有人擅長講原理，側重點不同，教材就要跟著調整。
與其先把教材定稿再讓講師自己想辦法，
我們的做法是讓教材去配合講師。
所以「還在修」不是進度落後，是這門課本來就要經過的過程。
尚未開始的是對外開課跟招生，這部分還要等公司安排。`,
  // P17 對內扉頁
  `接下來是對內這條線。
這裡有兩套完整的系統，還有一些零星的支援，我一個一個講。`,
  // P18 神秘客業務流程
  `先說明業務背景，因為不是每個部門都熟悉這件事。
程曦是承接客服外包的公司，所以服務品質必須被檢核。
做法是各個專案組之間每個月隨機互相打電話，
用腳本上的常見問題去問對方，
檢驗接聽的同仁有沒有用符合禮儀、符合標準起手台詞的方式，說出正確的答案。
這種電話我們叫做神秘客訪問電話。
打完之後，各組要寫抽查成果跟分數交給主管，
主管再彙整出報表、還有未達滿分的經緯報告，
然後把分數向下發給各專案組、向上呈報給高級主管。
這整套流程是每個月都要跑一次的。`,
  // P19 痛點
  `這套流程原本是這樣跑的。
第一，資料是散的。做同一件事，要開五份甚至更多份的 Excel 表單。
第二，每個月都要重來一次。每到月底就要重開所有報表，
還要重新設定一份 Google 表單發給大家填。
第三，收回來之後要手動算分數，
然後一直複製貼上到要給不同對象的 Excel 裡面——
給專案組一份、給主管一份、給高級主管又一份。
這裡面最花時間的其實不是判斷，是搬資料。`,
  // P20 解法
  `我的解法核心只有一句話：東西都放在同一個後台，用權限決定誰看得到什麼。
這樣做的好處是，資料不需要再搬移了。
過去要複製貼上到三份不同的 Excel，是因為三個對象要看不同的東西；
現在資料只有一份，我只要針對不同對象改變「顯示的內容」就好。
而且這個設定不是每個月都要改，是設定一次就一直適用，
直到當事人轉職或離職才需要動。
開發的部分，我是全程用 replit 這類 AI 工具做的。`,
  // P21 權限設計
  `權限的部分，我依照會涉及神秘客電訪的人的層級，分成三種等級，
不同的登入帳號會開放不同的功能。
這件事聽起來只是技術設定，但它其實是整套系統能成立的關鍵。
因為過去要拆成好幾份 Excel，本來就是為了讓不同的人看到不同的東西；
如果我只是把五份表單合成一份而沒有權限控制，那反而會出事。
所以是先有權限分層，才有辦法把資料集中。`,
  // P22 功能 1-4
  `系統裡總共有八項功能，我分兩頁講。
前四項是主要流程。
第一是設計新問題，把檢核用的題目建進系統。
第二是抽選問題並分配到指定的專案組——這一項直接取代了過去每個月重設 Google 表單的工作。
第三是評分介面，讓評分的人在同一個地方打分數，不用再自己在 Excel 上算。
第四是報表匯出，要給不同對象的報表直接從系統產生，不用再複製貼上。
這四項合起來，就把「每個月都要重來一次」這個痛點解決掉了。`,
  // P23 功能 5-8
  `後四項是配套。
第五是檢討報告，過去未達滿分要另外寫的經緯報告，現在直接在系統裡提交。
第六是申覆。這一項是我做完前面之後才發現需要的——
因為分數會影響到專案組，如果被評分的人覺得判定有問題，
應該要有一個正式的管道處理，而不是私下用訊息吵。
第七是帳號跟身分組的建立，還有移轉機制。
公司會有人轉職、有人離職，如果每次人事異動都要工程師進資料庫改，
那這套系統遲早會沒人維護。有了移轉機制，管理者自己就能處理。
第八是通知系統，分成全服公告跟個人通知兩種。
第七、第八這兩項不是酷炫的功能，但它們決定這套系統能不能長期活下去。`,
  // P24 前後對比
  `這頁是前後對比，我覺得最能說明這套系統的價值。
左邊是舊流程，右邊是現在。
各位可以看到，改變的重點不是把工作做得更快，
而是把「每個月都要重做一次」的工作，變成「設定一次就好」。
尤其最下面那一列，人事異動——
過去有人離職，那些表單就要重新整理一次；
現在管理者自己在後台移轉就結束了。`,
  // P25 成效
  `這套系統目前的狀態是已經上線，實際在使用中。
使用的單位是所有需要打神秘客訪問電話的單位，以及負責管理的管理者。
效益的部分我要誠實說明：
這個「節省近 90%」是我個人的推估，不是實測出來的數字。
我的推估依據是，過去每個月月底要開所有 Excel 報表、
設一次新的 Google 表單、手動算分數、再一直複製貼上給不同對象；
現在這些全部在同一個後台完成，資料不用搬，而且不需要每個月都改。
如果主管覺得需要，我可以後續實際去量測作業時間，把真實數字補上。`,
  // P26 PRD 問題
  `接下來是第二套系統，PRD 機器人。
它要解決的問題是：想法很抽象，但開發需要規格。
從「我想做一個什麼」到一份能真正動工的 PRD，中間那一段是最卡的。
過去這件事怎麼做？靠有經驗的人一次次開會追問，
或是新人自己對著一份空白文件硬寫。
前者很花時間，後者品質完全看人。
我的解法是做一個會反問的機器人。
你丟一個抽象的概念進去，它產出一份 PRD 初稿跟技術文件。
其實這件事跟神秘客系統是同一個主題：
把原本「靠人」的流程，變成「靠系統」也做得到。`,
  // P27 PRD 機制
  `講一下它怎麼運作，關鍵在「反問」這兩個字。
一般的文件工具是你先把需求寫好，它幫你排版整理；
這個是反過來的——你只要有一個模糊的想法就可以開始。
它會大量反問，你回答、它再問，一路把你沒想到的地方問出來，
問到夠了，再收斂成一份 PRD 初稿跟技術文件。
為什麼這樣有效？
因為我發現使用者常常不是不願意寫，是不知道該寫什麼。
面對一份空白文件很痛苦，但回答問題容易多了。
這個設計其實也回到我前面講的換位思考——
我沒有假設使用者應該要會寫規格，而是設計成他不會也沒關係。
補充一件事：這個「反問」的做法不是我憑空想的。
我在攸你資訊當 AI 提示工程師的時候，
工作就是一個一個去訪談同事的工作流程、畫成流程圖，
再為每個人客製他專用的指令集。
那時候是我用人力在做這件事，一次只能服務一個人。
PRD 機器人等於是把當年那套訪談流程自動化，
讓它可以同時服務很多人，而且不需要我在場。`,
  // P28 PRD 現況
  `目前的狀態我也要說清楚：
它已經開發完成，但還在持續更新版本，
現在是公司內部大約十位其他部門的同事一起在測試，
還沒有全面對公司內其他人開放。
我的做法是收到 bug 或改善建議就繼續修。
這是我第一次做「跨部門一起測」的工具，
老實說收回饋比寫功能難很多。
十個人給的建議常常互相衝突，
我要判斷哪些是真的需求、哪些只是個人習慣。
這件事還沒結束，我還在持續修。`,
  // P29 其他支援
  `除了這兩套系統，這段期間也處理了一些零星的支援跟小工具需求。
這些單獨看都不大，但加起來也佔掉一些時間，所以還是列一下。

（時間不夠時本頁可略過）`,
  // P30 工時分佈
  `這是我四個月的工時分佈，先說明這是推估值。
對外課程加起來大概七成，其中教材撰寫佔比最高，
因為寫教材要把一個主題徹底弄懂，花的時間比上課本身多很多。
對內兩套系統加起來不到三成，
神秘客大概兩成、PRD 機器人一成。
這個比例跟前面講的對外七、對內三是一致的。`,
  // P31 學習心得
  `講一下心得，分三個部分。
收穫最多的是「從使用者身上學需求」。
不管是寫教材還是做系統，難的都不是技術，是搞懂對方到底要什麼。
神秘客的申覆功能就是例子——它不在任何一份需求文件上，
是我理解流程之後才發現需要的。
到後來我乾脆把這件事本身做成工具，那就是 PRD 機器人。
最有成就感的是東西真的被用起來了。
神秘客每個月都有人登入在用，PRD 機器人有十個跨部門同事在測。
做出一個能動的東西不難，做出一個有人真的在用的東西完全不一樣。
最挑戰的是判斷什麼不該做。
用 AI 開發真的很快，快到你想到什麼功能都做得出來，
所以反而要克制，要一直問自己這個功能是使用者需要的，還是我想做的。`,
  // P32 最重要的一課
  `如果要我只講一件這幾個月最重要的收穫，是這個：
從「做出功能」，變成「交付一個會被用的系統」。
我剛到職的時候會覺得，功能寫完就是做完了。
但實際做下來才知道，功能做完只是一半。
你還要想：誰在用、有人離職了怎麼辦、被評分的人不服氣怎麼辦、
十個人的意見不一樣要聽誰的。
這些都不是技術問題，可是不解決，系統做得再好也不會有人用。
我覺得這是我從工程師往產品的方向走，很關鍵的一步。`,
  // P33 職涯規劃
  `職涯規劃分短中長期。
短期半年有四件事。
第一，我想持續主動去探索 AI 的新工具跟新做法。
這個領域變得很快，我不想只會我現在會的這些。
第二，把手上這兩套工具從「能用」做到「穩定」——
權限、資料、錯誤處理、維運，這些現在都還不夠硬。
第三，對外課程正式開賣之後，我希望能擔任主力講師，
並且把教材推到定稿。
第四，開始把這兩個專案的經驗沉澱成可複製的方法。
中期一年，我希望能獨立負責一條完整的產品線，
從題目發想、需求訪談，一路到上線跟維運，
並且把那套方法真的用在第三、第四個專案上，驗證它站得住。
長期三到五年，目標是成為能同時掌握對外課程產品跟對內系統落地的產品負責人，
讓 AI 導入在公司裡變成一件標準的事，而不是每次都從零開始的個案。
這個長期目標其實跟我當初應徵時寫的一樣——
我說我想為下一個十年的「辦公室 AI 化」做準備。
四個多月下來我更確定這件事值得做，而且程曦是可以做這件事的地方。`,
  // P34 建議
  `最後是給公司的三個建議，主題是新人入職跟知識傳承。
第一個，新人報告的目標建議由主管跟新人一起確認。
我理解目前的十二週目標表是為了讓新人有依循，
但它是所有新人共用的版本，跟實際派的工作常常會有落差。
我的建議是到職第一週由主管跟新人一起確認個人化的目標，中間再對一次，
這樣新人會更清楚自己被期待什麼。
第二個，建議各部門列一份自己每月重複作業的清單。
神秘客這套流程本來不在任何待辦清單上，老實說是碰巧發現的。
如果有這樣一份清單，我們部門就可以有系統地評估哪些該先做，而不是等人來說。
這個做法我在前公司實際跑過——
我在台北數位廣告導入 Notion 的時候，就是先盤點各部門重複的作業，
再依需求建自動化工作流，最後把跨部門溝通耗時降了大約 35%。
所以這不是我隨口提的想法，是驗證過會有效的做法。
第三個，內部工具做完之後要有交接跟維護機制。
我自己做完這兩套系統之後最擔心的就是這件事——
系統上線只是開始，如果沒有指定的維護窗口跟使用說明，
它遲早會變成只有我看得懂的黑箱。`,
  // P35 結尾
  `以上就是我這四個多月的報告，謝謝各位的聆聽。
有任何問題都歡迎提出。`,
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
    "prd-02-dialogue", "prd-03-output",
    "course-01-slide", "course-02-slide", "course-03-slide", "course-04-slide",
    "course-showcase-01", "course-showcase-02", "course-showcase-03",
  ].forEach(check);
  if (missing.length) {
    console.log(`\n📷 尚缺 ${missing.length} 張圖（目前顯示為灰底佔位方塊）：`);
    missing.forEach((m) => console.log(`   - report/assets/${m}.png`));
    console.log(`   放圖說明見 report/assets/README.md`);
  }
});
