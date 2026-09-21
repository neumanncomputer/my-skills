/**
 * consulting-deck / deck_lib.js
 * コンサル風戦略資料をレイアウト固定で生成するための pptxgenjs ラッパー。
 * 座標・フォント・色はすべてこのファイルの定数で管理する。
 * スライドを作るときは必ずこのライブラリの関数を使い、座標を直接書かないこと。
 */
const pptxgen = require("pptxgenjs");

// ---------------------------------------------------------------- 定数
// スライド: 16:9 (LAYOUT_WIDE = 13.333 x 7.5 inch)
const G = {
  W: 13.333, H: 7.5,
  M: 0.5,                 // 左右マージン
  CW: 12.333,             // コンテンツ幅
  SECTION: { x: 0.5, y: 0.3, w: 10.8, h: 0.3 },
  STATUS: { x: 11.533, y: 0.25, w: 1.3, h: 0.34 },   // ステータス枠（右上）
  TITLE: { x: 0.5, y: 0.62, w: 12.333, h: 0.85 },
  RULE_TOP_Y: 1.55,
  BODY: { x: 0.5, y: 1.8, w: 12.333, h: 4.85 },   // 本文エリア（ここから出さない）
  TAKEAWAY_H: 0.6,        // 下部「示唆」ボックスの高さ
  // 近接のルール: 関係する要素は近く（NEAR）、別のまとまりは離す（GAP）
  NEAR: 0.1,              // 見出し→本文など、関係する要素の間隔
  GAP: 0.4,               // ブロック（まとまり）同士の間隔。NEARの4倍
  RULE_BOTTOM_Y: 6.85,
  SOURCE: { x: 0.5, y: 6.9, w: 11.3, h: 0.4 },
  PAGE: { x: 12.033, y: 6.9, w: 0.8, h: 0.4 },
};

const C = {
  text: "2D2D2D",      // 本文・タイトル
  sub: "7D7D7D",       // セクション表示・出典・補足
  line: "DEDEDE",      // 罫線・枠
  fill: "F2F2F2",      // 非強調の塗り
  orange: "D04A02",    // 最重要の強調（1スライド1〜2か所）
  tangerine: "EB8C00", // グラフ2系列目
  yellow: "FFB600",    // グラフ3系列目
  red: "E0301E",       // リスク・警告（限定使用）
  tint: "FBEAE0",      // オレンジの淡色（強調エリアの背景）
  white: "FFFFFF",
  muted: "BFBFBF",     // グラフの非注目部分
  grid: "EDEDED",      // グラフの目盛線
};

// 枠線（この3種類以外は使わない）
const B = {
  std: { color: C.line, width: 0.75 },    // 標準：囲み・表・上下罫線
  thin: { color: C.line, width: 0.5 },    // 細線：リスト内の区切り（アジェンダ、ロードマップの行）
  key: { color: C.orange, width: 1.5 },   // 強調：最重要のカード1つだけ
  gap: (w = 1) => ({ color: C.white, width: w }), // 隣接する塗り同士のすき間（白線）
};

const F = {
  face: "Meiryo",
  cover: 32, title: 20, section: 10, body: 16, sub: 14, label: 12,
  small: 11, note: 9, stat: 28,
};

const LINE_SPACING = 1.25;   // 行間（倍率）
const PARA_AFTER = 8;        // 箇条書きの項目同士の間隔(pt)
const SUB_AFTER = 3;         // 親項目→子項目、子項目同士の間隔(pt)。親子は近く、項目同士は離す
const LIMIT = { orange: 2, boldRatio: 0.4 }; // 太字は個数制限なし。本文に占める割合が40%を超えたら使いすぎの警告

// ---------------------------------------------------------------- 文字装飾の記法
// **太字**   → 太字（本文中は1スライド3か所まで）
// ==強調==   → オレンジ太字（1スライド2か所まで）
function parseRich(str, base = {}) {
  const runs = [];
  const re = /(\*\*[^*]+\*\*|==[^=]+==)/g;
  let last = 0, m;
  while ((m = re.exec(str)) !== null) {
    if (m.index > last) runs.push({ text: str.slice(last, m.index), options: { ...base } });
    const tok = m[0];
    if (tok.startsWith("**")) runs.push({ text: tok.slice(2, -2), options: { ...base, bold: true }, _b: 1 });
    else runs.push({ text: tok.slice(2, -2), options: { ...base, bold: true, color: C.orange }, _o: 1 });
    last = m.index + tok.length;
  }
  if (last < str.length) runs.push({ text: str.slice(last), options: { ...base } });
  if (runs.length === 0) runs.push({ text: "", options: { ...base } });
  return runs;
}
const stripMarks = (s) => String(s).replace(/\*\*|==/g, "");

// ---------------------------------------------------------------- はみ出し推定
function textWidthPt(s, size) {
  let w = 0;
  for (const ch of stripMarks(s)) w += /[\u0000-\u00ff]/.test(ch) ? size * 0.55 : size;
  return w;
}
function estHeightIn(paras, wIn) {
  const avail = wIn * 72 - 10;
  let pt = 0;
  for (const p of paras) {
    const lines = String(p.text).split("\n")
      .reduce((a, l) => a + Math.max(1, Math.ceil(textWidthPt(l, p.size) / (avail - (p.indent || 0)))), 0);
    pt += lines * p.size * LINE_SPACING * 1.2 + PARA_AFTER;
  }
  return pt / 72;
}

// ---------------------------------------------------------------- Deck
class Deck {
  constructor(opts = {}) {
    this.opts = opts;
    this.pres = new pptxgen();
    this.pres.layout = "LAYOUT_WIDE";
    this.pres.title = opts.title || "";
    this.pres.author = opts.author || "";
    this.warnings = [];
    this.titles = [];
    this.count = 0;
    this._defineMasters();
  }

  _defineMasters() {
    // ステータス表示: ミドルグレーの枠線＋太字。本文の強調色（オレンジ）とは区別し、本文より一段控えめにする
    const status = this.opts.status
      ? [
          { rect: { ...G.STATUS, fill: { color: C.white }, line: { color: C.sub, width: 0.75 } } },
          { text: { text: this.opts.status, options: { ...G.STATUS, fontFace: F.face, fontSize: F.small, bold: true, color: C.sub, align: "center", valign: "middle", margin: 0, charSpacing: 1 } } },
        ]
      : [];
    const rules = [
      { line: { x: G.M, y: G.RULE_TOP_Y, w: G.CW, h: 0, line: { ...B.std } } },
      { line: { x: G.M, y: G.RULE_BOTTOM_Y, w: G.CW, h: 0, line: { ...B.std } } },
    ];
    const titlePh = { placeholder: { options: { name: "title", type: "title", ...G.TITLE, fontFace: F.face, fontSize: F.title, bold: true, color: C.text, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.15 }, text: "" } };
    const slideNumber = { ...G.PAGE, fontFace: F.face, fontSize: F.note, color: C.sub, align: "right", valign: "top", margin: 0 };

    // 生成用（本文は関数で描画）
    this.pres.defineSlideMaster({ title: "CONTENT", background: { color: C.white }, objects: [...rules, ...status, titlePh], slideNumber });
    // 手動追加用（PowerPointの「新しいスライド」で選ぶと同じ位置に枠が出る）
    this.pres.defineSlideMaster({
      title: "CONTENT_手動追加用", background: { color: C.white },
      objects: [...rules, ...status, titlePh,
        { placeholder: { options: { name: "section", type: "body", ...G.SECTION, fontFace: F.face, fontSize: F.section, color: C.sub, margin: 0 }, text: "セクション名" } },
        { placeholder: { options: { name: "body", type: "body", ...G.BODY, fontFace: F.face, fontSize: F.body, color: C.text, valign: "top" }, text: "本文" } },
        { placeholder: { options: { name: "source", type: "body", ...G.SOURCE, fontFace: F.face, fontSize: F.note, color: C.sub, margin: 0, valign: "middle" }, text: "出典：" } }],
      slideNumber,
    });
    this.pres.defineSlideMaster({ title: "DIVIDER", background: { color: C.white }, objects: [...status], slideNumber });
    this.pres.defineSlideMaster({ title: "COVER", background: { color: C.white }, objects: [] });
  }

  // ---------------- 共通フレーム
  _content({ title, section, source, note }) {
    const s = this.pres.addSlide({ masterName: "CONTENT" });
    this.count++;
    s._lint = { orange: 0, boldChars: 0, bodyChars: 0, id: `p${this.count + 0}「${stripMarks(title || "").slice(0, 20)}」` };
    if (!title) this._warn(s, "メッセージタイトルがありません");
    s.addText(parseRich(title || "", {}), { placeholder: "title", align: "left" });
    this._countMarks(s, title, true);
    if (estHeightIn([{ text: title || "", size: F.title }], G.TITLE.w) > G.TITLE.h + 0.05)
      this._warn(s, "タイトルが2行を超える可能性。短くしてください");
    if (section) s.addText(section, { ...G.SECTION, fontFace: F.face, fontSize: F.section, color: C.sub, margin: 0, isTextBox: true });
    const foot = [source ? `出典：${source}` : null, note ? `注：${note}` : null].filter(Boolean).join("　");
    if (foot) s.addText(foot, { ...G.SOURCE, fontFace: F.face, fontSize: F.note, color: C.sub, margin: 0, valign: "top", isTextBox: true });
    this.titles.push({ page: this.count + 0, section: section || "", title: stripMarks(title || "") });
    return s;
  }

  _warn(s, msg) { this.warnings.push(`${s && s._lint ? s._lint.id : ""} ${msg}`); }

  _countMarks(s, str, isTitle = false) {
    const t = String(str || "");
    s._lint.orange += (t.match(/==[^=]+==/g) || []).length;
    if (!isTitle) {
      s._lint.boldChars += (t.match(/\*\*[^*]+\*\*/g) || []).reduce((a, m) => a + m.length - 4, 0);
      s._lint.bodyChars += stripMarks(t).length;
    }
  }

  _finishLint(s) {
    if (s._lint.orange > LIMIT.orange) this._warn(s, `オレンジ強調が${s._lint.orange}か所（上限${LIMIT.orange}）`);
    const r = s._lint.bodyChars ? s._lint.boldChars / s._lint.bodyChars : 0;
    if (r > LIMIT.boldRatio) this._warn(s, `本文の太字が多すぎます（本文の${Math.round(r * 100)}%）。太字だけ拾い読みして意味が通る最小限に絞る`);
  }

  // 本文エリア（takeawayがあれば下部を確保）
  _area(takeaway) {
    const b = { ...G.BODY };
    if (takeaway) b.h -= G.TAKEAWAY_H + G.GAP;
    return b;
  }

  _takeaway(s, takeaway) {
    if (!takeaway) return;
    const y = G.BODY.y + G.BODY.h - G.TAKEAWAY_H;
    s.addShape(this.pres.shapes.RECTANGLE, { x: G.M, y, w: G.CW, h: G.TAKEAWAY_H, fill: { color: C.tint }, line: { color: C.tint } });
    s.addText(parseRich(takeaway, { fontFace: F.face, fontSize: F.body, color: C.text, bold: true }),
      { x: G.M + 0.2, y, w: G.CW - 0.4, h: G.TAKEAWAY_H, valign: "middle", margin: 0, isTextBox: true });
    this._countMarks(s, takeaway, true);
    if (estHeightIn([{ text: takeaway, size: F.body }], G.CW - 0.4) > G.TAKEAWAY_H + 0.1) this._warn(s, "示唆ボックスが1行に収まりません");
  }

  // 箇条書き（文字列 or {text, sub:[...]}）
  _bullets(s, points, box, opts = {}) {
    const size = opts.size || F.body;
    const arr = [], paras = [];
    (points || []).forEach((p) => {
      const item = typeof p === "string" ? { text: p } : p;
      this._countMarks(s, item.text);
      const hasSub = (item.sub || []).length > 0;
      parseRich(item.text, { fontFace: F.face, fontSize: size, color: C.text }).forEach((r, i, all) => {
        const o = { ...r.options };
        if (i === 0) { o.bullet = opts.bullet === false ? false : { indent: 14 }; o.paraSpaceAfter = hasSub ? SUB_AFTER : PARA_AFTER; o.lineSpacingMultiple = LINE_SPACING; }
        if (i === all.length - 1) o.breakLine = true;
        arr.push({ text: r.text, options: o });
      });
      paras.push({ text: item.text, size, indent: 14 });
      (item.sub || []).forEach((sb, si, subs) => {
        this._countMarks(s, sb);
        parseRich(sb, { fontFace: F.face, fontSize: F.sub, color: C.text }).forEach((r, i, all) => {
          const o = { ...r.options };
          if (i === 0) { o.bullet = { indent: 12 }; o.indentLevel = 1; o.paraSpaceAfter = si === subs.length - 1 ? PARA_AFTER : SUB_AFTER; o.lineSpacingMultiple = LINE_SPACING; }
          if (i === all.length - 1) o.breakLine = true;
          arr.push({ text: r.text, options: o });
        });
        paras.push({ text: sb, size: F.sub, indent: 40 });
      });
    });
    if (arr.length) delete arr[arr.length - 1].options.breakLine;
    if (!arr.length) return;
    s.addText(arr, { ...box, valign: opts.valign || "top", margin: 3, isTextBox: true, fit: "none" });
    if (estHeightIn(paras, box.w) > box.h) this._warn(s, `本文があふれる可能性（${opts.label || "本文"}）。文字を削るかスライドを分割`);
  }

  _header(s, x, y, w, label, style = "gray") {
    const fill = style === "key" ? C.orange : style === "white" ? C.white : C.fill;
    const color = style === "key" ? C.white : C.text;
    s.addShape(this.pres.shapes.RECTANGLE, { x, y, w, h: 0.45, fill: { color: fill }, line: { color: style === "white" ? C.line : fill } });
    s.addText(stripMarks(label), { x, y, w, h: 0.45, fontFace: F.face, fontSize: F.label, bold: true, color, align: "center", valign: "middle", margin: 0, isTextBox: true });
  }

  // ================================================================ レイアウト
  /** 表紙 */
  cover({ title, subtitle, date, author }) {
    const s = this.pres.addSlide({ masterName: "COVER" });
    this.count++;
    s.addText(stripMarks(title), { x: 0.8, y: 2.3, w: 11.7, h: 1.5, fontFace: F.face, fontSize: F.cover, bold: true, color: C.text, valign: "bottom", margin: 0, isTextBox: true });
    if (subtitle) s.addText(stripMarks(subtitle), { x: 0.8, y: 3.95, w: 11.7, h: 0.6, fontFace: F.face, fontSize: 16, color: C.sub, valign: "top", margin: 0, isTextBox: true });
    s.addShape(this.pres.shapes.LINE, { x: 0.8, y: 5.9, w: 11.733, h: 0, line: { ...B.std } });
    s.addText([date, author].filter(Boolean).join("　｜　"), { x: 0.8, y: 6.0, w: 9, h: 0.4, fontFace: F.face, fontSize: F.label, color: C.sub, margin: 0, isTextBox: true });
    if (this.opts.status) {
      s.addShape(this.pres.shapes.RECTANGLE, { ...G.STATUS, fill: { color: C.white }, line: { color: C.sub, width: 0.75 } });
      s.addText(this.opts.status, { ...G.STATUS, fontFace: F.face, fontSize: F.small, bold: true, color: C.sub, align: "center", valign: "middle", margin: 0, charSpacing: 1, isTextBox: true });
    }
    this.titles.push({ page: this.count, section: "表紙", title: stripMarks(title) });
    return s;
  }

  /** エグゼクティブサマリー: lead(最重要の結論) + points[{head, body}]（数は内容次第） */
  executiveSummary({ title, section = "エグゼクティブサマリー", lead, points = [], source, note, takeaway }) {
    const s = this._content({ title, section, source, note });
    const a = this._area(takeaway);
    let y = a.y;
    if (lead) {
      const h = 0.8;
      s.addShape(this.pres.shapes.RECTANGLE, { x: a.x, y, w: a.w, h, fill: { color: C.tint }, line: { color: C.tint } });
      s.addText(parseRich(lead, { fontFace: F.face, fontSize: 18, bold: true, color: C.text }), { x: a.x + 0.2, y, w: a.w - 0.4, h, valign: "middle", margin: 0, isTextBox: true });
      this._countMarks(s, lead, true);
      if (estHeightIn([{ text: lead, size: 18 }], a.w - 0.4) > h + 0.05) this._warn(s, "サマリーの結論（lead）が2行を超えます");
      y += h + G.GAP;
    }
    const n = points.length;
    if (n > 6) this._warn(s, `要点が${n}個。統合するか本編に委ねることを検討`);
    const rowH = n ? (a.y + a.h - y) / n : 0;
    points.forEach((p, i) => {
      const ry = y + i * rowH;
      if (i > 0) s.addShape(this.pres.shapes.LINE, { x: a.x, y: ry, w: a.w, h: 0, line: { ...B.std } });
      s.addText(String(i + 1).padStart(2, "0"), { x: a.x, y: ry, w: 0.5, h: rowH, fontFace: F.face, fontSize: F.body, bold: true, color: C.sub, valign: "middle", margin: 0, isTextBox: true });
      s.addText(parseRich(p.head, { fontFace: F.face, fontSize: F.body, bold: true, color: C.text }), { x: a.x + 0.55, y: ry, w: 2.4, h: rowH, valign: "middle", margin: 0, isTextBox: true });
      this._countMarks(s, p.head, true);
      this._countMarks(s, p.body || "");
      s.addText(parseRich(p.body || "", { fontFace: F.face, fontSize: F.body, color: C.text }), { x: a.x + 3.15, y: ry, w: a.w - 3.15, h: rowH, valign: "middle", margin: 0, lineSpacingMultiple: LINE_SPACING, isTextBox: true });
      if (estHeightIn([{ text: p.body || "", size: F.body }], a.w - 3.15) > rowH + 0.08) this._warn(s, `サマリー要点${i + 1}の本文があふれる可能性`);
    });
    this._takeaway(s, takeaway);
    this._finishLint(s);
    return s;
  }

  /** アジェンダ: items[文字列], current(強調する番号 0始まり / 省略可) */
  agenda({ title = "本日のアジェンダ", items = [], current = null }) {
    const s = this._content({ title, section: "アジェンダ" });
    const a = G.BODY, rowH = Math.min(0.7, a.h / Math.max(items.length, 1));
    items.forEach((it, i) => {
      const y = a.y + 0.1 + i * rowH, on = current === i, dim = current !== null && !on;
      s.addText(String(i + 1), { x: a.x + 0.3, y, w: 0.6, h: rowH, fontFace: F.face, fontSize: 18, bold: true, color: on ? C.orange : dim ? C.line : C.sub, valign: "middle", margin: 0, isTextBox: true });
      s.addText(stripMarks(it), { x: a.x + 1.0, y, w: 10, h: rowH, fontFace: F.face, fontSize: 16, bold: on, color: dim ? C.sub : C.text, valign: "middle", margin: 0, isTextBox: true });
      if (i < items.length - 1) s.addShape(this.pres.shapes.LINE, { x: a.x + 1.0, y: y + rowH, w: 8, h: 0, line: { ...B.thin } });
    });
    return s;
  }

  /** 汎用カラム（前提条件・スコープ等）: columns[{label, points, key}] 2〜4列 */
  columns({ title, section, columns = [], source, note, takeaway }) {
    const s = this._content({ title, section, source, note });
    const a = this._area(takeaway), n = columns.length || 1;
    const w = (a.w - G.GAP * (n - 1)) / n;
    columns.forEach((c, i) => {
      const x = a.x + i * (w + G.GAP);
      this._header(s, x, a.y, w, c.label, c.key ? "key" : "gray");
      this._bullets(s, c.points, { x, y: a.y + 0.45 + G.NEAR, w, h: a.h - 0.45 - G.NEAR }, { label: c.label });
    });
    this._takeaway(s, takeaway);
    this._finishLint(s);
    return s;
  }

  /** 前提条件・スコープ（columnsの定型） */
  scope({ title, section = "前提条件・スコープ", inScope = [], outOfScope = [], assumptions = [], source, note, takeaway }) {
    const cols = [{ label: "対象範囲", points: inScope }, { label: "対象外", points: outOfScope }];
    if (assumptions.length) cols.push({ label: "前提・仮定", points: assumptions });
    return this.columns({ title, section, columns: cols, source, note, takeaway });
  }

  /** 現状→課題→施策（横並び）: columns[{label, points}]、最後の列を強調 */
  flow({ title, section, columns = [], source, note, takeaway }) {
    const s = this._content({ title, section, source, note });
    const a = this._area(takeaway), n = columns.length || 1, arrow = 0.35;
    const w = (a.w - arrow * (n - 1) - 0.1 * 2 * (n - 1)) / n;
    columns.forEach((c, i) => {
      const x = a.x + i * (w + arrow + 0.2);
      const key = c.key !== undefined ? c.key : i === n - 1;
      this._header(s, x, a.y, w, c.label, key ? "key" : "gray");
      s.addShape(this.pres.shapes.RECTANGLE, { x, y: a.y + 0.45, w, h: a.h - 0.45, fill: { color: C.white }, line: { ...B.std } });
      this._bullets(s, c.points, { x: x + 0.1, y: a.y + 0.6, w: w - 0.2, h: a.h - 0.7 }, { label: c.label });
      if (i < n - 1) s.addShape(this.pres.shapes.CHEVRON, { x: x + w + 0.1, y: a.y + a.h / 2 - 0.2, w: arrow, h: 0.4, fill: { color: C.line }, line: { color: C.line } });
    });
    this._takeaway(s, takeaway);
    this._finishLint(s);
    return s;
  }

  /** 左に要点・右に図（グラフ or 描画関数）: visual = {chart:{type, data, options}} | (slide, box, deck) => void */
  textVisual({ title, section, points = [], visual, visualTitle, source, note, takeaway, ratio = 0.4 }) {
    const s = this._content({ title, section, source, note });
    const a = this._area(takeaway);
    const lw = (a.w - G.GAP) * ratio, rx = a.x + lw + G.GAP, rw = a.w - lw - G.GAP;
    this._bullets(s, points, { x: a.x, y: a.y, w: lw, h: a.h }, { label: "左の要点" });
    let vy = a.y, vh = a.h;
    if (visualTitle) {
      s.addText(stripMarks(visualTitle), { x: rx, y: a.y, w: rw, h: 0.35, fontFace: F.face, fontSize: F.label, bold: true, color: C.text, margin: 0, isTextBox: true });
      vy += 0.4; vh -= 0.4;
    }
    const box = { x: rx, y: vy, w: rw, h: vh };
    if (typeof visual === "function") visual(s, box, this);
    else if (visual && visual.chart) this.chart(s, box, visual.chart);
    else s.addShape(this.pres.shapes.RECTANGLE, { ...box, fill: { color: C.fill }, line: { color: C.fill } });
    this._takeaway(s, takeaway);
    this._finishLint(s);
    return s;
  }

  /**
   * ネイティブグラフ（配色・フォント・凡例を固定）
   * type: bar | line | pie | doughnut
   * highlight: 注目させる箇所。指定したものだけオレンジ、それ以外はグレーになる
   *   - 1系列の棒グラフ / 円・ドーナツ: 項目(labels)のindex
   *   - 複数系列の棒・折れ線: 系列(data)のindex
   * 凡例ルール:
   *   - 円: 各スライスの外側に「項目名＋%」を直接表示（グラフのそばに凡例）
 *   - ドーナツ: 各セグメント上に「項目名＋%」を直接表示（グラフの中に凡例）
   *   - 複数系列の棒・折れ線: グラフ上部（すぐそば）に凡例
   *   - 1系列の棒: 項目名は軸ラベル、系列名はvisualTitleで示す
   */
  chart(s, box, { type = "bar", data, options = {}, highlight = null }) {
    const T = { bar: this.pres.charts.BAR, line: this.pres.charts.LINE, pie: this.pres.charts.PIE, doughnut: this.pres.charts.DOUGHNUT }[type];
    const isPie = type === "pie" || type === "doughnut";
    const multi = data.length > 1;
    const GRAYS = ["A6A6A6", "C4C4C4", "DEDEDE", "EDEDED", "B4B4B4"];
    const MUTED = C.muted;
    let colors = [C.orange, C.tangerine, C.yellow, C.sub, C.line];
    let vary = false;
    if (isPie) {
      vary = true;
      const n = data[0].values.length;
      if (highlight !== null) { let g = 0; colors = data[0].values.map((_, i) => (i === highlight ? C.orange : GRAYS[g++ % GRAYS.length])); }
      else colors = [C.orange, C.tangerine, C.yellow, ...GRAYS].slice(0, Math.max(n, 1));
    } else if (!multi && highlight !== null) {
      vary = true;
      colors = data[0].values.map((_, i) => (i === highlight ? C.orange : MUTED));
    } else if (multi && highlight !== null) {
      { let g = 0; colors = data.map((_, i) => (i === highlight ? C.orange : ["BFBFBF", "D9D9D9", "A6A6A6"][g++ % 3])); }
    }
    if (highlight === null && (isPie || multi)) this._warn(s, "グラフの注目箇所（highlight）が未指定。伝えたい点だけオレンジにする");
    const base = {
      ...box, chartColors: colors, fontFace: F.face, varyColors: vary,
      dataLabelFontSize: F.small, dataLabelFontFace: F.face, dataLabelColor: C.text,
    };
    if (isPie) {
      Object.assign(base, {
        showLegend: false, showLabel: true, showPercent: true, showValue: false,
        dataLabelPosition: type === "pie" ? "outEnd" : "bestFit", dataLabelColor: C.text, dataBorder: { pt: 1, color: C.white },
        holeSize: type === "doughnut" ? 55 : undefined,
      });
    } else {
      Object.assign(base, {
        catAxisLabelFontFace: F.face, valAxisLabelFontFace: F.face, catAxisLabelFontSize: F.small, valAxisLabelFontSize: F.small,
        catAxisLabelColor: C.sub, valAxisLabelColor: C.sub, catAxisLineShow: true,
        valGridLine: { color: C.grid, size: 0.5 }, catGridLine: { style: "none" },
        showValue: true, dataLabelPosition: type === "bar" ? "outEnd" : "t",
        showLegend: multi, legendPos: "t", legendFontSize: F.small, legendFontFace: F.face, legendColor: C.text,
        barGapWidthPct: 60, lineSize: 2, lineDataSymbolSize: 6,
      });
    }
    s.addChart(T, data, { ...base, ...options });
  }

  /** 2×2マトリクス: items[{label, x:0-1, y:0-1, highlight}], quadrants[左上,右上,左下,右下], keyQuadrant(0-3) */
  matrix2x2({ title, section, xAxis = {}, yAxis = {}, quadrants = [], keyQuadrant = null, items = [], points = [], source, note, takeaway }) {
    const s = this._content({ title, section, source, note });
    const a = this._area(takeaway);
    const mw = 7.2, mx = a.x + 0.55, my = a.y, mh = a.h - 0.5, qw = mw / 2, qh = mh / 2;
    [[0, 0], [1, 0], [0, 1], [1, 1]].forEach(([cx, cy], i) => {
      const on = keyQuadrant === i;
      s.addShape(this.pres.shapes.RECTANGLE, { x: mx + cx * qw, y: my + cy * qh, w: qw, h: qh, fill: { color: on ? C.tint : C.fill }, line: B.gap(2) });
      if (quadrants[i]) s.addText(stripMarks(quadrants[i]), { x: mx + cx * qw + 0.1, y: my + cy * qh + 0.08, w: qw - 0.2, h: 0.35, fontFace: F.face, fontSize: F.small, bold: true, color: on ? C.orange : C.sub, margin: 0, isTextBox: true });
    });
    // 軸
    s.addText(`${xAxis.low || "低"}　←　${xAxis.label || ""}　→　${xAxis.high || "高"}`, { x: mx, y: my + mh + 0.08, w: mw, h: 0.35, fontFace: F.face, fontSize: F.small, color: C.sub, align: "center", margin: 0, isTextBox: true });
    s.addText(`${yAxis.low || "低"}　←　${yAxis.label || ""}　→　${yAxis.high || "高"}`, { x: mx - 0.55 - mh / 2 + 0.2, y: my + mh / 2 - 0.18, w: mh, h: 0.35, rotate: 270, fontFace: F.face, fontSize: F.small, color: C.sub, align: "center", margin: 0, isTextBox: true });
    // プロット
    items.forEach((it) => {
      const d = 0.22, px = mx + it.x * mw - d / 2, py = my + (1 - it.y) * mh - d / 2;
      s.addShape(this.pres.shapes.OVAL, { x: px, y: py, w: d, h: d, fill: { color: it.highlight ? C.orange : C.sub }, line: B.gap(1) });
      s.addText(stripMarks(it.label), { x: px + d + 0.05, y: py - 0.07, w: 2.2, h: 0.36, fontFace: F.face, fontSize: F.small, bold: !!it.highlight, color: it.highlight ? C.orange : C.text, margin: 0, isTextBox: true });
    });
    const rx = mx + mw + G.GAP + 0.1;
    this._bullets(s, points, { x: rx, y: a.y, w: a.x + a.w - rx, h: a.h }, { label: "右の要点" });
    this._takeaway(s, takeaway);
    this._finishLint(s);
    return s;
  }

  /** 比較表: options[案名], criteria[{name, values[]}], recommend(推奨案index), points(表の下の補足・任意) */
  comparison({ title, section, options = [], criteria = [], recommend = null, points = [], source, note, takeaway }) {
    const s = this._content({ title, section, source, note });
    const a = this._area(takeaway);
    const firstW = 2.6, colW = (a.w - firstW) / Math.max(options.length, 1);
    const hdr = (t, key) => ({ text: stripMarks(t), options: { bold: true, color: key ? C.white : C.text, fill: { color: key ? C.orange : C.fill }, align: "center", valign: "middle" } });
    const rows = [[hdr("評価軸", false), ...options.map((o, i) => hdr(o, i === recommend))]];
    criteria.forEach((c) => {
      rows.push([
        { text: stripMarks(c.name), options: { bold: true, color: C.text, valign: "middle" } },
        ...c.values.map((v, i) => ({ text: parseRich(String(v), {}), options: { color: C.text, align: "center", valign: "middle", fill: { color: i === recommend ? C.tint : C.white } } })),
      ]);
      c.values.forEach((v) => this._countMarks(s, String(v)));
    });
    const tableH = points.length ? a.h * 0.72 : a.h;
    const rowH = Math.min(0.6, tableH / rows.length);
    s.addTable(rows, { x: a.x, y: a.y, w: a.w, colW: [firstW, ...options.map(() => colW)], rowH, fontFace: F.face, fontSize: F.label, border: { type: "solid", pt: 0.75, color: C.line }, margin: 4 });
    if (points.length) this._bullets(s, points, { x: a.x, y: a.y + rows.length * rowH + 0.2, w: a.w, h: a.h - rows.length * rowH - 0.2 }, { label: "表の下の補足" });
    this._takeaway(s, takeaway);
    this._finishLint(s);
    return s;
  }

  /** 期待効果・KPI: kpis[{name, current, target, unit, note, key}], points(任意) */
  kpi({ title, section = "期待効果・KPI", kpis = [], points = [], source, note, takeaway }) {
    const s = this._content({ title, section, source, note });
    const a = this._area(takeaway);
    const n = Math.min(kpis.length, 4) || 1, w = (a.w - G.GAP * (n - 1)) / n, cardH = points.length ? 2.3 : a.h;
    kpis.slice(0, 4).forEach((k, i) => {
      const x = a.x + i * (w + G.GAP), key = k.key !== undefined ? k.key : i === 0;
      s.addShape(this.pres.shapes.RECTANGLE, { x, y: a.y, w, h: cardH, fill: { color: C.white }, line: key ? { ...B.key } : { ...B.std } });
      s.addText(stripMarks(k.name), { x: x + 0.2, y: a.y + 0.2, w: w - 0.4, h: 0.35, fontFace: F.face, fontSize: F.label, bold: true, color: C.text, margin: 0, valign: "top", isTextBox: true });
      s.addText([
        { text: String(k.target), options: { fontSize: F.stat, bold: true, color: key ? C.orange : C.text } },
        { text: ` ${k.unit || ""}`, options: { fontSize: F.label, color: C.text } },
      ], { x: x + 0.2, y: a.y + 0.55 + G.NEAR, w: w - 0.4, h: 0.6, fontFace: F.face, margin: 0, valign: "middle", isTextBox: true });
      if (k.current !== undefined && k.current !== null && k.current !== "" && k.current !== "―") s.addText(`現状 ${k.current}${k.unit ? " " + k.unit : ""} → 目標`, { x: x + 0.2, y: a.y + 1.25, w: w - 0.4, h: 0.3, fontFace: F.face, fontSize: F.small, color: C.sub, margin: 0, isTextBox: true });
      if (k.note) s.addText(stripMarks(k.note), { x: x + 0.2, y: a.y + 1.55 + G.GAP / 2, w: w - 0.4, h: cardH - 1.85, fontFace: F.face, fontSize: F.small, color: C.text, margin: 0, valign: "top", isTextBox: true });
    });
    if (kpis.length > 4) this._warn(s, "KPIは1スライド4つまで。残りはAppendixへ");
    if (points.length) this._bullets(s, points, { x: a.x, y: a.y + cardH + G.GAP, w: a.w, h: a.h - cardH - G.GAP }, { label: "KPIの補足" });
    this._takeaway(s, takeaway);
    this._finishLint(s);
    return s;
  }

  /** ロードマップ: periods[期間ラベル], tracks[{name, tasks[{label, start, end, key}]}] start/endはperiodsのindex（endを含む） */
  roadmap({ title, section = "ロードマップ", periods = [], tracks = [], milestones = [], source, note, takeaway }) {
    const s = this._content({ title, section, source, note });
    const a = this._area(takeaway);
    const labW = 2.4, pw = (a.w - labW) / Math.max(periods.length, 1), headH = 0.45;
    const msH = milestones.length ? 0.45 : 0;
    const rowH = Math.min(1.0, (a.h - headH - msH) / Math.max(tracks.length, 1));
    periods.forEach((p, i) => {
      const x = a.x + labW + i * pw;
      s.addShape(this.pres.shapes.RECTANGLE, { x, y: a.y, w: pw, h: headH, fill: { color: C.fill }, line: B.gap(1) });
      s.addText(p, { x, y: a.y, w: pw, h: headH, fontFace: F.face, fontSize: F.label, bold: true, color: C.text, align: "center", valign: "middle", margin: 0, isTextBox: true });
    });
    milestones.forEach((m) => {
      const cx = a.x + labW + (m.at + 0.5) * pw;
      s.addShape(this.pres.shapes.DIAMOND, { x: cx - 0.1, y: a.y + headH + 0.12, w: 0.2, h: 0.2, fill: { color: C.orange }, line: { color: C.orange } });
      s.addText(stripMarks(m.label), { x: cx + 0.15, y: a.y + headH + 0.05, w: pw * 1.5, h: 0.35, fontFace: F.face, fontSize: F.small, color: C.orange, bold: true, margin: 0, isTextBox: true });
    });
    tracks.forEach((t, r) => {
      const y = a.y + headH + msH + r * rowH;
      s.addText(stripMarks(t.name), { x: a.x, y, w: labW - 0.15, h: rowH, fontFace: F.face, fontSize: F.label, bold: true, color: C.text, valign: "middle", margin: 0, isTextBox: true });
      s.addShape(this.pres.shapes.LINE, { x: a.x, y: y + rowH, w: a.w, h: 0, line: { ...B.thin } });
      (t.tasks || []).forEach((k) => {
        const x = a.x + labW + k.start * pw + 0.05, w = (k.end - k.start + 1) * pw - 0.1, bh = Math.min(0.5, rowH - 0.2);
        s.addShape(this.pres.shapes.RECTANGLE, { x, y: y + (rowH - bh) / 2, w, h: bh, fill: { color: k.key ? C.orange : C.line }, line: { color: k.key ? C.orange : C.line } });
        s.addText(stripMarks(k.label), { x: x + 0.08, y: y + (rowH - bh) / 2, w: w - 0.16, h: bh, fontFace: F.face, fontSize: F.small, bold: !!k.key, color: k.key ? C.white : C.text, valign: "middle", margin: 0, isTextBox: true });
        if (textWidthPt(k.label, F.small) > (w - 0.16) * 72 * 2) this._warn(s, `ロードマップのバー「${stripMarks(k.label)}」の文字が収まらない可能性`);
      });
    });
    this._takeaway(s, takeaway);
    this._finishLint(s);
    return s;
  }

  /** ネクストステップ: actions[{what, who, when}], decisions[意思決定事項] */
  nextSteps({ title, section = "ネクストステップ", actions = [], decisions = [], source, note }) {
    const s = this._content({ title, section, source, note });
    const a = G.BODY, hasDec = decisions.length > 0, tw = hasDec ? 7.9 : a.w;
    const hdr = (t) => ({ text: t, options: { bold: true, color: C.text, fill: { color: C.fill }, valign: "middle" } });
    const rows = [[hdr("アクション"), hdr("担当"), hdr("期限")]];
    actions.forEach((x) => { rows.push([{ text: parseRich(x.what, {}) }, { text: x.who || "" }, { text: x.when || "" }]); this._countMarks(s, x.what); });
    s.addTable(rows, { x: a.x, y: a.y, w: tw, colW: [tw - 3.2, 1.7, 1.5], rowH: 0.5, fontFace: F.face, fontSize: F.label, color: C.text, valign: "middle", border: { type: "solid", pt: 0.75, color: C.line }, margin: 4 });
    if (hasDec) {
      const x = a.x + tw + G.GAP, w = a.w - tw - G.GAP;
      this._header(s, x, a.y, w, "ご判断いただきたい事項", "key");
      s.addShape(this.pres.shapes.RECTANGLE, { x, y: a.y + 0.45, w, h: a.h - 0.45, fill: { color: C.white }, line: { ...B.std } });
      this._bullets(s, decisions, { x: x + 0.1, y: a.y + 0.6, w: w - 0.2, h: a.h - 0.7 }, { label: "意思決定事項" });
    }
    this._finishLint(s);
    return s;
  }

  /** 箇条書きのみ（Appendix等で文章中心のとき） */
  text({ title, section, points = [], source, note, takeaway }) {
    const s = this._content({ title, section, source, note });
    const a = this._area(takeaway);
    this._bullets(s, points, a);
    this._takeaway(s, takeaway);
    this._finishLint(s);
    return s;
  }

  /** 中扉（Appendix等） */
  divider({ label = "Appendix", sub }) {
    const s = this.pres.addSlide({ masterName: "DIVIDER" });
    this.count++;
    s.addText(label, { x: 0.8, y: 2.9, w: 11.7, h: 1.0, fontFace: F.face, fontSize: F.cover, bold: true, color: C.text, valign: "bottom", margin: 0, isTextBox: true });
    if (sub) s.addText(stripMarks(sub), { x: 0.8, y: 4.0, w: 11.7, h: 0.5, fontFace: F.face, fontSize: F.body, color: C.sub, margin: 0, isTextBox: true });
    this.titles.push({ page: this.count, section: label, title: `（中扉）${label}` });
    return s;
  }

  // ---------------- 出力
  /** 縦読み用：メッセージタイトルだけを順に出力 */
  storyline() {
    return this.titles.map((t) => `p${t.page}　[${t.section}]　${t.title}`).join("\n");
  }

  async save(file) {
    // pptxgenjs は1段落内の2つ目以降のランにも <a:pPr> を書き出し、箇条書き記号が消える。
    // 段落内でランの後ろに出る pPr を除去してから保存する。
    const JSZip = require(require.resolve("jszip", { paths: [require("path").dirname(require.resolve("pptxgenjs")), __dirname] }));
    const buf = await this.pres.write({ outputType: "nodebuffer" });
    const zip = await JSZip.loadAsync(buf);
    for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
      const xml = await zip.file(name).async("string");
      zip.file(name, xml.replace(/(<\/a:r>)<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, "$1"));
    }
    require("fs").writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
    console.log("=== 縦読み（メッセージタイトルの流れ） ===\n" + this.storyline());
    if (this.warnings.length) console.log("\n=== 要確認 ===\n" + this.warnings.join("\n"));
    else console.log("\n=== 要確認 === なし");
    return file;
  }
}

module.exports = { Deck, G, C, B, F, parseRich };
