// クロスボーダーキャリアパス発表資料：見た目の候補（A〜D）を同じ内容で生成する
// 使い方: node build_candidates.js [出力フォルダ]
const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');
const path = require('path');

const OUT = process.argv[2] || __dirname;
const F = 'Meiryo UI';
const W = 13.333, H = 7.5, M = 0.6, CW = W - 2 * M;
const INK = '2E3440', SUB = '5B6472', MUTED = '8A93A3', LINE = 'D5DAE3', CARD = 'F3F5F9', NAVY = '1F2A44', WHITE = 'FFFFFF';
// 虹の7色（図形用）と、白地の文字に使える濃い色（コントラスト4.4以上）
const RB = ['E8505B', 'F39C33', 'F2C230', '5DBB63', '3FA7D6', '4A6FC7', '9A6BC4'];
const DEEP = ['C73B4A', 'A85800', '8F6E00', '2F7F3A', '1D72A3', '3A5BB8', '7A4DAE'];

// B型シミュレーション（標準シナリオ）の資金残高（万円）：開設前〜36か月目
const CASH = [-900, -1056, -1214, -1290, -1348, -1389, -1412, -1432, -1435, -1420, -1403, -1369, -1316, -1242, -1152, -1061, -971, -880, -790, -699, -609, -518, -428, -337, -247, -158, -69, 33, 135, 237, 339, 441, 542, 644, 746, 848, 950];

function mix(hex, a) { // 白と混ぜた淡い色（a=色の割合）
  const c = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16));
  return c.map(v => Math.round(255 - (255 - v) * a).toString(16).padStart(2, '0')).join('').toUpperCase();
}
function lum(hex) {
  const f = c => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  const [r, g, b] = [0, 2, 4].map(i => f(parseInt(hex.slice(i, i + 2), 16) / 255));
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}
function contrast(a, b) { const x = lum(a), y = lum(b); return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05); }
const onColor = hex => (contrast(WHITE, hex) >= 3 ? WHITE : INK);

// ---------------- 発表原稿（全候補共通・約5分） ----------------
const NOTES = [
  '本日は「全国にトータルサポートを。」というテーマで、クロスボーダーキャリアパスの候補生として、これから取り組みたいことをお話しします。',
  '2019年に奈良津1組で入社し、放課後等デイサービスの現場を経験しました。2024年に愛媛県今治市の日吉へ、2026年9月からは西条へ異動し、管理者兼エリアマネージャーとして愛媛エリアを担っています。',
  '一つ目のきっかけは保護者の声です。異動先ではさまざまな困難もありましたが、愛媛の地にもキャレオス、夢門塾はあり、利用者の居場所として確立していました。そんな中、高校生になる利用者の保護者から「ゆうゆうを卒業したあとがとても不安」「愛媛にキャレオスの就労があったらとっても安心する」という言葉をいただきました。嬉しい一方で、胸の痛む言葉でした。',
  '二つ目は、サービス支給量の決定基準の制定です。支給日数の上限が手帳や個別サポートの有無で決まるようになり、「お留守番ができるのか」「行き場が見つからない」「放デイ利用中を理由に児童クラブを断られた」といった不安の声が聞かれました。キャレオスには放課後児童クラブの委託事業があり、愛媛でも受け皿になれるのではと考えました。',
  '三つ目は、愛媛でナンバーワンになるためです。愛媛には、児童から就労、介護まで多数展開する法人があります。総量規制や人口減少の中で選ばれ続けるには、同等以上のトータルサポートを、地域に根差して提供する必要があります。',
  'クロスボーダーで関わりたいのは、第三管理部の事業開発課・地域共創係と、マーケティング課・トータルサポート推進係です。地域共創係では新規事業の立ち上げ方と、自治体・地域企業・学校と関係を築く動き方を。推進係では、トータルサポートの全体像と事業部のつなぎ方、放デイから就労へ切れ目なくつなぐ導線づくりを学びたいと考えています。',
  '具体的な第一歩は、愛媛での就労継続支援B型です。西条市で定員20名の事業所を開設した場合の試算です。標準シナリオでは1日18人の利用で、6か月目に単月黒字、27か月目に初期投資を回収し、3年目の営業利益は約1,140万円です。黒字ラインは1日11人前後。放デイ卒業生の受け皿になれるかが集客の土台です。',
  'では、クロスボーダーの期間に何をするか。最初の3か月は、2つの係の業務に同行して学びながら、日吉・西条の保護者アンケートと、西条市・今治市への事前相談で、ニーズと制度のリスクを確かめます。半年までにB型の事業計画書を社内の形式で作り、受託先や特別支援学校との連携を広げます。1年までに役員へ提案し、承認いただければ開設準備へ進みます。',
  'ここからは、まだ妄想の域を出ない仮構想です。まず愛媛で、B型を第一歩に、放デイ・児童クラブ・就労をつなぐトータルサポートの形をつくります。次に四国支社として四国全域へ。インドネシアの特定技能などの人材も含めて考えています。そして関西支社、中部支社と、全国へ。まだキャレオスにいない人材、ポジションだからこそ、確立させることに価値があると考えています。',
  '全国にトータルサポートを届ける。その始まりを、愛媛で実践する人材になります。ご清聴ありがとうございました。',
];
const SLIDE_TITLES = ['表紙', 'これまでの歩み', 'きっかけ① 保護者の声', 'きっかけ② サービス支給量の決定基準の制定', 'きっかけ③ 愛媛でナンバーワンになるために',
  'クロスボーダーで関わりたい部署', 'まずは愛媛にB型を ― 開設の試算', 'クロスボーダー期間の行動計画', '描く未来 ― 愛媛から全国へ（仮構想）', '結び'];

// ---------------- アイコン（白と濃紺の2色を用意） ----------------
const ICON_NAMES = ['FaHandshake', 'FaRoute', 'FaBoxOpen', 'FaSeedling', 'FaLaptop', 'FaMapMarkedAlt', 'FaBuilding', 'FaFlag'];
const ICONS = {};
async function loadIcons() {
  for (const n of ICON_NAMES) {
    ICONS[n] = {};
    for (const [k, col] of [['w', '#FFFFFF'], ['k', '#' + INK]]) {
      const svg = RDS.renderToStaticMarkup(React.createElement(fa[n], { color: col, size: 256 }));
      const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
      ICONS[n][k] = 'image/png;base64,' + buf.toString('base64');
    }
  }
}

// ---------------- 部品 ----------------
function makeKit(pres) {
  const SH = pres.shapes;
  const k = {
    tx(s, text, o) {
      const fs = o.fontSize || 14, lh = o.lh || 1.3;
      const opts = Object.assign({ fontFace: F, color: INK, margin: 0, isTextBox: true, valign: 'top', lineSpacing: Math.round(fs * lh * 10) / 10 }, o);
      delete opts.lh;
      s.addText(text, opts);
    },
    oval(s, x, y, d, fill, o = {}) { s.addShape(SH.OVAL, Object.assign({ x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0.5 } }, o)); },
    ring(s, x, y, d, color, width) { s.addShape(SH.OVAL, { x, y, w: d, h: d, fill: { color: WHITE, transparency: 100 }, line: { color, width } }); },
    rrect(s, x, y, w, h, fill, r = 0.14, o = {}) { s.addShape(SH.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: r, fill: { color: fill }, line: { color: fill, width: 0.5 } }, o)); },
    rect(s, x, y, w, h, fill, o = {}) { s.addShape(SH.RECTANGLE, Object.assign({ x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0.5 } }, o)); },
    hline(s, x, y, w, color, width) { s.addShape(SH.LINE, { x, y, w, h: 0, line: { color, width } }); },
    chev(s, x, y, w, h, color) { s.addShape(SH.CHEVRON, { x, y, w, h, fill: { color }, line: { color, width: 0.5 } }); },
    arc(s, cx, cy, r, color, range, thick) {
      s.addShape(SH.BLOCK_ARC, { x: cx - r, y: cy - r, w: r * 2, h: r * 2, fill: { color }, line: { color, width: 0.5 }, angleRange: range, arcThicknessRatio: thick / r });
    },
    icon(s, name, x, y, d, onHex) { s.addImage({ data: ICONS[name][onHex === WHITE ? 'w' : 'k'], x, y, w: d, h: d }); },
    dots(s, x, y, d = 0.13, gap = 0.08) { RB.forEach((c, i) => k.oval(s, x + i * (d + gap), y, d, c)); },
  };
  return k;
}

// ---------------- テーマ共通の既定動作 ----------------
const baseTheme = {
  c(i) { return RB[i]; },            // 図形の色
  d(i) { return DEEP[i]; },          // 文字の強調色
  mark(i) { return RB[i]; },         // 小さな目印の色
  acc() { return DEEP[5]; },         // 結論文の色
  chevColor() { return 'C9CFDA'; },
  barHL() { return RB[5]; },
  lineColor() { return RB[5]; },
  pageNum(s, k, n) { k.tx(s, String(n), { x: W - M - 0.6, y: 6.98, w: 0.6, h: 0.3, fontSize: 11, color: SUB, align: 'right' }); },
  card(s, k, x, y, w, h) { k.rrect(s, x, y, w, h, CARD, 0.14); },
  node(s, k, x, cy, i, ctx) { k.oval(s, x, cy - 0.18, 0.36, this.mark(i, ctx), { line: { color: WHITE, width: 3 } }); },
  badge(s, k, x, y, d, text, i, ctx) {
    const c = this.c(i, ctx); k.oval(s, x, y, d, c);
    k.tx(s, text, { x, y, w: d, h: d, fontSize: Math.round(d * 30), bold: true, color: onColor(c), align: 'center', valign: 'middle', lh: 1.0 });
  },
  iconBadge(s, k, x, y, d, name, i, ctx) {
    const c = this.c(i, ctx); k.oval(s, x, y, d, c);
    const p = d * 0.24; k.icon(s, name, x + p, y + p, d - 2 * p, onColor(c));
  },
  label(s, k, x, y, w, h, text, i, ctx) { // 小さなラベル（知識・動き方・STEP）
    const c = this.d(i, ctx); k.rrect(s, x, y, w, h, c, h / 2);
    k.tx(s, text, { x, y, w, h, fontSize: 13, bold: true, color: WHITE, align: 'center', valign: 'middle', lh: 1.0 });
  },
  pill(s, k, x, y, w, h, text, i, ctx) {
    k.rrect(s, x, y, w, h, WHITE, h / 2, { line: { color: this.c(i, ctx), width: 1.75 } });
    k.tx(s, text, { x, y, w, h, fontSize: 15, align: 'center', valign: 'middle', lh: 1.0 });
  },
  stat(s, k, x, y, w, h, value, label, i, ctx) {
    this.card(s, k, x, y, w, h, i, ctx);
    k.tx(s, value, { x: x + 0.1, y: y + h * 0.16, w: w - 0.2, h: 0.62, fontSize: value.length >= 7 ? 24 : 28, bold: true, color: this.d(i, ctx), align: 'center', valign: 'middle', lh: 1.0 });
    k.tx(s, label, { x: x + 0.15, y: y + h * 0.16 + 0.72, w: w - 0.3, h: h * 0.84 - 0.8, fontSize: 12, color: SUB, align: 'center', lh: 1.3 });
  },
  take(s, k, text, y, ctx, o = {}) {
    const x = o.x ?? M, w = o.w ?? CW, h = o.h ?? 0.6, fs = o.fontSize ?? 20;
    if (o.boxed) {
      k.rrect(s, x, y, w, h, mix(RB[5], 0.1), 0.14);
      k.tx(s, text, { x: x + 0.35, y, w: w - 0.7, h, fontSize: fs, bold: true, color: this.acc(ctx), valign: 'middle' });
    } else {
      k.tx(s, text, { x, y, w, h, fontSize: fs, bold: true, color: this.acc(ctx), valign: 'middle', lh: o.lh });
    }
  },
};

// ================= 候補A：レインボーアーチ =================
const A = Object.assign(Object.create(baseTheme), {
  id: 'A', file: '候補A_レインボーアーチ', name: '候補A　レインボーアーチ',
  desc: '白地に虹のアーチ。やさしく明るい印象（前回デザインの改良版）',
  base(s, k, ctx, title, n) {
    s.background = { color: WHITE };
    k.dots(s, M, 0.5);
    k.tx(s, title, { x: M, y: 0.74, w: 11.4, h: 0.7, fontSize: 30, bold: true });
    this.pageNum(s, k, n);
  },
  titleSlide(pres, k) {
    const s = pres.addSlide(); s.background = { color: WHITE };
    RB.forEach((c, i) => k.arc(s, W - 1.2, H + 0.7, 5.0 - i * 0.46, c, [180, 270], 0.3));
    k.dots(s, 0.9, 1.72);
    k.tx(s, 'クロスボーダーキャリアパス　候補生発表', { x: 0.9, y: 2.0, w: 7.5, h: 0.4, fontSize: 16, color: SUB });
    k.tx(s, '全国に\nトータルサポートを。', { x: 0.9, y: 2.5, w: 7.8, h: 1.8, fontSize: 48, bold: true, lh: 1.25 });
    k.tx(s, 'その始まりを、愛媛から。', { x: 0.9, y: 4.45, w: 7, h: 0.5, fontSize: 22, bold: true, color: DEEP[5] });
    k.tx(s, '放課後等デイサービス　管理者 兼 エリアマネージャー\n夢門塾ゆうゆう西条', { x: 0.9, y: 5.3, w: 7, h: 0.8, fontSize: 14, color: SUB, lh: 1.5 });
    return s;
  },
  closing(pres, k) {
    const s = pres.addSlide(); s.background = { color: WHITE };
    RB.forEach((c, i) => k.arc(s, -0.2, -0.2, 4.4 - i * 0.42, c, [0, 90], 0.28));
    k.tx(s, 'その始まりを、愛媛で実践する人材に。', { x: 3.5, y: 2.5, w: 9.23, h: 0.5, fontSize: 20, color: SUB, align: 'right' });
    k.tx(s, '全国に\nトータルサポートを。', { x: 3.5, y: 3.15, w: 9.23, h: 1.8, fontSize: 48, bold: true, align: 'right', lh: 1.25 });
    k.dots(s, W - M - (7 * 0.13 + 6 * 0.08), 5.25);
    k.tx(s, 'ご清聴ありがとうございました', { x: 3.5, y: 5.6, w: 9.23, h: 0.4, fontSize: 15, color: SUB, align: 'right' });
    return s;
  },
});

// ================= 候補B：レインボー・ジャーニー =================
const B = Object.assign(Object.create(baseTheme), {
  id: 'B', file: '候補B_レインボージャーニー', name: '候補B　レインボー・ジャーニー',
  desc: 'スライドが進むごとに虹の色が進み、最後に虹が完成する構成',
  ki(i, ctx) { return ctx && ctx.key !== 'all' ? ctx.key : i; },
  c(i, ctx) { return RB[this.ki(i, ctx)]; },
  d(i, ctx) { return DEEP[this.ki(i, ctx)]; },
  mark(i, ctx) { return RB[this.ki(i, ctx)]; },
  acc(ctx) { return ctx.key === 'all' ? DEEP[5] : ctx.key === 2 ? INK : DEEP[ctx.key]; },
  chevColor(ctx) { return mix(RB[ctx.key === 'all' ? 5 : ctx.key], 0.45); },
  barHL(ctx) { return RB[ctx.key]; },
  lineColor(ctx) { return RB[ctx.key]; },
  card(s, k, x, y, w, h, i, ctx) { k.rrect(s, x, y, w, h, mix(RB[this.ki(i, ctx)], 0.12), 0.14); },
  base(s, k, ctx, title, n) {
    s.background = { color: WHITE };
    const d = 0.95, x = M, y = 0.42;
    if (ctx.key === 'all') { // 最後は虹が完成した円
      RB.forEach((c, j) => { const dd = d - j * 0.065; k.oval(s, x + (d - dd) / 2, y + (d - dd) / 2, dd, c); });
      k.oval(s, x + 0.2, y + 0.2, d - 0.4, WHITE);
      k.tx(s, String(ctx.no).padStart(2, '0'), { x, y, w: d, h: d, fontSize: 15, bold: true, color: INK, align: 'center', valign: 'middle', lh: 1.0 });
    } else {
      k.oval(s, x, y, d, RB[ctx.key]);
      k.tx(s, String(ctx.no).padStart(2, '0'), { x, y, w: d, h: d, fontSize: 22, bold: true, color: onColor(RB[ctx.key]), align: 'center', valign: 'middle', lh: 1.0 });
    }
    k.tx(s, title, { x: M + 1.2, y: 0.42, w: 10.9, h: 0.95, fontSize: 30, bold: true, valign: 'middle' });
    const dd = 0.15, gap = 0.09;
    RB.forEach((c, j) => {
      const done = ctx.key === 'all' || j <= ctx.key;
      if (done) k.oval(s, M + j * (dd + gap), 7.05, dd, c); else k.ring(s, M + j * (dd + gap), 7.05, dd, LINE, 1.25);
    });
    this.pageNum(s, k, n);
  },
  take(s, k, text, y, ctx, o = {}) {
    if (o.boxed) return baseTheme.take.call(this, s, k, text, y, ctx, o);
    const x = o.x ?? M, w = o.w ?? CW, h = o.h ?? 0.6, fs = o.fontSize ?? 20;
    const c = ctx.key === 'all' ? RB[5] : RB[ctx.key];
    if (o.bullet !== false) k.oval(s, x, y + (h - 0.2) / 2, 0.2, c);
    const off = o.bullet !== false ? 0.38 : 0;
    k.tx(s, text, { x: x + off, y, w: w - off, h, fontSize: fs, bold: true, color: this.acc(ctx), valign: 'middle', lh: o.lh });
  },
  titleSlide(pres, k) {
    const s = pres.addSlide(); s.background = { color: WHITE };
    let x = 0.9;
    RB.forEach((c, i) => {
      const d = 0.85 + i * 0.28;
      k.oval(s, x, 6.95 - d, d, c, { line: { color: WHITE, width: 2.5 } });
      if (i === 0) k.tx(s, '愛媛', { x, y: 6.95 - d, w: d, h: d, fontSize: 13, bold: true, color: WHITE, align: 'center', valign: 'middle', lh: 1.0 });
      if (i === 6) k.tx(s, '全国', { x, y: 6.95 - d, w: d, h: d, fontSize: 24, bold: true, color: WHITE, align: 'center', valign: 'middle', lh: 1.0 });
      x += d * 0.8;
    });
    k.tx(s, 'クロスボーダーキャリアパス　候補生発表', { x: 0.9, y: 0.95, w: 8, h: 0.4, fontSize: 16, color: SUB });
    k.tx(s, '全国に\nトータルサポートを。', { x: 0.9, y: 1.4, w: 9, h: 1.8, fontSize: 48, bold: true, lh: 1.25 });
    k.tx(s, 'その始まりを、愛媛から。', { x: 0.9, y: 3.3, w: 8, h: 0.5, fontSize: 22, bold: true, color: DEEP[5] });
    k.tx(s, '放課後等デイサービス　管理者 兼 エリアマネージャー\n夢門塾ゆうゆう西条', { x: 0.9, y: 3.95, w: 7, h: 0.8, fontSize: 14, color: SUB, lh: 1.5 });
    return s;
  },
  closing(pres, k) {
    const s = pres.addSlide(); s.background = { color: WHITE };
    const cx = W / 2, cy = 4.25, R = 3.1, d = 0.95;
    RB.forEach((c, j) => {
      const th = (180 - j * 30) * Math.PI / 180;
      k.oval(s, cx + R * Math.cos(th) - d / 2, cy - R * Math.sin(th) - d / 2, d, c);
    });
    k.tx(s, 'その始まりを、愛媛で実践する人材に。', { x: cx - 2.4, y: 3.3, w: 4.8, h: 0.45, fontSize: 16, color: SUB, align: 'center' });
    k.tx(s, '全国にトータルサポートを。', { x: 1.2, y: 4.95, w: W - 2.4, h: 0.8, fontSize: 44, bold: true, align: 'center', valign: 'middle', lh: 1.2 });
    k.tx(s, 'ご清聴ありがとうございました', { x: 1.2, y: 5.95, w: W - 2.4, h: 0.4, fontSize: 15, color: SUB, align: 'center' });
    return s;
  },
});

// ================= 候補C：エグゼクティブ（紺×虹） =================
const C = Object.assign(Object.create(baseTheme), {
  id: 'C', file: '候補C_エグゼクティブ', name: '候補C　エグゼクティブ',
  desc: '表紙と結びは紺。本文は白地に紺の文字で、役員会議向けの落ち着いた印象',
  c() { return NAVY; },
  d() { return NAVY; },
  acc() { return NAVY; },
  chevColor() { return 'B7C0CE'; },
  barHL() { return NAVY; },
  lineColor() { return NAVY; },
  card(s, k, x, y, w, h) { k.rect(s, x, y, w, h, WHITE, { line: { color: 'D6DBE4', width: 1 } }); },
  label(s, k, x, y, w, h, text) {
    k.rect(s, x, y, w, h, NAVY);
    k.tx(s, text, { x, y, w, h, fontSize: 13, bold: true, color: WHITE, align: 'center', valign: 'middle', lh: 1.0 });
  },
  pill(s, k, x, y, w, h, text, i) {
    k.rect(s, x, y, w, h, WHITE, { line: { color: 'B7C0CE', width: 1 } });
    k.oval(s, x + 0.2, y + (h - 0.14) / 2, 0.14, RB[i]);
    k.tx(s, text, { x: x + 0.42, y, w: w - 0.5, h, fontSize: 15, valign: 'middle', lh: 1.0 });
  },
  take(s, k, text, y, ctx, o = {}) {
    const x = o.x ?? M, w = o.w ?? CW, h = o.h ?? 0.62, fs = o.fontSize ?? 18;
    k.rect(s, x, y, w, h, NAVY);
    k.tx(s, text, { x: x + 0.3, y, w: w - 0.6, h, fontSize: fs, bold: true, color: WHITE, valign: 'middle', lh: o.lh });
  },
  base(s, k, ctx, title, n) {
    s.background = { color: WHITE };
    k.rect(s, M, 0.54, 0.13, 0.13, RB[ctx.key === 'all' ? 5 : ctx.key]);
    k.tx(s, ctx.tag, { x: M + 0.25, y: 0.46, w: 6, h: 0.3, fontSize: 11, bold: true, color: NAVY, charSpacing: 3, valign: 'middle', lh: 1.0 });
    k.tx(s, title, { x: M, y: 0.82, w: 12, h: 0.62, fontSize: 28, bold: true, color: NAVY });
    k.tx(s, '全国にトータルサポートを。｜クロスボーダーキャリアパス 候補生発表', { x: M, y: 7.0, w: 9, h: 0.25, fontSize: 9, color: MUTED, lh: 1.0 });
    k.tx(s, String(n), { x: W - M - 0.6, y: 6.98, w: 0.6, h: 0.28, fontSize: 10, bold: true, color: NAVY, align: 'right', lh: 1.0 });
  },
  titleSlide(pres, k) {
    const s = pres.addSlide(); s.background = { color: NAVY };
    RB.forEach((c, i) => { const r = 3.6 - i * 0.48; k.ring(s, 11.4 - r, 3.75 - r, r * 2, c, 7); });
    k.dots(s, 0.9, 1.6);
    k.tx(s, 'クロスボーダーキャリアパス　候補生発表', { x: 0.9, y: 1.9, w: 7, h: 0.4, fontSize: 15, color: 'A9B4C8' });
    k.tx(s, '全国に\nトータルサポートを。', { x: 0.9, y: 2.35, w: 7.2, h: 1.8, fontSize: 46, bold: true, color: WHITE, lh: 1.25 });
    k.tx(s, 'その始まりを、愛媛から。', { x: 0.9, y: 4.3, w: 7, h: 0.5, fontSize: 21, bold: true, color: RB[2] });
    k.tx(s, '放課後等デイサービス　管理者 兼 エリアマネージャー\n夢門塾ゆうゆう西条', { x: 0.9, y: 5.2, w: 7, h: 0.8, fontSize: 13, color: 'A9B4C8', lh: 1.5 });
    return s;
  },
  closing(pres, k) {
    const s = pres.addSlide(); s.background = { color: NAVY };
    RB.forEach((c, i) => { const r = 3.6 - i * 0.48; k.ring(s, 1.6 - r, 3.75 - r, r * 2, c, 7); });
    k.tx(s, 'その始まりを、愛媛で実践する人材に。', { x: 5.8, y: 2.5, w: W - M - 5.8, h: 0.5, fontSize: 19, color: 'A9B4C8', align: 'right' });
    k.tx(s, '全国に\nトータルサポートを。', { x: 5.8, y: 3.1, w: W - M - 5.8, h: 1.8, fontSize: 46, bold: true, color: WHITE, align: 'right', lh: 1.25 });
    k.tx(s, 'ご清聴ありがとうございました', { x: 5.8, y: 5.35, w: W - M - 5.8, h: 0.4, fontSize: 15, color: 'A9B4C8', align: 'right' });
    return s;
  },
});

// ================= 候補D：パステル =================
const D = Object.assign(Object.create(baseTheme), {
  id: 'D', file: '候補D_パステル', name: '候補D　パステル',
  desc: '淡い虹色の丸をあしらった、やわらかく温かい印象',
  barHL() { return RB[0]; },
  lineColor() { return DEEP[4]; },
  chevColor() { return mix(RB[5], 0.4); },
  card(s, k, x, y, w, h, i) { k.rrect(s, x, y, w, h, mix(RB[i], 0.16), 0.3); },
  pill(s, k, x, y, w, h, text, i) {
    k.rrect(s, x, y, w, h, mix(RB[i], 0.32), h / 2);
    k.tx(s, text, { x, y, w, h, fontSize: 15, align: 'center', valign: 'middle', lh: 1.0 });
  },
  take(s, k, text, y, ctx, o = {}) {
    const x = o.x ?? M, w = o.w ?? CW, h = o.h ?? 0.66, fs = o.fontSize ?? 19;
    k.rrect(s, x, y, w, h, mix(RB[5], 0.14), Math.min(h / 2, 0.4));
    k.tx(s, text, { x: x + 0.4, y, w: w - 0.8, h, fontSize: fs, bold: true, color: DEEP[5], valign: 'middle', lh: o.lh });
  },
  base(s, k, ctx, title, n) {
    s.background = { color: WHITE };
    const a = (n + 3) % 7, b = (n + 5) % 7, c = (n + 1) % 7;
    k.oval(s, W - 1.55, -1.25, 2.6, mix(RB[a], 0.26));
    k.oval(s, W - 2.85, 0.35, 1.15, mix(RB[b], 0.32));
    k.oval(s, W - 3.35, 1.25, 0.32, RB[c]);
    k.oval(s, 0.36, 0.5, 0.82, mix(RB[n % 7], 0.45));
    k.tx(s, title, { x: M, y: 0.62, w: 9.3, h: 0.7, fontSize: 30, bold: true });
    k.oval(s, W - M - 0.46, 6.88, 0.46, mix(RB[n % 7], 0.35));
    k.tx(s, String(n), { x: W - M - 0.46, y: 6.88, w: 0.46, h: 0.46, fontSize: 12, bold: true, align: 'center', valign: 'middle', lh: 1.0 });
  },
  titleSlide(pres, k) {
    const s = pres.addSlide(); s.background = { color: WHITE };
    [[8.2, 0.7, 3.3, 4, 0.22], [10.6, 2.9, 2.9, 2, 0.3], [8.7, 4.3, 2.3, 0, 0.22], [11.3, 0.2, 1.6, 3, 0.28], [7.9, 3.75, 1.15, 6, 0.3]]
      .forEach(([x, y, d, ci, a]) => k.oval(s, x, y, d, mix(RB[ci], a)));
    [[10.7, 2.35, 0.3, 1], [8.1, 2.95, 0.22, 4], [12.4, 6.3, 0.28, 6], [7.5, 6.0, 0.2, 3], [11.0, 6.7, 0.18, 5]]
      .forEach(([x, y, d, ci]) => k.oval(s, x, y, d, RB[ci]));
    k.tx(s, 'クロスボーダーキャリアパス　候補生発表', { x: 0.9, y: 1.9, w: 6.8, h: 0.4, fontSize: 16, color: SUB });
    k.tx(s, '全国に\nトータルサポートを。', { x: 0.9, y: 2.35, w: 6.8, h: 1.8, fontSize: 46, bold: true, lh: 1.25 });
    k.tx(s, 'その始まりを、愛媛から。', { x: 0.9, y: 4.3, w: 6.8, h: 0.5, fontSize: 22, bold: true, color: DEEP[5] });
    k.tx(s, '放課後等デイサービス　管理者 兼 エリアマネージャー\n夢門塾ゆうゆう西条', { x: 0.9, y: 5.15, w: 6.8, h: 0.8, fontSize: 14, color: SUB, lh: 1.5 });
    return s;
  },
  closing(pres, k) {
    const s = pres.addSlide(); s.background = { color: WHITE };
    [[-0.8, -0.9, 3.0, 5, 0.22], [1.7, 0.45, 1.0, 1, 0.32], [10.9, 5.0, 3.2, 3, 0.24], [10.1, 6.35, 0.9, 6, 0.3], [11.9, 0.5, 1.3, 2, 0.3]]
      .forEach(([x, y, d, ci, a]) => k.oval(s, x, y, d, mix(RB[ci], a)));
    [[2.9, 1.3, 0.24, 0], [10.3, 1.5, 0.2, 4], [1.2, 6.3, 0.26, 6], [9.6, 5.9, 0.18, 1]].forEach(([x, y, d, ci]) => k.oval(s, x, y, d, RB[ci]));
    k.tx(s, 'その始まりを、愛媛で実践する人材に。', { x: 1.5, y: 2.3, w: W - 3, h: 0.5, fontSize: 20, color: SUB, align: 'center' });
    k.tx(s, '全国に\nトータルサポートを。', { x: 1.5, y: 2.9, w: W - 3, h: 1.8, fontSize: 46, bold: true, align: 'center', lh: 1.25 });
    k.tx(s, 'ご清聴ありがとうございました', { x: 1.5, y: 4.95, w: W - 3, h: 0.4, fontSize: 15, color: SUB, align: 'center' });
    return s;
  },
});

// ================= 候補E：ブルーウェーブ（参考画像のデザイン） =================
// 背景の波・虹色の光はPowerPointの図形では描けないため、SVGで画像を作って背景に使う
const ASSET = {};
const EC = ['0B3A9E', '1F5FD6', '1E88D0', '2E5BE8', '4B3BC9', '6A3FC9', '9A4DC9'];   // 図形用（濃紺→青→紫）
const ED = ['0B3A9E', '1F5FD6', '176AA6', '2E4FD8', '4B3BC9', '6A3FC9', '8E3FB8'];   // 文字用（白地でコントラスト4.5以上）
const EORB = [['4E86FF', '0B2F9E'], ['66AEFF', '1F5FD6'], ['5CCBF5', '1676BA'], ['7A95FF', '2E4FD8'], ['9A86FF', '4B3BC9'], ['B58CFF', '6A3FC9'], ['E28BE6', '8E3FB8']];
const EDEEP = '0B3A9E', ELABEL = '3F6FC6', ETEXT = '2D4C8A';

function streakPath(x0, y0, x1, y1, h0, h1) { // 太さが先細りする帯
  const L = Math.hypot(x1 - x0, y1 - y0), nx = -(y1 - y0) / L, ny = (x1 - x0) / L;
  const p = (x, y, h, sg) => `${(x + nx * h * sg).toFixed(1)},${(y + ny * h * sg).toFixed(1)}`;
  return `M ${p(x0, y0, h0, 1)} L ${p(x1, y1, h1, 1)} L ${p(x1, y1, h1, -1)} L ${p(x0, y0, h0, -1)} Z`;
}
const BLUR = [2, 4, 6, 12, 18, 30, 40].map(v => `<filter id="b${v}" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="${v}"/></filter>`).join('');
const stops = arr => arr.map(([o, c, a]) => `<stop offset="${o}" stop-color="#${c}" stop-opacity="${a}"/>`).join('');
const lg = (id, x1, y1, x2, y2, st) => `<linearGradient id="${id}" gradientUnits="userSpaceOnUse" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">${stops(st)}</linearGradient>`;
const IRIS = [[0, '2A4BFF', 0.95], [0.3, '5B3BF5', 0.95], [0.48, '9B45F0', 0.9], [0.62, 'E070E0', 0.75], [0.76, 'FF9FC8', 0.5], [0.88, 'FFE0A0', 0.3], [1, 'FFFFFF', 0]];

function svgTitle() { // 表紙・結び：上に大きな青い波、対角線に虹色の光、右下に波
  const sx0 = 80, sy0 = 1427, sx1 = 2500, sy1 = 338;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="2400" height="1350" viewBox="0 0 2400 1350"><defs>${BLUR}
  <radialGradient id="hz1" cx="0" cy="0.5" r="0.45"><stop offset="0" stop-color="#9FE6FF" stop-opacity="0.28"/><stop offset="1" stop-color="#9FE6FF" stop-opacity="0"/></radialGradient>
  <radialGradient id="hz2" cx="1" cy="1" r="0.55"><stop offset="0" stop-color="#CFE6FF" stop-opacity="0.55"/><stop offset="1" stop-color="#CFE6FF" stop-opacity="0"/></radialGradient>
  ${lg('deep', 0, 0, 2400, 0, [[0, '0A2FB8', 0.95], [0.45, '1347E0', 0.95], [0.8, '2F86FF', 0.85], [1, '8FD0FF', 0.5]])}
  ${lg('soft', 0, 0, 2400, 0, [[0, '7FC0FF', 0.4], [0.5, '5AA0FF', 0.45], [1, 'A8DDFF', 0.3]])}
  ${lg('deep2', 1150, 0, 2460, 0, [[0, '0A2FB8', 0.9], [0.55, '2F86FF', 0.9], [1, '9FDBFF', 0.7]])}
  ${lg('iris', sx0, sy0, sx1, sy1, IRIS)}
  ${lg('band', 0, 1500, 1500, 800, [[0, '1E3FE0', 0.55], [0.5, '6A5CFF', 0.3], [1, 'FFFFFF', 0]])}
  ${lg('rib', 150, 1480, 1650, 820, [[0, '0A2FB8', 0.95], [0.35, '1F5FE8', 0.9], [0.6, '6C5CFF', 0.6], [0.85, 'B98CFF', 0.25], [1, 'FFFFFF', 0]])}
  </defs>
  <rect width="2400" height="1350" fill="#FFFFFF"/>
  <rect width="2400" height="1350" fill="url(#hz1)"/><rect width="2400" height="1350" fill="url(#hz2)"/>
  <path d="M -60,80 C 700,220 1450,80 2460,-200 L 2460,50 C 1450,320 700,440 -60,360 Z" fill="url(#soft)" filter="url(#b40)"/>
  <path d="M -60,20 C 700,140 1450,20 2460,-230 L 2460,-190 C 1450,70 700,195 -60,70 Z" fill="#8ED8FF" fill-opacity="0.55" filter="url(#b6)"/>
  <path d="M 300,-20 C 700,30 1100,30 1500,-30 L 1500,-8 C 1100,58 700,60 300,8 Z" fill="#1D5CF0" fill-opacity="0.45" filter="url(#b2)"/>
  <path d="M -60,165 C 700,290 1450,165 2460,-110 L 2460,-55 C 1450,225 700,365 -60,250 Z" fill="url(#deep)" filter="url(#b2)"/>
  <path d="M -60,182 C 700,305 1450,178 2460,-98" stroke="#FFFFFF" stroke-opacity="0.55" stroke-width="4" fill="none" filter="url(#b2)"/>
  <path d="${streakPath(-20, 1640, 1560, 900, 130, 60)}" fill="url(#band)" filter="url(#b30)"/>
  <path d="${streakPath(150, 1590, 1700, 890, 70, 30)}" fill="url(#rib)" fill-opacity="0.35" filter="url(#b18)"/>
  <path d="${streakPath(150, 1590, 1700, 890, 38, 14)}" fill="url(#rib)" filter="url(#b4)"/>
  <path d="${streakPath(150, 1566, 1700, 880, 3, 1)}" fill="#FFFFFF" fill-opacity="0.5" filter="url(#b2)"/>
  <path d="${streakPath(sx0, sy0, sx1, sy1, 120, 45)}" fill="url(#iris)" fill-opacity="0.32" filter="url(#b40)"/>
  <path d="${streakPath(sx0, sy0, sx1, sy1, 34, 8)}" fill="url(#iris)" filter="url(#b6)"/>
  <path d="${streakPath(sx0, sy0, sx1, sy1, 5, 1.5)}" fill="#FFFFFF" fill-opacity="0.55" filter="url(#b2)"/>
  <path d="M 1100,1400 C 1600,1180 2000,1120 2460,1150 L 2460,1300 C 2000,1250 1600,1300 1150,1420 Z" fill="#6FB6FF" fill-opacity="0.25" filter="url(#b18)"/>
  <path d="M 1000,1400 C 1500,1240 1950,1200 2460,1260 L 2460,1400 Z" fill="#CFEAFF" fill-opacity="0.7" filter="url(#b6)"/>
  <path d="M 1150,1370 C 1550,1210 1950,1150 2460,1190 L 2460,1240 C 1950,1215 1600,1265 1250,1380 Z" fill="url(#deep2)" filter="url(#b2)"/>
  <path d="M 1300,1360 C 1700,1230 2050,1200 2460,1215 L 2460,1232 C 2050,1222 1700,1255 1320,1375 Z" fill="#7FD0FF" fill-opacity="0.7" filter="url(#b2)"/>
  </svg>`;
}
function svgContent() { // 本文：右上に波、左下に虹色の光（文字の場所は白いまま）
  const sx0 = -60, sy0 = 1318, sx1 = 1750, sy1 = 1262;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="2400" height="1350" viewBox="0 0 2400 1350"><defs>${BLUR}
  <radialGradient id="hz" cx="1" cy="0" r="0.45"><stop offset="0" stop-color="#D6ECFF" stop-opacity="0.6"/><stop offset="1" stop-color="#D6ECFF" stop-opacity="0"/></radialGradient>
  ${lg('deep', 1460, 0, 2460, 0, [[0, '1447E6', 0], [0.35, '1447E6', 0.85], [1, '0A2FB8', 0.95]])}
  ${lg('bl', -60, 0, 1950, 0, [[0, '1447E6', 0.45], [0.55, '9FDBFF', 0.3], [1, 'FFFFFF', 0]])}
  ${lg('iris', sx0, sy0, sx1, sy1, [[0, '2A4BFF', 0.9], [0.35, '7B3FF2', 0.9], [0.6, 'E070E0', 0.7], [0.8, 'FFB3C8', 0.4], [1, 'FFFFFF', 0]])}
  </defs>
  <rect width="2400" height="1350" fill="#FFFFFF"/><rect width="2400" height="1350" fill="url(#hz)"/>
  <path d="M 1350,-80 C 1800,20 2120,110 2460,280 L 2460,110 C 2120,0 1800,-60 1350,-120 Z" fill="#7FC0FF" fill-opacity="0.35" filter="url(#b18)"/>
  <path d="M 1250,-40 C 1750,40 2100,120 2460,320 L 2460,355 C 2100,160 1750,75 1230,-15 Z" fill="#9FDBFF" fill-opacity="0.6" filter="url(#b4)"/>
  <path d="M 1480,-30 C 1880,15 2160,80 2460,200 L 2460,250 C 2160,120 1880,50 1460,-8 Z" fill="url(#deep)" filter="url(#b2)"/>
  <path d="M 1500,-22 C 1890,22 2165,88 2460,208" stroke="#FFFFFF" stroke-opacity="0.5" stroke-width="3" fill="none" filter="url(#b2)"/>
  <path d="M 1700,-20 C 2000,0 2250,40 2460,90 L 2460,105 C 2250,55 2000,18 1690,-8 Z" fill="#8ED8FF" fill-opacity="0.7" filter="url(#b2)"/>
  <path d="M -60,1290 C 500,1250 1200,1270 1950,1370 L -60,1370 Z" fill="url(#bl)" filter="url(#b6)"/>
  <path d="${streakPath(sx0, sy0, sx1, sy1, 30, 10)}" fill="url(#iris)" fill-opacity="0.35" filter="url(#b12)"/>
  <path d="${streakPath(sx0, sy0, sx1, sy1, 7, 2)}" fill="url(#iris)" filter="url(#b2)"/>
  </svg>`;
}
function svgOrb([light, dark]) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="256" height="256"><defs>
  <linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#${light}"/><stop offset="1" stop-color="#${dark}"/></linearGradient>
  <radialGradient id="h" cx="0.3" cy="0.25" r="0.5"><stop offset="0" stop-color="#FFFFFF" stop-opacity="0.4"/><stop offset="1" stop-color="#FFFFFF" stop-opacity="0"/></radialGradient></defs>
  <circle cx="128" cy="128" r="126" fill="url(#g)"/><circle cx="128" cy="128" r="126" fill="url(#h)"/></svg>`;
}
async function makeAssets(dir) {
  const fs = require('fs'); fs.mkdirSync(dir, { recursive: true });
  ASSET.bgTitle = path.join(dir, 'wave_title.jpg'); ASSET.bgContent = path.join(dir, 'wave_content.jpg');
  await sharp(Buffer.from(svgTitle())).jpeg({ quality: 90 }).toFile(ASSET.bgTitle);
  await sharp(Buffer.from(svgContent())).jpeg({ quality: 90 }).toFile(ASSET.bgContent);
  ASSET.orb = [];
  for (let i = 0; i < EORB.length; i++) {
    const f = path.join(dir, `orb_${i}.png`); await sharp(Buffer.from(svgOrb(EORB[i]))).png().toFile(f); ASSET.orb.push(f);
  }
  ASSET.axis = path.join(dir, 'axis.png');
  await sharp(Buffer.from(`<svg xmlns="http://www.w3.org/2000/svg" width="2400" height="12"><defs>${lg('a', 0, 0, 2400, 0, [[0, '1F5FD6', 0.9], [0.55, '6A3FC9', 0.8], [0.85, 'E070E0', 0.55], [1, 'FFB3C8', 0.15]])}</defs><rect width="2400" height="12" rx="6" fill="url(#a)"/></svg>`)).png().toFile(ASSET.axis);
}

const E = Object.assign(Object.create(baseTheme), {
  id: 'E', file: '候補E_ブルーウェーブ', name: '候補E　ブルーウェーブ',
  desc: '参考画像をもとに、白地に流れる青い波と虹色の光。透明感のある上品な印象',
  contentMaster: 'WAVE_CONTENT',
  defineMasters(pres) {
    pres.defineSlideMaster({ title: 'WAVE_TITLE', background: { path: ASSET.bgTitle }, objects: [] });
    pres.defineSlideMaster({ title: 'WAVE_CONTENT', background: { path: ASSET.bgContent }, objects: [] });
  },
  c(i) { return EC[i]; },
  d(i) { return ED[i]; },
  mark(i) { return EC[i]; },
  acc() { return EDEEP; },
  chevColor() { return 'B9CCF2'; },
  barHL() { return '1F5FD6'; },
  lineColor() { return '1F5FD6'; },
  card(s, k, x, y, w, h) {
    k.rrect(s, x, y, w, h, 'F4F8FF', 0.16, { shadow: { type: 'outer', color: '1F4FB0', opacity: 0.14, blur: 10, offset: 3, angle: 90 } });
  },
  axis(s, k, x, y, w) { s.addImage({ path: ASSET.axis, x, y: y - 0.03, w, h: 0.06 }); },
  dot(s, k, x, y, d, i) { s.addImage({ path: ASSET.orb[i], x, y, w: d, h: d }); },
  node(s, k, x, cy, i) { k.oval(s, x - 0.04, cy - 0.22, 0.44, WHITE); s.addImage({ path: ASSET.orb[i], x, y: cy - 0.18, w: 0.36, h: 0.36 }); },
  badge(s, k, x, y, d, text, i) {
    s.addImage({ path: ASSET.orb[i], x, y, w: d, h: d });
    k.tx(s, text, { x, y, w: d, h: d, fontSize: Math.round(d * 30), bold: true, color: WHITE, align: 'center', valign: 'middle', lh: 1.0 });
  },
  iconBadge(s, k, x, y, d, name, i) {
    s.addImage({ path: ASSET.orb[i], x, y, w: d, h: d });
    const p = d * 0.24; k.icon(s, name, x + p, y + p, d - 2 * p, WHITE);
  },
  pill(s, k, x, y, w, h, text, i) {
    k.rrect(s, x, y, w, h, WHITE, h / 2, { line: { color: EC[i], width: 1.5 } });
    k.tx(s, text, { x, y, w, h, fontSize: 15, color: ETEXT, align: 'center', valign: 'middle', lh: 1.0 });
  },
  take(s, k, text, y, ctx, o = {}) {
    const x = o.x ?? M, w = o.w ?? CW, h = o.h ?? 0.6, fs = o.fontSize ?? 20;
    if (o.boxed) {
      k.rrect(s, x, y, w, h, 'EAF0FE', 0.16);
      s.addImage({ path: ASSET.orb[5], x: x + 0.35, y: y + (h - 0.24) / 2, w: 0.24, h: 0.24 });
      k.tx(s, text, { x: x + 0.78, y, w: w - 1.1, h, fontSize: fs, bold: true, color: EDEEP, valign: 'middle' });
      return;
    }
    const mark = o.bullet !== false;
    if (mark) s.addImage({ path: ASSET.orb[5], x, y: y + (h - 0.24) / 2, w: 0.24, h: 0.24 });
    const off = mark ? 0.42 : 0;
    k.tx(s, text, { x: x + off, y, w: w - off, h, fontSize: fs, bold: true, color: EDEEP, valign: 'middle', lh: o.lh });
  },
  base(s, k, ctx, title, n) {
    k.tx(s, ctx.tag, { x: M, y: 0.46, w: 6, h: 0.3, fontSize: 12, bold: true, color: ELABEL, charSpacing: 4, valign: 'middle', lh: 1.0 });
    k.tx(s, title, { x: M, y: 0.8, w: 10.2, h: 0.65, fontSize: 30, bold: true, color: EDEEP });
    k.tx(s, String(n), { x: W - M - 0.6, y: 6.98, w: 0.6, h: 0.3, fontSize: 11, bold: true, color: ELABEL, align: 'right', lh: 1.0 });
  },
  titleSlide(pres, k) {
    const s = pres.addSlide({ masterName: 'WAVE_TITLE' });
    k.tx(s, 'CROSS-BORDER CAREER PATH', { x: 0.73, y: 2.08, w: 7, h: 0.32, fontSize: 13, bold: true, color: ELABEL, charSpacing: 5, valign: 'middle', lh: 1.0 });
    k.tx(s, 'クロスボーダーキャリアパス　候補生発表', { x: 0.73, y: 2.42, w: 7.5, h: 0.36, fontSize: 15, color: ETEXT, valign: 'middle', lh: 1.1 });
    k.tx(s, '全国に\nトータルサポートを。', { x: 0.73, y: 2.85, w: 7.4, h: 1.7, fontSize: 46, bold: true, color: EDEEP, lh: 1.22 });
    k.tx(s, 'その始まりを、\n　愛媛から。', { x: 9.4, y: 4.3, w: W - M - 9.4, h: 1.0, fontSize: 24, color: EDEEP, lh: 1.45 });
    k.tx(s, '放課後等デイサービス\n管理者 兼 エリアマネージャー\n夢門塾ゆうゆう西条', { x: 9.4, y: 5.42, w: W - M - 9.4, h: 0.8, fontSize: 12, color: ETEXT, lh: 1.45 });
    return s;
  },
  closing(pres, k) {
    const s = pres.addSlide({ masterName: 'WAVE_TITLE' });
    k.tx(s, 'CROSS-BORDER CAREER PATH', { x: 0.73, y: 2.08, w: 7, h: 0.32, fontSize: 13, bold: true, color: ELABEL, charSpacing: 5, valign: 'middle', lh: 1.0 });
    k.tx(s, 'その始まりを、愛媛で実践する人材に。', { x: 0.73, y: 2.42, w: 7.5, h: 0.36, fontSize: 16, color: ETEXT, valign: 'middle', lh: 1.1 });
    k.tx(s, '全国に\nトータルサポートを。', { x: 0.73, y: 2.85, w: 7.4, h: 1.7, fontSize: 46, bold: true, color: EDEEP, lh: 1.22 });
    k.tx(s, 'ご清聴\n　ありがとうございました', { x: 9.4, y: 4.3, w: W - M - 9.4, h: 1.0, fontSize: 19, color: EDEEP, lh: 1.5 });
    return s;
  },
});

// ---------------- 本文スライド（全候補共通のレイアウト） ----------------
function buildDeck(T) {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.title = '全国にトータルサポートを。';
  const k = makeKit(pres);
  if (T.defineMasters) T.defineMasters(pres);
  const add = (ctx, title, n) => {
    const s = T.contentMaster ? pres.addSlide({ masterName: T.contentMaster }) : pres.addSlide();
    T.base(s, k, ctx, title, n); return s;
  };

  // 1. 表紙
  T.titleSlide(pres, k).addNotes(NOTES[0]);

  // 2. これまでの歩み
  {
    const ctx = { no: 1, key: 0, tag: 'PROFILE' }; const s = add(ctx, 'これまでの歩み', 2);
    const items = [
      ['入社', '2019.10.28', '夢門塾ゆうゆう\n奈良津1組', '放課後等デイサービスの\n現場からスタート', 0],
      ['異動', '2024.4.1', '愛媛県今治市\n夢門塾ゆうゆう日吉', '困難も経験しながら、愛媛で\n利用者の居場所づくりに向き合う', 3],
      ['異動', '2026.9.1', '愛媛県西条市\n夢門塾ゆうゆう西条', '管理者 兼 エリアマネージャー\nとして愛媛エリアを担う', 5],
    ];
    const cw = 3.8, gap = (CW - 3 * cw) / 2, ly = 3.25;
    if (T.axis) T.axis(s, k, M, ly, CW); else k.hline(s, M, ly, CW, LINE, 2);
    items.forEach(([lab, d, place, desc, ci], i) => {
      const x = M + i * (cw + gap);
      k.tx(s, lab, { x, y: 1.95, w: cw, h: 0.32, fontSize: 13, color: SUB });
      k.tx(s, d, { x, y: 2.27, w: cw, h: 0.55, fontSize: 26, bold: true, color: T.d(ci, ctx), lh: 1.1 });
      T.node(s, k, x + 0.2, ly, ci, ctx);
      T.card(s, k, x, 3.65, cw, 2.5, ci, ctx);
      k.tx(s, place, { x: x + 0.3, y: 3.92, w: cw - 0.6, h: 0.9, fontSize: 18, bold: true });
      k.tx(s, desc, { x: x + 0.3, y: 4.98, w: cw - 0.6, h: 1.2, fontSize: 14, color: SUB, lh: 1.45 });
    });
    s.addNotes(NOTES[1]);
  }

  // 3. きっかけ① 保護者の声
  {
    const ctx = { no: 2, key: 1, tag: 'WHY' }; const s = add(ctx, 'きっかけ ①　保護者の声', 3);
    k.tx(s, '異動先ではさまざまな困難もあった。それでも愛媛の地にもキャレオス・夢門塾はあり、\n利用者の「居場所」として確立していた。', { x: M, y: 1.72, w: CW, h: 0.85, fontSize: 15, color: SUB, lh: 1.45 });
    [['「ゆうゆうを卒業したあとが\nとても不安」', 0], ['「愛媛にキャレオスの就労が\nあったらとっても安心する」', 4]].forEach(([t, ci], i) => {
      const w = (CW - 0.33) / 2, x = M + i * (w + 0.33);
      T.card(s, k, x, 2.75, w, 2.3, ci, ctx);
      k.tx(s, '“', { x: x + 0.3, y: 2.82, w: 1, h: 0.8, fontSize: 60, bold: true, color: T.mark(ci, ctx), lh: 1.0 });
      k.tx(s, t, { x: x + 0.45, y: 3.55, w: w - 0.8, h: 1.25, fontSize: 21, bold: true, lh: 1.4 });
    });
    k.tx(s, '― 高校生になる利用者の保護者より', { x: M, y: 5.15, w: CW, h: 0.35, fontSize: 12, color: SUB, align: 'right' });
    T.take(s, k, '嬉しい、けれど胸の痛いことば。卒業後の「次の居場所」を愛媛に。', 5.72, ctx);
    s.addNotes(NOTES[2]);
  }

  // 4. きっかけ② 支給量の決定基準
  {
    const ctx = { no: 3, key: 2, tag: 'WHY' }; const s = add(ctx, 'きっかけ ②　サービス支給量の決定基準の制定', 4);
    const cols = [
      ['制度の変化', '受給者証の支給日数の上限が、\n手帳や個別サポートの有無で\n決まるように。', 1],
      ['現場の不安', '・支給日数が減り、お留守番が\n　できるか不安\n・次の行き場がすぐに見つからない\n・放デイ利用中を理由に、\n　児童クラブを断られた', 0],
      ['参入の機会', 'キャレオスには放課後\n児童クラブの委託事業がある。\n愛媛でも受け皿になれないか。', 3],
    ];
    const cw = 3.8, gap = (CW - 3 * cw) / 2, y = 1.8, h = 3.55;
    cols.forEach(([hd, body, ci], i) => {
      const x = M + i * (cw + gap);
      T.card(s, k, x, y, cw, h, ci, ctx);
      T.badge(s, k, x + 0.3, y + 0.3, 0.55, String(i + 1), ci, ctx);
      k.tx(s, hd, { x: x + 1.02, y: y + 0.3, w: cw - 1.2, h: 0.55, fontSize: 19, bold: true, valign: 'middle', lh: 1.1 });
      const bodyText = Array.isArray(body)
        ? body.map((t, j) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: j < body.length - 1, paraSpaceAfter: 8 } }))
        : body;
      k.tx(s, bodyText, { x: x + 0.3, y: y + 1.1, w: cw - 0.6, h: h - 1.3, fontSize: 14, lh: 1.5 });
      if (i < 2) k.chev(s, x + cw + (gap - 0.18) / 2, y + h / 2 - 0.18, 0.18, 0.36, T.chevColor(ctx));
    });
    T.take(s, k, '放デイだけでは支えきれない。地域の「放課後の受け皿」にキャレオスが。', 5.72, ctx);
    s.addNotes(NOTES[3]);
  }

  // 5. きっかけ③ 愛媛でナンバーワン
  {
    const ctx = { no: 4, key: 3, tag: 'WHY' }; const s = add(ctx, 'きっかけ ③　愛媛でナンバーワンになるために', 5);
    T.card(s, k, M, 1.8, 5.9, 4.35, 4, ctx);
    k.tx(s, '愛媛の福祉を支える地域大手（例：来島会）', { x: M + 0.35, y: 2.1, w: 5.2, h: 0.45, fontSize: 16, bold: true, valign: 'middle' });
    ['児童発達支援', '就労継続支援', '相談支援', 'グループホーム', '介護'].forEach((v, i) => {
      T.pill(s, k, M + 0.35 + (i % 2) * 2.65, 2.8 + Math.floor(i / 2) * 0.82, 2.45, 0.6, v, i + 2, ctx);
    });
    k.tx(s, 'ライフステージを通して多数展開', { x: M + 0.35, y: 5.38, w: 5.2, h: 0.35, fontSize: 13, color: SUB });
    const rx = 6.9, rw = W - M - rx;
    k.tx(s, '10年後・20年後を見据えると', { x: rx, y: 1.85, w: rw, h: 0.45, fontSize: 19, bold: true, valign: 'middle' });
    ['総量規制のはじまり', '人口減少', '「選ばれ続ける」ための、地域に根差した福祉'].forEach((t, i) => {
      const y = 2.6 + i * 0.78;
      if (T.dot) T.dot(s, k, rx, y + 0.13, 0.26, [0, 3, 5][i], ctx); else k.oval(s, rx, y + 0.13, 0.26, T.mark([0, 1, 3][i], ctx));
      k.tx(s, t, { x: rx + 0.45, y, w: rw - 0.45, h: 0.52, fontSize: 16, valign: 'middle', lh: 1.2 });
    });
    T.take(s, k, '同等、それ以上の\nトータルサポートが必要。', 4.95, ctx, { x: rx, w: rw, h: 1.2, fontSize: 23, lh: 1.35, bullet: false });
    s.addNotes(NOTES[4]);
  }

  // 6. クロスボーダーで関わりたい部署
  {
    const ctx = { no: 5, key: 4, tag: 'LEARN' }; const s = add(ctx, 'クロスボーダーで関わりたい部署', 6);
    k.tx(s, '愛媛などで福祉施設を広げるための「知識」と「動き方」を身につける', { x: M, y: 1.7, w: CW, h: 0.45, fontSize: 16, color: SUB, valign: 'middle' });
    const units = [
      ['第三管理部　事業開発課', '地域共創係', 'FaHandshake', 3, [
        ['知識', '新規事業の立ち上げ方\n（市場調査・事業計画・収支・指定申請）'],
        ['動き方', '自治体・地域企業・学校と関係を築き、\n地域に根差した事業をつくる']]],
      ['第三管理部　マーケティング課', 'トータルサポート推進係', 'FaRoute', 4, [
        ['知識', 'トータルサポートの全体像と、\n事業部どうしのつなぎ方'],
        ['動き方', '利用者・保護者の声から、\n放デイ→就労へ切れ目なくつなぐ導線づくり']]],
    ];
    const cw = (CW - 0.33) / 2, y = 2.3;
    units.forEach(([dept, team, ic, ci, rows], i) => {
      const x = M + i * (cw + 0.33);
      T.card(s, k, x, y, cw, 3.4, ci, ctx);
      T.iconBadge(s, k, x + 0.35, y + 0.32, 0.72, ic, ci, ctx);
      k.tx(s, dept, { x: x + 1.3, y: y + 0.3, w: cw - 1.5, h: 0.32, fontSize: 13, color: SUB });
      k.tx(s, team, { x: x + 1.3, y: y + 0.62, w: cw - 1.5, h: 0.5, fontSize: 22, bold: true, lh: 1.2 });
      rows.forEach(([lab, t], j) => {
        const yy = y + 1.45 + j * 0.95;
        T.label(s, k, x + 0.35, yy, 0.95, 0.38, lab, ci, ctx);
        k.tx(s, t, { x: x + 1.47, y: yy - 0.02, w: cw - 1.72, h: 0.85, fontSize: 14, lh: 1.4 });
      });
    });
    T.take(s, k, '現場（愛媛）の声 × 本部の知識と動き方 ＝ 自分のボーダーを越える', 5.95, ctx);
    s.addNotes(NOTES[5]);
  }

  // 7. B型の試算（まずは愛媛にB型を）
  {
    const ctx = { no: 6, key: 5, tag: 'FIRST STEP' }; const s = add(ctx, 'まずは愛媛にB型を　― 開設の試算', 7);
    T.card(s, k, M, 1.75, 6.95, 4.0, 6, ctx);
    s.addChart(pres.charts.LINE, [{ name: '資金残高（万円）', labels: CASH.map((_, i) => (i % 6 === 0 ? String(i) : '')), values: CASH }], {
      x: M + 0.15, y: 1.85, w: 6.65, h: 3.8,
      chartColors: [T.lineColor(ctx)], lineSize: 2.5, lineDataSymbol: 'none',
      valAxisMinVal: -1600, valAxisMaxVal: 1200, valAxisMajorUnit: 400, valAxisLabelFormatCode: '#,##0',
      valAxisLabelFontFace: F, valAxisLabelFontSize: 10, valAxisLabelColor: SUB, valAxisLineShow: false,
      catAxisLabelFontFace: F, catAxisLabelFontSize: 10, catAxisLabelColor: SUB, catAxisLabelRotate: 0, catAxisLabelPos: 'low',
      catAxisLineColor: '9AA3B2',
      valGridLine: { color: 'E3E7EE', size: 0.75 }, catGridLine: { style: 'none' },
      showLegend: false,
      showTitle: true, title: '資金残高の推移（万円）　横軸：開設からの月数', titleFontFace: F, titleFontSize: 13, titleColor: INK,
    });
    const rx = 7.8, tw = (W - M - rx - 0.2) / 2;
    [['18人', '1日の平均利用者数\n（安定期）', 0], ['6か月目', '単月で黒字化', 1], ['27か月目', '初期投資900万円を回収', 3], ['1,140万円', '3年目の営業利益\n（概算）', 5]]
      .forEach(([v, l, ci], i) => T.stat(s, k, rx + (i % 2) * (tw + 0.2), 1.75 + Math.floor(i / 2) * 2.1, tw, 1.9, v, l, ci, ctx));
    T.take(s, k, '黒字ラインは1日11人前後。工賃が愛媛平均（2.3万円）に届けば、年間約100万円の上乗せ。', 5.98, ctx, { h: 0.52, fontSize: 17 });
    k.tx(s, '※西条市・定員20名・2027年4月開設・初期投資900万円の標準シナリオ。人件費・家賃は仮置き。報酬は令和8年6月の見直しと新規指定の特例（984/1000）を反映。', { x: M, y: 6.6, w: CW, h: 0.28, fontSize: 10, color: MUTED, lh: 1.1 });
    s.addNotes(NOTES[6]);
  }

  // 8. クロスボーダー期間の行動計画
  {
    const ctx = { no: 7, key: 6, tag: 'ACTION PLAN' }; const s = add(ctx, 'クロスボーダー期間の行動計画', 8);
    const phases = [
      ['〜3か月', '学ぶ・聞く', '・2つの係の業務に同行\n・保護者アンケート\n　（日吉・西条）\n・西条市・今治市へ\n　事前相談', 'アンケート結果\n行政ヒアリング記録', 0],
      ['〜6か月', '形にする', '・B型の事業計画書を\n　社内の形式で作成\n・受託先・連携先の開拓\n・特別支援学校と連携', '事業計画書（案）\n連携先リスト', 2],
      ['〜1年', '提案する', '・役員へ事業計画を提案\n・承認後、物件・採用・\n　指定申請へ\n・保護者説明会の開催', '提案資料\n開設スケジュール', 4],
      ['その先', '広げる', '・B型の開設・運営\n・放課後児童クラブへの\n　参入を検討\n・四国支社構想へ', 'B型の開設\n次の展開案', 6],
    ];
    const cw = 2.85, gap = (CW - 4 * cw) / 3, y = 1.75, h = 4.1;
    phases.forEach(([per, verb, body, out, ci], i) => {
      const x = M + i * (cw + gap);
      T.card(s, k, x, y, cw, h, ci, ctx);
      T.badge(s, k, x + 0.28, y + 0.28, 0.62, String(i + 1), ci, ctx);
      k.tx(s, per, { x: x + 1.05, y: y + 0.24, w: cw - 1.15, h: 0.42, fontSize: 20, bold: true, valign: 'middle', lh: 1.1 });
      k.tx(s, verb, { x: x + 1.05, y: y + 0.64, w: cw - 1.15, h: 0.3, fontSize: 13, bold: true, color: T.d(ci, ctx), valign: 'middle', lh: 1.1 });
      k.tx(s, body, { x: x + 0.28, y: y + 1.15, w: cw - 0.4, h: 1.75, fontSize: 12.5, lh: 1.42 });
      T.label(s, k, x + 0.28, y + 3.0, 0.95, 0.32, '成果物', ci, ctx);
      k.tx(s, out, { x: x + 0.28, y: y + 3.4, w: cw - 0.4, h: 0.6, fontSize: 12, color: SUB, lh: 1.35 });
      if (i < 3) k.chev(s, x + cw + (gap - 0.13) / 2, y + h / 2 - 0.15, 0.13, 0.3, T.chevColor(ctx));
    });
    T.take(s, k, 'ニーズを数字にし、制度のリスクを早めに確かめ、半年で事業計画書を形にする。', 6.03, ctx, { h: 0.55, fontSize: 18 });
    s.addNotes(NOTES[7]);
  }

  // 9. 描く未来
  {
    const ctx = { no: 8, key: 'all', tag: 'VISION' }; const s = add(ctx, '描く未来　― 愛媛から全国へ（仮構想）', 9);
    const steps = [
      ['STEP 1', '愛媛で実践', 'B型を第一歩に、\n放デイ・児童クラブ・\n就労をつなぐ', 'FaSeedling', 0],
      ['STEP 2', '四国支社', '愛媛から四国全域へ\n（インドネシアの特定技能\n等の人材も含めて）', 'FaMapMarkedAlt', 2],
      ['STEP 3', '関西・中部支社', 'エリアごとの拠点を\n各地に広げていく', 'FaBuilding', 4],
      ['GOAL', '全国へ', '全国にキャレオスの\nトータルサポートを', 'FaFlag', 6],
    ];
    const cw = 2.85, gap = (CW - 4 * cw) / 3, y = 1.8, h = 3.35;
    steps.forEach(([st, hd, body, ic, ci], i) => {
      const x = M + i * (cw + gap);
      T.card(s, k, x, y, cw, h, ci, ctx);
      T.iconBadge(s, k, x + 0.3, y + 0.3, 0.78, ic, ci, ctx);
      k.tx(s, st, { x: x + 1.25, y: y + 0.52, w: cw - 1.4, h: 0.35, fontSize: 13, bold: true, color: T.d(ci, ctx), valign: 'middle', lh: 1.0 });
      k.tx(s, hd, { x: x + 0.3, y: y + 1.3, w: cw - 0.5, h: 0.5, fontSize: 21, bold: true, lh: 1.15 });
      k.tx(s, body, { x: x + 0.3, y: y + 1.92, w: cw - 0.45, h: 1.3, fontSize: 13, color: SUB, lh: 1.45 });
      if (i < 3) k.chev(s, x + cw + (gap - 0.13) / 2, y + h / 2 - 0.15, 0.13, 0.3, T.chevColor(ctx));
    });
    T.take(s, k, 'まだキャレオスにいない人材・ポジション。そこを確立させることに価値がある。', 5.45, ctx, { boxed: true, h: 0.95, fontSize: 19 });
    s.addNotes(NOTES[8]);
  }

  // 10. 結び
  T.closing(pres, k).addNotes(NOTES[9]);
  return pres;
}

module.exports = { NOTES, SLIDE_TITLES, THEMES: [A, B, C, D, E] };

// 使い方: node build_candidates.js [出力フォルダ] [候補ID（例: E または A,B）]
if (require.main === module) {
  (async () => {
    await loadIcons();
    await makeAssets(path.join(__dirname, 'assets'));
    const ids = (process.argv[3] || 'A,B,C,D,E').split(',');
    for (const T of [A, B, C, D, E].filter(t => ids.includes(t.id))) {
      const f = await buildDeck(T).writeFile({ fileName: path.join(OUT, T.file + '.pptx') });
      console.log(f);
    }
  })();
}
