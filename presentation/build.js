const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // 10 x 5.625
pres.title = '全国にトータルサポートを。';

const F = 'Meiryo UI';
const INK = '2E3440', SUB = '6B7280', LIGHT = 'F4F6FA', WHITE = 'FFFFFF';
const RB = ['E8505B', 'F39C33', 'F2C230', '5DBB63', '3FA7D6', '4A6FC7', '9A6BC4'];

// rainbow dot motif (7 small dots)
function dots(slide, x, y, d = 0.11, gap = 0.07) {
  RB.forEach((c, i) => slide.addShape(pres.shapes.OVAL, { x: x + i * (d + gap), y, w: d, h: d, fill: { color: c }, line: { color: c } }));
}
function header(slide, title, n) {
  slide.background = { color: WHITE };
  dots(slide, 0.5, 0.42);
  slide.addText(title, { x: 0.5, y: 0.58, w: 9, h: 0.7, fontFace: F, fontSize: 26, bold: true, color: INK, margin: 0, isTextBox: true });
  slide.addText(String(n), { x: 9.0, y: 5.2, w: 0.5, h: 0.25, fontFace: F, fontSize: 10, color: SUB, align: 'right', margin: 0, isTextBox: true });
}
function t(slide, text, o) { slide.addText(text, Object.assign({ fontFace: F, color: INK, margin: 0, isTextBox: true, valign: 'top' }, o)); }
function card(slide, x, y, w, h, fill = LIGHT) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: fill } });
}
function num(slide, x, y, label, color, d = 0.42, fs = 14) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { color } });
  t(slide, label, { x, y, w: d, h: d, fontSize: fs, bold: true, color: WHITE, align: 'center', valign: 'middle' });
}

// ---------- 1. Title ----------
{
  const s = pres.addSlide(); s.background = { color: WHITE };
  // rainbow arcs, bottom-right
  RB.forEach((c, i) => {
    const r = 3.6 - i * 0.34;
    s.addShape(pres.shapes.BLOCK_ARC, { x: 8.6 - r, y: 5.625 + 0.4 - r, w: r * 2, h: r * 2, fill: { color: c }, line: { color: c, width: 0 }, angleRange: [180, 270], arcThicknessRatio: 0.09 });
  });
  dots(s, 0.7, 1.35);
  t(s, 'クロスボーダーキャリアパス　候補生発表', { x: 0.7, y: 1.6, w: 6, h: 0.4, fontSize: 14, color: SUB });
  t(s, '全国に\nトータルサポートを。', { x: 0.7, y: 2.05, w: 6.2, h: 1.7, fontSize: 40, bold: true, lineSpacingMultiple: 1.05 });
  t(s, 'その始まりを、愛媛から。', { x: 0.7, y: 3.8, w: 6, h: 0.45, fontSize: 18, color: RB[5], bold: true });
  t(s, '放課後等デイサービス　管理者 兼 エリアマネージャー\n夢門塾ゆうゆう西条', { x: 0.7, y: 4.5, w: 6, h: 0.6, fontSize: 12, color: SUB });
  s.addNotes('本日は「全国にトータルサポートを。」というテーマで、クロスボーダーキャリアパスの候補生として、私がこれから何をしていきたいかをお話しします。（約15秒）');
}

// ---------- 2. Career ----------
{
  const s = pres.addSlide(); header(s, 'これまでの歩み', 2);
  const items = [
    ['2019.10.28', '入社', '夢門塾ゆうゆう\n奈良津1組', '放課後等デイサービスの現場からスタート', RB[0]],
    ['2024.4.1', '異動', '愛媛県今治市\n夢門塾ゆうゆう日吉', '縁のない土地で、一から地域との関係づくり', RB[3]],
    ['2026.9.1', '異動', '愛媛県西条市\n夢門塾ゆうゆう西条', '管理者 兼 エリアマネージャーとして愛媛エリアを担う', RB[5]],
  ];
  const y0 = 2.5;
  s.addShape(pres.shapes.LINE, { x: 1.0, y: y0, w: 8.0, h: 0, line: { color: 'D5DAE3', width: 2 } });
  items.forEach(([d, k, place, desc, c], i) => {
    const x = 0.6 + i * 3.0;
    s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y0 - 0.16, w: 0.32, h: 0.32, fill: { color: c }, line: { color: WHITE, width: 3 } });
    t(s, d, { x, y: 1.78, w: 2.7, h: 0.35, fontSize: 18, bold: true, color: c });
    t(s, k, { x, y: 1.5, w: 1.0, h: 0.28, fontSize: 11, color: SUB });
    card(s, x, 2.9, 2.7, 1.85);
    t(s, place, { x: x + 0.2, y: 3.05, w: 2.35, h: 0.7, fontSize: 13.5, bold: true });
    t(s, desc, { x: x + 0.2, y: 3.85, w: 2.3, h: 0.8, fontSize: 11.5, color: SUB });
  });
  s.addNotes('2019年に入社し、奈良津1組で放課後等デイサービスの現場を経験しました。2024年に愛媛県今治市の日吉へ異動、そして今月、西条へ異動し、管理者兼エリアマネージャーとして愛媛エリアを担っています。（約30秒）');
}

// ---------- 3. Trigger 1: voice ----------
{
  const s = pres.addSlide(); header(s, 'きっかけ ①　保護者の声', 3);
  t(s, '異動先で困難もあった。それでも愛媛の地にもキャレオス・夢門塾はあり、\n利用者の「居場所」として確立していた。', { x: 0.5, y: 1.4, w: 9, h: 0.65, fontSize: 13, color: SUB });
  const q = [['「ゆうゆうを卒業したあとが\nとても不安」', RB[0]], ['「愛媛にキャレオスの就労が\nあったらとっても安心する」', RB[4]]];
  q.forEach(([txt, c], i) => {
    const x = 0.5 + i * 4.6;
    card(s, x, 2.3, 4.4, 1.75);
    t(s, '“', { x: x + 0.2, y: 2.25, w: 0.6, h: 0.7, fontSize: 48, bold: true, color: c });
    t(s, txt, { x: x + 0.35, y: 2.9, w: 3.9, h: 0.95, fontSize: 17, bold: true });
  });
  t(s, '— 高校生になる利用者の保護者より', { x: 0.5, y: 4.15, w: 9, h: 0.3, fontSize: 11, color: SUB, align: 'right' });
  t(s, '嬉しい、けれど胸の痛いことば。卒業後の「次の居場所」を愛媛に。', { x: 0.5, y: 4.6, w: 9, h: 0.4, fontSize: 16, bold: true, color: RB[5] });
  s.addNotes('一つ目のきっかけは保護者の声です。異動先ではさまざまな困難もありましたが、愛媛の地にもキャレオス、夢門塾はあり、利用者の居場所として確立していました。そんな中、高校生になる利用者の保護者から「ゆうゆうを卒業したあとがとても不安」「愛媛にキャレオスの就労があったらとっても安心する」という言葉をいただきました。嬉しい一方で、胸の痛む言葉でした。（約40秒）');
}

// ---------- 4. Trigger 2: policy ----------
{
  const s = pres.addSlide(); header(s, 'きっかけ ②　サービス支給量の決定基準の制定', 4);
  const cols = [
    ['制度の変化', '受給者証の支給日数の上限が、手帳や個別サポートの有無で決まるように。', RB[1]],
    ['現場の不安', '・支給日数が減り、お留守番ができるか\n・次の行き場がすぐに見つからない\n・放デイ利用中を理由に児童クラブを断られた', RB[0]],
    ['参入の機会', 'キャレオスには放課後児童クラブの委託事業がある。愛媛でも受け皿になれないか。', RB[3]],
  ];
  cols.forEach(([h, b, c], i) => {
    const x = 0.5 + i * 3.1;
    card(s, x, 1.5, 2.8, 2.75);
    num(s, x + 0.25, 1.8, String(i + 1), c);
    t(s, h, { x: x + 0.8, y: 1.83, w: 1.9, h: 0.4, fontSize: 16, bold: true, valign: 'middle' });
    t(s, b, { x: x + 0.25, y: 2.45, w: 2.35, h: 2.3, fontSize: 12.5, lineSpacingMultiple: 1.2 });
    if (i < 2) s.addShape(pres.shapes.CHEVRON, { x: x + 2.87, y: 2.72, w: 0.16, h: 0.3, fill: { color: 'C9CFDA' }, line: { color: 'C9CFDA' } });
  });
  t(s, '放デイだけでは支えきれない。地域の「放課後の受け皿」にキャレオスが。', { x: 0.5, y: 4.5, w: 9, h: 0.4, fontSize: 15, bold: true, color: RB[5] });
  s.addNotes('二つ目は、サービス支給量の決定基準の制定です。支給日数の上限が手帳や個別サポートの有無で決まるようになり、日数を減らされた方から「お留守番ができるのか」「行き場が見つからない」「放デイ利用中を理由に児童クラブを断られた」という不安の声が聞かれました。キャレオスには放課後児童クラブの委託事業があり、愛媛でも参入の機会があると考えました。（約40秒）');
}

// ---------- 5. Trigger 3: No.1 ----------
{
  const s = pres.addSlide(); header(s, 'きっかけ ③　愛媛でナンバーワンになるために', 5);
  // left: local leader
  card(s, 0.5, 1.5, 4.3, 3.4);
  t(s, '愛媛の福祉を支える地域大手（例：来島会）', { x: 0.75, y: 1.7, w: 3.9, h: 0.35, fontSize: 13, bold: true });
  const svc = ['児童発達支援', '就労継続支援', '相談支援', 'グループホーム', '介護'];
  svc.forEach((v, i) => {
    const x = 0.75 + (i % 2) * 1.95, y = 2.2 + Math.floor(i / 2) * 0.62;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.8, h: 0.48, rectRadius: 0.24, fill: { color: WHITE }, line: { color: RB[i + 2], width: 1.5 } });
    t(s, v, { x, y, w: 1.8, h: 0.48, fontSize: 12, align: 'center', valign: 'middle' });
  });
  t(s, 'ライフステージを通して多数展開', { x: 0.75, y: 4.35, w: 3.9, h: 0.35, fontSize: 11, color: SUB });
  // right: future
  t(s, '10年後・20年後を見据えると', { x: 5.2, y: 1.55, w: 4.3, h: 0.4, fontSize: 15, bold: true });
  const risks = [['総量規制のはじまり', RB[0]], ['人口減少', RB[1]], ['「選ばれ続ける」ための地域に根差した福祉', RB[3]]];
  risks.forEach(([r, c], i) => {
    num(s, 5.2, 2.1 + i * 0.62, '', c, 0.22);
    t(s, r, { x: 5.55, y: 2.05 + i * 0.62, w: 4.0, h: 0.35, fontSize: 13.5, valign: 'middle' });
  });
  t(s, '同等、それ以上の\nトータルサポートが必要。', { x: 5.2, y: 4.0, w: 4.3, h: 0.9, fontSize: 18, bold: true, color: RB[5] });
  s.addNotes('三つ目は、愛媛でナンバーワンになるためです。愛媛には来島会のように、児童から就労、介護まで多数展開する法人があります。総量規制や人口減少の中で選ばれ続けるには、同等以上のトータルサポートを地域に根差して提供する必要があります。（約30秒）');
}

// ---------- 6. Cross-border learning ----------
{
  const s = pres.addSlide(); header(s, 'クロスボーダーで関わりたい部署', 6);
  t(s, '愛媛などで福祉施設を広げるための「知識」と「動き方」を身につける', { x: 0.5, y: 1.3, w: 9, h: 0.35, fontSize: 13, color: SUB });
  const units = [
    ['第三管理部　事業開発課', '地域共創係', RB[3], [
      ['知識', '新規事業の立ち上げ方（市場調査・事業計画・収支の見立て・指定申請）'],
      ['動き方', '自治体・地域企業・学校と関係を築き、地域に根差した事業をつくる'],
    ]],
    ['第三管理部　マーケティング課', 'トータルサポート推進係', RB[4], [
      ['知識', 'キャレオスのトータルサポートの全体像と、事業部どうしのつなぎ方'],
      ['動き方', '利用者・保護者の声から、放デイ → 就労へ切れ目なくつなぐ導線をつくる'],
    ]],
  ];
  units.forEach(([dept, team, c, rows], i) => {
    const x = 0.5 + i * 4.6, y = 1.85;
    card(s, x, y, 4.4, 2.75);
    t(s, dept, { x: x + 0.25, y: y + 0.2, w: 3.9, h: 0.3, fontSize: 11, color: SUB });
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: y + 0.6, w: 0.22, h: 0.22, fill: { color: c }, line: { color: c } });
    t(s, team, { x: x + 0.57, y: y + 0.5, w: 3.6, h: 0.42, fontSize: 18, bold: true, valign: 'middle' });
    rows.forEach(([k, v], j) => {
      const yy = y + 1.1 + j * 0.78;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.25, y: yy, w: 0.75, h: 0.3, rectRadius: 0.15, fill: { color: c }, line: { color: c } });
      t(s, k, { x: x + 0.25, y: yy, w: 0.75, h: 0.3, fontSize: 10, bold: true, color: WHITE, align: 'center', valign: 'middle' });
      t(s, v, { x: x + 1.15, y: yy - 0.03, w: 3.05, h: 0.7, fontSize: 11.5, lineSpacingMultiple: 1.1 });
    });
  });
  t(s, '現場（愛媛）の声 × 本部の知識と動き方 ＝ 自分のボーダーを越える', { x: 0.5, y: 4.78, w: 9, h: 0.35, fontSize: 15, bold: true, color: RB[5] });
  s.addNotes('クロスボーダーで関わりたいのは、第三管理部の事業開発課・地域共創係と、マーケティング課・トータルサポート推進係です。地域共創係では、市場調査や事業計画、行政との協議といった新規事業の立ち上げ方と、自治体や地域企業、学校と関係を築く動き方を学びたいです。トータルサポート推進係では、キャレオスのトータルサポートの全体像と事業部どうしのつなぎ方、そして放課後等デイサービスから就労へ切れ目なくつなぐ導線づくりを学びたいと考えています。現場で聞いた愛媛の声と、本部の知識と動き方を掛け合わせて、自分のボーダーを越えます。（約45秒）');
}

// ---------- 7. First step: B-type simulation ----------
{
  const s = pres.addSlide(); header(s, 'まずは、愛媛に就労継続支援B型を', 7);
  t(s, '西条市・定員20名で開設した場合の試算（標準シナリオ・2027年4月開設を想定）', { x: 0.5, y: 1.3, w: 9, h: 0.35, fontSize: 13, color: SUB });
  const stats = [['18人', '1日の平均利用者数\n（安定期）', RB[0]], ['6か月目', '単月で黒字化', RB[1]], ['27か月目', '初期投資\n約900万円を回収', RB[3]], ['1,140万円', '3年目の営業利益\n（概算）', RB[5]]];
  stats.forEach(([v, l, c], i) => {
    const x = 0.5 + i * 2.3;
    card(s, x, 1.85, 2.1, 1.75);
    t(s, v, { x: x + 0.1, y: 2.05, w: 1.9, h: 0.6, fontSize: 22, bold: true, color: c, align: 'center', valign: 'middle' });
    t(s, l, { x: x + 0.1, y: 2.75, w: 1.9, h: 0.65, fontSize: 11, color: SUB, align: 'center' });
  });
  const keys = [['放デイ卒業生の受け皿', '採算ラインは1日12人前後。日吉・西条の卒業予定者の把握が集客の土台'], ['工賃1.8万円の壁', '令和8年6月の見直しで区分の基準額が上昇。作業の開拓が収益に直結']];
  keys.forEach(([h, b], i) => {
    const x = 0.5 + i * 4.6;
    num(s, x, 3.85, String(i + 1), RB[4 + i * 2], 0.36, 12);
    t(s, h, { x: x + 0.5, y: 3.83, w: 3.9, h: 0.35, fontSize: 13.5, bold: true, valign: 'middle' });
    t(s, b, { x: x + 0.5, y: 4.2, w: 3.9, h: 0.6, fontSize: 11, color: SUB });
  });
  t(s, '※人件費・家賃・初期投資は仮置きの試算。報酬は令和8年6月の見直し・新規指定事業所の特例（984/1000）を反映', { x: 0.5, y: 4.95, w: 8.3, h: 0.3, fontSize: 9, color: SUB });
  s.addNotes('具体的な第一歩として、愛媛に就労継続支援B型をつくる構想を試算しました。西条市で定員20名、2027年4月開設を想定した標準シナリオでは、1日の平均利用者18人で、6か月目に単月黒字、27か月目に初期投資を回収、3年目の営業利益は約1,140万円です。ポイントは二つです。採算ラインは1日12人前後で、放デイの卒業生の受け皿になれるかが集客の土台になること。そして平均工賃1.8万円の壁を越える作業の開拓です。この試算を実際の数字で磨き上げることも、地域共創係・推進係で学びたいことです。（約45秒）');
}

// ---------- 7. Vision ----------
{
  const s = pres.addSlide(); header(s, '描く未来　― 愛媛から全国へ（仮構想）', 8);
  const steps = [
    ['STEP 1', '愛媛で実践', 'B型を第一歩に、\n放デイ・児童クラブ・\n就労をつなぐ', RB[0]],
    ['STEP 2', '四国支社', '愛媛から四国全域へ\nインドネシア特定技能\n等の人材も含めて', RB[2]],
    ['STEP 3', '関西・中部支社', 'エリアごとの拠点を\n各地に広げていく', RB[4]],
    ['GOAL', '全国へ', '全国にキャレオスの\nトータルサポートを', RB[6]],
  ];
  steps.forEach(([k, h, b, c], i) => {
    const x = 0.5 + i * 2.3, y = 1.55;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.1, h: 2.35, rectRadius: 0.12, fill: { color: LIGHT }, line: { color: LIGHT } });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.2, y: y + 0.22, w: 0.95, h: 0.32, rectRadius: 0.16, fill: { color: c }, line: { color: c } });
    t(s, k, { x: x + 0.2, y: y + 0.22, w: 0.95, h: 0.32, fontSize: 10, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    t(s, h, { x: x + 0.2, y: y + 0.7, w: 1.75, h: 0.45, fontSize: 17, bold: true });
    t(s, b, { x: x + 0.2, y: y + 1.25, w: 1.85, h: 1.0, fontSize: 10.5, color: SUB, lineSpacingMultiple: 1.15 });
  });
  card(s, 0.5, 4.15, 9.0, 0.85, 'EEF1FB');
  t(s, 'まだキャレオスにいない人材・ポジション。そこを確立させることに価値がある。', { x: 0.75, y: 4.15, w: 8.5, h: 0.85, fontSize: 15, bold: true, color: RB[5], valign: 'middle' });
  s.addNotes('ここからは、まだ妄想の域を出ない仮構想です。まず愛媛で、B型を第一歩に、放課後等デイサービス、児童クラブ、就労をつなぐトータルサポートの形をつくります。次に四国支社として、愛媛だけでなく四国全域へ。インドネシアの特定技能などの人材も含めて考えています。そして関西支社、中部支社と、全国へ広げていく。まだキャレオスにいない人材、ポジションだからこそ、それを確立させることに価値があると思っています。（約45秒）');
}

// ---------- 8. Closing ----------
{
  const s = pres.addSlide(); s.background = { color: WHITE };
  RB.forEach((c, i) => {
    const r = 3.1 - i * 0.3;
    s.addShape(pres.shapes.BLOCK_ARC, { x: -r, y: -r, w: r * 2, h: r * 2, fill: { color: c }, line: { color: c, width: 0 }, angleRange: [0, 90], arcThicknessRatio: 0.09 });
  });
  t(s, 'その始まりを、愛媛で実践する人材に。', { x: 1.0, y: 2.0, w: 8.5, h: 0.5, fontSize: 18, color: SUB, align: 'right' });
  t(s, '全国に\nトータルサポートを。', { x: 3.0, y: 2.55, w: 6.5, h: 1.6, fontSize: 40, bold: true, align: 'right', lineSpacingMultiple: 1.05 });
  dots(s, 9.5 - 7 * 0.18 + 0.07, 4.45);
  t(s, 'ご清聴ありがとうございました', { x: 3.0, y: 4.75, w: 6.5, h: 0.4, fontSize: 13, color: SUB, align: 'right' });
  s.addNotes('「ゆうゆうを卒業したあとも安心できる」と言っていただける地域をつくるために。全国にトータルサポートを届ける、その始まりを愛媛で実践する人材になります。ご清聴ありがとうございました。（約20秒）');
}

pres.writeFile({ fileName: 'crossborder_career_presentation.pptx' }).then(f => console.log(f));
