// 温湿度管理 基礎版 スライド生成スクリプト
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");
const path = require("path");
const { applyTheme } = require(process.env.SKILL_DIR + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const FONT = "Meiryo UI";
const THEME = {
  name: "TempHumid",
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: {
    dk1: "1B1F24", lt1: "FFFFFF", dk2: "14213D", lt2: "EEF3F7",
    accent1: "0E8C9A", // 青緑(湿度・管理)
    accent2: "E8772E", // オレンジ(熱・注意)
    accent3: "3D5A80", // 落ち着いた青
    accent4: "C0392B", // 赤(危険)
    accent5: "8AA1B1", // グレー青
    accent6: "2A9D8F",
    hlink: "0E8C9A", folHlink: "3D5A80",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "新しい製品の物流担当になったら — 温度と湿度の管理 入門";
pres.author = "物流研修";
const C = pres.SchemeColor;

const COURSE = "温度と湿度の管理 入門(基礎版)";

pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.5, w: 8.8, h: 1.5, fontSize: 32, bold: true, color: C.background1, valign: "middle", fontFace: FONT }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.lt1 },
  margin: [0.4, 0.5, 0.4, 0.5],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.28, w: 7.6, h: 0.62, fontSize: 24, bold: true, color: C.text2, valign: "middle", fontFace: FONT, margin: 0 }, text: "" } },
    { text: { text: COURSE, options: { x: 0.5, y: 5.25, w: 5, h: 0.25, fontSize: 9, color: C.accent5, fontFace: FONT, margin: 0 } } },
  ],
  slideNumber: { x: 9.0, y: 5.25, w: 0.5, h: 0.25, fontSize: 9, color: HEX.accent5, fontFace: FONT, align: "right" },
});

// ---- icons ----
const iconCache = {};
async function icon(name, color) {
  const key = name + color;
  if (iconCache[key]) return iconCache[key];
  const Comp = md[name];
  if (!Comp) throw new Error("icon not found: " + name);
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  iconCache[key] = "image/png;base64," + buf.toString("base64");
  return iconCache[key];
}
async function circleIcon(slide, name, x, y, d, bg, fg = "FFFFFF") {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg }, objectName: "icon-circle" });
  const pad = d * 0.22;
  slide.addImage({ data: await icon(name, fg), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, altText: name });
}

// ---- helpers ----
const T = (text, opt) => ({ text, options: opt || {} });
function chip(slide, label) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.25, y: 0.36, w: 1.25, h: 0.42, rectRadius: 0.21, fill: { color: HEX.lt2 }, line: { color: HEX.lt2 }, objectName: "lecture-chip" });
  slide.addText(label, { x: 8.25, y: 0.36, w: 1.25, h: 0.42, fontSize: 11, bold: true, color: C.accent1, align: "center", valign: "middle", isTextBox: true, margin: 0 });
}
function source(slide, text) {
  slide.addText(text, { x: 0.5, y: 4.88, w: 9.0, h: 0.32, fontSize: 9, color: C.accent5, isTextBox: true, margin: 0, valign: "top" });
}
function content(sec, lec, title) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText(title, { placeholder: "title" });
  chip(s, lec);
  return s;
}
function card(slide, x, y, w, h, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill: { color: fill || HEX.lt2 }, line: { color: fill || HEX.lt2 }, objectName: "card" });
}
function tx(slide, text, x, y, w, h, opt) {
  slide.addText(text, Object.assign({ x, y, w, h, fontSize: 14, color: C.text1, isTextBox: true, valign: "top", margin: 0.05 }, opt || {}));
}
// アイコン付きの列カード
async function iconCards(slide, items, y, h) {
  const n = items.length, gap = 0.3, W = 9.0, w = (W - gap * (n - 1)) / n;
  for (let i = 0; i < n; i++) {
    const it = items[i], x = 0.5 + i * (w + gap);
    card(slide, x, y, w, h, it.fill);
    await circleIcon(slide, it.icon, x + 0.25, y + 0.25, 0.62, it.color || HEX.accent1);
    tx(slide, it.head, x + 1.0, y + 0.27, w - 1.15, 0.6, { fontSize: 16, bold: true, color: C.text2, valign: "middle" });
    tx(slide, it.body, x + 0.25, y + 1.0, w - 0.45, h - 1.15, { fontSize: 14 });
  }
}
// 横の流れ図
async function flow(slide, steps, y, color) {
  const n = steps.length, gap = 0.22, W = 9.0, w = (W - gap * (n - 1)) / n;
  for (let i = 0; i < n; i++) {
    const x = 0.5 + i * (w + gap), st = steps[i];
    card(slide, x, y, w, 1.95, st.fill || HEX.lt2);
    await circleIcon(slide, st.icon, x + (w - 0.6) / 2, y + 0.18, 0.6, st.color || color || HEX.accent1);
    tx(slide, st.head, x + 0.06, y + 0.84, w - 0.12, 0.5, { fontSize: 13, bold: true, align: "center", valign: "middle", color: C.text2 });
    tx(slide, st.body, x + 0.1, y + 1.36, w - 0.2, 0.56, { fontSize: 11.5, align: "center", color: C.text1 });
    if (i < n - 1) slide.addShape(pres.shapes.RIGHT_TRIANGLE ? pres.shapes.ISOSCELES_TRIANGLE : pres.shapes.RECTANGLE, { x: x + w + 0.04, y: y + 0.85, w: 0.14, h: 0.25, rotate: 90, fill: { color: HEX.accent5 }, line: { color: HEX.accent5 }, objectName: "arrow" });
  }
}
async function exerciseSlide(sec, lec, title, points, task, notes) {
  const s = content(sec, lec, title);
  card(s, 0.5, 1.15, 4.3, 3.55, HEX.lt2);
  tx(s, "この回のポイント", 0.75, 1.3, 3.8, 0.4, { fontSize: 16, bold: true, color: C.accent1 });
  tx(s, points.map((p, i) => T(p, { bullet: true, breakLine: i < points.length - 1, paraSpaceAfter: 6 })), 0.75, 1.75, 3.85, 2.85, { fontSize: 14 });
  card(s, 5.1, 1.15, 4.4, 3.55, HEX.dk2);
  await circleIcon(s, "MdEditNote", 5.35, 1.35, 0.62, HEX.accent2);
  tx(s, "演習(動画を止めて考えましょう)", 6.1, 1.38, 3.3, 0.6, { fontSize: 15, bold: true, color: C.background1, valign: "middle" });
  tx(s, task, 5.35, 2.15, 3.95, 2.45, { fontSize: 14, color: C.background1 });
  s.addNotes(notes);
  return s;
}

(async () => {
  // =============== 第1回 ===============
  let sec = "第1回 なぜ新しい製品で温湿度管理が必要なのか", L = "第1回";
  pres.addSection({ title: sec });
  {
    const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: sec });
    s.addText("新しい製品の物流担当になったら\n温度と湿度の管理 入門", { placeholder: "title" });
    s.addText("基礎版 ― 電子部品・充電用電池の保管・荷役・輸配送", { x: 0.6, y: 3.0, w: 8.8, h: 0.45, fontSize: 16, color: C.background2, isTextBox: true, margin: 0 });
    const lec = ["第1回 なぜ温湿度管理が必要か", "第2回 条件と基準の読み方", "第3回 保管", "第4回 荷役と輸配送", "第5回 測る・記録する・動く"];
    s.addText(lec.map((t, i) => T(t, { breakLine: i < lec.length - 1 })), { x: 0.6, y: 3.6, w: 8.8, h: 1.4, fontSize: 12, color: C.background2, isTextBox: true, margin: 0, lineSpacingMultiple: 1.1 });
    await circleIcon(s, "MdDeviceThermostat", 8.35, 0.55, 0.9, HEX.accent2);
    await circleIcon(s, "MdWaterDrop", 7.3, 0.55, 0.9, HEX.accent1);
    s.addNotes("この講座は、「新しい製品の物流担当になったら、温度と湿度の管理 入門」の基礎版です。対象は、半導体などの電子部品や、充電用の電池といった、新しい製品の物流を担当することになった方です。全5回で、温度と湿度の管理の基本を、保管、荷役、輸配送の順に学びます。第1回では、なぜ温度と湿度の管理が必要なのか。第2回では、条件と基準の読み方。第3回では保管。第4回では荷役と輸配送。第5回では、測る、記録する、外れたときに動く、を学びます。なお、この講座の音声は、音声合成ソフトによるAI音声です。また、スライドで紹介する数値や制度には、出典を示しています。");
  }
  {
    const s = content(sec, L, "温度管理は、食品だけの話ではない");
    await iconCards(s, [
      { icon: "MdRestaurant", head: "食品・医薬品", body: "冷凍・冷蔵などの温度管理が早くから発達してきた分野。管理の考え方の手本になる", color: HEX.accent5 },
      { icon: "MdMemory", head: "電子部品", body: "半導体チップ、IC、センサーなど。湿気・結露・静電気に弱いものがある", color: HEX.accent1 },
      { icon: "MdBatteryChargingFull", head: "充電用電池", body: "電池そのものと、電池を組み込んだ製品。高温や強い衝撃が事故につながるおそれがある", color: HEX.accent2 },
    ], 1.2, 2.3);
    card(s, 0.5, 3.75, 9.0, 0.95, HEX.dk2);
    tx(s, "これまで温度を気にせず運んできた方も、担当する製品が変われば、温湿度の管理が必要になります", 0.75, 3.8, 8.5, 0.85, { fontSize: 15, bold: true, color: C.background1, valign: "middle" });
    s.addNotes("温度管理と聞くと、冷凍食品や医薬品を思い浮かべる方が多いと思います。実際、温度管理の考え方は、食品や医薬品の物流で早くから発達してきました。この講座でも、その考え方を手本として借ります。一方で、温度や湿度の管理が必要なのは、食品や医薬品だけではありません。半導体チップやセンサーなどの電子部品には、湿気や結露、静電気に弱いものがあります。また、充電用の電池や、電池を組み込んだ製品は、高温や強い衝撃が事故につながるおそれがあります。これまで温度を気にせずに運んできた方も、こうした製品を担当すると、温度と湿度の管理が必要になります。");
  }
  {
    const s = content(sec, L, "製品によって「守るもの」が違う");
    card(s, 0.5, 1.2, 4.35, 3.45, HEX.lt2);
    await circleIcon(s, "MdMemory", 0.75, 1.4, 0.7, HEX.accent1);
    tx(s, "電子部品", 1.6, 1.45, 3.0, 0.6, { fontSize: 20, bold: true, color: C.text2, valign: "middle" });
    tx(s, [T("守るもの: 品質", { bold: true, breakLine: true, color: C.accent1 }), T("湿気を吸うと、取り付けの工程で不具合の原因になることがある", { bullet: true, breakLine: true }), T("急な温度変化で、結露が起きる", { bullet: true, breakLine: true }), T("乾燥しすぎると、静電気が起きやすい", { bullet: true })], 0.75, 2.25, 3.9, 2.3, { fontSize: 14, paraSpaceAfter: 4 });
    card(s, 5.15, 1.2, 4.35, 3.45, "FDF0E7");
    await circleIcon(s, "MdBatteryChargingFull", 5.4, 1.4, 0.7, HEX.accent2);
    tx(s, "充電用電池", 6.25, 1.45, 3.0, 0.6, { fontSize: 20, bold: true, color: C.text2, valign: "middle" });
    tx(s, [T("守るもの: 品質 + 安全", { bold: true, breakLine: true, color: C.accent2 }), T("メーカーが示す保存温度の範囲を守る", { bullet: true, breakLine: true }), T("高温は発火につながるおそれ", { bullet: true, breakLine: true }), T("圧迫・振動・落下などの強い衝撃を避ける", { bullet: true })], 5.4, 2.25, 3.9, 2.3, { fontSize: 14, paraSpaceAfter: 4 });
    source(s, "出典: 製品評価技術基盤機構(NITE)「リチウムイオン電池関連製品の製造と安全性」/ 電子部品の湿気の扱いは業界の規格(IPC/JEDEC J-STD-033)の考え方による");
    s.addNotes("温度と湿度の管理で何を守るのかは、製品によって違います。電子部品で主に守るのは、品質です。湿気に弱い電子部品は、湿気を吸ったまま基板に取り付ける工程に進むと、不具合の原因になることがあります。また、急な温度の変化で結露が起きたり、反対に乾燥しすぎると静電気が起きやすくなったりします。一方、充電用の電池では、品質に加えて、安全を守る必要があります。製品評価技術基盤機構、ナイトの資料では、電池を使うときの条件として、保存するときの許容温度の範囲を守ること、そして、圧迫や振動、衝撃を避けることが挙げられています。高温や強い衝撃は、発火につながるおそれがあるためです。");
  }
  {
    const s = content(sec, L, "電池は、高温が危ない");
    const stats = [["2,140件", "リチウムイオン電池を搭載した製品の事故(2021〜2025年の5年間)", HEX.accent3], ["1,901件", "そのうち火災が起きた事故(約9割)", HEX.accent4], ["8月", "事故が最も多くなる月。春から夏にかけて気温の上昇とともに増える", HEX.accent2]];
    for (let i = 0; i < 3; i++) {
      const x = 0.5 + i * 3.1;
      card(s, x, 1.2, 2.8, 2.6, HEX.lt2);
      tx(s, stats[i][0], x + 0.15, 1.4, 2.5, 0.9, { fontSize: 34, bold: true, color: stats[i][2], align: "center", valign: "middle" });
      tx(s, stats[i][1], x + 0.2, 2.4, 2.4, 1.3, { fontSize: 13, align: "center" });
    }
    tx(s, "※ モバイルバッテリーなど消費者向け製品の事故の集計。物流の荷物の事故件数ではありませんが、「高温」と「衝撃」が危ないことを示しています", 0.5, 3.95, 9.0, 0.7, { fontSize: 12, color: C.text2 });
    source(s, "出典: 製品評価技術基盤機構(NITE)ニュースリリース「『夏バテ(夏のバッテリー)』にご注意を」2026年7月15日");
    s.addNotes("電池で特に注意したいのが、高温です。製品評価技術基盤機構、ナイトによると、2021年から2025年までの5年間に通知された、リチウムイオン電池を搭載した製品の事故は2140件ありました。そのうち、火災が起きた事故は1901件で、およそ9割にのぼります。また、事故は春から夏にかけて、気温の上昇とともに増え、8月にピークを迎えます。ナイトは、事故を防ぐポイントの一つとして、高温となる場所では使用・保管しないことを挙げています。なお、この数字はモバイルバッテリーなど、消費者が使う製品の事故の集計で、物流の荷物の事故件数ではありません。それでも、高温と衝撃が危ないことは、物流で電池を扱うときにも同じです。");
  }
  {
    const s = content(sec, L, "温湿度管理は「つながり」で守る");
    await flow(s, [
      { icon: "MdFactory", head: "メーカー", body: "条件を決め、伝える" },
      { icon: "MdWarehouse", head: "倉庫", body: "保管・荷役で守る" },
      { icon: "MdLocalShipping", head: "運送会社", body: "輸送中に守る" },
      { icon: "MdStore", head: "納品先", body: "受け取りで守る" },
    ], 1.3);
    card(s, 0.5, 3.55, 9.0, 1.15, "FDF0E7");
    tx(s, [T("どこか1か所で条件が外れると、ほかの人たちの管理が生かせなくなります。", { bold: true, breakLine: true }), T("自分の担当の前後で、だれが何をしているかを知ることが出発点です。")], 0.75, 3.68, 8.5, 0.95, { fontSize: 14, color: C.text2 });
    s.addNotes("温度と湿度の管理は、だれか1人では守れません。メーカーが条件を決めて伝え、倉庫が保管と荷役で守り、運送会社が輸送中に守り、納品先が受け取りで守ります。この流れのどこか1か所で条件が外れると、ほかの人たちがどれだけ丁寧に管理していても、その努力が生かせなくなってしまいます。ですから、まず自分の担当の前と後ろで、だれが何をしているのかを知ることが、温度と湿度の管理の出発点になります。例えば、メーカーが条件を伝え忘れれば、倉庫も運送会社も、その製品に温度管理が必要だと気づけません。反対に、倉庫がどれだけ丁寧に保管しても、トラックの荷台で長い時間、日に当たってしまえば意味がありません。この講座では、自分の担当だけでなく、前後の工程にも目を向けることを意識して学んでいきましょう。");
  }
  await exerciseSlide(sec, L, "第1回のまとめ", ["温度・湿度の管理は、食品・医薬品以外の製品でも必要", "電子部品は主に「品質」、電池は「品質」と「安全」を守る", "電池は高温と強い衝撃が危ない", "管理は、メーカー・倉庫・運送・納品先のつながりで守る"], "担当する製品を1つ選び、温度や湿度が外れたら何が困るかを書き出してみましょう。\n(品質のこと、安全のこと、取引先のこと)",
    "第1回のまとめです。温度と湿度の管理は、食品や医薬品だけでなく、電子部品や電池でも必要です。電子部品では主に品質を、電池では品質と安全を守ります。特に電池は、高温と強い衝撃が危険です。そして、温度と湿度の管理は、メーカー、倉庫、運送会社、納品先のつながりで守ります。ここで演習です。動画を止めて、あなたが担当する製品を1つ選び、温度や湿度が外れたら何が困るかを書き出してみてください。品質のこと、安全のこと、取引先との関係のこと、の3つの面から考えてみましょう。書き出した内容は、第2回以降の演習でも使います。次回は、温度と湿度の条件と基準の読み方を学びます。");

  // =============== 第2回 ===============
  sec = "第2回 条件と基準の読み方"; L = "第2回";
  pres.addSection({ title: sec });
  {
    const s = content(sec, L, "「常温」は何℃? 言葉ではなく数字で確かめる");
    card(s, 0.5, 1.2, 4.35, 3.45, HEX.lt2);
    tx(s, "例: 医薬品の決まり(日本薬局方)", 0.75, 1.35, 3.9, 0.4, { fontSize: 15, bold: true, color: C.accent1 });
    const rows = [["常温", "15〜25℃"], ["室温", "1〜30℃"], ["冷所", "1〜15℃"]];
    rows.forEach((r, i) => {
      tx(s, r[0], 0.9, 1.95 + i * 0.75, 1.4, 0.6, { fontSize: 20, bold: true, color: C.text2, valign: "middle" });
      tx(s, r[1], 2.3, 1.95 + i * 0.75, 2.4, 0.6, { fontSize: 24, bold: true, color: C.accent1, valign: "middle" });
    });
    await circleIcon(s, "MdHelpOutline", 5.4, 1.4, 0.7, HEX.accent2);
    tx(s, "同じ言葉でも、範囲は会社や業界で違う", 6.25, 1.4, 3.1, 0.75, { fontSize: 16, bold: true, color: C.text2, valign: "middle" });
    tx(s, [T("「常温で」「定温で」と言われたら、必ず「何℃から何℃か」「湿度は何%以下か」を確かめる", { bullet: true, breakLine: true }), T("数字が書かれた文書(仕様書・指示書・契約)をもらう", { bullet: true })], 5.4, 2.4, 4.0, 2.2, { fontSize: 14, paraSpaceAfter: 6 });
    source(s, "出典: 第十九改正日本薬局方 通則(令和8年4月10日 厚生労働省告示第193号)");
    s.addNotes("第2回は、条件と基準の読み方です。この回では、温度や湿度の条件を、どう確かめ、どう読めばよいのかを学びます。言葉のあいまいさ、条件の根拠の種類、メーカーの仕様書、湿度の指定、そして電池の充電率の順に見ていきます。まず、「常温」と言われたら、何度のことだと思いますか。例えば、医薬品の決まりである日本薬局方では、常温は15度から25度、室温は1度から30度、冷所は1度から15度と決められています。ところが、同じ「常温」という言葉でも、会社や業界によって指している範囲は違います。ですから、「常温で」「定温で」と言われたら、必ず、何度から何度までか、湿度は何パーセント以下か、を数字で確かめてください。そして、口頭だけでなく、数字が書かれた仕様書や指示書、契約書などの文書をもらうようにしましょう。");
  }
  {
    const s = content(sec, L, "温湿度の条件の根拠は4種類");
    const items = [
      ["MdGavel", "① 法令", "守らなければならない決まり。例: 消防法(電池の保管)", HEX.accent4],
      ["MdPublic", "② 業界・国際の規格", "業界団体や国際機関が決めたルール。例: 電子部品の湿気の扱い", HEX.accent3],
      ["MdDescription", "③ メーカーの仕様書", "製品ごとの条件。保存温度、湿度、取り扱いの注意", HEX.accent1],
      ["MdHandshake", "④ 契約・取引先の指示", "取引先との約束。法令より厳しいことも多い", HEX.accent2],
    ];
    for (let i = 0; i < 4; i++) {
      const x = 0.5 + (i % 2) * 4.65, y = 1.15 + Math.floor(i / 2) * 1.8;
      card(s, x, y, 4.35, 1.6, HEX.lt2);
      await circleIcon(s, items[i][0], x + 0.22, y + 0.25, 0.62, items[i][3]);
      tx(s, items[i][1], x + 1.0, y + 0.22, 3.2, 0.45, { fontSize: 16, bold: true, color: C.text2, valign: "middle" });
      tx(s, items[i][2], x + 1.0, y + 0.72, 3.2, 0.8, { fontSize: 13 });
    }
    tx(s, "条件を聞いたら「その根拠はどれですか」と確かめると、どこまで厳しく守るべきかがわかります", 0.5, 4.75, 9.0, 0.4, { fontSize: 13, bold: true, color: C.accent1 });
    s.addNotes("温度や湿度の条件には、根拠があります。大きく分けると4種類です。1つ目は法令で、必ず守らなければならない決まりです。例えば、電池の保管には消防法が関わります。2つ目は業界や国際の規格で、例えば電子部品の湿気の扱いには、業界の規格があります。3つ目はメーカーの仕様書で、製品ごとの保存温度や湿度、取り扱いの注意が書かれています。4つ目は契約や取引先の指示で、法令より厳しい条件が求められることも少なくありません。条件を聞いたら、「その根拠はどれですか」と確かめてみてください。根拠がわかると、どこまで厳しく守るべきかが判断しやすくなります。");
  }
  {
    const s = content(sec, L, "出発点は、メーカーの仕様書");
    await circleIcon(s, "MdDescription", 0.6, 1.3, 0.9, HEX.accent1);
    tx(s, "電池の「保存温度」の推奨値は、資料によってばらばら", 1.7, 1.35, 7.8, 0.8, { fontSize: 18, bold: true, color: C.text2, valign: "middle" });
    card(s, 0.5, 2.35, 9.0, 1.45, HEX.lt2);
    tx(s, [T("インターネット上には、さまざまな「推奨値」がありますが、数値は資料ごとに違います。", { breakLine: true }), T("公的機関の資料でも、電池を使うときの条件として「保存時の許容温度範囲」を守ることが挙げられています。その範囲を示すのは、製品のメーカーです。")], 0.75, 2.45, 8.5, 1.3, { fontSize: 14 });
    card(s, 0.5, 3.95, 9.0, 0.8, HEX.dk2);
    tx(s, "担当する製品の仕様書・取扱説明書で、温度・湿度・充電率などの条件を確かめる", 0.75, 4.0, 8.5, 0.7, { fontSize: 15, bold: true, color: C.background1, valign: "middle" });
    source(s, "出典: 製品評価技術基盤機構(NITE)「リチウムイオン電池関連製品の製造と安全性」");
    s.addNotes("では、具体的な温度の条件は何を見ればよいのでしょうか。出発点は、メーカーの仕様書です。例えば電池の保存温度について調べると、インターネット上にはさまざまな推奨値が出てきますが、その数値は資料ごとにばらばらです。製品評価技術基盤機構、ナイトの資料でも、電池を使うときの条件として、保存するときの許容温度の範囲を守ることが挙げられています。そして、その範囲を示すのは、製品を作ったメーカーです。ですから、一般的な数字をうのみにせず、担当する製品の仕様書や取扱説明書で、温度、湿度、そして電池であれば充電率などの条件を確かめてください。もし仕様書に条件が書かれていなかったり、書かれていても意味がわかりにくかったりする場合は、遠慮せずにメーカーや出荷元に質問しましょう。質問した内容と回答は、記録に残しておくと、後で担当者が変わっても困りません。");
  }
  {
    const s = content(sec, L, "湿度の指定を読む ― 湿気に弱い電子部品");
    await iconCards(s, [
      { icon: "MdInventory2", head: "防湿袋", body: "湿気を通しにくい袋で密封する", color: HEX.accent1 },
      { icon: "MdGrain", head: "乾燥剤", body: "袋の中の湿気を吸い取る", color: HEX.accent3 },
      { icon: "MdPalette", head: "湿度インジケーター", body: "袋の中の湿気を色の変化で示すカード", color: HEX.accent6 },
    ], 1.15, 2.2);
    card(s, 0.5, 3.55, 9.0, 1.15, "FDF0E7");
    tx(s, [T("湿気への弱さは、業界の規格でレベル分けされています(MSL: 湿気への敏感さのレベル)。", { breakLine: true }), T("物流の担当者は「袋を勝手に開けない」「表示されている条件を守る」「開封時の確認方法をメーカーに聞く」が基本です。")], 0.75, 3.63, 8.5, 1.0, { fontSize: 13.5 });
    source(s, "参考: 業界規格 IPC/JEDEC J-STD-033(湿気に敏感な部品の取り扱い)。具体的な条件は規格・メーカーの資料で確認");
    s.addNotes("次に、湿度の指定の読み方です。湿気に弱い電子部品は、多くの場合、3つのもので守られています。1つ目は防湿袋で、湿気を通しにくい袋に密封します。2つ目は乾燥剤で、袋の中の湿気を吸い取ります。3つ目は湿度インジケーターで、袋の中の湿気を色の変化で知らせるカードです。電子部品の湿気への弱さは、業界の規格で、MSLと呼ばれるレベルに分けられています。物流の担当者として大切なのは、まず、袋を勝手に開けないこと。次に、袋や箱に表示されている条件を守ること。そして、開けたときの確認方法をメーカーに聞いておくことです。防湿袋に穴が開いていたり、封が破れていたりするのを見つけたら、そのまま出荷せずに、出荷元に報告しましょう。");
  }
  {
    const s = content(sec, L, "電池は「充電率」も条件になる");
    card(s, 0.5, 1.2, 3.6, 3.45, HEX.dk2);
    tx(s, "充電率とは", 0.75, 1.4, 3.1, 0.45, { fontSize: 16, bold: true, color: C.background1 });
    tx(s, "電池にどれだけ電気がたまっているかの割合。満充電を100%として表します", 0.75, 1.9, 3.15, 1.3, { fontSize: 14, color: C.background1 });
    await circleIcon(s, "MdBattery3Bar", 1.7, 3.25, 1.1, HEX.accent2);
    tx(s, "例: 消防法の保管の特例(2023年改正)", 4.4, 1.25, 5.1, 0.45, { fontSize: 16, bold: true, color: C.text2 });
    card(s, 4.4, 1.75, 5.1, 1.2, HEX.lt2);
    tx(s, [T("電池だけを保管する倉庫の特例の条件の一つに", { breakLine: true }), T("「蓄電池の充電率は60%以下」", { bold: true, color: C.accent2 })], 4.6, 1.85, 4.7, 1.0, { fontSize: 15, valign: "middle" });
    tx(s, [T("航空輸送などのルールにも、充電率の条件がある", { bullet: true, breakLine: true }), T("充電率はメーカーや出荷元に確かめる", { bullet: true, breakLine: true }), T("規制は見直しが続いている(第3回で扱います)", { bullet: true })], 4.4, 3.1, 5.1, 1.6, { fontSize: 14, paraSpaceAfter: 4 });
    source(s, "出典: 消防庁 報道資料(令和5年12月6日)危険物の規制に関する政令等の一部改正");
    s.addNotes("電池の場合、温度や湿度に加えて、充電率も条件になることがあります。充電率とは、電池にどれだけ電気がたまっているかの割合で、満充電を100パーセントとして表します。例えば、2023年の消防法の改正では、電池だけを保管する倉庫について、建物の規制を緩める特例ができました。その条件の一つが、蓄電池の充電率を60パーセント以下にすることです。また、航空輸送などのルールにも、充電率の条件があります。充電率は物流の現場では測れないことが多いので、メーカーや出荷元に確かめるようにしましょう。また、充電率の条件は、保管の場面と輸送の場面で違うことがあります。どの場面の条件なのかも、あわせて確かめておきましょう。なお、こうした規制は見直しが続いています。詳しくは第3回で扱います。");
  }
  await exerciseSlide(sec, L, "製品の温度と、周りの温度 ― 第2回のまとめ", ["庫内の空気の温度と、製品そのものの温度は違う(冷えるのも温まるのも時間差がある)", "条件は数字で、文書で確かめる", "根拠は4種類。出発点はメーカーの仕様書", "電子部品は湿気、電池は充電率も条件になる"], "担当する製品の温度・湿度の条件と、その根拠を確かめましょう。\n・何℃から何℃か、湿度は何%以下か\n・根拠は4種類のどれか\n※ 配布物「温湿度条件の確認シート」を使います",
    "第2回のまとめの前に、1つ大切なことをお伝えします。倉庫や荷台の空気の温度と、製品そのものの温度は違います。空気が冷えても、製品はすぐには冷えませんし、温まるときも時間差があります。温度計が示しているのが空気の温度なのか、製品の温度なのかを意識してください。では、まとめです。条件は数字で、そして文書で確かめること。根拠は4種類あり、出発点はメーカーの仕様書であること。電子部品では湿気が、電池では充電率も条件になること。ここで演習です。配布物の温湿度条件の確認シートを使って、担当する製品の条件と、その根拠を書き出してみてください。");

  // =============== 第3回 ===============
  sec = "第3回 保管の温湿度管理"; L = "第3回";
  pres.addSection({ title: sec });
  {
    const s = content(sec, L, "倉庫の中には、温度の「むら」がある");
    card(s, 0.5, 1.2, 5.0, 3.45, HEX.lt2);
    // 倉庫の断面図(簡略)
    s.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 1.65, w: 4.2, h: 2.75, fill: { color: "FFFFFF" }, line: { color: HEX.accent5, width: 1.5 }, objectName: "warehouse" });
    const spots = [[1.1, 1.85, "天井付近\n熱がたまることも", HEX.accent2], [3.6, 1.85, "窓・壁際\n日差し・外気", HEX.accent2], [1.1, 3.35, "扉の近く\n出入りで変化", HEX.accent1], [3.6, 3.35, "奥・床付近\n倉庫により違う", HEX.accent3]];
    tx(s, "例: 温度が変わりやすいことがある場所", 0.9, 1.25, 4.2, 0.3, { fontSize: 12, bold: true, color: C.accent1 });
    spots.forEach(p => { s.addShape(pres.shapes.OVAL, { x: p[0], y: p[1], w: 0.25, h: 0.25, fill: { color: p[3] }, line: { color: p[3] }, objectName: "spot" }); tx(s, p[2], p[0] + 0.3, p[1] - 0.05, 1.45, 0.8, { fontSize: 11, color: C.text2 }); });
    await circleIcon(s, "MdThermostat", 5.8, 1.3, 0.7, HEX.accent1);
    tx(s, "温度マッピング", 6.65, 1.35, 2.8, 0.6, { fontSize: 17, bold: true, color: C.text2, valign: "middle" });
    tx(s, [T("使い始める前に、場所ごとの温度を測って、むらを知る", { bullet: true, breakLine: true }), T("結果をもとに、温度計を置く場所と、製品を置く場所を決める", { bullet: true, breakLine: true }), T("季節や設備が変わったら測り直す", { bullet: true })], 5.8, 2.15, 3.7, 2.5, { fontSize: 14, paraSpaceAfter: 6 });
    source(s, "参考: 厚生労働省「医薬品の適正流通(GDP)ガイドライン」(平成30年12月)3.3 温度及び環境管理 ― 医薬品で使われている考え方を借りています");
    s.addNotes("第3回は、保管の温度と湿度の管理です。この回では、倉庫で製品を保管するときに、温度と湿度のどこに気をつければよいかを学びます。温度のむら、夏の高温と結露、湿度、そして電池の保管と消防法の順に見ていきます。倉庫の中の温度は、どこでも同じではありません。例えば、天井付近に熱がたまったり、窓や壁の近くが日差しや外気の影響を受けたり、扉の近くが人や荷物の出入りで温度が変わったりすることがあります。どこで温度が変わりやすいかは、倉庫によって違います。こうした温度のむらを知る方法が、温度マッピングです。倉庫を使い始める前に、場所ごとの温度を測り、その結果をもとに、温度計を置く場所と、製品を置く場所を決めます。季節が変わったときや、空調などの設備を変えたときは、測り直します。これは、医薬品の物流のガイドラインで求められている考え方ですが、電子部品や電池の保管にも役立ちます。");
  }
  {
    const s = content(sec, L, "夏の高温と、結露");
    await iconCards(s, [
      { icon: "MdWbSunny", head: "夏の高温", body: "屋根や壁からの熱、直射日光で、倉庫の中は外より暑くなることがある。電池は特に注意", color: HEX.accent2, fill: "FDF0E7" },
      { icon: "MdWaterDrop", head: "結露", body: "冷えた荷物を、暖かく湿った場所に出すと、表面に水滴がつく。電子部品の大敵", color: HEX.accent1 },
    ], 1.15, 2.05);
    card(s, 0.5, 3.4, 9.0, 1.3, HEX.dk2);
    tx(s, [T("結露が起きるしくみ", { bold: true, breakLine: true, color: C.accent2 }), T("空気が含める水分の量は、温度が低いほど少なくなります。冷たい物の表面で空気が冷やされ、含みきれなくなった水分が水滴になります(水滴ができ始める温度を「露点」といいます)。冷えた荷物は、周りの温度になじませてから開けましょう")], 0.75, 3.48, 8.5, 1.15, { fontSize: 13.5, color: C.background1 });
    source(s, "参考: 熊本県立大学 辻原万規彦「建築環境工学I 第7回 湿度/結露」2023年5月30日(結露の原因と防止の考え方)");
    s.addNotes("季節ごとに注意したいことがあります。夏は高温です。屋根や壁からの熱や直射日光で、倉庫の中が外より暑くなることがあります。電池を保管するときは特に注意が必要です。もう一つは結露です。冷えた荷物を、暖かく湿った場所に出すと、表面に水滴がつきます。これが結露で、電子部品の大敵です。結露が起きるしくみを確認しておきましょう。空気が含むことのできる水分の量は、温度が低いほど少なくなります。冷たい物の表面で空気が冷やされると、含みきれなくなった水分が水滴になります。水滴ができ始める温度を露点といいます。ですから、冷えた荷物は、すぐに開けずに、周りの温度になじませてから開けることが大切です。");
  }
  {
    const s = content(sec, L, "湿度は高すぎても、低すぎても困る");
    // 横軸の図
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.0, w: 8.4, h: 0.35, fill: { color: HEX.lt2 }, line: { color: HEX.lt2 }, objectName: "scale" });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.0, w: 2.2, h: 0.35, fill: { color: "F6C9A8" }, line: { color: "F6C9A8" }, objectName: "low" });
    s.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: 2.0, w: 2.2, h: 0.35, fill: { color: "A8DCE1" }, line: { color: "A8DCE1" }, objectName: "high" });
    tx(s, "乾燥", 0.8, 1.55, 2.2, 0.4, { fontSize: 14, bold: true, color: C.accent2, align: "center" });
    tx(s, "条件の範囲", 3.0, 1.55, 4.0, 0.4, { fontSize: 14, bold: true, color: C.text2, align: "center" });
    tx(s, "高湿度", 7.0, 1.55, 2.2, 0.4, { fontSize: 14, bold: true, color: C.accent1, align: "center" });
    await iconCards(s, [
      { icon: "MdBolt", head: "低すぎると", body: "静電気が起きやすい。電子部品が静電気で壊れることがある", color: HEX.accent2, fill: "FDF0E7" },
      { icon: "MdWaterDrop", head: "高すぎると", body: "湿気を吸う、結露する、カビが生える、段ボールが弱くなる", color: HEX.accent1 },
    ], 2.6, 1.6);
    tx(s, "電子部品は、帯電防止の包装や、決められた取り扱い方法を守ることが基本です", 0.5, 4.35, 9.0, 0.4, { fontSize: 13, bold: true, color: C.text2 });
    source(s, "参考: 村田製作所「静電気の基礎」、JAXA 技術標準(静電気放電対策)。湿度の具体的な数値は製品の条件で確認");
    s.addNotes("湿度は、高すぎても、低すぎても困ります。湿度が高すぎると、製品が湿気を吸ったり、結露したり、カビが生えたり、段ボールが弱くなったりします。反対に、空調の効いた室内などで湿度が極端に低くなると、静電気が起きやすくなります。電子部品の中には、静電気で壊れてしまうものがあります。ですから、湿度は、製品の条件で決められた範囲に保つことが大切です。具体的な数値は、製品や会社によって違いますので、仕様書で確かめてください。また、電子部品は、帯電防止の包装や、決められた取り扱い方法を守ることが基本です。湿度計がない場所で電子部品を扱っている場合は、まず、湿度を測るところから始めてみてください。");
  }
  {
    const s = content(sec, L, "電池を大量に保管するときは、消防法");
    await circleIcon(s, "MdLocalFireDepartment", 0.6, 1.25, 0.9, HEX.accent4);
    tx(s, "リチウムイオン電池の電解液は、消防法の「危険物」にあたる", 1.7, 1.3, 7.8, 0.8, { fontSize: 18, bold: true, color: C.text2, valign: "middle" });
    card(s, 0.5, 2.3, 4.35, 2.4, HEX.lt2);
    tx(s, [T("ポイント", { bold: true, color: C.accent4, breakLine: true }), T("電解液は燃えやすい液体(第4類)", { bullet: true, breakLine: true }), T("保管する電池の量が一定を超えると、倉庫は消防法の基準に合った施設にする必要がある", { bullet: true })], 0.75, 2.4, 3.95, 2.2, { fontSize: 14, paraSpaceAfter: 4 });
    card(s, 5.15, 2.3, 4.35, 2.4, "FDF0E7");
    tx(s, [T("確かめる相手", { bold: true, color: C.accent2, breakLine: true }), T("自社の倉庫が対象になるか", { bullet: true, breakLine: true }), T("電池を組み込んだ製品の場合の扱い", { bullet: true, breakLine: true }), T("→ 社内の担当部署、倉庫会社、所轄の消防に相談する", { breakLine: false })], 5.4, 2.4, 3.95, 2.2, { fontSize: 14, paraSpaceAfter: 4 });
    source(s, "出典: 消防庁 検討会資料「現行の消防法令の基準について」、消防庁 報道資料(令和5年12月6日・令和7年5月14日)");
    s.addNotes("電池の保管では、消防法も関わってきます。リチウムイオン電池の中の電解液は、燃えやすい液体で、消防法の危険物の第4類にあたります。そのため、保管する電池の量が一定を超えると、その倉庫は、消防法の基準に合った施設にする必要があります。この講座では、具体的な量や設備の基準までは扱いません。大切なのは、自社の倉庫が対象になるかどうかを確かめることです。また、電池を組み込んだ製品の場合にどう扱われるかも、確認が必要です。こうしたことは、自分で判断せず、社内の担当部署や倉庫会社、そして所轄の消防に相談してください。倉庫会社に保管を任せている場合も、電池を預けることを事前に伝え、受け入れられるかどうかを確かめておきましょう。");
  }
  {
    const s = content(sec, L, "【変わりやすい情報】電池の保管規制の見直し");
    const tl = [
      ["2023年12月", "公布(電池の規定は翌日施行)", "電池だけを保管する倉庫の特例。条件: 充電率60%以下、水が浸みこむ素材(段ボールなど)で包装する、スプリンクラーなど"],
      ["2025年5月", "公布(翌日施行)", "容器に入れずに保管・運搬できる方法を整備(段ボールなどで包装、耐火性のある箱 など)。電池の製造・充放電を行う施設の特例"],
    ];
    for (let i = 0; i < 2; i++) {
      const y = 1.15 + i * 1.65;
      card(s, 0.5, y, 9.0, 1.45, i === 0 ? HEX.lt2 : "FDF0E7");
      tx(s, tl[i][0], 0.7, y + 0.15, 1.9, 0.5, { fontSize: 20, bold: true, color: C.text2 });
      tx(s, tl[i][1], 0.7, y + 0.7, 1.9, 0.5, { fontSize: 12, color: C.text2 });
      tx(s, tl[i][2], 2.7, y + 0.15, 6.6, 1.2, { fontSize: 14, valign: "middle" });
    }
    tx(s, "最新の内容は消防庁の資料や所轄の消防で確認を(最終確認日: 2026年11月24日)", 0.5, 4.42, 9.0, 0.4, { fontSize: 12, bold: true, color: C.accent4 });
    source(s, "出典: 消防庁 報道資料「危険物の規制に関する政令の一部を改正する政令等の公布」令和5年12月6日・令和7年5月14日");
    s.addNotes("この画面は、変わりやすい情報です。電池の保管の規制は、ここ数年、見直しが続いています。2023年12月には、電池だけを保管する倉庫について、建物の規制を緩める特例ができました。その条件として、充電率を60パーセント以下にすること、水が浸みこむ段ボールなどの素材で包装すること、スプリンクラーなどの消火設備を設けることなどが定められています。さらに2025年5月には、容器に入れずに保管したり運搬したりできる方法などが整備されました。このように、規制は変わり続けています。最新の内容は、消防庁の資料や所轄の消防で確かめてください。この画面の情報は、2026年11月24日に確認したものです。");
  }
  await exerciseSlide(sec, L, "第3回のまとめ", ["倉庫の中には温度のむらがある。温度マッピングで知る", "夏は高温、温度差があると結露。冷えた荷物は周りの温度になじませてから開ける", "湿度は高すぎても低すぎても困る", "電池の大量保管は消防法。自分で判断せず相談する"], "自社の倉庫で、温度や湿度が外れやすい場所を書き出してみましょう。\n・天井付近、窓際、扉の近く\n・夏と冬で変わる場所",
    "第3回のまとめです。倉庫の中には温度のむらがあり、それを知る方法が温度マッピングでした。夏は高温に、温度差があるときは結露に注意し、冷えた荷物は周りの温度になじませてから開けます。湿度は、高すぎても低すぎても困ります。そして、電池を大量に保管するときは消防法が関わるため、自分で判断せずに相談することが大切です。演習です。動画を止めて、自社の倉庫で、温度や湿度が外れやすい場所を書き出してみてください。天井付近、窓際、扉の近く、そして夏と冬で状態が変わる場所に注目してみましょう。書き出した場所は、第5回で学ぶ、温度計やデータロガーを置く場所を考えるときにも役立ちます。");

  // =============== 第4回 ===============
  sec = "第4回 荷役と輸配送の温湿度管理"; L = "第4回";
  pres.addSection({ title: sec });
  {
    const s = content(sec, L, "荷役は「すき間の時間」");
    await flow(s, [
      { icon: "MdWarehouse", head: "倉庫で保管", body: "管理されている", color: HEX.accent1 },
      { icon: "MdPallet", head: "積み込み", body: "荷待ちも管理の外", color: HEX.accent2, fill: "FDF0E7" },
      { icon: "MdLocalShipping", head: "輸送", body: "管理されている", color: HEX.accent1 },
      { icon: "MdSwapHoriz", head: "積み替え", body: "荷降ろしも管理の外", color: HEX.accent2, fill: "FDF0E7" },
      { icon: "MdStore", head: "納品先", body: "受け取りを待つ間は?", color: HEX.accent3 },
    ], 1.2);
    card(s, 0.5, 3.45, 9.0, 1.25, HEX.lt2);
    tx(s, [T("見落としやすいのは、保管と輸送の「間」の時間です。", { bold: true, breakLine: true }), T("出荷を待つ置き場、トラックの荷待ち、中継地、納品先での受け取り待ち。ここで温度や湿度が外れやすくなります。")], 0.75, 3.55, 8.5, 1.05, { fontSize: 14, color: C.text2 });
    s.addNotes("第4回は、荷役と輸配送の温度と湿度の管理です。この回では、製品を動かすときに、温度や湿度が外れやすい場面と、その防ぎ方を学びます。すき間の時間、夏の車内、車両の選び方、積み方と衝撃、そして電池の輸送のルールの順に見ていきます。倉庫で保管している間と、トラックで運んでいる間は、温度や湿度が管理されています。ところが、その間には、積み込み、荷待ち、積み替え、荷降ろしといった荷役の時間があります。この、すき間の時間に、製品は温度管理の外に出てしまいます。また、納品先で受け取りを待っている間も、温度が管理されていないかもしれません。出荷を待つ置き場、トラックの荷待ち、中継地、納品先での受け取り待ち。こうした保管と輸送の間の時間こそ、温度や湿度が外れやすい場所です。この時間は、倉庫の担当者と運送会社の担当者の、ちょうど境目にあたるため、だれの責任なのかがあいまいになりがちです。だれが、どこで、どのくらいの時間まで管理するのかを、あらかじめ決めておくことが大切です。");
  }
  {
    const s = content(sec, L, "夏の車内は、思った以上に高温になる");
    card(s, 0.5, 1.2, 4.35, 2.8, "FDF0E7");
    tx(s, "気温35℃の日、炎天下に4時間駐車した車内", 0.7, 1.32, 4.0, 0.7, { fontSize: 14, bold: true, color: C.text2 });
    tx(s, "52〜57℃", 0.7, 2.0, 4.0, 0.9, { fontSize: 40, bold: true, color: C.accent2, align: "center", valign: "middle" });
    tx(s, "(車内の最高温度。JAFの試験)", 0.7, 2.95, 4.0, 0.4, { fontSize: 12, align: "center", color: C.text2 });
    await circleIcon(s, "MdWarning", 5.3, 1.3, 0.7, HEX.accent4);
    tx(s, "物流での注意", 6.15, 1.35, 3.2, 0.6, { fontSize: 17, bold: true, color: C.text2, valign: "middle" });
    tx(s, [T("温度管理のできない車両の荷台や、日の当たる置き場に放置しない", { bullet: true, breakLine: true }), T("荷待ちや休憩で止まっている間の温度にも注意", { bullet: true, breakLine: true }), T("電池は特に高温を避ける", { bullet: true })], 5.3, 2.2, 4.2, 2.4, { fontSize: 14, paraSpaceAfter: 6 });
    tx(s, "※ 乗用車での試験の例。トラックの荷台の温度は車両や条件によって違います", 0.5, 4.15, 9.0, 0.4, { fontSize: 12, color: C.text2 });
    source(s, "出典: JAFユーザーテスト「車内の環境」(JAF Mate 2025年8月8日の記事で紹介。記事名は参考文献に記載)");
    s.addNotes("夏の車内の温度は、思った以上に高くなります。日本自動車連盟、ジャフの試験では、気温35度の日に、炎天下に4時間駐車した乗用車の車内の最高温度は、52度から57度に達しました。これは乗用車での試験の例で、トラックの荷台の温度は、車両や条件によって違います。それでも、温度管理のできない車両の荷台や、日の当たる置き場に製品を放置すると、高温になるおそれがあることがわかります。荷待ちや休憩で止まっている間の温度にも注意してください。特に電池は、高温を避けることが大切です。荷待ちが長くなりそうなときは、置き場所や待ち方を、前もって運送会社や倉庫と相談しておきましょう。");
  }
  {
    const s = content(sec, L, "温度を保つ車両を確かめる");
    await iconCards(s, [
      { icon: "MdAcUnit", head: "冷蔵冷凍車", body: "外の気温にかかわらず、荷台の中を冷やすことができる装置を持つ車両", color: HEX.accent1 },
      { icon: "MdDeviceThermostat", head: "保温車", body: "外の気温にかかわらず、荷台の中を一定の温度に保つ装置を持つ車両", color: HEX.accent3 },
    ], 1.15, 1.85);
    card(s, 0.5, 3.2, 9.0, 1.5, HEX.dk2);
    tx(s, [T("「保冷車」「定温車」などの呼び方は会社によって違う。手配するときは次を確かめる", { bold: true, breakLine: true, color: C.accent2 }), T("何℃から何℃に保てるか / 冷やす・温める装置があるか(断熱だけか) / 温度を記録できるか", { breakLine: false })], 0.75, 3.3, 8.5, 1.3, { fontSize: 14, color: C.background1 });
    source(s, "出典: 国土交通省 自動車の用途区分の資料(車体の形状「冷蔵冷凍車」「保温車」)");
    s.addNotes("温度を保って運ぶには、それに合った車両が必要です。国土交通省の車両の区分では、外の気温にかかわらず荷台の中を冷やすことのできる装置を持つ車両を、冷蔵冷凍車と呼びます。また、荷台の中を一定の温度に保つ装置を持つ車両を、保温車と呼びます。ただし、現場では、保冷車や定温車など、さまざまな呼び方が使われていて、会社によって意味が違うことがあります。ですから、車両を手配するときは、名前ではなく中身を確かめてください。何度から何度に保てるのか。冷やしたり温めたりする装置があるのか、それとも断熱だけなのか。そして、温度を記録できるのか、です。また、温度帯の違う荷物と一緒に積む場合は、荷台の中を仕切れるかどうかも確かめておくとよいでしょう。");
  }
  {
    const s = content(sec, L, "積み方・衝撃・振動にも注意");
    await iconCards(s, [
      { icon: "MdBatteryAlert", head: "電池", body: "圧迫・振動・落下などの強い衝撃で内部が傷つくと、発熱・発火につながるおそれ。落とさない、つぶさない、無理に積み重ねない", color: HEX.accent2, fill: "FDF0E7" },
      { icon: "MdMemory", head: "電子部品", body: "衝撃に加え、温度差による結露にも注意。防湿袋や帯電防止の包装を傷つけない", color: HEX.accent1 },
    ], 1.15, 2.45);
    card(s, 0.5, 3.8, 9.0, 0.9, HEX.lt2);
    tx(s, "積み方の決まり(向き、段数、固定の方法)は、メーカーや出荷元に確かめ、作業手順に書いておく", 0.75, 3.88, 8.5, 0.75, { fontSize: 14, bold: true, color: C.text2, valign: "middle" });
    source(s, "出典: 製品評価技術基盤機構(NITE)「リチウムイオン電池関連製品の製造と安全性」、NITE ニュースリリース(2025年6月26日)");
    s.addNotes("荷役と輸送では、温度だけでなく、積み方や衝撃、振動にも注意が必要です。製品評価技術基盤機構、ナイトの資料では、電池について、電極が変形するような圧迫や、振動、衝撃を避けることが挙げられています。強い衝撃で電池の内部が傷つくと、発熱や発火につながるおそれがあるためです。落とさない、つぶさない、無理に積み重ねない、が基本です。電子部品も衝撃に注意が必要で、さらに、温度差による結露にも気をつけます。防湿袋や帯電防止の包装を傷つけないようにしましょう。積み方の決まり、例えば向きや段数、固定の方法は、メーカーや出荷元に確かめて、作業手順に書いておきましょう。手順に書いておけば、担当者が変わっても、同じ注意を引き継ぐことができます。");
  }
  {
    const s = content(sec, L, "電池の輸送は「運び方」と「形」でルールが違う");
    await iconCards(s, [
      { icon: "MdFlight", head: "航空", body: "危険物のルールが特に細かい。充電率などの条件もある", color: HEX.accent3 },
      { icon: "MdDirectionsBoat", head: "海上", body: "海上輸送の危険物のルールに従う", color: HEX.accent3 },
      { icon: "MdLocalShipping", head: "陸上", body: "国内では消防法の運搬のルールなど", color: HEX.accent3 },
    ], 1.15, 1.75);
    card(s, 0.5, 3.1, 9.0, 1.6, HEX.lt2);
    tx(s, [T("電池の「形」でも扱いが変わる", { bold: true, breakLine: true, color: C.accent2 }), T("電池そのものか、製品に組み込まれているか(または同梱か)で区分が違う(例: 航空では国連番号が UN3480 と UN3481 に分かれる)。", { breakLine: true }), T("→ 社内の担当部署・運送会社・航空会社に確かめる", { bold: true })], 0.75, 3.18, 8.5, 1.45, { fontSize: 14 });
    source(s, "参考: 航空会社の危険物の案内、消防庁 報道資料(令和7年5月14日)。規則は毎年のように改訂されるため、最新の内容を確認");
    s.addNotes("電池の輸送には、危険物のルールがあります。そのルールは、運び方によって違います。航空輸送は特にルールが細かく、充電率などの条件もあります。海上輸送には海上輸送の危険物のルールがあり、国内の陸上輸送では、消防法の運搬のルールなどが関わります。さらに、電池の形によっても扱いが変わります。電池そのものなのか、製品に組み込まれているのか、または製品と一緒に梱包されているのかで、区分が違います。例えば航空輸送では、国連番号という危険物の番号が、UN3480とUN3481に分かれます。こうしたルールは毎年のように改訂されます。自分で判断せず、社内の担当部署や運送会社、航空会社に確かめてください。");
  }
  await exerciseSlide(sec, L, "第4回のまとめ", ["保管と輸送の「間」の時間に、温度・湿度が外れやすい", "夏の車内や置き場は高温になる", "車両は名前ではなく中身(温度・装置・記録)で確かめる", "電池は衝撃を避け、輸送のルールは担当部署に確かめる"], "担当する製品の輸送ルートを書き出し、温度管理が途切れる場所に印をつけましょう。\n・出荷を待つ置き場\n・荷待ち、中継地\n・納品先での受け取り待ち",
    "第4回のまとめです。温度や湿度が外れやすいのは、保管と輸送の間の、すき間の時間でした。夏の車内や置き場は、思った以上に高温になります。車両は、名前ではなく、保てる温度、装置、記録ができるかという中身で確かめます。そして電池は衝撃を避け、輸送のルールは担当部署に確かめます。演習です。動画を止めて、担当する製品の輸送ルートを書き出し、温度管理が途切れる場所に印をつけてみてください。出荷を待つ置き場、荷待ちや中継地、納品先での受け取り待ちに注目しましょう。印をつけた場所ごとに、だれが管理しているのかも書き添えておくと、関係者に相談するときに役立ちます。");

  // =============== 第5回 ===============
  sec = "第5回 測る・記録する・外れたときに動く"; L = "第5回";
  pres.addSection({ title: sec });
  {
    const s = content(sec, L, "温度と湿度を測る ― データロガー");
    const pts = [["MdTimer", "記録の間隔", "何分ごとに記録するか。短すぎると電池やデータ量の負担、長すぎると変化を見逃す"], ["MdPlace", "置く場所", "温度マッピングの結果をもとに、外れやすい場所に置く"], ["MdTune", "校正", "計器のずれを、決めた間隔で確かめる"], ["MdWaterDrop", "測るもの", "温度だけか、湿度もか。電子部品では湿度も重要"]];
    for (let i = 0; i < 4; i++) {
      const x = 0.5 + (i % 2) * 4.65, y = 1.15 + Math.floor(i / 2) * 1.75;
      card(s, x, y, 4.35, 1.55, HEX.lt2);
      await circleIcon(s, pts[i][0], x + 0.22, y + 0.22, 0.6, HEX.accent1);
      tx(s, pts[i][1], x + 1.0, y + 0.22, 3.2, 0.45, { fontSize: 16, bold: true, color: C.text2, valign: "middle" });
      tx(s, pts[i][2], x + 1.0, y + 0.7, 3.2, 0.8, { fontSize: 12.5 });
    }
    tx(s, "※ 特定の製品はおすすめしていません。必要な機能を決めてから選びましょう", 0.5, 4.65, 9.0, 0.3, { fontSize: 11.5, color: C.text2 });
    source(s, "参考: 厚生労働省「医薬品の適正流通(GDP)ガイドライン」3.3・3.4(温度マッピング、機器の校正、警報)");
    s.addNotes("第5回は、測る、記録する、外れたときに動く、です。この回が最後の回になります。温度や湿度をどう測り、記録をどう使い、条件から外れたときにどう動けばよいかを学びます。温度や湿度を自動で記録する機器を、データロガーといいます。選ぶときの観点は4つです。1つ目は記録の間隔です。何分ごとに記録するかを決めます。短すぎると電池やデータ量の負担が増え、長すぎると変化を見逃します。2つ目は置く場所です。温度マッピングの結果をもとに、外れやすい場所に置きます。3つ目は校正で、計器のずれを、決めた間隔で確かめます。4つ目は何を測るかです。電子部品を扱うなら、温度だけでなく湿度も重要です。なお、この講座では特定の製品はおすすめしていません。必要な機能を決めてから選びましょう。");
  }
  {
    const s = content(sec, L, "記録は「取る」だけでなく「使う」もの");
    await flow(s, [
      { icon: "MdSensors", head: "記録する", body: "決めた間隔で自動記録" },
      { icon: "MdSearch", head: "確かめる", body: "だれが、いつ、何を見るか" },
      { icon: "MdNotificationsActive", head: "気づく", body: "外れたら警報・連絡" },
      { icon: "MdShare", head: "伝える", body: "取引先に記録を渡す" },
    ], 1.2);
    card(s, 0.5, 3.45, 9.0, 1.25, "FDF0E7");
    tx(s, [T("よくある失敗: 記録はあるのに、だれも見ていない", { bold: true, breakLine: true, color: C.accent2 }), T("「だれが・いつ・何を確認するか」と「外れたときの連絡先」を決めておくことで、記録が役に立ちます。")], 0.75, 3.55, 8.5, 1.05, { fontSize: 14 });
    s.addNotes("温度や湿度の記録は、取るだけでは意味がありません。使って初めて役に立ちます。まず、決めた間隔で自動的に記録します。次に、だれが、いつ、何を確かめるのかを決めておきます。そして、条件から外れたら、警報や連絡で気づけるようにします。さらに、取引先から求められたときに、記録を渡せるようにしておきます。よくある失敗は、記録はあるのに、だれも見ていないことです。問題が起きてから記録を見返して、初めて外れていたことに気づく、ということがないよう、確認する人と時期、外れたときの連絡先を、あらかじめ決めておきましょう。また、記録をいつまで残しておくかも、取引先との取り決めや社内のルールで決めておくと安心です。");
  }
  {
    const s = content(sec, L, "条件から外れたら、どう動くか");
    await flow(s, [
      { icon: "MdCampaign", head: "① 報告", body: "決められた連絡先へ、すぐに", color: HEX.accent2 },
      { icon: "MdBlock", head: "② 分ける", body: "その製品を出荷せず、分けて保管", color: HEX.accent2 },
      { icon: "MdFactCheck", head: "③ 判断", body: "使えるかを判断するのはメーカー・品質保証部門", color: HEX.accent3 },
      { icon: "MdEditNote", head: "④ 記録", body: "いつ・どこで・何℃・何時間", color: HEX.accent1 },
      { icon: "MdBuild", head: "⑤ 再発防止", body: "原因を調べ、手順を直す", color: HEX.accent1 },
    ], 1.2);
    card(s, 0.5, 3.45, 9.0, 1.25, HEX.dk2);
    tx(s, [T("「少しだけだから大丈夫」と自分で判断しない", { bold: true, breakLine: true, color: C.accent2 }), T("使えるかどうかの最終判断は、メーカーや品質保証部門の役割です。物流の担当者は、すぐに報告し、正確に記録することが大切です。")], 0.75, 3.55, 8.5, 1.05, { fontSize: 14, color: C.background1 });
    source(s, "参考: 厚生労働省「医薬品の適正流通(GDP)ガイドライン」9.2.2(輸送中の温度逸脱の報告と調査)");
    s.addNotes("温度や湿度が、条件から外れてしまったら、どう動けばよいでしょうか。基本は5つのステップです。1つ目は報告です。決められた連絡先に、すぐに知らせます。2つ目は分けることです。その製品を出荷せず、ほかの製品と分けて保管します。3つ目は判断です。その製品が使えるかどうかを判断するのは、メーカーや品質保証部門です。4つ目は記録です。いつ、どこで、何度に、何時間さらされたのかを記録します。5つ目は再発防止で、原因を調べて手順を直します。大切なのは、少しだけだから大丈夫と、自分で判断しないことです。物流の担当者の役割は、すぐに報告し、正確に記録することです。外れたことを隠したり、報告が遅れたりすると、問題のある製品が取引先に届いてしまうおそれがあります。すぐに報告できる雰囲気をつくることも、管理する立場の大切な役割です。");
  }
  {
    const s = content(sec, L, "電池の異常に気づいたら");
    await iconCards(s, [
      { icon: "MdWhatshot", head: "熱い", body: "触れなくても熱を感じる", color: HEX.accent4, fill: "FDECEA" },
      { icon: "MdOpenInFull", head: "ふくらんでいる", body: "電池や製品の膨張・変形", color: HEX.accent4, fill: "FDECEA" },
      { icon: "MdAir", head: "におう・煙が出る", body: "いつもと違うにおい、煙", color: HEX.accent4, fill: "FDECEA" },
    ], 1.15, 1.85);
    card(s, 0.5, 3.2, 9.0, 1.5, HEX.dk2);
    tx(s, [T("安全を最優先に", { bold: true, breakLine: true, color: C.accent2 }), T("近づきすぎず、すぐに責任者に知らせ、社内の緊急時の手順に従う。火災になったら119番通報", { bullet: true, breakLine: true }), T("緊急時の手順と連絡先を、前もって確かめておく", { bullet: true })], 0.75, 3.28, 8.5, 1.35, { fontSize: 14, color: C.background1 });
    source(s, "出典: 消防庁「リチウムイオン電池対策の取組状況」(膨張などの異常時の使用中止の呼びかけ)、NITE ニュースリリース(2025年6月26日)");
    s.addNotes("電池を扱うときは、異常のサインにも気をつけてください。電池や製品が熱い、ふくらんでいる、いつもと違うにおいや煙がある。こうしたときは、異常のおそれがあります。消防庁も、電池の膨張などの異常が生じた場合には使用を中止するよう呼びかけています。異常に気づいたら、安全を最優先にします。近づきすぎず、すぐに責任者に知らせて、社内の緊急時の手順に従ってください。火災になったら、119番に通報します。そして何より大切なのは、こうした緊急時の手順と連絡先を、前もって確かめておくことです。異常は、いつ、どこで見つかるかわかりません。倉庫の作業者や運送会社のドライバーにも、異常のサインと連絡先を伝えておきましょう。");
  }
  {
    const s = content(sec, L, "講座のまとめ ― 温湿度管理の4つのステップ");
    await flow(s, [
      { icon: "MdRule", head: "条件を確かめる", body: "数字で、文書で、根拠も" },
      { icon: "MdTravelExplore", head: "外れる場所を探す", body: "保管・荷役・輸送の「すき間」" },
      { icon: "MdSensors", head: "測って記録する", body: "置く場所、間隔、校正" },
      { icon: "MdDirectionsRun", head: "外れたら動く", body: "報告・分ける・記録" },
    ], 1.2);
    card(s, 0.5, 3.45, 9.0, 1.25, HEX.lt2);
    tx(s, [T("次の一歩", { bold: true, breakLine: true, color: C.accent1 }), T("担当する製品の条件を確認シートにまとめ、上司や関係者と共有しましょう。応用編(完全版)では、急に温度管理が必要になったときの対応を、ケーススタディで学ぶ予定です。")], 0.75, 3.55, 8.5, 1.05, { fontSize: 14 });
    s.addNotes("最後に、講座全体をまとめます。温度と湿度の管理は、4つのステップで考えます。1つ目は、条件を確かめること。数字で、文書で、そして根拠も確かめます。2つ目は、外れやすい場所を探すこと。保管、荷役、輸送のすき間の時間に注目します。3つ目は、測って記録すること。置く場所、記録の間隔、校正を決めます。4つ目は、外れたら動くこと。報告し、分けて保管し、記録します。この4つは、どんな製品にも共通する考え方です。担当する製品が変わっても、この順番で考えれば、何を確かめればよいかが見えてきます。次の一歩として、担当する製品の条件を、確認シートにまとめて、上司や関係者と共有してみてください。今後予定している応用編では、急に温度管理が必要になったときの対応を、ケーススタディで学ぶ予定です。ご受講ありがとうございました。");
  }
  {
    const s = content(sec, L, "参考文献");
    const refs = [
      "製品評価技術基盤機構(NITE)ニュースリリース「『夏バテ(夏のバッテリー)』にご注意を」2026年7月15日 https://www.nite.go.jp/data/000161880.pdf / 同「ご用心」2025年6月26日 https://www.nite.go.jp/data/000158238.pdf",
      "製品評価技術基盤機構(NITE)「リチウムイオン電池関連製品の製造と安全性」 https://www.nite.go.jp/data/000101539.pdf",
      "消防庁 報道資料(令和5年12月6日・令和7年5月14日)危険物の規制に関する政令等の一部改正 https://www.fdma.go.jp/",
      "消防庁 検討会資料「現行の消防法令の基準について」/「リチウムイオン電池対策の取組状況」 https://www.fdma.go.jp/",
      "厚生労働省「医薬品の適正流通(GDP)ガイドライン」平成30年12月 https://www.mhlw.go.jp/content/11120000/000466214.pdf",
      "第十九改正日本薬局方 通則(令和8年4月10日 厚生労働省告示第193号) https://jpdb.nihs.go.jp/kyokuhou/",
      "国土交通省 自動車の用途区分(冷蔵冷凍車・保温車) https://www.mlit.go.jp/jidosha/kensatoroku/kensa/PDF/kubun-3-1-36.pdf",
      "JAF Mate「車内放置でスマホが発火? 猛暑のリチウムイオン電池事故に要注意【2025年版】」2025年8月8日(JAFユーザーテスト「車内の環境」)",
      "熊本県立大学 辻原万規彦「建築環境工学I 第7回 湿度/結露」2023年5月30日 / 村田製作所「静電気の基礎」/ IPC/JEDEC J-STD-033(業界規格)",
    ];
    tx(s, refs.map((r, i) => T(r, { breakLine: i < refs.length - 1 })), 0.5, 1.1, 9.0, 3.9, { fontSize: 10.5, paraSpaceAfter: 4 });
    s.addNotes("こちらは、この講座の作成にあたって参照した資料の一覧です。電池の事故や安全については製品評価技術基盤機構、電池の保管の規制については消防庁、温度の管理の考え方については厚生労働省の医薬品のガイドライン、車両の区分については国土交通省の資料などをもとにしています。法令やルールは変わることがありますので、実際の業務では、最新の情報を確かめてください。この講座が、新しい製品の物流を担当する皆さんの、最初の一歩の助けになれば幸いです。");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})().catch(e => { console.error(e); process.exit(1); });
