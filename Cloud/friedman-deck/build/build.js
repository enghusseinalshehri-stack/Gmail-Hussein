const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { applyTheme } = require("/root/.claude/skills/synced/e2e490b6-6862-4a13-bca0-000ef8bb5452_1adf728e-0325-4668-9271-a0b825074c66/pptx/scripts/apply_theme.js");

const B = __dirname;
const FS = JSON.parse(fs.readFileSync(path.join(B, "formulas.json")));
const OUT = path.join(B, "..", "friedman-test.pptx");

const THEME = {
  name: "Friedman Dark",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "F1F5F9", lt1: "0B1220", dk2: "94A3B8", lt2: "152038",
    accent1: "FBBF24", accent2: "2DD4BF", accent3: "F472B6", accent4: "60A5FA",
    accent5: "A78BFA", accent6: "34D399", hlink: "60A5FA", folHlink: "A78BFA",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.rtlMode = true;
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "اختبار فريدمان لمقارنة أكثر من عينتين مرتبطتين";
const C = pres.SchemeColor;
const R = 12.63, L = 0.7, W = 11.93;

// ---------- layouts ----------
pres.defineSlideMaster({ title: "TITLE", background: { color: C.background1 }, objects: [] });
pres.defineSlideMaster({
  title: "CONTENT", background: { color: C.background1 },
  objects: [
    { ellipse: { x: L, y: 0.62, w: 0.34, h: 0.34, fill: { color: C.background1 }, line: { color: C.accent1, width: 1.5 } } },
    { text: { text: "اختبار فريدمان  |  Friedman Test", options: { x: 8.2, y: 6.95, w: 4.43, h: 0.3, fontSize: 11, color: C.text2, align: "right", rtlMode: true, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: L + 0.6, y: 0.4, w: W - 0.6, h: 1.0, fontSize: 28, bold: true, color: C.text1, align: "right", valign: "middle", rtlMode: true, margin: 0 }, text: "" } },
  ],
  slideNumber: { x: L, y: 6.95, w: 0.8, h: 0.3, fontSize: 11, color: C.text2, align: "left" },
});
pres.defineSlideMaster({
  title: "DIVIDER", background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 4.4, y: 3.05, w: 8.23, h: 1.9, fontSize: 44, bold: true, color: C.text1, align: "right", valign: "top", rtlMode: true, margin: 0 }, text: "" } },
  ],
});

// ---------- animation / helper registry ----------
const anims = {}, trans = {};
let slideNo = 0, uid = 0, cur;
function reg(name, fx, at, dur) { (anims[slideNo] = anims[slideNo] || []).push({ name, fx, at: at || 0, dur: dur || 500 }); }
let lastSec = "المقدمة";
function newSlide(master, section, transition) {
  if (section === "AUTO") section = lastSec;
  slideNo++; trans[slideNo] = transition || "fade";
  cur = pres.addSlide({ masterName: master, sectionTitle: section });
  return cur;
}
const nm = () => "o" + (++uid);
function tx(text, o, fx, at, dur) {
  const name = nm();
  if (Array.isArray(text)) text = text.map((x) => (typeof x === "string" ? { text: x } : x));
  cur.addText(text, Object.assign({ isTextBox: true, rtlMode: true, align: "right", valign: "top", margin: 0, color: C.text1, fontSize: 16, lang: "ar-SA", objectName: name }, o));
  if (fx) reg(name, fx, at, dur);
  return name;
}
function shp(type, o, fx, at, dur) {
  const name = nm();
  cur.addShape(type, Object.assign({ objectName: name }, o));
  if (fx) reg(name, fx, at, dur);
  return name;
}
function card(x, y, w, h, o, fx, at, dur) {
  return shp(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: 0.1, fill: { color: C.background2 }, line: { type: "none" } }, o || {}), fx, at, dur);
}
function img(file, o, fx, at, dur) {
  const name = nm();
  cur.addImage(Object.assign({ path: file, objectName: name }, o));
  if (fx) reg(name, fx, at, dur);
  return name;
}
function fm(key, x, y, w, h, fx, at, o) {
  const s = FS[key]; const k = Math.min(w / s.w, h / s.h, 1.6);
  const iw = s.w * k, ih = s.h * k;
  const al = (o && o.align) || "center";
  const ix = al === "right" ? x + w - iw : al === "left" ? x : x + (w - iw) / 2;
  return img(path.join(B, "f", key + ".png"), { x: ix, y: y + (h - ih) / 2, w: iw, h: ih, altText: "formula " + key }, fx, at);
}
function ttl(t, fx) {
  const name = nm();
  cur.addText(t, { placeholder: "title", lang: "ar-SA", objectName: name });
  reg(name, fx || "fade", 0, 500);
}
const sb = (t) => ({ text: t, options: { subscript: true } });
const bd = (t) => ({ text: t, options: { bold: true } });

const thead = (cols) => cols.map((c) => ({ text: c, options: { bold: true, color: C.text2, fill: { color: C.background1 }, align: "center", fontSize: 13, rtlMode: true } }));
function table(rows, header, x, y, w, colW, rowH, sumRow, fx, at) {
  const name = nm();
  const data = [thead(header)];
  rows.forEach((r, i) => data.push(r.map((v) => ({ text: String(v), options: { align: "center", fontSize: 15, color: C.text1, fill: { color: i % 2 ? C.background1 : C.background2 }, rtlMode: true } }))));
  if (sumRow) data.push(sumRow.map((v) => ({ text: String(v), options: { align: "center", fontSize: 15, bold: true, color: C.background1, fill: { color: C.accent1 }, rtlMode: true } })));
  cur.addTable(data, { x, y, w, colW, rowH, border: { type: "none" }, valign: "middle", objectName: name, margin: [0, 0.05, 0, 0.05] });
  if (fx) reg(name, fx, at || 0, 600);
}
let secNo = 0;
function divider(text) {
  secNo++;
  const num = String(secNo).padStart(2, "0");
  lastSec = `${num} ${text.replace(/\n/g, " ")}`;
  pres.addSection({ title: lastSec });
  newSlide("DIVIDER", `${num} ${text.replace(/\n/g, " ")}`, "push");
  shp(pres.shapes.OVAL, { x: 0.9, y: 2.2, w: 3.0, h: 3.0, fill: { type: "none" }, line: { color: C.accent2, width: 1.5 } }, "zoom", 0, 800);
  shp(pres.shapes.OVAL, { x: 1.25, y: 2.55, w: 2.3, h: 2.3, fill: { type: "none" }, line: { color: C.accent1, width: 1.5 } }, "zoom", 200, 800);
  tx(num, { x: 0.9, y: 2.2, w: 3.0, h: 3.0, fontSize: 72, bold: true, color: C.accent1, align: "center", valign: "middle", rtlMode: false }, "fade", 500);
  tx("القسم " + num, { x: 4.4, y: 2.55, w: 8.23, h: 0.4, fontSize: 16, color: C.accent2, bold: true }, "fade", 700);
  const n = nm();
  cur.addText(text, { placeholder: "title", lang: "ar-SA", objectName: n });
  reg(n, "float", 900, 600);
}
function badge(x, y, d, t, fill, fx, at) {
  shp(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill || C.accent1 }, line: { type: "none" } }, fx, at);
  tx(t, { x, y, w: d, h: d, align: "center", valign: "middle", fontSize: 14, bold: true, color: C.background1, rtlMode: false }, fx, at);
}

// =====================================================
// 1. Title
pres.addSection({ title: "المقدمة" });
newSlide("TITLE", "المقدمة", "fade");
shp(pres.shapes.OVAL, { x: 0.9, y: 0.8, w: 2.8, h: 2.8, fill: { type: "none" }, line: { color: C.accent1, width: 1.5 } }, "zoom", 0, 800);
shp(pres.shapes.OVAL, { x: 1.4, y: 1.3, w: 1.8, h: 1.8, fill: { type: "none" }, line: { color: C.accent2, width: 1.5 } }, "zoom", 250, 800);
shp(pres.shapes.OVAL, { x: 2.05, y: 1.95, w: 0.5, h: 0.5, fill: { color: C.accent1 }, line: { type: "none" } }, "zoom", 500, 600);
tx("FRIEDMAN TEST", { x: 5.0, y: 3.35, w: R - 5.0, h: 0.4, fontSize: 16, color: C.accent2, bold: true, charSpacing: 6, rtlMode: false }, "fade", 900);
tx("اختبار فريدمان لمقارنة\nأكثر من عينتين مرتبطتين", { x: 3.0, y: 3.9, w: R - 3.0, h: 2.2, fontSize: 48, bold: true }, "float", 1100, 700);
tx("إجراء لامعلمي (Nonparametric) للعينات المرتبطة", { x: 3.0, y: 6.25, w: R - 3.0, h: 0.4, fontSize: 18, color: C.text2 }, "fade", 1800);

// 2. divider
divider("تعريف اختبار فريدمان ومفهومه");

// 3. definition cards
newSlide("CONTENT", "AUTO");
ttl("اختبار فريدمان إجراء لامعلمي لمقارنة أكثر من عينتين غير مستقلتين");
{
  const cw = 3.78, xs = [R - cw, R - 2 * cw - 0.295, L];
  const data = [
    ["التعريف", "لامعلمي", C.accent1, 40, "إجراء إحصائي لامعلمي (Nonparametric Statistical Procedure) لمقارنة أكثر من عينتين غير مستقلتين (Dependent) أو مرتبطتين (Related)."],
    ["المقابل المعلمي", "ANOVA", C.accent2, 44, "تحليل التباين للمقاييس المتكررة (Repeated Measures ANOVA)."],
    ["ما يختبره", "θ", C.accent3, 54, "وسيط المجتمعات (Population Medians)، وتُصاغ الفرضيات وفقًا للمجتمع."],
  ];
  data.forEach((d, i) => {
    const t = 300 + i * 450;
    card(xs[i], 1.8, cw, 4.75, {}, "float", t, 600);
    tx(d[0], { x: xs[i] + 0.3, y: 2.1, w: cw - 0.6, h: 0.35, fontSize: 14, bold: true, color: C.text2 }, "fade", t + 150);
    tx(d[1], { x: xs[i] + 0.3, y: 2.55, w: cw - 0.6, h: 1.0, fontSize: d[3], bold: true, color: d[2], rtlMode: d[1] !== "لامعلمي" ? false : true, align: d[1] === "لامعلمي" ? "right" : "right" }, "zoom", t + 250, 600);
    tx(d[4], { x: xs[i] + 0.3, y: 3.85, w: cw - 0.6, h: 2.4, fontSize: 20, color: C.text1 }, "fade", t + 450);
  });
}

// 4. example
newSlide("CONTENT", "AUTO");
ttl("مثال توضيحي للعينات المرتبطة: نسب القبول في الكليات");
tx("قطاع تعليمي يقارن نسب خريجي المدارس المقبولين في الكليات خلال السنوات الخمس الماضية.", { x: L, y: 1.5, w: W, h: 0.5, fontSize: 18, color: C.text2 }, "fade", 200);
{
  const tw = 3.7, g = 0.415, xs = [R - tw, R - 2 * tw - g, L];
  [["10", "مدارس", C.accent1], ["5", "مجموعات (سنوات)", C.accent2], ["10", "قيم في كل مجموعة", C.accent3]].forEach((d, i) => {
    const t = 500 + i * 400;
    card(xs[i], 2.4, tw, 2.3, {}, "fade", t, 500);
    tx(d[0], { x: xs[i], y: 2.6, w: tw, h: 1.3, fontSize: 72, bold: true, color: d[2], align: "center", valign: "middle", rtlMode: false }, "zoom", t + 200, 600);
    tx(d[1], { x: xs[i], y: 3.95, w: tw, h: 0.5, fontSize: 18, align: "center" }, "fade", t + 400);
  });
}
card(L, 5.1, W, 1.3, { fill: { color: C.background2 } }, "float", 2000, 600);
shp(pres.shapes.OVAL, { x: R - 0.8, y: 5.55, w: 0.4, h: 0.4, fill: { color: C.accent1 }, line: { type: "none" } }, "zoom", 2200, 400);
tx("العينات مرتبطة لأن لكل مدرسة نسبة سنوية تتكرر عبر السنوات", { x: L + 0.3, y: 5.1, w: W - 1.4, h: 1.3, fontSize: 22, bold: true, valign: "middle" }, "fade", 2300);

// 5. what it does not tell
newSlide("CONTENT", "AUTO");
ttl("ما الذي لا يخبرنا به اختبار فريدمان؟");
card(7.03, 1.7, 5.6, 4.9, {}, "float", 300, 600);
tx("النتيجة المعنوية", { x: 7.33, y: 2.0, w: 5.0, h: 0.4, fontSize: 14, bold: true, color: C.accent2 }, "fade", 600);
tx("النتيجة المعنوية تعني أن عينة واحدة على الأقل مختلفة عن العينات الأخرى.", { x: 7.33, y: 2.6, w: 5.0, h: 3.6, fontSize: 28, bold: true }, "fade", 800, 600);
tx("لا يخبرنا عن:", { x: L, y: 1.7, w: 6.0, h: 0.35, fontSize: 14, bold: true, color: C.text2 }, "fade", 1100);
["مكان الاختلاف", "عدد الاختلافات"].forEach((t, i) => {
  card(L, 2.15 + i * 0.85, 6.0, 0.7, { fill: { color: C.background2 } }, "wipeR", 1200 + i * 350, 500);
  tx(t, { x: L + 0.3, y: 2.15 + i * 0.85, w: 5.4, h: 0.7, fontSize: 20, strike: true, color: C.accent3, valign: "middle" }, "wipeR", 1200 + i * 350, 500);
});
card(L, 3.95, 6.0, 2.65, { fill: { color: C.background2 } }, "float", 2000, 600);
tx("لتحديدها نستخدم اختلافات العينات أو الاختبارات البعدية (Post Hoc Tests)، ومنها اختبار ويلكوكسن لإشارات الرتب (Wilcoxon Signed Rank Test).", { x: L + 0.3, y: 4.15, w: 5.4, h: 2.3, fontSize: 22 }, "fade", 2300);

// 6. divider
divider("خطوات إجراء الاختبار");

// 7. steps
newSlide("CONTENT", "AUTO");
ttl("حساب إحصاء اختبار فريدمان");
{
  const steps = [
    ["إنشاء جدول البيانات: الأفراد في الصفوف، والحالات في الأعمدة."],
    ["ترتيب قيم الحالات لكل فرد على حدة (أفقيًا داخل الصف)."],
    [{ text: "جمع الرتب لكل حالة للحصول على " }, { text: "R" }, sb("i"), { text: "." }],
    ["اختيار الصيغة: بلا رتب متساوية، أو مع رتب متساوية (Ties)."],
    ["تحديد درجة الحرية والقيمة الحرجة ثم المقارنة واتخاذ القرار."],
  ];
  steps.forEach((s, i) => {
    const y = 1.65 + i * 0.99, t = 300 + i * 450;
    card(5.9, y, 6.73, 0.8, {}, "wipeR", t, 500);
    badge(R - 0.65, y + 0.15, 0.5, String(i + 1).padStart(2, "0"), C.accent1, "fade", t + 200);
    tx(s, { x: i === 4 ? 8.1 : 6.1, y, w: i === 4 ? 3.7 : 5.7, h: 0.8, fontSize: 16, valign: "middle" }, "fade", t + 200);
  });
  fm("dfs", 6.1, 1.65 + 4 * 0.99, 1.7, 0.8, "zoom", 2300);
  card(L, 1.65, 4.9, 2.4, {}, "float", 800, 600);
  tx("بلا رتب متساوية", { x: L + 0.25, y: 1.8, w: 4.4, h: 0.35, fontSize: 14, bold: true, color: C.accent2 }, "fade", 1000);
  fm("f1s", L + 0.25, 2.2, 4.4, 1.7, "wipeL", 1300, { align: "center" });
  card(L, 4.25, 4.9, 2.5, {}, "float", 1400, 600);
  tx("مع رتب متساوية (Ties)", { x: L + 0.25, y: 4.4, w: 4.4, h: 0.35, fontSize: 14, bold: true, color: C.accent2 }, "fade", 1600);
  fm("f2s", L + 0.25, 4.8, 4.4, 1.8, "wipeL", 1900, { align: "center" });
}

// 8. divider
divider("القانون والرموز");

// helper: symbol rows
function symRows(rows, y0, pitch, rh, t0) {
  rows.forEach((r, i) => {
    const y = y0 + i * pitch, t = t0 + i * 350;
    card(L, y, W, rh, {}, "wipeR", t, 450);
    fm(r[0], R - 1.2, y, 1.0, rh, "fade", t + 150);
    tx(r[1], { x: L + 0.3, y, w: W - 1.8, h: rh, fontSize: 16, valign: "middle" }, "fade", t + 150);
  });
}
// 9. formula without ties
newSlide("CONTENT", "AUTO");
ttl("الصيغة: عند عدم وجود رتب متساوية");
card(L, 1.5, W, 2.0, {}, "fade", 200, 500);
fm("f1", L, 1.5, W, 2.0, "wipeL", 500, { align: "center" });
symRows([
  ["sFr", "إحصاء اختبار فريدمان"],
  ["sn", "عدد الصفوف، أي عدد الأفراد أو العناصر"],
  ["sk", "عدد الأعمدة، أي عدد الحالات"],
  ["sRi", "مجموع الرتب للعمود أو الحالة i"],
], 3.75, 0.72, 0.6, 1100);

// 10. formula with ties
newSlide("CONTENT", "AUTO");
ttl("الصيغة: عند وجود رتب متساوية");
card(L, 1.5, W, 2.0, {}, "fade", 200, 500);
fm("f2", 6.9, 1.5, 5.5, 2.0, "wipeL", 500);
fm("f3", 1.0, 1.5, 5.4, 2.0, "wipeL", 900);
symRows([
  ["sCF", "معامل تصحيح تساوي الرتب (Ties Correction)"],
  ["srij", "رتبة القيمة الخاصة بالعنصر j للحالة i"],
  ["sn", "عدد الصفوف، أي عدد الأفراد أو العناصر"],
  ["sk", "عدد الأعمدة، أي عدد الحالات"],
  ["sRi", "مجموع الرتب للعمود أو الحالة i"],
], 3.75, 0.62, 0.5, 1300);

// 11. df
newSlide("CONTENT", "AUTO");
ttl("درجة الحرية والقيم الحرجة");
card(3.9, 1.6, 5.53, 1.4, { fill: { color: C.background2 } }, "zoom", 300, 600);
fm("df", 3.9, 1.6, 5.53, 1.4, "fade", 700);
{
  const cw = 5.82;
  [[R - cw, "العينات الصغيرة", "عند حساب قيمة إحصاء فريدمان فإنه يمكننا مقارنتها مع جدول القيم الحرجة لاختبار المجموعات للفروقات المعنوية.", C.accent1],
   [L, "العينات الكبيرة", "يمكن تقريب العينات كبيرة الحجم باستخدام جدول توزيع كاي تربيع χ².", C.accent2]].forEach((d, i) => {
    const t = 1000 + i * 500;
    card(d[0], 3.3, cw, 3.1, {}, "float", t, 600);
    tx(d[1], { x: d[0] + 0.35, y: 3.6, w: cw - 0.7, h: 0.5, fontSize: 24, bold: true, color: d[3] }, "fade", t + 250);
    tx(d[2], { x: d[0] + 0.35, y: 4.3, w: cw - 0.7, h: 1.9, fontSize: 22 }, "fade", t + 450);
  });
}

// 12. divider
divider("مسألة مختارة\n(العينات الصغيرة بلا رتب متساوية)");

// 13. problem
newSlide("CONTENT", "AUTO");
ttl("المسألة: خصم الراتب لتقليل التأخر (بلا رتب متساوية)");
card(7.6, 1.7, 5.03, 4.4, {}, "float", 300, 600);
tx("مديرة تطبق على 7 موظفين خصم 10 دولارات عن كل يوم تأخر في الشهر الأول، ثم 20 دولارًا في الشهر الثاني؛ خط الأساس = عدد مرات التأخر قبل التطبيق.", { x: 7.9, y: 2.0, w: 4.43, h: 3.8, fontSize: 20 }, "fade", 700);
table([
  [1, 16, 13, 12], [2, 10, 5, 2], [3, 7, 8, 9], [4, 13, 11, 5], [5, 17, 2, 6], [6, 10, 7, 9], [7, 11, 6, 7],
].map((r) => [r[3], r[2], r[1], r[0]]), ["شهر 2", "شهر 1", "خط الأساس", "الموظف"], L, 1.7, 6.5, [1.6, 1.6, 1.6, 1.7], 0.52, null, "fade", 500);

// 14. ranks + hypotheses
newSlide("CONTENT", "AUTO");
ttl("جدول الرتب ومجموعها");
table([
  [1, 3, 2, 1], [2, 3, 2, 1], [3, 1, 2, 3], [4, 3, 2, 1], [5, 3, 1, 2], [6, 3, 1, 2], [7, 3, 1, 2],
].map((r) => [r[3], r[2], r[1], r[0]]), ["شهر 2", "شهر 1", "خط الأساس", "الموظف"], L, 1.6, 6.4, [1.4, 1.4, 1.6, 2.0], 0.5, ["12", "11", "19", "R (مجموع الرتب)"], "fade", 300);
card(7.4, 1.6, 5.23, 2.4, {}, "float", 900, 600);
tx([bd("أولًا: الفرضية الصفرية (Null Hypothesis): "), { text: "تشير إلى أن أيًا من استراتيجيتي المديرة لا تُغيّر عدد مرات تأخر الموظفين." }], { x: 7.65, y: 1.8, w: 4.73, h: 1.4, fontSize: 14 }, "fade", 1200);
fm("h0", 7.65, 3.2, 4.73, 0.6, "wipeL", 1500);
card(7.4, 4.2, 5.23, 2.5, {}, "float", 1700, 600);
tx([bd("ثانيًا: فرضية البحث (Research Hypothesis): "), { text: "تشير إلى أن إحدى الاستراتيجيتين أو كلتيهما ستُخفض عدد مرات تأخر الموظفين شهريًا." }], { x: 7.65, y: 4.4, w: 4.73, h: 1.4, fontSize: 14 }, "fade", 2000);
fm("ha", 11.7, 5.85, 0.68, 0.6, "fade", 2300);
tx("إحدى الاستراتيجيتين أو كلتاهما تخفضان التأخر", { x: 7.65, y: 5.85, w: 3.95, h: 0.6, fontSize: 14, valign: "middle" }, "fade", 2300);

// 15. calculation
newSlide("TITLE", "AUTO", "wipe");
shp(pres.shapes.RECTANGLE, { x: 8.6, y: 0, w: 4.733, h: 7.5, fill: { color: C.accent1 }, line: { type: "none" } }, "wipeR", 0, 700);
tx("التعويض في الصيغة", { x: 1.2, y: 1.15, w: 6.9, h: 0.4, fontSize: 16, bold: true, color: C.accent2, align: "left" }, "fade", 200);
["c1", "c2", "c3", "c4", "c5"].forEach((k, i) => {
  const s = FS[k], sc = Math.min(1, 7.3 / s.w);
  img(path.join(B, "f", k + ".png"), { x: 1.2, y: 1.9 + i * 0.95 + (0.8 - s.h * sc) / 2, w: s.w * sc, h: s.h * sc, altText: "step " + k }, "wipeL", 500 + i * 650, 550);
});
tx("النتيجة", { x: 8.9, y: 2.4, w: 4.1, h: 0.5, fontSize: 22, bold: true, color: C.background1, align: "center" }, "fade", 500);
{
  const t0 = 3900;
  const fin = tx("5.429", { x: 8.9, y: 3.0, w: 4.1, h: 1.6, fontSize: 80, bold: true, color: C.background1, align: "center", valign: "middle", rtlMode: false });
  reg(fin, "zoom", t0, 700);
  tx("إحصاء فريدمان المحسوب", { x: 8.9, y: 4.85, w: 4.1, h: 0.5, fontSize: 18, color: C.background1, align: "center" }, "fade", t0 + 600);
}

// 16. divider
divider("تفسير النتيجة واتخاذ القرار الإحصائي");

// 17. decision rule
newSlide("CONTENT", "AUTO");
ttl("قاعدة القرار: مقارنة القيمة المحسوبة بالقيمة الحرجة");
{
  const cw = 5.82;
  [[R - cw, "إذا كانت القيمة الحرجة ≤ القيمة المحسوبة:", "نرفض الفرضية الصفرية", C.accent3],
   [L, "إذا كانت القيمة الحرجة > القيمة المحسوبة:", "لا نرفض الفرضية الصفرية", C.accent2]].forEach((d, i) => {
    const t = 300 + i * 400;
    card(d[0], 1.6, cw, 1.5, {}, "float", t, 500);
    tx(d[1], { x: d[0] + 0.3, y: 1.8, w: cw - 0.6, h: 0.4, fontSize: 16, color: C.text2 }, "fade", t + 200);
    tx(d[2], { x: d[0] + 0.3, y: 2.3, w: cw - 0.6, h: 0.6, fontSize: 24, bold: true, color: d[3] }, "fade", t + 300);
  });
  const tw = 2.4, xs = [R - tw, R - 2 * tw - 0.3, R - 3 * tw - 0.6];
  [["المسألة", "k=3 ، n=7", 30], ["قيمة Fr المحسوبة", "5.429", 44], ["القيمة الحرجة", "7.14", 44]].forEach((d, i) => {
    const t = 1200 + i * 400;
    card(xs[i], 3.4, tw, 2.0, {}, "float", t, 500);
    tx(d[0], { x: xs[i], y: 3.65, w: tw, h: 0.4, fontSize: 14, color: C.text2, align: "center" }, "fade", t + 200);
    tx(d[1], { x: xs[i], y: 4.2, w: tw, h: 0.9, fontSize: d[2], bold: true, align: "center", valign: "middle", rtlMode: false }, "zoom", t + 300, 500);
  });
  card(L, 3.4, 3.83, 2.0, { fill: { color: C.accent2 } }, "float", 2400, 500);
  tx([bd("لا نرفض الفرضية الصفرية "), { text: "(p>0.05) لأن القيمة الحرجة أكبر من القيمة المحسوبة." }], { x: L + 0.3, y: 3.4, w: 3.23, h: 2.0, fontSize: 16, color: C.background1, valign: "middle" }, "fade", 2700);
  card(L, 5.7, W, 0.95, { fill: { color: C.accent1 } }, "wipeR", 3200, 600);
  tx("البيانات لا تدعم خصم 10 دولارات أو 20 دولارًا.", { x: L, y: 5.7, w: W, h: 0.95, fontSize: 24, bold: true, color: C.background1, align: "center", valign: "middle" }, "fade", 3500);
}

// 18. divider
divider("التطبيق العملي\nباستخدام برنامج SPSS");

// 19. SPSS data
newSlide("CONTENT", "AUTO");
ttl("بيانات المسألة: العينات صغيرة الحجم مع رتب متساوية");
card(L, 1.7, 5.9, 4.9, {}, "float", 300, 600);
tx([
  { text: "بعد فشل المديرة في تخفيض حالات تأخر موظفيها للحضور باستخدام طريقة الخصم من الراتب، قررت تطبيق طريقة أخرى حيث قامت بمكافأة الموظفين الذين لا يتأخرون بالحضور.", options: { breakLine: true, paraSpaceAfter: 10 } },
  { text: "الأولى: تقوم المديرة على مدى شهر بمكافأة الموظفين بمقدار 10 دولارات لكل يوم لا يتأخر فيه الموظفون.", options: { breakLine: true, paraSpaceAfter: 10 } },
  { text: "الثانية: تقوم المديرة وخلال الشهر التالي بمكافأتهم بمقدار 20 دولارًا لكل يوم يصل فيه الموظفون في الوقت المحدد." },
], { x: L + 0.3, y: 1.95, w: 5.3, h: 4.4, fontSize: 18 }, "fade", 700, 600);
table([
  [1, 16, 17, 11], [2, 10, 5, 2], [3, 7, 8, 0], [4, 13, 9, 5], [5, 17, 2, 2], [6, 10, 10, 9], [7, 11, 6, 5],
].map((r) => [r[3], r[2], r[1], r[0]]), ["شهر 2", "شهر 1", "خط الأساس", "الموظف"], 6.9, 1.7, 5.73, [1.4, 1.4, 1.5, 1.43], 0.55, null, "fade", 500);

// 20. hypotheses in SPSS
newSlide("CONTENT", "AUTO");
ttl("الفرضيتان في تطبيق SPSS");
card(L, 1.6, W, 2.45, {}, "float", 300, 600);
tx([bd("الفرضية الصفرية (Null Hypothesis): "), { text: "تشير إلى أن أيًا من الاستراتيجيتين لا تُغيّر عدد مرات تأخر الموظفين." }], { x: L + 0.35, y: 1.85, w: W - 0.7, h: 1.0, fontSize: 20 }, "fade", 700);
fm("h0", L + 0.35, 3.0, W - 0.7, 0.8, "wipeL", 1000);
card(L, 4.3, W, 2.4, {}, "float", 1300, 600);
tx([bd("فرضية البحث (Research Hypothesis): "), { text: "تشير إلى أن إحدى الاستراتيجيتين أو كلتيهما ستخفض عدد مرات تأخر الموظفين." }], { x: L + 0.35, y: 4.5, w: W - 0.7, h: 1.0, fontSize: 20 }, "fade", 1700);
fm("ha", R - 0.35 - 0.7, 5.6, 0.7, 0.8, "fade", 2000);
tx("إحدى الاستراتيجيتين أو كلتاهما ستخفض عدد مرات تأخر الموظفين شهريًا.", { x: L + 0.35, y: 5.6, w: W - 1.6, h: 0.8, fontSize: 18, valign: "middle" }, "fade", 2000);

// 21. variables + data entry
newSlide("CONTENT", "AUTO");
ttl("تعريف المتغيرات وإدخال البيانات");
[[6.83, "01", "p-000.jpg", 'من "Variable View" اكتب أسماء المتغيرات في عمود "Name": Baseline و Month_1 و Month_2'],
 [L, "02", "p-001.jpg", 'من "Data View" أدخل البيانات أسفل أسماء المتغيرات.']].forEach((d, i) => {
  const t = 300 + i * 700;
  card(d[0], 1.6, 5.8, 5.1, {}, "float", t, 600);
  badge(d[0] + 5.8 - 0.75, 1.8, 0.45, d[1], C.accent1, "zoom", t + 200);
  img(path.join(B, d[2]), { x: d[0] + 0.5, y: 2.35, w: 4.8, h: 3.6, altText: "SPSS " + d[1] }, "fade", t + 300, 600);
  tx(d[3], { x: d[0] + 0.3, y: 6.05, w: 5.2, h: 0.6, fontSize: 14 }, "fade", t + 500);
});

// 22. analysis path
newSlide("CONTENT", "AUTO");
ttl("تنفيذ التحليل");
{
  const chips = [["Analyze", 1.8], ["Nonparametric Tests", 3.0], ["Legacy Dialogs", 2.4], ["K Related Samples", 2.9]];
  let x = R;
  chips.forEach((c, i) => {
    x -= c[1];
    const t = 200 + i * 350;
    card(x, 1.55, c[1], 0.6, { fill: { color: C.background2 } }, "wipeR", t, 400);
    tx(c[0], { x, y: 1.55, w: c[1], h: 0.6, fontSize: 15, bold: true, color: C.accent1, align: "center", valign: "middle", rtlMode: false }, "wipeR", t, 400);
    if (i < 3) { x -= 0.6; tx("←", { x, y: 1.55, w: 0.6, h: 0.6, fontSize: 22, bold: true, color: C.accent2, align: "center", valign: "middle", rtlMode: false }, "fade", t + 250); }
  });
  const cw = 3.78, xs = [8.85, 4.775, 0.7];
  const step = [
    'اختر Analyze ثم Nonparametric Tests ثم Legacy Dialogs ثم K Related Samples.',
    'انقل المتغيرات بزر السهم إلى صندوق "Test Variables".',
    'اختبار "Friedman" مختار افتراضيًا (As Default)؛ اضغط "OK".',
  ];
  step.forEach((s, i) => {
    const t = 1500 + i * 600;
    card(xs[i], 2.45, cw, 4.3, {}, "float", t, 600);
    badge(xs[i] + cw - 0.65, 2.65, 0.45, "0" + (i + 1), C.accent1, "zoom", t + 200);
    tx(s, { x: xs[i] + 0.25, y: 2.65, w: cw - 1.1, h: 1.1, fontSize: 14, valign: "top" }, "fade", t + 250);
  });
  img(path.join(B, "p-002.jpg"), { x: xs[0] + 0.19, y: 3.95, w: 3.4, h: 2.55, altText: "SPSS menu" }, "fade", 1900, 600);
  img(path.join(B, "p-003.jpg"), { x: xs[1] + 0.19, y: 4.15, w: 3.4, h: 1.93, altText: "SPSS dialog" }, "fade", 2500, 600);
  card(xs[2] + 0.4, 4.1, cw - 0.8, 0.75, { fill: { color: C.background1 } }, "wipeR", 3100, 450);
  tx("Friedman  ✓", { x: xs[2] + 0.4, y: 4.1, w: cw - 0.8, h: 0.75, fontSize: 20, bold: true, color: C.accent2, align: "center", valign: "middle", rtlMode: false }, "wipeR", 3100, 450);
  card(xs[2] + 0.4, 5.2, cw - 0.8, 0.75, { fill: { color: C.accent1 } }, "zoom", 3500, 450);
  tx("OK", { x: xs[2] + 0.4, y: 5.2, w: cw - 0.8, h: 0.75, fontSize: 22, bold: true, color: C.background1, align: "center", valign: "middle", rtlMode: false }, "zoom", 3500, 450);
}

// 23. divider
divider("تفسير جداول مخرجات SPSS");

// 24. outputs
newSlide("CONTENT", "AUTO");
ttl("مخرجات SPSS: جدولا Ranks و Test Statistics");
img(path.join(B, "p-004.jpg"), { x: L, y: 1.7, w: 5.6, h: 4.2, altText: "SPSS output" }, "fade", 300, 700);
tx("متوسط الرتب لكل حالة", { x: 6.8, y: 1.7, w: 5.83, h: 0.35, fontSize: 14, bold: true, color: C.text2 }, "fade", 700);
[["2.64", "Baseline", C.accent1], ["2.29", "Month_1", C.accent2], ["1.07", "Month_2", C.accent3]].forEach((d, i) => {
  const x = R - 1.8 - i * 2.015, t = 900 + i * 350;
  card(x, 2.15, 1.8, 1.3, {}, "float", t, 500);
  tx(d[0], { x, y: 2.3, w: 1.8, h: 0.7, fontSize: 32, bold: true, color: d[2], align: "center", valign: "middle", rtlMode: false }, "zoom", t + 200, 500);
  tx(d[1], { x, y: 3.0, w: 1.8, h: 0.3, fontSize: 13, color: C.text2, align: "center", rtlMode: false }, "fade", t + 300);
});
card(6.8, 3.65, 5.83, 1.0, {}, "float", 2100, 500);
tx("قيمة إحصاء فريدمان المحسوبة", { x: 7.8, y: 3.65, w: 4.55, h: 1.0, fontSize: 17, valign: "middle" }, "fade", 2300);
tx("10.231", { x: 7.0, y: 3.65, w: 2.2, h: 1.0, fontSize: 30, bold: true, color: C.accent2, align: "left", valign: "middle", rtlMode: false }, "zoom", 2400, 500);
card(6.8, 4.85, 5.83, 1.85, {}, "float", 2700, 500);
tx("وحيث إن هذا الاختبار يستخدم توزيع مربع كاي فإن اختبار فريدمان يسمى chi-square وعدد العناصر 7 ودرجة الحرية df = 2 والدلالة Asymp. Sig = .006.", { x: 7.05, y: 5.0, w: 5.33, h: 1.55, fontSize: 17, valign: "middle" }, "fade", 3000);

// 25. reading
newSlide("CONTENT", "AUTO");
ttl("قراءة المخرجات واتخاذ القرار");
{
  const tw = 2.8, g = (W - 4 * tw) / 3;
  [["N", "7", C.text1], ["DF", "2", C.text1], ["CHI-SQUARE", "10.231", C.text1], ["P", "0.006", C.background1]].forEach((d, i) => {
    const x = R - tw - i * (tw + g), t = 300 + i * 450, hi = i === 3;
    card(x, 1.7, tw, 2.2, { fill: { color: hi ? C.accent1 : C.background2 } }, "float", t, 500);
    tx(d[1], { x, y: 1.95, w: tw, h: 1.2, fontSize: 52, bold: true, color: hi ? C.background1 : C.accent2, align: "center", valign: "middle", rtlMode: false }, "zoom", t + 200, 600);
    tx(d[0], { x, y: 3.2, w: tw, h: 0.4, fontSize: 14, bold: true, color: hi ? C.background1 : C.text2, align: "center", rtlMode: false, charSpacing: 3 }, "fade", t + 300);
  });
}
card(L, 4.2, W, 2.5, {}, "float", 2300, 600);
tx([{ text: "نتيجة الاختبار معنوية: (F" }, sb("r"), { text: " = 10.23، p < 0.05)." }], { x: L + 0.4, y: 4.45, w: W - 0.8, h: 0.6, fontSize: 24, bold: true, color: C.accent1 }, "fade", 2700);
tx("وفقًا للنتائج نرفض الفرضية الصفرية، أي يوجد فرق معنوي بين حالتين أو أكثر من الحالات الثلاث، ولمقارنة الحالات زوجيًا يمكن استخدام اختلافات العينات.", { x: L + 0.4, y: 5.2, w: W - 0.8, h: 1.3, fontSize: 20 }, "fade", 3100);

// 26. reference
newSlide("CONTENT", "AUTO");
ttl("المرجع");
shp(pres.shapes.OVAL, { x: 0.9, y: 2.0, w: 3.2, h: 3.2, fill: { type: "none" }, line: { color: C.accent1, width: 1.5 } }, "zoom", 200, 800);
shp(pres.shapes.OVAL, { x: 1.4, y: 2.5, w: 2.2, h: 2.2, fill: { type: "none" }, line: { color: C.accent2, width: 1.5 } }, "zoom", 450, 800);
card(4.6, 1.9, 8.03, 3.0, {}, "float", 700, 600);
tx("كورد، جريجوري، وفورمان، ديل. (2020). الإحصاء اللامعلمي: خطوة بخطوة (وسيم بن سلمان النصير، مترجم). مركز البحوث والدراسات.", { x: 4.95, y: 2.2, w: 7.33, h: 2.4, fontSize: 22, valign: "middle" }, "fade", 1100);

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  fs.writeFileSync(path.join(B, "anims.json"), JSON.stringify({ anims, trans }));
  console.log("ok", slideNo, "slides");
})();
