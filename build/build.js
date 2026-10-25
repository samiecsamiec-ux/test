// Generátor prezentace „OSA a inzulinová rezistence" – FÁZE 3
// Všechna čísla pocházejí z FAZE1_tabulka_evidence.md (ověřeno v PubMedu 25. 10. 2026).
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { applyTheme } = require("/root/.claude/skills/synced/0d41d7db-9120-4970-bad2-969fb9a805c5_82b02416-39f2-41be-afac-a300b20f2efe/pptx/scripts/apply_theme.js");

const REFS = JSON.parse(fs.readFileSync(path.join(__dirname, "refs.json"), "utf8"));
const refByPmid = Object.fromEntries(REFS.map(r => [r.pmid, r]));

// ---------------- THEME ----------------
const THEME = {
  name: "OSA IR Kardiologie",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B2A41", lt1: "FFFFFF", dk2: "3D5A80", lt2: "EEF2F7",
    accent1: "C8553D", // OSA / hypoxie (terakota)
    accent2: "2A9D8F", // metabolismus (teal)
    accent3: "E9C46A", // zlatá – zvýraznění
    accent4: "8D99AE", // šedá – neutrální / NS
    accent5: "6B4C9A", // fialová – guideline
    accent6: "264653", // tmavý teal
    hlink: "2A9D8F", folHlink: "6B4C9A",
  },
};
const HX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "OSA a inzulinová rezistence: mechanismy, evidence a klinické důsledky pro kardiologa";
pres.author = "Centrum preventivní medicíny";
pres.lang = "cs-CZ";
const C = pres.SchemeColor;
const W = 13.333, H = 7.5;

// ---------------- LAYOUTS ----------------
const FOOT = "OSA a inzulinová rezistence – pohled kardiologa · rešerše PubMed 25. 10. 2026";
pres.defineSlideMaster({
  title: "TITLE", background: { color: HX.dk1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 1.6, w: 11.9, h: 2.2, fontSize: 40, bold: true, color: HX.lt1, align: "left", valign: "bottom", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.7, y: 3.9, w: 11.9, h: 0.9, fontSize: 20, color: HX.lt2, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT", background: { color: HX.lt1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.35, w: 8.9, h: 0.95, fontSize: 25, bold: true, color: HX.dk1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.5, y: 7.05, w: 9.5, h: 0.3, fontSize: 9, color: HX.accent4, align: "left", margin: 0 } } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.6, h: 0.3, fontSize: 9, color: HX.accent4, align: "right" },
});
pres.defineSlideMaster({
  title: "DARK", background: { color: HX.dk1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.35, w: 8.9, h: 0.95, fontSize: 25, bold: true, color: HX.lt1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.5, y: 7.05, w: 9.5, h: 0.3, fontSize: 9, color: HX.accent4, align: "left", margin: 0 } } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.6, h: 0.3, fontSize: 9, color: HX.accent4, align: "right" },
});

// ---------------- HELPERS ----------------
const TAGCOL = {
  "MA RCT": HX.accent2, "RCT": HX.accent2, "MA obs.": HX.dk2, "kohorta": HX.dk2, "průřez": HX.accent4,
  "mechanistické": HX.accent1, "zvířecí": HX.accent1, "MR": HX.accent5, "guideline": HX.accent5, "syntéza": HX.accent6, "placeholder": HX.accent4, "rešerše": HX.accent4,
};
function chip(slide, labels) {
  // štítky úrovně evidence vpravo nahoře
  let x = 12.85;
  const arr = Array.isArray(labels) ? labels : [labels];
  for (let i = arr.length - 1; i >= 0; i--) {
    const lab = arr[i];
    const key = Object.keys(TAGCOL).find(k => lab.startsWith(k)) || "syntéza";
    const w = Math.max(0.95, 0.12 * lab.length + 0.25);
    x -= w;
    slide.addShape(pres.ShapeType.roundRect, { x, y: 0.42, w, h: 0.34, fill: { color: TAGCOL[key] }, line: { color: TAGCOL[key] }, rectRadius: 0.08, objectName: "tag-" + lab });
    slide.addText(lab, { x, y: 0.42, w, h: 0.34, fontSize: 10, bold: true, color: HX.lt1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "tag-text-" + lab });
    x -= 0.08;
  }
}
function cite(slide, text) {
  slide.addText(text, { x: 0.5, y: 6.6, w: 12.3, h: 0.42, fontSize: 9.5, color: HX.accent4, italic: true, align: "left", valign: "bottom", margin: 0, isTextBox: true, objectName: "citace" });
}
function msg(slide, text, y = 1.35, h = 0.55, opts = {}) {
  slide.addText(text, Object.assign({ x: 0.5, y, w: 12.3, h, fontSize: 15, color: HX.dk2, align: "left", valign: "top", margin: 0, isTextBox: true, objectName: "sdeleni" }, opts));
}
function box(slide, x, y, w, h, title, lines, color, opts = {}) {
  slide.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: opts.fill || HX.lt2 }, line: { color: opts.fill || HX.lt2 }, rectRadius: 0.1, objectName: "box-" + title });
  slide.addText(title, { x: x + 0.15, y: y + 0.1, w: w - 0.3, h: 0.36, fontSize: opts.titleSize || 12.5, bold: true, color: color || HX.dk1, margin: 0, isTextBox: true, valign: "top", objectName: "box-title-" + title });
  const items = lines.map((t, i) => ({ text: t, options: { bullet: opts.bullet === false ? false : { indent: 10 }, breakLine: i < lines.length - 1, paraSpaceAfter: 3 } }));
  slide.addText(items, { x: x + 0.15, y: y + 0.48, w: w - 0.3, h: h - 0.58, fontSize: opts.fontSize || 11, color: opts.textColor || HX.dk1, margin: 0, isTextBox: true, valign: "top", objectName: "box-body-" + title });
}
function stat(slide, x, y, w, big, small, color) {
  slide.addText(big, { x, y, w, h: 0.9, fontSize: 40, bold: true, color, align: "left", valign: "bottom", margin: 0, isTextBox: true, objectName: "stat-" + big });
  slide.addText(small, { x, y: y + 0.92, w, h: 0.75, fontSize: 12, color: HX.lt2, align: "left", valign: "top", margin: 0, isTextBox: true, objectName: "stat-l-" + big });
}

// Forest plot kreslený z tvarů (bodový odhad + 95 % CI). rows: {label, est, lo, hi, color, val, dim}
function forest(slide, o) {
  const { x, y, w, h, rows, xmin, xmax, ref, ticks, xlabel } = o;
  const log = o.scale === "log";
  const lw = o.labelW || 3.6, vw = o.valueW || 2.6;
  const px = x + lw, pw = w - lw - vw, vx = px + pw + 0.15;
  const n = rows.length, rh = (h - 0.55) / n;
  const sc = v => {
    let t = log ? (Math.log(v) - Math.log(xmin)) / (Math.log(xmax) - Math.log(xmin)) : (v - xmin) / (xmax - xmin);
    return px + Math.min(1, Math.max(0, t)) * pw;
  };
  // header
  if (o.head) {
    slide.addText(o.head[0], { x, y: y - 0.02, w: lw, h: 0.3, fontSize: 10, bold: true, color: HX.accent4, margin: 0, isTextBox: true, objectName: "fh1" });
    slide.addText(o.head[1], { x: vx, y: y - 0.02, w: vw, h: 0.3, fontSize: 10, bold: true, color: HX.accent4, margin: 0, isTextBox: true, objectName: "fh2" });
  }
  const top = y + 0.3;
  // zebra + rows
  rows.forEach((r, i) => {
    const ry = top + i * rh, cy = ry + rh / 2;
    if (r.group) {
      slide.addText(r.label, { x, y: ry, w: lw + pw + vw, h: rh, fontSize: 10.5, bold: true, color: HX.dk2, margin: 0, isTextBox: true, valign: "middle", objectName: "fg-" + i });
      return;
    }
    if (i % 2 === 0) slide.addShape(pres.ShapeType.rect, { x, y: ry, w, h: rh, fill: { color: HX.lt2 }, line: { color: HX.lt2 }, objectName: "zebra-" + i });
    const col = r.color || HX.dk1;
    slide.addText(r.label, { x: x + 0.05, y: ry, w: lw - 0.1, h: rh, fontSize: r.dim ? 10 : 10.5, italic: !!r.dim, color: r.dim ? HX.accent4 : HX.dk1, margin: 0, isTextBox: true, valign: "middle", objectName: "fl-" + i });
    const lo = Math.max(xmin, r.lo), hi = Math.min(xmax, r.hi);
    slide.addShape(pres.ShapeType.line, { x: sc(lo), y: cy, w: Math.max(0.01, sc(hi) - sc(lo)), h: 0, line: { color: col, width: r.dim ? 1 : 1.75, dashType: r.dim ? "dash" : "solid" }, objectName: "ci-" + i });
    if (r.lo < xmin) slide.addText("◂", { x: px - 0.12, y: cy - 0.12, w: 0.2, h: 0.24, fontSize: 9, color: col, margin: 0, isTextBox: true, objectName: "arrL-" + i });
    if (r.hi > xmax) slide.addText("▸", { x: px + pw - 0.08, y: cy - 0.12, w: 0.2, h: 0.24, fontSize: 9, color: col, margin: 0, isTextBox: true, objectName: "arrR-" + i });
    const ms = r.big ? 0.2 : 0.14;
    slide.addShape(r.diamond ? pres.ShapeType.diamond : pres.ShapeType.rect, { x: sc(r.est) - ms / 2, y: cy - ms / 2, w: ms, h: ms, fill: { color: r.dim ? HX.lt1 : col }, line: { color: col, width: 1 }, objectName: "pt-" + i });
    slide.addText(r.val, { x: vx, y: ry, w: vw, h: rh, fontSize: 10, color: r.dim ? HX.accent4 : HX.dk1, margin: 0, isTextBox: true, valign: "middle", objectName: "fv-" + i });
  });
  // axis
  const ay = top + n * rh + 0.05;
  slide.addShape(pres.ShapeType.line, { x: px, y: ay, w: pw, h: 0, line: { color: HX.accent4, width: 1 }, objectName: "axis" });
  slide.addShape(pres.ShapeType.line, { x: sc(ref), y: top, w: 0, h: ay - top, line: { color: HX.accent1, width: 1, dashType: "dash" }, objectName: "ref" });
  ticks.forEach(t => {
    slide.addShape(pres.ShapeType.line, { x: sc(t), y: ay, w: 0, h: 0.06, line: { color: HX.accent4, width: 1 }, objectName: "tick-" + t });
    slide.addText(String(t).replace(".", ","), { x: sc(t) - 0.3, y: ay + 0.05, w: 0.6, h: 0.2, fontSize: 9, color: HX.accent4, align: "center", margin: 0, isTextBox: true, objectName: "tl-" + t });
  });
  if (xlabel) slide.addText(xlabel, { x: px, y: ay + 0.22, w: pw, h: 0.22, fontSize: 9, color: HX.accent4, align: "center", margin: 0, isTextBox: true, objectName: "xlabel" });
}

const chartBase = {
  catAxisLabelColor: HX.dk1, valAxisLabelColor: HX.accent4, catAxisLabelFontSize: 10, valAxisLabelFontSize: 9,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt",
  valGridLine: { color: "DDE3EA", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false,
  showValue: true, dataLabelFontSize: 10, dataLabelColor: HX.dk1,
};

let notesStore = [];
const NOTES = JSON.parse(fs.readFileSync(path.join(__dirname, "notes.json"), "utf8"));
function addNotes(slide, n, txt) { const t = NOTES[String(n)] || txt; slide.addNotes(t); notesStore.push({ n, words: t.split(/\s+/).length }); }

// ============================================================
// SLIDE 1 – ÚVOD
// ============================================================
pres.addSection({ title: "Úvod a epidemiologie" });
{
  const s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Úvod a epidemiologie" });
  s.addText("OSA a inzulinová rezistence", { placeholder: "title" });
  s.addText("Mechanismy, evidence a klinické důsledky pro kardiologa", { placeholder: "body" });
  s.addText("Centrum preventivní medicíny · [jméno přednášejícího – doplnit]", { x: 0.7, y: 4.7, w: 11.9, h: 0.4, fontSize: 14, color: HX.accent4, margin: 0, isTextBox: true, objectName: "autor" });
  stat(s, 0.7, 5.35, 3.6, "~936 mil.", "dospělých 30–69 let s OSA (AHI ≥ 5) celosvětově\nBenjafield, Lancet Respir Med 2019 · modelový odhad", HX.accent3);
  stat(s, 4.7, 5.35, 3.6, "82–93 %", "mužů / žen se středně těžkou–těžkou OSA bez diagnózy\nYoung, Sleep 1997 · průřez", HX.accent1);
  stat(s, 8.7, 5.35, 4.0, "40–80 %", "prevalence OSA u hypertenze, HF, ICHS, FS, CMP\nAHA scientific statement, Circulation 2021", HX.accent2);
  addNotes(s, 1, `Dobrý den, děkuji za pozvání. Dnes bych chtěl mluvit o obstrukční spánkové apnoi z pohledu, který kardiologům často uniká: jako o metabolickém onemocnění, které se potkává s inzulinovou rezistencí a diabetem přímo v našich ambulancích. Proč bychom to měli řešit právě my? Tři čísla. Podle modelového odhadu z roku 2019 má OSA s AHI nad 5 přibližně 936 milionů dospělých ve věku 30 až 69 let a 425 milionů má středně těžkou až těžkou formu (Benjafield 2019). Je to odhad financovaný výrobcem CPAP a závisí na přepočtu mezi skórovacími kritérii, ale řádově ho nikdo nezpochybňuje. Druhé číslo: ve Wisconsinské kohortě z devadesátých let nebylo klinicky diagnostikováno 82 procent mužů a 93 procent žen se středně těžkou až těžkou apnoí, a to v populaci bez bariér přístupu k péči (Young 1997). A třetí: podle stanoviska AHA z roku 2021 se prevalence OSA u pacientů s hypertenzí, srdečním selháním, ischemickou chorobou, plicní hypertenzí, fibrilací síní a po cévní mozkové příhodě pohybuje mezi 40 a 80 procenty (Yeghiazarians 2021). Těmito pacienty je plná každá naše ambulance. Pojďme se nejdřív podívat, o kterých diagnózách konkrétně mluvíme.`);
}

// ============================================================
// SLIDE 2 – OSA u KV pacientů
// ============================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Úvod a epidemiologie" });
  s.addText("OSA u našich pacientů: rezistentní hypertenze, fibrilace síní, srdeční selhání", { placeholder: "title" });
  chip(s, ["průřez"]);
  msg(s, "U rezistentní hypertenze je OSA nejčastější „sekundární“ nález, u FS je 2× častější než u ostatních kardiaků se stejným BMI, u HFrEF má SDB tři čtvrtiny pacientů.");
  s.addChart(pres.ChartType.bar, [{
    name: "Prevalence (%)",
    labels: ["Rezistentní HTN, AHI > 15 (Pedrosa 2011, n = 125)", "Rezistentní HTN, AHI ≥ 10 (Logan 2001, n = 41)", "FS před kardioverzí, Berlin dotazník (Gami 2004, n = 151)", "Kardiologická kontrola, stejný BMI (Gami 2004, n = 312)", "HFrEF – obstrukční SA (Oldenburg 2007, n = 700)", "HFrEF – centrální SA (Oldenburg 2007, n = 700)"],
    values: [64, 83, 49, 32, 36, 40],
  }], Object.assign({}, chartBase, {
    x: 0.5, y: 2.0, w: 8.3, h: 4.45, barDir: "bar", barGapWidthPct: 45,
    chartColors: [HX.accent1, HX.accent1, HX.accent2, HX.accent4, HX.dk2, HX.accent4],
    valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 20, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"',
    catAxisOrientation: "maxMin",
  }));
  box(s, 9.1, 2.0, 3.75, 4.45, "Co je za čísly", [
    "Pedrosa 2011: systematické hledání sekundárních příčin; OSA 64,0 % vs. primární hyperaldosteronismus 5,6 %",
    "Logan 2001: n = 41, BMI 34, terciární centrum – číslo uvádět s n",
    "Gami 2004: OSA dle dotazníku, ne PSG; adj. OR FS–OSA 2,19 (1,40–3,42)",
    "Oldenburg 2007: SDB celkem 76 %; u HF je CSA stejně častá jako OSA – rozlišovat",
    "Všechny studie průřezové → asociace, ne kauzalita",
  ], HX.dk1);
  cite(s, "Pedrosa, Hypertension 2011 · Logan, J Hypertens 2001 · Gami, Circulation 2004 · Oldenburg, Eur J Heart Fail 2007");
  addNotes(s, 2, `Když se podíváme na konkrétní kardiologické populace, jsou čísla přesvědčivá, i když všechna pocházejí z průřezových studií, tedy ukazují asociaci, nikoli příčinu. U rezistentní hypertenze ukázalo systematické vyšetření 125 pacientů v São Paulu na všechny známé sekundární příčiny OSA s AHI nad 15 u 64 procent, zatímco primární hyperaldosteronismus jen u 5,6 procenta (Pedrosa 2011). Starší torontská práce na pouhých 41 pacientech našla OSA u 83 procent, ale populace měla BMI 34 a práh AHI 10 (Logan 2001). U fibrilace síní měla OSA podle Berlínského dotazníku skoro polovina pacientů před kardioverzí, 49 procent, oproti 32 procentům v běžné kardiologické ambulanci se stejným BMI; adjustované OR bylo 2,19 (Gami 2004). A u 700 pacientů se srdečním selháním s ejekční frakcí pod 40 procent mělo poruchu dýchání ve spánku 76 procent, z toho 36 procent obstrukční a 40 procent centrální (Oldenburg 2007). Pro kardiologa je důležité, že u srdečního selhání je centrální apnoe stejně častá jako obstrukční a léčí se jinak. Pozor na limity: malé vzorky, dotazník místo polysomnografie, obézní populace. Ale jedno sdělení drží: OSA není u našich pacientů vzácná komorbidita, je to pravidlo. A většina z nich má zároveň poruchu glukózového metabolismu, ke které se dostáváme teď.`);
}

// ============================================================
// SLIDE 3 – Překryv OSA–obezita–DM2
// ============================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Úvod a epidemiologie" });
  s.addText("Překryv OSA – obezita – DM2: kolik z toho je obezita?", { placeholder: "title" });
  chip(s, ["průřez", "MA obs."]);
  msg(s, "OSA má 36–87 % diabetiků podle BMI a metody; po adjustaci na BMI a obvod krku část asociace mizí, část přetrvává – obezita vysvětluje hodně, ne vše.");
  s.addChart(pres.ChartType.bar, [{
    name: "Prevalence OSA u DM2 (%)",
    labels: ["Sleep AHEAD, obézní DM2, BMI 36,5, PSG (Foster 2009, n = 305)", "Diabetologická ambulance, AHI ≥ 15 (Einhorn 2007, n = 279)", "MA 23 studií, objektivně potvrzená OSA (Sun 2026)", "MA 23 studií – jen PSG-potvrzená (Sun 2026)"],
    values: [86.6, 36, 37, 54],
  }], Object.assign({}, chartBase, {
    x: 0.5, y: 2.0, w: 6.6, h: 3.6, barDir: "bar", barGapWidthPct: 40,
    chartColors: [HX.accent1, HX.dk2, HX.accent2, HX.accent2],
    valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 20, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"',
    catAxisOrientation: "maxMin", showTitle: true, title: "Prevalence OSA u diabetiků 2. typu", titleFontSize: 11, titleColor: HX.dk2,
  }));
  s.addText("MA 2026: pooled 37 % (95 % CI 27–49), I² ≈ 99 %; severoamerických 86 % pochází z jediné studie (Sun 2026). Obousměrně: DM u OSA OR 2,29 (1,93–2,72), OSA u DM OR 2,12 (1,73–2,60) (Huang & Chen 2024).", { x: 0.5, y: 5.65, w: 6.6, h: 0.85, fontSize: 10.5, color: HX.dk2, margin: 0, isTextBox: true, valign: "top", objectName: "ma-pozn" });
  box(s, 7.4, 2.0, 2.65, 4.5, "Asociace po adjustaci MIZÍ", [
    "Sleep Heart Health Study, n = 5 874 (Resnick 2003)",
    "Po adjustaci na věk, pohlaví, BMI, rasu a obvod krku rozdíly v RDI a hypoxémii mezi diabetiky a nediabetiky vymizely",
    "Nezávislé zůstalo jen periodické dýchání: OR 1,80 (1,02–3,15)",
    "„largely explained by obesity“",
  ], HX.accent1, { fontSize: 10.5 });
  box(s, 10.2, 2.0, 2.65, 4.5, "Asociace po adjustaci PŘETRVÁVÁ", [
    "Oxford, muži s DM2, oxymetrie (West 2006): OSA 17 % vs. 6 % v obecné populaci",
    "BMI vysvětlil 13 % variance, diabetes dalších 8 % – nezávislý, ale malý příspěvek",
    "SHHS, OGTT (Seicean 2008): asociace SDB s IGT/okultním DM podobná u ne-nadvážných i nadvážných",
  ], HX.accent2, { fontSize: 10.5 });
  cite(s, "Foster, Diabetes Care 2009 · Einhorn, Endocr Pract 2007 · Sun, Can J Diabetes 2026 · Huang & Chen, J Diabetes Investig 2024 · Resnick, Diabetes Care 2003 · West, Thorax 2006 · Seicean, Diabetes Care 2008");
  addNotes(s, 3, `Jak silný je překryv OSA s diabetem? Záleží na tom, koho a jak měříme. Ve studii Sleep AHEAD mělo z 305 obézních diabetiků s průměrným BMI 36,5 apnoi 86,6 procenta, téměř čtvrtina těžkou, a žádný symptom včetně spavosti ji nepredikoval (Foster 2009). V běžné diabetologické ambulanci s jednokanálovým záznamem to bylo 36 procent při AHI nad 15 (Einhorn 2007). Letošní metaanalýza 23 studií dává poolovanou prevalenci objektivně potvrzené OSA u diabetiků 37 procent, při čistě polysomnografickém potvrzení 54 procent, ale s heterogenitou 99 procent; severoamerických 86 procent označuje za nereprezentativní (Sun 2026). Vztah je obousměrný: diabetes je u OSA 2,3krát častější a OSA u diabetu 2,1krát (Huang a Chen 2024). A teď klíčová otázka pro poctivého kardiologa: kolik z toho je prostě obezita? Ve Sleep Heart Health Study rozdíly v obstrukčních parametrech mezi diabetiky a nediabetiky po adjustaci na věk, pohlaví, BMI a obvod krku vymizely; nezávislé zůstalo jen periodické dýchání (Resnick 2003). Naopak v Oxfordu zůstal diabetes nezávislým prediktorem OSA i po korekci na BMI, ale vysvětloval jen 8 procent variance oproti 13 procentům u BMI (West 2006), a asociace apnoe s poruchou glukózové tolerance byla podobná u neobézních i obézních (Seicean 2008). Obezita tedy vysvětluje hodně, ale ne všechno. Jak by mohla apnoe sama zhoršovat inzulinovou senzitivitu?`);
}

// ============================================================
// SLIDE 4 – Mechanismy I
// ============================================================
pres.addSection({ title: "Patofyziologie a obousměrnost" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Patofyziologie a obousměrnost" });
  s.addText("Mechanismy I: intermitentní hypoxie a fragmentace spánku snižují inzulinovou senzitivitu – i u zdravých", { placeholder: "title" });
  chip(s, ["mechanistické (lidé)", "zvířecí"]);
  msg(s, "Dvě experimentální expozice u zdravých dobrovolníků sníží inzulinovou senzitivitu bez kompenzačního vzestupu sekrece inzulinu; zvířecí clamp ukazuje přímý tkáňový efekt hypoxie.");
  // schéma
  const sy = 2.15;
  const node = (x, y, w, h, t, fill, tc) => { s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.12, objectName: "n-" + t }); s.addText(t, { x, y, w, h, fontSize: 12, bold: true, color: tc || HX.lt1, align: "center", valign: "middle", margin: 0.05, isTextBox: true, objectName: "nt-" + t }); };
  const arrow = (x, y, w) => s.addShape(pres.ShapeType.rightArrow, { x, y, w, h: 0.3, fill: { color: HX.accent4 }, line: { color: HX.accent4 }, objectName: "arr" + x + y });
  node(0.5, sy + 0.55, 1.7, 0.9, "Obstrukční\nspánková apnoe", HX.dk1);
  arrow(2.3, sy + 0.85, 0.5);
  node(2.9, sy, 2.1, 0.8, "Intermitentní hypoxie\n(desaturace / reoxygenace)", HX.accent1);
  node(2.9, sy + 1.2, 2.1, 0.8, "Fragmentace spánku\n(arousaly, ↓ SWS)", HX.accent3, HX.dk1);
  arrow(5.1, sy + 0.25, 0.5); arrow(5.1, sy + 1.45, 0.5);
  node(5.7, sy + 0.55, 1.9, 0.9, "↓ inzulinová senzitivita\nbez ↑ sekrece inzulinu", HX.accent2);
  // chart before/after
  s.addChart(pres.ChartType.bar, [
    { name: "Kontrola / baseline", labels: ["5 h intermitentní hypoxie\nv bdění, n = 13 (Louis 2009)", "2 noci fragmentace spánku,\nn = 11 (Stamatakis 2010)"], values: [4.1, 5.02] },
    { name: "Expozice", labels: ["5 h intermitentní hypoxie\nv bdění, n = 13 (Louis 2009)", "2 noci fragmentace spánku,\nn = 11 (Stamatakis 2010)"], values: [3.4, 3.76] },
  ], Object.assign({}, chartBase, {
    x: 7.9, y: 1.95, w: 4.95, h: 3.1, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
    chartColors: [HX.accent4, HX.accent1], showLegend: true, legendPos: "b", legendFontSize: 9, legendFontFace: "+mn-lt", legendColor: HX.dk1,
    dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0#", valAxisMinVal: 0, valAxisMaxVal: 6,
    showTitle: true, title: "Inzulinová senzitivita S_I (IVGTT, minimal model), (mU/l)⁻¹·min⁻¹", titleFontSize: 10, titleColor: HX.dk2,
  }));
  box(s, 0.5, 4.55, 7.1, 1.95, "Co data říkají – a co ne", [
    "Louis & Punjabi 2009: S_I 4,1 → 3,4 (p = 0,018), S_G 1,9 → 1,3; sekrece inzulinu beze změny; posun k sympatiku",
    "Stamatakis & Punjabi 2010: S_I 5,02 → 3,76 (p < 0,0001); ↑ ranní kortizol a sympatická aktivita – bez hypoxie",
    "Tasali 2008 (PNAS): selektivní suprese SWS → „marked decrease“ S_I, korelující s mírou redukce SWS (procento neuvádím – neověřeno z plného textu)",
    "Limity: akutní expozice, n 11–13, zdraví dobrovolníci → mechanismus, ne klinický efekt",
  ], HX.dk1, { fontSize: 10.5 });
  box(s, 7.9, 5.2, 4.95, 1.3, "ZVÍŘECÍ: Iiyori 2007 (myš, clamp)", [
    "IH 9 h: inzulinová senzitivita 49,4 → 38,8 mg/kg/min (p < 0,005); blokáda autonomního systému pokles neodvrátila → přímý tkáňový efekt",
  ], HX.accent1, { fontSize: 10.5, fill: "FBEDE9" });
  cite(s, "Louis & Punjabi, J Appl Physiol 2009 · Stamatakis & Punjabi, Chest 2010 · Tasali, PNAS 2008 · Iiyori, Am J Respir Crit Care Med 2007 (zvířecí)");
  addNotes(s, 4, `Mechanisticky má OSA dvě hlavní složky: intermitentní hypoxii a fragmentaci spánku. Obě lze u lidí experimentálně navodit odděleně, a to je pro kauzální uvažování cenné. Třináct zdravých dobrovolníků dýchalo pět hodin v bdělém stavu intermitentně hypoxickou směs s průměrně 24 desaturacemi za hodinu; inzulinová senzitivita z intravenózního glukózového testu klesla ze 4,1 na 3,4 a sekrece inzulinu to nekompenzovala, zároveň se posunula autonomní rovnováha k sympatiku (Louis a Punjabi 2009). Jedenáct zdravých dobrovolníků mělo dvě noci spánek fragmentovaný akustickými a mechanickými podněty bez jakékoli hypoxie; inzulinová senzitivita klesla z 5,02 na 3,76 a stoupl ranní kortizol i sympatická aktivita (Stamatakis a Punjabi 2010). Selektivní potlačení hlubokého spánku u mladých zdravých dospělých vedlo k výraznému poklesu inzulinové senzitivity, který koreloval s mírou redukce hlubokého spánku (Tasali 2008); přesné procento neuvádím, protože ho nemám ověřené z plného textu. U myší potom intermitentní hypoxie snížila inzulinovou senzitivitu měřenou clampem z 49,4 na 38,8 miligramů na kilogram a minutu a blokáda autonomního systému to neodvrátila, což ukazuje na přímý tkáňový efekt hypoxie (Iiyori 2007). Pozor: jde o akutní expozice u malých skupin zdravých lidí; dokládají mechanismus, ne klinický efekt. Které mediátory to spojují s tím, co léčíme my?`);
}

// ============================================================
// SLIDE 5 – Mechanismy II
// ============================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Patofyziologie a obousměrnost" });
  s.addText("Mechanismy II: sympatikus, osa HPA, zánět, adipokiny, jaterní inzulinová rezistence", { placeholder: "title" });
  chip(s, ["mechanistické (lidé)", "MA obs."]);
  msg(s, "Mediátory jsou kardiologovi důvěrně známé – a sympatická hyperaktivita je atributem apnoe, nikoli obezity samotné.");
  const cards = [
    ["Sympatikus", "MSNA 61 ± 8 u obézních s okultní OSA vs. 42 ± 3 u obézních bez OSA a 41 ± 3 u štíhlých (bursts/100 tepů) – Narkiewicz 1998\nVysoká MSNA i v bdění; vrchol na konci apnoe – Somers 1995", HX.accent1],
    ["Osa HPA / kortizol", "24h kortizol vyšší u OSA než u kontrol; CPAP ↓ vs. baseline, ale rozdíl vs. sham-CPAP nesignifikantní – Kritikou 2016 (RCT crossover)\nMA 6 RCT: SMD −0,39 (−0,75 až −0,03) – Ken-Dror 2021", HX.accent5],
    ["Jedna noc bez CPAP", "Vysazení CPAP → ↑ noční volné mastné kyseliny (p = 0,007), glykémie (p = 0,028), kortizol (p = 0,037), sTK (p = 0,017) úměrně počtu událostí – Chopra 2017 (RCT crossover, n = 31)", HX.dk2],
    ["Zánět", "IN VITRO: IH selektivně aktivuje NF-κB, ne HIF-1\nPacienti: TNF-α 2,56 vs. 1,25 pg/ml (p < 0,001), po 6 týdnech CPAP 1,24 – Ryan 2005\nMA 51 studií: SMD CRP 1,77, IL-6 2,16 (CI neuvedeny; bez adj. na BMI) – Nadeem 2013", HX.accent1],
    ["Adipokiny", "Leptin 13,7 ± 1,3 vs. 9,2 ± 1,2 ng/ml u stejně obézních mužů s/bez OSA (p = 0,02) – Phillips 2000\nAdiponektin: SMD −0,71 (−0,92 až −0,49), 20 studií – Lu 2019", HX.accent2],
    ["Jaterní IR / MASLD", "MA 18 průřezových studií: OR NAFLD 2,01 (1,36–2,97) histologie, 2,99 (1,79–4,99) radiologie; NASH 2,37; fibróza 2,16 – Musso 2013; efekt nezávislý na BMI v meta-regresi\nVázáno na stupeň noční hypoxémie – Mesarwi 2019", HX.accent2],
  ];
  cards.forEach((c, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 4.15, y = 2.0 + row * 2.25, w = 3.95, h = 2.1;
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: HX.lt2 }, line: { color: HX.lt2 }, rectRadius: 0.1, objectName: "card-" + c[0] });
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.15, y: y + 0.15, w: 0.4, h: 0.4, fill: { color: c[2] }, line: { color: c[2] }, objectName: "dot-" + c[0] });
    s.addText(String(i + 1), { x: x + 0.15, y: y + 0.15, w: 0.4, h: 0.4, fontSize: 12, bold: true, color: HX.lt1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "dotn-" + c[0] });
    s.addText(c[0], { x: x + 0.65, y: y + 0.15, w: w - 0.8, h: 0.4, fontSize: 13, bold: true, color: HX.dk1, valign: "middle", margin: 0, isTextBox: true, objectName: "ct-" + c[0] });
    s.addText(c[1], { x: x + 0.15, y: y + 0.62, w: w - 0.3, h: h - 0.72, fontSize: 10, color: HX.dk1, valign: "top", margin: 0, isTextBox: true, objectName: "cb-" + c[0] });
  });
  cite(s, "Somers, J Clin Invest 1995 · Narkiewicz, Circulation 1998 · Kritikou, Eur Respir J 2016 · Ken-Dror, Clin Endocrinol 2021 · Chopra, JCEM 2017 · Ryan, Circulation 2005 · Nadeem, JCSM 2013 · Phillips, Am J Physiol 2000 · Lu, Lipids Health Dis 2019 · Musso, Obes Rev 2013 · Mesarwi, AJRCCM 2019");
  addNotes(s, 5, `Mediátory jsou vám důvěrně známé. Sympatikus: pacienti s OSA mají vysokou svalovou sympatickou aktivitu i v bdělém stavu a na konci apnoe stoupá na trojnásobek (Somers 1995). Důležité je, že samotná obezita bez apnoe sympatikus nezvyšuje: obézní bez OSA měli 42 vzruchů na 100 tepů stejně jako štíhlí, zatímco obézní se skrytou apnoí 61 (Narkiewicz 1998). Osa HPA: u relativně neobézních pacientů byl 24hodinový kortizol vyšší než u kontrol a CPAP ho snížil proti výchozí hodnotě, ale rozdíl proti sham-CPAP nebyl signifikantní (Kritikou 2016); metaanalýza šesti RCT dává malý efekt, SMD minus 0,39 (Ken-Dror 2021). Když pacientům adaptovaným na CPAP přístroj na jednu noc vysadíme, během spánku stoupnou volné mastné kyseliny, glykémie, kortizol i systolický tlak úměrně počtu respiračních událostí (Chopra 2017). Zánět: in vitro aktivuje intermitentní hypoxie selektivně NF-kappa B, nikoli adaptivní HIF-1, a u pacientů byl TNF-alfa 2,56 oproti 1,25 pikogramů na mililitr a po šesti týdnech CPAP se normalizoval (Ryan 2005). Adipokiny: vyšší leptin u stejně obézních mužů s OSA, 13,7 oproti 9,2 nanogramů na mililitr (Phillips 2000), a nižší adiponektin, SMD minus 0,71 (Lu 2019). A játra: OSA zdvojnásobuje až ztrojnásobuje riziko steatotické choroby jater, OR 2,01 až 2,99, nezávisle na BMI v meta-regresi (Musso 2013), v závislosti na stupni noční hypoxémie (Mesarwi 2019). Většina těchto dat je ale průřezová a často neadjustovaná na BMI; berte je jako mechanistický rámec, ne jako důkaz. Kterým směrem vztah vlastně běží?`);
}

// ============================================================
// SLIDE 6 – Obousměrnost
// ============================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Patofyziologie a obousměrnost" });
  s.addText("Obousměrnost: obezita a IR → OSA, OSA → IR/DM2 – ale asymetricky", { placeholder: "title" });
  chip(s, ["kohorta", "MR", "průřez"]);
  msg(s, "Přibrání o 10 % zvedne AHI o ~32 %; OSA zvyšuje riziko DM2 i po adjustaci na BMI a pas (HR 1,37), opačný směr je po adjustaci hraniční (HR 1,08).");
  const nb = (x, y, w, h, t, fill) => { s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.15, objectName: "nb-" + t }); s.addText(t, { x, y, w, h, fontSize: 15, bold: true, color: HX.lt1, align: "center", valign: "middle", margin: 0.05, isTextBox: true, objectName: "nbt-" + t }); };
  nb(0.5, 2.75, 2.6, 1.3, "Obezita\ninzulinová rezistence\nDM2", HX.accent2);
  nb(9.0, 2.75, 2.6, 1.3, "Obstrukční\nspánková apnoe", HX.accent1);
  s.addShape(pres.ShapeType.rightArrow, { x: 3.3, y: 2.7, w: 5.5, h: 0.5, fill: { color: HX.accent2 }, line: { color: HX.accent2 }, objectName: "a1" });
  s.addShape(pres.ShapeType.leftArrow, { x: 3.3, y: 3.6, w: 5.5, h: 0.5, fill: { color: HX.accent1 }, line: { color: HX.accent1 }, objectName: "a2" });
  s.addText("DM2 → OSA: HR 1,53 (1,32–1,77) → po adj. na BMI + pas 1,08 (1,00–1,16); inzulinem léčený DM 1,43 (1,11–1,83) – Huang 2018\n+10 % hmotnosti → +32 % AHI (20–45 %), 6× odds středně těžké–těžké SDB (2,2–17,0) – Peppard 2000 · MR: BMI → OSA kauzálně, genetická korelace r = 0,72 – Strausz 2021", { x: 3.3, y: 1.95, w: 5.5, h: 0.72, fontSize: 9.5, color: HX.dk1, valign: "bottom", margin: 0, isTextBox: true, objectName: "t1" });
  s.addText("OSA → DM2: HR 2,06 (1,86–2,28) → po adj. na BMI + pas 1,37 (1,24–1,53); 146 519 osob, 3 kohorty, 10–18 let – Huang 2018", { x: 3.3, y: 4.15, w: 5.5, h: 0.5, fontSize: 9.5, color: HX.dk1, valign: "top", margin: 0, isTextBox: true, objectName: "t2" });
  box(s, 0.5, 5.05, 4.0, 1.45, "Mechanismus DM2 → OSA: dysautonomie", [
    "Neobézní diabetici s autonomní neuropatií (BMI ~24): obstrukční události u > 30 %, u diabetiků bez neuropatie u nikoho; ↓ periferní chemosenzitivita k CO₂ – Bottini 2003; Ficker 1998 (26 % vs. 0 %)",
  ], HX.accent2, { fontSize: 10, bullet: false });
  box(s, 4.7, 5.05, 4.0, 1.45, "Mendelovská randomizace: PRO", [
    "OSA → DM2 OR 2,34 (1,59–3,44); DM2 → OSA OR 1,03 (1,01–1,05) – Hrytsenko 2025",
    "OSA → T2D OR 2,02 (1,19–3,43) – Xia 2023 (3 instrumenty)",
  ], HX.accent5, { fontSize: 10 });
  box(s, 8.9, 5.05, 3.95, 1.45, "Mendelovská randomizace: PROTI", [
    "OSA → T2D OR 3,58 (1,06–12,11), ale po odstranění lokusu FTO (obezita) OR 1,30 (0,68–2,50) – Ding 2023 → kauzalita nezávislá na obezitě není v MR jednoznačná",
  ], HX.accent5, { fontSize: 10, bullet: false });
  cite(s, "Peppard, JAMA 2000 · Strausz, Eur Respir J 2021 · Huang, Diabetes Care 2018 · Bottini, Eur Respir J 2003 · Ficker, Eur Respir J 1998 · Hrytsenko, Commun Med 2025 · Xia 2023 · Ding, Front Psychiatry 2023");
  addNotes(s, 6, `Vztah je obousměrný, ale ne symetrický. Směr obezita k apnoi je jasný: ve Wisconsinské kohortě předpovídal desetiprocentní nárůst hmotnosti nárůst AHI o 32 procent a šestinásobné riziko vzniku středně těžké až těžké apnoe, byť s širokým intervalem 2,2 až 17 (Peppard 2000); mendelovská randomizace ve FinnGen kauzalitu BMI k OSA podporuje a genetická korelace obou znaků je 0,72 (Strausz 2021). Nejlepší humánní data pro oba směry najednou pocházejí ze tří amerických kohort se 146 tisíci osobami sledovanými 10 až 18 let: OSA zvyšovala riziko diabetu s HR 2,06, po adjustaci na BMI a obvod pasu 1,37; diabetes zvyšoval riziko OSA s HR 1,53, ale po adjustaci na adipozitu už jen 1,08, jasně pouze u inzulinem léčených, HR 1,43 (Huang 2018). Mechanismem opačného směru může být autonomní neuropatie: u neobézních diabetiků s dysautonomií měla obstrukční události více než třetina, u diabetiků bez neuropatie nikdo, a měli sníženou chemosenzitivitu k oxidu uhličitému (Bottini 2003; Ficker 1998). Mendelovská randomizace pro směr OSA k diabetu je rozporná: jedna analýza dává OR 2,34 (Hrytsenko 2025), jiná ukazuje, že signál zmizí po odstranění obezitního lokusu FTO (Ding 2023). Kauzalitu u lidí tedy považujme za pravděpodobnou, ne prokázanou. Podívejme se na prospektivní data.`);
}

// ============================================================
// SLIDE 7 – MA kohort
// ============================================================
pres.addSection({ title: "Observační evidence" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Observační evidence" });
  s.addText("Observační evidence: OSA a incidence DM2 – metaanalýzy prospektivních kohort", { placeholder: "title" });
  chip(s, ["MA obs."]);
  msg(s, "Tři metaanalýzy kohort ukazují RR 1,40–2,02 pro incidentní DM2 s gradientem podle tíže a lineárním vztahem k AHI – jde o asociaci; I² a adjustace na BMI v abstraktech chybí.");
  forest(s, {
    x: 0.5, y: 2.05, w: 12.35, h: 4.3, scale: "log", xmin: 0.5, xmax: 4, ref: 1, ticks: [0.5, 1, 2, 4], xlabel: "Relativní riziko incidentního DM2 (log škála)",
    head: ["Metaanalýza · expozice", "RR (95 % CI)"], labelW: 5.4, valueW: 2.4,
    rows: [
      { label: "Wang 2013 · 6 kohort, n 5 953, 332 DM, FU 2,7–16 let", group: true },
      { label: "   mírná OSA vs. bez OSA", est: 1.22, lo: 0.91, hi: 1.63, color: HX.accent4, val: "1,22 (0,91–1,63) NS" },
      { label: "   středně těžká–těžká OSA vs. bez OSA", est: 1.63, lo: 1.09, hi: 2.45, color: HX.dk2, val: "1,63 (1,09–2,45)" },
      { label: "Anothaisintawee 2016 · 36 studií, 1 061 555 osob (všechny spánkové expozice)", group: true },
      { label: "   OSA vs. bez OSA", est: 2.02, lo: 1.57, hi: 2.61, color: HX.dk2, val: "2,02 (1,57–2,61)" },
      { label: "   pro srovnání: nadváha", est: 2.99, lo: 2.42, hi: 3.72, color: HX.accent4, dim: true, val: "2,99 (2,42–3,72)" },
      { label: "   pro srovnání: fyzická inaktivita", est: 1.20, lo: 1.11, hi: 1.32, color: HX.accent4, dim: true, val: "1,20 (1,11–1,32)" },
      { label: "Qie 2020 · 16 kohort, n 338 912, 19 355 DM2, medián FU 10,5 let", group: true },
      { label: "   OSA vs. bez OSA (binární)", est: 1.40, lo: 1.32, hi: 1.48, color: HX.dk2, big: true, val: "1,40 (1,32–1,48)" },
      { label: "   na každých +5 událostí AHI/h (dose–response, lineární)", est: 1.08, lo: 1.01, hi: 1.14, color: HX.accent1, val: "1,08 (1,01–1,14)" },
    ],
  });
  cite(s, "Wang, Respirology 2013 · Anothaisintawee, Sleep Med Rev 2016 · Qie, J Diabetes 2020 — I² a adjustace zahrnutých studií na BMI: v abstraktech neuvedeny (NEOVĚŘENO)");
  addNotes(s, 7, `Co říkají prospektivní kohorty o tom, zda apnoe předchází vzniku diabetu? Máme tři metaanalýzy. První zahrnula šest kohort s necelými šesti tisíci účastníky a 332 novými případy diabetu: středně těžká až těžká OSA měla relativní riziko 1,63, mírná 1,22 a ta nebyla signifikantní (Wang 2013). Druhá srovnala poruchy spánku s tradičními rizikovými faktory: OSA měla riziko 2,02, pro srovnání nadváha 2,99, rodinná anamnéza 2,33 a fyzická inaktivita 1,20 (Anothaisintawee 2016). Největší, se 16 kohortami a 339 tisíci účastníky sledovanými v mediánu 10,5 roku, dává riziko 1,40 s velmi úzkým intervalem a lineární vztah dávka–odpověď: každých pět událostí za hodinu navíc zvyšuje riziko diabetu o 8 procent (Qie 2020). Pro kardiologa je to riziko řádově srovnatelné s faktory, které běžně vkládáme do skórovacích systémů. Dva limity musím říct nahlas. U žádné z těchto metaanalýz nemám z abstraktu heterogenitu ani informaci, zda byly zahrnuté studie adjustovány na BMI, protože plné texty nebyly dostupné. A jde o asociaci, kterou reziduální confounding obezitou může částečně vysvětlovat. Proto se podívejme na jednotlivé kohorty, které na adipozitu adjustovaly důsledně.`);
}

// ============================================================
// SLIDE 8 – Dávka–odpověď a nezávislost na BMI
// ============================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Observační evidence" });
  s.addText("Dávka–odpověď a nezávislost na BMI – a poctivá protiváha", { placeholder: "title" });
  chip(s, ["kohorta", "průřez"]);
  msg(s, "Riziko přetrvává po adjustaci na BMI i obvod pasu a platí i u obézních; gradient sahá až k AHI ≥ 50 – ale Wisconsinská kohorta longitudinálně signifikanci neukázala.");
  forest(s, {
    x: 0.5, y: 2.05, w: 8.6, h: 4.3, scale: "log", xmin: 0.5, xmax: 6, ref: 1, ticks: [0.5, 1, 2, 4], xlabel: "HR / OR (log škála)",
    head: ["Studie · srovnání", "HR / OR (95 % CI)"], labelW: 4.2, valueW: 1.9,
    rows: [
      { label: "Nagayoshi 2016 · ARIC/SHHS, n 1 453, 13 let, incidentní DM", group: true },
      { label: "   těžká OSA, adj. demografie + chování", est: 2.75, lo: 1.81, hi: 4.19, color: HX.accent4, val: "2,75 (1,81–4,19)" },
      { label: "   + BMI a obvod pasu", est: 1.71, lo: 1.08, hi: 2.71, color: HX.dk2, big: true, val: "1,71 (1,08–2,71)" },
      { label: "   jen obézní (BMI ≥ 30)", est: 2.03, lo: 1.20, hi: 3.44, color: HX.dk2, val: "2,03 (1,20–3,44)" },
      { label: "Priou 2012 · n 1 599 bez DM, HbA1c > 6,0 %, adj. BMI + pas", group: true },
      { label: "   AHI 15–< 30 vs. < 5", est: 1.80, lo: 1.19, hi: 2.72, color: HX.accent1, val: "1,80 (1,19–2,72)" },
      { label: "   AHI 30–< 50", est: 2.02, lo: 1.31, hi: 3.14, color: HX.accent1, val: "2,02 (1,31–3,14)" },
      { label: "   AHI ≥ 50", est: 2.96, lo: 1.58, hi: 5.54, color: HX.accent1, val: "2,96 (1,58–5,54)" },
      { label: "Seicean 2008 · SHHS, n 2 588 · okultní DM (OGTT), adj. BMI + pas", est: 1.7, lo: 1.1, hi: 2.7, color: HX.dk2, val: "1,7 (1,1–2,7)" },
      { label: "Reichmuth 2005 · Wisconsin, n 1 387 · incidentní DM za 4 roky, AHI ≥ 15", est: 1.62, lo: 0.67, hi: 3.65, color: HX.accent1, dim: true, val: "1,62 (0,67–3,65) NS" },
    ],
  });
  box(s, 9.35, 2.05, 3.5, 2.45, "Přesné měření IR: Punjabi 2009", [
    "118 nediabetiků, FSIVGTT + minimal model, % tělesného tuku",
    "Pokles inzulinové senzitivity o 26,7 / 36,5 / 43,7 % u mírné / střední / těžké SDB nezávisle na % tuku (CI neuvedeny)",
    "↓ disposition index (β-buňka) u střední–těžké SDB; korelace s desaturací",
  ], HX.dk1, { fontSize: 10 });
  box(s, 9.35, 4.65, 3.5, 1.7, "Protiváha", [
    "Wisconsin průřezově OR 2,30 (1,28–4,11), ale po 4 letech OR 1,62 (0,67–3,65) – krátké sledování, málo událostí; „not clear that SDB is causal“ (Reichmuth 2005)",
  ], HX.accent1, { fontSize: 10, bullet: false, fill: "FBEDE9" });
  cite(s, "Nagayoshi, Sleep Med 2016 · Priou, Diabetes Care 2012 · Seicean, Diabetes Care 2008 · Reichmuth, AJRCCM 2005 · Punjabi & Beamer, AJRCCM 2009 — vše observační: asociace");
  addNotes(s, 8, `Nejsilnější populační důkaz nezávislosti na obezitě pochází z kohorty ARIC a Sleep Heart Health Study: 1453 nediabetiků sledovaných 13 let. Těžká OSA měla pro nový diabetes HR 2,75, po adjustaci na BMI i obvod pasu 1,71, a mezi obézními zůstala 2,03 (Nagayoshi 2016). Průřezová francouzská data u 1599 pacientů bez diabetu ukazují gradient: pravděpodobnost prediabetického HbA1c nad 6 procent rostla s tíží apnoe z OR 1,80 přes 2,02 až na 2,96 při AHI nad 50, po adjustaci na BMI i obvod pasu (Priou 2012). U 118 nediabetiků měřených přesnou metodou FSIVGTT klesla inzulinová senzitivita o 27, 37 a 44 procent podle tíže apnoe nezávisle na procentu tělesného tuku, a klesala i funkce beta-buněk (Punjabi 2009). Asociace platila i u neobézních účastníků Sleep Heart Health Study, OR 1,7 pro okultní diabetes (Seicean 2008). Teď protiváha, kterou považuji za důležitou: ve Wisconsinské kohortě byla průřezová asociace jasná, OR 2,30, ale čtyřleté sledování dalo OR 1,62 s intervalem 0,67 až 3,65, tedy nesignifikantní (Reichmuth 2005). Důvodem je krátké sledování a málo událostí, ale je to připomínka, že kauzalita není uzavřená. A ještě jedna věc: možná měříme špatnou veličinu.`);
}

// ============================================================
// SLIDE 9 – Hypoxic burden
// ============================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Observační evidence" });
  s.addText("AHI není všechno: hypoxic burden a hypoxemické metriky", { placeholder: "title" });
  chip(s, ["kohorta"]);
  msg(s, "Pro KV mortalitu predikuje hypoxic burden tam, kde AHI selhává; pro diabetes predikují T90, průměrná SpO₂ a ODI lépe než AHI – prospektivní evidence pro hypoxic burden a DM2 je zatím jediná studie.");
  // schéma desaturační křivky
  const bx = 0.5, by = 2.1, bw = 4.3, bh = 2.3;
  s.addShape(pres.ShapeType.roundRect, { x: bx, y: by, w: bw, h: bh, fill: { color: HX.lt2 }, line: { color: HX.lt2 }, rectRadius: 0.1, objectName: "hb-box" });
  s.addText("Hypoxic burden = plocha pod desaturační křivkou vázaná na respirační události (hloubka × trvání × frekvence), %min/h", { x: bx + 0.15, y: by + 0.1, w: bw - 0.3, h: 0.7, fontSize: 10.5, bold: true, color: HX.dk1, margin: 0, isTextBox: true, objectName: "hb-t" });
  // baseline line + dips
  const ly = by + 1.2;
  s.addShape(pres.ShapeType.line, { x: bx + 0.3, y: ly, w: bw - 0.6, h: 0, line: { color: HX.accent4, width: 1, dashType: "dash" }, objectName: "spo2-base" });
  s.addText("SpO₂ před událostí", { x: bx + 0.3, y: ly - 0.25, w: 2, h: 0.2, fontSize: 8, color: HX.accent4, margin: 0, isTextBox: true, objectName: "spo2-l" });
  [[0.3, 0.8, 0.35], [1.5, 0.5, 0.2], [2.4, 1.1, 0.5]].forEach((d, i) => {
    s.addShape(pres.ShapeType.triangle, { x: bx + 0.3 + d[0], y: ly, w: d[1], h: d[2], fill: { color: HX.accent1, transparency: 30 }, line: { color: HX.accent1, width: 0.75 }, flipV: true, objectName: "dip-" + i });
  });
  s.addText("hloubka", { x: bx + 0.3 + 2.4 + 1.12, y: ly + 0.05, w: 0.6, h: 0.2, fontSize: 7.5, color: HX.accent1, margin: 0, isTextBox: true, objectName: "dip-l1" });
  s.addText("trvání", { x: bx + 0.3 + 2.4 + 0.3, y: ly + 0.52, w: 0.6, h: 0.2, fontSize: 7.5, color: HX.accent1, margin: 0, isTextBox: true, objectName: "dip-l2" });
  s.addText("Azarbarzin, Eur Heart J 2019 · Parekh 2024: více definic (HB, SASHB, T90, ODI)", { x: bx + 0.15, y: by + bh - 0.35, w: bw - 0.3, h: 0.3, fontSize: 8.5, italic: true, color: HX.accent4, margin: 0, isTextBox: true, objectName: "hb-c" });
  forest(s, {
    x: 5.0, y: 2.05, w: 7.85, h: 4.3, scale: "log", xmin: 0.5, xmax: 12, ref: 1, ticks: [0.5, 1, 2, 4, 8], xlabel: "HR / IRR (log škála)",
    head: ["Studie · expozice (outcome)", "HR / IRR (95 % CI)"], labelW: 3.9, valueW: 1.9,
    rows: [
      { label: "Azarbarzin 2019 · hypoxic burden (KV MORTALITA, ne DM)", group: true },
      { label: "   MrOS, n 2 743 · 4. kvintil vs. nižší", est: 1.81, lo: 1.25, hi: 2.62, color: HX.accent1, val: "1,81 (1,25–2,62)" },
      { label: "   MrOS · 5. kvintil", est: 2.73, lo: 1.71, hi: 4.36, color: HX.accent1, val: "2,73 (1,71–4,36)" },
      { label: "   SHHS, n 5 111 · 5. kvintil (AHI nepredikoval)", est: 1.96, lo: 1.11, hi: 3.43, color: HX.accent1, val: "1,96 (1,11–3,43)" },
      { label: "Incidentní DM2 – hypoxemické metriky", group: true },
      { label: "   Appleton 2025 · MAILES, n 536 mužů, 8,3 let · nejnižší kvartil průměrné SpO₂", est: 4.2, lo: 1.7, hi: 10.3, color: HX.dk2, val: "IRR 4,2 (1,7–10,3)" },
      { label: "   Appleton 2025 · kontinuální HBI (na jednotku); kategorický HBI/AHI/ODI NS", est: 1.015, lo: 1.006, hi: 1.024, color: HX.dk2, val: "IRR 1,015 (1,006–1,024)" },
      { label: "   Wojeck 2023 · DREAM, n 650 · T90 (na 1 %); AHI bez asociace", est: 1.009, lo: 1.001, hi: 1.017, color: HX.dk2, val: "HR 1,009 (1,001–1,017)" },
      { label: "   Wojeck 2023 · průměrná noční SpO₂ (na 1 %)", est: 0.914, lo: 0.857, hi: 0.975, color: HX.accent2, val: "HR 0,914 (0,857–0,975)" },
    ],
  });
  box(s, 0.5, 4.55, 4.3, 1.8, "Další kohorty: hypoxémie > AHI", [
    "Toronto, n 8 678: AHI > 30 ≈ +30 % hazard DM; nezávisle i T90 a REM-AHI (Kendzerska 2014; CI neuvedeno)",
    "Uppsala, muži, 11 let: ODI > 5 → DM OR 4,4 (1,1–18,1) (Lindberg 2012)",
    "MA hypoxic burden → DM: neexistuje",
  ], HX.dk1, { fontSize: 10 });
  cite(s, "Azarbarzin, Eur Heart J 2019 · Appleton, J Sleep Res 2025 · Wojeck, JCSM 2023 · Kendzerska, AJRCCM 2014 · Lindberg, Chest 2012 — observační; evidence pro HB → DM2 je novější a omezená");
  addNotes(s, 9, `AHI počítá události, ale neříká nic o tom, jak hluboké a dlouhé jsou desaturace. Hypoxic burden je plocha pod desaturační křivkou vázaná na respirační události, měřená od saturace před událostí. Ve dvou kohortách, MrOS a Sleep Heart Health Study, predikoval kardiovaskulární mortalitu tam, kde AHI nepredikoval: v nejvyšších kvintilech HR 1,81 a 2,73 u starších mužů a 1,96 v SHHS (Azarbarzin 2019). Zdůrazňuji, že tam je outcome kardiovaskulární mortalita, nikoli diabetes. Pro diabetes je obraz podobný, ale evidence je čerstvá a omezená. V torontské kohortě 8678 pacientů měli ti s AHI nad 30 o 30 procent vyšší hazard diabetu a nezávisle predikoval i čas se saturací pod 90 procent (Kendzerska 2014). U 650 veteránů predikovaly incidentní diabetes T90 a průměrná noční saturace, zatímco AHI ne (Wojeck 2023). V australské mužské kohortě byl nejnižší kvartil průměrné saturace spojen s 4,2násobným rizikem diabetu a kontinuální hypoxic burden byl signifikantní, kategorické AHI ani ODI nikoli; šlo ale jen o 52 událostí (Appleton 2025). V Uppsale predikovalo ODI nad 5 diabetes s OR 4,4 u 141 mužů (Lindberg 2012). Metaanalýza pro hypoxic burden a diabetes neexistuje. Praktické sdělení: až budete číst polygrafický nález, dívejte se i na ODI, T90 a průměrnou saturaci, ne jen na AHI. Ale je to léčitelné? Přejdeme k intervencím.`);
}

// ============================================================
// SLIDE 10 – CPAP nediabetici
// ============================================================
pres.addSection({ title: "Intervence" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Intervence" });
  s.addText("Intervence I: CPAP a inzulinová rezistence u nediabetiků", { placeholder: "title" });
  chip(s, ["MA RCT", "RCT"]);
  msg(s, "Metaanalýzy RCT ukazují malý pokles HOMA-IR (−0,4 až −0,55) bez změny glykémie; IPD MA sham-RCT a RCT s FSIVGTT efekt nenašly ani u adherentních.");
  forest(s, {
    x: 0.5, y: 2.05, w: 8.4, h: 3.2, scale: "lin", xmin: -1.2, xmax: 0.4, ref: 0, ticks: [-1, -0.5, 0], xlabel: "Rozdíl HOMA-IR, CPAP vs. sham/kontrola (MD)",
    head: ["Metaanalýza RCT · nediabetici", "MD (95 % CI); I²"], labelW: 3.9, valueW: 2.2,
    rows: [
      { label: "Iftikhar 2013 · 5 RCT, n 244", est: -0.44, lo: -0.82, hi: -0.06, color: HX.accent2, val: "−0,44 (−0,82 až −0,06); I² 0 %" },
      { label: "Iftikhar 2015 · aktualizace (n NEOVĚŘENO)", est: -0.43, lo: -0.75, hi: -0.11, color: HX.accent2, val: "−0,43 (−0,75 až −0,11)" },
      { label: "Yang 2013 · 12 studií, 3–24 týdnů", est: -0.55, lo: -0.91, hi: -0.20, color: HX.accent2, val: "−0,55 (−0,91 až −0,20)" },
      { label: "Abud 2019 · 9 RCT, n 443, GRADE low", est: -0.39, lo: -0.69, hi: -0.08, color: HX.accent2, val: "−0,39 (−0,69 až −0,08); I² 57 %" },
    ],
  });
  s.addText("Glykémie nalačno: bez změny ve všech MA (Iftikhar 2013; Yang 2013; Abud 2019). Viscerální tuk a adiponektin beze změny (Iftikhar 2015).", { x: 0.5, y: 5.35, w: 8.4, h: 0.5, fontSize: 10.5, color: HX.dk2, margin: 0, isTextBox: true, objectName: "fbg" });
  box(s, 9.15, 2.05, 3.7, 2.1, "Proti: IPD MA sham-RCT (Hoyos 2019)", [
    "3 sham-kontrolované RCT, n 128, 2–3 měsíce",
    "Ani CPAP, ani dávka CPAP nezměnily žádný metabolický parametr",
    "Hmotnost +1,17 kg (0,37–1,97); +0,30 kg na každou hodinu/noc",
  ], HX.accent1, { fontSize: 10, fill: "FBEDE9" });
  box(s, 9.15, 4.3, 3.7, 2.05, "Proti: RCT s FSIVGTT (Chirinos 2014)", [
    "181 obézních s AHI ≥ 15, bez DM; CPAP vs. redukce hmotnosti vs. obojí, 24 týdnů",
    "Inzulinová rezistence klesla jen s úbytkem hmotnosti – ne u CPAP samotného, ani u adherentních (4,0 h/noc)",
  ], HX.accent1, { fontSize: 10, fill: "FBEDE9" });
  cite(s, "Iftikhar, Ann Am Thorac Soc 2013 · Iftikhar, JCSM 2015 · Yang, Sleep Breath 2013 · Abud, Sleep Med 2019 · Hoyos, J Sleep Res 2019 · Chirinos, N Engl J Med 2014 — adherence v MA neuvedena");
  addNotes(s, 10, `Pokud apnoe zhoršuje inzulinovou rezistenci, měl by ji CPAP zlepšit. U nediabetiků to metaanalýzy randomizovaných studií skutečně naznačují. HOMA-IR klesl o 0,44 v pěti RCT bez heterogenity (Iftikhar 2013), o 0,43 v aktualizaci, bez změny viscerálního tuku a adiponektinu (Iftikhar 2015), o 0,55 ve dvanácti studiích (Yang 2013) a o 0,39 v devíti RCT s heterogenitou 57 procent a nízkou jistotou evidence (Abud 2019). Glykémie nalačno se nezměnila v žádné z nich. Je to klinicky významné? Pokles HOMA o 0,4 autoři přirovnávají k efektu thiazolidindionů u nediabetiků, ale to je interpretace, ne měření. A jsou tu dvě studie, které výsledek zpochybňují. Metaanalýza individuálních dat ze tří sham-kontrolovaných RCT u 128 osob nenašla žádný efekt CPAP ani jeho dávky na metabolické parametry, zato přírůstek hmotnosti 1,17 kilogramu (Hoyos 2019). A v RCT se 181 obézními pacienty měřila inzulinová senzitivita metodou FSIVGTT: zlepšila se jen v ramenech s redukcí hmotnosti, nikoli u samotného CPAP, a to ani u adherentních, při průměrné adherenci 4 hodiny (Chirinos 2014). Zlepšení HOMA-IR je tedy malé a ne zcela konzistentní. A u diabetiků, kde by na tom záleželo nejvíc?`);
}

// ============================================================
// SLIDE 11 – CPAP a HbA1c u DM2
// ============================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Intervence" });
  s.addText("Intervence II: CPAP a HbA1c u DM2 – rozporné metaanalýzy, negativní velká RCT", { placeholder: "title" });
  chip(s, ["MA RCT", "RCT"]);
  msg(s, "Starší MA negativní, Herth 2023 pozitivní (−0,24 %, I² 71 %, publikační bias), nejnovější a největší Wu 2026 negativní pro HbA1c ve všech délkách – HOMA-IR klesá jen při léčbě ≥ 24 týdnů.");
  forest(s, {
    x: 0.5, y: 2.05, w: 8.6, h: 3.75, scale: "lin", xmin: -1.0, xmax: 0.4, ref: 0, ticks: [-1, -0.5, 0], xlabel: "Rozdíl HbA1c (%), CPAP vs. kontrola (MD)",
    head: ["Metaanalýza RCT · DM2", "MD % (95 % CI); I²"], labelW: 4.0, valueW: 2.3,
    rows: [
      { label: "Zhu 2018 · 6 RCT, n 496", est: 0.05, lo: -0.14, hi: 0.24, color: HX.accent4, val: "+0,05 (−0,14 až 0,24) NS" },
      { label: "Labarca 2018 · 6 RCT, n 581; ≥ 4 h subanalýza též NS", est: -0.10, lo: -0.25, hi: 0.04, color: HX.accent4, val: "−0,10 (−0,25 až 0,04) NS" },
      { label: "Herth 2023 · 9 RCT, n 807 (Egger p = 0,039)", est: -0.24, lo: -0.43, hi: -0.06, color: HX.accent2, big: true, val: "−0,24 (−0,43 až −0,06); I² 71 %" },
      { label: "Herth 2023 · bez outlieru (Zhao 2022, 8,3 h/noc)", est: -0.16, lo: -0.28, hi: -0.03, color: HX.accent2, val: "−0,16 (−0,28 až −0,03); I² 39 %" },
      { label: "Wu 2026 · 14 RCT, n 1 273 · léčba 12–< 24 týdnů (4 RCT)", est: -0.05, lo: -0.24, hi: 0.13, color: HX.accent4, val: "−0,05 (−0,24 až 0,13); I² 0 %" },
      { label: "Wu 2026 · léčba ≥ 24 týdnů (4 RCT)", est: -0.15, lo: -0.47, hi: 0.16, color: HX.accent4, val: "−0,15 (−0,47 až 0,16); I² 72 %" },
    ],
  });
  s.addText("Shang 2021 (7 RCT, n 691): HbA1c SMD −0,32 (−0,60 až −0,03) – jiná škála, proto není v grafu. Wu 2026: HOMA-IR při léčbě ≥ 24 týdnů (2 RCT) MD −1,67 (−2,57 až −0,77), I² 0 %; BMI +0,74 kg/m² (0,20–1,27); adherenci nebylo možné adjustovat.", { x: 0.5, y: 5.85, w: 8.6, h: 0.7, fontSize: 10, color: HX.dk2, margin: 0, isTextBox: true, objectName: "shang" });
  box(s, 9.35, 2.05, 3.5, 4.45, "Velká pragmatická RCT: negativní", [
    "Shaw 2016: n 298, DM2 s HbA1c 6,5–8,5 %, 6 měsíců, bez sham – žádný rozdíl HbA1c ani u adherentních, horší kontroly či těžší OSA; DTK −3,5 vs. −1,5",
    "Loffler 2020 (SAVE substudie): n 888, KV pacienti, medián 4,3 roku – bez rozdílu v glukóze, HbA1c ani nových diagnózách DM",
    "Aurora 2023 (HYPNOS): n 184, CGM 3 měsíce, adherence 5,4 h – bez rozdílu ve variabilitě glykémie",
    "Banghøj 2021: n 72, CGM 12 týdnů, 5,6 h – negativní i v subanalýze ≥ 4 h",
  ], HX.dk1, { fontSize: 10 });
  cite(s, "Zhu, Sleep Breath 2018 · Labarca, Clin Respir J 2018 · Shang, Diabetes Obes Metab 2021 · Herth, Eur Respir Rev 2023 · Wu, BMJ Open Respir Res 2026 · Shaw, AJRCCM 2016 · Loffler, Diabetes Care 2020 · Aurora, Chest 2023 · Banghøj, Endocrinol Diabetes Metab 2021");
  addNotes(s, 11, `U diabetiků je otázka klinicky nejzajímavější a odpověď nejrozpornější. Dvě starší metaanalýzy RCT jsou negativní: HbA1c plus 0,05 (Zhu 2018) a minus 0,10 procenta, bez efektu i v podskupině s adherencí nad 4 hodiny (Labarca 2018). Pak přišla pozitivní metaanalýza 11 RCT s 964 pacienty: HbA1c minus 0,24 procenta, ale s heterogenitou 71 procent a známkami publikačního bias; po vyloučení jedné odlehlé studie s adherencí 8,3 hodiny byl efekt minus 0,16 procenta (Herth 2023). A letošní největší metaanalýza 14 RCT s 1273 pacienty je pro HbA1c negativní ve všech délkách léčby, zato HOMA-IR při léčbě nad 24 týdnů klesl o 1,67 ve dvou RCT a BMI stoupl o 0,74 (Wu 2026); autoři sami píší, že adherenci nemohli adjustovat. Velká pragmatická RCT to potvrzují: 298 diabetiků s dobře kontrolovaným HbA1c po šesti měsících bez rozdílu, ani u adherentních (Shaw 2016); substudie SAVE po 4,3 letech bez rozdílu v HbA1c ani v nových diagnózách diabetu (Loffler 2020); kontinuální monitorace glykémie u 184 pacientů bez rozdílu při adherenci 5,4 hodiny (Aurora 2023) a u 72 pacientů negativní i v subanalýze adherentních (Banghøj 2021). Můj závěr: CPAP sám u běžného diabetika HbA1c klinicky významně nesníží. Proč se tedy některé studie liší?`);
}

// ============================================================
// SLIDE 12 – Adherence jako dávka
// ============================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Intervence" });
  s.addText("Adherence jako dávka: proč se RCT rozcházejí", { placeholder: "title" });
  chip(s, ["RCT", "MA RCT"]);
  msg(s, "Jediné clampové RCT (3,6 h/noc) je negativní; pozitivní signál dávají protokoly s ~8 h/noc a meta-regrese – metabolismus zřejmě potřebuje vyšší dávku než symptomy.");
  // dose axis
  const ax = 0.9, ay = 4.05, aw = 11.5, hmin = 2, hmax = 9;
  const hx = h => ax + (h - hmin) / (hmax - hmin) * aw;
  s.addShape(pres.ShapeType.line, { x: ax, y: ay, w: aw, h: 0, line: { color: HX.accent4, width: 1.5 }, objectName: "dose-axis" });
  for (let h = 2; h <= 9; h++) { s.addShape(pres.ShapeType.line, { x: hx(h), y: ay, w: 0, h: 0.08, line: { color: HX.accent4, width: 1 }, objectName: "dt" + h }); s.addText(h === 9 ? "9 h/noc CPAP" : h + " h", { x: hx(h) - 0.5, y: ay + 0.1, w: 1.0, h: 0.22, fontSize: 9, color: HX.accent4, align: "center", margin: 0, isTextBox: true, objectName: "dtl" + h }); }
  s.addShape(pres.ShapeType.rect, { x: hx(4), y: 2.05, w: 0.02, h: 2.0, fill: { color: HX.accent3 }, line: { color: HX.accent3 }, objectName: "thr4" });
  s.addText("„≥ 4 h“ – obvyklý práh adherence", { x: hx(4) + 0.08, y: 2.0, w: 2.2, h: 0.25, fontSize: 8.5, italic: true, color: HX.accent4, margin: 0, isTextBox: true, objectName: "thr4l" });
  // lvl: 1/2 = nad osou (blíž/dál), -1/-2 = pod osou
  const pts = [
    { h: 2.5, lvl: -1, col: HX.accent4, t: "Lam 2017 · DM2, n 64, 3 měs.\nHbA1c NS (2,5 h*)" },
    { h: 3.6, lvl: 1, col: HX.accent4, t: "West 2007 · DM2, sham, n 42, 3 měs.\nCLAMP negativní (3,6 h)" },
    { h: 4.9, lvl: -1, dx: -0.55, col: HX.accent4, t: "Shaw 2016 · DM2, n 298, 6 měs.\nHbA1c NS i u adherentních (4,9 h*)" },
    { h: 5.2, lvl: 2, col: HX.accent2, t: "Martínez-Cerón 2016 · DM2, n 50, HbA1c 7,6 %\nHbA1c ↓, HOMA ↓ (5,2 h*)" },
    { h: 5.4, lvl: -2, dx: 0.3, col: HX.accent4, t: "Aurora 2023 · DM2, CGM, n 184\nNS (5,4 h)" },
    { h: 5.6, lvl: 1, dx: 0.6, col: HX.accent4, t: "Banghøj 2021 · DM2, CGM, n 72\nNS i ≥ 4 h (5,6 h*)" },
    { h: 7.0, lvl: -1, col: HX.accent2, t: "Chasens 2022 · DM2, sham, n 98\nITT NS; ≥ 7 h → HbA1c ↓ (b 0,54; p = 0,001)" },
    { h: 7.9, lvl: 1, col: HX.accent2, t: "Mokhlesi 2017 · DM2, in-lab 1 týden\n24h glykémie −13,5 mg/dl (7,9 h*)" },
    { h: 8.0, lvl: -2, col: HX.accent2, t: "Pamidi 2015 · prediabetes, in-lab 2 týdny\nAUC glukózy ↓, inz. senzitivita ↑ (8 h předepsáno)" },
  ];
  const TB = 0.7; // výška textového bloku
  pts.forEach((p, i) => {
    const x = hx(p.h);
    s.addShape(pres.ShapeType.ellipse, { x: x - 0.11, y: ay - 0.11, w: 0.22, h: 0.22, fill: { color: p.col }, line: { color: HX.lt1, width: 1 }, objectName: "dp" + i });
    let ty, ly1, lh;
    if (p.lvl === 1) { ty = ay - 0.25 - TB; ly1 = ay - 0.25; lh = 0.14; }
    else if (p.lvl === 2) { ty = ay - 1.15 - TB; ly1 = ay - 1.15; lh = 1.04; }
    else if (p.lvl === -1) { ty = ay + 0.4; ly1 = ay + 0.11; lh = 0.29; }
    else { ty = ay + 1.3; ly1 = ay + 0.11; lh = 1.19; }
    s.addShape(pres.ShapeType.line, { x, y: ly1, w: 0, h: lh, line: { color: p.col, width: 0.75, dashType: Math.abs(p.lvl) === 2 ? "dash" : "solid" }, objectName: "dl" + i });
    s.addText(p.t, { x: x - 1.05 + (p.dx || 0), y: ty, w: 2.1, h: TB, fontSize: 8.5, color: p.col === HX.accent4 ? HX.dk2 : HX.accent2, align: "center", valign: p.lvl > 0 ? "bottom" : "top", margin: 0, isTextBox: true, objectName: "dtx" + i });
  });
  s.addText("Meta-regrese (Herth 2023, 11 RCT, n 964): každá hodina nočního užívání navíc → HbA1c −0,17 % (−0,31 až −0,04; p = 0,010); vyšší výchozí HbA1c → větší pokles (−0,34; p = 0,001).", { x: 0.5, y: 6.1, w: 8.6, h: 0.45, fontSize: 10, bold: true, color: HX.accent2, margin: 0, isTextBox: true, valign: "middle", objectName: "metareg" });
  s.addText("● pozitivní  ● negativní / NS   * adherence dle tab. 1 v MA Herth 2023 / Wu 2026 (sekundární zdroj)", { x: 9.3, y: 6.1, w: 3.55, h: 0.45, fontSize: 8, color: HX.accent4, margin: 0, isTextBox: true, valign: "middle", objectName: "leg" });
  cite(s, "West, Thorax 2007 · Lam, Sleep Breath 2017 · Shaw, AJRCCM 2016 · Martínez-Cerón, AJRCCM 2016 · Banghøj 2021 · Aurora, Chest 2023 · Chasens, Endocr Pract 2022 · Mokhlesi, Diabetes Obes Metab 2017 · Pamidi, AJRCCM 2015 · Herth, Eur Respir Rev 2023");
  addNotes(s, 12, `Odpověď se zřejmě jmenuje dávka. Jediné randomizované studie, která měřila inzulinovou rezistenci zlatým standardem, hyperinzulinemickým euglykemickým clampem, měla sham-kontrolu, 42 diabetiků a tři měsíce léčby a byla negativní, při adherenci 3,6 hodiny za noc (West 2007). Největší RCT bylo negativní při 4,9 hodinách i mezi adherentními (Shaw 2016; adherence dle metaanalýzy Herth 2023). Pozitivní španělské RCT s 50 pacienty mělo 5,2 hodiny, ale vyšší výchozí HbA1c kolem 7,6 procenta a jediné centrum (Martínez-Cerón 2016). Sham-kontrolované RCT u 98 diabetiků bylo v ITT negativní, ale pokles HbA1c rostl s kumulativními hodinami a užívání nad 7 hodin za noc bylo spojeno se signifikantním poklesem (Chasens 2022). A obě laboratorní studie, kde byl CPAP zajištěn celou osmihodinovou noc, byly pozitivní: u prediabetiků se za dva týdny zlepšila glukózová tolerance i inzulinová senzitivita (Pamidi 2015) a u diabetiků klesla 24hodinová glykémie o 13,5 miligramů na decilitr za jediný týden (Mokhlesi 2017). Meta-regrese to shrnuje: každá hodina nočního užívání navíc znamená o 0,17 procenta HbA1c méně (Herth 2023). Hypotéza tedy zní, že metabolický efekt vyžaduje výrazně vyšší dávku než efekt na spavost; práh 4 hodiny, který používáme pro adherenci, ho zpravidla neodhalí. Pokud je ale cílem metabolismus, máme silnější nástroj.`);
}

// ============================================================
// SLIDE 13 – Hmotnost vs CPAP
// ============================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Intervence" });
  s.addText("Intervence III: redukce hmotnosti je metabolicky silnější než CPAP", { placeholder: "title" });
  chip(s, ["RCT", "MA RCT", "MA obs."]);
  msg(s, "V RCT CPAP vs. redukce hmotnosti vs. obojí klesly CRP, inzulinová rezistence a TG jen s úbytkem hmotnosti; hubnutí snižuje AHI o 6–23/h, bariatrie o ~38/h – zpravidla bez normalizace.");
  s.addChart(pres.ChartType.bar, [
    { name: "Změna hmotnosti (kg)", labels: ["CPAP", "Redukce hmotnosti", "Obojí"], values: [0, -6.8, -7.0] },
  ], Object.assign({}, chartBase, {
    x: 0.5, y: 2.0, w: 3.2, h: 2.6, barDir: "col", barGapWidthPct: 50, chartColors: [HX.accent4, HX.accent2, HX.dk2],
    dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", valAxisMinVal: -8, valAxisMaxVal: 1,
    showTitle: true, title: "Chirinos 2014 · hmotnost, 24 týdnů (kg)", titleFontSize: 10, titleColor: HX.dk2,
  }));
  s.addChart(pres.ChartType.bar, [
    { name: "Změna STK (mm Hg)", labels: ["CPAP", "Redukce hmotnosti", "Obojí"], values: [-3.0, -6.8, -14.1] },
  ], Object.assign({}, chartBase, {
    x: 3.9, y: 2.0, w: 3.2, h: 2.6, barDir: "col", barGapWidthPct: 50, chartColors: [HX.accent4, HX.accent2, HX.dk2],
    dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", valAxisMinVal: -16, valAxisMaxVal: 1,
    showTitle: true, title: "Chirinos 2014 · STK per-protokol (mm Hg)", titleFontSize: 10, titleColor: HX.dk2,
  }));
  s.addText("Chirinos 2014 (NEJM): 181 obézních, AHI ≥ 15, CRP > 1, bez DM. CRP, inzulinová rezistence (FSIVGTT) a TG klesly pouze v ramenech s redukcí hmotnosti; CPAP adherence 4,0 h/noc. Per-protokol TK = nerandomizovaná analýza.", { x: 0.5, y: 4.65, w: 6.6, h: 0.75, fontSize: 10, color: HX.dk2, margin: 0, isTextBox: true, objectName: "chir" });
  forest(s, {
    x: 7.35, y: 2.05, w: 5.5, h: 3.35, scale: "lin", xmin: -45, xmax: 5, ref: 0, ticks: [-40, -30, -20, -10, 0], xlabel: "Změna AHI (událostí/h) vs. kontrola / před",
    head: ["Redukce hmotnosti → AHI", "Δ AHI (95 % CI)"], labelW: 2.45, valueW: 1.6,
    rows: [
      { label: "Foster 2009 · ILI vs. edukace, DM2, 1 rok (RCT)", est: -9.7, lo: -9.7, hi: -9.7, color: HX.accent2, val: "−9,7 (SE 2,0)" },
      { label: "Johansson 2009 · VLED 9 týdnů, muži (RCT)", est: -23, lo: -30, hi: -15, color: HX.accent2, val: "−23 (−30 až −15)" },
      { label: "Araghi 2013 · MA 7 RCT, lifestyle, I² 86 %", est: -6.04, lo: -11.18, hi: -0.90, color: HX.accent2, val: "−6,0 (−11,2 až −0,9)" },
      { label: "Mitchell 2014 · MA 4 RCT, intenzivní lifestyle", est: -16.09, lo: -25.64, hi: -6.54, color: HX.accent2, val: "−16,1 (−25,6 až −6,5)" },
      { label: "Greenburg 2009 · MA před/po bariatrii, 342 pac.", est: -38.2, lo: -44.4, hi: -31.9, color: HX.dk2, val: "−38,2 (−44,4 až −31,9)" },
      { label: "Dixon 2012 · bandáž vs. konzervativní, RCT, 2 roky", est: -11.5, lo: -28.3, hi: 5.3, color: HX.accent4, dim: true, val: "−11,5 (−28,3 až 5,3) NS" },
    ],
  });
  s.addText("Efekt ILI na AHI přetrval 4 roky i přes ~50 % regain hmotnosti; remise OSA 20,7 % vs. 3,6 % (Kuna 2013). Po bariatrii reziduální AHI 15,8/h – „should not expect a cure“ (Greenburg 2009).", { x: 7.35, y: 5.45, w: 5.5, h: 0.95, fontSize: 10, color: HX.dk2, margin: 0, isTextBox: true, objectName: "kuna" });
  cite(s, "Chirinos, N Engl J Med 2014 · Foster, Arch Intern Med 2009 · Kuna, Sleep 2013 · Johansson, BMJ 2009 · Araghi, Sleep 2013 · Mitchell, Sleep Med 2014 · Greenburg, Am J Med 2009 · Dixon, JAMA 2012");
  addNotes(s, 13, `Tím nástrojem je hmotnost. V randomizované studii, kterou jsem už zmínil, bylo 181 obézních pacientů s apnoí rozděleno k CPAP, k programu redukce hmotnosti nebo k obojímu na 24 týdnů: hmotnost klesla o 6,8 a 7,0 kilogramu v ramenech s redukcí a vůbec ne u CPAP; CRP, inzulinová rezistence a triglyceridy klesly jen tam, kde pacienti hubli; tlak klesl ve všech ramenech, v per-protokolové analýze nejvíce u kombinace, o 14,1 mm Hg (Chirinos 2014). Opačným směrem: hubnutí snižuje apnoi. Intenzivní životní styl u obézních diabetiků snížil hmotnost o 10,8 kilogramu a AHI o 9,7 události za hodinu proti kontrole (Foster 2009) a efekt na AHI přetrval čtyři roky i přes zpětný nárůst poloviny hmotnosti, s remisí apnoe u 20,7 oproti 3,6 procenta (Kuna 2013). Devítitýdenní nízkoenergetická dieta u mužů snížila hmotnost o 20 kilogramů a AHI o 23 (Johansson 2009). Metaanalýzy dávají pokles AHI o 16 událostí při intenzivní intervenci (Mitchell 2014), ale jen o 6 v sedmi RCT s heterogenitou 86 procent (Araghi 2013). Bariatrie sníží AHI z 55 na 16, což je stále středně těžká apnoe (Greenburg 2009), a v jediném RCT nebyl rozdíl proti konzervativní léčbě signifikantní, byť při malém vzorku (Dixon 2012). Hubnutí tedy pomáhá metabolicky i na apnoi, ale zřídka ji vyléčí. A tím se dostáváme k inkretinům.`);
}

// ============================================================
// SLIDE 14 – Inkretiny
// ============================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Intervence" });
  s.addText("Intervence IV: inkretiny u OSA – SCALE, SURMOUNT-OSA, TRIUMPH-OSA", { placeholder: "title" });
  chip(s, ["RCT", "MA RCT"]);
  msg(s, "Tirzepatid snížil AHI o 20–24/h proti placebu za 52 týdnů, hmotnost o 16–17 % a HOMA-IR o ~50 %; u DM2 je efekt retatrutidu menší; chybí tvrdé KV outcomy i přímé srovnání s CPAP.");
  forest(s, {
    x: 0.5, y: 2.05, w: 8.4, h: 4.3, scale: "lin", xmin: -40, xmax: 5, ref: 0, ticks: [-40, -30, -20, -10, 0], xlabel: "Rozdíl změny AHI vs. placebo (událostí/h)",
    head: ["RCT · lék · populace", "Δ AHI vs. placebo (95 % CI)"], labelW: 4.1, valueW: 1.9,
    rows: [
      { label: "SCALE Sleep Apnea (Blackman 2016) · liraglutid 3 mg, 32 t, n 359, bez DM", est: -6.1, lo: -11.0, hi: -1.2, color: HX.dk2, val: "−6,1 (−11,0 až −1,2)" },
      { label: "SURMOUNT-OSA 1 (Malhotra 2024) · tirzepatid, 52 t, bez PAP, n 234", est: -20.0, lo: -25.8, hi: -14.2, color: HX.accent2, big: true, val: "−20,0 (−25,8 až −14,2)" },
      { label: "SURMOUNT-OSA 2 (Malhotra 2024) · tirzepatid, s PAP, n 235", est: -23.8, lo: -29.6, hi: -17.9, color: HX.accent2, big: true, val: "−23,8 (−29,6 až −17,9)" },
      { label: "TRIUMPH-1-OSA (Grunstein 2026) · retatrutid 9 mg, 80 t, bez DM2", est: -24.4, lo: -34.1, hi: -14.7, color: HX.accent2, val: "−24,4 (−34,1 až −14,7)" },
      { label: "TRIUMPH-2-OSA (Grunstein 2026) · retatrutid 9 mg, 80 t, s DM2", est: -14.4, lo: -21.6, hi: -7.2, color: HX.accent1, val: "−14,4 (−21,6 až −7,2)" },
      { label: "MA 6 RCT GLP-1 RA (Carvalhal 2026) · n 1 069", est: -9.99, lo: -16.47, hi: -3.50, color: HX.dk2, diamond: true, val: "−10,0 (−16,5 až −3,5); I² 87 %" },
    ],
  });
  box(s, 9.15, 2.05, 3.7, 4.3, "SURMOUNT-OSA: kardiometabolické výsledky (Nat Med 2026, sekundární analýza)", [
    "Hmotnost −16,1 % (−18,0 až −14,2) / −17,3 %",
    "Hypoxic burden (SASHB) −70,1 / −61,3 %min/h",
    "STK −7,9 (−11,0 až −4,9) / −4,3 mm Hg (48. t)",
    "hsCRP −28,9 % / −45,1 %",
    "HOMA-IR −48,0 % / −54,5 % (p < 0,001)",
    "Změny AHI/SASHB nezávisle mediovaly hsCRP, HOMA-IR a TG; STK mediován hmotností",
    "Limity: efficacy estimand, 52 týdnů, bez DM, bez tvrdých KV outcomes; efekt vs. placebo až od 20. týdne (Malhotra 2025)",
  ], HX.accent2, { fontSize: 9.5, titleSize: 11 });
  cite(s, "Blackman, Int J Obes 2016 · Malhotra, N Engl J Med 2024 · Malhotra, Nat Med 2026 · Malhotra, Sleep Med 2025 · Grunstein, N Engl J Med 2026 · Carvalhal, Sleep Breath 2026 · Pack, Sleep 2025 (fórum) — AASM guideline k tirzepatidu u OSA k 10/2026 nenalezena");
  addNotes(s, 14, `Farmakoterapie obezity změnila i spánkovou medicínu. Liraglutid 3 miligramy snížil za 32 týdnů AHI o 12,2 oproti 6,1 na placebu, rozdíl 6,1 události, při rozdílu hmotnosti 4,2 procenta (Blackman 2016). SURMOUNT-OSA s tirzepatidem je první velké RCT s AHI jako primárním cílem: u 469 obézních nediabetiků se středně těžkou až těžkou apnoí klesl AHI za 52 týdnů o 20,0 události za hodinu proti placebu u pacientů bez PAP a o 23,8 u pacientů na PAP (Malhotra 2024). Sekundární analýza ukazuje hmotnost minus 16 až 17 procent, hypoxic burden minus 61 až 70 procent-minut za hodinu, systolický tlak minus 7,9 a 4,3 mm Hg, hsCRP minus 29 až 45 procent a HOMA-IR minus 48 až 55 procent, přičemž změny apnoických metrik nezávisle mediovaly část efektu na hsCRP, HOMA-IR a triglyceridy, zatímco tlak byl mediován hmotností (Malhotra 2026). Letošní studie s retatrutidem ukazují podobný efekt u nediabetiků, rozdíl AHI až minus 24, ale menší u diabetiků, minus 14 (Grunstein 2026). Metaanalýza šesti RCT s liraglutidem a tirzepatidem dává pokles AHI o 10 s heterogenitou 87 procent (Carvalhal 2026). Limity pro kardiologa: efekt proti placebu nastupuje až po 20 týdnech, zatímco CPAP působí okamžitě (Malhotra 2025); chybí tvrdé kardiovaskulární outcomy u populace s OSA i přímé srovnání s CPAP a žádná AASM guideline k tirzepatidu u apnoe zatím není. Teď k tomu, co kardiology zajímá nejvíc: tvrdé outcomy CPAP.`);
}

// ============================================================
// SLIDE 15 – SAVE/ISAACC/RICCADSA
// ============================================================
pres.addSection({ title: "Tvrdé KV outcomy" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Tvrdé KV outcomy" });
  s.addText("Tvrdé KV outcomy CPAP: SAVE, ISAACC, RICCADSA", { placeholder: "title" });
  chip(s, ["RCT"]);
  msg(s, "Tři RCT u KV pacientů (převážně nespavých) s adherencí 2,8–3,3 h/noc jsou v ITT neutrální; benefit se objevuje jen v nerandomizovaných on-treatment analýzách.");
  forest(s, {
    x: 0.5, y: 2.05, w: 12.35, h: 3.3, scale: "log", xmin: 0.1, xmax: 2, ref: 1, ticks: [0.1, 0.25, 0.5, 1, 2], xlabel: "HR primárního složeného KV endpointu, CPAP vs. usual care (log škála)",
    head: ["RCT · populace · follow-up · adherence", "HR (95 % CI)"], labelW: 6.0, valueW: 2.2,
    rows: [
      { label: "SAVE (McEvoy 2016) · n 2 717, ICHS/CMP, 45–75 let, FU 3,7 r · adherence 3,3 h/noc · ITT", est: 1.10, lo: 0.91, hi: 1.32, color: HX.dk2, big: true, val: "1,10 (0,91–1,32); p = 0,34" },
      { label: "ISAACC (Sánchez-de-la-Torre 2020) · n 1 264, nespaví po AKS, FU 3,35 r · 2,78 h/noc · ITT", est: 0.89, lo: 0.68, hi: 1.17, color: HX.dk2, big: true, val: "0,89 (0,68–1,17); p = 0,40" },
      { label: "RICCADSA (Peker 2016) · n 244, po revaskularizaci, ESS < 10, FU 57 měs. · ITT", est: 0.80, lo: 0.46, hi: 1.41, color: HX.dk2, big: true, val: "0,80 (0,46–1,41); p = 0,45" },
      { label: "RICCADSA · adjustovaná on-treatment analýza, CPAP ≥ 4 h vs. < 4 h/bez – NERANDOMIZOVANÉ", est: 0.29, lo: 0.10, hi: 0.86, color: HX.accent1, dim: true, val: "0,29 (0,10–0,86); p = 0,026" },
    ],
  });
  box(s, 0.5, 5.4, 6.0, 1.15, "Co studie ukázaly navíc", [
    "SAVE: AHI 29,0 → 3,7/h; ↓ chrápání a spavost, ↑ kvalita života a nálada; žádný efekt na jednotlivé KV endpointy",
    "ISAACC: pacienti s OSA neměli více příhod než bez OSA – HR 1,01 (0,76–1,35); bez vztahu ke compliance či tíži",
  ], HX.dk1, { fontSize: 10 });
  box(s, 6.75, 5.4, 6.1, 1.15, "Pozor na on-treatment analýzy", [
    "Adherentní pacienti se liší od neadherentních (healthy-adherer bias) – HR 0,29 není důkaz účinku",
    "SAVE propensity-matched ≥ 4 h: čísla z plného textu NEJM nebyla ověřitelná (web blokován) – neuvádím",
  ], HX.accent1, { fontSize: 10, fill: "FBEDE9" });
  cite(s, "McEvoy, N Engl J Med 2016 · Sánchez-de-la-Torre, Lancet Respir Med 2020 · Peker, Am J Respir Crit Care Med 2016");
  addNotes(s, 15, `Tři randomizované studie testovaly, zda CPAP sníží kardiovaskulární příhody u pacientů, které léčíme my. SAVE: 2717 pacientů ve věku 45 až 75 let s koronárním nebo cerebrovaskulárním onemocněním a středně těžkou až těžkou apnoí, průměrně 3,7 roku; primární složený endpoint nastal u 17,0 oproti 15,4 procenta, HR 1,10, při průměrné adherenci 3,3 hodiny za noc; AHI přitom klesl z 29 na 3,7 a zlepšila se spavost i kvalita života (McEvoy 2016). ISAACC: 1264 nespavých pacientů po akutním koronárním syndromu, HR 0,89, adherence 2,78 hodiny; a pacienti s apnoí neměli více příhod než pacienti bez apnoe, HR 1,01 (Sánchez-de-la-Torre 2020). RICCADSA: 244 pacientů po revaskularizaci bez denní spavosti, v ITT HR 0,80 a nesignifikantní; v adjustované on-treatment analýze měli pacienti užívající CPAP nad 4 hodiny HR 0,29 (Peker 2016). To poslední číslo je lákavé, ale je nerandomizované: pacienti, kteří přístroj používají, se liší od těch, kteří ne, a to i ve věcech, které neumíme změřit. Pro kardiologa shrnutí: v ITT žádná z těchto studií neprokázala prevenci příhod a CPAP tedy zatím nemůžeme předepisovat jako kardioprotektivum. Co říkají metaanalýzy a co adherence?`);
}

// ============================================================
// SLIDE 16 – MA + adherence KV
// ============================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Tvrdé KV outcomy" });
  s.addText("Metaanalýzy a IPD metaanalýza: ITT nula, adherence možná", { placeholder: "title" });
  chip(s, ["MA RCT", "kohorta"]);
  msg(s, "IPD MA tří RCT: ITT HR 1,01, při ≥ 4 h/den HR 0,69 (kauzální inference z observačních dat); MA se v modifikaci adherencí rozcházejí; kohorty ukazují dávkově závislý pokles MACE – s rizikem healthy-adherer bias.");
  forest(s, {
    x: 0.5, y: 2.05, w: 12.35, h: 4.45, scale: "log", xmin: 0.2, xmax: 8, ref: 1, ticks: [0.25, 0.5, 1, 2, 4, 8], xlabel: "RR / HR / OR pro MACE či KV příhody (log škála); ◆ = nerandomizovaná adherenční podskupina; ○ = observační",
    head: ["Syntéza · srovnání", "Efekt (95 % CI)"], labelW: 6.0, valueW: 2.2,
    rows: [
      { label: "Metaanalýzy RCT – ITT", group: true },
      { label: "   Yu 2017 (JAMA) · 10 RCT, n 7 266 · MACE; meta-regrese: adherence bez vlivu", est: 0.77, lo: 0.53, hi: 1.13, color: HX.dk2, val: "RR 0,77 (0,53–1,13)" },
      { label: "   Khan 2018 (EHJ) · 7 RCT, n 4 268 · MACE", est: 0.74, lo: 0.47, hi: 1.17, color: HX.dk2, val: "RR 0,74 (0,47–1,17); I² 48 %" },
      { label: "   Labarca 2020 · 8 RCT, n 5 817 · MACE; podskupiny spavost/adherence/tíže bez efektu", est: 0.87, lo: 0.70, hi: 1.10, color: HX.dk2, val: "RR 0,87 (0,70–1,10)" },
      { label: "   IPD MA SAVE + ISAACC + RICCADSA (Sánchez-de-la-Torre 2023) · n 4 186 · MACCE, ITT", est: 1.01, lo: 0.87, hi: 1.17, color: HX.dk2, big: true, val: "HR 1,01 (0,87–1,17)" },
      { label: "Adherenční podskupiny (nerandomizované)", group: true },
      { label: "   Khan 2018 · adherence ≥ 4 h/noc", est: 0.43, lo: 0.23, hi: 0.80, color: HX.accent2, diamond: true, val: "RR 0,43 (0,23–0,80); I² 0 %" },
      { label: "   IPD MA 2023 · ≥ 4 h/den, marginální strukturální model", est: 0.69, lo: 0.52, hi: 0.92, color: HX.accent2, diamond: true, val: "HR 0,69 (0,52–0,92)" },
      { label: "   Almuzainy 2026 · 11 RCT, n 6 139 · ITT RR 1,00 (0,87–1,14); adherentní ≥ 4 h", est: 0.71, lo: 0.52, hi: 0.97, color: HX.accent2, diamond: true, val: "RR 0,71 (0,52–0,97)" },
      { label: "Observační kohorty", group: true },
      { label: "   Marin 2005 · muži, 10 let · neléčená těžká OSA vs. zdraví, fatální KV příhody", est: 2.87, lo: 1.17, hi: 7.51, color: HX.accent1, dim: true, val: "OR 2,87 (1,17–7,51)" },
      { label: "   Gervès-Pinquié 2022 · n 5 138, 6,6 let · PAP 6–7 h vs. 0–4 h/noc, MACE", est: 0.75, lo: 0.62, hi: 0.92, color: HX.accent1, dim: true, val: "HR 0,75 (0,62–0,92)" },
    ],
  });
  cite(s, "Yu, JAMA 2017 · Khan, Eur Heart J 2018 · Labarca, Sleep Med Rev 2020 · Sánchez-de-la-Torre, JAMA 2023 · Almuzainy, J Saudi Heart Assoc 2026 · Marin, Lancet 2005 · Gervès-Pinquié, AJRCCM 2022");
  addNotes(s, 16, `Metaanalýzy ITT dat jsou konzistentně nulové: relativní riziko 0,77 v deseti RCT bez vztahu k adherenci v meta-regresi (Yu 2017), 0,74 v sedmi RCT (Khan 2018) a 0,87 v osmi RCT bez modifikace spavostí, adherencí ani tíží apnoe (Labarca 2020). Nejkvalitnější syntézou je metaanalýza individuálních dat SAVE, ISAACC a RICCADSA s 4186 pacienty a průměrnou adherencí 3,1 hodiny: v ITT HR 1,01; v on-treatment analýze marginálními strukturálními modely měli pacienti s adherencí nad 4 hodiny HR 0,69 (Sánchez-de-la-Torre 2023). Podobně Khan 2018 dává u adherentních riziko 0,43 a letošní metaanalýza 11 RCT riziko 0,71 pro MACE a 0,36 pro kardiovaskulární mortalitu (Almuzainy 2026). Všechny adherenční analýzy jsou ale nerandomizované a metaanalýzy se v tom, zda adherence efekt modifikuje, rozcházejí; to je třeba říct otevřeně. Observační data říkají totéž, co intuice: neléčená těžká OSA měla u mužů za deset let 2,9krát vyšší riziko fatálních příhod než zdraví (Marin 2005) a ve francouzské kohortě 5138 pacientů klesalo riziko MACE s hodinami PAP, HR 0,75 při šesti až sedmi hodinách (Gervès-Pinquié 2022). Zdravý adherující pacient je ale jiný pacient než ten, kdo přístroj odloží. Proč tedy RCT vyšly neutrálně?`);
}

// ============================================================
// SLIDE 17 – Proč neutrální
// ============================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Tvrdé KV outcomy" });
  s.addText("Proč jsou RCT neutrální: selekce, fenotyp, adherence, metrika", { placeholder: "title" });
  chip(s, ["průřez", "kohorta", "syntéza"]);
  msg(s, "Čtyři důvody, které snižují sílu RCT – a které jsou vysvětlením, nikoli důkazem přínosu.");
  const tiles = [
    ["1,2–20,9 %", "Selekce", "pacientů s těžkou OSA ze spánkové laboratoře by splnilo kritéria 5 RCT; vylučovala hlavně spavost, KV onemocnění, noční hypoxémie (Reynor 2022, n 3 965)", HX.accent1],
    ["HR 1,7–2,4", "Fenotyp", "„excessively sleepy“ subtyp (16,7 %) měl v SHHS vyšší riziko incidentního KVO než ostatní subtypy (Mazzotti 2019) – a právě spaví byli z RCT vyloučeni; jiná kohorta po adjustaci nepotvrdila (Trzepizur 2022)", HX.accent5],
    ["2,8–3,3 h", "Adherence", "průměrné noční užívání v SAVE a ISAACC; v IPD MA 3,1 h – hluboko pod klinickou praxí; RCT podhodnocené (Pack 2021)", HX.dk2],
    ["HR 1,21 / 1,34", "Metrika", "hypoxic burden a T90 byly jedinými prediktory MACE u 5 358 pacientů bez KVO; zařazení do RCT šlo podle AHI (Trzepizur 2022; Azarbarzin 2019)", HX.accent2],
  ];
  tiles.forEach((t, i) => {
    const x = 0.5 + i * 3.1, y = 2.1, w = 2.9, h = 3.6;
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: HX.lt2 }, line: { color: HX.lt2 }, rectRadius: 0.1, objectName: "tile-" + t[1] });
    s.addText(t[0], { x: x + 0.15, y: y + 0.15, w: w - 0.3, h: 0.8, fontSize: 26, bold: true, color: t[3], margin: 0, isTextBox: true, valign: "middle", objectName: "tile-n-" + t[1] });
    s.addText(t[1], { x: x + 0.15, y: y + 0.95, w: w - 0.3, h: 0.4, fontSize: 14, bold: true, color: HX.dk1, margin: 0, isTextBox: true, objectName: "tile-t-" + t[1] });
    s.addText(t[2], { x: x + 0.15, y: y + 1.4, w: w - 0.3, h: h - 1.5, fontSize: 10.5, color: HX.dk1, margin: 0, isTextBox: true, valign: "top", objectName: "tile-b-" + t[1] });
  });
  s.addText("Důsledek pro praxi: CPAP indikujeme pro symptomy, tlak a kvalitu života (AASM 2019: PAP u spavosti STRONG, u hypertenze CONDITIONAL; žádné doporučení k prevenci KV příhod) – a kardiometabolické riziko řešíme hmotností.", { x: 0.5, y: 5.85, w: 12.35, h: 0.65, fontSize: 11.5, color: HX.dk2, bold: true, margin: 0, isTextBox: true, objectName: "dusledek" });
  cite(s, "Reynor, Sleep 2022 · Mazzotti, AJRCCM 2019 · Trzepizur, AJRCCM 2022 · Pack, Sleep 2021 · Azarbarzin, Eur Heart J 2019 · Javaheri, JACC 2024 · Patil, JCSM 2019 (AASM)");
  addNotes(s, 17, `Čtyři důvody, které nejsou výmluvou, ale ani důkazem přínosu. Zaprvé selekce: kritéria pěti RCT by splnilo jen 1,2 až 20,9 procenta pacientů s těžkou apnoí ze spánkové laboratoře; vylučovala hlavně nadměrná spavost, preexistující kardiovaskulární onemocnění a noční hypoxémie (Reynor 2022). Zadruhé fenotyp: právě nadměrně spaví pacienti, kteří byli z etických důvodů vyloučeni, měli ve Sleep Heart Health Study 1,7 až 2,4krát vyšší riziko incidentního kardiovaskulárního onemocnění než ostatní subtypy (Mazzotti 2019); jiná kohorta to po adjustaci nepotvrdila (Trzepizur 2022), takže i to zůstává otevřené. Zatřetí adherence 2,8 až 3,3 hodiny, hluboko pod tím, co vidíme u pacientů, kteří přijdou pro symptomy; podle analýzy biasů byly RCT na detekci efektu u adherentních podhodnocené (Pack 2021). A začtvrté metrika: do studií se zařazovalo podle AHI, zatímco riziko nese hypoxic burden a T90, s HR 1,21 a 1,34 pro MACE ve francouzské kohortě (Trzepizur 2022). Z toho plyne praktická věc, kterou říká i guideline AASM z roku 2019: CPAP indikujeme pro spavost, pro kvalitu života a podmíněně pro hypertenzi, nikoli abychom pacientovi zabránili v infarktu (Patil 2019); a pokud chceme cílit na kardiometabolické riziko, musíme pracovat s hmotností. Koho tedy máme v ambulanci aktivně hledat?`);
}

// ============================================================
// SLIDE 18 – Koho screenovat
// ============================================================
pres.addSection({ title: "Klinická praxe" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Klinická praxe" });
  s.addText("Klinická praxe I: koho screenovat", { placeholder: "title" });
  chip(s, ["guideline (2021–2026)", "průřez"]);
  msg(s, "Screening má smysl u rezistentní hypertenze, rekurentní FS, HF s podezřením na SDB, metabolického syndromu a DM2 – a nespoléhat na symptomy ani ESS.");
  const rows = [
    [{ text: "Indikace", options: { bold: true, color: HX.lt1, fill: { color: HX.dk1 } } }, { text: "Zdroj a znění", options: { bold: true, color: HX.lt1, fill: { color: HX.dk1 } } }, { text: "Síla / poznámka", options: { bold: true, color: HX.lt1, fill: { color: HX.dk1 } } }],
    ["Rezistentní / špatně kontrolovaná hypertenze", "AHA 2021: screening doporučen. ESC 2024 HTN: „OSAS should be suspected … in all patients with resistant hypertension, and in patients with non-dipping“", "AHA statement (bez tříd). ESC 2024: narativní znění, samostatná třída/LoE nenalezena (plný text neověřen)"],
    ["Rekurentní FS po kardioverzi / ablaci", "AHA 2021: screening doporučen. ESC 2024 FS: „Management of OSA may be considered as part of comprehensive management of risk factors … to reduce recurrence and progression“; screening pouze symptomovým dotazníkem se nedoporučuje", "ESC 2024: třída IIb (LoE k ověření v plném textu)"],
    ["Srdeční selhání NYHA II–IV s podezřením na SDB / spavostí", "AHA 2021: formální spánkové vyšetření „reasonable“", "U HF je CSA stejně častá jako OSA (Oldenburg 2007); AASM 2017: PSG, ne domácí test"],
    ["Plicní hypertenze; tachy-brady, VT, přeživší SCD", "AHA 2021: screening doporučen (PH); zvážit při podezření (arytmie)", "AHA statement"],
    ["DM2 / prediabetes", "ADA Standards of Care 2026: screening sleep health vč. poruch spánku (rec. 5.56, sekce 5 – znění a stupeň k ověření v plném textu); IDF 2008: oboustranný screening „strongly recommends“", "ADA stupeň NEOVĚŘENO; prevalence OSA u DM2 37–87 % (Sun 2026; Foster 2009)"],
    ["Metabolický syndrom", "HypnoLaus: horní kvartil AHI → MetS OR 2,80 (1,86–4,29) (Heinzer 2015)", "průřezová asociace; ~60 % u MetS dle přehledu (Drager 2013)"],
  ];
  s.addTable(rows, { x: 0.5, y: 1.95, w: 12.35, colW: [2.9, 6.25, 3.2], fontSize: 9.5, color: HX.dk1, fontFace: "Calibri", border: { type: "solid", pt: 0.5, color: "DDE3EA" }, fill: { color: HX.lt1 }, valign: "middle", margin: 0.05, rowH: [0.35, 0.62, 0.72, 0.5, 0.45, 0.72, 0.5], objectName: "tab-screening" });
  s.addText("Varování: u obézních diabetiků ve Sleep AHEAD nepredikoval přítomnost ani tíži OSA žádný symptom včetně spavosti (ESS 7,9 ± 4,6 při 86,6 % prevalenci) – Foster 2009. Epworth měří spavost, není to screening OSA (Johns 1991).", { x: 0.5, y: 5.95, w: 12.35, h: 0.6, fontSize: 10.5, bold: true, color: HX.accent1, margin: 0, isTextBox: true, objectName: "varovani" });
  cite(s, "Yeghiazarians, Circulation 2021 (AHA) · McEvoy, Eur Heart J 2024 (ESC HTN) · Van Gelder, Eur Heart J 2024 (ESC FS) · ADA, Diabetes Care 2026 · Shaw, Diabetes Res Clin Pract 2008 (IDF) · Heinzer, Lancet Respir Med 2015 · Foster, Diabetes Care 2009");
  addNotes(s, 18, `Doporučení pro kardiologa jsou poměrně konzistentní, i když s různou silou. Stanovisko AHA z roku 2021 doporučuje screening OSA u rezistentní nebo špatně kontrolované hypertenze, plicní hypertenze a rekurentní fibrilace síní po kardioverzi či ablaci a u srdečního selhání NYHA II až IV s podezřením na poruchu dýchání ve spánku považuje formální spánkové vyšetření za rozumné (Yeghiazarians 2021). Evropská guideline pro fibrilaci síní 2024 uvádí, že management OSA lze zvážit jako součást komplexní kontroly rizikových faktorů ke snížení rekurence, třída IIb, a že screening pouze symptomovým dotazníkem se nedoporučuje (Van Gelder 2024); úroveň evidence jsem nemohl ověřit z plného textu. Guideline pro hypertenzi 2024 říká, že na OSA máme myslet u všech pacientů s rezistentní hypertenzí a u non-dipperů (McEvoy 2024), bez samostatné třídy doporučení. Standardy ADA 2026 zahrnují screening zdraví spánku včetně poruch spánku u prediabetu a diabetu; přesné znění a stupeň ověříme z plného textu před přednáškou. Už v roce 2008 IDF důrazně doporučila oboustranné zvažování obou diagnóz (Shaw 2008). U metabolického syndromu měl horní kvartil AHI v populační studii OR 2,80 (Heinzer 2015). A jedno varování: u obézních diabetiků ve Sleep AHEAD nepredikoval přítomnost ani tíži OSA žádný symptom včetně spavosti (Foster 2009). Nespoléhejte na to, že pacient je ospalý. Jak tedy screenovat?`);
}

// ============================================================
// SLIDE 19 – STOP-BANG a referování
// ============================================================
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Klinická praxe" });
  s.addText("Klinická praxe II: STOP-BANG, domácí testování, kdy referovat", { placeholder: "title" });
  chip(s, ["MA obs.", "guideline (2017, 2019)"]);
  msg(s, "STOP-BANG ≥ 3 má senzitivitu 94 % pro AHI ≥ 15, ale specificitu 34 % a v nízkoprevalenční populaci PPV 16 % – vylučuje, nepotvrzuje; dotazník nikdy nenahrazuje test.");
  // algorithm
  const nb = (x, y, w, h, t, fill, tc, fs) => { s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.12, objectName: "alg-" + t.slice(0, 12) }); s.addText(t, { x, y, w, h, fontSize: fs || 10.5, bold: true, color: tc || HX.lt1, align: "center", valign: "middle", margin: 0.05, isTextBox: true, objectName: "algt-" + t.slice(0, 12) }); };
  const ar = (x, y, w) => s.addShape(pres.ShapeType.rightArrow, { x, y, w, h: 0.28, fill: { color: HX.accent4 }, line: { color: HX.accent4 }, objectName: "alga" + x + "-" + y });
  nb(0.5, 2.0, 2.3, 1.1, "Kardiologický pacient s indikací (slide 18)\n+ STOP-BANG (0–8)", HX.dk1);
  ar(2.9, 2.41, 0.4);
  nb(3.4, 2.0, 2.0, 0.6, "0–2: nízké riziko", HX.accent4);
  nb(3.4, 2.75, 2.0, 0.6, "3–4: střední – doplnit BMI > 35, krk, muž", HX.accent3, HX.dk1, 9.5);
  nb(3.4, 3.5, 2.0, 0.6, "5–8: vysoké riziko", HX.accent1);
  ar(5.5, 2.16, 0.4); ar(5.5, 2.91, 0.4); ar(5.5, 3.66, 0.4);
  nb(6.0, 2.0, 2.4, 0.6, "sledovat; při přetrvávajícím podezření test", HX.lt2, HX.dk1, 9.5);
  nb(6.0, 2.75, 2.4, 1.35, "Nekomplikovaný pacient:\ndomácí test (HSAT) nebo PSG\n\nNegativní/nejednoznačný HSAT → PSG", HX.accent2, HX.lt1, 9.5);
  nb(6.0, 4.25, 2.4, 1.0, "Významné kardiorespirační onemocnění (HF), hypoventilace, CMP, opioidy → rovnou PSG", HX.dk2, HX.lt1, 9.5);
  ar(8.5, 3.3, 0.4);
  nb(9.0, 2.75, 1.9, 2.5, "Spánková laboratoř:\ndiagnóza, fenotyp (AHI, ODI, T90, HB), indikace PAP / alternativ; metabolická péče a hmotnost paralelně", HX.dk1, HX.lt1, 9.5);
  s.addText("AASM 2017 (Kapur): dotazníky a predikční algoritmy nepoužívat k diagnóze bez PSG/HSAT (STRONG); PSG u nekomplikovaných se zvýšeným rizikem nebo HSAT (STRONG); PSG místo HSAT u významného kardiorespiračního onemocnění (STRONG). AASM 2019 (Patil): PAP u spavosti STRONG, u hypertenze CONDITIONAL.", { x: 0.5, y: 5.4, w: 10.4, h: 1.1, fontSize: 9.5, color: HX.dk2, margin: 0, isTextBox: true, valign: "top", objectName: "aasm" });
  box(s, 11.05, 2.0, 1.8, 3.3, "STOP-BANG ≥ 3 (Nagappa 2015, 17 studií, n 9 206)", [
    "Senzitivita 90 / 94 / 96 % pro AHI ≥ 5 / 15 / 30",
    "Specificita 49 / 34 / 25 %",
    "NPV 46 / 75 / 90 %",
    "Obecná populace (SHHS): PPV 16 % pro RDI ≥ 15",
    "P(těžká OSA): skóre 3 → 25 %, 5 → 45 %, 7–8 → 75 %",
  ], HX.dk1, { fontSize: 9, titleSize: 9.5 });
  cite(s, "Chung, Anesthesiology 2008 · Chung, Chest 2016 · Nagappa, PLoS One 2015 · Kapur, J Clin Sleep Med 2017 (AASM) · Patil, J Clin Sleep Med 2019 (AASM)");
  addNotes(s, 19, `Nástrojem první linie je STOP-BANG: osm otázek na chrápání, únavu, pozorované apnoe, tlak, BMI, věk, obvod krku a pohlaví, původně validovaných u chirurgických pacientů (Chung 2008). Metaanalýza 17 studií s 9206 pacienty dává při skóre 3 a více ve spánkových laboratořích senzitivitu 90, 94 a 96 procent pro AHI nad 5, 15 a 30, ale specificitu jen 49, 34 a 25 procent (Nagappa 2015). Negativní prediktivní hodnota pro těžkou apnoi je 90 procent, pozitivní prediktivní hodnota pro středně těžkou až těžkou apnoi v obecné populaci pouhých 16 procent (Nagappa 2015). STOP-BANG tedy vylučuje, nepotvrzuje. Pravděpodobnost těžké apnoe roste se skóre: 25 procent při skóre 3, 45 při 5 a 75 při 7 až 8 (Nagappa 2015). Praktická stratifikace: 0 až 2 nízké riziko, 5 až 8 vysoké, u 3 až 4 doplnit BMI nad 35, obvod krku nebo mužské pohlaví (Chung 2016). Guideline AASM 2017 říká jasně: dotazníky a predikční algoritmy nesmějí nahradit test; u nekomplikovaných pacientů se zvýšeným rizikem stačí domácí test, ale u pacientů s významným kardiorespiračním onemocněním, hypoventilací, po mozkové příhodě nebo na opioidech má být polysomnografie, a negativní domácí test při přetrvávajícím podezření vyžaduje polysomnografii (Kapur 2017). Pro naše pacienty se srdečním selháním to znamená rovnou spánkovou laboratoř, kde se rozliší i centrální apnoe. A jak to vypadá u našich pacientů?`);
}

// ============================================================
// SLIDES 20–21 – REZERVOVANÉ (prázdné)
// ============================================================
pres.addSection({ title: "Vlastní data CPM" });
for (const which of ["A", "B"]) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Vlastní data CPM" });
  s.addText("Vlastní data – Centrum preventivní medicíny (doplníme)", { placeholder: "title" });
  chip(s, ["placeholder " + which]);
  s.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.5, w: 12.35, h: 5.0, fill: { color: HX.lt1 }, line: { color: HX.accent4, width: 1, dashType: "dash" }, objectName: "plocha-" + which });
  const sugg = which === "A"
    ? `DOPLNIT: vlastní data CPM (slide A).\n\nNávrh typu dat (pouze návrh, nikoli obsah): prevalence vysokého rizika OSA podle STOP-BANG (skóre ≥ 3 a ≥ 5) u pacientů CPM s inzulinovou rezistencí, prediabetem nebo DM2, případně podle kategorií BMI. Logicky navazuje na slide 3 (prevalence OSA u DM2: 36–87 %, Sun 2026; Foster 2009) a na slide 19 (vlastnosti STOP-BANG, Nagappa 2015) a umožní srovnání našich pacientů s literaturou.\n\nMůstek k dalšímu slidu: „Druhá věc, kterou jsme se u našich pacientů dívali…“\n\nČas: 1 min.`
    : `DOPLNIT: vlastní data CPM (slide B).\n\nNávrh typu dat (pouze návrh, nikoli obsah): vztah ODI nebo AHI z domácí polygrafie k HOMA-IR či HbA1c v kohortě CPM, ideálně s adjustací na BMI a obvod pasu (srovnání s gradientem Priou 2012 a Punjabi 2009 ze slidu 8 a s hypoxemickými metrikami ze slidu 9); alternativně podíl pacientů, u kterých screening vedl k potvrzení OSA a k zahájení léčby (uzavírá algoritmus ze slidu 19).\n\nMůstek k take-home: „Co si z toho tedy odnést do ambulance?“\n\nČas: 1 min.`;
  addNotes(s, which === "A" ? 20 : 21, sugg);
}

// ============================================================
// SLIDE 22 – Take-home
// ============================================================
pres.addSection({ title: "Závěr" });
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "Závěr" });
  s.addText("Take-home messages a otevřené otázky", { placeholder: "title" });
  chip(s, ["syntéza"]);
  const th = [
    ["1", "OSA je u kardiologických a metabolických pacientů pravidlem, ne výjimkou – a většina o ní neví.", "AHA 2021; Young 1997"],
    ["2", "Asociace OSA → IR/DM2 je konzistentní a dávkově závislá a přetrvává po adjustaci na BMI i pas; kauzalita u lidí není dokázána, obezita zůstává hlavním confounderem.", "Nagayoshi 2016; Priou 2012; Resnick 2003; Ding 2023"],
    ["3", "CPAP zlepšuje spavost a TK, mírně HOMA-IR u nediabetiků; HbA1c a tvrdé KV outcomy v ITT nemění – efekt zřejmě vyžaduje dávku vysoko nad 4 h/noc.", "Wu 2026; Sánchez-de-la-Torre 2023; Herth 2023"],
    ["4", "Redukce hmotnosti včetně inkretinů je metabolicky silnější než CPAP; u OSA populace ale bez tvrdých KV outcomes a bez přímého srovnání s CPAP.", "Chirinos 2014; Malhotra 2024/2026"],
    ["5", "Screenovat cíleně, dotazníkem vylučovat, testem potvrzovat; u srdečního selhání rovnou PSG; nespoléhat na spavost.", "Nagappa 2015; Kapur 2017; Foster 2009"],
  ];
  th.forEach((t, i) => {
    const y = 1.45 + i * 0.86;
    s.addShape(pres.ShapeType.ellipse, { x: 0.5, y: y + 0.08, w: 0.5, h: 0.5, fill: { color: i === 2 || i === 3 ? HX.accent2 : HX.accent1 }, line: { color: i === 2 || i === 3 ? HX.accent2 : HX.accent1 }, objectName: "th-dot-" + i });
    s.addText(t[0], { x: 0.5, y: y + 0.08, w: 0.5, h: 0.5, fontSize: 14, bold: true, color: HX.lt1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "th-n-" + i });
    s.addText([{ text: t[1], options: { fontSize: 13, color: HX.lt1, breakLine: true } }, { text: t[2], options: { fontSize: 9.5, color: HX.accent4, italic: true } }], { x: 1.2, y, w: 7.6, h: 0.82, margin: 0, isTextBox: true, valign: "middle", objectName: "th-t-" + i });
  });
  box(s, 9.1, 1.45, 3.75, 4.3, "Otevřené otázky / gap v evidenci", [
    "Které fenotypy OSA (spaví, vysoký hypoxic burden) profitují z léčby na KV outcomy? – RCT u spavých eticky obtížné",
    "Je hypoxic burden lepší cíl a kritérium zařazení než AHI? – pro DM2 jediná prospektivní studie",
    "Jaká je prahová dávka CPAP pro metabolický efekt (≥ 7 h?) – pouze meta-regrese a malé in-lab studie",
    "Inkretiny vs. CPAP head-to-head; tvrdé KV outcomy inkretinů u OSA",
    "Třídy doporučení ESC/ADA k OSA – doplnit z plných textů",
  ], HX.lt1, { fontSize: 10, fill: HX.dk2, textColor: HX.lt1 });
  s.addText("Děkuji za pozornost.", { x: 0.5, y: 5.9, w: 8.3, h: 0.5, fontSize: 16, italic: true, color: HX.accent3, margin: 0, isTextBox: true, objectName: "diky" });
  addNotes(s, 22, `Dovolte mi pět závěrů. Zaprvé: OSA je u kardiologických a metabolických pacientů pravidlem, ne výjimkou, a většina z nich o ní neví (Yeghiazarians 2021; Young 1997). Zadruhé: asociace apnoe s inzulinovou rezistencí a diabetem je konzistentní a dávkově závislá a přetrvává po adjustaci na BMI i obvod pasu (Nagayoshi 2016; Priou 2012), ale kauzalita u lidí není dokázána a obezita zůstává hlavním confounderem (Resnick 2003; Ding 2023). Zatřetí: CPAP spolehlivě zlepšuje spavost a tlak a mírně inzulinovou rezistenci u nediabetiků, ale HbA1c ani tvrdé kardiovaskulární outcomy v ITT analýzách nemění (Wu 2026; Sánchez-de-la-Torre 2023); efekt zřejmě závisí na dávce vysoko nad čtyřmi hodinami (Herth 2023). Začtvrté: redukce hmotnosti včetně inkretinů je metabolicky silnější než CPAP (Chirinos 2014; Malhotra 2024), u populace s apnoí ale zatím bez tvrdých kardiovaskulárních outcomů a bez přímého srovnání s CPAP. Zapáté: screenujte cíleně, dotazníkem vylučujte a testem potvrzujte; u srdečního selhání rovnou polysomnografií a nespoléhejte na spavost (Nagappa 2015; Kapur 2017; Foster 2009). A otevřené otázky, které zůstávají: které fenotypy apnoe profitují z léčby, zda je hypoxic burden lepší cíl než AHI, jaká je prahová dávka CPAP pro metabolický efekt a zda inkretiny obstojí proti CPAP v přímém srovnání. Děkuji za pozornost a těším se na diskusi.`);
}

// ============================================================
// LITERATURA – Vancouver
// ============================================================
const ORDER = [
  // kap. 1
  "31300334", "23589584", "9406321", "34148375", "21968750", "11725173", "15249509", "17027333",
  // kap. 2
  "19279303", "17669711", "42607927", "39705149", "12610025", "16928713", "18268072", "28527878", "18544448", "25682233",
  // kap. 3
  "19265062", "19542260", "18172212", "17272786", "7560081", "9727547", "26541531", "34323304", "28595341", "16246965", "24127144", "10899061", "30684961", "23387384", "30422676", "11874813", "11874812",
  // kap. 4
  "11122588", "33243845", "30072403", "14582920", "9543264", "40999017", "36846222",
  // kap. 5
  "22988888", "26687279", "31872550", "27810258", "22688546", "19011148", "16192452", "30376054", "24897551", "36689314", "39349356", "22499826", "15353412",
  // kap. 6
  "23607839", "25700870", "22411171", "31518943", "30450787", "24918371", "28812180", "30073792", "33146450", "37673425", "41825890", "26926656", "32291275", "37062349", "33855195", "17557769", "27817148", "26910598", "35131440", "27860160", "25897569",
  // kap. 7
  "19786682", "23633746", "19959590", "24082315", "25192671", "19486716", "22990273", "27005405", "38912654", "41540105", "41135142", "42904878", "42837024", "40184186",
  // kap. 8
  "27571048", "31839558", "26914592", "28697252", "29069399", "32248026", "37787793", "42110829", "15781100", "35816570", "34739082", "30764637", "34648724", "33165616", "30736887",
  // kap. 9
  "39210715", "39210723", "41358897", "18431116", "26378880", "26658438", "28162150", "1798888",
];
function vancouver(r, i) {
  let au = r.authors.slice(0, 6).join(", ");
  if (r.authors.length > 6) au += ", et al";
  let t = r.title.replace(/\.$/, "");
  let vol = r.volume ? `;${r.volume}` : "";
  if (r.volume && r.issue) vol += `(${r.issue})`;
  let pg = r.pages ? `:${r.pages}` : "";
  if (!r.volume) vol = " [online]";
  return `${i}. ${au}. ${t}. ${r.journal}. ${r.year}${vol}${pg}. PMID: ${r.pmid}. doi: ${r.doi || "–"}`;
}
const refTexts = ORDER.map((p, i) => { const r = refByPmid[p]; if (!r) throw new Error("missing ref " + p); return vancouver(r, i + 1); });
const PER = 18;
const nRefSlides = Math.ceil(refTexts.length / PER);
for (let k = 0; k < nRefSlides; k++) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Závěr" });
  s.addText(`Literatura (${k + 1}/${nRefSlides}) – Vancouver, PMID / DOI`, { placeholder: "title" });
  if (k === 0) chip(s, ["rešerše PubMed 25. 10. 2026"]);
  const chunk = refTexts.slice(k * PER, (k + 1) * PER);
  const half = Math.ceil(chunk.length / 2);
  [chunk.slice(0, half), chunk.slice(half)].forEach((col, ci) => {
    if (!col.length) return;
    s.addText(col.map((t, j) => ({ text: t, options: { breakLine: j < col.length - 1, paraSpaceAfter: 3 } })), { x: 0.5 + ci * 6.25, y: 1.35, w: 6.1, h: 5.6, fontSize: 9, color: HX.dk1, margin: 0, isTextBox: true, valign: "top", objectName: "refs-" + k + "-" + ci });
  });
  s.addNotes(k === 0 ? "Seznam literatury – není součástí mluveného času. Všechna PMID/DOI ověřena v PubMedu 25. 10. 2026; podrobná tabulka evidence viz FAZE1_tabulka_evidence.md." : "Seznam literatury – pokračování.");
}

(async () => {
  const out = path.join(__dirname, "..", "OSA_inzulinova_rezistence_kardiolog.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  fs.writeFileSync(path.join(__dirname, "notes_wordcount.json"), JSON.stringify(notesStore, null, 1));
  console.log("written", out, "ref slides:", nRefSlides, "refs:", refTexts.length);
  console.log(notesStore.map(n => `${n.n}:${n.words}`).join("  "));
})();
