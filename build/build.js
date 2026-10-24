// Builds «محاضرة الكسوف الكلي ٢٠٢٧» (PowerPoint, Arabic RTL) from the book's figures.
// Usage: node build.js   (expects the book repository at ../../eclipse, or ECLIPSE_REPO=...)
const path = require("path");
const fs = require("fs");
const sharp = require("sharp");
const pptxgen = require("pptxgenjs");

const REPO = process.env.ECLIPSE_REPO || path.resolve(__dirname, "../../eclipse");
const ASSETS = path.resolve(__dirname, "../assets");
const OUT = path.resolve(__dirname, "../محاضرة_الكسوف_الكلي_2027.pptx");
const SKILL_THEME = path.join(__dirname, "apply_theme.js");

// ------------------------------------------------------------------ palette (the book's own)
const THEME = {
  name: "Eclipse 2027",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "1F1F1D", lt1: "FBFAF7", dk2: "0F1B2D", lt2: "EEF1F5",
    accent1: "F7B733", accent2: "1D4E89", accent3: "B3261E", accent4: "1F7A4D",
    accent5: "E88A2C", accent6: "8A8984", hlink: "1D4E89", folHlink: "6D4C8B",
  },
};
const K = {
  navy: "0F1B2D", navy2: "1A2A44", navy3: "24385A", paper: "FBFAF7", ink: "1F1F1D", ink2: "4A4A46",
  muted: "8A8984", sun: "F7B733", sunDeep: "E88A2C", red: "B3261E", green: "1F7A4D", blue: "1D4E89",
  line: "E4E2DC", card: "FFFFFF", tint: "EEF1F5", ice: "C9D6EA",
};
const W = 13.333, H = 7.5, M = 0.6;

// ------------------------------------------------------------------ images
const IMG = {
  corona: "data/nasa_images/corona.jpg",
  diamond: "data/nasa_images/diamond_ring.jpg",
  beads: "data/nasa_images/baily_beads.jpg",
  prom: "data/nasa_images/prominences.jpg",
  issShadow: "data/nasa_images/iss_shadow.jpg",
  viewing: "data/nasa_images/viewing.jpg",
  partialSeq: "data/nasa_images/partial_sequence.jpg",
  coincidence: "book/figures/ch1/fig_1_01_cosmic_coincidence.png",
  shadowGeom: "book/figures/ch1/fig_1_04_shadow_geometry.png",
  solarLunar: "book/figures/ch1/fig_1_06_solar_vs_lunar.png",
  inclination: "book/figures/ch1/fig_1_07_orbit_inclination.png",
  contacts: "book/figures/ch1/fig_1_10_four_contacts.png",
  saros: "book/figures/ch1/fig_1_22_saros_cycle.png",
  saros136: "book/figures/ch1/fig_1_23_saros136_paths.png",
  paths2001: "book/figures/ch1/fig_1_27_central_paths_2001_2050.png",
  skyLuxor: "book/figures/ch1/fig_1_28_sky_during_totality_luxor.png",
  whatChanges: "book/figures/ch1/fig_1_30_what_changes.png",
  raApophis: "book/figures/ch2/fig_2_02_myth_ra_apophis.png",
  akhenaten: "book/figures/ch2/fig_2_04_novel_akhenaten_1338bc.png",
  dendera: "book/figures/ch2/fig_2_05_novel_dendera_eclipses.png",
  haytham: "book/figures/ch2/fig_2_06_ibn_al_haytham_camera.png",
  ibnYunus: "book/figures/ch2/fig_2_07_ibn_yunus.png",
  e993: "book/figures/ch2/fig_2_10_novel_eclipse_993.png",
  e1004: "book/figures/ch2/fig_2_11_novel_eclipse_1004.png",
  lunarCairo: "book/figures/ch2/fig_2_12_novel_lunar_cairo.png",
  novelTable: "book/figures/ch2/fig_2_13_novel_eclipses_table.png",
  sohag1882: "book/figures/ch2/fig_2_15_eclipse_1882_comet_tewfik.png",
  salloum: "book/figures/ch2/fig_2_18_eclipse_2006_story.png",
  lastTotal: "book/figures/ch2/fig_2_23_last_total_eclipse_cities.png",
  monuments: "book/figures/ch2/fig_2_24_monuments_durations.png",
  monumentsMap: "book/figures/ch2/fig_2_25_monuments_map.png",
  regional: "maps_ar/02_regional_overview_totality_and_obscuration.png",
  egypt: "maps_ar/10_country_egypt.png",
  snapshots: "book/figures/ch3/fig_3_04_shadow_snapshots_grid.png",
  longest: "book/figures/ch3/fig_3_24_longest_totality_point.png",
  partialVsTotal: "book/figures/ch3/fig_3_21_partial_vs_total.png",
  luxorCard: "book/figures/ch3/fig_3_25_city_card_luxor.png",
  bahariya: "book/figures/ch6/fig_6_16_bahariya_oasis.png",
  golden: "book/figures/ch4/fig_4_01_golden_rule.png",
  safeUnsafe: "book/figures/ch4/fig_4_02_safe_vs_unsafe.png",
  glassesLuxor: "book/figures/ch4/fig_4_09_glasses_timeline_luxor.png",
  myths: "book/figures/ch4/fig_4_15_myths_facts.png",
  heat: "book/figures/ch4/fig_4_18_heat_safety.png",
  pinholes: "book/figures/ch4/fig_4_07_natural_pinholes.png",
  ancientRecords: "book/figures/ch8/fig_8_13_ancient_records_table.png",
  assyria: "book/figures/ch8/fig_8_03_eclipse_763bc_assyria.png",
  eddington: "book/figures/ch8/fig_8_10_eddington_1919.png",
  helium: "book/figures/ch8/fig_8_08_helium_1868.png",
  deltaT: "book/figures/ch8/fig_8_14_delta_t_curve.png",
  phone: "book/figures/ch5/fig_5_21_smartphone_tips.png",
  types: "book/figures/ch1/fig_1_05_eclipse_types.png",
  planLuxor: "book/figures/ch5/fig_5_14_totality_plan_luxor.png",
  phaseSettings: "book/figures/ch5/fig_5_05_phase_settings.png",
};
const DIM = {};

async function prepareImages() {
  fs.mkdirSync(ASSETS, { recursive: true });
  for (const [key, rel] of Object.entries(IMG)) {
    const out = path.join(ASSETS, key + ".jpg");
    if (!fs.existsSync(out)) {
      const src = path.join(REPO, rel);
      if (!fs.existsSync(src)) throw new Error("missing image " + src);
      await sharp(src).flatten({ background: "#fbfaf7" })
        .resize({ width: 1700, height: 1700, fit: "inside", withoutEnlargement: true })
        .jpeg({ quality: 84, mozjpeg: true }).toFile(out);
    }
    const m = await sharp(out).metadata();
    DIM[key] = { path: out, w: m.width, h: m.height };
  }
}

// ------------------------------------------------------------------ text helpers
const R = { rtlMode: true, align: "right", lang: "ar-EG" };
const t = (text, o = {}) => ({ text, options: { ...R, ...o } });

function bullets(items, o = {}) {
  return items.map((it, i) => {
    const last = i === items.length - 1;
    if (Array.isArray(it)) { // [bold lead, rest]
      return [
        { text: it[0], options: { ...R, bold: true, color: o.leadColor || K.ink, bullet: { indent: 18 }, ...o.run } },
        { text: it[1], options: { ...R, breakLine: !last, ...o.run } },
      ];
    }
    return [{ text: it, options: { ...R, bullet: { indent: 18 }, breakLine: !last, paraSpaceAfter: o.gap ?? 8, ...o.run } }];
  }).flat().map((r) => { if (r.options.bullet) r.options.paraSpaceAfter = o.gap ?? 8; return r; });
}

function fit(key, box) {
  const d = DIM[key];
  const r = d.w / d.h;
  let w = box.w, h = box.w / r;
  if (h > box.h) { h = box.h; w = box.h * r; }
  const ax = box.ax ?? 0.5, ay = box.ay ?? 0.5;
  return { path: d.path, x: box.x + (box.w - w) * ax, y: box.y + (box.h - h) * ay, w, h };
}

function addImg(slide, key, box, o = {}) {
  const p = fit(key, box);
  slide.addImage({ ...p, objectName: o.name || key, altText: o.alt || key,
    ...(o.frame ? { line: { color: K.line, width: 0.75 } } : {}) });
  return p;
}

// ------------------------------------------------------------------ deck
async function build() {
  await prepareImages();
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.rtlMode = true;
  pres.title = "الكسوف الكلي للشمس ٢ أغسطس ٢٠٢٧ — محاضرة للمرشدين السياحيين والأثريين";
  pres.subject = "الكسوف والخسوف، دوريتهما، الكسوف في الحضارة الإسلامية، كسوف ٢٠٢٧ في مصر";
  pres.author = "مبنية على كتاب «الكسوف الكلي للشمس — ٢ أغسطس ٢٠٢٧»";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  // layouts
  const footerText = { text: { text: "كسوف ٢ أغسطس ٢٠٢٧ | محاضرة تأهيلية للمرشدين السياحيين والأثريين",
    options: { x: 5.2, y: 7.02, w: 7.55, h: 0.3, fontSize: 10, color: K.muted, ...R } } };
  pres.defineSlideMaster({
    title: "CONTENT", background: { color: K.paper },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.28, w: W - 2 * M, h: 0.75,
        fontSize: 32, bold: true, color: K.navy, valign: "middle", ...R, margin: 0 }, text: "" } },
      footerText,
    ],
    slideNumber: { x: M, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: K.muted, align: "left" },
  });
  pres.defineSlideMaster({
    title: "DARK", background: { color: K.navy },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.28, w: W - 2 * M, h: 0.75,
        fontSize: 32, bold: true, color: "FFFFFF", valign: "middle", ...R, margin: 0 }, text: "" } },
      { text: { text: "كسوف ٢ أغسطس ٢٠٢٧ | محاضرة تأهيلية للمرشدين السياحيين والأثريين",
        options: { x: 5.2, y: 7.02, w: 7.55, h: 0.3, fontSize: 10, color: "8C9AB3", ...R } } },
    ],
    slideNumber: { x: M, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: "8C9AB3", align: "left" },
  });
  pres.defineSlideMaster({ title: "SECTION", background: { color: K.navy }, objects: [] });
  pres.defineSlideMaster({ title: "COVER", background: { color: K.navy }, objects: [] });

  let section = "";
  // title may be [arabic, english]: the English line sits under the Arabic title
  const add = (master, title) => {
    let en;
    if (Array.isArray(title)) [title, en] = title;
    const s = pres.addSlide({ masterName: master, sectionTitle: section });
    if (title) s.addText(title, { placeholder: "title", ...R });
    if (en) s.addText(en, { x: M, y: 1.0, w: W - 2 * M, h: 0.34, fontSize: 14, italic: true, align: "right", margin: 0,
      color: master === "DARK" ? K.sun : K.sunDeep, isTextBox: true, lang: "en-US", objectName: "title-en" });
    return s;
  };

  // eclipse-disc motif: dark disc + gold ring (+ optional label)
  const disc = (s, x, y, d, label, o = {}) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: o.fill || K.navy },
      line: { color: K.sun, width: o.ring || 2.5 }, objectName: o.name || "disc",
      shadow: { type: "outer", color: "F7B733", opacity: 0.35, blur: 8, offset: 0, angle: 0 } });
    if (label) s.addText(label, { x, y, w: d, h: d, align: "center", valign: "middle", fontSize: o.size || 20,
      bold: true, color: K.sun, margin: 0, isTextBox: true, rtlMode: true });
  };

  const card = (s, x, y, w, h, o = {}) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.12, fill: { color: o.fill || K.card }, line: { color: o.line || K.line, width: 0.75 },
    objectName: o.name || "card",
    shadow: o.flat ? undefined : { type: "outer", color: "000000", opacity: 0.08, blur: 6, offset: 2, angle: 90 } });

  const caption = (s, text, y = 6.62, o = {}) => s.addText(text, { x: o.x ?? M, y, w: o.w ?? (W - 2 * M), h: 0.3,
    fontSize: 10, color: o.color || K.muted, italic: true, isTextBox: true, ...R, margin: 0 });

  // figure + text column (text on the right, figure on the left: reading order RTL)
  const figText = (title, img, body, notes, o = {}) => {
    const s = add(o.dark ? "DARK" : "CONTENT", title);
    const split = o.split ?? 7.35;
    addImg(s, img, { x: M, y: 1.45, w: split - M - 0.3, h: 5.1 }, { frame: !o.dark && !o.noFrame });
    const tx = split + 0.05, tw = W - M - tx;
    if (o.lead) {
      s.addText(o.lead, { x: tx, y: 1.45, w: tw, h: 0.9, fontSize: 20, bold: true, color: o.dark ? K.sun : K.blue,
        isTextBox: true, ...R, margin: 0, valign: "top" });
    }
    s.addText(bullets(body, { run: { fontSize: o.fs || 16, color: o.dark ? "E8ECF3" : K.ink },
      leadColor: o.dark ? K.sun : K.navy }), { x: tx, y: o.lead ? 2.4 : 1.5, w: tw, h: o.lead ? 4.2 : 5.1,
      valign: "top", isTextBox: true, margin: 0, ...R });
    if (o.cap) caption(s, o.cap, 6.62, { color: o.dark ? "8C9AB3" : K.muted });
    s.addNotes(notes);
    return s;
  };

  const sectionSlide = (num, title, sub, notes) => {
    let en;
    if (Array.isArray(title)) [title, en] = title;
    section = title;
    pres.addSection({ title });
    const s = pres.addSlide({ masterName: "SECTION", sectionTitle: section });
    // big eclipse: corona glow behind a dark disc
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: 1.55, w: 4.4, h: 4.4, fill: { color: K.sun, transparency: 82 },
      line: { color: K.sun, width: 0 }, objectName: "corona-glow" });
    s.addShape(pres.shapes.OVAL, { x: 1.35, y: 1.9, w: 3.7, h: 3.7, fill: { color: K.sun, transparency: 60 },
      line: { color: K.sun, width: 0 }, objectName: "corona-inner" });
    disc(s, 1.6, 2.15, 3.2, num, { size: 72, ring: 4, name: "section-disc" });
    s.addText(title, { x: 5.9, y: 1.85, w: 6.85, h: 2.0, fontSize: 34, bold: true, color: "FFFFFF",
      valign: "bottom", isTextBox: true, ...R, margin: 0 });
    s.addText(sub, { x: 5.9, y: 3.95, w: 6.85, h: 1.0, fontSize: 17, color: K.ice, valign: "top",
      isTextBox: true, ...R, margin: 0 });
    if (en) s.addText(en, { x: 5.9, y: 5.0, w: 6.85, h: 0.6, fontSize: 16, italic: true, color: K.sun, align: "right",
      isTextBox: true, margin: 0, lang: "en-US" });
    s.addNotes(notes);
    return s;
  };

  // a lighter divider inside a stage (e.g. ٢-أ)
  const partSlide = (tag, title, en, sub, notes) => {
    const s = pres.addSlide({ masterName: "SECTION", sectionTitle: section });
    disc(s, W - M - 1.5, 2.55, 1.5, tag, { size: 26, name: "part-disc" });
    s.addText(title, { x: M, y: 2.35, w: W - 2 * M - 1.9, h: 1.0, fontSize: 34, bold: true, color: "FFFFFF",
      isTextBox: true, ...R, margin: 0, valign: "middle" });
    s.addText(en, { x: M, y: 3.35, w: W - 2 * M - 1.9, h: 0.45, fontSize: 17, italic: true, color: K.sun, align: "right",
      isTextBox: true, margin: 0, lang: "en-US" });
    s.addText(sub, { x: M, y: 3.9, w: W - 2 * M - 1.9, h: 0.9, fontSize: 16, color: K.ice, isTextBox: true, ...R, margin: 0, valign: "top" });
    s.addNotes(notes);
    return s;
  };
  const breakSlide = (label, next) => {
    const s = pres.addSlide({ masterName: "SECTION", sectionTitle: section });
    addImg(s, "partialSeq", { x: M, y: 0.9, w: W - 2 * M, h: 3.6 }, { name: "break-sequence", alt: "مراحل الكسوف، صورة ناسا" });
    s.addText("استراحة — ١٠ دقائق", { x: M, y: 4.8, w: W - 2 * M, h: 0.8, fontSize: 36, bold: true, color: K.sun, align: "center",
      isTextBox: true, rtlMode: true, margin: 0 });
    s.addText("Break — 10 minutes", { x: M, y: 5.55, w: W - 2 * M, h: 0.4, fontSize: 16, italic: true, color: K.ice, align: "center", isTextBox: true, margin: 0 });
    s.addText("بعد الاستراحة: " + next, { x: M, y: 6.1, w: W - 2 * M, h: 0.5, fontSize: 18, color: "FFFFFF", align: "center",
      isTextBox: true, rtlMode: true, margin: 0 });
    s.addText("صورة: مراحل كسوف ٢٠٢٤ — ناسا (NASA/Aubrey Gemignani)", { x: M, y: 6.95, w: W - 2 * M, h: 0.3, fontSize: 10, color: "8C9AB3",
      align: "center", isTextBox: true, rtlMode: true, margin: 0 });
    s.addNotes(label);
    return s;
  };

  // ================================================================== 1. cover
  section = "الافتتاح";
  pres.addSection({ title: section });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: section });
    addImg(s, "corona", { x: 0, y: 0, w: 6.2, h: H, ax: 0.5 }, { name: "cover-corona", alt: "الهالة الشمسية أثناء الكلية، صورة ناسا" });
    s.addText("محاضرة تأهيلية للمرشدين السياحيين والأثريين", { x: 6.4, y: 1.0, w: 6.35, h: 0.5, fontSize: 18,
      color: K.sun, bold: true, isTextBox: true, ...R, margin: 0 });
    s.addText("الكسوف الكلي للشمس", { x: 6.4, y: 1.55, w: 6.35, h: 0.95, fontSize: 38, bold: true, color: "FFFFFF",
      isTextBox: true, ...R, margin: 0 });
    s.addText("٢ أغسطس ٢٠٢٧ — كسوف القرن فوق صعيد مصر", { x: 6.4, y: 2.55, w: 6.35, h: 0.6, fontSize: 22,
      color: K.ice, isTextBox: true, ...R, margin: 0 });
    s.addText("The Total Solar Eclipse of 2 August 2027 — a training lecture for tour guides & archaeologists", { x: 6.4, y: 3.15, w: 6.35, h: 0.55,
      fontSize: 14, italic: true, color: K.sun, align: "right", isTextBox: true, margin: 0, lang: "en-US" });
    s.addText([
      t("الكسوف والخسوف: كيف يحدثان وكم مرة يتكرران", { breakLine: true }),
      t("الكسوف والخسوف في الحضارة الإسلامية وفي القاهرة الفاطمية", { breakLine: true }),
      t("مسار كسوف ٢٠٢٧ ومواعيده في المدن والمواقع الأثرية المصرية"),
    ], { x: 6.4, y: 3.9, w: 6.35, h: 1.5, fontSize: 16, color: "E8ECF3", paraSpaceAfter: 6, isTextBox: true, margin: 0, valign: "top" });
    s.addText("المادة العلمية: كتاب «الكسوف الكلي للشمس — ٢ أغسطس ٢٠٢٧» ورواية «أرصاد وأطوار»", {
      x: 6.4, y: 5.9, w: 6.35, h: 0.6, fontSize: 12, color: "8C9AB3", isTextBox: true, ...R, margin: 0 });
    s.addText("صورة الخلفية: ناسا (NASA/Aubrey Gemignani)", { x: 0.25, y: 7.05, w: 4, h: 0.3, fontSize: 9,
      color: "C9D6EA", isTextBox: true, align: "left", margin: 0 });
    s.addNotes(
      "الترحيب وتقديم المحاضرة. هدفنا اليوم أن يخرج المرشد السياحي والأثري قادرًا على ثلاثة أمور: " +
      "أن يشرح للزائر ما سيحدث في السماء بلغة بسيطة دقيقة، وأن يربط الحدث بتاريخ مصر وبالحضارة الإسلامية، " +
      "وأن يدير مجموعته يوم الكسوف بأمان: متى يلبسون النظارة ومتى ينزعونها، وأين يقفون، وكيف يتقون الحر. " +
      "كل الأرقام في هذه المحاضرة محسوبة في الكتاب من عناصر بِسِل الرسمية التي نشرتها ناسا لهذا الكسوف، " +
      "والأوقات كلها بتوقيت مصر الصيفي (التوقيت العالمي + ٣).");
  }

  // ================================================================== 2. why it matters
  {
    const s = add("CONTENT", ["لماذا يهمّ هذا الحدث المرشد السياحي؟", "Why this event matters to tour guides"]);
    const cards = [
      ["٦ د ٢٠ ث", "مدة الكلية في الأقصر", "الكرنك ووادي الملوك والدير البحري داخل الظل الكامل، والشمس شبه عمودية فوق الرأس"],
      ["٦٥٠ سنة", "منذ آخر كسوف كلي فوق الأقصر", "آخر كلية مرت بطيبة كانت في ١٠ يناير ١٣٧٧م تقريبًا، والقاهرة منذ ١٢٤١م"],
      ["ملايين", "زائر متوقَّع في الصعيد", "سيأتي الزوار من أجل ست دقائق؛ المرشد هو من يحوّلها إلى تجربة آمنة لا تُنسى"],
    ];
    cards.forEach(([big, label, desc], i) => {
      const x = W - M - (i + 1) * 3.9 - i * 0.2 + 0.2 * 0, w = 3.9;
      const cx = W - M - w - i * (w + 0.25);
      card(s, cx, 1.55, w, 3.6, { name: "why-card-" + (i + 1) });
      disc(s, cx + w - 1.0, 1.8, 0.65, "", { name: "why-disc-" + (i + 1), ring: 2 });
      s.addText(big, { x: cx + 0.25, y: 2.55, w: w - 0.5, h: 0.85, fontSize: 36, bold: true, color: K.sunDeep,
        isTextBox: true, ...R, margin: 0 });
      s.addText(label, { x: cx + 0.25, y: 3.4, w: w - 0.5, h: 0.5, fontSize: 17, bold: true, color: K.navy,
        isTextBox: true, ...R, margin: 0 });
      s.addText(desc, { x: cx + 0.25, y: 3.9, w: w - 0.5, h: 1.15, fontSize: 14, color: K.ink2,
        isTextBox: true, ...R, margin: 0, valign: "top" });
      void x;
    });
    card(s, M, 5.45, W - 2 * M, 1.15, { fill: K.navy, line: K.navy, name: "why-question", flat: true });
    s.addText([
      t("السؤال الذي سيسمعه كل مرشد يوم الكسوف: ", { bold: true, color: K.sun }),
      t("«متى أنزع النظارة؟ وهل أرى الكسوف كليًا من هنا؟» — هذه المحاضرة تعطيك الجواب لكل مدينة وموقع أثري.", { color: "FFFFFF" }),
    ], { x: M + 0.35, y: 5.5, w: W - 2 * M - 0.7, h: 1.05, fontSize: 17, valign: "middle", isTextBox: true, margin: 0 });
    s.addNotes(
      "ابدأ بالمفارقة: الأقصر — أغنى مدينة في العالم بالآثار — لم تشهد كسوفًا كليًا منذ نحو ستة قرون ونصف " +
      "(١٠ يناير ١٣٧٧م بحساب الكتاب، وهو تقدير تقريبي لأن تباطؤ دوران الأرض في العصور الوسطى غير معروف بدقة). " +
      "والقاهرة لم ترَه منذ ٦ أكتوبر ١٢٤١م. ويوم ٢ أغسطس ٢٠٢٧ تدوم الكلية في الأقصر ٦ دقائق و٢٠ ثانية. " +
      "دور المرشد: الشرح، والربط بالتاريخ، والسلامة. وسنعود في آخر المحاضرة إلى خطة عملية للمرشد يوم الكسوف.");
  }

  {
    const s = add("CONTENT", ["خطة المحاضرة: ثلاث مراحل في ثلاث ساعات", "Lecture plan: three stages in three hours"]);
    const st = [
      ["١", "المرحلة الأولى", "٠:٠٠ – ٠:٥٠", "تعريف الكسوف والخسوف وهندستهما", "الظل وأنواع الكسوف، الكسوف والخسوف، العقدتان والمواسم، التماسات الأربعة وظواهر الكلية، دورة ساروس، وكم يتكرر الكسوف في العالم وعلى المكان نفسه"],
      ["٢", "المرحلة الثانية", "١:٠٠ – ١:٥٠", "الكسوف في تاريخ مصر والحضارة الإسلامية", "أ) الكسوفات التي حدثت في مصر: من رع وإخناتون ودندرة إلى القاهرة الفاطمية وسوهاج ١٨٨٢ والسلوم ٢٠٠٦\nب) الحديث النبوي وصلاة الكسوف وابن يونس وابن الهيثم"],
      ["٣", "المرحلة الثالثة", "٢:٠٠ – ٣:٠٠", "كسوف ٢ أغسطس ٢٠٢٧: الرصد والتصوير", "المسار ومواعيده في مدن مصر، مراحله كاملة، كيف يبدو في الأقصر وشكل الشمس في كل مرحلة، الرصد الآمن، التصوير بالهاتف والكاميرا، وخطة المرشد"],
    ];
    st.forEach(([n, lab, time, head, body], i) => {
      const w = 3.85, x = W - M - w - i * (w + 0.27), y = 1.55, h = 4.55;
      card(s, x, y, w, h, { name: "stage-card-" + n });
      disc(s, x + w - 1.15, y + 0.25, 0.85, n, { size: 26, name: "stage-disc-" + n });
      s.addText(lab, { x: x + 0.25, y: y + 0.25, w: w - 1.5, h: 0.42, fontSize: 18, bold: true, color: K.sunDeep, isTextBox: true, ...R, margin: 0 });
      s.addText(time, { x: x + 0.25, y: y + 0.67, w: w - 1.5, h: 0.38, fontSize: 15, color: K.muted, isTextBox: true, ...R, margin: 0 });
      s.addText(head, { x: x + 0.25, y: y + 1.25, w: w - 0.5, h: 0.85, fontSize: 18, bold: true, color: K.navy, isTextBox: true, ...R, margin: 0, valign: "top" });
      s.addText(body, { x: x + 0.25, y: y + 2.1, w: w - 0.5, h: 2.35, fontSize: 14, color: K.ink2, isTextBox: true, ...R, margin: 0, valign: "top" });
    });
    s.addText("استراحتان من ١٠ دقائق بعد المرحلتين الأولى والثانية", { x: M, y: 6.25, w: W - 2 * M, h: 0.4, fontSize: 15, italic: true,
      color: K.blue, isTextBox: true, ...R, margin: 0 });
    s.addNotes("المحاضرة ثلاث ساعات في ثلاث مراحل: (١) التعريف والهندسة نحو خمسين دقيقة، ثم استراحة؛ " +
      "(٢) الكسوف في تاريخ مصر وفي الحضارة الإسلامية نحو خمسين دقيقة، ثم استراحة؛ " +
      "(٣) كسوف ٢٠٢٧ تفصيلًا: المسار والمواعيد، ومراحله، وشكله في الأقصر، والرصد الآمن، والتصوير، وخطة المرشد — نحو ساعة.");
  }

  sectionSlide("١", ["المرحلة الأولى: تعريف الكسوف والخسوف وهندستهما", "Stage 1 — Definitions and the geometry of eclipses"], "الساعة الأولى (٠:٠٠ – ٠:٥٠): كيف يحدثان؟ وكم مرة يتكرران في العالم وعلى المكان نفسه؟",
    "المرحلة الأولى (نحو خمسين دقيقة): الأساس العلمي. الرسالة التي يحفظها المرشد: الكسوف = القمر يحجب الشمس في آخر الشهر القمري، والخسوف = ظل الأرض يغطي القمر ليلة البدر. ثم الدورية: كسوف كلي في مكان ما كل ١٨ شهرًا تقريبًا، لكن المدينة الواحدة تنتظره قرونًا.");

  figText(["مصادفة كونية: قرصان بالحجم نفسه", "A cosmic coincidence: two discs of the same apparent size"], "coincidence", [
    ["٤٠٠ مرة: ", "قطر الشمس أكبر من قطر القمر بنحو ٤٠٠ مرة، وهي أبعد عنا بنحو ٤٠٠ مرة أيضًا — قطر زاوي نصف درجة (Angular diameter)"],
    ["النتيجة: ", "يبدو القرصان في السماء بالحجم نفسه تقريبًا: نحو نصف درجة"],
    ["القمر قريب — الحضيض (Perigee): ", "يبدو أكبر من الشمس فيحجبها كلها ← كسوف كلي (Total)"],
    ["القمر بعيد — الأوج (Apogee): ", "تبقى حلقة مضيئة حوله ← كسوف حلقي (Annular)"],
  ], "جملة تصلح للسائح: «لو كان القمر أصغر قليلًا أو أبعد قليلًا لما رأينا كسوفًا كليًا قط». " +
    "النسبة ٤٠٠ إلى ٤٠٠ مصادفة لا قانون، والقمر يبتعد عن الأرض نحو ٣٫٨ سم كل سنة، فبعد مئات الملايين من السنين لن يحدث كسوف كلي. " +
    "في ٢٠٢٧ القمر قريب من الحضيض، فيبدو أكبر من الشمس بنحو ٤٪ (قدر الكسوف ١٫٠٣٩)، ولهذا تطول الكلية.",
    { lead: "لماذا يمكن أن يحدث كسوف كلي أصلًا؟", cap: "الشكل: كتاب «الكسوف الكلي للشمس — ٢٠٢٧»، الشكل ١-١" });

  figText(["الظل الكامل وشبه الظل وأنواع الكسوف", "Umbra, penumbra and the types of solar eclipse"], "shadowGeom", [
    ["الظل الكامل (Umbra): ", "لا يصله أي ضوء من قرص الشمس؛ من يقف فيه يرى كسوفًا كليًا"],
    ["شبه الظل (Penumbra): ", "يحجب جزءًا من القرص فقط؛ من يقف فيه يرى كسوفًا جزئيًا"],
    ["الظل المضاد (Antumbra): ", "إذا لم يبلغ رأس الظل الأرض نرى حلقة من نار: كسوف حلقي"],
    ["الهجين (Hybrid): ", "حلقي في جزء من المسار وكلي في جزء آخر، مثل كسوف ١٠٠٤م فوق القاهرة"],
  ], "شرح بالرسم: مخروطا الظل. الظل الكامل بقعة صغيرة على الأرض عرضها في ٢٠٢٧ لا يتجاوز ٢٥٩ كيلومترًا، " +
    "أما شبه الظل فيغطي آلاف الكيلومترات — ولهذا يرى أهل القاهرة والإسكندرية الكسوف جزئيًا فقط. " +
    "نبّه إلى أن الفرق بين الجزئي والكلي ليس فرقًا في الدرجة بل في النوع: حتى ٩٩٪ ليس كليًا.",
    { cap: "الشكل ٤-١ من الكتاب: مخروطا ظل القمر" });

  // solar vs lunar comparison
  {
    const s = add("CONTENT", ["الكسوف والخسوف: ما الفرق؟", "Solar eclipse vs. lunar eclipse"]);
    const cols = [
      { head: "كسوف الشمس (Solar eclipse)", color: K.sunDeep, rows: [
        ["الترتيب", "الشمس ← القمر ← الأرض"],
        ["متى؟", "عند المحاق: آخر الشهر القمري (٢٨ أو ٢٩)"],
        ["من يراه؟", "شريط ضيق على الأرض (الكلي)، وما حوله جزئيًا"],
        ["مدة الكلية", "دقائق: أقصاها نحو ٧ دقائق ونصف"],
        ["السلامة", "لا يُنظر إليه إلا بنظارة معتمدة، إلا لحظة الكلية"],
      ] },
      { head: "خسوف القمر (Lunar eclipse)", color: K.blue, rows: [
        ["الترتيب", "الشمس ← الأرض ← القمر"],
        ["متى؟", "عند البدر: منتصف الشهر القمري (١٤ أو ١٥)"],
        ["من يراه؟", "كل من كان القمر فوق أفقه: نصف الكرة الأرضية"],
        ["مدة الكلية", "قد تبلغ ساعة وأربعين دقيقة"],
        ["السلامة", "آمن تمامًا بالعين المجردة، ويحمرّ القمر"],
      ] },
    ];
    cols.forEach((c, i) => {
      const w = 5.85, x = i === 0 ? W - M - w : M, y = 1.5;
      card(s, x, y, w, 5.05, { name: "compare-" + i });
      s.addShape(pres.shapes.OVAL, { x: x + w - 0.95, y: y + 0.28, w: 0.6, h: 0.6, fill: { color: i === 0 ? K.navy : "8E3B2A" },
        line: { color: c.color, width: 2.5 }, objectName: "compare-icon-" + i });
      const [hAr, hEn] = c.head.replace(")", "").split(" (");
      s.addText(hAr, { x: x + 0.3, y: y + 0.18, w: w - 1.4, h: 0.5, fontSize: 24, bold: true, color: c.color,
        isTextBox: true, ...R, margin: 0 });
      s.addText(hEn, { x: x + 0.3, y: y + 0.66, w: w - 1.4, h: 0.3, fontSize: 13, italic: true, color: K.muted, align: "right",
        isTextBox: true, margin: 0, lang: "en-US" });
      c.rows.forEach(([k, v], j) => {
        const yy = y + 1.1 + j * 0.78;
        s.addText(k, { x: x + w - 1.85, y: yy, w: 1.55, h: 0.65, fontSize: 15, bold: true, color: K.navy,
          isTextBox: true, ...R, margin: 0, valign: "top" });
        s.addText(v, { x: x + 0.3, y: yy, w: w - 2.3, h: 0.7, fontSize: 15, color: K.ink,
          isTextBox: true, ...R, margin: 0, valign: "top" });
      });
    });
    s.addNotes("هذه الشريحة هي ما يُسأل عنه المرشد أكثر من غيره. الفارق الجوهري: الكسوف يقع دائمًا في آخر الشهر الهجري " +
      "(كسوف ٢٠٢٧ يوافق ٢٨ أو ٢٩ صفر ١٤٤٩هـ بحسب الرؤية)، والخسوف دائمًا في منتصف الشهر ليلة البدر. " +
      "وفي اللغة يُستعمل اللفظان للشمس والقمر معًا، وقد شاع في الاصطلاح: الكسوف للشمس والخسوف للقمر، كما في قوله تعالى «وَخَسَفَ الْقَمَرُ». " +
      "أقصى كلية شمسية ممكنة نحو ٧ دقائق و٣٠ ثانية؛ وأطول خسوف كلي للقمر يقارب ساعة و٤٥ دقيقة.");
  }

  figText(["لماذا لا يحدث الكسوف كل شهر؟", "Why not every month? Orbital inclination, lunar nodes and eclipse seasons"], "inclination", [
    ["ميل المدار ٥° (Inclination): ", "مدار القمر مائل نحو خمس درجات على مدار الأرض، فيمر القمر في أغلب الشهور فوق الشمس أو تحتها"],
    ["العقدتان (Lunar nodes): ", "لا يقع الكسوف إلا إذا وافق المحاقُ وجودَ القمر قرب إحدى نقطتي تقاطع المدارين"],
    ["موسما الكسوف (Eclipse seasons): ", "مرتان كل عام تقريبًا؛ في ٢٠٢٧ موسم فبراير (كسوف حلقي) وموسم يوليو–أغسطس (كسوفنا الكلي)"],
    ["في السنة: ", "كسوفان للشمس على الأقل وخمسة على الأكثر، ومجموع الكسوفات والخسوفات لا يزيد على سبعة"],
  ], "تشبيه مفيد: مدار الأرض ومدار القمر طوقان مائلان أحدهما على الآخر قليلًا، يتقاطعان في نقطتين فقط. " +
    "ولا يصطف الثلاثة على خط واحد إلا عندما يكون القمر عند إحدى النقطتين في المحاق أو البدر. " +
    "ولهذا يأتي الكسوف والخسوف غالبًا متقاربين: خسوف قبل الكسوف أو بعده بأسبوعين.",
    { cap: "الشكل ٧-١ من الكتاب: ميل مدار القمر وعقدتاه" });

  figText(["مراحل الكسوف الكلي: التماسات الأربعة", "Phases of a total eclipse: the four contacts (C1–C4)"], "contacts", [
    ["التماس الأول C1 — First contact (١١:٤٠:٠٧): ", "القمر يلمس حافة الشمس — يبدأ الجزئي"],
    ["التماس الثاني C2 — Second contact (١٣:٠١:٥٨): ", "يكتمل الحجب — تبدأ الكلية وتُنزع النظارة"],
    ["الذروة — Maximum (١٣:٠٥:٠٨): ", "منتصف الكلية والشمس على ارتفاع ٨٢° تقريبًا"],
    ["التماس الثالث C3 — Third contact (١٣:٠٨:١٨): ", "يعود أول شعاع — تُلبس النظارة فورًا"],
    ["التماس الرابع C4 — Fourth contact (١٤:٢٦:٢٥): ", "ينفصل القرصان وينتهي الكسوف"],
  ], "الأوقات المذكورة لمدينة الأقصر بتوقيت مصر الصيفي، محسوبة في الكتاب. " +
    "المدة الكلية للحدث من أوله إلى آخره ساعتان و٤٦ دقيقة، أما الكلية فست دقائق وعشرون ثانية فقط. " +
    "ملحوظة: تذكر رواية «أرصاد وأطوار» للأقصر ١٣:٠٢:١٤–١٣:٠٨:٣٦ (٦ د ٢٢ ث)؛ والفرق ثوانٍ مصدره اختلاف موضع النقطة المختارة داخل المدينة وطريقة الحساب. " +
    "اعتمد في الإرشاد أرقام الكتاب المحسوبة بعناصر ناسا، وأبلغ الزوار بأن يكونوا مستعدين قبل الموعد بدقيقتين.",
    { lead: "مثال: الأقصر يوم ٢ أغسطس ٢٠٢٧", fs: 15, cap: "الشكل ١٠-١ من الكتاب؛ الأوقات بتوقيت مصر الصيفي (ت.ع +٣)" });

  // phenomena photo grid (dark)
  {
    const s = add("DARK", ["ما سيراه الزوار خلال الكلية", "What visitors will see during totality"]);
    const ph = [["diamond", "الخاتم الماسي (Diamond ring)", "لحظة قبل الكلية وبعدها مباشرة"],
      ["beads", "حبات بيلي (Baily's beads)", "ضوء الشمس عبر وديان حافة القمر"],
      ["corona", "الهالة (Corona)", "الغلاف الخارجي للشمس، لا يُرى إلا في الكلية"],
      ["prom", "الشواظ (Prominences)", "ألسنة حمراء فوق الكروموسفير (Chromosphere) عند حافة القمر"]];
    const w = 2.85, gap = 0.2, h = 3.6;
    ph.forEach(([k, name, desc], i) => {
      const x = W - M - w - i * (w + gap), y = 1.5;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: h + 1.2, fill: { color: K.navy2 }, line: { color: K.navy3, width: 0.75 },
        objectName: "photo-card-" + i });
      const d = DIM[k], r = d.w / d.h;
      // cover-crop into the frame
      const fw = w, fh = h;
      let iw = fw, ih = fw / r; if (ih < fh) { ih = fh; iw = fh * r; }
      s.addImage({ path: d.path, x, y, w: iw, h: ih, sizing: { type: "crop", x: (iw - fw) / 2, y: (ih - fh) / 2, w: fw, h: fh },
        objectName: "photo-" + k, altText: name });
      const [nAr, nEn] = name.replace(/\)$/, "").split(" (");
      s.addText(nAr, { x: x + 0.15, y: y + h + 0.06, w: w - 0.3, h: 0.38, fontSize: 17, bold: true, color: K.sun,
        isTextBox: true, ...R, margin: 0 });
      s.addText(nEn, { x: x + 0.15, y: y + h + 0.42, w: w - 0.3, h: 0.26, fontSize: 12, italic: true, color: K.ice, align: "right",
        isTextBox: true, margin: 0, lang: "en-US" });
      s.addText(desc, { x: x + 0.15, y: y + h + 0.7, w: w - 0.3, h: 0.48, fontSize: 12, color: "E8ECF3",
        isTextBox: true, ...R, margin: 0, valign: "top" });
    });
    caption(s, "صور ناسا للاستخدام التعليمي: NASA/Aubrey Gemignani (الخاتم، الحبات، الهالة)، NASA/Keegan Barber (الشواظ)", 6.55, { color: "8C9AB3" });
    s.addNotes("هذه صور حقيقية من كسوفي ٢٠١٧ و٢٠٢٤. رتّبها للزوار زمنيًا: " +
      "قبل الكلية بثوانٍ تظهر حبات بيلي ثم الخاتم الماسي — وهنا ما زالت النظارة على العين حتى يختفي آخر ضوء. " +
      "ثم الكروموسفير الوردي والشواظ عند حافة القمر، ثم الهالة البيضاء الممتدة ملايين الكيلومترات. " +
      "وفي نهاية الكلية يتكرر الخاتم الماسي في الجهة المقابلة — وهو إشارة لبس النظارة فورًا.");
  }

  figText(["دورة ساروس: الكسوف يعود بعد ١٨ سنة", "The Saros cycle: eclipses repeat every 18 years"], "saros", [
    ["٦٥٨٥٫٣ يومًا (Saros = 223 synodic months): ", "١٨ سنة و١١ يومًا و٨ ساعات تقريبًا، تعود بعدها الشمس والقمر والعقدة إلى الوضع نفسه"],
    ["الثلث يوم: ", "بسبب الساعات الثماني يقع الكسوف التالي أبعد غربًا بنحو ١٢٠° طولًا"],
    ["٥٤ سنة وشهر: ", "بعد ثلاث دورات (Exeligmos) يعود الكسوف قريبًا من خط الطول نفسه"],
    ["سلسلة ساروس ١٣٦ (Saros 136): ", "كسوف ٢٠٢٧ من عائلة كسوف ٢٠ يونيو ١٩٥٥ أطول كسوف في القرن العشرين؛ والتالي ١٢ أغسطس ٢٠٤٥"],
  ], "عرف البابليون هذه الدورة من سجلات الكسوف المتراكمة، وسمّاها إدموند هالي «ساروس». " +
    "أعضاء سلسلة ساروس ١٣٦ المعاصرون: ١٩٥٥ (٧ د ٠٨ ث)، ١٩٧٣ (الكسوف الذي طاردته طائرة الكونكورد)، ١٩٩١، ٢٠٠٩، ثم ٢٠٢٧ (٦ د ٢٣ ث)، ثم ٢٠٤٥. " +
    "والمدد تتناقص ببطء في هذه السلسلة. وتذكر رواية «أرصاد وأطوار» أن الحُسّاب في القاهرة الفاطمية كانوا يعرفون هذا «الدور» الذي تعود به الكسوفات.",
    { cap: "الشكل ٢٢-١ من الكتاب: الأشهر الثلاثة المتوافقة في دورة ساروس" });

  // native chart: eclipse statistics 1901–2050
  {
    const s = add("CONTENT", ["كسوف الشمس في العالم بالأرقام (١٩٠١–٢٠٥٠)", "Solar eclipses worldwide in numbers, 1901–2050"]);
    s.addChart(pres.charts.BAR, [{ name: "عدد الكسوفات", labels: ["جزئي", "حلقي", "كلي", "هجين"], values: [120, 106, 99, 11] }], {
      x: M, y: 1.5, w: 6.9, h: 5.0, barDir: "col", chartColors: ["8A8984", "E88A2C", "0F1B2D", "1D4E89"],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelColor: K.ink, dataLabelFontFace: "+mn-lt",
      catAxisLabelFontSize: 15, catAxisLabelColor: K.ink2, catAxisLabelFontFace: "+mn-lt",
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
      showTitle: true, title: "عدد كسوفات الشمس حسب النوع — ٣٣٦ كسوفًا في ١٥٠ سنة", titleFontSize: 14, titleColor: K.ink2,
      titleFontFace: "+mn-lt", barGapWidthPct: 60, varyColors: true, objectName: "chart-eclipse-types",
    });
    const stats = [["≈ ٢٫٢", "كسوف للشمس كل سنة في المتوسط"], ["كل ١٨ شهرًا", "يقع كسوف كلي في مكان ما على الأرض"],
      ["≈ ٣٧٥ سنة", "متوسط عودة الكسوف الكلي إلى النقطة نفسها"]];
    stats.forEach(([big, lab], i) => {
      const y = 1.55 + i * 1.68;
      card(s, 7.9, y, W - M - 7.9, 1.5, { name: "stat-" + i });
      s.addText(big, { x: 8.1, y: y + 0.12, w: W - M - 8.3, h: 0.75, fontSize: 30, bold: true, color: i === 2 ? K.red : K.sunDeep,
        isTextBox: true, ...R, margin: 0 });
      s.addText(lab, { x: 8.1, y: y + 0.85, w: W - M - 8.3, h: 0.5, fontSize: 15, color: K.ink2, isTextBox: true, ...R, margin: 0 });
    });
    caption(s, "المصدر: حسابات الكتاب من تقويم JPL DE421؛ ومتوسط العودة إلى النقطة نفسها تقدير جان ميوس (Meeus)");
    s.addNotes("بين ١٩٠١ و٢٠٥٠ وقع ٣٣٦ كسوفًا للشمس بحسب حسابات الكتاب: ١٢٠ جزئيًا، و١٠٦ حلقيًا، و٩٩ كليًا، و١١ هجينًا. " +
      "أي أن الكسوف الكلي يقع في مكان ما من الأرض كل سنة ونصف تقريبًا، لكنه غالبًا فوق المحيطات أو مناطق نائية. " +
      "أما النقطة الواحدة فلا يمر بها الكسوف الكلي إلا مرة كل ٣٧٥ سنة في المتوسط (تقدير الفلكي جان ميوس)، " +
      "مع تفاوت كبير: بعض المدن تراه مرتين في عقود قليلة، وبعضها لا يراه ألف عام.");
  }

  // lunar periodicity + upcoming lunar eclipses from Cairo (native table)
  {
    const s = add("CONTENT", ["الخسوف: أكثر حظًّا لأهل المكان الواحد", "Lunar eclipses: periodicity and upcoming eclipses seen from Cairo"]);
    s.addText(bullets([
      ["نصف الأرض يراه (Night side): ", "كل من كان القمر فوق أفقه يرى الخسوف، فلا يحتاج إلى السفر"],
      ["الخسوف الكلي (Total lunar eclipse): ", "يراه المكان الواحد مرة كل سنتين إلى ثلاث في المتوسط، مقابل قرون للكسوف الكلي"],
      ["القمر الدموي (Blood Moon): ", "لا يختفي القمر في الكلية بل يحمرّ بضوء الشروق والغروب المنكسر في جو الأرض"],
      ["تدريب للمرشد: ", "الخسوفان الكليان القادمان فوق مصر فرصة لتدريب المجموعات قبل كل موسم"],
    ], { run: { fontSize: 15, color: K.ink }, gap: 8 }), { x: 7.65, y: 1.5, w: W - M - 7.65, h: 5.0, valign: "top", isTextBox: true, margin: 0, ...R });
    const hdr = (x) => ({ text: x, options: { bold: true, color: "FFFFFF", fill: { color: K.navy }, align: "center", fontSize: 13, ...R, align: "center" } });
    const cell = (x, o = {}) => ({ text: x, options: { fontSize: 13, color: K.ink, align: "center", ...R, align: "center", ...o } });
    // RTL: first column on the right → columns listed right-to-left reversed
    const rows = [
      ["الرؤية من القاهرة", "مدة الكلية", "الذروة (القاهرة)", "النوع", "التاريخ"],
      ["القمر فوق الأفق منذ أول المساء", "—", "٢١:٢٠", "جزئي صغير", "٦ يوليو ٢٠٢٨"],
      ["القمر منخفض في الشرق بعد الغروب", "٧٢ دقيقة", "١٨:٥٣", "كلي", "٣١ ديسمبر ٢٠٢٨"],
      ["القمر تحت الأفق وقت الذروة", "١٠٣ دقائق", "٠٦:٢٢", "كلي عميق", "٢٦ يونيو ٢٠٢٩"],
      ["ممتاز: القمر قرب سمت الرأس", "٥٥ دقيقة", "٠٠:٤٣", "كلي", "٢٠/٢١ ديسمبر ٢٠٢٩"],
      ["جيد بعد الغروب", "—", "٢١:٣٣", "جزئي", "١٥ يونيو ٢٠٣٠"],
    ].map((r) => r.slice().reverse());
    const tbl = rows.map((r, i) => r.map((c, j) => i === 0 ? hdr(c) : cell(c, j === 0 ? { bold: true, color: K.navy } : j === 2 && /^كلي/.test(c) ? { bold: true, color: K.red } : {})));
    s.addTable(tbl, { x: M, y: 1.55, w: 6.75, colW: [1.35, 0.9, 1.15, 1.05, 2.3], rowH: 0.62, border: { type: "solid", color: K.line, pt: 0.75 },
      fill: { color: K.card }, valign: "middle", objectName: "lunar-table" });
    caption(s, "حسابات خاصة بهذه المحاضرة بأدوات الكتاب (DE421)؛ الأوقات بتوقيت القاهرة: شتاءً ت.ع+٢ وصيفًا ت.ع+٣ بافتراض استمرار التوقيت الصيفي", 6.0, { x: M, w: 6.75 });
    s.addNotes("الحساب أُجري خصيصًا لهذه المحاضرة بالأدوات نفسها التي حُسب بها الكتاب. " +
      "أفضل خسوف قادم للتدريب: ليلة ٢٠ إلى ٢١ ديسمبر ٢٠٢٩، الذروة نحو ٠٠:٤٣ بتوقيت القاهرة والقمر على ارتفاع نحو ٧٧°، والكلية ٥٥ دقيقة. " +
      "وخسوف ٣١ ديسمبر ٢٠٢٨ — ليلة رأس السنة — كلي لمدة ٧٢ دقيقة والقمر يرتفع في الشرق بعد الغروب، وهي فرصة سياحية ممتازة. " +
      "أما خسوف ٢٦ يونيو ٢٠٢٩ فهو من أعمق الخسوفات (قدر ١٫٨٥) لكن ذروته تقع والقمر تحت أفق القاهرة صباحًا. " +
      "الخسوفان الجزئيان ٦ يوليو ٢٠٢٨ (قدر ٠٫٣٩) و١٥ يونيو ٢٠٣٠ (قدر ٠٫٥١). الأوقات تقريبية بدقيقة.");
  }

  breakSlide("استراحة بعد المرحلة الأولى.", "الكسوف في تاريخ مصر والحضارة الإسلامية");

  sectionSlide("٢", ["المرحلة الثانية: الكسوف في تاريخ مصر والحضارة الإسلامية", "Stage 2 — Eclipses in Egypt's history and in Islamic civilization"], "الساعة الثانية (١:٠٠ – ١:٥٠): من رع وإخناتون إلى ابن يونس، ومن الحديث النبوي إلى صلاة الكسوف",
    "المرحلة الثانية (نحو خمسين دقيقة) في جزأين: (أ) الكسوفات التي حدثت في مصر من الأسطورة إلى الحساب، (ب) الكسوف والخسوف في الحضارة الإسلامية. تنبيه منهجي: لا يوجد نص مصري قديم معروف يؤرخ كسوفًا بعينه تأريخًا قاطعًا؛ فكل ربط بين نص وكسوف اجتهاد، والحساب يبيّن ما هو ممكن.");

  partSlide("٢-أ", "الكسوفات التي حدثت في مصر", "Eclipses that happened over Egypt", "من الأسطورة وسجلات العالم القديم إلى إخناتون ودندرة، فالقاهرة الفاطمية، فبعثات سوهاج والسلوم", "الجزء الأول من المرحلة الثانية: نبدأ بسياق عالمي قصير عن سجلات الكسوف، ثم ننتقل إلى مصر.");

  figText(["سجلات الكسوف في العالم القديم", "Eclipse records of the ancient world"], "ancientRecords", [
    ["الصين: ", "عظام العرافة (Oracle bones) زمن أسرة شانغ نحو ١٢٠٠ ق.م، وقصة الفلكيين «هسي وهو»"],
    ["أوغاريت: ", "لوح طيني يصف اختفاء الشمس نهارًا، يُربط بكسوف نحو ١٢٢٣ ق.م (والتاريخ مختلف فيه)"],
    ["آشور: ", "كسوف ١٥ يونيو ٧٦٣ ق.م في «سجلات الأعوام» (Eponym chronicle): مرتكز تأريخ الشرق القديم"],
    ["الإغريق: ", "كسوف ٢٨ مايو ٥٨٥ ق.م أوقف معركة الميديين والليديين، ونُسب التنبؤ به إلى طاليس"],
    ["بابل: ", "رصد الكهنة الكسوفات قرونًا فاكتشفوا تكرارها كل ١٨ سنة (ساروس)"],
  ], "مدخل مهم للأثريين: الكسوف من أدق أدوات التأريخ، لأن الحساب الفلكي يعطي يومه وساعته بلا لبس. " +
    "كسوف ٧٦٣ ق.م المسجل في قوائم الحكام الآشوريين (في سنة بور-ساجالي) هو المسمار الذي تُعلَّق عليه أزمنة الشرق القديم كله. " +
    "وقصة طاليس يرويها هيرودوت، وكسوف ثوسيديدس (٣ أغسطس ٤٣١ ق.م) كان حلقيًا رآه أهل أثينا جزئيًا ٨٣٪. " +
    "مصدر الجدول: كتاب «الكسوفات» لسيفاريال (١٩١٥) بعد تدقيق الكتاب له بالحساب؛ مع التنبيه إلى أن الجانب التنجيمي في ذلك الكتاب معتقد قديم لا علم.",
    { split: 7.0, fs: 15, cap: "الشكل ١٣-٢ من الكتاب: أشهر سجلات الكسوف المكتوبة في العالم القديم" });

  figText(["رع وأبوفيس: الثعبان الذي يحاول ابتلاع الشمس", "Ra and Apophis: the serpent that tries to swallow the Sun"], "raApophis", [
    ["الأسطورة: ", "يبحر رع في قاربه كل ليلة عبر العالم السفلي، ويحاول الثعبان أبوفيس ابتلاعه"],
    ["قراءة محتملة: ", "يربط بعض الباحثين الأسطورة بالظواهر التي تحجب الشمس كالكسوف والعواصف"],
    ["للمرشد: ", "استعملها مدخلًا للحكاية، مع التنبيه إلى أن الربط بالكسوف تفسير حديث لا نص صريح"],
  ], "عند الحديث في مقابر وادي الملوك (حيث كتب العالم الآخر: الأمدوات وكتاب البوابات) يمكن ربط رحلة رع الليلية بالحدث: " +
    "في ٢ أغسطس ٢٠٢٧ يتحول نهار وادي الملوك إلى ليل لمدة ٦ دقائق و٢٠ ثانية. " +
    "والرواية «أرصاد وأطوار» تسمي أحد فصولها «اليوم الذي ابتلع فيه الثعبان الشمس كلها» على هذا المعنى.",
    { cap: "الشكل ٢-٣ من الكتاب" });

  figText(["«كسوف إخناتون»: ١٤ مايو ١٣٣٨ ق.م فوق العمارنة", "The “Akhenaten eclipse”: 14 May 1338 BC over Amarna"], "akhenaten", [
    ["الحساب (Retrocalculation): ", "كسوف كلي مرّ ظله فوق أخت-آتون (تل العمارنة) بعد الظهر بقليل، والشمس على ارتفاع ٥٧°"],
    ["المدة: ", "نحو ٥ دقائق ونصف بحساب الكتاب، ونحو ٤ دقائق ونصف في الرواية"],
    ["عدم اليقين (ΔT, Delta T): ", "خطأ في فرق التوقيت بنحو ٨٠٠ ثانية يجعل المدة بين ٤ د ٤٧ ث و٦ د ١٣ ث، وتبقى العمارنة داخل الكلية"],
    ["تنبيه: ", "ربطه بالسنة ١٦ من حكم إخناتون اختيار روائي، لا يثبته نص مصري مؤرخ"],
  ], "هذه من أقوى الحكايات لمرشدي المنيا وتل العمارنة: عاصمة إله قرص الشمس آتون شهدت كسوفًا كليًا للشمس. " +
    "مصدرها رواية «أرصاد وأطوار» التي يقول مؤلفها في ملحقه إن الحساب الفلكي الحديث يثبت وقوع الكسوف، وإن ربطه بالسنة السادسة عشرة غير يقيني. " +
    "وأعاد الكتاب الحساب بنظرية ERFA مع معادلات إسبيناك ومييس لفرق التوقيت. " +
    "وفي ٢٠٢٧ يعود الظل إلى تل العمارنة: الكلية هناك ٥ دقائق و٣٣ ثانية.",
    { split: 7.0, cap: "الشكل ٤-٣ من الكتاب (محسوب، بالتقويم اليولياني)" });

  figText(["سقف دندرة: كسوف وخسوف منقوشان في الحجر", "The Dendera zodiac: a solar and a lunar eclipse carved in stone"], "dendera", [
    ["السقف الدائري (Dendera zodiac): ", "بطلمي من نحو ٥٠ ق.م، نُقل إلى باريس سنة ١٨٢١ وهو اليوم في اللوفر، وفي المعبد نسخة مصبوبة"],
    ["قرص الكسوف: ", "إيزيس تمسك قردًا من ذيله ← كسوف ٧ مارس ٥١ ق.م (حلقي، رآه أهل دندرة جزئيًا ١٧٪ عصرًا)"],
    ["قرص الخسوف: ", "عين حورس ← خسوف كلي للقمر ٢٥ سبتمبر ٥٢ ق.م (قدر ١٫٨١)"],
    ["٢٠٢٧: ", "الكلية في دندرة ٦ دقائق و١٤ ثانية"],
  ], "القراءة للباحثين المحدثين، والحساب في الكتاب أكد أن الحدثين وقعا وكانا مرئيين من صعيد مصر. " +
    "قف بالزوار في معبد دندرة أمام النسخة المصبوبة في سقف المقصورة، ثم أخبرهم أن الظل الذي نقش الكهنة رمزه سيمر فوق رؤوسهم بعد أكثر من ألفي سنة. " +
    "وتنقل الرواية أن كهنة دندرة كانوا يقولون في نقوشهم إنهم ينسخون عن «كتاب قديم من زمن الأولين».",
    { split: 7.3, cap: "الشكل ٥-٣ من الكتاب (محسوب)" });

  // novel eclipses: two figures side by side
  {
    const s = add("CONTENT", ["سماء القاهرة الفاطمية كما في رواية «أرصاد وأطوار»", "The sky of Fatimid Cairo in the novel “Arsad wa Atwar”"]);
    addImg(s, "e993", { x: 6.95, y: 1.45, w: 5.8, h: 4.3 }, { frame: true });
    addImg(s, "e1004", { x: M, y: 1.45, w: 5.8, h: 4.3 }, { frame: true });
    s.addText([t("٢٠ أغسطس ٩٩٣م (٢٩ جمادى الآخرة ٣٨٣هـ): ", { bold: true, color: K.navy }),
      t("كسوف كلي مرّ ظله قبالة الساحل الشمالي، وبلغ في القاهرة ٩٦٪ ضحى", { color: K.ink })],
      { x: 6.95, y: 5.85, w: 5.8, h: 0.75, fontSize: 14, isTextBox: true, margin: 0, valign: "top" });
    s.addText([t("٢٤ يناير ١٠٠٤م (٢٩ ربيع الأول ٣٩٤هـ): ", { bold: true, color: K.navy }),
      t("كسوف هجين، غربت فيه الشمس على القاهرة مكسوفة ٩٧٪ خلف الأهرام", { color: K.ink })],
      { x: M, y: 5.85, w: 5.8, h: 0.75, fontSize: 14, isTextBox: true, margin: 0, valign: "top" });
    s.addNotes("تجري أحداث الرواية في قاهرة العزيز بالله والحاكم بأمر الله، ويفصل مؤلفها في ملحق ختامي بين التاريخي والمتخيَّل. " +
      "أخذنا الأحداث الفلكية التاريخية فقط وأعاد الكتاب حسابها، فجاءت كلها صحيحة. " +
      "كسوف ٩٩٣م: في الرواية يرصده العزيز بالله في طست ماء في صحن القصر ومعه ابنه الحاكم ابن ثماني سنين، ويصل الخبر من الإسكندرية بأن الناس رأوا النجوم ضحى وصلّوا صلاة الكسوف. " +
      "كسوف ١٠٠٤م: يصعد الحاكم المقطم ليرى الشمس تغرب مكسوفة خلف الأهرام «قرنين من نار». " +
      "وبينهما كسوف ٧ أبريل ١٠٠٠م الجزئي (نحو نصف القرص قرب الظهر) الذي يعلّم فيه ابن يونس تلميذه الرصد في الطست. " +
      "تنبيه للمرشد: الشخصيات المتخيلة (سيف، ريحانة…) والمنامات والنبوءات من خيال الرواية؛ أما التواريخ الفلكية فحقيقية ومحسوبة.");
  }

  figText(["خسوف ١ مارس ١٠٠٢م: ظل الأرض المستدير", "The total lunar eclipse of 1 March 1002: Earth's round shadow"], "lunarCairo", [
    ["الحدث: ", "خسوف كلي للقمر ليلة ١٤ ربيع الآخر ٣٩٢هـ، رُئي كاملًا من القاهرة"],
    ["الدرس — ظل الأرض (Earth's umbra): ", "حافة ظل الأرض على القمر مستديرة دائمًا، والجسم الوحيد الذي يلقي ظلًا مستديرًا من كل جهة هو الكرة"],
    ["لماذا يحمرّ؟ ", "في الرواية يقول ابن يونس للخليفة: «لا أعلم» — ونعرف اليوم أنه ضوء الشمس المنكسر في جو الأرض"],
    ["خسوفات أخرى: ", "٦ نوفمبر ٩٧٩م (ربيع الآخر ٣٦٩هـ) و٢٢ أبريل ٩٨١م (شوال ٣٧٠هـ)، جزئيان"],
  ], "مشهد الخسوف في الرواية من أجمل ما يُروى للزوار ليلًا على الأسطح والمراكب النيلية: " +
    "«هي مستديرة في كل خسوف رأيته، والقمر عالٍ أو منخفض، في الشرق أو في الغرب. وليس في الدنيا جسم يلقي ظلًا مستديرًا من كل جهة إلا الكرة. هذا ظلنا يا سيف». " +
    "وهو استدلال قديم على كروية الأرض يعود إلى أرسطو، وكان معروفًا لفلكيي الحضارة الإسلامية. " +
    "فرق صغير: تؤرخ الرواية خسوف سنة ٣٧٠هـ بالعاشر من شوال، والحساب يضعه في الرابع عشر منه، لأن الخسوف لا يقع إلا والقمر بدر.",
    { dark: false, split: 7.5, cap: "الشكل ١٢-٣ من الكتاب: خسوفات القمر كما رُئيت من القاهرة (محسوبة)" });

  figText(["الرواية والحساب: جدول التحقق", "Novel vs. computation: verification table"], "novelTable", [
    ["كل الأحداث صحيحة: ", "ستة كسوفات وخسوفات في القاهرة الفاطمية، وكسوف إخناتون، وقرصا دندرة"],
    ["فروق صغيرة: ", "تاريخ خسوف ٣٧٠هـ، ومدة كلية إخناتون، وثانيتان في مدة الأقصر ٢٠٢٧"],
    ["أحداث سماوية أخرى: ", "قران الزهرة وعطارد ١٩ مايو ١٠٠٠م رصده ابن يونس، والمستعر الأعظم ٣٠ أبريل ١٠٠٦م رصده علي بن رضوان"],
    ["للمرشد: ", "الرواية مدخل قصصي ممتاز، مع الفصل الواضح بين التاريخي والمتخيل"],
  ], "المستعر الأعظم (سوبرنوفا) سنة ١٠٠٦م في برج العقرب هو ألمع نجم مستعر في التاريخ المدوَّن، " +
    "ورآه علي بن رضوان المصري من القاهرة وهو في الثامنة عشرة، فوصف سطوعه بأنه مرتان ونصف إلى ثلاثة أمثال الزهرة، وضوءه أكثر من ربع ضوء القمر. " +
    "هذه الأحداث مذكورة في ملحق المؤلف «ما هو حقيقي في هذا الكتاب». " +
    "اقترح على الزوار المهتمين قراءة الرواية قبل الرحلة.",
    { split: 7.3, fs: 15, cap: "الشكل ١٣-٣ من الكتاب: الكسوفات والخسوفات الحقيقية في الرواية مقارنة بالحساب" });

  figText(["مصر الحديثة وبعثات الكسوف", "Modern Egypt and eclipse expeditions"], "sohag1882", [
    ["سوهاج ١٨٨٢: ", "بعثات أوروبية ترصد الكلية فتكشف مذنبًا قرب الشمس لم يُرَ قبلها ولا بعدها، سُمّي «مذنب توفيق»"],
    ["أسوان ١٩٠٥: ", "آخر كسوف كلي فوق أسوان (٣٠ أغسطس)"],
    ["السلوم ٢٠٠٦: ", "كسوف ٢٩ مارس الكلي جمع الآلاف في الصحراء الغربية"],
    ["٢٠٢٧: ", "أطول وأوسع من ٢٠٠٦، ويمر بقلب مصر السياحية"],
  ], "كسوف ١٧ مايو ١٨٨٢ رُصد من سوهاج، وظهر في صور الكلية مذنب ساطع قرب الشمس سُمّي باسم الخديوي توفيق، ولم يُشاهد إلا في تلك الدقائق. " +
    "وكسوف السلوم ٢٠٠٦ مقارنة مفيدة للزوار الذين يسألون عن خبرة مصر في استقبال سياحة الكسوف.",
    { cap: "الشكل ١٥-٣ من الكتاب" });

  figText(["كسوفات صنعت العلم", "Eclipses that made science: helium (1868) and relativity (1919)"], "eddington", [
    ["١٧١٥ — هالي (Halley): ", "أول خريطة تتنبأ بمسار الظل فوق إنجلترا، وجاء التنبؤ قريبًا جدًا من الواقع"],
    ["١٨٥١ — أول صورة: ", "التقط بيركوفسكي في كونيغسبرغ أول صورة فوتوغرافية ناجحة للكلية"],
    ["١٨٦٨ — الهيليوم (Helium): ", "خط أصفر في طيف الشواظ (Spectrum) كشف عنصرًا جديدًا سُمّي باسم الشمس"],
    ["١٩١٩ — أينشتاين: ", "بعثتا إدنغتون قاستا انحناء ضوء النجوم قرب الشمس (١٫٧٥″)، فأيّدتا النسبية العامة (General relativity)"],
    ["١٩٧٣ — الكونكورد: ", "طائرة طاردت الظل ٧٤ دقيقة من الكلية فوق الصحراء الكبرى"],
  ], "رسالة للزائر: الكسوف ليس مشهدًا جميلًا فحسب؛ إنه مختبر طبيعي. فيه اكتُشف عنصر كيميائي، وفيه اختُبرت أشهر نظرية في الفيزياء الحديثة. " +
    "واليوم تصنع المراصد الفضائية كسوفًا صناعيًا بالمِكْسَفة (Coronagraph) مثل SOHO وProba-3، لكن الكسوف الطبيعي ما زال يكشف الهالة الداخلية بوضوح لا تبلغه معظم الأجهزة. " +
    "وكسوف ١٩٧٣ من سلسلة ساروس ١٣٦ نفسها التي ينتمي إليها كسوف ٢٠٢٧.",
    { split: 7.3, fs: 15, cap: "الشكل ١٠-٢ من الكتاب: تجربة إدنغتون ١٩١٩" });

  figText(["على المكان نفسه: متى رأت مدن مصر الكسوف الكلي؟", "Recurrence at one place: when did Egypt's cities last see totality?"], "lastTotal", [
    ["القاهرة والإسكندرية: ", "آخر كلية في ٦ أكتوبر ١٢٤١م — وفي ٢٠٢٧ جزئي فقط"],
    ["الأقصر: ", "آخر كلية في ١٠ يناير ١٣٧٧م — ثم ٢ أغسطس ٢٠٢٧"],
    ["أسوان: ", "آخر كلية في ٣٠ أغسطس ١٩٠٥ — وفي ٢٠٢٧ جزئي ٩٩٫٨٪"],
    ["السلوم: ", "رأت الكلية في ٢٩ مارس ٢٠٠٦ (٤ دقائق تقريبًا)"],
  ], "هذه الشريحة تجيب عن سؤال «نفس المكان». " +
    "التواريخ قبل ١٩٠٠ تقريبية لأن فرق التوقيت ΔT (تباطؤ دوران الأرض) غير معروف بدقة في العصور الوسطى؛ " +
    "والتواريخ قبل ١٥٨٢ بالتقويم اليولياني. والفكرة للزائر: «أنتم تشهدون ما لم يشهده أهل الأقصر منذ أكثر من ستة قرون». " +
    "وكذلك تذكر الرواية (الفصل الأخير) أن آخر ظل مر بطيبة كان في العاشر من يناير ١٣٧٧م — متفقة مع الحساب.",
    { split: 7.6, cap: "الشكل ٢٣-٣ من الكتاب (محسوب: ERFA قبل ١٩٠٠ وDE421 بعدها)" });

  partSlide("٢-ب", "الكسوف والخسوف في الحضارة الإسلامية", "Eclipses in Islamic civilization", "من حديث «لا ينكسفان لموت أحد» وصلاة الكسوف إلى أرصاد ابن يونس وابن الهيثم في القاهرة", "الجزء الثاني من المرحلة الثانية: ثلاث رسائل — الإسلام نزع عن الكسوف صفة النذير، وشرع له عبادة، ثم ازدهر رصده وحسابه علميًا في بغداد والرقة والقاهرة وغزنة.");

  // hadith + computed eclipse of 632
  {
    const s = add("CONTENT", ["«لا ينكسفان لموت أحد ولا لحياته»", "“They are not eclipsed for anyone's death or life” — the eclipse of 632 CE"]);
    card(s, 5.35, 1.5, W - M - 5.35, 3.05, { fill: K.navy, line: K.navy, name: "hadith-card", flat: true });
    s.addText([
      t("«إِنَّ الشَّمْسَ وَالْقَمَرَ آيَتَانِ مِنْ آيَاتِ اللَّهِ، لَا يَنْكَسِفَانِ لِمَوْتِ أَحَدٍ وَلَا لِحَيَاتِهِ، فَإِذَا رَأَيْتُمْ ذَلِكَ فَادْعُوا اللَّهَ وَكَبِّرُوا وَصَلُّوا وَتَصَدَّقُوا»",
        { color: "FFFFFF", fontSize: 21, breakLine: true }),
      t("صحيح البخاري، من حديث عائشة رضي الله عنها", { color: K.sun, fontSize: 14 }),
    ], { x: 5.65, y: 1.65, w: W - M - 5.95, h: 2.75, valign: "middle", isTextBox: true, margin: 0, paraSpaceAfter: 10 });
    s.addText(bullets([
      ["المناسبة: ", "انكسفت الشمس يوم مات إبراهيم ابن النبي ﷺ، فقال الناس: انكسفت لموته"],
      ["الرسالة: ", "قطع الصلة بين الكسوف وأحداث البشر — موقف علمي سابق لعصره"],
      ["للمرشد: ", "جواب جاهز لكل سؤال عن «نحس» الكسوف أو تأثيره في المصائر"],
    ], { run: { fontSize: 16, color: K.ink }, gap: 10 }), { x: 5.35, y: 4.75, w: W - M - 5.35, h: 1.85, valign: "top", isTextBox: true, margin: 0, ...R });
    // computed box
    card(s, M, 1.5, 4.5, 5.1, { name: "calc-632" });
    disc(s, M + 3.55, 1.72, 0.7, "", { name: "calc-disc", ring: 2 });
    s.addText("ماذا يقول الحساب؟", { x: M + 0.25, y: 1.75, w: 3.2, h: 0.6, fontSize: 19, bold: true, color: K.blue, isTextBox: true, ...R, margin: 0 });
    const facts = [["التاريخ", "٢٧ يناير ٦٣٢م (أواخر شوال ١٠هـ)"], ["النوع", "كسوف حلقي مرّ مساره جنوب الجزيرة"],
      ["في المدينة", "جزئي، حُجب نحو ٧٧٪ من قرص الشمس"], ["الوقت", "ضحى، نحو الثامنة صباحًا بالتوقيت المحلي"]];
    facts.forEach(([k, v], i) => {
      const y = 2.6 + i * 0.95;
      s.addText(k, { x: M + 0.25, y, w: 4.0, h: 0.35, fontSize: 13, bold: true, color: K.sunDeep, isTextBox: true, ...R, margin: 0 });
      s.addText(v, { x: M + 0.25, y: y + 0.33, w: 4.0, h: 0.55, fontSize: 15, color: K.ink, isTextBox: true, ...R, margin: 0, valign: "top" });
    });
    caption(s, "حساب خاص بالمحاضرة (ERFA + ΔT لإسبيناك ومييس)، تقريبي لبعد العهد", 6.62, { x: M, w: 4.5 });
    s.addNotes("نص الحديث في صحيح البخاري (كتاب الكسوف) من حديث عائشة رضي الله عنها، وله روايات أخرى عن المغيرة بن شعبة وأبي مسعود. " +
      "حسبنا لهذه المحاضرة الكسوف المقترن بوفاة إبراهيم ابن النبي ﷺ: كسوف حلقي في ٢٧ يناير ٦٣٢م، رآه أهل المدينة جزئيًا بنسبة احتجاب نحو ٧٧٪ " +
      "ضحىً والشمس على ارتفاع نحو ١٦° (نحو الثامنة صباحًا بالتوقيت الشمسي المحلي). ويوافق بالتقويم الحسابي ٢٨ أو ٢٩ شوال سنة ١٠هـ. " +
      "الحساب تقريبي لأن ΔT في القرن السابع غير معروف بدقة، لكن احتجاب نحو ثلاثة أرباع القرص لا يتغير كثيرًا. " +
      "الرسالة للزائر الأجنبي: قبل أربعة عشر قرنًا أعلن نبي الإسلام أن الكسوف آية كونية لا علاقة لها بموت أحد أو حياته.");
  }

  // prayer + Dhuhr coincidence
  {
    const s = add("CONTENT", ["صلاة الكسوف والخسوف — ووقت الظهر يوم ٢٠٢٧", "The eclipse prayer (Salat al-Kusuf) and Dhuhr time on eclipse day"]);
    const steps = [["سنة مؤكدة", "صلاة جماعة في المسجد عند حدوث الكسوف أو الخسوف، ووقتها من بدئه حتى ينجلي"],
      ["صفتها", "ركعتان، في كل ركعة قيامان وركوعان عند الجمهور، مع إطالة القراءة"],
      ["مع الصلاة", "الدعاء والتكبير والصدقة، وخطبة يُذكَّر فيها الناس، كما في الحديث"],
      ["المصطلح", "يُستعمل اللفظان في اللغة للنيّرين، واشتهر: الكسوف للشمس والخسوف للقمر «وَخَسَفَ الْقَمَرُ» (القيامة: ٨)"]];
    steps.forEach(([k, v], i) => {
      const y = 1.5 + i * 1.12;
      card(s, 6.15, y, W - M - 6.15, 1.0, { name: "prayer-" + i });
      s.addText(k, { x: W - M - 1.9, y: y + 0.1, w: 1.7, h: 0.8, fontSize: 16, bold: true, color: K.navy, isTextBox: true, ...R, margin: 0, valign: "middle" });
      s.addText(v, { x: 6.35, y: y + 0.08, w: W - M - 6.35 - 2.0, h: 0.85, fontSize: 14, color: K.ink, isTextBox: true, ...R, margin: 0, valign: "middle" });
    });
    card(s, M, 1.5, 5.25, 4.35, { fill: K.navy, line: K.navy, name: "dhuhr-card", flat: true });
    s.addText("انتبه: الكلية بعد أذان الظهر بدقائق", { x: M + 0.3, y: 1.7, w: 4.65, h: 0.55, fontSize: 18, bold: true, color: K.sun, isTextBox: true, ...R, margin: 0 });
    const rows = [["الأقصر", "١٢:٥٦", "١٣:٠١:٥٨"], ["سوهاج", "١٢:٥٩", "١٢:٥٨:٣٨"], ["أسوان", "١٢:٥٥", "جزئي ٩٩٫٨٪"], ["القاهرة", "١٣:٠١", "جزئي ٩٤٫٨٪"]];
    const head = ["المدينة", "الزوال", "بداية الكلية"];
    [head, ...rows].forEach((r, i) => {
      const y = 2.4 + i * 0.6;
      r.forEach((c, j) => s.addText(c, { x: M + 0.3 + (2 - j) * 1.55, y, w: 1.55, h: 0.5, fontSize: i === 0 ? 13 : 15,
        bold: i === 0 || j === 0, color: i === 0 ? "8C9AB3" : (j === 2 ? K.sun : "FFFFFF"), align: "center", isTextBox: true, rtlMode: true, margin: 0 }));
    });
    s.addText("الزوال (منتصف النهار الشمسي) محسوب لهذه المحاضرة، ويُرفع الأذان بعده بدقائق.", { x: M + 0.3, y: 5.15, w: 4.65, h: 0.6,
      fontSize: 12, color: K.ice, isTextBox: true, ...R, margin: 0 });
    s.addText("الإثنين ٢ أغسطس ٢٠٢٧ = ٢٨ أو ٢٩ صفر ١٤٤٩هـ (آخر الشهر، كما في كل كسوف)", { x: M, y: 6.0, w: W - 2 * M, h: 0.5,
      fontSize: 15, bold: true, color: K.blue, isTextBox: true, ...R, margin: 0 });
    s.addNotes("نقطة عملية مهمة يغفل عنها كثيرون: الزوال في الأقصر يوم ٢ أغسطس ٢٠٢٧ نحو ١٢:٥٥:٤٥ بتوقيت مصر الصيفي، " +
      "فيُرفع أذان الظهر قرابة ١٢:٥٧–١٣:٠٠، وتبدأ الكلية ١٣:٠١:٥٨. أي أن المساجد ستصلي الظهر ثم صلاة الكسوف في قلب الحدث، " +
      "وستتجمع حشود قرب المساجد والساحات. خطط لمواقع مجموعتك بعيدًا عن الزحام، واشرح للزائر الأجنبي معنى ما يسمعه من تكبير وصلاة. " +
      "التاريخ الهجري: الحساب العددي يعطي ٢٨ صفر، وتذكر الرواية ٢٩ صفر ١٤٤٩هـ؛ والفرق يوم بحسب رؤية الهلال.");
  }

  // Muslim astronomers cards
  {
    const s = add("CONTENT", ["علماء مسلمون رصدوا الكسوف وحسبوه", "Muslim astronomers who observed and computed eclipses"]);
    const ppl = [
      ["البتّاني", "الرقة (ت ٣١٧هـ/٩٢٩م)", "أرصاد مؤرخة للكسوف والخسوف؛ لاحظ تغير القطر الظاهري للشمس والقمر فقال بإمكان الكسوف الحلقي"],
      ["ابن يونس", "القاهرة (ت ٣٩٩هـ/١٠٠٩م)", "«الزيج الحاكمي الكبير» في ٨١ بابًا وأكثر من مئة رصدة مؤرخة؛ أرصاده للكسوف تُستعمل اليوم لقياس تباطؤ دوران الأرض"],
      ["ابن الهيثم", "البصرة والقاهرة (ت نحو ١٠٤٠م)", "«مقالة في صورة الكسوف»: درس صورة الشمس المكسوفة عبر ثقب في الحجرة المظلمة — أصل كاميرا الثقب الآمنة"],
      ["البيروني", "خوارزم وغزنة (ت ٤٤٠هـ/١٠٤٨م)", "أرصاد دقيقة للكسوف، ونبّه إلى خطر النظر المباشر إلى الشمس وفضّل رصد صورتها في الماء"],
    ];
    ppl.forEach(([n, where, what], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const w = 5.85, h = 2.35, x = col === 0 ? W - M - w : M, y = 1.5 + row * (h + 0.25);
      card(s, x, y, w, h, { name: "scholar-" + i });
      disc(s, x + w - 1.0, y + 0.25, 0.7, "", { name: "scholar-disc-" + i, ring: 2 });
      s.addText(n, { x: x + 0.25, y: y + 0.2, w: w - 1.45, h: 0.5, fontSize: 21, bold: true, color: K.navy, isTextBox: true, ...R, margin: 0 });
      s.addText(where, { x: x + 0.25, y: y + 0.68, w: w - 1.45, h: 0.4, fontSize: 13, color: K.sunDeep, bold: true, isTextBox: true, ...R, margin: 0 });
      s.addText(what, { x: x + 0.25, y: y + 1.1, w: w - 0.5, h: 1.15, fontSize: 14, color: K.ink, isTextBox: true, ...R, margin: 0, valign: "top" });
    });
    s.addNotes("أربعة أسماء يحسن أن يحفظها المرشد. ابن يونس وابن الهيثم عاشا في القاهرة الفاطمية نفسها، " +
      "وهما شخصيتان تاريخيتان في رواية «أرصاد وأطوار». يذكر ملحق الرواية أن ابن يونس ترك الزيج الحاكمي في ٨١ بابًا وفيه أكثر من مئة رصدة مؤرخة، " +
      "يميّز رصده من رصد غيره ويكتب أخطاء أسلافه وأخطاء نفسه. " +
      "وأرصاد الكسوف العربية — مع سجلات بابل والصين — هي اليوم من أدق وسائل قياس فرق التوقيت ΔT، أي تباطؤ دوران الأرض عبر القرون.");
  }

  // Ibn Yunus method
  figText(["كيف رصد ابن يونس الكسوف في القاهرة؟", "How Ibn Yunus observed eclipses in Fatimid Cairo"], "ibnYunus", [
    ["طست الماء (Water basin): ", "يُنظر إلى صورة الشمس في طست نحاس مملوء ماءً، لا إليها مباشرة"],
    ["الأصابع (Digits): ", "يُقاس المنكسف بالأصابع: اثنا عشر إصبعًا لقطر القرص كله"],
    ["التوقيت: ", "بارتفاع الشمس بظل شاخص، أو بارتفاع نجم معلوم كالسماك الرامح ليلًا"],
    ["التأريخ مرتين: ", "بالهجري وبسنيّ يزدجرد الفارسية التي لا كبيس فيها"],
    ["الأمانة: ", "«وكانت أصابع الكسوف مساوية لما أوجبه الحساب»، ويثبت خطأه ثم تصحيحه"],
  ], "النصوص منقولة في ملحق رواية «أرصاد وأطوار» عن «الزيج الكبير الحاكمي» كما نشره كوسان دي برسفال في باريس سنة ١٨٠٤ من مخطوط ليدن. " +
    "اقرأ للزوار عبارة ابن يونس: «قد رصدت هذا الكسوف أنا وابني أبو الحسن ومُفلِح، فطلعت الشمس منكسفة… ولم يزل الكسوف يزيد… " +
    "ورصدناها بالماء رصدًا محكمًا، فانجلت فلم يبق فيها من الكسوف شيء». " +
    "وفي الرواية مشهد يصفع فيه الشيخ تلميذه حين رفع عينه إلى الشمس: «في كل بيت رصدٍ أعمى، كان في شبابه أحمق ظنّ أن نظرة واحدة لا تضر» — " +
    "وهي أفضل مقدمة لدرس السلامة. تنبيه: طست الماء لا يُنصح به اليوم، فالانعكاس على الماء ما زال ساطعًا مؤذيًا؛ البديل الحديث كاميرا الثقب والنظارة المعتمدة.",
    { split: 6.9, fs: 15, cap: "الشكل ٧-٣ من الكتاب؛ النصوص من ملحق رواية «أرصاد وأطوار» عن طبعة كوسان دي برسفال (١٨٠٤)" });

  figText(["من ابن الهيثم إلى أهلّة الأشجار: الرصد عبر الثقب", "From Ibn al-Haytham to tree crescents: the camera obscura / pinhole"], "haytham", [
    ["المبدأ — الحجرة المظلمة (Camera obscura): ", "ثقب صغير في جدار حجرة مظلمة يرسم على الحائط المقابل صورة مقلوبة للشمس المكسوفة"],
    ["ابن الهيثم: ", "لاحظ أن صورة الكسوف تحتفظ بشكل الهلال مهما كان شكل الثقب إذا صغر"],
    ["اليوم — كاميرا الثقب (Pinhole): ", "مصفاة المطبخ أو أوراق الشجر ترسم عشرات الأهلّة على الأرض أثناء الجزئي"],
    ["للمرشد: ", "نشاط آمن لكل الأعمار في الساعة الأولى من الكسوف"],
  ], "اربط العلم بالتاريخ: الحجرة المظلمة (Camera obscura) التي درسها ابن الهيثم في القاهرة هي جدّة الكاميرا الحديثة. " +
    "وفي الرواية وصف لظلال سعف النخل وقد صارت «أهلّة صغيرة» على البلاط أثناء كسوف ٩٩٣م. " +
    "اطلب من الزوار في المعابد أن ينظروا إلى ظلال النخيل والأشجار على الأرض قبيل الكلية — سيرون مئات الأهلّة.",
    { cap: "الشكل ٦-٣ من الكتاب: مبدأ الحجرة المظلمة" });

  figText(["فرق التوقيت ΔT: السجلات القديمة تقيس دوران الأرض", "Delta T: ancient records measure Earth's slowing rotation"], "deltaT", [
    ["الفكرة: ", "دوران الأرض يتباطأ ببطء بفعل المد والجزر، فيتراكم فرق بين الزمن المنتظم وزمن الدوران"],
    ["المقدار: ", "نحو ست ساعات في القرن الثامن قبل الميلاد"],
    ["الأثر: ", "الأرض تدور ١٥° كل ساعة، فخطأ ساعة واحدة يزيح مسار الكسوف القديم نحو ١٥° طولًا"],
    ["القيمة الأثرية: ", "مقارنة الكسوف المدوَّن (في بابل والصين وزيج ابن يونس) بالحساب هي أدق مقياس لـ ΔT"],
  ], "هذه نقطة تجمع الأثري والفلكي: السجل التاريخي المؤرخ يصحح الحساب الفلكي، والحساب الفلكي يؤرخ السجل. " +
    "ولهذا تبقى تواريخ ومسارات الكسوفات قبل ١٩٠٠ في الكتاب «تقريبية»، ويُرسم حولها نطاق عدم اليقين كما في كسوف إخناتون. " +
    "وأرصاد ابن يونس في القاهرة الفاطمية من السجلات التي استعملها علماء القرن العشرين في حساب تباطؤ دوران الأرض.",
    { split: 7.3, cap: "الشكل ١٤-٢ من الكتاب: منحنى ΔT عبر ثلاثة آلاف عام (محسوب)" });

  breakSlide("استراحة بعد المرحلة الثانية.", "كسوف ٢ أغسطس ٢٠٢٧: الرصد والتصوير");

  sectionSlide("٣", ["المرحلة الثالثة: كسوف ٢ أغسطس ٢٠٢٧", "Stage 3 — The 2 August 2027 eclipse: timing, observing, photographing"], "الساعة الثالثة (٢:٠٠ – ٣:٠٠): المسار والمواعيد في مصر، المراحل كاملة، شكله في الأقصر، الرصد الآمن والتصوير",
    "المرحلة الثالثة (نحو ساعة) وهي قلب المحاضرة. كل الأرقام من عناصر بِسِل الرسمية التي نشرتها ناسا (فريد إسبيناك، ΔT = ٧٦ ثانية)، وتحقق منها الكتاب مستقلًا من تقويم JPL DE421 فاتفقت المدد في حدود ثلث ثانية.");

  partSlide("٣-أ", "المسار ومواعيده في مصر", "The path and the times across Egypt", "من الأطلسي إلى المحيط الهندي، ثم مدينة مدينة داخل المسار وخارجه", "الجزء الأول من المرحلة الثالثة.");

  // key numbers (native stat callouts)
  {
    const s = add("DARK", ["الحدث بالأرقام", "The eclipse of 2 August 2027 in numbers"]);
    const st = [["٦ د ٢٣ ث", "أطول مدة للكلية (Max. duration)", "غرب طهطا بمحافظة سوهاج"], ["٢٥٩ كم", "أقصى عرض للمسار (Path width)", "قرب خط عرض ٢٢° شمالًا"],
      ["١١ دولة", "يمر بها مسار الكلية", "من إسبانيا إلى الصومال"], ["٨١°", "ارتفاع الشمس (Sun altitude)", "الشمس شبه عمودية فوق الرأس"],
      ["١٫٠٣٩", "قدر الكسوف (Magnitude)", "القمر أكبر ظاهريًا من الشمس بنحو ٤٪"], ["٣ س ٢٣ د", "عمر الظل على الأرض", "من ١١:٢٥ إلى ١٤:٤٨ بتوقيت مصر"]];
    st.forEach(([big, lab, sub], i) => {
      const col = i % 3, row = Math.floor(i / 3);
      const w = 3.85, h = 2.45, x = W - M - w - col * (w + 0.25), y = 1.5 + row * (h + 0.25);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: K.navy2 }, line: { color: K.navy3, width: 0.75 }, objectName: "num-" + i });
      s.addText(big, { x: x + 0.3, y: y + 0.3, w: w - 0.6, h: 0.95, fontSize: 40, bold: true, color: K.sun, isTextBox: true, ...R, margin: 0 });
      s.addText(lab, { x: x + 0.3, y: y + 1.3, w: w - 0.6, h: 0.5, fontSize: 15, bold: true, color: "FFFFFF", isTextBox: true, ...R, margin: 0 });
      s.addText(sub, { x: x + 0.3, y: y + 1.8, w: w - 0.6, h: 0.45, fontSize: 14, color: K.ice, isTextBox: true, ...R, margin: 0 });
    });
    s.addNotes("أرقام يحفظها المرشد: أطول مدة ٦ دقائق و٢٣ ثانية، أقصى عرض للمسار ٢٥٩ كم، ١١ دولة، الشمس على ارتفاع ٨١°. " +
      "كسوف ٢٠٢٧ من أطول الكسوفات الكلية في القرن الحادي والعشرين، وسر طوله ثلاثة: القمر قريب من الحضيض فيبدو كبيرًا، " +
      "والأرض قرب الأوج فتبدو الشمس أصغر قليلًا، ومحور الظل يمر قريبًا من مركز الأرض (غاما ≈ ٠٫١٤)، فيكون الظل أبطأ فوق صعيد مصر. " +
      "الظل يلامس الأرض من ٠٨:٢٥ إلى ١١:٤٨ بالتوقيت العالمي.");
  }

  figText(["المسار من الأطلسي إلى المحيط الهندي", "Path of totality: from the Atlantic to the Indian Ocean"], "regional", [
    ["البداية: ", "فجرًا في المحيط الأطلسي، ثم مضيق جبل طارق وجنوب إسبانيا"],
    ["شمال أفريقيا: ", "شمال المغرب والجزائر وتونس وليبيا"],
    ["الذروة: ", "فوق صعيد مصر قرب الظهر"],
    ["النهاية: ", "السعودية واليمن والصومال، ثم المحيط الهندي"],
  ], "اللون الأزرق: مدة الكلية، والبرتقالي: نسبة تغطية الشمس خارج المسار. " +
    "أمثلة من الدول الأخرى للمقارنة (من ملحق الكتاب): طنجة ٤ د ٥٠ ث، طريفة ٤ د ٣٨ ث، الأقصر ٦ د ٢٠ ث، ومكة جزئي، وجدة قريبة من المسار. " +
    "الرسالة: مصر تحظى بأطول مدة على اليابسة، ومع آثارها وطقسها الصافي في أغسطس تصبح الوجهة الأولى في العالم لهذا الكسوف.",
    { split: 8.2, fs: 15, cap: "خريطة ٢ من أطلس الكتاب (بالعربية، محسوبة)" });

  figText(["رحلة الظل فوق مصر", "The Moon's umbra crossing Egypt"], "snapshots", [
    ["١٢:٤٢: ", "يدخل الظل الكامل مصر من الغرب، فتبدأ الكلية في سيوة"],
    ["١٢:٥٠: ", "الواحات البحرية"],
    ["١٢:٥٧–١٣:٠٢: ", "المنيا وأسيوط وسوهاج والأقصر"],
    ["١٣:٠٦: ", "القصير على البحر الأحمر"],
    ["١٣:١٧: ", "آخر الكلية في برنيس، ثم يعبر الظل البحر الأحمر إلى السعودية"],
  ], "يعبر الظل الكامل مصر من سيوة إلى برنيس في نحو ٣٥ دقيقة فقط. " +
    "لاحظ كيف يصغر الظل ويتباطأ قرب مصر لأن الشمس شبه عمودية فوقها. " +
    "سرعة الظل فوق مصر نحو ألفي كيلومتر في الساعة — أي أن الظلام يأتي من جهة الغرب-الشمال الغربي مسرعًا، فوجّه أنظار مجموعتك نحو تلك الجهة قبيل الكلية لترى «جدار الظل» قادمًا.",
    { split: 7.0, cap: "الشكل ٤-٤ من الكتاب: مواقع الظل كل ربع ساعة (محسوبة)" });

  // Egypt map full width
  {
    const s = add("CONTENT", ["مسار الكلية فوق مصر", "Path of totality over Egypt"]);
    addImg(s, "egypt", { x: M, y: 1.35, w: 8.3, h: 5.3 }, { frame: true });
    s.addText(bullets([
      ["داخل المسار (Path of totality): ", "سيوة، الواحات البحرية، المنيا، أسيوط، سوهاج، قنا، الأقصر، إدفو، كوم أمبو، القصير، مرسى علم، برنيس"],
      ["على الحافة: ", "أسوان ٩٩٫٨٪ والخارجة ٩٩٫٥٪ وسفاجا ٩٩٫٦٪ — جزئي رغم القرب"],
      ["خارجه: ", "القاهرة ٩٤٫٨٪ والإسكندرية ٩٣٫٤٪ والغردقة ٩٨٫٦٪"],
      ["النجمة: ", "نقطة أطول مدة (Greatest duration)؛ والمعيّن: نقطة أعظم كسوف (Greatest eclipse)"],
    ], { run: { fontSize: 15, color: K.ink }, gap: 10 }), { x: 9.15, y: 1.45, w: W - M - 9.15, h: 5.15, valign: "top", isTextBox: true, margin: 0, ...R });
    s.addNotes("خريطة مصر من أطلس الكتاب: الخطان الأزرقان حدّا مسار الكلية، والخط الأحمر المتقطع خط المركز. " +
      "كلما اقتربت من الخط الأحمر طالت الكلية. أسوان تقع خارج الحد الجنوبي بقليل، فلا يراها أهل أسوان كلية — " +
      "انصح مجموعات أسوان بالتوجه شمالًا إلى كوم أمبو (٣ د ٢١ ث) أو إدفو (٥ د ٣٠ ث). " +
      "وأهل القاهرة يحتاجون إلى السفر جنوبًا: أقرب موقع ممتاز الواحات البحرية (نحو ٣٧٠ كم) أو المنيا.");
  }

  figText(["أطول مدة أم أعظم كسوف؟", "Greatest duration vs. greatest eclipse"], "longest", [
    ["أطول مدة (Greatest duration): ", "٦ د ٢٣٫٢ ث عند ٢٦°٤٨٫٨′ ش و٣١°٠٨٫٠′ ق، في الصحراء نحو ٣٦ كم غرب طهطا، الساعة ١٣:٠٠:٢١"],
    ["أعظم كسوف (Greatest eclipse): ", "أقرب مرور لمحور الظل من مركز الأرض، عند ٢٥°٣٠٫٤′ ش و٣٣°١٠٫٨′ ق، نحو ٦٠ كم شرق الأقصر، الساعة ١٣:٠٦:٣٤ ومدته ٦ د ٢٢٫٧ ث"],
    ["المدن: ", "نجع حمادي وسوهاج ٦ د ٢٢ ث، أبيدوس وقوص ٦ د ٢١ ث، الأقصر وبرنيس ٦ د ٢٠ ث"],
    ["الخلاصة: ", "الفرق ثوانٍ لا يدركها الراصد؛ الأهم سماء صافية وأفق مفتوح قرب خط المركز"],
  ], "يكثر الخلط بين النقطتين في الإعلام. ناسا تنشر الاثنتين، وتظهر الثانية على خرائطها مقرّبة إلى ٦ د ٢٣ ث. " +
    "والمنحنى منبسط جدًا: لا يبلغ ٦ د ٢٣ ث كاملة إلا شريط ضيق في الصحراء. " +
    "فإذا سألك زائر «أين أطول كسوف؟» فالجواب الصادق: «الفرق بين هنا وهناك ثانية أو ثانيتان، ولن تشعر به».",
    { split: 6.8, fs: 14, cap: "الشكل ٢٤-٤ من الكتاب (محسوب)" });

  // totality table (native)
  {
    const s = add("CONTENT", ["مواعيد الكلية في مدن المسار", "Totality times in cities inside the path (Egypt summer time, UT+3)"]);
    const data = [
      ["نجع حمادي", "قنا", "١٣:٠٠:٣٣", "١٣:٠٦:٥٦", "٦ د ٢٢ ث", "٨١٫٦°"],
      ["سوهاج", "سوهاج", "١٢:٥٨:٣٨", "١٣:٠٥:٠٠", "٦ د ٢٢ ث", "٨١٫٢°"],
      ["أبيدوس (البلينا)", "سوهاج", "١٢:٥٩:٤١", "١٣:٠٦:٠٢", "٦ د ٢١ ث", "٨١٫٥°"],
      ["الأقصر", "الأقصر", "١٣:٠١:٥٨", "١٣:٠٨:١٨", "٦ د ٢٠ ث", "٨١٫٨°"],
      ["برنيس", "البحر الأحمر", "١٣:١٠:٣٤", "١٣:١٦:٥٤", "٦ د ٢٠ ث", "٨٠٫٨°"],
      ["الباويطي", "الجيزة", "١٢:٥٠:١٠", "١٢:٥٦:٢٨", "٦ د ١٧ ث", "٧٨٫٧°"],
      ["دندرة", "قنا", "١٣:٠١:٢١", "١٣:٠٧:٣٥", "٦ د ١٤ ث", "٨١٫٤°"],
      ["قنا", "قنا", "١٣:٠١:٢٧", "١٣:٠٧:٣٨", "٦ د ١١ ث", "٨١٫٣°"],
      ["أسيوط", "أسيوط", "١٢:٥٦:٤٦", "١٣:٠٢:٥٣", "٦ د ٠٧ ث", "٨٠٫٦°"],
      ["إسنا", "الأقصر", "١٣:٠٢:٤٣", "١٣:٠٨:٢٨", "٥ د ٤٤ ث", "٨٢٫٢°"],
      ["مرسى علم", "البحر الأحمر", "١٣:٠٧:٥٤", "١٣:١٣:٢٥", "٥ د ٣١ ث", "٨٠٫٨°"],
      ["إدفو", "أسوان", "١٣:٠٤:٠١", "١٣:٠٩:٣١", "٥ د ٣٠ ث", "٨٢٫٣°"],
      ["سيوة", "مطروح", "١٢:٤٢:٢٦", "١٢:٤٧:٥٥", "٥ د ٢٨ ث", "٧٥٫٥°"],
      ["المنيا", "المنيا", "١٢:٥٥:٢٨", "١٢:٥٩:٣٧", "٤ د ٠٩ ث", "٧٩٫٦°"],
      ["كوم أمبو", "أسوان", "١٣:٠٦:٠٤", "١٣:٠٩:٢٥", "٣ د ٢١ ث", "٨٢٫٦°"],
      ["القصير", "البحر الأحمر", "١٣:٠٦:٣٠", "١٣:٠٩:٠٢", "٢ د ٣١ ث", "٨٠٫٦°"],
    ];
    const head = ["المدينة", "المحافظة", "بداية الكلية", "نهاية الكلية", "المدة", "ارتفاع الشمس"];
    const half = [data.slice(0, 8), data.slice(8)];
    half.forEach((rows, k) => {
      const x = k === 0 ? 6.75 : M, w = 5.95;
      const cols = [0, 2, 3, 4]; // city, C2, C3, duration
      const tbl = [cols.map((c) => head[c]), ...rows.map((r) => cols.map((c) => r[c]))].map((r) => r.slice().reverse())
        .map((r, i) => r.map((c, j) => ({ text: c, options: { ...R, align: "center", fontSize: 13,
          bold: i === 0 || j === 3 || (j === 0 && i > 0), color: i === 0 ? "FFFFFF" : j === 0 ? K.blue : K.ink,
          fill: { color: i === 0 ? K.navy : (i % 2 ? K.card : K.tint) } } })));
      s.addTable(tbl, { x, y: 1.5, w, colW: [1.3, 1.45, 1.45, 1.75], rowH: 0.52, border: { type: "solid", color: K.line, pt: 0.5 },
        valign: "middle", objectName: "totality-table-" + (k + 1) });
    });
    caption(s, "الأوقات بتوقيت مصر الصيفي (ت.ع +٣)، والراصد عند مستوى سطح البحر؛ قد تختلف التماسات ثواني قليلة بسبب تضاريس حافة القمر", 6.35);
    s.addNotes("هذا الجدول من ملحق الكتاب (جدول المدن المصرية الخمسين). رُتّبت المدن من الأطول كلية إلى الأقصر. " +
      "ارتفاع الشمس عند الذروة: نجع حمادي ٨١٫٦°، الأقصر ٨١٫٨°، إدفو ٨٢٫٣°، سيوة ٧٥٫٥°. " +
      "لاحظ أن مدن البحر الأحمر الجنوبية (مرسى علم وبرنيس والشلاتين ٥ د ٣٦ ث) خيار ممتاز للسياحة الشاطئية، " +
      "أما القصير فعلى حافة المسار (٢ د ٣١ ث) والغردقة وسفاجا خارجه.");
  }

  // partial-eclipse timetable (native)
  {
    const s = add("CONTENT", ["مواعيد الكسوف الجزئي في المدن خارج المسار", "Partial eclipse times in cities outside the path (UT+3)"]);
    const rows = [
      ["أسوان", "١١:٤٢:٤٢", "١٣:٠٨:١٧", "١٤:٢٩:٤٢", "٩٩٫٨٪"],
      ["سفاجا", "١١:٤١:٤٦", "١٣:٠٦:٠٤", "١٤:٢٦:٢٨", "٩٩٫٦٪"],
      ["الخارجة", "١١:٣٥:٥٩", "١٣:٠١:١٢", "١٤:٢٣:٢٣", "٩٩٫٥٪"],
      ["الغردقة", "١١:٤٠:٥٧", "١٣:٠٤:٥٩", "١٤:٢٥:١٧", "٩٨٫٦٪"],
      ["شرم الشيخ", "١١:٤١:٢٦", "١٣:٠٤:٥٨", "١٤:٢٤:٤٤", "٩٥٫٦٪"],
      ["الجيزة", "١١:٣٢:٥٩", "١٢:٥٥:٤٣", "١٤:١٦:٠٧", "٩٥٫٠٪"],
      ["القاهرة", "١١:٣٣:٠١", "١٢:٥٥:٤٣", "١٤:١٦:٠٥", "٩٤٫٨٪"],
      ["الإسكندرية", "١١:٢٩:٣٣", "١٢:٥١:٣٣", "١٤:١١:٥٣", "٩٣٫٤٪"],
      ["أبو سمبل", "١١:٤٢:٢٩", "١٣:٠٨:٣٢", "١٤:٣٠:٢٧", "٩٢٫٠٪"],
      ["بورسعيد", "١١:٣٤:٢٩", "١٢:٥٦:٠٤", "١٤:١٥:١٤", "٨٨٫٥٪"],
    ];
    const head = ["المدينة", "البداية (C1)", "الذروة (Max)", "النهاية (C4)", "الاحتجاب"];
    const tbl = [head, ...rows].map((r, i) => r.slice().reverse().map((c, j) => ({ text: c, options: { ...R, align: "center", fontSize: 14,
      bold: i === 0 || j === 4 || j === 0, color: i === 0 ? "FFFFFF" : j === 4 ? K.blue : j === 0 ? K.sunDeep : K.ink,
      fill: { color: i === 0 ? K.navy : (i % 2 ? K.card : K.tint) } } })));
    s.addTable(tbl, { x: M, y: 1.5, w: 8.2, colW: [1.4, 1.65, 1.65, 1.65, 1.85], rowH: 0.43, border: { type: "solid", color: K.line, pt: 0.5 },
      valign: "middle", objectName: "partial-table" });
    card(s, 9.1, 1.5, W - M - 9.1, 4.75, { fill: K.navy, line: K.navy, name: "partial-note", flat: true });
    s.addText([t("في هذه المدن", { bold: true, color: K.sun, fontSize: 20, breakLine: true }),
      t("لا توجد لحظة كلية، فالنظارة على العين من البداية إلى النهاية.", { color: "FFFFFF", fontSize: 16, breakLine: true }),
      t("الحدث كله نحو ساعتين وأربعين دقيقة، وذروته بين ١٢:٥١ (الإسكندرية) و١٣:٠٨ (أسوان).", { color: K.ice, fontSize: 15 })],
      { x: 9.35, y: 1.7, w: W - M - 9.6, h: 4.35, valign: "top", isTextBox: true, margin: 0, paraSpaceAfter: 12 });
    caption(s, "المصدر: جدول المدن في ملحق الكتاب (محسوب)؛ توقيت مصر الصيفي", 6.62, { x: M, w: 8.2 });
    s.addNotes("هذا الجدول للمرشدين العاملين في القاهرة والإسكندرية وأسوان والبحر الأحمر وسيناء. " +
      "الاحتجاب (Obscuration) هو النسبة المحجوبة من مساحة قرص الشمس. حتى ٩٩٫٨٪ في أسوان لا يعطي ظلامًا ولا هالة. " +
      "اقترح على هذه المجموعات الانتقال إلى أقرب نقطة داخل المسار إن أمكن.");
  }

  // partial cities chart
  {
    const s = add("CONTENT", ["خارج المسار: الجزئي يبقى جزئيًا مهما بلغ", "Outside the path: a partial eclipse stays partial"]);
    const cities = ["أسوان", "سفاجا", "الخارجة", "الغردقة", "بني سويف", "شرم الشيخ", "الجيزة", "القاهرة", "الإسكندرية", "أبو سمبل", "بورسعيد"];
    const vals = [99.8, 99.6, 99.5, 98.6, 98.2, 95.6, 95.0, 94.8, 93.4, 92.0, 88.5];
    s.addChart(pres.charts.BAR, [{ name: "نسبة تغطية قرص الشمس ٪", labels: cities.slice().reverse(), values: vals.slice().reverse() }], {
      x: M, y: 1.45, w: 7.3, h: 5.15, barDir: "bar", chartColors: [K.sunDeep], showValue: true, dataLabelPosition: "outEnd",
      dataLabelFormatCode: "0.0", dataLabelFontSize: 12, dataLabelColor: K.ink, dataLabelFontFace: "+mn-lt",
      catAxisLabelFontSize: 13, catAxisLabelColor: K.ink2, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "minMax",
      valAxisMinVal: 80, valAxisMaxVal: 100, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      showLegend: false, showTitle: true, title: "نسبة تغطية قرص الشمس عند الذروة (٪)", titleFontSize: 13, titleColor: K.ink2,
      titleFontFace: "+mn-lt", barGapWidthPct: 45, objectName: "chart-partial-cities",
    });
    card(s, 8.25, 1.5, W - M - 8.25, 2.25, { fill: K.red, line: K.red, name: "warn-card", flat: true });
    s.addText([t("٩٩٫٨٪ ≠ كلي", { bold: true, fontSize: 26, color: "FFFFFF", breakLine: true }),
      t("الجزء الباقي من الشمس في أسوان أسطع بآلاف المرات من الهالة: لا ليل، ولا هالة، ولا نزع للنظارة أبدًا", { fontSize: 15, color: "FFFFFF" })],
      { x: 8.5, y: 1.62, w: W - M - 8.75, h: 2.0, valign: "middle", isTextBox: true, margin: 0, paraSpaceAfter: 6 });
    addImg(s, "partialVsTotal", { x: 8.25, y: 3.95, w: W - M - 8.25, h: 2.3 }, { frame: true });
    caption(s, "المصدر: جدول المدن في ملحق الكتاب (محسوب)؛ الصورة: الشكل ٢١-٤", 6.62, { x: 8.25, w: W - M - 8.25 });
    s.addNotes("هذه أخطر رسالة في المحاضرة للسلامة وللتسويق معًا. الفرق بين ٩٩٪ و١٠٠٪ هو الفرق بين «يوم غائم قليلًا» و«ليل في الظهيرة». " +
      "لا تبع للسائح تجربة كلية في أسوان أو الغردقة أو القاهرة. وفي هذه المدن تبقى النظارة على العين طوال الحدث. " +
      "اقتراح عملي: مجموعات الغردقة تنتقل جنوبًا إلى القصير أو مرسى علم، ومجموعات أسوان شمالًا إلى كوم أمبو أو إدفو، ومجموعات القاهرة إلى الواحات البحرية أو المنيا.");
  }

  partSlide("٣-ب", "كيف سيبدو الكسوف في الأقصر؟", "What the eclipse will look like in Luxor", "المراحل كاملة، وشكل الشمس في كل مرحلة، والسماء لحظة الكلية، والمواقع الأثرية تحت الظل", "الجزء الثاني من المرحلة الثالثة: الأقصر نموذجًا، ويمكن تطبيق الشرائح نفسها على أي مدينة ببطاقتها في الكتاب.");

  figText(["بطاقة الأقصر", "Luxor eclipse card"], "luxorCard", [
    ["بداية الكسوف: ", "١١:٤٠:٠٧"],
    ["بداية الكلية: ", "١٣:٠١:٥٨"],
    ["نهاية الكلية: ", "١٣:٠٨:١٨ — ٦ د ٢٠ ث"],
    ["نهاية الكسوف: ", "١٤:٢٦:٢٥"],
    ["ارتفاع الشمس: ", "٨١٫٨° — قرب سمت الرأس"],
  ], "لكل مدينة في الكتاب بطاقة مماثلة (٢٥ بطاقة) — اطبع بطاقة مدينتك واحملها يوم الكسوف. " +
    "في الأقصر: نحو ساعة و٢٢ دقيقة من الجزئي قبل الكلية، ثم ست دقائق وعشرون ثانية من الكلية، ثم ساعة و١٨ دقيقة من الجزئي بعدها.",
    { split: 5.4, lead: "الأقصر: كل ما تحتاجه في بطاقة", cap: "الشكل ٢٥-٤ من الكتاب" });

  // complete phases at Luxor: what to see and do (native table)
  {
    const s = add("CONTENT", ["مراحل الكسوف كاملة في الأقصر: ماذا ترى وماذا تفعل؟", "Complete eclipse timeline at Luxor: what to see, what to do"]);
    const rows = [
      ["١١:٤٠:٠٧", "التماس الأول (C1)", "«عضّة» صغيرة على الحافة الغربية العليا للشمس", "النظارة على العين"],
      ["١٢:٠٠ – ١٢:٤٠", "الجزئي المتزايد", "الاحتجاب ١٤٪ ← ٦٨٪؛ أهلّة الضوء تحت الأشجار", "كاميرا الثقب ومصفاة المطبخ"],
      ["١٢:٥٥", "قبل الكلية بسبع دقائق", "نحو ٩١٪: ضوء معدني باهت، الحرارة تنخفض، الظلال حادة", "اشرب، واستعد، وراقب الغرب"],
      ["١٣:٠١", "الدقيقة الأخيرة", "الأحزمة الظلية (Shadow bands)، ظل القمر قادم من الغرب، حبات بيلي ثم الخاتم الماسي", "النظارة حتى يختفي آخر ضوء"],
      ["١٣:٠١:٥٨ – ١٣:٠٨:١٨", "الكلية (Totality) ٦ د ٢٠ ث", "الهالة، الشواظ الحمراء، الزهرة والمشتري وعطارد، غروب على الأفق كله", "انزعوا النظارات وانظروا"],
      ["١٣:٠٨:١٨", "التماس الثالث (C3)", "الخاتم الماسي الثاني على الجهة المقابلة", "النظارات فورًا!"],
      ["١٣:١٥ – ١٤:٠٠", "الجزئي المتناقص", "الاحتجاب ٩١٪ ← ٢٤٪؛ يعود الضوء والحرارة", "صور جماعية، والماء والظل"],
      ["١٤:٢٦:٢٥", "التماس الرابع (C4)", "ينفصل القرصان", "نهاية الكسوف"],
    ];
    const head = ["الوقت", "المرحلة", "ماذا ترى؟", "ماذا تفعل؟"];
    const tot = (i) => i === 5;
    const tbl = [head, ...rows].map((r, i) => r.slice().reverse().map((c, j) => ({ text: c, options: { ...R, align: j === 3 ? "center" : "right",
      fontSize: 13, bold: i === 0 || j === 3 || j === 2, valign: "middle",
      color: i === 0 ? "FFFFFF" : tot(i) ? (j === 3 ? K.sun : "FFFFFF") : j === 3 ? K.blue : j === 0 ? K.red : K.ink,
      fill: { color: i === 0 ? K.navy : tot(i) ? K.navy2 : (i % 2 ? K.card : K.tint) } } })));
    s.addTable(tbl, { x: M, y: 1.5, w: W - 2 * M, colW: [2.75, 5.0, 2.35, 2.03], rowH: 0.55, border: { type: "solid", color: K.line, pt: 0.5 },
      objectName: "phases-table" });
    caption(s, "الأوقات ونسب الاحتجاب محسوبة لموقع الأقصر (توقيت مصر الصيفي)", 6.62);
    s.addNotes("هذا هو «سيناريو» اليوم كاملًا للأقصر، من التماس الأول إلى الرابع. " +
      "ساعتان و٤٦ دقيقة في المجموع، منها ٦ دقائق و٢٠ ثانية فقط من الكلية. " +
      "ركّز على الدقيقة الأخيرة قبل الكلية: يأتي ظل القمر من جهة الغرب-الشمال الغربي بسرعة، وقد تظهر خطوط متموجة باهتة على الأرض الفاتحة (الأحزمة الظلية). " +
      "ثم حبات بيلي والخاتم الماسي: النظارة على العين حتى يختفي آخر ضوء، ثم تُنزع. وعند أول وميض في نهاية الكلية تُلبس فورًا. " +
      "لمدينة أخرى: أضف الفرق بين أوقاتها وأوقات الأقصر من جدول المدن.");
  }

  // Sun's shape at Luxor, phase by phase (native shapes from computed geometry)
  {
    const s = add("DARK", ["شكل الشمس خلال الكسوف كما يُرى من الأقصر", "The Sun's shape through the eclipse as seen from Luxor"]);
    const ph = require("./luxor_phases.json");
    const names = ["التماس الأول", "", "", "", "", "الكلية", "", "", "", "التماس الرابع"];
    const r = 0.5, cw = (W - 2 * M) / 5;
    // suns + moons first, labels last (the Moon disc is painted in the background colour)
    ph.forEach((p, i) => {
      const col = 4 - (i % 5), row = Math.floor(i / 5);
      const cx = M + cw * (col + 0.5), cy = 2.35 + row * 2.45;
      if (p.obs > 0.999) {
        s.addShape(pres.shapes.OVAL, { x: cx - r * 2.0, y: cy - r * 2.0, w: r * 4.0, h: r * 4.0, fill: { color: "FFFFFF", transparency: 88 },
          line: { color: "FFFFFF", width: 0 }, objectName: "corona-outer" });
        s.addShape(pres.shapes.OVAL, { x: cx - r * 1.45, y: cy - r * 1.45, w: r * 2.9, h: r * 2.9, fill: { color: "FFFFFF", transparency: 70 },
          line: { color: "FFFFFF", width: 0 }, objectName: "corona-inner" });
      } else {
        s.addShape(pres.shapes.OVAL, { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color: K.sun }, line: { color: K.sun, width: 0 },
          objectName: "sun-" + p.time });
      }
      const mr = r * p.rm, mx = cx + p.dx * r, my = cy - p.dy * r;
      s.addShape(pres.shapes.OVAL, { x: mx - mr, y: my - mr, w: 2 * mr, h: 2 * mr, fill: { color: K.navy },
        line: { color: p.obs > 0.999 ? "FFFFFF" : "3A4D70", width: p.obs > 0.999 ? 1 : 0.75 }, objectName: "moon-" + p.time });
    });
    ph.forEach((p, i) => {
      const col = 4 - (i % 5), row = Math.floor(i / 5);
      const cx = M + cw * (col + 0.5), cy = 2.35 + row * 2.45;
      const ar = (v) => String(v).replace(/[0-9]/g, (dg) => "٠١٢٣٤٥٦٧٨٩"[dg]);
      const pct = p.obs > 0.999 ? "كلي ١٠٠٪" : p.obs < 0.0005 ? "٠٪" : ar((p.obs * 100).toFixed(0)) + "٪";
      s.addText([t(ar(p.time), { bold: true, color: "FFFFFF", fontSize: 16, breakLine: true, align: "center" }),
        t((names[i] ? names[i] + " — " : "") + pct, { color: p.obs > 0.999 ? K.sun : K.ice, fontSize: 13, align: "center" })],
        { x: cx - cw / 2 + 0.05, y: cy + 0.95, w: cw - 0.1, h: 0.75, isTextBox: true, margin: 0, valign: "top" });
    });
    s.addText("الشمال إلى أعلى والغرب إلى اليمين؛ الشمس شبه عمودية فوق الرأس (٧١°–٨٢°)", { x: M, y: 6.6, w: W - 2 * M, h: 0.3, fontSize: 11,
      italic: true, color: "8C9AB3", isTextBox: true, ...R, margin: 0 });
    s.addNotes("هذه الرسوم محسوبة خصيصًا لهذه المحاضرة من عناصر بِسِل الرسمية (ناسا) لموقع الأقصر: موضع قرص القمر على قرص الشمس في عشر لحظات. " +
      "يدخل القمر من الحافة الغربية العليا للشمس (يمين الصورة إلى أعلى) ويخرج من الحافة الشرقية السفلى. " +
      "نسبة الاحتجاب: ١٢:٠٠ ← ١٤٪، ١٢:٢٠ ← ٣٨٪، ١٢:٤٠ ← ٦٨٪، ١٢:٥٥ ← ٩١٪، ثم الكلية ١٣:٠١:٥٨–١٣:٠٨:١٨، " +
      "ثم ١٣:١٥ ← ٩١٪، ١٣:٣٥ ← ٦٠٪، ١٤:٠٠ ← ٢٤٪، والنهاية ١٤:٢٦:٢٥. " +
      "قرص القمر أكبر ظاهريًا من قرص الشمس بنحو ٨٪ في قطره، ولهذا تطول الكلية. " +
      "ولأن الشمس قريبة من سمت الرأس، فإن «أعلى» و«أسفل» يتغيران بحسب اتجاه وقوف الراصد؛ الرسم بالاتجاهات السماوية: الشمال إلى أعلى.");
  }

  figText(["ماذا يحدث حولك أثناء الكلية؟", "What happens around you during totality"], "skyLuxor", [
    ["الكواكب والنجوم: ", "تظهر الزهرة ملاصقة للشمس تقريبًا، والمشتري وعطارد، والشعرى اليمانية"],
    ["الضوء: ", "يهبط إلى ما يشبه الشفق (Twilight)، مع «غروب» على امتداد الأفق كله (360° sunset)"],
    ["الحرارة: ", "تنخفض بضع درجات، ويتغير اتجاه الريح أحيانًا"],
    ["الكائنات: ", "تسكت الطيور، وتعود بعض الحيوانات إلى مأواها"],
    ["الظلال: ", "قبيل الكلية تصبح حادّة غريبة، وقد تظهر «الأحزمة الظلية» (Shadow bands) على الأرض الفاتحة"],
  ], "خريطة سماء الأقصر لحظة ذروة الكلية (١٣:٠٥) من الكتاب. الشمس في مركز الدائرة تقريبًا (ارتفاعها نحو ٨٢°)، " +
    "فليُنبَّه الزوار إلى أن ينظروا إلى أعلى، لا إلى الأفق. ويمكن للمرشد أن يجعلها لعبة: " +
    "«من يرى الزهرة أولًا؟» — وهي أسطع جرم بعد الشمس والقمر، وستكون قريبة جدًا من الشمس المكسوفة. " +
    "وأوصف الغروب الدائري: الأفق كله برتقالي لأننا نرى خارج حدود الظل في كل اتجاه.",
    { dark: true, split: 6.4, cap: "الشكل ٢٨-١ من الكتاب: سماء الأقصر ١٣:٠٥ (محسوبة)" });

  figText(["آثار مصر تحت ظل القمر", "Egypt's monuments under the Moon's shadow"], "monuments", [
    ["أكثر من ٦ دقائق: ", "أبيدوس والكرنك (٦ د ٢١ ث)، الدير البحري ووادي الملوك وتمثالا ممنون ومعبد الأقصر ومدينة هابو (٦ د ٢٠ ث)، دندرة (٦ د ١٤ ث)"],
    ["أكثر من ٥ دقائق: ", "إسنا، تل العمارنة، معبد آمون في سيوة، إدفو، الأشمونين"],
    ["أقل: ", "بني حسن (٤ د ٣٧ ث)، كوم أمبو (٣ د ٢١ ث)"],
    ["جزئي: ", "فيلة ٩٩٫٧٪، سقارة والأهرام ٩٥٪، سانت كاترين ٩٤٫٢٪، أبو سمبل ٩٢٪"],
  ], "هذه الشريحة هي الأهم للأثريين. يمر مسار الكلية بأغنى بقاع الأرض بالآثار. " +
    "تذكير: الشمس على ارتفاع ٨١–٨٢° في الأقصر، أي فوق الرأس تقريبًا؛ فلا تتوقع صورة «الشمس المكسوفة فوق المسلة» إلا بعدسة واسعة جدًا من زاوية منخفضة. " +
    "وأهرام الجيزة خارج المسار (٩٥٪) — صحّح هذا لمن يظن أن الأهرام ستشهد الكلية.",
    { split: 7.0, fs: 14, cap: "الشكل ٢٤-٣ من الكتاب: مدة الكلية عند المواقع الأثرية (محسوبة)" });

  figText(["خريطة المواقع الأثرية في وادي النيل", "Archaeological sites of the Nile Valley inside the path"], "monumentsMap", [
    ["الأقصر الكبرى: ", "الضفتان داخل المسار بفارق ثوانٍ؛ اختر الموقع بحسب الزحام والظل والخدمات"],
    ["أبيدوس ودندرة: ", "أقرب إلى خط المركز وأقل ازدحامًا من الكرنك ووادي الملوك"],
    ["مصر الوسطى: ", "تل العمارنة والأشمونين وبني حسن: تجربة تجمع آثار إخناتون والكسوف"],
    ["تنظيم الزيارة: ", "نسّق مسبقًا مع إدارة الموقع بشأن مواعيد الدخول والبقاء في الساحات المفتوحة"],
  ], "انصح بالمواقع المفتوحة ذات الأفق الواسع لرؤية الظل قادمًا من الغرب. " +
    "داخل المعابد المسقوفة لن يُرى شيء من السماء. واعلم أن تنظيم المواقع الأثرية يوم الكسوف (ساعات الدخول والأعداد) قرار الجهات المختصة؛ تابع تعليماتها الرسمية.",
    { split: 6.6, cap: "الشكل ٢٥-٣ من الكتاب (محسوب)" });

  figText(["الواحات البحرية: الخيار الأقرب للقاهرة", "Bahariya Oasis: the closest option to Cairo"], "bahariya", [
    ["الباويطي: ", "الكلية ١٢:٥٠:١٠–١٢:٥٦:٢٨، أي ٦ د ١٧ ث، والشمس على ارتفاع ٧٩°"],
    ["الحيز ومنديشة: ", "٦ د ١٦ ث و٦ د ١٥ ث"],
    ["خط المركز (Central line): ", "٦ د ٢٢ ث على بعد نحو ٢٢ كم جنوب الباويطي"],
    ["المزايا: ", "أفق صحراوي مفتوح، وسماء صافية عادة في أغسطس، ونحو ٣٧٠ كم فقط من القاهرة"],
  ], "الواحات البحرية تتبع إداريًا محافظة الجيزة، وتقع داخل مسار الكلية بالكامل. " +
    "وهي خيار ممتاز للمجموعات القادمة من القاهرة التي لا تريد السفر إلى الصعيد، مع إمكان الجمع بين الكسوف والصحراء البيضاء والسوداء. " +
    "وكذلك سيوة (٥ د ٢٨ ث) لمن يريد تجربة الواحة ومعبد آمون.",
    { split: 6.9, cap: "الشكل ١٦-٩ من الكتاب (محسوب)" });

  partSlide("٣-ج", "كيف ترصده بأمان؟", "How to observe safely", "النظارة المعتمدة، والوسائل الآمنة، ومتى تُنزع النظارة، والخرافات، والحر", "الجزء الثالث من المرحلة الثالثة: السلامة.");

  figText(["القاعدة الذهبية لسلامة العين", "The golden rule of eye safety"], "golden", [
    ["دائمًا: ", "نظارة كسوف (Eclipse glasses) مطابقة للمعيار ISO 12312-2، سليمة غير مخدوشة"],
    ["الاستثناء الوحيد: ", "داخل مسار الكلية، وبين التماسين الثاني والثالث فقط"],
    ["ممنوع: ", "النظارات الشمسية، والأفلام، والزجاج المدخّن، والنظر عبر كاميرا أو منظار بلا مرشح"],
    ["زجاج اللحام: ", "رقم ١٤ فقط مقبول؛ الأقل منه غير آمن"],
    ["الخطر: ", "اعتلال الشبكية الشمسي (Solar retinopathy) بلا ألم، وقد يكون دائمًا"],
  ], "ضوء الشمس — حتى لو حُجب ٩٩٪ منه — قادر على إتلاف الشبكية دون أن يشعر الإنسان بألم، لأن الشبكية بلا مستقبلات للألم. " +
    "ابدأ هذه الشريحة بعبارة ابن يونس في الرواية: «في كل بيت رصد أعمى، كان في شبابه أحمق ظنّ أن نظرة واحدة لا تضر». " +
    "اشترِ النظارات مبكرًا من مصدر موثوق، وافحص كل نظارة قبل التوزيع، واحمل احتياطيًا.",
    { split: 6.9, fs: 15, cap: "الشكل ١-٥ من الكتاب" });

  figText(["الوسائل الآمنة وغير الآمنة", "Safe vs. unsafe ways to watch the Sun"], "safeUnsafe", [
    ["آمن: ", "نظارة كسوف معتمدة (ISO 12312-2)، مرشح شمسي (Solar filter) أمام العدسة، كاميرا الثقب، زجاج لحام رقم ١٤"],
    ["غير آمن: ", "النظارات الشمسية (Sunglasses)، أفلام الأشعة، الأقراص المدمجة، الزجاج المدخّن، الانعكاس في الماء"],
    ["الأخطر: ", "منظار أو تلسكوب أو عدسة كاميرا بلا مرشح أمامها: يركّز الضوء فيحرق العين في ثانية"],
    ["الفحص: ", "النظارة السليمة لا يُرى منها إلا الشمس؛ أي خدش أو ثقب ← تُستبعد"],
  ], "نبّه إلى الفرق بين الرصد القديم والحديث: طست الماء عند ابن يونس كان أفضل المتاح قبل ألف سنة، لكن الانعكاس على الماء يبقى ساطعًا ولا يُعتمد اليوم. " +
    "ونبّه المصورين: المرشح يوضع أمام العدسة لا خلفها، ويُنزع فقط في الكلية داخل المسار. " +
    "ومن الأخطاء الشائعة أن يلبس الشخص نظارة الكسوف ثم ينظر عبر منظار بلا مرشح: النظارة لا تحميه حينئذ.",
    { split: 6.9, fs: 15, cap: "الشكل ٢-٥ من الكتاب" });

  figText(["متى يلبس الزائر النظارة ومتى ينزعها؟ (الأقصر)", "Glasses on, glasses off: the Luxor timeline"], "glassesLuxor", [
    ["١١:٤٠: ", "النظارة على العين لمتابعة بداية الجزئي"],
    ["١٣:٠١:٥٨: ", "يختفي آخر ضوء ← انزعوا النظارات"],
    ["٦ دقائق و٢٠ ثانية: ", "انظروا بالعين المجردة إلى الهالة والكواكب"],
    ["١٣:٠٨:١٨: ", "أول وميض للخاتم الماسي ← النظارات فورًا"],
    ["١٤:٢٦: ", "نهاية الكسوف"],
  ], "استعمل منبّهًا صوتيًا أو تطبيق Solar Eclipse Timer، وعيّن في كل مجموعة مساعدًا يراقب الوقت. " +
    "قاعدة بسيطة للزائر: «إذا رأيت أي ضوء أبيض ساطع من الشمس، فالنظارة على عينك». " +
    "وفي المدن خارج المسار لا توجد لحظة نزع إطلاقًا. وفي الكتاب خطوط زمنية مماثلة لسوهاج وأسيوط وأسوان والقاهرة والغردقة.",
    { split: 7.8, fs: 15, cap: "الشكل ٩-٥ من الكتاب (محسوب)" });

  figText(["خرافات شائعة سيسألك عنها الزوار", "Common myths visitors will ask about"], "myths", [
    ["«الكسوف يضر الحوامل والطعام»: ", "لا أصل علميًا لذلك"],
    ["«لا تخرج وقت الكسوف»: ", "الخطر الوحيد هو النظر إلى الشمس بلا مرشح"],
    ["«الكسوف نذير»: ", "«لا ينكسفان لموت أحد ولا لحياته» — حديث نبوي"],
    ["«الأشعة أخطر أثناء الكسوف»: ", "الشمس هي الشمس؛ الخطر أن الناس يحدّقون فيها"],
  ], "اجمع بين العلم والتراث في الرد: الحديث النبوي ينفي النذير، والحساب الفلكي يتنبأ بالكسوف بالثانية قبله بقرون. " +
    "وفي رواية «أرصاد وأطوار» مشهد ساخر: المنجمون الذين أنذروا الخليفة بموت ملك عند كسوف ٩٩٣م «قبضوا أجرهم مرتين: مرة على الإنذار، ومرة على أن الإنذار لم يقع».",
    { split: 6.6, fs: 15, cap: "الشكل ١٥-٥ من الكتاب" });

  figText(["ظهيرة أغسطس في الصعيد: الحر خطر حقيقي", "August noon in Upper Egypt: heat is a real hazard"], "heat", [
    ["الشمس فوق الرأس: ", "ارتفاعها ٨١–٨٢°، فلا ظل تقريبًا عند الجدران والأشجار"],
    ["الحرارة (Heat stress): ", "ظهيرة أغسطس في الأقصر وأسوان تتجاوز ٤٠° عادةً"],
    ["الماء: ", "لتر لكل ساعة لكل شخص، مع أملاح"],
    ["الوقاية: ", "قبعة واسعة، ومظلات، وملابس فاتحة، ومكان استراحة مظلل قريب"],
    ["الإسعاف: ", "حدّد أقرب نقطة إسعاف، وراقب كبار السن والأطفال"],
  ], "انخفاض الحرارة بضع درجات أثناء الكلية لا يعوّض ساعات الانتظار في الشمس. " +
    "خطط لوصول المجموعة إلى الموقع قبل التماس الأول بوقت كافٍ، ووفر ظلًا متنقلًا. " +
    "الخطر الأكبر يوم الكسوف عمليًا هو الإجهاد الحراري والزحام المروري، لا الكسوف نفسه.",
    { split: 6.6, fs: 15, cap: "الشكل ١٨-٥ من الكتاب" });

  partSlide("٣-د", "كيف تصوّره؟", "How to photograph it", "بالهاتف الذكي، ثم بالكاميرا: المرشح، والوضع اليدوي، والتعريض المتدرج، وخطة الدقائق الست", "الجزء الرابع من المرحلة الثالثة: التصوير.");

  figText(["تصوير الكسوف بالهاتف", "Photographing the eclipse with a smartphone"], "phone", [
    ["أثناء الجزئي: ", "ضع مرشحًا شمسيًا أو نظارة كسوف أمام عدسة الهاتف، واقفل التركيز والتعريض (AE/AF lock)"],
    ["أثناء الكلية: ", "انزع المرشح؛ صوّر الأفق الدائري والناس بعدسة واسعة (Wide angle) — الشمس فوق الرأس"],
    ["الوضع الليلي: ", "أطفئ الفلاش، وثبّت الهاتف على حامل (Tripod)"],
    ["الأهم: ", "لا تقضِ الدقائق الست خلف الشاشة؛ صوّر دقيقة ثم انظر بعينيك"],
  ], "سيطلب كثير من الزوار مساعدتهم في التصوير. الشمس في الأقصر على ارتفاع ٨٢° تقريبًا، فصورة «الشمس المكسوفة فوق المعبد» تحتاج عدسة واسعة جدًا وزاوية منخفضة. " +
    "وفي الكتاب فصل كامل للتصوير: مثلث التعريض (Exposure triangle)، وجداول التعريض المتدرج (Bracketing) لكل ظاهرة، وخطة تصوير ثانية بثانية للأقصر وسوهاج. " +
    "نصيحة المصورين المحترفين المنقولة في الكتاب عن آلان هول: جرّب الحامل ورأسه في الساعة نفسها قبل الحدث بأيام، لأن التصوير قرب سمت الرأس صعب.",
    { split: 6.8, fs: 15, cap: "الشكل ٢١-٧ من الكتاب" });

  figText(["تصوير الكسوف بالكاميرا: خطة الدقائق الست في الأقصر", "Camera photography: the six-minute totality plan for Luxor"], "planLuxor", [
    ["الوضع اليدوي (Manual mode): ", "فتحة ثابتة مثل f/8 وحساسية منخفضة ISO 100–400، وتغيير زمن الغالق وحده"],
    ["المرشح الشمسي (Solar filter): ", "أمام العدسة طوال الجزئي؛ يُنزع قبل التماس الثاني بثوانٍ ويُعاد بعد الثالث"],
    ["التعريض المتدرج (Bracketing): ", "سلسلة لقطات من ١/٤٠٠٠ ث حتى ١ ث تقريبًا لتغطي الهالة الداخلية والخارجية"],
    ["البعد البؤري: ", "٤٠٠–٨٠٠ مم لقرص الشمس والهالة؛ عدسة واسعة للمشهد لأن الشمس فوق الرأس"],
    ["الأتمتة: ", "برامج مثل Eclipse Orchestrator تلتقط التسلسل آليًا، فتتفرغ عيناك"],
  ], "خطة التصوير ثانية بثانية للأقصر من الكتاب. القاعدة الذهبية للمصور: ارفع عينيك عن الكاميرا في منتصف الكلية (١٣:٠٥:٠٨) واستمتع. " +
    "مثلث التعريض (Exposure triangle): فتحة العدسة (Aperture) وزمن الغالق (Shutter speed) والحساسية (ISO). " +
    "في الكسوف نستعمل الوضع اليدوي لأن الوضع الآلي ينخدع بالسماء المظلمة حول الهالة. " +
    "وفي الكتاب جداول تعريض محسوبة لكل ظاهرة عند f/5.6 وf/8 وf/11، وخطة مماثلة لسوهاج. " +
    "تحدي مصر الخاص: الشمس على ارتفاع ٨١–٨٢°، فجرّب الحامل ورأسه قبل الحدث بأيام في الساعة نفسها.",
    { dark: true, split: 6.3, fs: 15, cap: "الشكل ١٤-٧ من الكتاب: خطة التصوير خلال الكلية في الأقصر (محسوبة)" });

  figText(["إعدادات الكاميرا لكل مرحلة", "Camera settings for each phase of the eclipse"], "phaseSettings", [
    ["الجزئي (Partial): ", "مع المرشح، تعريض قصير، صورة كل ٥–١٠ دقائق لصنع تسلسل المراحل"],
    ["حبات بيلي والخاتم الماسي: ", "بلا مرشح، تعريض قصير جدًا وتصوير متتابع سريع (Burst)"],
    ["الشواظ والكروموسفير: ", "تعريضات قصيرة في أول الكلية وآخرها"],
    ["الهالة (Corona): ", "سلسلة من الأقصر إلى الأطول، ثم دمج HDR بعد الحدث"],
    ["المدرّج التكراري (Histogram): ", "راجعه في الجزئي فقط؛ لا وقت لذلك في الكلية"],
  ], "الشكل يلخص إعدادات كل مرحلة كما في الفصل السابع من الكتاب. " +
    "ذكّر المصورين بأن أكثر الأخطاء شيوعًا: نسيان نزع المرشح عند الكلية، أو نسيان إعادته بعدها، والتركيز التلقائي (Autofocus) الذي يضيع في الظلام — ثبّت التركيز يدويًا على اللانهاية مسبقًا.",
    { split: 7.4, fs: 15, cap: "الشكل ٥-٧ من الكتاب: الإعدادات المقترحة لكل مرحلة (محسوبة)" });

  section = "الختام";
  pres.addSection({ title: section });

  // guide timeline (native process)
  {
    const s = add("CONTENT", ["خطة المرشد: من الآن حتى نهاية الكسوف", "The guide's plan: from now until fourth contact"]);
    const steps = [
      ["قبل سنة", "حجز الموقع والنقل؛ دراسة بطاقة المدينة؛ قراءة الكتاب والرواية"],
      ["قبل ٣ أشهر", "شراء نظارات معتمدة + ٢٠٪ احتياطي؛ تجربة التطبيقات"],
      ["قبل أسبوع", "مراجعة الطقس والطرق؛ إحاطة المجموعة بالسلامة"],
      ["يوم الكسوف صباحًا", "الوصول مبكرًا؛ توزيع النظارات والماء؛ تحديد الظل والإسعاف"],
      ["١١:٤٠ ← ١٣:٠٠", "متابعة الجزئي بالنظارة؛ أهلّة الأشجار؛ شرح ما سيأتي"],
      ["١٣:٠٢ ← ١٣:٠٨", "«انزعوا النظارات!» ثم «النظارات!» مع أول وميض"],
    ];
    const n = steps.length, gap = 0.18, w = (W - 2 * M - gap * (n - 1)) / n;
    s.addShape(pres.shapes.LINE, { x: M + 0.3, y: 2.05, w: W - 2 * M - 0.6, h: 0, line: { color: K.sun, width: 2, dashType: "dash" }, objectName: "timeline-line" });
    steps.forEach(([when, what], i) => {
      const x = W - M - w - i * (w + gap);
      disc(s, x + w / 2 - 0.4, 1.65, 0.8, String("١٢٣٤٥٦"[i]), { size: 20, name: "step-disc-" + (i + 1) });
      card(s, x, 2.75, w, 3.1, { name: "step-card-" + (i + 1), fill: i >= 4 ? K.navy : K.card, line: i >= 4 ? K.navy : K.line });
      s.addText(when, { x: x + 0.15, y: 2.9, w: w - 0.3, h: 0.7, fontSize: 15, bold: true, color: i >= 4 ? K.sun : K.blue,
        isTextBox: true, ...R, margin: 0, valign: "top" });
      s.addText(what, { x: x + 0.15, y: 3.6, w: w - 0.3, h: 2.15, fontSize: 14, color: i >= 4 ? "FFFFFF" : K.ink,
        isTextBox: true, ...R, margin: 0, valign: "top" });
    });
    s.addText([t("بعد الكلية: ", { bold: true, color: K.navy }), t("ساعة و١٨ دقيقة من الجزئي — وقت مناسب للصور الجماعية وجمع النظارات وشرح ما رأيناه وربطه بابن يونس وإخناتون.", { color: K.ink })],
      { x: M, y: 6.0, w: W - 2 * M, h: 0.6, fontSize: 15, isTextBox: true, margin: 0, valign: "middle" });
    s.addNotes("الأوقات في الخطوتين الأخيرتين للأقصر؛ عدّلها حسب بطاقة مدينتك. " +
      "تدرب على عبارتي «انزعوا النظارات» و«النظارات الآن» بالعربية والإنجليزية: Glasses OFF — totality! / Glasses ON — now! " +
      "وانتبه إلى أن أذان الظهر وصلاة الكسوف سيقعان في الوقت نفسه تقريبًا. " +
      "وفي الكتاب فصل كامل عن التطبيقات (Xavier Jubier، Stellarium، PhotoPills، Solar Eclipse Timer) وجدول استعداد زمني.");
  }

  // bilingual glossary (native tables)
  {
    const terms = [
      ["كسوف الشمس", "Solar eclipse", "القمر يحجب الشمس، في المحاق"],
      ["خسوف القمر", "Lunar eclipse", "ظل الأرض يغطي القمر، في البدر"],
      ["كسوف كلي", "Total eclipse", "حجب قرص الشمس كله"],
      ["كسوف حلقي", "Annular eclipse", "تبقى حلقة مضيئة حول القمر"],
      ["كسوف جزئي", "Partial eclipse", "حجب جزء من القرص فقط"],
      ["كسوف هجين", "Hybrid eclipse", "حلقي في جزء من المسار وكلي في آخر"],
      ["الظل الكامل", "Umbra", "منطقة الكلية"],
      ["شبه الظل", "Penumbra", "منطقة الكسوف الجزئي"],
      ["مسار الكلية", "Path of totality", "الشريط الذي تُرى فيه الكلية"],
      ["خط المركز", "Central line", "منتصف المسار، وفيه أطول مدة"],
      ["الكلية", "Totality", "الفترة بين التماسين الثاني والثالث"],
      ["التماسات", "Contacts (C1–C4)", "لحظات بدء المراحل وانتهائها"],
      ["قدر الكسوف", "Eclipse magnitude", "نسبة قطر الشمس الذي يغطيه القمر"],
      ["نسبة الاحتجاب", "Obscuration", "نسبة مساحة القرص المحجوبة"],
      ["الهالة", "Corona", "الغلاف الخارجي للشمس"],
      ["الكروموسفير", "Chromosphere", "الطبقة الوردية فوق سطح الشمس"],
      ["الشواظ", "Prominences", "ألسنة غاز حمراء عند حافة الشمس"],
      ["الخاتم الماسي", "Diamond ring", "آخر وميض قبل الكلية وأوله بعدها"],
      ["حبات بيلي", "Baily's beads", "ضوء الشمس عبر وديان حافة القمر"],
      ["الأحزمة الظلية", "Shadow bands", "خطوط متموجة على الأرض قبيل الكلية"],
      ["العقدتان", "Lunar nodes", "نقطتا تقاطع مداري القمر والأرض"],
      ["دورة ساروس", "Saros cycle", "١٨ سنة و١١ يومًا و٨ ساعات"],
      ["الحضيض / الأوج", "Perigee / Apogee", "أقرب / أبعد نقطة للقمر من الأرض"],
      ["المحاق / البدر", "New Moon / Full Moon", "وقت الكسوف / وقت الخسوف"],
      ["فرق التوقيت", "Delta T (ΔT)", "أثر تباطؤ دوران الأرض على الحساب"],
      ["عناصر بِسِل", "Besselian elements", "معاملات حساب ظروف الكسوف"],
      ["نظارة الكسوف", "Eclipse glasses (ISO 12312-2)", "مرشح آمن للنظر إلى الشمس"],
      ["اعتلال الشبكية الشمسي", "Solar retinopathy", "تلف العين بالنظر إلى الشمس"],
      ["كاميرا الثقب", "Pinhole projector", "رصد غير مباشر آمن"],
      ["الزيج", "Zij (astronomical tables)", "جداول فلكية في الحضارة الإسلامية"],
    ];
    [terms.slice(0, 15), terms.slice(15)].forEach((part, k) => {
      const s = add("CONTENT", ["مسرد المصطلحات (" + (k === 0 ? "١" : "٢") + " من ٢)", "Bilingual glossary of eclipse terms (" + (k + 1) + "/2)"]);
      const head = ["المصطلح", "English", "المعنى"];
      const rows = [head, ...part].map((r, i) => r.slice().reverse().map((c, j) => ({ text: c, options: {
        fontSize: 13, valign: "middle",
        ...(j === 1 ? { align: "center", lang: "en-US", italic: i > 0, color: i === 0 ? "FFFFFF" : K.blue } : { ...R, color: i === 0 ? "FFFFFF" : (j === 2 ? K.navy : K.ink) }),
        bold: i === 0 || j === 2, fill: { color: i === 0 ? K.navy : (i % 2 ? K.card : K.tint) } } })));
      s.addTable(rows, { x: M, y: 1.5, w: W - 2 * M, colW: [5.5, 3.6, 3.03], rowH: 0.33, border: { type: "solid", color: K.line, pt: 0.5 },
        objectName: "glossary-" + (k + 1) });
      s.addNotes("مسرد ثنائي اللغة للمصطلحات التي سيستعملها المرشد مع الزوار الأجانب. " +
        "اطبعه بطاقة صغيرة. والمسرد الكامل في ملحق الكتاب.");
    });
  }

  // conclusion (dark)
  {
    const s = add("DARK", ["خلاصة: سبع رسائل يحملها المرشد", "Summary: seven key messages"]);
    const msgs = [
      "الكسوف في آخر الشهر القمري، والخسوف في منتصفه",
      "الكسوف الكلي كل ١٨ شهرًا في مكان ما، وعلى المدينة نفسها مرة كل قرون",
      "الخسوف الكلي يُرى من المكان نفسه كل سنتين إلى ثلاث — أقربها ٣١ ديسمبر ٢٠٢٨",
      "«لا ينكسفان لموت أحد ولا لحياته» — والقاهرة الفاطمية رصدت الكسوف بالطست والأصابع",
      "الأقصر: كلية ٦ د ٢٠ ث من ١٣:٠١:٥٨، لأول مرة منذ ١٣٧٧م",
      "أسوان والقاهرة والغردقة خارج المسار: النظارة لا تُنزع أبدًا",
      "الحر والزحام هما الخطر العملي الأكبر يوم الكسوف",
    ];
    msgs.forEach((m, i) => {
      const col = i < 4 ? 0 : 1, row = i < 4 ? i : i - 4;
      const w = 5.9, x = col === 0 ? W - M - w : M, y = 1.55 + row * 1.25;
      disc(s, x + w - 0.75, y + 0.1, 0.7, "١٢٣٤٥٦٧"[i], { size: 18, name: "msg-disc-" + i, fill: K.navy2 });
      s.addText(m, { x, y, w: w - 0.95, h: 0.95, fontSize: 17, color: "FFFFFF", isTextBox: true, ...R, margin: 0, valign: "middle" });
    });
    s.addNotes("لخّص الرسائل السبع، ثم افتح باب الأسئلة. " +
      "ذكّر الحضور بأن جداول جميع المدن المصرية الخمسين، وبطاقات ٢٥ مدينة، وخرائط المحافظات، موجودة في الكتاب.");
  }

  // sources
  {
    const s = add("CONTENT", ["المصادر", "Sources & references"]);
    s.addText(bullets([
      ["الكتاب: ", "«الكسوف الكلي للشمس — ٢ أغسطس ٢٠٢٧: دليل علمي مصوّر لمصر والعالم العربي» (أكتوبر ٢٠٢٦) — مصدر الأشكال والجداول والأوقات"],
      ["الرواية: ", "«أرصاد وأطوار: سر الحاكم — ثغرة الكسوف»، ولا سيما «ملحق المؤلف: ما هو حقيقي في هذا الكتاب» وجدولها الزمني"],
      ["ناسا: ", "عناصر بِسِل لكسوف ٢٠٢٧ (Fred Espenak, NASA GSFC)، وصور ناسا المفتوحة للاستخدام التعليمي"],
      ["الحساب: ", "تقويم JPL DE421 (Skyfield)، ونظرية ERFA مع معادلات إسبيناك ومييس لفرق التوقيت ΔT للعصور القديمة"],
      ["ابن يونس: ", "«الزيج الحاكمي الكبير»، قطعة مخطوط ليدن، نشرة كوسان دي برسفال، باريس ١٨٠٤"],
      ["الحديث: ", "صحيح البخاري، كتاب الكسوف"],
      ["مراجع أخرى في الكتاب: ", "Sepharial, Eclipses (1915)؛ Allan Hall, How to Photograph an Eclipse (2017)؛ Jean Meeus"],
    ], { run: { fontSize: 15, color: K.ink }, gap: 9 }), { x: M, y: 1.5, w: W - 2 * M, h: 4.9, valign: "top", isTextBox: true, margin: 0, ...R });
    caption(s, "حسابات خاصة بهذه المحاضرة: كسوف ٦٣٢م في المدينة، وخسوفات ٢٠٢٨–٢٠٣٠ من القاهرة، ووقت الزوال يوم ٢ أغسطس ٢٠٢٧ — بأدوات الكتاب نفسها", 6.5);
    s.addNotes("كل الأرقام قابلة للتحقق: شيفرة الحساب في مستودع الكتاب.");
  }

  await pres.writeFile({ fileName: OUT });
  if (SKILL_THEME) {
    const { applyTheme } = require(SKILL_THEME);
    await applyTheme(OUT, THEME);
  }
  console.log("wrote", OUT);
}

build().catch((e) => { console.error(e); process.exit(1); });
