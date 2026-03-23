const pptxgen = require("pptxgenjs");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Janne Kammonen";
pptx.title = "Kuituverkon perusteet";

// Theme colors (no # prefix)
const BG_TITLE = "0A0E1A";
const BG_CONTENT = "111827";
const TEXT_PRIMARY = "E2E8F0";
const TEXT_DIM = "94A3B8";
const CYAN = "06B6D4";
const GOLD = "FACC15";
const PURPLE = "8B5CF6";
const GREEN = "22C55E";
const RED = "EF4444";
const CARD_BG = "1E293B";
const CARD_BG2 = "1F2937";
const TABLE_HEADER = "0F172A";
const TABLE_ROW1 = "1E293B";
const TABLE_ROW2 = "162032";

// Helper: add title bar to content slides
function addTitleBar(slide, title) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: "100%", h: 0.9,
    fill: { color: CARD_BG2 },
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0.85, w: "100%", h: 0.06,
    fill: { color: CYAN },
  });
  slide.addText(title, {
    x: 0.6, y: 0.15, w: 11, h: 0.6,
    fontSize: 32, fontFace: "Arial", bold: true, color: TEXT_PRIMARY,
  });
}

// Helper: section divider slide
function addSectionSlide(title, subtitle) {
  const slide = pptx.addSlide();
  slide.background = { fill: BG_TITLE };
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 4.5, y: 2.2, w: 4, h: 0.06, fill: { color: CYAN },
  });
  slide.addText(title, {
    x: 0.5, y: 2.5, w: 12, h: 1.2,
    fontSize: 44, fontFace: "Arial", bold: true, color: TEXT_PRIMARY, align: "center",
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 4.5, y: 3.8, w: 4, h: 0.06, fill: { color: CYAN },
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 1, y: 4.1, w: 11, h: 0.6,
      fontSize: 18, fontFace: "Arial", color: TEXT_DIM, align: "center",
    });
  }
  return slide;
}

// Helper: add a card (colored rectangle with text)
function addCard(slide, x, y, w, h, text, opts = {}) {
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: opts.fill || CARD_BG },
    rectRadius: 0.1,
  });
  if (opts.accentTop) {
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: x + 0.05, y, w: w - 0.1, h: 0.04,
      fill: { color: opts.accentTop },
    });
  }
  slide.addText(text, {
    x: x + 0.15, y: y + (opts.accentTop ? 0.12 : 0.08), w: w - 0.3, h: h - 0.2,
    fontSize: opts.fontSize || 13, fontFace: "Arial", color: opts.color || TEXT_PRIMARY,
    valign: opts.valign || "top", bold: opts.bold || false,
    lineSpacingMultiple: 1.2,
  });
}

// Helper: footer note
function addFooter(slide, text) {
  slide.addText(text, {
    x: 0.5, y: 7.0, w: 12, h: 0.4,
    fontSize: 10, fontFace: "Arial", color: TEXT_DIM, italic: true,
  });
}

// =============================================
// SLIDE 1: Title Slide
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_TITLE };

  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: "100%", h: 0.08, fill: { color: CYAN },
  });

  slide.addText("Kuituverkon perusteet", {
    x: 0.5, y: 1.5, w: 12, h: 1.4,
    fontSize: 54, fontFace: "Arial", bold: true, color: TEXT_PRIMARY, align: "center",
  });

  slide.addText("Valokuidusta PON-verkkoon \u2014 koulutuspaketti", {
    x: 0.5, y: 3.0, w: 12, h: 0.8,
    fontSize: 22, fontFace: "Arial", color: CYAN, align: "center",
  });

  // Decorative fiber lines
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 2, y: 4.2, w: 9, h: 0.03, fill: { color: GOLD },
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 3, y: 4.4, w: 7, h: 0.03, fill: { color: CYAN },
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 4, y: 4.6, w: 5, h: 0.03, fill: { color: PURPLE },
  });

  slide.addText("Janne Kammonen", {
    x: 0.5, y: 5.4, w: 12, h: 0.5,
    fontSize: 16, fontFace: "Arial", color: TEXT_DIM, align: "center",
  });
}

// =============================================
// SLIDE 2: Section — Valon fysiikka
// =============================================
addSectionSlide("Osa I: Valon fysiikka", "Valo, aallonpituus ja kokonaisheijastus");

// =============================================
// SLIDE 3: Mita valo on
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Mit\u00e4 valo on?");

  // EM spectrum card
  addCard(slide, 0.4, 1.2, 5.8, 2.8, [
    { text: "S\u00e4hk\u00f6magneettinen s\u00e4teily", options: { fontSize: 16, bold: true, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "Valo on s\u00e4hk\u00f6magneettista s\u00e4teily\u00e4, joka etenee aaltoliikkeen\u00e4. N\u00e4kyv\u00e4 valo on vain pieni osa EM-spektri\u00e4.", options: { fontSize: 13, breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "Tietoliikennekuidussa k\u00e4ytet\u00e4\u00e4n infrapuna-aluetta:", options: { fontSize: 13, breakLine: true } },
    { text: "  850 nm \u2014 monimoodi, lyhyet v\u00e4lit", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: "  1310 nm \u2014 yksimoodi, upstream", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: "  1490 nm \u2014 yksimoodi, downstream", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: "  1550 nm \u2014 pitkät v\u00e4lit, CATV", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: "  1625 nm \u2014 monitorointi (OTDR)", options: { fontSize: 12, color: GOLD } },
  ], { accentTop: CYAN });

  // Key concept card
  addCard(slide, 6.6, 1.2, 6, 2.8, [
    { text: "Miksi infrapuna?", options: { fontSize: 16, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "Lasin vaimennus on pienimmillaan 1260\u20131625 nm alueella.", options: { fontSize: 13, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Vaimennusminimi:", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "  ~0.20 dB/km @ 1550 nm", options: { fontSize: 14, color: GREEN, breakLine: true } },
    { text: "  ~0.35 dB/km @ 1310 nm", options: { fontSize: 14, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "N\u00e4kyv\u00e4ll\u00e4 valolla vaimennus olisi >10 dB/km \u2014 ei k\u00e4ytt\u00f6kelpoinen.", options: { fontSize: 12, color: TEXT_DIM } },
  ], { accentTop: GOLD });

  // EM spectrum bar at bottom
  const spectrumColors = ["8B5CF6","3B82F6","06B6D4","22C55E","FACC15","F97316","EF4444"];
  const spectrumLabels = ["UV","Violet","Blue","Green","Yellow","Orange","Red"];
  for (let i = 0; i < 7; i++) {
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 1.0 + i * 1.2, y: 4.5, w: 1.2, h: 0.5,
      fill: { color: spectrumColors[i] },
    });
    slide.addText(spectrumLabels[i], {
      x: 1.0 + i * 1.2, y: 4.5, w: 1.2, h: 0.5,
      fontSize: 9, fontFace: "Arial", color: "FFFFFF", align: "center", valign: "middle", bold: true,
    });
  }
  // IR box
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 9.4, y: 4.5, w: 2.5, h: 0.5,
    fill: { color: "7F1D1D" },
  });
  slide.addText("INFRAPUNA (telecom)", {
    x: 9.4, y: 4.5, w: 2.5, h: 0.5,
    fontSize: 10, fontFace: "Arial", color: RED, align: "center", valign: "middle", bold: true,
  });
  slide.addText("380 nm                                                         780 nm                  850\u20131625 nm", {
    x: 1.0, y: 5.05, w: 11, h: 0.3,
    fontSize: 9, fontFace: "Arial", color: TEXT_DIM,
  });

  addFooter(slide, "EM = s\u00e4hk\u00f6magneettinen, nm = nanometri (10\u207b\u2079 m)");
}

// =============================================
// SLIDE 4: Aallonpituus
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Aallonpituus (\u03bb)");

  addCard(slide, 0.4, 1.2, 6, 2.2, [
    { text: "Aallonpituus = \u03bb (lambda)", options: { fontSize: 16, bold: true, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Mitataan nanometrein\u00e4 (nm). Yksi nanometri = 0.000001 mm.", options: { fontSize: 13, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Eri aallonpituuksilla on eri ominaisuudet:", options: { fontSize: 13, breakLine: true } },
    { text: "  Eri vaimennus, eri dispersio, eri k\u00e4ytt\u00f6tarkoitus", options: { fontSize: 12, color: TEXT_DIM } },
  ], { accentTop: CYAN });

  addCard(slide, 6.8, 1.2, 5.8, 2.2, [
    { text: "Analogia", options: { fontSize: 16, bold: true, color: PURPLE, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Aallonpituus on kuin radiokanava:", options: { fontSize: 13, breakLine: true } },
    { text: "  Yksi kuitu voi kuljettaa useita \u03bb samanaikaisesti (WDM).", options: { fontSize: 13, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Jokainen \u03bb on itsen\u00e4inen datakanava.", options: { fontSize: 13 } },
  ], { accentTop: PURPLE });

  // Wavelength table
  const wlRows = [
    [
      { text: "Aallonpituus", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
      { text: "Alue", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
      { text: "K\u00e4ytt\u00f6", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
      { text: "Vaimennus", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
    ],
    [
      { text: "850 nm", options: { color: GOLD, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "Monimoodi", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "LAN, datakeskus", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "~2.5 dB/km", options: { color: RED, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
    [
      { text: "1310 nm", options: { color: CYAN, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "O-band", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "GPON upstream", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "~0.35 dB/km", options: { color: GOLD, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
    ],
    [
      { text: "1490 nm", options: { color: CYAN, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "S-band", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "GPON downstream", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "~0.25 dB/km", options: { color: GREEN, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
    [
      { text: "1550 nm", options: { color: CYAN, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "C-band", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "Pitk\u00e4t v\u00e4lit, DWDM", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "~0.20 dB/km", options: { color: GREEN, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
    ],
  ];
  slide.addTable(wlRows, {
    x: 0.8, y: 3.8, w: 11.4,
    colW: [2.2, 2.2, 4.0, 3.0],
    border: { type: "solid", pt: 0.5, color: "2D3748" },
    rowH: [0.4, 0.35, 0.35, 0.35, 0.35],
    fontFace: "Arial",
  });

  addFooter(slide, "\u03bb = lambda, kreikkalainen kirjain aallonpituudelle");
}

// =============================================
// SLIDE 5: Valo kuidussa
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Valo kuidussa");

  addCard(slide, 0.4, 1.2, 4.0, 3.0, [
    { text: "Kokonaisheijastus", options: { fontSize: 16, bold: true, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Valo etenee kuidussa kokonaisheijastuksen avulla.", options: { fontSize: 13, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Ydin (core) = korkeampi taitekerroin", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: "Kuori (cladding) = matalampi taitekerroin", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Valo \"pomppii\" ytimen ja kuoren rajapinnalla eik\u00e4 p\u00e4\u00e4se ulos.", options: { fontSize: 12, color: TEXT_DIM } },
  ], { accentTop: CYAN });

  addCard(slide, 4.7, 1.2, 3.9, 3.0, [
    { text: "Yksimoodi (SM)", options: { fontSize: 16, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Ydin: 9 \u00b5m", options: { fontSize: 13, color: GREEN, breakLine: true } },
    { text: "Kuori: 125 \u00b5m", options: { fontSize: 13, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Yksi valopolku \u2192 ei moodidispersiota", options: { fontSize: 12, breakLine: true } },
    { text: "Pitk\u00e4t v\u00e4lit (jopa 100+ km)", options: { fontSize: 12, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "K\u00e4yt\u00f6ss\u00e4 k\u00e4yt\u00e4nn\u00f6ss\u00e4 kaikissa FTTH-verkoissa", options: { fontSize: 11, color: TEXT_DIM } },
  ], { accentTop: GOLD });

  addCard(slide, 8.9, 1.2, 3.8, 3.0, [
    { text: "Monimoodi (MM)", options: { fontSize: 16, bold: true, color: PURPLE, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Ydin: 50 tai 62.5 \u00b5m", options: { fontSize: 13, color: GREEN, breakLine: true } },
    { text: "Kuori: 125 \u00b5m", options: { fontSize: 13, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Monta valopolkua \u2192 dispersio rajoittaa", options: { fontSize: 12, breakLine: true } },
    { text: "Lyhyet v\u00e4lit (max ~550 m)", options: { fontSize: 12, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "K\u00e4ytt\u00f6: datakeskukset, LAN", options: { fontSize: 11, color: TEXT_DIM } },
  ], { accentTop: PURPLE });

  // SM vs MM comparison bar
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 4.6, w: 12.2, h: 1.2,
    fill: { color: CARD_BG }, rectRadius: 0.1,
  });
  slide.addText("SM 9/125 \u00b5m", {
    x: 0.8, y: 4.7, w: 4, h: 0.4,
    fontSize: 14, fontFace: "Arial", bold: true, color: GOLD,
  });
  slide.addText("FTTH-verkon standardi. Kaikki PON-teknologiat k\u00e4ytt\u00e4v\u00e4t yksimoodikuitua.", {
    x: 0.8, y: 5.1, w: 11, h: 0.4,
    fontSize: 12, fontFace: "Arial", color: TEXT_PRIMARY,
  });

  addFooter(slide, "SM = Single-Mode, MM = Multi-Mode, \u00b5m = mikrometri");
}

// =============================================
// SLIDE 6: Section — Kuituverkko
// =============================================
addSectionSlide("Osa II: Kuituverkko", "Rakenneosat, kuitutyypit ja liittimet");

// =============================================
// SLIDE 7: Aallonpituusikkunat
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Aallonpituusikkunat (Bands)");

  const bandRows = [
    [
      { text: "Band", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
      { text: "Alue (nm)", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
      { text: "Nimi", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
      { text: "Vaimennus", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
      { text: "K\u00e4ytt\u00f6", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
    ],
    [
      { text: "O", options: { color: CYAN, fill: { color: TABLE_ROW1 }, fontSize: 11, bold: true } },
      { text: "1260\u20131360", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "Original", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "~0.35 dB/km", options: { color: GOLD, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "GPON/XGS-PON US", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
    [
      { text: "E", options: { color: CYAN, fill: { color: TABLE_ROW2 }, fontSize: 11, bold: true } },
      { text: "1360\u20131460", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "Extended", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "Vesihuippu!", options: { color: RED, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "V\u00e4ltet\u00e4\u00e4n (OH\u207b)", options: { color: TEXT_DIM, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
    ],
    [
      { text: "S", options: { color: CYAN, fill: { color: TABLE_ROW1 }, fontSize: 11, bold: true } },
      { text: "1460\u20131530", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "Short", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "~0.25 dB/km", options: { color: GREEN, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "GPON DS (1490)", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
    [
      { text: "C", options: { color: CYAN, fill: { color: TABLE_ROW2 }, fontSize: 11, bold: true } },
      { text: "1530\u20131565", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "Conventional", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "~0.20 dB/km", options: { color: GREEN, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "XGS-PON DS, DWDM", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
    ],
    [
      { text: "L", options: { color: CYAN, fill: { color: TABLE_ROW1 }, fontSize: 11, bold: true } },
      { text: "1565\u20131625", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "Long", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "~0.22 dB/km", options: { color: GREEN, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "CWDM, OTDR", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
  ];
  slide.addTable(bandRows, {
    x: 0.5, y: 1.2, w: 12,
    colW: [1.2, 2.2, 2.2, 2.4, 4.0],
    border: { type: "solid", pt: 0.5, color: "2D3748" },
    rowH: [0.42, 0.38, 0.38, 0.38, 0.38, 0.38],
    fontFace: "Arial",
  });

  // Water peak callout
  addCard(slide, 0.5, 4.0, 5.5, 1.6, [
    { text: "Vesihuippu (OH\u207b)", options: { fontSize: 14, bold: true, color: RED, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "E-bandin kohdalla (~1383 nm) vaimennus nousee voimakkaasti hydroksyyli-ionien vuoksi.", options: { fontSize: 12, breakLine: true } },
    { text: "Moderni G.652.D-kuitu on \"low water peak\" ja mahdollistaa E-bandin k\u00e4yt\u00f6n.", options: { fontSize: 12, color: GREEN } },
  ], { accentTop: RED });

  addCard(slide, 6.3, 4.0, 6.3, 1.6, [
    { text: "Vaimennusk\u00e4yr\u00e4n muoto", options: { fontSize: 14, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Paras: C-band ~0.20 dB/km (1550 nm)", options: { fontSize: 12, color: GREEN, breakLine: true } },
    { text: "Hyv\u00e4: L-band ~0.22 dB/km", options: { fontSize: 12, color: GREEN, breakLine: true } },
    { text: "OK: O-band ~0.35 dB/km (1310 nm)", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: "Huono: 850 nm ~2.5 dB/km (monimoodi)", options: { fontSize: 12, color: RED } },
  ], { accentTop: GOLD });

  addFooter(slide, "OH\u207b = hydroksyyli-ioni, OTDR = Optical Time Domain Reflectometer");
}

// =============================================
// SLIDE 8: Kuituverkon rakenneosat
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Kuituverkon rakenneosat");

  // Top row: Network elements
  const elements = [
    { label: "OLT", desc: "Optical Line Terminal\nKeskuslaite, operaattori", color: PURPLE, x: 0.3 },
    { label: "ODF", desc: "Optical Distribution Frame\nKuitujen ristikytkent\u00e4", color: CYAN, x: 2.45 },
    { label: "FDH", desc: "Fiber Distribution Hub\nJakokaappi (ulko)", color: GOLD, x: 4.6 },
    { label: "FDT/FAT", desc: "Fiber Distribution/Access Terminal\nTalohaarikkeet", color: GREEN, x: 6.75 },
    { label: "ONT", desc: "Optical Network Terminal\nAsiakasp\u00e4\u00e4te", color: CYAN, x: 8.9 },
    { label: "Splitter", desc: "Passiivinen jakaja\n1:N (N=2\u201364)", color: PURPLE, x: 11.05 },
  ];

  elements.forEach(el => {
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: el.x, y: 1.15, w: 1.95, h: 1.8,
      fill: { color: CARD_BG }, rectRadius: 0.08,
    });
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: el.x + 0.03, y: 1.15, w: 1.89, h: 0.04,
      fill: { color: el.color },
    });
    slide.addText(el.label, {
      x: el.x + 0.1, y: 1.3, w: 1.75, h: 0.4,
      fontSize: 15, fontFace: "Arial", bold: true, color: el.color,
    });
    slide.addText(el.desc, {
      x: el.x + 0.1, y: 1.7, w: 1.75, h: 1.1,
      fontSize: 10, fontFace: "Arial", color: TEXT_PRIMARY, lineSpacingMultiple: 1.15,
    });
  });

  // Arrow line between elements
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.3, y: 3.1, w: 12.7, h: 0.03, fill: { color: CYAN },
  });

  // Bottom: Cable types
  addCard(slide, 0.3, 3.5, 4.0, 2.5, [
    { text: "Kaapelit", options: { fontSize: 15, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Runkokaapeli", options: { fontSize: 13, bold: true, color: CYAN, breakLine: true } },
    { text: "  48\u2013288 kuitua, kanavistossa", options: { fontSize: 11, breakLine: true } },
    { text: "Jakelukaapeli", options: { fontSize: 13, bold: true, color: CYAN, breakLine: true } },
    { text: "  12\u201348 kuitua, alueverkko", options: { fontSize: 11, breakLine: true } },
    { text: "Talokaapeli (drop)", options: { fontSize: 13, bold: true, color: CYAN, breakLine: true } },
    { text: "  1\u20134 kuitua, asiakkaalle", options: { fontSize: 11 } },
  ], { accentTop: GOLD });

  addCard(slide, 4.6, 3.5, 4.0, 2.5, [
    { text: "Kanavisto", options: { fontSize: 15, bold: true, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Mikroputki", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "  5/3.5\u201314/10 mm, kaapelien suojaus", options: { fontSize: 11, breakLine: true } },
    { text: "Suojaputki", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "  40\u2013110 mm, sis\u00e4lt\u00e4\u00e4 mikroputket", options: { fontSize: 11, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Puhallustekniikka: kaapeli puhalletaan mikroputkeen paineilmalla", options: { fontSize: 10, color: TEXT_DIM } },
  ], { accentTop: CYAN });

  addCard(slide, 8.9, 3.5, 4.0, 2.5, [
    { text: "Topologia", options: { fontSize: 15, bold: true, color: PURPLE, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "OLT \u2192 ODF \u2192 Runko \u2192 FDH", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: "FDH \u2192 Jakelu \u2192 FDT/FAT", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: "FDT/FAT \u2192 Drop \u2192 ONT", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Splitterit sijoitetaan yleens\u00e4 FDH:n tai FDT:n yhteyteen.", options: { fontSize: 10, color: TEXT_DIM } },
  ], { accentTop: PURPLE });
}

// =============================================
// SLIDE 9: Kuitutyypit
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Kuitutyypit (ITU-T suositukset)");

  const fiberRows = [
    [
      { text: "Tyyppi", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
      { text: "Min. taivutuss\u00e4de", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
      { text: "K\u00e4ytt\u00f6kohde", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
      { text: "Huomiot", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
    ],
    [
      { text: "G.652.D", options: { color: CYAN, fill: { color: TABLE_ROW1 }, fontSize: 12, bold: true } },
      { text: "30 mm", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 12 } },
      { text: "Runko- ja jakeluverkko", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 12 } },
      { text: "Standardi SM-kuitu, low water peak", options: { color: TEXT_DIM, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
    [
      { text: "G.657.A1", options: { color: GOLD, fill: { color: TABLE_ROW2 }, fontSize: 12, bold: true } },
      { text: "10 mm", options: { color: GREEN, fill: { color: TABLE_ROW2 }, fontSize: 12 } },
      { text: "Taloverkko, sis\u00e4johdotus", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 12 } },
      { text: "Yhteensopiva G.652.D kanssa", options: { color: TEXT_DIM, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
    ],
    [
      { text: "G.657.A2", options: { color: GOLD, fill: { color: TABLE_ROW1 }, fontSize: 12, bold: true } },
      { text: "7.5 mm", options: { color: GREEN, fill: { color: TABLE_ROW1 }, fontSize: 12 } },
      { text: "Ahdas sis\u00e4asennus, MDU", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 12 } },
      { text: "Paras taivutuskesto, yhteensopiva", options: { color: TEXT_DIM, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
  ];
  slide.addTable(fiberRows, {
    x: 0.5, y: 1.2, w: 12,
    colW: [2.5, 2.5, 3.5, 3.5],
    border: { type: "solid", pt: 0.5, color: "2D3748" },
    rowH: [0.42, 0.4, 0.4, 0.4],
    fontFace: "Arial",
  });

  addCard(slide, 0.5, 3.3, 5.8, 2.6, [
    { text: "Taivutuss\u00e4de k\u00e4yt\u00e4nn\u00f6ss\u00e4", options: { fontSize: 15, bold: true, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Liian pieni taivutuss\u00e4de = vaimennus kasvaa!", options: { fontSize: 13, color: RED, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "G.652.D: sopii kanavistoon ja jakokaappeihin", options: { fontSize: 12, breakLine: true } },
    { text: "G.657.A1: sopii talojakamoon ja nousuun", options: { fontSize: 12, breakLine: true } },
    { text: "G.657.A2: sopii ahtaisiin nurkkiin, MDU-asennuksiin", options: { fontSize: 12, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Kaikki G.657-kuidut ovat yhteensopivia G.652.D kanssa \u2192 voidaan jatkaa samaan verkkoon.", options: { fontSize: 11, color: GREEN } },
  ], { accentTop: CYAN });

  addCard(slide, 6.6, 3.3, 6, 2.6, [
    { text: "Valintaohje", options: { fontSize: 15, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Runkoverkko:", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "  G.652.D \u2014 edullisin, riitt\u00e4v\u00e4 taivutus", options: { fontSize: 12, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Jakeluverkko:", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "  G.652.D tai G.657.A1", options: { fontSize: 12, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Drop + sis\u00e4verkko:", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "  G.657.A1 tai A2", options: { fontSize: 12, color: CYAN } },
  ], { accentTop: GOLD });
}

// =============================================
// SLIDE 10: Liittimet
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Kuituliittimet");

  // SC/APC card
  addCard(slide, 0.4, 1.2, 4.0, 2.8, [
    { text: "SC/APC", options: { fontSize: 18, bold: true, color: GREEN, breakLine: true } },
    { text: "Vihre\u00e4 liitin", options: { fontSize: 13, color: GREEN, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Angled Physical Contact (8\u00b0 kulma)", options: { fontSize: 12, breakLine: true } },
    { text: "Paluuvaimennus: \u226565 dB", options: { fontSize: 13, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "PON-verkon standardi", options: { fontSize: 12, bold: true, breakLine: true } },
    { text: "Heijastus ohjautuu pois kuidun akselilta \u2192 paras WDM-yhteensopivuus", options: { fontSize: 11, color: TEXT_DIM } },
  ], { accentTop: GREEN });

  // SC/UPC card
  addCard(slide, 4.7, 1.2, 4.0, 2.8, [
    { text: "SC/UPC", options: { fontSize: 18, bold: true, color: "3B82F6", breakLine: true } },
    { text: "Sininen liitin", options: { fontSize: 13, color: "3B82F6", breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Ultra Physical Contact (suora)", options: { fontSize: 12, breakLine: true } },
    { text: "Paluuvaimennus: \u226550 dB", options: { fontSize: 13, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Dataverkot, mittalaitteet", options: { fontSize: 12, bold: true, breakLine: true } },
    { text: "EI PON-verkkoon \u2014 heijastukset h\u00e4iritsev\u00e4t", options: { fontSize: 11, color: RED } },
  ], { accentTop: "3B82F6" });

  // LC card
  addCard(slide, 8.9, 1.2, 3.8, 2.8, [
    { text: "LC", options: { fontSize: 18, bold: true, color: CYAN, breakLine: true } },
    { text: "Pieni liitin", options: { fontSize: 13, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Lucent Connector", options: { fontSize: 12, breakLine: true } },
    { text: "Puolet SC:n koosta", options: { fontSize: 13, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "OLT, aktiivilaitteet, SFP-moduulit", options: { fontSize: 12, bold: true, breakLine: true } },
    { text: "Saatavana APC ja UPC versioina", options: { fontSize: 11, color: TEXT_DIM } },
  ], { accentTop: CYAN });

  // Cleaning note
  addCard(slide, 0.4, 4.3, 12.3, 1.5, [
    { text: "Puhdistus on kriittist\u00e4!", options: { fontSize: 16, bold: true, color: RED, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Likainen liitin = heijastuksia + vaimennusta. Puhdista AINA ennen kytkent\u00e4\u00e4.", options: { fontSize: 13, breakLine: true } },
    { text: "Ty\u00f6kalut: kuivapuhdistin (IBC, Cletop), isopropanoli + nukkaamaton liina, tarkastusmikroskooppi (200\u2013400x).", options: { fontSize: 12, color: TEXT_DIM } },
  ], { accentTop: RED });

  addFooter(slide, "APC = Angled Physical Contact, UPC = Ultra Physical Contact, RL = Return Loss");
}

// =============================================
// SLIDE 11: Laserturvallisuus
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Laserturvallisuus");

  addCard(slide, 0.4, 1.2, 6, 2.5, [
    { text: "PON-laitteet: Class 1 / 1M", options: { fontSize: 16, bold: true, color: GREEN, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Class 1: Turvallinen kaikissa olosuhteissa", options: { fontSize: 13, breakLine: true } },
    { text: "Class 1M: Turvallinen paljaalle silm\u00e4lle", options: { fontSize: 13, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "OLT, ONT ja SFP-moduulit ovat Class 1 tai 1M.", options: { fontSize: 12, color: TEXT_DIM, breakLine: true } },
    { text: "\u00c4L\u00c4 koskaan katso kuidun p\u00e4\u00e4h\u00e4n optisilla laitteilla!", options: { fontSize: 12, color: RED } },
  ], { accentTop: GREEN });

  addCard(slide, 6.8, 1.2, 5.8, 2.5, [
    { text: "VFL: Class 2 / 3R", options: { fontSize: 16, bold: true, color: RED, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Visual Fault Locator = n\u00e4kyv\u00e4 punainen laser", options: { fontSize: 13, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Class 2: < 1 mW, silm\u00e4n sulkurefleksi suojaa", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: "Class 3R: 1\u20135 mW, vaatii varovaisuutta!", options: { fontSize: 12, color: RED, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "K\u00e4yt\u00e4 aina suojalaseja Class 3R kanssa.", options: { fontSize: 12, color: RED, bold: true } },
  ], { accentTop: RED });

  // Safety labels
  addCard(slide, 0.4, 4.1, 12.3, 1.8, [
    { text: "Turvallisuusmerkinn\u00e4t ja k\u00e4yt\u00e4nn\u00f6t", options: { fontSize: 15, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Kaikissa laserlaitteissa oltava luokkamerkint\u00e4 (IEC 60825-1)", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "\u00c4l\u00e4 koskaan katso suoraan kuidun p\u00e4\u00e4h\u00e4n \u2014 k\u00e4yt\u00e4 tehomittaria", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Suojaa avoimet liittimet aina tulpilla", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Kuitukatkokset ja puhtaat p\u00e4\u00e4t eiv\u00e4t tarkoita \"turvallista\" \u2014 mittaa ensin", options: { fontSize: 12, bullet: true } },
  ], { accentTop: GOLD });
}

// =============================================
// SLIDE 12: Section — PON ja suunnittelu
// =============================================
addSectionSlide("Osa III: PON ja suunnittelu", "Arkkitehtuurit, h\u00e4vi\u00f6budjetti ja yhteiselo");

// =============================================
// SLIDE 13: PON-arkkitehtuurit
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "PON-arkkitehtuurit");

  const ponRows = [
    [
      { text: "Teknologia", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 11 } },
      { text: "DS / US", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 11 } },
      { text: "DS \u03bb (nm)", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 11 } },
      { text: "US \u03bb (nm)", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 11 } },
      { text: "Budjetti", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 11 } },
      { text: "Standardi", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 11 } },
    ],
    [
      { text: "GPON", options: { color: CYAN, fill: { color: TABLE_ROW1 }, fontSize: 11, bold: true } },
      { text: "2.5 / 1.25 Gbps", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "1490", options: { color: GOLD, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "1310", options: { color: GOLD, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "B+ 28, C+ 32", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 10 } },
      { text: "G.984", options: { color: TEXT_DIM, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
    [
      { text: "XGS-PON", options: { color: CYAN, fill: { color: TABLE_ROW2 }, fontSize: 11, bold: true } },
      { text: "10 / 10 Gbps", options: { color: GREEN, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "1577", options: { color: GOLD, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "1270", options: { color: GOLD, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "N1 29, N2 31", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 10 } },
      { text: "G.9807", options: { color: TEXT_DIM, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
    ],
    [
      { text: "25GS-PON", options: { color: PURPLE, fill: { color: TABLE_ROW1 }, fontSize: 11, bold: true } },
      { text: "25 / 25 Gbps", options: { color: GREEN, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "1358", options: { color: GOLD, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "1270", options: { color: GOLD, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "N1 29", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 10 } },
      { text: "G.9806", options: { color: TEXT_DIM, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
    [
      { text: "50G-PON", options: { color: PURPLE, fill: { color: TABLE_ROW2 }, fontSize: 11, bold: true } },
      { text: "50 / 50 Gbps", options: { color: GREEN, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "1340", options: { color: GOLD, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "1280", options: { color: GOLD, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "N1 29", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 10 } },
      { text: "G.9804", options: { color: TEXT_DIM, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
    ],
    [
      { text: "P2P (ei PON)", options: { color: TEXT_DIM, fill: { color: TABLE_ROW1 }, fontSize: 11, bold: true } },
      { text: "1\u2013100 Gbps", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "1310/1550", options: { color: TEXT_DIM, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "1310/1550", options: { color: TEXT_DIM, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "Linkki-koht.", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 10 } },
      { text: "802.3", options: { color: TEXT_DIM, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
  ];
  slide.addTable(ponRows, {
    x: 0.3, y: 1.15, w: 12.4,
    colW: [2.0, 2.2, 1.6, 1.6, 2.2, 2.8],
    border: { type: "solid", pt: 0.5, color: "2D3748" },
    rowH: [0.4, 0.36, 0.36, 0.36, 0.36, 0.36],
    fontFace: "Arial",
  });

  addCard(slide, 0.3, 3.7, 6, 2.2, [
    { text: "PON = Passive Optical Network", options: { fontSize: 15, bold: true, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Ei aktiivisia laitteita OLT:n ja ONT:n v\u00e4lill\u00e4", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Yksi kuitu jaetaan splittereill\u00e4 (1:32 tai 1:64)", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Passiivisuus = ei s\u00e4hk\u00f6\u00e4, ei huoltoa kent\u00e4ll\u00e4", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "GPON on yleisin, XGS-PON yleistyy nopeasti", options: { fontSize: 12, bullet: true, color: GREEN } },
  ], { accentTop: CYAN });

  addCard(slide, 6.6, 3.7, 6.1, 2.2, [
    { text: "P2P vs PON", options: { fontSize: 15, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "P2P: Oma kuitu jokaiselle asiakkaalle", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "  + Paras suorituskyky, ei jakoa", options: { fontSize: 11, color: GREEN, breakLine: true } },
    { text: "  - Kallis (paljon kuituja ja portteja)", options: { fontSize: 11, color: RED, breakLine: true } },
    { text: "PON: Jaettu kuitu, jaettu kaistanleveys", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "  + Edullinen, v\u00e4hemm\u00e4n kuituja", options: { fontSize: 11, color: GREEN, breakLine: true } },
    { text: "  - Kaista jaetaan k\u00e4ytt\u00e4jien kesken", options: { fontSize: 11, color: RED } },
  ], { accentTop: GOLD });
}

// =============================================
// SLIDE 14: GPON-signalointi
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "GPON-signalointi");

  addCard(slide, 0.4, 1.2, 6, 3.0, [
    { text: "Downstream (1490 nm)", options: { fontSize: 16, bold: true, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Broadcast-periaate:", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "OLT l\u00e4hett\u00e4\u00e4 kaiken kaikille ONT:ille", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Jokainen ONT lukee vain omat kehyksens\u00e4 (GEM Port ID)", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Salaus: AES-128", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "Koska kaikki ONT:t n\u00e4kev\u00e4t kaiken datan, downstream on AES-128-salattu.", options: { fontSize: 11, color: TEXT_DIM, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Nopeus: 2.488 Gbps jaettu kaikille ONT:ille", options: { fontSize: 12, color: GOLD } },
  ], { accentTop: CYAN });

  addCard(slide, 6.8, 1.2, 5.8, 3.0, [
    { text: "Upstream (1310 nm)", options: { fontSize: 16, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "TDMA Burst-l\u00e4hetys:", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "ONT:t l\u00e4hett\u00e4v\u00e4t vuorotellen aikav\u00e4leiss\u00e4", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "OLT jakaa aikav\u00e4lit DBA-algoritmilla", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "DBA = Dynamic Bandwidth Allocation", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "Aktiiviset ONT:t saavat enemm\u00e4n kaistaa, hiljaiset v\u00e4hemm\u00e4n.", options: { fontSize: 11, color: TEXT_DIM, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Nopeus: 1.244 Gbps jaettu kaikille ONT:ille", options: { fontSize: 12, color: GOLD } },
  ], { accentTop: GOLD });

  // Timing diagram illustration
  addCard(slide, 0.4, 4.5, 12.3, 1.4, [
    { text: "Aikajakoperiaate (TDMA)", options: { fontSize: 14, bold: true, color: PURPLE, breakLine: true } },
    { text: "ONT1 [====]                ONT2 [====]                ONT3 [====]                ONT1 [====] ...", options: { fontSize: 11, color: CYAN, breakLine: true, fontFace: "Courier New" } },
    { text: "Ranging-prosessi synkronoi ONT:iden et\u00e4isyydet. Guard time est\u00e4\u00e4 p\u00e4\u00e4llekk\u00e4isyydet.", options: { fontSize: 11, color: TEXT_DIM } },
  ], { accentTop: PURPLE });
}

// =============================================
// SLIDE 15: Haviobudjetti
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "H\u00e4vi\u00f6budjetti (Loss Budget)");

  // Loss components
  addCard(slide, 0.4, 1.15, 6.0, 2.8, [
    { text: "H\u00e4vi\u00f6komponentit", options: { fontSize: 15, bold: true, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Kuituvaimennus:", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "  DS: 0.30 dB/km @ 1490 nm", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: "  US: 0.35 dB/km @ 1310 nm (m\u00e4\u00e4r\u00e4\u00e4v\u00e4!)", options: { fontSize: 12, color: RED, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Liitokset:", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "  Liittimet: 0.2\u20130.3 dB / kpl", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: "  Hitsausjatkokset: 0.05\u20130.1 dB / kpl", options: { fontSize: 12, color: GREEN, breakLine: true } },
    { text: "  Mek. jatkokset: 0.1\u20130.3 dB / kpl", options: { fontSize: 12, color: GOLD } },
  ], { accentTop: CYAN });

  // Splitter losses
  const splitRows = [
    [
      { text: "Jakosuhde", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 11 } },
      { text: "H\u00e4vi\u00f6 (dB)", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 11 } },
      { text: "Jakosuhde", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 11 } },
      { text: "H\u00e4vi\u00f6 (dB)", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 11 } },
    ],
    [
      { text: "1:2", options: { color: CYAN, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "3.5", options: { color: GREEN, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "1:16", options: { color: CYAN, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "14.5", options: { color: GOLD, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
    [
      { text: "1:4", options: { color: CYAN, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "7.0", options: { color: GREEN, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "1:32", options: { color: CYAN, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
      { text: "17.5", options: { color: RED, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
    ],
    [
      { text: "1:8", options: { color: CYAN, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "10.5", options: { color: GOLD, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "1:64", options: { color: CYAN, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
      { text: "21.0", options: { color: RED, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
  ];
  slide.addTable(splitRows, {
    x: 6.8, y: 1.15, w: 5.8,
    colW: [1.5, 1.4, 1.5, 1.4],
    border: { type: "solid", pt: 0.5, color: "2D3748" },
    rowH: [0.38, 0.35, 0.35, 0.35],
    fontFace: "Arial",
  });
  slide.addText("Splitterin h\u00e4vi\u00f6t", {
    x: 6.8, y: 2.65, w: 5.8, h: 0.3,
    fontSize: 10, fontFace: "Arial", color: TEXT_DIM, align: "center",
  });

  // Budget classes
  const budgetRows = [
    [
      { text: "Luokka", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
      { text: "Max h\u00e4vi\u00f6 (dB)", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
      { text: "Teknologia", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
      { text: "Tyypillinen k\u00e4ytt\u00f6", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 12 } },
    ],
    [
      { text: "B+ (GPON)", options: { color: CYAN, fill: { color: TABLE_ROW1 }, fontSize: 12, bold: true } },
      { text: "28 dB", options: { color: GOLD, fill: { color: TABLE_ROW1 }, fontSize: 12 } },
      { text: "GPON", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 12 } },
      { text: "Perus GPON", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
    [
      { text: "C+ (GPON)", options: { color: CYAN, fill: { color: TABLE_ROW2 }, fontSize: 12, bold: true } },
      { text: "32 dB", options: { color: GREEN, fill: { color: TABLE_ROW2 }, fontSize: 12 } },
      { text: "GPON", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 12 } },
      { text: "Pitk\u00e4t v\u00e4lit, 1:64", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
    ],
    [
      { text: "N1 (XGS)", options: { color: PURPLE, fill: { color: TABLE_ROW1 }, fontSize: 12, bold: true } },
      { text: "29 dB", options: { color: GOLD, fill: { color: TABLE_ROW1 }, fontSize: 12 } },
      { text: "XGS-PON", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 12 } },
      { text: "Perus XGS-PON", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 11 } },
    ],
    [
      { text: "N2 (XGS)", options: { color: PURPLE, fill: { color: TABLE_ROW2 }, fontSize: 12, bold: true } },
      { text: "31 dB", options: { color: GREEN, fill: { color: TABLE_ROW2 }, fontSize: 12 } },
      { text: "XGS-PON", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 12 } },
      { text: "Extended reach", options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 11 } },
    ],
  ];
  slide.addTable(budgetRows, {
    x: 0.4, y: 4.3, w: 12.2,
    colW: [2.5, 2.5, 2.5, 4.7],
    border: { type: "solid", pt: 0.5, color: "2D3748" },
    rowH: [0.4, 0.35, 0.35, 0.35, 0.35],
    fontFace: "Arial",
  });

  addFooter(slide, "H\u00e4vi\u00f6budjetti = OLT:n ja ONT:n v\u00e4lisen optisen reitin kokonaish\u00e4vi\u00f6. US 1310 nm on m\u00e4\u00e4r\u00e4\u00e4v\u00e4 (suurempi vaimennus).");
}

// =============================================
// SLIDE 16: Splitteritopologiat
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Splitteritopologiat");

  // Star/Tree
  addCard(slide, 0.4, 1.2, 4.0, 2.6, [
    { text: "T\u00e4hti / Puu", options: { fontSize: 16, bold: true, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Yksi splitter FDH:ssa jakaa kaikille", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Yksinkertainen, helppo hallita", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Tyypillisesti 1:32 tai 1:64", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Yleisin topologia FTTH:ssa", options: { fontSize: 12, color: GREEN, bold: true } },
  ], { accentTop: CYAN });

  // Daisy-chain
  addCard(slide, 4.7, 1.2, 4.0, 2.6, [
    { text: "Ketju (Daisy-chain)", options: { fontSize: 16, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Splitterit per\u00e4kk\u00e4in reitill\u00e4", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Jokaisessa pisteess\u00e4 osa valosta haaroitetaan", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Sopii: tie- ja rautatiereittien varsille", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Haaste: ep\u00e4tasainen tehojakauma", options: { fontSize: 12, color: RED } },
  ], { accentTop: GOLD });

  // Cascade
  addCard(slide, 8.9, 1.2, 3.8, 2.6, [
    { text: "Kaskadi (2-vaihe)", options: { fontSize: 16, bold: true, color: PURPLE, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "1. vaihe: FDH (esim. 1:4)", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "2. vaihe: FDT (esim. 1:8)", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Yht. 1:32 (4 x 8)", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Joustava, laajennettava", options: { fontSize: 12, color: GREEN, bold: true } },
  ], { accentTop: PURPLE });

  // Comparison card
  addCard(slide, 0.4, 4.2, 12.3, 1.8, [
    { text: "Vertailu", options: { fontSize: 15, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "T\u00e4hti: Yksinkertaisin hallita, yksi vikapiste splitteriss\u00e4. Paras tiiviille alueelle.", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Ketju: V\u00e4hiten kuitua pitkill\u00e4 reiteill\u00e4, mutta vaikein optimoida.", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Kaskadi: Paras joustavuus \u2014 laajennettavissa ilman uutta kuitua. Suosituin uusissa verkoissa.", options: { fontSize: 12, bullet: true, color: GREEN } },
  ], { accentTop: GOLD });
}

// =============================================
// SLIDE 17: GPON + XGS-PON yhteiselo
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "GPON + XGS-PON yhteiselo");

  addCard(slide, 0.4, 1.2, 6, 3.2, [
    { text: "WDM1r \u2014 yhteiselo samalla kuidulla", options: { fontSize: 15, bold: true, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "GPON ja XGS-PON k\u00e4ytt\u00e4v\u00e4t eri aallonpituuksia:", options: { fontSize: 13, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "GPON DS:  1490 nm (S-band)", options: { fontSize: 12, color: CYAN, breakLine: true } },
    { text: "GPON US:  1310 nm (O-band)", options: { fontSize: 12, color: CYAN, breakLine: true } },
    { text: "XGS DS:   1577 nm (L-band)", options: { fontSize: 12, color: PURPLE, breakLine: true } },
    { text: "XGS US:   1270 nm (O-band)", options: { fontSize: 12, color: PURPLE, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "WDM1r CE (Coexistence Element) yhdist\u00e4\u00e4 molemmat OLT:t samaan kuituun.", options: { fontSize: 12, color: GOLD } },
  ], { accentTop: CYAN });

  addCard(slide, 6.8, 1.2, 5.8, 3.2, [
    { text: "Siirtym\u00e4strategia", options: { fontSize: 15, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "1. Asenna WDM1r CE OLT-p\u00e4\u00e4h\u00e4n", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "2. Lis\u00e4\u00e4 XGS-PON OLT rinnalle", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "3. Vaihda ONT:t asiakaskohtaisesti", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "4. GPON-ONT:t toimivat koko ajan!", options: { fontSize: 12, bullet: true, color: GREEN, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Hyv\u00e4:", options: { fontSize: 13, bold: true, color: GREEN, breakLine: true } },
    { text: "Ei uutta kaivua tai kuituasennusta", options: { fontSize: 12, color: GREEN, breakLine: true } },
    { text: "ONT-vaihto on ainoa fy. muutos", options: { fontSize: 12, color: GREEN, breakLine: true } },
    { text: "Asteittainen siirtym\u00e4 mahdollinen", options: { fontSize: 12, color: GREEN } },
  ], { accentTop: GOLD });

  // Wavelength diagram
  addCard(slide, 0.4, 4.7, 12.3, 1.2, [
    { text: "Aallonpituusallokaatio:", options: { fontSize: 13, bold: true, color: TEXT_PRIMARY } },
  ], { accentTop: PURPLE });

  // Wavelength bars
  const bands = [
    { label: "XGS US\n1270", x: 1.5, w: 1.2, color: PURPLE },
    { label: "GPON US\n1310", x: 3.0, w: 1.2, color: CYAN },
    { label: "GPON DS\n1490", x: 5.5, w: 1.2, color: CYAN },
    { label: "XGS DS\n1577", x: 8.0, w: 1.2, color: PURPLE },
  ];
  bands.forEach(b => {
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: b.x, y: 5.15, w: b.w, h: 0.55,
      fill: { color: b.color }, rectRadius: 0.05,
    });
    slide.addText(b.label, {
      x: b.x, y: 5.15, w: b.w, h: 0.55,
      fontSize: 10, fontFace: "Arial", color: "FFFFFF", align: "center", valign: "middle", bold: true,
    });
  });
}

// =============================================
// SLIDE 18: WDM-periaate
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "WDM-periaate (Wavelength Division Multiplexing)");

  addCard(slide, 0.4, 1.2, 6, 3.0, [
    { text: "Yksi kuitu, monta kanavaa", options: { fontSize: 16, bold: true, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "WDM = eri aallonpituudet samassa kuidussa samanaikaisesti", options: { fontSize: 13, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "MUX (Multiplexer):", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "  Yhdist\u00e4\u00e4 useita \u03bb yhdeksi kuiduksi", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "DEMUX (Demultiplexer):", options: { fontSize: 13, bold: true, breakLine: true } },
    { text: "  Erottelee \u03bb takaisin omiin kuituihinsa", options: { fontSize: 12, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "Passiivisia komponentteja \u2014 ei s\u00e4hk\u00f6\u00e4!", options: { fontSize: 12, color: GREEN } },
  ], { accentTop: CYAN });

  addCard(slide, 6.8, 1.2, 5.8, 3.0, [
    { text: "WDM-teknologiat", options: { fontSize: 16, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "CWDM (Coarse):", options: { fontSize: 14, bold: true, color: CYAN, breakLine: true } },
    { text: "  20 nm kanavav\u00e4li", options: { fontSize: 12, breakLine: true } },
    { text: "  Max 18 kanavaa (1270\u20131610 nm)", options: { fontSize: 12, breakLine: true } },
    { text: "  Edullinen, ei l\u00e4mp\u00f6tilans\u00e4\u00e4t\u00f6\u00e4", options: { fontSize: 12, color: GREEN, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "DWDM (Dense):", options: { fontSize: 14, bold: true, color: PURPLE, breakLine: true } },
    { text: "  0.8 nm kanavav\u00e4li (100 GHz)", options: { fontSize: 12, breakLine: true } },
    { text: "  40\u201396+ kanavaa C-bandissa", options: { fontSize: 12, breakLine: true } },
    { text: "  Kallis, l\u00e4mp\u00f6tilans\u00e4\u00e4t\u00f6 vaaditaan", options: { fontSize: 12, color: GOLD } },
  ], { accentTop: GOLD });

  addCard(slide, 0.4, 4.5, 12.3, 1.4, [
    { text: "WDM mahdollistaa kuituverkon kapasiteetin moninkertaistamisen ilman uutta kuitua.", options: { fontSize: 14, bold: true, color: GREEN, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "CWDM/DWDM-teknologiat k\u00e4sitell\u00e4\u00e4n tarkemmin erillisess\u00e4 koulutuksessa.", options: { fontSize: 12, color: TEXT_DIM } },
  ], { accentTop: PURPLE });
}

// =============================================
// SLIDE 19: Yhteenveto
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Yhteenveto");

  // Three summary columns
  addCard(slide, 0.4, 1.2, 4.0, 2.6, [
    { text: "Valon fysiikka", options: { fontSize: 15, bold: true, color: CYAN, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Infrapuna 850\u20131625 nm", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Kokonaisheijastus pit\u00e4\u00e4 valon kuidussa", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "SM 9/125 \u00b5m = FTTH-standardi", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Eri \u03bb = eri ominaisuudet", options: { fontSize: 12, bullet: true } },
  ], { accentTop: CYAN });

  addCard(slide, 4.7, 1.2, 4.0, 2.6, [
    { text: "Kuituverkko", options: { fontSize: 15, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "OLT \u2192 FDH \u2192 FDT \u2192 ONT", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "G.652.D + G.657.A1/A2", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "SC/APC (vihre\u00e4) PON-verkkoon", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Puhdistus on kriittist\u00e4!", options: { fontSize: 12, bullet: true, color: RED } },
  ], { accentTop: GOLD });

  addCard(slide, 8.9, 1.2, 3.8, 2.6, [
    { text: "PON ja suunnittelu", options: { fontSize: 15, bold: true, color: PURPLE, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "GPON 2.5G, XGS-PON 10G", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "H\u00e4vi\u00f6budjetti m\u00e4\u00e4r\u00e4\u00e4 kantaman", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "Kaskadi = joustavin topologia", options: { fontSize: 12, bullet: true, breakLine: true } },
    { text: "WDM = kapasiteetti x N", options: { fontSize: 12, bullet: true } },
  ], { accentTop: PURPLE });

  // Glossary
  addCard(slide, 0.4, 4.1, 12.3, 2.5, [
    { text: "Sanasto", options: { fontSize: 15, bold: true, color: GOLD, breakLine: true } },
    { text: " ", options: { fontSize: 4, breakLine: true } },
    { text: "APC = Angled Physical Contact  |  CWDM = Coarse WDM  |  DBA = Dynamic Bandwidth Allocation", options: { fontSize: 10, color: TEXT_DIM, breakLine: true } },
    { text: "DS = Downstream  |  DWDM = Dense WDM  |  FTTH = Fiber To The Home", options: { fontSize: 10, color: TEXT_DIM, breakLine: true } },
    { text: "FDH = Fiber Distribution Hub  |  FDT = Fiber Distribution Terminal  |  FAT = Fiber Access Terminal", options: { fontSize: 10, color: TEXT_DIM, breakLine: true } },
    { text: "GPON = Gigabit PON  |  ODF = Optical Distribution Frame  |  OLT = Optical Line Terminal", options: { fontSize: 10, color: TEXT_DIM, breakLine: true } },
    { text: "ONT = Optical Network Terminal  |  OTDR = Optical Time Domain Reflectometer", options: { fontSize: 10, color: TEXT_DIM, breakLine: true } },
    { text: "PON = Passive Optical Network  |  SM = Single-Mode  |  TDMA = Time Division Multiple Access", options: { fontSize: 10, color: TEXT_DIM, breakLine: true } },
    { text: "UPC = Ultra Physical Contact  |  US = Upstream  |  VFL = Visual Fault Locator  |  WDM = Wavelength Division Multiplexing", options: { fontSize: 10, color: TEXT_DIM, breakLine: true } },
    { text: "XGS-PON = 10 Gigabit Symmetric PON  |  \u03bb = aallonpituus (lambda)", options: { fontSize: 10, color: TEXT_DIM } },
  ], { accentTop: GOLD });
}

// =============================================
// Generate file
// =============================================
const outputPath = "/Users/jannekammonen/JK/jkammone/Claude/Repositories/koulutus/Kuituverkon_perusteet.pptx";
pptx.writeFile({ fileName: outputPath })
  .then(() => console.log("OK: " + outputPath))
  .catch(err => { console.error("ERROR:", err); process.exit(1); });
