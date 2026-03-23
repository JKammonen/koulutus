const pptxgen = require("pptxgenjs");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Janne Kammonen";
pptx.title = "CWDM & DWDM - Aallonpituuskanavointiteknologiat";

// Theme colors (no # prefix for pptxgenjs)
const BG_TITLE = "0A0E1A";
const BG_CONTENT = "111827";
const TEXT_PRIMARY = "E2E8F0";
const TEXT_DIM = "94A3B8";
const PURPLE = "8B5CF6";
const GOLD = "FACC15";
const CYAN = "06B6D4";
const BLUE = "3B82F6";
const GREEN = "22C55E";
const RED = "EF4444";
const CARD_BG = "1E293B";
const CARD_BG2 = "1A1F2E";
const TABLE_HEADER = "2D1B69";
const TABLE_ROW1 = "1E293B";
const TABLE_ROW2 = "162032";
const ORANGE = "F97316";
const WARN_BG = "422006";

// Helper: add title bar to content slides
function addTitleBar(slide, title) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: "100%", h: 0.9,
    fill: { color: CARD_BG2 },
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0.85, w: "100%", h: 0.06,
    fill: { color: PURPLE },
  });
  slide.addText(title, {
    x: 0.6, y: 0.15, w: 11, h: 0.6,
    fontSize: 32, fontFace: "Arial", bold: true, color: TEXT_PRIMARY,
  });
}

// Helper: section divider slide
function addSectionSlide(title) {
  const slide = pptx.addSlide();
  slide.background = { fill: BG_TITLE };
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 4.5, y: 2.2, w: 4, h: 0.06, fill: { color: PURPLE },
  });
  slide.addText(title, {
    x: 0.5, y: 2.5, w: 12, h: 1.2,
    fontSize: 44, fontFace: "Arial", bold: true, color: TEXT_PRIMARY, align: "center",
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 4.5, y: 3.8, w: 4, h: 0.06, fill: { color: PURPLE },
  });
  return slide;
}

// Helper: card background
function addCard(slide, x, y, w, h, color) {
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: w, h: h,
    fill: { color: color || CARD_BG },
    rectRadius: 0.1,
  });
}

// =============================================
// SLIDE 1: Title
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_TITLE };

  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: "100%", h: 0.08, fill: { color: PURPLE },
  });

  slide.addText("CWDM & DWDM", {
    x: 0.5, y: 1.8, w: 12, h: 1.4,
    fontSize: 54, fontFace: "Arial", bold: true, color: TEXT_PRIMARY, align: "center",
  });

  slide.addText("Aallonpituuskanavointiteknologiat", {
    x: 0.5, y: 3.2, w: 12, h: 0.8,
    fontSize: 24, fontFace: "Arial", color: PURPLE, align: "center",
  });

  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 3, y: 4.2, w: 7, h: 0.04, fill: { color: GOLD },
  });

  slide.addText("Coarse & Dense Wavelength Division Multiplexing", {
    x: 0.5, y: 4.5, w: 12, h: 0.6,
    fontSize: 16, fontFace: "Arial", color: TEXT_DIM, align: "center",
  });

  slide.addText("Janne Kammonen | Keypro Oy", {
    x: 0.5, y: 6.5, w: 12, h: 0.5,
    fontSize: 14, fontFace: "Arial", color: TEXT_DIM, align: "center",
  });
}

// =============================================
// SLIDE 2: Section - CWDM-tekniikka
// =============================================
addSectionSlide("CWDM-tekniikka");

// =============================================
// SLIDE 3: CWDM periaate
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "CWDM-periaate");

  // Left card: How it works
  addCard(slide, 0.4, 1.2, 6, 2.8, CARD_BG);
  slide.addText("Toimintaperiaate", {
    x: 0.6, y: 1.3, w: 5.5, h: 0.5,
    fontSize: 20, fontFace: "Arial", bold: true, color: PURPLE,
  });
  slide.addText([
    { text: "18 kanavaa", options: { bold: true, color: GOLD } },
    { text: " aallonpituusalueella 1270-1610 nm", options: { color: TEXT_PRIMARY } },
    { text: "\n\nKanavavaali: ", options: { color: TEXT_DIM } },
    { text: "20 nm", options: { bold: true, color: CYAN } },
    { text: "\n\nPassiivinen TFF MUX/DEMUX", options: { color: TEXT_PRIMARY } },
    { text: "\n(Thin Film Filter)", options: { color: TEXT_DIM } },
    { text: "\n\nStandardi: ", options: { color: TEXT_DIM } },
    { text: "ITU-T G.694.2", options: { bold: true, color: GOLD } },
  ], {
    x: 0.6, y: 1.85, w: 5.5, h: 2.0,
    fontSize: 15, fontFace: "Arial", lineSpacingMultiple: 0.95, valign: "top",
  });

  // Right card: Channel grid visual
  addCard(slide, 6.7, 1.2, 6, 2.8, CARD_BG);
  slide.addText("Kanavarasterit", {
    x: 6.9, y: 1.3, w: 5.5, h: 0.5,
    fontSize: 20, fontFace: "Arial", bold: true, color: PURPLE,
  });

  const channels = [
    "1270", "1290", "1310", "1330", "1350", "1370", "1390", "1410",
    "1430", "1450", "1470", "1490", "1510", "1530", "1550", "1570", "1590", "1610"
  ];
  const chColors = [
    "7C3AED", "6D28D9", "4F46E5", "3B82F6", "06B6D4", "14B8A6",
    "22C55E", "84CC16", "EAB308", "F59E0B", "F97316", "EF4444",
    "DC2626", "BE185D", "9333EA", "7C3AED", "6D28D9", "4F46E5"
  ];
  for (let i = 0; i < 18; i++) {
    const row = Math.floor(i / 6);
    const col = i % 6;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 6.9 + col * 0.9, y: 2.0 + row * 0.7, w: 0.8, h: 0.55,
      fill: { color: chColors[i] }, rectRadius: 0.05,
    });
    slide.addText(channels[i], {
      x: 6.9 + col * 0.9, y: 2.0 + row * 0.7, w: 0.8, h: 0.55,
      fontSize: 10, fontFace: "Arial", bold: true, color: "FFFFFF", align: "center", valign: "middle",
    });
  }
  slide.addText("18 kanavaa  |  20 nm vali  |  1270-1610 nm", {
    x: 6.9, y: 3.65, w: 5.5, h: 0.3,
    fontSize: 11, fontFace: "Arial", color: TEXT_DIM, align: "center",
  });

  // Bottom: MUX/DEMUX concept
  addCard(slide, 0.4, 4.3, 12.3, 2.6, CARD_BG2);
  slide.addText("MUX / DEMUX -periaate", {
    x: 0.6, y: 4.4, w: 5, h: 0.4,
    fontSize: 18, fontFace: "Arial", bold: true, color: CYAN,
  });

  // Visual: multiple wavelengths -> MUX -> single fiber -> DEMUX -> wavelengths
  const wlColors = [PURPLE, BLUE, CYAN, GREEN, GOLD, ORANGE];
  const wlLabels = ["1270", "1310", "1350", "1470", "1550", "1610"];
  for (let i = 0; i < 6; i++) {
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 0.8, y: 4.95 + i * 0.28, w: 1.2, h: 0.2,
      fill: { color: wlColors[i] },
    });
    slide.addText(wlLabels[i] + " nm", {
      x: 0.8, y: 4.95 + i * 0.28, w: 1.2, h: 0.2,
      fontSize: 8, fontFace: "Arial", bold: true, color: "FFFFFF", align: "center", valign: "middle",
    });
  }

  // MUX box
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 2.5, y: 5.0, w: 1.2, h: 1.5, fill: { color: "2D1B69" }, rectRadius: 0.1,
  });
  slide.addText("MUX", {
    x: 2.5, y: 5.4, w: 1.2, h: 0.5,
    fontSize: 14, fontFace: "Arial", bold: true, color: GOLD, align: "center", valign: "middle",
  });

  // Single fiber
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 4.0, y: 5.6, w: 3.5, h: 0.3, fill: { color: PURPLE },
  });
  slide.addText("1 kuitu", {
    x: 4.0, y: 5.25, w: 3.5, h: 0.3,
    fontSize: 12, fontFace: "Arial", bold: true, color: PURPLE, align: "center",
  });

  // DEMUX box
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 7.8, y: 5.0, w: 1.4, h: 1.5, fill: { color: "2D1B69" }, rectRadius: 0.1,
  });
  slide.addText("DEMUX", {
    x: 7.8, y: 5.4, w: 1.4, h: 0.5,
    fontSize: 14, fontFace: "Arial", bold: true, color: GOLD, align: "center", valign: "middle",
  });

  // Output wavelengths
  for (let i = 0; i < 6; i++) {
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 9.5, y: 4.95 + i * 0.28, w: 1.2, h: 0.2,
      fill: { color: wlColors[i] },
    });
    slide.addText(wlLabels[i] + " nm", {
      x: 9.5, y: 4.95 + i * 0.28, w: 1.2, h: 0.2,
      fontSize: 8, fontFace: "Arial", bold: true, color: "FFFFFF", align: "center", valign: "middle",
    });
  }
}

// =============================================
// SLIDE 4: CWDM edut
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "CWDM-edut");

  const advantages = [
    { icon: "x18", title: "Kapasiteetti", desc: "1 kuitu = jopa 18 signaalia\nKerrannaistaa kuituresurssit", color: PURPLE },
    { icon: "DFB", title: "Edulliset laserit", desc: "Jaahdyttamattomat DFB-laserit\n~1/3 DWDM-lasereiden hinnasta", color: GOLD },
    { icon: "20nm", title: "Leveaa toleranssi", desc: "20 nm kanavavaali sallii\nlammon aiheuttaman siirtyman", color: CYAN },
    { icon: "80km", title: "Hyvaa kantama", desc: "40-80 km ilman vahvistimia\nRiittaa metro/access-verkkoon", color: GREEN },
    { icon: "TFF", title: "Passiivinen MUX", desc: "Ei sahkoa vaativa laite\nThin Film Filter -tekniikka", color: BLUE },
    { icon: "OPEX", title: "Matala OPEX", desc: "Passiiviset komponentit\nEi huollettavia vahvistimia", color: ORANGE },
  ];

  for (let i = 0; i < 6; i++) {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.4 + col * 4.2;
    const y = 1.2 + row * 2.7;

    addCard(slide, x, y, 3.9, 2.4, CARD_BG);

    // Icon circle
    slide.addShape(pptx.shapes.OVAL, {
      x: x + 0.2, y: y + 0.25, w: 0.7, h: 0.7,
      fill: { color: advantages[i].color },
    });
    slide.addText(advantages[i].icon, {
      x: x + 0.2, y: y + 0.25, w: 0.7, h: 0.7,
      fontSize: 10, fontFace: "Arial", bold: true, color: "FFFFFF", align: "center", valign: "middle",
    });

    slide.addText(advantages[i].title, {
      x: x + 1.1, y: y + 0.25, w: 2.5, h: 0.5,
      fontSize: 18, fontFace: "Arial", bold: true, color: advantages[i].color,
    });
    slide.addText(advantages[i].desc, {
      x: x + 0.2, y: y + 1.1, w: 3.5, h: 1.1,
      fontSize: 13, fontFace: "Arial", color: TEXT_PRIMARY, lineSpacingMultiple: 1.1,
    });
  }
}

// =============================================
// SLIDE 5: G.652.D vaatimus
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "G.652.D-kuituvaatimus");

  // Warning card
  addCard(slide, 0.4, 1.2, 12.3, 2.2, WARN_BG);
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 1.2, w: 0.12, h: 2.2, fill: { color: GOLD },
  });
  slide.addText("HUOM!", {
    x: 0.8, y: 1.3, w: 3, h: 0.5,
    fontSize: 24, fontFace: "Arial", bold: true, color: GOLD,
  });
  slide.addText([
    { text: "Kaikki 18 CWDM-kanavaa toimivat vain ", options: { color: TEXT_PRIMARY } },
    { text: "G.652.D", options: { bold: true, color: GOLD } },
    { text: " -kuidussa!", options: { color: TEXT_PRIMARY } },
    { text: "\nVanhemmissa G.652.A/B -kuiduissa OH-piikki (1383 nm) vaimentaa kanavat 1370-1410 nm.", options: { color: TEXT_DIM } },
  ], {
    x: 0.8, y: 1.85, w: 11.5, h: 1.2,
    fontSize: 16, fontFace: "Arial", lineSpacingMultiple: 1.2,
  });

  // OH peak visualization
  addCard(slide, 0.4, 3.7, 12.3, 3.2, CARD_BG);
  slide.addText("OH-vaimennuspiikki", {
    x: 0.6, y: 3.8, w: 5, h: 0.5,
    fontSize: 20, fontFace: "Arial", bold: true, color: CYAN,
  });

  // Wavelength bar with OH peak
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 4.8, w: 11.5, h: 0.6, fill: { color: "1B3A2F" },
  });
  slide.addText("Kaytettavissa oleva kaista (G.652.D)", {
    x: 0.8, y: 4.8, w: 11.5, h: 0.6,
    fontSize: 12, fontFace: "Arial", color: GREEN, align: "center", valign: "middle",
  });

  // OH peak zone
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 3.8, y: 4.5, w: 1.8, h: 1.2, fill: { color: "3B1010" },
  });
  slide.addText("OH-piikki\n1370-1410 nm", {
    x: 3.8, y: 4.5, w: 1.8, h: 1.2,
    fontSize: 11, fontFace: "Arial", bold: true, color: RED, align: "center", valign: "middle",
  });

  // Labels
  slide.addText("1270 nm", {
    x: 0.6, y: 5.5, w: 1.5, h: 0.3,
    fontSize: 10, fontFace: "Arial", color: TEXT_DIM,
  });
  slide.addText("1610 nm", {
    x: 11.0, y: 5.5, w: 1.5, h: 0.3,
    fontSize: 10, fontFace: "Arial", color: TEXT_DIM, align: "right",
  });

  // Comparison table
  const compRows = [
    [
      { text: "Kuitutyyppi", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 13, fontFace: "Arial", align: "center" } },
      { text: "Kanavat", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 13, fontFace: "Arial", align: "center" } },
      { text: "OH-piikki", options: { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 13, fontFace: "Arial", align: "center" } },
    ],
    [
      { text: "G.652.D (low water peak)", options: { color: GREEN, fill: { color: TABLE_ROW1 }, fontSize: 12, fontFace: "Arial" } },
      { text: "18 kanavaa", options: { bold: true, color: GREEN, fill: { color: TABLE_ROW1 }, fontSize: 12, fontFace: "Arial", align: "center" } },
      { text: "Poistettu", options: { color: GREEN, fill: { color: TABLE_ROW1 }, fontSize: 12, fontFace: "Arial", align: "center" } },
    ],
    [
      { text: "G.652.A/B (vanha)", options: { color: RED, fill: { color: TABLE_ROW2 }, fontSize: 12, fontFace: "Arial" } },
      { text: "~14 kanavaa", options: { color: RED, fill: { color: TABLE_ROW2 }, fontSize: 12, fontFace: "Arial", align: "center" } },
      { text: "Korkea (>2 dB/km)", options: { color: RED, fill: { color: TABLE_ROW2 }, fontSize: 12, fontFace: "Arial", align: "center" } },
    ],
  ];
  slide.addTable(compRows, {
    x: 1.5, y: 6.0, w: 10, h: 0.9,
    border: { type: "solid", pt: 0.5, color: "374151" },
    colW: [4, 3, 3],
  });
}

// =============================================
// SLIDE 6: Section - CWDM vs DWDM
// =============================================
addSectionSlide("CWDM vs DWDM");

// =============================================
// SLIDE 7: Vertailutaulukko
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "CWDM vs DWDM -vertailu");

  const hdrOpts = { bold: true, color: TEXT_PRIMARY, fill: { color: TABLE_HEADER }, fontSize: 13, fontFace: "Arial", align: "center" };
  const r1 = (text, opts) => ({ text, options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW1 }, fontSize: 12, fontFace: "Arial", ...opts } });
  const r2 = (text, opts) => ({ text, options: { color: TEXT_PRIMARY, fill: { color: TABLE_ROW2 }, fontSize: 12, fontFace: "Arial", ...opts } });

  const tableRows = [
    [
      { text: "Ominaisuus", options: { ...hdrOpts } },
      { text: "CWDM", options: { ...hdrOpts, color: PURPLE } },
      { text: "DWDM", options: { ...hdrOpts, color: CYAN } },
    ],
    [r1("Kanavavaali"), r1("20 nm", { bold: true, color: PURPLE, align: "center" }), r1("0.8 / 0.4 nm", { bold: true, color: CYAN, align: "center" })],
    [r2("Kanavat / kuitu"), r2("18", { align: "center" }), r2("80-96", { align: "center" })],
    [r1("Laserit"), r1("DFB (jaahdyttamaton)", { align: "center" }), r1("DFB (jaahdytetty, tarkka)", { align: "center" })],
    [r2("Kantama"), r2("40-80 km", { align: "center" }), r2("Satoja km (EDFA)", { align: "center" })],
    [r1("EDFA-vahvistin"), r1("Ei tuettu", { color: RED, align: "center" }), r1("Kylla (C/L-band)", { color: GREEN, align: "center" })],
    [r2("MUX-tekniikka"), r2("TFF (passiivinen)", { align: "center" }), r2("AWG / TFF", { align: "center" })],
    [r1("Hinta / kanava"), r1("Edullinen", { color: GREEN, bold: true, align: "center" }), r1("Kallis", { color: RED, bold: true, align: "center" })],
    [r2("Kaytto"), r2("Access / metro", { align: "center" }), r2("Runko / long-haul", { align: "center" })],
  ];

  slide.addTable(tableRows, {
    x: 0.5, y: 1.2, w: 12, h: 5.5,
    border: { type: "solid", pt: 0.5, color: "374151" },
    colW: [3, 4.5, 4.5],
    rowH: [0.55, 0.55, 0.55, 0.55, 0.55, 0.55, 0.55, 0.55, 0.55],
  });
}

// =============================================
// SLIDE 8: Milloin CWDM, milloin DWDM
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Milloin CWDM, milloin DWDM?");

  // CWDM card
  addCard(slide, 0.4, 1.2, 5.8, 5.5, CARD_BG);
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 1.2, w: 5.8, h: 0.06, fill: { color: PURPLE },
  });
  slide.addText("CWDM", {
    x: 0.6, y: 1.4, w: 5.4, h: 0.6,
    fontSize: 28, fontFace: "Arial", bold: true, color: PURPLE,
  });
  slide.addText("Access & Metro", {
    x: 0.6, y: 2.0, w: 5.4, h: 0.4,
    fontSize: 16, fontFace: "Arial", color: GOLD,
  });

  const cwdmUses = [
    "Kuitupula: monta palvelua 1 kuidussa",
    "Eta-OLT: P2P-yhteydet kylaan",
    "Yrityspalvelut: erilliset aallonpituudet",
    "Rengasverkot: suojaus + kapasiteetti",
    "Metro-yhteydet: alle 80 km",
  ];
  cwdmUses.forEach((text, i) => {
    slide.addShape(pptx.shapes.OVAL, {
      x: 0.8, y: 2.65 + i * 0.7, w: 0.25, h: 0.25,
      fill: { color: PURPLE },
    });
    slide.addText(text, {
      x: 1.2, y: 2.55 + i * 0.7, w: 4.5, h: 0.5,
      fontSize: 13, fontFace: "Arial", color: TEXT_PRIMARY,
    });
  });

  // DWDM card
  addCard(slide, 6.7, 1.2, 5.8, 5.5, CARD_BG);
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 6.7, y: 1.2, w: 5.8, h: 0.06, fill: { color: CYAN },
  });
  slide.addText("DWDM", {
    x: 6.9, y: 1.4, w: 5.4, h: 0.6,
    fontSize: 28, fontFace: "Arial", bold: true, color: CYAN,
  });
  slide.addText("Backbone & Long-haul", {
    x: 6.9, y: 2.0, w: 5.4, h: 0.4,
    fontSize: 16, fontFace: "Arial", color: GOLD,
  });

  const dwdmUses = [
    "Runkoverkko: satoja km EDFA:lla",
    "Suuri kapasiteetti: 80-96 kanavaa",
    "Operaattoriyhteydet: 100G+ / kanava",
    "Merikaapelit: submarine-jarjestelmat",
    "Datakeskukset: terabit-yhteydet",
  ];
  dwdmUses.forEach((text, i) => {
    slide.addShape(pptx.shapes.OVAL, {
      x: 7.1, y: 2.65 + i * 0.7, w: 0.25, h: 0.25,
      fill: { color: CYAN },
    });
    slide.addText(text, {
      x: 7.5, y: 2.55 + i * 0.7, w: 4.5, h: 0.5,
      fontSize: 13, fontFace: "Arial", color: TEXT_PRIMARY,
    });
  });
}

// =============================================
// SLIDE 9: Section - CWDM kaytannossa
// =============================================
addSectionSlide("CWDM kaytannossa");

// =============================================
// SLIDE 10: CWDM + eta-OLT (CORRECTED)
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Kaytantoesimerkki: CWDM + eta-OLT");

  // Scenario description
  addCard(slide, 0.4, 1.1, 12.3, 1.1, CARD_BG2);
  slide.addText([
    { text: "Tilanne: ", options: { bold: true, color: GOLD } },
    { text: "Kyla 10 km paassa, runkokaapelissa 12 kuitua, tarve 128 asiakkaalle.", options: { color: TEXT_PRIMARY } },
  ], {
    x: 0.6, y: 1.2, w: 11.8, h: 0.8,
    fontSize: 15, fontFace: "Arial",
  });

  // Central office
  addCard(slide, 0.4, 2.5, 3.5, 3.0, CARD_BG);
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 2.5, w: 3.5, h: 0.5, fill: { color: "2D1B69" },
  });
  slide.addText("Keskus (CO)", {
    x: 0.4, y: 2.5, w: 3.5, h: 0.5,
    fontSize: 14, fontFace: "Arial", bold: true, color: GOLD, align: "center", valign: "middle",
  });
  slide.addText([
    { text: "Kytkin lahettaa\n", options: { color: TEXT_DIM, fontSize: 11 } },
    { text: "4 x 10G P2P\n", options: { bold: true, color: CYAN, fontSize: 14 } },
    { text: "uplink-yhteydet\n\n", options: { color: TEXT_DIM, fontSize: 11 } },
    { text: "CWDM MUX\n", options: { bold: true, color: PURPLE, fontSize: 13 } },
    { text: "yhdistaa 1 kuitupariin", options: { color: TEXT_DIM, fontSize: 11 } },
  ], {
    x: 0.6, y: 3.1, w: 3.1, h: 2.2,
    fontFace: "Arial", valign: "top",
  });

  // Arrow: CO -> Trunk
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 4.0, y: 3.9, w: 1.5, h: 0.25, fill: { color: PURPLE },
  });
  slide.addText("1 kuitu", {
    x: 4.0, y: 3.5, w: 1.5, h: 0.35,
    fontSize: 10, fontFace: "Arial", bold: true, color: PURPLE, align: "center",
  });
  slide.addText("10 km", {
    x: 4.0, y: 4.2, w: 1.5, h: 0.3,
    fontSize: 10, fontFace: "Arial", color: TEXT_DIM, align: "center",
  });

  // Village FDH
  addCard(slide, 5.6, 2.5, 3.5, 3.0, CARD_BG);
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 5.6, y: 2.5, w: 3.5, h: 0.5, fill: { color: "1B3A2F" },
  });
  slide.addText("Kyla (FDH)", {
    x: 5.6, y: 2.5, w: 3.5, h: 0.5,
    fontSize: 14, fontFace: "Arial", bold: true, color: GREEN, align: "center", valign: "middle",
  });
  slide.addText([
    { text: "CWDM DEMUX\n", options: { bold: true, color: PURPLE, fontSize: 13 } },
    { text: "erottaa 4 kanavaa\n\n", options: { color: TEXT_DIM, fontSize: 11 } },
    { text: "Eta-OLT\n", options: { bold: true, color: CYAN, fontSize: 14 } },
    { text: "4 GPON-porttia\n", options: { color: TEXT_PRIMARY, fontSize: 12 } },
    { text: "(1490/1310 nm)", options: { color: TEXT_DIM, fontSize: 11 } },
  ], {
    x: 5.8, y: 3.1, w: 3.1, h: 2.2,
    fontFace: "Arial", valign: "top",
  });

  // Arrow: FDH -> Customers
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 9.2, y: 3.9, w: 1.0, h: 0.25, fill: { color: GREEN },
  });

  // Customers
  addCard(slide, 10.3, 2.5, 2.5, 3.0, CARD_BG);
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 10.3, y: 2.5, w: 2.5, h: 0.5, fill: { color: "1A3048" },
  });
  slide.addText("Asiakkaat", {
    x: 10.3, y: 2.5, w: 2.5, h: 0.5,
    fontSize: 14, fontFace: "Arial", bold: true, color: CYAN, align: "center", valign: "middle",
  });
  slide.addText([
    { text: "4 x 1:32\n", options: { bold: true, color: GOLD, fontSize: 14 } },
    { text: "splitterit\n\n", options: { color: TEXT_DIM, fontSize: 11 } },
    { text: "= 128\n", options: { bold: true, color: GREEN, fontSize: 20 } },
    { text: "asiakasta", options: { color: TEXT_PRIMARY, fontSize: 12 } },
  ], {
    x: 10.4, y: 3.1, w: 2.2, h: 2.2,
    fontFace: "Arial", align: "center", valign: "top",
  });

  // CRITICAL WARNING
  addCard(slide, 0.4, 5.7, 12.3, 1.2, WARN_BG);
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 5.7, w: 0.12, h: 1.2, fill: { color: RED },
  });
  slide.addText([
    { text: "HUOM: ", options: { bold: true, color: RED, fontSize: 15 } },
    { text: "GPON-signaalia EI erotella CWDM-kanaville! GPON kayttaa kiinteita aallonpituuksia (DS 1490 nm, US 1310 nm). ", options: { color: TEXT_PRIMARY, fontSize: 13 } },
    { text: "CWDM kuljettaa P2P-uplinkit, PON toimii paikallisesti vakioaallonpituuksilla.", options: { bold: true, color: GOLD, fontSize: 13 } },
  ], {
    x: 0.8, y: 5.8, w: 11.6, h: 1.0,
    fontFace: "Arial", valign: "middle",
  });
}

// =============================================
// SLIDE 11: Haviobudjetti
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Haviobudjetti: kaksi erillist");

  // P2P uplink budget
  addCard(slide, 0.4, 1.2, 5.8, 4.0, CARD_BG);
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 1.2, w: 5.8, h: 0.06, fill: { color: PURPLE },
  });
  slide.addText("P2P uplink (CWDM)", {
    x: 0.6, y: 1.35, w: 5.4, h: 0.5,
    fontSize: 20, fontFace: "Arial", bold: true, color: PURPLE,
  });
  slide.addText("10 km, 10G SFP+ (~24 dB budjetti)", {
    x: 0.6, y: 1.85, w: 5.4, h: 0.4,
    fontSize: 12, fontFace: "Arial", color: TEXT_DIM,
  });

  const p2pItems = [
    { label: "CWDM MUX", value: "~2.0 dB", color: PURPLE },
    { label: "Kuitu 10 km (0.35 dB/km)", value: "~3.5 dB", color: BLUE },
    { label: "CWDM DEMUX", value: "~2.0 dB", color: PURPLE },
    { label: "Liitokset (6 kpl)", value: "~1.5 dB", color: TEXT_DIM },
  ];
  p2pItems.forEach((item, i) => {
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.7, y: 2.4 + i * 0.55, w: 5.0, h: 0.45,
      fill: { color: CARD_BG2 }, rectRadius: 0.05,
    });
    slide.addText(item.label, {
      x: 0.9, y: 2.4 + i * 0.55, w: 3, h: 0.45,
      fontSize: 13, fontFace: "Arial", color: item.color, valign: "middle",
    });
    slide.addText(item.value, {
      x: 3.9, y: 2.4 + i * 0.55, w: 1.7, h: 0.45,
      fontSize: 13, fontFace: "Arial", bold: true, color: TEXT_PRIMARY, align: "right", valign: "middle",
    });
  });

  // Total
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.7, y: 4.65, w: 5.0, h: 0.45,
    fill: { color: "2D1B69" }, rectRadius: 0.05,
  });
  slide.addText("YHTEENSA", {
    x: 0.9, y: 4.65, w: 3, h: 0.45,
    fontSize: 14, fontFace: "Arial", bold: true, color: GOLD, valign: "middle",
  });
  slide.addText("~9.0 dB", {
    x: 3.9, y: 4.65, w: 1.7, h: 0.45,
    fontSize: 14, fontFace: "Arial", bold: true, color: GREEN, align: "right", valign: "middle",
  });

  // Local PON budget
  addCard(slide, 6.7, 1.2, 5.8, 4.0, CARD_BG);
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 6.7, y: 1.2, w: 5.8, h: 0.06, fill: { color: GREEN },
  });
  slide.addText("Paikallinen PON", {
    x: 6.9, y: 1.35, w: 5.4, h: 0.5,
    fontSize: 20, fontFace: "Arial", bold: true, color: GREEN,
  });
  slide.addText("OLT -> splitter -> ONT (<1 km)", {
    x: 6.9, y: 1.85, w: 5.4, h: 0.4,
    fontSize: 12, fontFace: "Arial", color: TEXT_DIM,
  });

  const ponItems = [
    { label: "1:32 splitter", value: "~17.5 dB", color: GOLD },
    { label: "Kuitu <1 km", value: "~0.35 dB", color: BLUE },
    { label: "Liitokset (2-3 kpl)", value: "~0.75 dB", color: TEXT_DIM },
  ];
  ponItems.forEach((item, i) => {
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 7.0, y: 2.4 + i * 0.55, w: 5.0, h: 0.45,
      fill: { color: CARD_BG2 }, rectRadius: 0.05,
    });
    slide.addText(item.label, {
      x: 7.2, y: 2.4 + i * 0.55, w: 3, h: 0.45,
      fontSize: 13, fontFace: "Arial", color: item.color, valign: "middle",
    });
    slide.addText(item.value, {
      x: 10.2, y: 2.4 + i * 0.55, w: 1.7, h: 0.45,
      fontSize: 13, fontFace: "Arial", bold: true, color: TEXT_PRIMARY, align: "right", valign: "middle",
    });
  });

  // Total PON
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 7.0, y: 4.05, w: 5.0, h: 0.45,
    fill: { color: "1B3A2F" }, rectRadius: 0.05,
  });
  slide.addText("YHTEENSA", {
    x: 7.2, y: 4.05, w: 3, h: 0.45,
    fontSize: 14, fontFace: "Arial", bold: true, color: GOLD, valign: "middle",
  });
  slide.addText("~18.6 dB", {
    x: 10.2, y: 4.05, w: 1.7, h: 0.45,
    fontSize: 14, fontFace: "Arial", bold: true, color: GREEN, align: "right", valign: "middle",
  });

  slide.addText("GPON B+ budjetti: 28 dB -> reilusti varaa!", {
    x: 6.9, y: 4.6, w: 5.4, h: 0.4,
    fontSize: 12, fontFace: "Arial", bold: true, color: GREEN,
  });

  // Key insight
  addCard(slide, 0.4, 5.5, 12.3, 1.4, CARD_BG2);
  slide.addText([
    { text: "Avainhuomio: ", options: { bold: true, color: GOLD } },
    { text: "CWDM-havio (9 dB) pysyy helposti 10G SFP+ -budjetissa (~24 dB). PON-havio lasketaan erikseen lyhyelle paikallismatkalle, jossa standardi GPON B+ (28 dB) riittaa mainiosti.", options: { color: TEXT_PRIMARY } },
  ], {
    x: 0.6, y: 5.6, w: 11.8, h: 1.2,
    fontSize: 14, fontFace: "Arial", lineSpacingMultiple: 1.15,
  });
}

// =============================================
// SLIDE 12: Bidirektionaalinen CWDM
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Bidirektionaalinen CWDM (BiDi)");

  addCard(slide, 0.4, 1.2, 12.3, 2.3, CARD_BG);
  slide.addText("BiDi P2P -uplinkit", {
    x: 0.6, y: 1.3, w: 5, h: 0.5,
    fontSize: 22, fontFace: "Arial", bold: true, color: PURPLE,
  });
  slide.addText([
    { text: "BiDi-SFP kayttaa kahta eri aallonpituutta samassa kuidussa:\n", options: { color: TEXT_PRIMARY } },
    { text: "TX ja RX eri suuntiin WDM-diplex-suotimella.\n\n", options: { color: TEXT_DIM } },
    { text: "Esim. 10G BiDi: TX 1270 nm / RX 1330 nm (pari kayttaa kaaanteisesti)\n", options: { color: CYAN } },
    { text: "Tuplaa CWDM-kapasiteetin: 1 kuitu, 2 suuntaa, useita kanavia.", options: { bold: true, color: GOLD } },
  ], {
    x: 0.6, y: 1.9, w: 11.8, h: 1.4,
    fontSize: 14, fontFace: "Arial", lineSpacingMultiple: 1.15,
  });

  // Visual: BiDi on single fiber
  addCard(slide, 0.4, 3.8, 12.3, 2.0, CARD_BG2);
  slide.addText("Esimerkki: 4 BiDi-uplinkkia yhdella kuidulla", {
    x: 0.6, y: 3.9, w: 11, h: 0.5,
    fontSize: 16, fontFace: "Arial", bold: true, color: CYAN,
  });

  // Bidirectional arrows
  const bidiPairs = [
    { tx: "1270", rx: "1330", color1: PURPLE, color2: BLUE },
    { tx: "1350", rx: "1410", color1: CYAN, color2: GREEN },
    { tx: "1470", rx: "1530", color1: GOLD, color2: ORANGE },
    { tx: "1550", rx: "1610", color1: "EC4899", color2: RED },
  ];
  bidiPairs.forEach((pair, i) => {
    const y = 4.55 + i * 0.3;
    // TX arrow (right)
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 1.0, y: y, w: 4.5, h: 0.12, fill: { color: pair.color1 },
    });
    slide.addText(pair.tx + " nm ->", {
      x: 1.0, y: y - 0.15, w: 4.5, h: 0.15,
      fontSize: 8, fontFace: "Arial", color: pair.color1, align: "center",
    });
    // RX arrow (left)
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 6.0, y: y, w: 4.5, h: 0.12, fill: { color: pair.color2 },
    });
    slide.addText("<- " + pair.rx + " nm", {
      x: 6.0, y: y - 0.15, w: 4.5, h: 0.15,
      fontSize: 8, fontFace: "Arial", color: pair.color2, align: "center",
    });
  });

  // PON note
  addCard(slide, 0.4, 6.0, 12.3, 0.9, CARD_BG);
  slide.addText([
    { text: "GPON toimii aina vakioaallonpituuksilla: ", options: { color: TEXT_DIM } },
    { text: "DS 1490 nm / US 1310 nm", options: { bold: true, color: GREEN } },
    { text: " (paikallisesti OLT:n ja ONT:n valilla, ei CWDM-verkossa)", options: { color: TEXT_DIM } },
  ], {
    x: 0.6, y: 6.1, w: 11.8, h: 0.7,
    fontSize: 13, fontFace: "Arial", valign: "middle",
  });
}

// =============================================
// SLIDE 13: Kustannusvertailu
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Kustannusvertailu");

  // CWDM cost
  addCard(slide, 0.4, 1.3, 5.8, 3.0, CARD_BG);
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 1.3, w: 5.8, h: 0.06, fill: { color: GREEN },
  });
  slide.addText("CWDM-ratkaisu", {
    x: 0.6, y: 1.5, w: 5.4, h: 0.5,
    fontSize: 22, fontFace: "Arial", bold: true, color: GREEN,
  });

  const cwdmCosts = [
    { item: "4-kanavainen MUX + DEMUX", cost: "2 000 - 5 000 EUR" },
    { item: "4 x BiDi SFP+ (10G)", cost: "800 - 2 000 EUR" },
    { item: "Eta-OLT (kompakti)", cost: "3 000 - 8 000 EUR" },
  ];
  cwdmCosts.forEach((c, i) => {
    slide.addText(c.item, {
      x: 0.8, y: 2.2 + i * 0.55, w: 3.5, h: 0.45,
      fontSize: 13, fontFace: "Arial", color: TEXT_PRIMARY, valign: "middle",
    });
    slide.addText(c.cost, {
      x: 4.3, y: 2.2 + i * 0.55, w: 1.7, h: 0.45,
      fontSize: 13, fontFace: "Arial", bold: true, color: GREEN, align: "right", valign: "middle",
    });
  });

  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 3.85, w: 5.0, h: 0.04, fill: { color: "374151" },
  });
  slide.addText("YHTEENSA: ~6 000 - 15 000 EUR", {
    x: 0.8, y: 3.95, w: 5.0, h: 0.35,
    fontSize: 15, fontFace: "Arial", bold: true, color: GREEN,
  });

  // New cable cost
  addCard(slide, 6.7, 1.3, 5.8, 3.0, CARD_BG);
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 6.7, y: 1.3, w: 5.8, h: 0.06, fill: { color: RED },
  });
  slide.addText("Uusi kaapeli (vaihtoehto)", {
    x: 6.9, y: 1.5, w: 5.4, h: 0.5,
    fontSize: 22, fontFace: "Arial", bold: true, color: RED,
  });

  const cableCosts = [
    { item: "Kaivuutyot", cost: "15 000 - 40 000 EUR/km" },
    { item: "Kaapeli + asennus", cost: "5 000 - 15 000 EUR/km" },
    { item: "Luvat + suunnittelu", cost: "2 000 - 10 000 EUR/km" },
  ];
  cableCosts.forEach((c, i) => {
    slide.addText(c.item, {
      x: 7.1, y: 2.2 + i * 0.55, w: 3.5, h: 0.45,
      fontSize: 13, fontFace: "Arial", color: TEXT_PRIMARY, valign: "middle",
    });
    slide.addText(c.cost, {
      x: 10.6, y: 2.2 + i * 0.55, w: 1.7, h: 0.45,
      fontSize: 13, fontFace: "Arial", bold: true, color: RED, align: "right", valign: "middle",
    });
  });

  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 7.1, y: 3.85, w: 5.0, h: 0.04, fill: { color: "374151" },
  });
  slide.addText("10 km: ~220 000 - 650 000 EUR", {
    x: 7.1, y: 3.95, w: 5.0, h: 0.35,
    fontSize: 15, fontFace: "Arial", bold: true, color: RED,
  });

  // Savings highlight
  addCard(slide, 0.4, 4.6, 12.3, 1.5, "1B3A2F");
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 4.6, w: 0.12, h: 1.5, fill: { color: GREEN },
  });
  slide.addText([
    { text: "Saasto: ", options: { bold: true, color: GREEN, fontSize: 24 } },
    { text: "CWDM saastaa tyypillisesti ", options: { color: TEXT_PRIMARY, fontSize: 18 } },
    { text: "90-97%", options: { bold: true, color: GOLD, fontSize: 28 } },
    { text: " verrattuna uuden kaapelin rakentamiseen.", options: { color: TEXT_PRIMARY, fontSize: 18 } },
  ], {
    x: 0.8, y: 4.7, w: 11.6, h: 1.3,
    fontFace: "Arial", valign: "middle",
  });

  // When NOT worth it
  addCard(slide, 0.4, 6.3, 12.3, 0.7, CARD_BG2);
  slide.addText([
    { text: "Milloin CWDM ei kannata? ", options: { bold: true, color: GOLD } },
    { text: "Jos vapaat kuidut riittavat, suora kytkenta on yksinkertaisempi ja halvempi.", options: { color: TEXT_DIM } },
  ], {
    x: 0.6, y: 6.35, w: 11.8, h: 0.55,
    fontSize: 13, fontFace: "Arial", valign: "middle",
  });
}

// =============================================
// SLIDE 14: Paatospuu
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Paatospuu: teknologiavalinta");

  // Decision tree as cards
  // Level 1: Shared or dedicated?
  addCard(slide, 4.5, 1.2, 4, 0.8, "2D1B69");
  slide.addText("Jaettu vai dedikoitu yhteys?", {
    x: 4.5, y: 1.2, w: 4, h: 0.8,
    fontSize: 14, fontFace: "Arial", bold: true, color: GOLD, align: "center", valign: "middle",
  });

  // Level 2 left: GPON
  addCard(slide, 0.5, 2.4, 3.5, 0.8, CARD_BG);
  slide.addText("Jaettu -> GPON / XGS-PON", {
    x: 0.5, y: 2.4, w: 3.5, h: 0.8,
    fontSize: 13, fontFace: "Arial", bold: true, color: GREEN, align: "center", valign: "middle",
  });

  // Level 2 right: P2P
  addCard(slide, 9.0, 2.4, 3.5, 0.8, CARD_BG);
  slide.addText("Dedikoitu -> P2P (10G/1G)", {
    x: 9.0, y: 2.4, w: 3.5, h: 0.8,
    fontSize: 13, fontFace: "Arial", bold: true, color: CYAN, align: "center", valign: "middle",
  });

  // Connecting lines
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 2.2, y: 2.0, w: 0.06, h: 0.4, fill: { color: "4B5563" },
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 10.7, y: 2.0, w: 0.06, h: 0.4, fill: { color: "4B5563" },
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 2.2, y: 2.0, w: 8.5, h: 0.06, fill: { color: "4B5563" },
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 6.5, y: 1.95, w: 0.06, h: 0.15, fill: { color: "4B5563" },
  });

  // Level 3 from GPON: Need remote OLT?
  addCard(slide, 0.3, 3.6, 3.8, 0.8, CARD_BG2);
  slide.addText("Tarvitaanko eta-OLT?", {
    x: 0.3, y: 3.6, w: 3.8, h: 0.8,
    fontSize: 13, fontFace: "Arial", bold: true, color: GOLD, align: "center", valign: "middle",
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 2.2, y: 3.2, w: 0.06, h: 0.4, fill: { color: "4B5563" },
  });

  // Level 3 from P2P: Free fibers?
  addCard(slide, 8.8, 3.6, 3.8, 0.8, CARD_BG2);
  slide.addText("Vapaita kuituja?", {
    x: 8.8, y: 3.6, w: 3.8, h: 0.8,
    fontSize: 13, fontFace: "Arial", bold: true, color: GOLD, align: "center", valign: "middle",
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 10.7, y: 3.2, w: 0.06, h: 0.4, fill: { color: "4B5563" },
  });

  // Level 4 GPON: Yes -> CWDM uplinks
  addCard(slide, 0.3, 4.8, 3.8, 1.2, "1B3A2F");
  slide.addText([
    { text: "Kylla -> ", options: { color: GREEN, bold: true } },
    { text: "CWDM kuljettaa\nP2P-uplinkit eta-OLT:lle\nPON paikallisesti", options: { color: TEXT_PRIMARY } },
  ], {
    x: 0.3, y: 4.8, w: 3.8, h: 1.2,
    fontSize: 12, fontFace: "Arial", align: "center", valign: "middle",
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 2.2, y: 4.4, w: 0.06, h: 0.4, fill: { color: "4B5563" },
  });

  // Level 4 P2P: Yes -> direct
  addCard(slide, 8.8, 4.8, 1.8, 1.2, "1B3A2F");
  slide.addText([
    { text: "Kylla\n", options: { color: GREEN, bold: true } },
    { text: "Suora\nkytkenta", options: { color: TEXT_PRIMARY } },
  ], {
    x: 8.8, y: 4.8, w: 1.8, h: 1.2,
    fontSize: 12, fontFace: "Arial", align: "center", valign: "middle",
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 9.7, y: 4.4, w: 0.06, h: 0.4, fill: { color: "4B5563" },
  });

  // Level 4 P2P: No -> CWDM
  addCard(slide, 10.8, 4.8, 1.8, 1.2, "2D1B69");
  slide.addText([
    { text: "Ei\n", options: { color: RED, bold: true } },
    { text: "CWDM-\nkanavointi", options: { color: PURPLE, bold: true } },
  ], {
    x: 10.8, y: 4.8, w: 1.8, h: 1.2,
    fontSize: 12, fontFace: "Arial", align: "center", valign: "middle",
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 11.7, y: 4.4, w: 0.06, h: 0.4, fill: { color: "4B5563" },
  });

  // Summary bar
  addCard(slide, 0.4, 6.3, 12.3, 0.7, CARD_BG2);
  slide.addText([
    { text: "Muista: ", options: { bold: true, color: GOLD } },
    { text: "CWDM on aina P2P-uplink-ratkaisu, ei korvaa PON-splittausta. PON toimii vakioaallonpituuksilla.", options: { color: TEXT_PRIMARY } },
  ], {
    x: 0.6, y: 6.35, w: 11.8, h: 0.55,
    fontSize: 13, fontFace: "Arial", valign: "middle",
  });
}

// =============================================
// SLIDE 15: Yhteenveto
// =============================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: BG_CONTENT };
  addTitleBar(slide, "Yhteenveto");

  const takeaways = [
    {
      num: "1",
      title: "CWDM moninkertaistaa kuitukapasiteetin",
      desc: "18 kanavaa, 20 nm vali, passiivinen MUX/DEMUX. Sopii access- ja metro-verkkoon jossa kuituja rajallisesti.",
      color: PURPLE,
    },
    {
      num: "2",
      title: "CWDM kuljettaa P2P-uplinkit, ei PON-signaaleja",
      desc: "GPON kayttaa kiinteita aallonpituuksia (1490/1310 nm). CWDM-kanavilla siirretaan uplink-dataa eta-OLT:lle.",
      color: GOLD,
    },
    {
      num: "3",
      title: "Kustannustehokas vaihtoehto uudelle kaapelille",
      desc: "CWDM-laitteisto ~6 000-15 000 EUR vs. kaapelin kaivuu 220 000-650 000 EUR / 10 km. Saasto 90-97%.",
      color: GREEN,
    },
    {
      num: "4",
      title: "DWDM runkoverkkoihin, CWDM liityntaan",
      desc: "DWDM: 80-96 kanavaa, EDFA, satoja km. CWDM: edullisempi, riittaa kun kapasiteettitarve on kohtuullinen.",
      color: CYAN,
    },
  ];

  takeaways.forEach((t, i) => {
    const y = 1.2 + i * 1.5;
    addCard(slide, 0.4, y, 12.3, 1.3, CARD_BG);

    // Number circle
    slide.addShape(pptx.shapes.OVAL, {
      x: 0.7, y: y + 0.25, w: 0.7, h: 0.7,
      fill: { color: t.color },
    });
    slide.addText(t.num, {
      x: 0.7, y: y + 0.25, w: 0.7, h: 0.7,
      fontSize: 24, fontFace: "Arial", bold: true, color: "FFFFFF", align: "center", valign: "middle",
    });

    slide.addText(t.title, {
      x: 1.6, y: y + 0.1, w: 10.8, h: 0.5,
      fontSize: 18, fontFace: "Arial", bold: true, color: t.color,
    });
    slide.addText(t.desc, {
      x: 1.6, y: y + 0.6, w: 10.8, h: 0.6,
      fontSize: 13, fontFace: "Arial", color: TEXT_DIM, lineSpacingMultiple: 1.1,
    });
  });

  // Footer
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 7.2, w: "100%", h: 0.06, fill: { color: PURPLE },
  });
  slide.addText("Janne Kammonen | Keypro Oy", {
    x: 0.5, y: 7.0, w: 12, h: 0.3,
    fontSize: 11, fontFace: "Arial", color: TEXT_DIM, align: "center",
  });
}

// =============================================
// Generate
// =============================================
const outputPath = "/Users/jannekammonen/JK/jkammone/Claude/Repositories/koulutus/CWDM_DWDM.pptx";
pptx.writeFile({ fileName: outputPath })
  .then(() => console.log("OK: " + outputPath))
  .catch(err => console.error("FAIL:", err));
