/**
 * Stage 3 -- Compiler: Design Tokens + UI Plan -> .pptx
 *
 * A rendering engine that reads design_tokens.json and ui_plan.json,
 * dynamically calculates coordinates for each element type, and
 * generates a native .pptx file using pptxgenjs.
 *
 * Usage: node compiler.mjs <tokens_path> <plan_path> <output_path>
 */

import fs from "fs";
import path from "path";
import PptxGenJS from "pptxgenjs";

// ─── Configuration ───────────────────────────────────────────────────────────

const MARGIN = { left: 0.4, right: 0.4, top: 0.35, bottom: 0.45 };
const TITLE_AREA = { y: 0.55, h: 0.65 }; // Title strip position
const CONTENT_TOP = 1.45; // Where content starts below title

// ─── Helpers ─────────────────────────────────────────────────────────────────

function hexToColor(hex) {
  return hex ? hex.replace("#", "") : "333333";
}

function contentWidth(tokens) {
  return tokens.dimensions.width - MARGIN.left - MARGIN.right;
}

function contentHeight() {
  return 7.5 - CONTENT_TOP - MARGIN.bottom;
}

function loadImageAsBase64(imgPath, outputDir) {
  // imgPath is relative to the output dir (e.g. "assets/bg_cover.png")
  const fullPath = path.resolve(outputDir, imgPath);
  if (fs.existsSync(fullPath)) {
    const data = fs.readFileSync(fullPath);
    const ext = path.extname(fullPath).slice(1).toLowerCase();
    const mime = ext === "jpg" ? "jpeg" : ext;
    return `image/${mime};base64,${data.toString("base64")}`;
  }
  return null;
}

/** Determine which background key to use for a given layout type */
function getBgKey(layout, backgrounds) {
  const mapping = {
    cover: ["Cover", "1_Cover", "2_Cover", "0_Title Company"],
    divider: ["Divider", "C_Section blue"],
    content: ["Title only", "1_E_Title, Subtitle and Body", "Blank"],
    chart: ["Title only", "1_E_Title, Subtitle and Body", "Blank"],
    thank_you: ["1_Thank you", "Thank You", "1_Thank you"],
  };

  const candidates = mapping[layout] || mapping["content"];
  for (const key of candidates) {
    if (backgrounds[key]) return key;
  }
  // Fallback: return first available
  const keys = Object.keys(backgrounds);
  return keys.length > 0 ? keys[0] : null;
}


// ─── Slide Masters ───────────────────────────────────────────────────────────

function defineSlidesMasters(pptx, tokens, outputDir) {
  const backgrounds = tokens.backgrounds || {};

  const layoutTypes = ["cover", "divider", "content", "chart", "thank_you"];

  for (const layoutType of layoutTypes) {
    const bgKey = getBgKey(layoutType, backgrounds);
    const masterDef = {
      title: `MASTER_${layoutType.toUpperCase()}`,
    };

    if (bgKey && backgrounds[bgKey]) {
      const b64 = loadImageAsBase64(backgrounds[bgKey], outputDir);
      if (b64) {
        masterDef.background = { data: b64 };
      } else {
        masterDef.background = { color: hexToColor(tokens.colors.lt1 || "#FFFFFF") };
      }
    } else {
      masterDef.background = { color: hexToColor(tokens.colors.lt1 || "#FFFFFF") };
    }

    pptx.defineSlideMaster(masterDef);
  }
}


// ─── Add Title to Slide ──────────────────────────────────────────────────────

function addSlideTitle(slide, title, tokens) {
  if (!title) return;
  slide.addText(title.toUpperCase(), {
    x: MARGIN.left,
    y: TITLE_AREA.y,
    w: contentWidth(tokens),
    h: TITLE_AREA.h,
    fontSize: 28,
    fontFace: tokens.fonts.heading,
    color: hexToColor(tokens.colors.dk1),
    bold: true,
    valign: "middle",
  });
}


// ─── Element Renderers ───────────────────────────────────────────────────────

function renderGrid(slide, element, tokens) {
  const { columns, items } = element;
  const cols = Math.min(columns || 3, 4);
  const gap = 0.25;
  const cW = contentWidth(tokens);
  const cardW = (cW - (cols - 1) * gap) / cols;
  const rows = Math.ceil(items.length / cols);
  const rowGap = 0.25;
  const maxCardH = (contentHeight() - (rows - 1) * rowGap) / rows;
  const cardH = Math.min(maxCardH, 2.6);

  items.forEach((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const itemsInRow = row === rows - 1 && items.length % cols !== 0
      ? items.length % cols : cols;
    const rowW = itemsInRow * cardW + (itemsInRow - 1) * gap;
    const startX = MARGIN.left + (cW - rowW) / 2;
    const x = startX + col * (cardW + gap);
    const y = CONTENT_TOP + row * (cardH + rowGap);

    // Card background
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: hexToColor(tokens.colors.lt2) },
      line: { color: hexToColor(tokens.colors.accent1), width: 1.5 },
      rectRadius: 0.1,
    });

    // Heading
    slide.addText(item.heading || "", {
      x: x + 0.2, y: y + 0.2, w: cardW - 0.4, h: 0.5,
      fontSize: 16, bold: true,
      color: hexToColor(tokens.colors.accent1),
      fontFace: tokens.fonts.heading,
      valign: "top",
    });

    // Body
    slide.addText(item.body || "", {
      x: x + 0.2, y: y + 0.75, w: cardW - 0.4, h: cardH - 1.0,
      fontSize: 11,
      color: hexToColor(tokens.colors.dk2),
      fontFace: tokens.fonts.body,
      valign: "top",
      wrap: true,
    });
  });
}


function renderTimeline(slide, element, tokens) {
  const { steps } = element;
  if (!steps || steps.length === 0) return;
  const n = steps.length;
  const cW = contentWidth(tokens);
  const gap = 0.15;
  const boxW = Math.min(2.8, (cW - (n - 1) * gap - (n - 1) * 0.3) / n);
  const boxH = 2.0;
  const totalW = n * boxW + (n - 1) * (gap + 0.3);
  const startX = MARGIN.left + (cW - totalW) / 2;
  const midY = CONTENT_TOP + (contentHeight() - boxH) / 2;

  steps.forEach((step, i) => {
    const x = startX + i * (boxW + gap + 0.3);

    // Label badge
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y: midY - 0.35, w: boxW, h: 0.35,
      fill: { color: hexToColor(tokens.colors.accent1) },
      rectRadius: 0.05,
    });
    slide.addText(step.label || `${i + 1}`, {
      x, y: midY - 0.35, w: boxW, h: 0.35,
      fontSize: 11, bold: true, align: "center", valign: "middle",
      color: hexToColor(tokens.colors.lt1),
      fontFace: tokens.fonts.heading,
    });

    // Main box
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y: midY, w: boxW, h: boxH,
      fill: { color: hexToColor(tokens.colors.lt2) },
      line: { color: hexToColor(tokens.colors.accent1), width: 1 },
      rectRadius: 0.08,
    });

    // Title
    slide.addText((step.title || "").toUpperCase(), {
      x: x + 0.15, y: midY + 0.15, w: boxW - 0.3, h: 0.4,
      fontSize: 12, bold: true,
      color: hexToColor(tokens.colors.dk1),
      fontFace: tokens.fonts.heading,
      valign: "top",
    });

    // Description
    slide.addText(step.description || "", {
      x: x + 0.15, y: midY + 0.6, w: boxW - 0.3, h: boxH - 0.8,
      fontSize: 10,
      color: hexToColor(tokens.colors.dk2),
      fontFace: tokens.fonts.body,
      valign: "top", wrap: true,
    });

    // Arrow connector
    if (i < n - 1) {
      const arrowX = x + boxW + gap * 0.25;
      const arrowY = midY + boxH / 2 - 0.15;
      slide.addShape(pptx.shapes.RIGHT_ARROW, {
        x: arrowX, y: arrowY, w: 0.3, h: 0.3,
        fill: { color: hexToColor(tokens.colors.accent1) },
        line: { width: 0 },
      });
    }
  });
}


function renderHero(slide, element, tokens) {
  const cW = contentWidth(tokens);
  const cH = contentHeight();

  // Large heading
  slide.addText(element.heading || "", {
    x: MARGIN.left + 0.5, y: CONTENT_TOP + cH * 0.15,
    w: cW - 1.0, h: 1.2,
    fontSize: 36, bold: true,
    color: hexToColor(tokens.colors.dk1),
    fontFace: tokens.fonts.heading,
    align: "center", valign: "middle",
  });

  // Body text
  slide.addText(element.body || "", {
    x: MARGIN.left + 1.0, y: CONTENT_TOP + cH * 0.15 + 1.4,
    w: cW - 2.0, h: 1.5,
    fontSize: 16,
    color: hexToColor(tokens.colors.dk2),
    fontFace: tokens.fonts.body,
    align: "center", valign: "top",
    wrap: true,
  });
}


function renderBullets(slide, element, tokens) {
  const { items } = element;
  if (!items || items.length === 0) return;
  const cW = contentWidth(tokens);

  const textRows = items.map((item) => {
    const parts = [];
    if (item.bold_prefix) {
      parts.push({ text: item.bold_prefix + " ", options: { bold: true, fontSize: 13, color: hexToColor(tokens.colors.accent1) } });
    }
    parts.push({ text: item.text || "", options: { fontSize: 13, color: hexToColor(tokens.colors.dk2) } });
    return {
      text: parts,
      options: { bullet: { code: "2022" }, paraSpaceAfter: 8 },
    };
  });

  slide.addText(textRows, {
    x: MARGIN.left + 0.3, y: CONTENT_TOP + 0.2,
    w: cW - 0.6, h: contentHeight() - 0.4,
    fontFace: tokens.fonts.body,
    valign: "top", wrap: true,
  });
}


function renderChart(slide, element, tokens) {
  const chartTypeMap = {
    bar: pptx.charts.BAR,
    line: pptx.charts.LINE,
    pie: pptx.charts.PIE,
    column: pptx.charts.BAR,
    doughnut: pptx.charts.DOUGHNUT,
    area: pptx.charts.AREA,
  };

  const chartType = chartTypeMap[(element.chart_type || "bar").toLowerCase()] || pptx.charts.BAR;
  const chartData = (element.series || []).map((s) => ({
    name: s.name,
    labels: element.categories || [],
    values: s.values || [],
  }));

  if (chartData.length === 0) return;

  const cW = contentWidth(tokens);
  slide.addChart(chartType, chartData, {
    x: MARGIN.left + 1.0, y: CONTENT_TOP + 0.2,
    w: cW - 2.0, h: contentHeight() - 0.4,
    showTitle: false,
    showValue: true,
    chartColors: [
      hexToColor(tokens.colors.accent1),
      hexToColor(tokens.colors.accent2),
      hexToColor(tokens.colors.accent3),
      hexToColor(tokens.colors.accent4),
      hexToColor(tokens.colors.accent5),
      hexToColor(tokens.colors.accent6),
    ],
    catAxisLabelFontFace: tokens.fonts.body,
    valAxisLabelFontFace: tokens.fonts.body,
    dataLabelFontFace: tokens.fonts.body,
  });
}


function renderTable(slide, element, tokens) {
  const { headers, rows } = element;
  if (!headers || headers.length === 0) return;

  const tableRows = [
    headers.map((h) => ({
      text: h, options: {
        bold: true, fontSize: 11, align: "center",
        color: hexToColor(tokens.colors.lt1),
        fill: { color: hexToColor(tokens.colors.accent1) },
      }
    })),
    ...(rows || []).map((row, ri) =>
      (row.cells || []).map((cell) => ({
        text: cell, options: {
          fontSize: 10, align: "center",
          fill: { color: ri % 2 === 0 ? hexToColor(tokens.colors.lt2) : hexToColor(tokens.colors.lt1) },
        }
      }))
    ),
  ];

  const cW = contentWidth(tokens);
  slide.addTable(tableRows, {
    x: MARGIN.left + 0.3, y: CONTENT_TOP + 0.2,
    w: cW - 0.6,
    fontFace: tokens.fonts.body,
    border: { type: "solid", pt: 0.5, color: hexToColor(tokens.colors.dk2) },
    autoPage: false,
  });
}


function renderTwoColumn(slide, element, tokens) {
  const cW = contentWidth(tokens);
  const colW = (cW - 0.4) / 2;
  const cH = contentHeight();

  ["left", "right"].forEach((side, idx) => {
    const data = element[side];
    if (!data) return;
    const x = MARGIN.left + idx * (colW + 0.4);

    // Column card
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y: CONTENT_TOP, w: colW, h: cH,
      fill: { color: hexToColor(tokens.colors.lt2) },
      line: { color: hexToColor(tokens.colors.accent1), width: 1 },
      rectRadius: 0.08,
    });

    // Heading
    slide.addText((data.heading || "").toUpperCase(), {
      x: x + 0.25, y: CONTENT_TOP + 0.25, w: colW - 0.5, h: 0.5,
      fontSize: 18, bold: true,
      color: hexToColor(tokens.colors.accent1),
      fontFace: tokens.fonts.heading,
    });

    // Body
    slide.addText(data.body || "", {
      x: x + 0.25, y: CONTENT_TOP + 0.85, w: colW - 0.5, h: cH - 1.2,
      fontSize: 12,
      color: hexToColor(tokens.colors.dk2),
      fontFace: tokens.fonts.body,
      valign: "top", wrap: true,
    });
  });
}


function renderStatsRow(slide, element, tokens) {
  const { items } = element;
  if (!items || items.length === 0) return;
  const n = items.length;
  const cW = contentWidth(tokens);
  const gap = 0.3;
  const cardW = (cW - (n - 1) * gap) / n;
  const cardH = 2.2;
  const midY = CONTENT_TOP + (contentHeight() - cardH) / 2;

  items.forEach((item, i) => {
    const x = MARGIN.left + i * (cardW + gap);

    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y: midY, w: cardW, h: cardH,
      fill: { color: hexToColor(tokens.colors.lt2) },
      line: { color: hexToColor(tokens.colors.accent1), width: 1.5 },
      rectRadius: 0.1,
    });

    // Big number
    slide.addText(item.value || "", {
      x, y: midY + 0.2, w: cardW, h: 1.0,
      fontSize: 36, bold: true,
      color: hexToColor(tokens.colors.accent1),
      fontFace: tokens.fonts.heading,
      align: "center", valign: "middle",
    });

    // Label
    slide.addText(item.label || "", {
      x, y: midY + 1.3, w: cardW, h: 0.6,
      fontSize: 12,
      color: hexToColor(tokens.colors.dk2),
      fontFace: tokens.fonts.body,
      align: "center", valign: "top",
    });
  });
}


function renderQuote(slide, element, tokens) {
  const cW = contentWidth(tokens);
  const cH = contentHeight();
  const boxW = cW - 2.0;
  const boxH = 3.0;
  const x = MARGIN.left + 1.0;
  const y = CONTENT_TOP + (cH - boxH) / 2;

  // Quote background
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w: boxW, h: boxH,
    fill: { color: hexToColor(tokens.colors.lt2) },
    line: { color: hexToColor(tokens.colors.accent1), width: 2 },
    rectRadius: 0.1,
  });

  // Opening quotation mark
  slide.addText("\u201C", {
    x: x + 0.2, y: y + 0.1, w: 0.8, h: 0.8,
    fontSize: 60, color: hexToColor(tokens.colors.accent1),
    fontFace: tokens.fonts.heading, bold: true,
  });

  // Quote text
  slide.addText(element.quote || "", {
    x: x + 0.4, y: y + 0.6, w: boxW - 0.8, h: boxH - 1.6,
    fontSize: 16, italic: true,
    color: hexToColor(tokens.colors.dk1),
    fontFace: tokens.fonts.body,
    valign: "middle", align: "center", wrap: true,
  });

  // Attribution
  slide.addText(`-- ${element.attribution || ""}`, {
    x: x + 0.4, y: y + boxH - 0.9, w: boxW - 0.8, h: 0.5,
    fontSize: 12, italic: true,
    color: hexToColor(tokens.colors.dk2),
    fontFace: tokens.fonts.body,
    align: "right",
  });
}


function renderComparison(slide, element, tokens) {
  const cW = contentWidth(tokens);
  const colW = (cW - 0.5) / 2;
  const cH = contentHeight();

  ["left", "right"].forEach((side, idx) => {
    const data = element[side];
    if (!data) return;
    const x = MARGIN.left + idx * (colW + 0.5);
    const accent = idx === 0 ? tokens.colors.accent1 : tokens.colors.accent3 || tokens.colors.accent2;

    // Header bar
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y: CONTENT_TOP, w: colW, h: 0.5,
      fill: { color: hexToColor(accent) },
      rectRadius: 0.06,
    });
    slide.addText((data.title || "").toUpperCase(), {
      x, y: CONTENT_TOP, w: colW, h: 0.5,
      fontSize: 14, bold: true, align: "center", valign: "middle",
      color: hexToColor(tokens.colors.lt1),
      fontFace: tokens.fonts.heading,
    });

    // Points
    const pointTexts = (data.points || []).map((p) => ({
      text: p,
      options: { bullet: { code: "2022" }, paraSpaceAfter: 6, fontSize: 11, color: hexToColor(tokens.colors.dk2) },
    }));

    slide.addText(pointTexts, {
      x: x + 0.15, y: CONTENT_TOP + 0.6, w: colW - 0.3, h: cH - 0.8,
      fontFace: tokens.fonts.body, valign: "top", wrap: true,
    });
  });
}


function renderIconGrid(slide, element, tokens) {
  const { columns, items } = element;
  const cols = Math.min(columns || 3, 4);
  const gap = 0.25;
  const cW = contentWidth(tokens);
  const cardW = (cW - (cols - 1) * gap) / cols;
  const rows = Math.ceil(items.length / cols);
  const rowGap = 0.25;
  const cardH = Math.min((contentHeight() - (rows - 1) * rowGap) / rows, 2.4);

  items.forEach((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = MARGIN.left + col * (cardW + gap);
    const y = CONTENT_TOP + row * (cardH + rowGap);

    // Card
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: hexToColor(tokens.colors.lt2) },
      line: { color: hexToColor(tokens.colors.accent2), width: 1 },
      rectRadius: 0.08,
    });

    // Icon (emoji)
    slide.addText(item.icon || "", {
      x: x + 0.15, y: y + 0.15, w: 0.6, h: 0.6,
      fontSize: 28,
    });

    // Title
    slide.addText(item.title || "", {
      x: x + 0.15, y: y + 0.8, w: cardW - 0.3, h: 0.4,
      fontSize: 13, bold: true,
      color: hexToColor(tokens.colors.dk1),
      fontFace: tokens.fonts.heading,
    });

    // Description
    slide.addText(item.description || "", {
      x: x + 0.15, y: y + 1.2, w: cardW - 0.3, h: cardH - 1.5,
      fontSize: 10,
      color: hexToColor(tokens.colors.dk2),
      fontFace: tokens.fonts.body,
      valign: "top", wrap: true,
    });
  });
}


function renderFunnel(slide, element, tokens) {
  const { steps } = element;
  if (!steps || steps.length === 0) return;
  const n = steps.length;
  const cW = contentWidth(tokens);
  const cH = contentHeight();
  const stepH = Math.min((cH - 0.3) / n, 1.0);
  const maxW = cW - 1.0;
  const minW = maxW * 0.3;

  steps.forEach((step, i) => {
    const ratio = 1 - (i / Math.max(n - 1, 1)) * 0.7;
    const w = minW + (maxW - minW) * ratio;
    const x = MARGIN.left + (cW - w) / 2;
    const y = CONTENT_TOP + i * stepH + 0.1;

    const colors = [tokens.colors.accent1, tokens.colors.accent2, tokens.colors.accent3, tokens.colors.accent4, tokens.colors.accent5];
    const fillColor = colors[i % colors.length];

    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y, w, h: stepH - 0.08,
      fill: { color: hexToColor(fillColor) },
      rectRadius: 0.06,
    });

    const labelText = step.value ? `${step.label} (${step.value})` : step.label;
    slide.addText(labelText, {
      x, y, w, h: stepH - 0.08,
      fontSize: 13, bold: true, align: "center", valign: "middle",
      color: hexToColor(tokens.colors.lt1),
      fontFace: tokens.fonts.heading,
    });
  });
}


function renderPyramid(slide, element, tokens) {
  const { levels } = element;
  if (!levels || levels.length === 0) return;
  const n = levels.length;
  const cW = contentWidth(tokens);
  const cH = contentHeight();
  const stepH = Math.min((cH - 0.3) / n, 1.0);
  const maxW = cW - 1.0;
  const minW = maxW * 0.25;

  levels.forEach((level, i) => {
    const ratio = i / Math.max(n - 1, 1);
    const w = minW + (maxW - minW) * ratio;
    const x = MARGIN.left + (cW - w) / 2;
    const y = CONTENT_TOP + i * stepH + 0.1;

    const colors = [tokens.colors.accent1, tokens.colors.accent2, tokens.colors.accent3, tokens.colors.accent4];
    const fillColor = colors[i % colors.length];

    slide.addShape(pptx.shapes.TRAPEZOID, {
      x, y, w, h: stepH - 0.08,
      fill: { color: hexToColor(fillColor) },
    });

    const txt = level.description ? `${level.label}: ${level.description}` : level.label;
    slide.addText(txt, {
      x, y, w, h: stepH - 0.08,
      fontSize: 12, bold: true, align: "center", valign: "middle",
      color: hexToColor(tokens.colors.lt1),
      fontFace: tokens.fonts.heading,
    });
  });
}


function renderMatrix(slide, element, tokens) {
  const { quadrants, x_axis, y_axis } = element;
  if (!quadrants || quadrants.length < 4) return;
  const cW = contentWidth(tokens);
  const cH = contentHeight();
  const gapLabel = 0.5;
  const matW = (cW - gapLabel - 0.5) / 2;
  const matH = (cH - gapLabel - 0.3) / 2;

  const colors = [tokens.colors.accent1, tokens.colors.accent3, tokens.colors.accent2, tokens.colors.accent4];
  const positions = [
    { col: 0, row: 0 }, // TL
    { col: 1, row: 0 }, // TR
    { col: 0, row: 1 }, // BL
    { col: 1, row: 1 }, // BR
  ];

  // Axis labels
  if (x_axis) {
    slide.addText(x_axis, {
      x: MARGIN.left + gapLabel, y: CONTENT_TOP + cH - 0.35,
      w: cW - gapLabel, h: 0.35,
      fontSize: 10, align: "center",
      color: hexToColor(tokens.colors.dk2),
      fontFace: tokens.fonts.body,
    });
  }
  if (y_axis) {
    slide.addText(y_axis, {
      x: MARGIN.left, y: CONTENT_TOP,
      w: gapLabel, h: cH - gapLabel,
      fontSize: 10, align: "center", valign: "middle",
      color: hexToColor(tokens.colors.dk2),
      fontFace: tokens.fonts.body,
      rotate: 270,
    });
  }

  quadrants.forEach((q, idx) => {
    const pos = positions[idx];
    const x = MARGIN.left + gapLabel + pos.col * (matW + 0.15);
    const y = CONTENT_TOP + pos.row * (matH + 0.15);

    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y, w: matW, h: matH,
      fill: { color: hexToColor(colors[idx]) + "33" },
      line: { color: hexToColor(colors[idx]), width: 1.5 },
      rectRadius: 0.08,
    });

    // Label
    slide.addText((q.label || "").toUpperCase(), {
      x: x + 0.15, y: y + 0.1, w: matW - 0.3, h: 0.4,
      fontSize: 12, bold: true,
      color: hexToColor(colors[idx]),
      fontFace: tokens.fonts.heading,
    });

    // Items
    const itemText = (q.items || []).map((it) => ({
      text: it,
      options: { bullet: { code: "2022" }, fontSize: 9, paraSpaceAfter: 3, color: hexToColor(tokens.colors.dk2) },
    }));
    if (itemText.length > 0) {
      slide.addText(itemText, {
        x: x + 0.15, y: y + 0.5, w: matW - 0.3, h: matH - 0.7,
        fontFace: tokens.fonts.body, valign: "top", wrap: true,
      });
    }
  });
}


function renderSWOT(slide, element, tokens) {
  // Transform SWOT into matrix format
  const matrixElement = {
    x_axis: "",
    y_axis: "",
    quadrants: [
      { label: "Strengths", items: element.strengths || [] },
      { label: "Weaknesses", items: element.weaknesses || [] },
      { label: "Opportunities", items: element.opportunities || [] },
      { label: "Threats", items: element.threats || [] },
    ],
  };
  renderMatrix(slide, matrixElement, tokens);
}


function renderCycle(slide, element, tokens) {
  const { steps } = element;
  if (!steps || steps.length === 0) return;
  const n = steps.length;
  const cW = contentWidth(tokens);
  const cH = contentHeight();
  const centerX = MARGIN.left + cW / 2;
  const centerY = CONTENT_TOP + cH / 2;
  const radius = Math.min(cW, cH) / 2 - 0.8;
  const nodeW = 2.0;
  const nodeH = 0.9;

  steps.forEach((step, i) => {
    const angle = (2 * Math.PI * i) / n - Math.PI / 2;
    const x = centerX + radius * Math.cos(angle) - nodeW / 2;
    const y = centerY + radius * Math.sin(angle) - nodeH / 2;

    const colors = [tokens.colors.accent1, tokens.colors.accent2, tokens.colors.accent3, tokens.colors.accent4, tokens.colors.accent5];

    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y, w: nodeW, h: nodeH,
      fill: { color: hexToColor(colors[i % colors.length]) },
      rectRadius: 0.1,
    });

    const txt = step.description ? `${step.title}\n${step.description}` : step.title;
    slide.addText(txt, {
      x, y, w: nodeW, h: nodeH,
      fontSize: 10, bold: true, align: "center", valign: "middle",
      color: hexToColor(tokens.colors.lt1),
      fontFace: tokens.fonts.heading,
      wrap: true,
    });

    // Arrow to next
    if (n > 1) {
      const nextAngle = (2 * Math.PI * ((i + 1) % n)) / n - Math.PI / 2;
      const midAngle = (angle + nextAngle) / 2 + (nextAngle < angle ? Math.PI : 0);
      const arrowR = radius * 0.75;
      const ax = centerX + arrowR * Math.cos(midAngle) - 0.12;
      const ay = centerY + arrowR * Math.sin(midAngle) - 0.12;
      slide.addText("\u27A1", {
        x: ax, y: ay, w: 0.3, h: 0.3,
        fontSize: 14, align: "center", valign: "middle",
        color: hexToColor(tokens.colors.accent1),
        rotate: (midAngle * 180) / Math.PI + 90,
      });
    }
  });
}


function renderGauge(slide, element, tokens) {
  const cW = contentWidth(tokens);
  const cH = contentHeight();
  const value = Math.min(Math.max(element.value || 0, 0), 100);
  const centerX = MARGIN.left + cW / 2;
  const centerY = CONTENT_TOP + cH * 0.55;
  const radius = 1.8;

  // Background arc (full circle as placeholder)
  slide.addShape(pptx.shapes.ARC, {
    x: centerX - radius, y: centerY - radius,
    w: radius * 2, h: radius * 2,
    fill: { color: hexToColor(tokens.colors.lt2) },
    line: { color: hexToColor(tokens.colors.accent1), width: 8 },
    angleRange: [180, 360],
  });

  // Value text
  const displayValue = `${element.value}${element.unit || "%"}`;
  slide.addText(displayValue, {
    x: centerX - 1.5, y: centerY - 0.5, w: 3.0, h: 1.0,
    fontSize: 48, bold: true, align: "center", valign: "middle",
    color: hexToColor(tokens.colors.accent1),
    fontFace: tokens.fonts.heading,
  });

  // Label
  slide.addText(element.label || "", {
    x: centerX - 2, y: centerY + 0.7, w: 4, h: 0.5,
    fontSize: 16, align: "center",
    color: hexToColor(tokens.colors.dk2),
    fontFace: tokens.fonts.body,
  });
}


function renderKPICards(slide, element, tokens) {
  const { items } = element;
  if (!items || items.length === 0) return;
  const n = items.length;
  const cW = contentWidth(tokens);
  const gap = 0.25;
  const cardW = (cW - (n - 1) * gap) / n;
  const cardH = 2.5;
  const midY = CONTENT_TOP + (contentHeight() - cardH) / 2;

  items.forEach((item, i) => {
    const x = MARGIN.left + i * (cardW + gap);

    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y: midY, w: cardW, h: cardH,
      fill: { color: hexToColor(tokens.colors.lt2) },
      line: { color: hexToColor(tokens.colors.accent1), width: 1.5 },
      rectRadius: 0.1,
    });

    // Value
    slide.addText(item.value || "", {
      x, y: midY + 0.2, w: cardW, h: 0.9,
      fontSize: 32, bold: true, align: "center", valign: "middle",
      color: hexToColor(tokens.colors.accent1),
      fontFace: tokens.fonts.heading,
    });

    // Label
    slide.addText(item.label || "", {
      x, y: midY + 1.1, w: cardW, h: 0.5,
      fontSize: 12, align: "center",
      color: hexToColor(tokens.colors.dk2),
      fontFace: tokens.fonts.body,
    });

    // Trend arrow + change
    if (item.change) {
      const arrow = item.trend === "up" ? "\u25B2" : item.trend === "down" ? "\u25BC" : "\u25C6";
      const trendColor = item.trend === "up" ? (tokens.colors.accent3 || tokens.colors.accent6) : item.trend === "down" ? tokens.colors.accent1 : tokens.colors.dk2;
      slide.addText(`${arrow} ${item.change}`, {
        x, y: midY + 1.6, w: cardW, h: 0.5,
        fontSize: 14, bold: true, align: "center",
        color: hexToColor(trendColor),
        fontFace: tokens.fonts.body,
      });
    }
  });
}


function renderImageText(slide, element, tokens) {
  // Renders as two-column with a placeholder rectangle on image side
  const cW = contentWidth(tokens);
  const colW = (cW - 0.4) / 2;
  const cH = contentHeight();
  const content = element.content || {};
  const imgLeft = content.image_side === "left";

  const textX = MARGIN.left + (imgLeft ? colW + 0.4 : 0);
  const imgX = MARGIN.left + (imgLeft ? 0 : colW + 0.4);

  // Image placeholder
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: imgX, y: CONTENT_TOP, w: colW, h: cH,
    fill: { color: hexToColor(tokens.colors.lt2) },
    line: { color: hexToColor(tokens.colors.accent2), width: 1 },
    rectRadius: 0.08,
  });
  slide.addText("[Image]", {
    x: imgX, y: CONTENT_TOP, w: colW, h: cH,
    fontSize: 18, align: "center", valign: "middle",
    color: hexToColor(tokens.colors.dk2) + "66",
    fontFace: tokens.fonts.body,
  });

  // Text side
  slide.addText((content.heading || "").toUpperCase(), {
    x: textX + 0.2, y: CONTENT_TOP + 0.3, w: colW - 0.4, h: 0.6,
    fontSize: 20, bold: true,
    color: hexToColor(tokens.colors.accent1),
    fontFace: tokens.fonts.heading,
  });
  slide.addText(content.body || "", {
    x: textX + 0.2, y: CONTENT_TOP + 1.0, w: colW - 0.4, h: cH - 1.5,
    fontSize: 12,
    color: hexToColor(tokens.colors.dk2),
    fontFace: tokens.fonts.body,
    valign: "top", wrap: true,
  });
}


function renderWaterfall(slide, element, tokens) {
  // Render waterfall as a bar chart with manual bars
  const { steps } = element;
  if (!steps || steps.length === 0) return;
  const n = steps.length;
  const cW = contentWidth(tokens);
  const cH = contentHeight();
  const gap = 0.15;
  const barW = (cW - (n - 1) * gap) / n;
  const maxVal = Math.max(...steps.map(s => Math.abs(s.value)));
  const scale = (cH - 1.0) / maxVal;

  let cumulative = 0;
  steps.forEach((step, i) => {
    const x = MARGIN.left + i * (barW + gap);
    const val = step.value;
    const barH = Math.abs(val) * scale * 0.4;
    const isPositive = val >= 0;
    const baseY = CONTENT_TOP + cH - 0.8;
    const y = step.is_total
      ? baseY - barH
      : isPositive
        ? baseY - (cumulative + val) * scale * 0.4
        : baseY - cumulative * scale * 0.4;

    const fillColor = step.is_total ? tokens.colors.accent1 : isPositive ? tokens.colors.accent3 || tokens.colors.accent6 : tokens.colors.accent1;

    slide.addShape(pptx.shapes.RECTANGLE, {
      x, y, w: barW, h: Math.max(barH, 0.2),
      fill: { color: hexToColor(fillColor) },
    });

    // Label
    slide.addText(step.label, {
      x, y: baseY + 0.05, w: barW, h: 0.4,
      fontSize: 9, align: "center",
      color: hexToColor(tokens.colors.dk2),
      fontFace: tokens.fonts.body,
    });

    // Value
    slide.addText(String(val), {
      x, y: y - 0.3, w: barW, h: 0.3,
      fontSize: 9, bold: true, align: "center",
      color: hexToColor(tokens.colors.dk1),
      fontFace: tokens.fonts.body,
    });

    if (!step.is_total) cumulative += val;
  });
}


// ─── Dispatch ────────────────────────────────────────────────────────────────

const RENDERERS = {
  grid: renderGrid,
  timeline: renderTimeline,
  hero: renderHero,
  bullets: renderBullets,
  chart: renderChart,
  table: renderTable,
  two_column: renderTwoColumn,
  stats_row: renderStatsRow,
  quote: renderQuote,
  comparison: renderComparison,
  icon_grid: renderIconGrid,
  waterfall: renderWaterfall,
  funnel: renderFunnel,
  pyramid: renderPyramid,
  matrix: renderMatrix,
  swot: renderSWOT,
  cycle: renderCycle,
  gauge: renderGauge,
  kpi_cards: renderKPICards,
  image_text: renderImageText,
};


// ─── Main Compile Function ──────────────────────────────────────────────────

let pptx; // Module-level for shape/chart references

function compile(tokensPath, planPath, outputPath) {
  const tokens = JSON.parse(fs.readFileSync(tokensPath, "utf-8"));
  const plan = JSON.parse(fs.readFileSync(planPath, "utf-8"));
  const outputDir = path.dirname(tokensPath);

  pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

  // Define slide masters with composite backgrounds
  defineSlidesMasters(pptx, tokens, outputDir);

  console.log(`Compiling ${plan.slides.length} slides...`);

  plan.slides.forEach((slideDef, idx) => {
    const masterName = `MASTER_${(slideDef.layout || "content").toUpperCase()}`;
    const slide = pptx.addSlide({ masterName });

    console.log(`  Slide ${idx + 1}: [${slideDef.layout}] ${slideDef.title}`);

    // Cover slide
    if (slideDef.layout === "cover") {
      slide.addText(slideDef.title || "", {
        x: 0.4, y: 2.8, w: 7.0, h: 1.2,
        fontSize: 40, bold: true,
        color: hexToColor(tokens.colors.dk1),
        fontFace: tokens.fonts.heading,
      });
      if (slideDef.subtitle) {
        slide.addText(slideDef.subtitle, {
          x: 0.4, y: 4.1, w: 7.0, h: 0.8,
          fontSize: 20,
          color: hexToColor(tokens.colors.dk2),
          fontFace: tokens.fonts.body,
        });
      }
      return;
    }

    // Thank you slide — just the background
    if (slideDef.layout === "thank_you") {
      return;
    }

    // Divider slide
    if (slideDef.layout === "divider") {
      slide.addText(slideDef.title || "", {
        x: 0.4, y: 3.0, w: 11.0, h: 1.0,
        fontSize: 32, bold: true,
        color: hexToColor(tokens.colors.dk1),
        fontFace: tokens.fonts.heading,
      });
      return;
    }

    // Content / Chart slides — add title + render elements
    addSlideTitle(slide, slideDef.title, tokens);

    (slideDef.elements || []).forEach((element) => {
      const renderer = RENDERERS[element.type];
      if (renderer) {
        renderer(slide, element, tokens);
      } else {
        console.log(`    WARNING: Unknown element type '${element.type}', skipping.`);
      }
    });
  });

  // Save
  const outDir = path.dirname(outputPath);
  if (outDir && !fs.existsSync(outDir)) {
    fs.mkdirSync(outDir, { recursive: true });
  }

  pptx.writeFile({ fileName: outputPath })
    .then(() => {
      console.log(`\nPresentation saved to: ${outputPath}`);
    })
    .catch((err) => {
      console.error(`Error saving presentation: ${err}`);
      process.exit(1);
    });
}


// ─── CLI ─────────────────────────────────────────────────────────────────────

const args = process.argv.slice(2);
if (args.length < 3) {
  console.error("Usage: node compiler.mjs <tokens_path> <plan_path> <output_path>");
  process.exit(1);
}

compile(args[0], args[1], args[2]);
