export const CM_PER_INCH = 2.54;

export const SLIDE = Object.freeze({
  widthCm: 33.867,
  heightCm: 19.05,
  widthIn: 33.867 / CM_PER_INCH,
  heightIn: 19.05 / CM_PER_INCH,
  marginX: 1.2 / CM_PER_INCH,
  topBarY: 0,
  topBarH: 1.55 / CM_PER_INCH,
  level2Y: 1.55 / CM_PER_INCH,
  level2H: 1.35 / CM_PER_INCH,
  level3Y: 2.9 / CM_PER_INCH,
  level3H: 1.15 / CM_PER_INCH,
  contentY: 4.25 / CM_PER_INCH,
  contentBottom: 18.0 / CM_PER_INCH,
  footerY: 18.25 / CM_PER_INCH,
  footerH: 0.8 / CM_PER_INCH,
});

export const THEMES = Object.freeze({
  dark: Object.freeze({
    name: "dark",
    background: "0A151C",
    panel: "122832",
    panelAlt: "173640",
    primary: "1E869F",
    cyan: "7BC1CC",
    gold: "C69A4B",
    text: "F4F7F6",
    mutedText: "B9D0D2",
    rule: "28505C",
    conclusion: "ED7D31",
    conclusionText: "FFFFFF",
    evidence: "FFC000",
  }),
  light: Object.freeze({
    name: "light",
    background: "23899A",
    panel: "FFFDFC",
    panelAlt: "D8F0EF",
    primary: "1E869F",
    cyan: "7BC1CC",
    gold: "C69A4B",
    text: "15343A",
    mutedText: "41646B",
    rule: "A8D7D9",
    conclusion: "ED7D31",
    conclusionText: "FFFFFF",
    evidence: "FFC000",
  }),
});

// Backward-compatible alias. The Skill defaults to the dark written style specification.
export const COLORS = THEMES.dark;

export const FONTS = Object.freeze({
  title: "Microsoft YaHei",
  body: "SimHei",
  source: "KaiTi",
  latin: "Times New Roman",
});

export const FONT_SIZE = Object.freeze({
  titleDefault: 32,
  titleMinimum: 28,
  bodyDefault: 28,
  bodyMinimum: 24,
  sourceDefault: 26,
  sourceMinimum: 24,
  conclusionDefault: 34,
  conclusionMinimum: 32,
  metadata: 14,
});

export function cm(value) {
  return value / CM_PER_INCH;
}

export function resolveTheme(theme = "dark") {
  if (typeof theme === "object" && theme) return theme;
  return THEMES[theme] || THEMES.dark;
}

export function configureHistoryDeck(pptx, metadata = {}) {
  pptx.defineLayout({
    name: "HISTORY_CLASSROOM_WIDE",
    width: SLIDE.widthIn,
    height: SLIDE.heightIn,
  });
  pptx.layout = "HISTORY_CLASSROOM_WIDE";
  pptx.lang = "zh-CN";
  if (metadata.author) pptx.author = metadata.author;
  if (metadata.company) pptx.company = metadata.company;
  if (metadata.subject) pptx.subject = metadata.subject;
  if (metadata.title) pptx.title = metadata.title;
}

export function addClassroomFrame(pptx, slide, options) {
  const {
    slideNumber,
    level1,
    level2,
    level3 = "",
    courseName = "",
    organization = "高中历史课堂",
    theme = "dark",
  } = options;

  if (!String(level1 || "").trim()) {
    throw new Error(`Slide ${slideNumber}: level1 is required for a classroom content slide.`);
  }
  if (!String(level2 || "").trim()) {
    throw new Error(`Slide ${slideNumber}: level2 is required for a classroom content slide.`);
  }

  const palette = resolveTheme(theme);
  const prefix = `S${String(slideNumber).padStart(2, "0")}`;
  const contentW = SLIDE.widthIn - 2 * SLIDE.marginX;
  slide.background = { color: palette.background };

  // A consistent top bar, gold marker, and stepped title path rebuild the
  // Rebuild the written cloud-school visual language without a source template.
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: SLIDE.topBarY,
    w: SLIDE.widthIn,
    h: SLIDE.topBarH,
    fill: { color: palette.panel },
    line: { color: palette.panel, transparency: 100 },
    objectName: `${prefix}_TopBar`,
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: cm(0.22),
    h: SLIDE.topBarH,
    fill: { color: palette.gold },
    line: { color: palette.gold, transparency: 100 },
    objectName: `${prefix}_GoldMarker`,
  });
  slide.addText(level1, {
    x: SLIDE.marginX,
    y: SLIDE.topBarY,
    w: contentW - cm(4),
    h: SLIDE.topBarH,
    fontFace: FONTS.title,
    fontSize: FONT_SIZE.titleDefault,
    bold: true,
    color: palette.text,
    margin: 0.05,
    valign: "mid",
    breakLine: false,
    objectName: `${prefix}_L1Title`,
  });
  slide.addText(organization, {
    x: SLIDE.widthIn - cm(5.2),
    y: SLIDE.topBarY,
    w: cm(4.0),
    h: SLIDE.topBarH,
    fontFace: FONTS.title,
    fontSize: 18,
    bold: true,
    color: palette.cyan,
    align: "right",
    valign: "mid",
    margin: 0.03,
    objectName: `${prefix}_Brand`,
  });
  slide.addText(level2, {
    x: SLIDE.marginX,
    y: SLIDE.level2Y,
    w: contentW,
    h: SLIDE.level2H,
    fontFace: FONTS.title,
    fontSize: FONT_SIZE.titleDefault,
    bold: true,
    color: palette.name === "dark" ? palette.cyan : palette.panel,
    margin: 0.04,
    valign: "mid",
    breakLine: false,
    objectName: `${prefix}_L2Title`,
  });
  slide.addText(level3, {
    x: SLIDE.marginX,
    y: SLIDE.level3Y,
    w: contentW,
    h: SLIDE.level3H,
    fontFace: FONTS.title,
    fontSize: FONT_SIZE.titleMinimum,
    bold: true,
    color: palette.name === "dark" ? palette.text : palette.panel,
    margin: 0.04,
    valign: "mid",
    breakLine: false,
    objectName: `${prefix}_L3Title`,
  });
  slide.addShape(pptx.ShapeType.line, {
    x: SLIDE.marginX,
    y: cm(4.1),
    w: contentW,
    h: 0,
    line: { color: palette.gold, width: 1.5 },
    objectName: `${prefix}_HeaderRule`,
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: SLIDE.footerY,
    w: SLIDE.widthIn,
    h: SLIDE.footerH,
    fill: { color: palette.panel },
    line: { color: palette.panel, transparency: 100 },
    objectName: `${prefix}_Footer`,
  });
  slide.addText(courseName, {
    x: SLIDE.marginX,
    y: SLIDE.footerY,
    w: cm(18),
    h: SLIDE.footerH,
    fontFace: FONTS.title,
    fontSize: FONT_SIZE.metadata,
    color: palette.mutedText,
    bold: true,
    margin: 0.03,
    valign: "mid",
    objectName: `${prefix}_FooterCourse`,
  });
  slide.addText(String(slideNumber), {
    x: cm(30.6),
    y: SLIDE.footerY,
    w: cm(2.0),
    h: SLIDE.footerH,
    fontFace: FONTS.latin,
    fontSize: FONT_SIZE.metadata,
    color: palette.cyan,
    bold: true,
    align: "right",
    margin: 0.03,
    valign: "mid",
    objectName: `${prefix}_PageNumber`,
  });
}

export function addContentPanel(pptx, slide, options = {}) {
  const {
    x = SLIDE.marginX,
    y = SLIDE.contentY,
    w = SLIDE.widthIn - 2 * SLIDE.marginX,
    h = SLIDE.contentBottom - SLIDE.contentY,
    theme = "dark",
    alt = false,
    objectName = "ContentPanel",
  } = options;
  const palette = resolveTheme(theme);
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: 0.06,
    fill: { color: alt ? palette.panelAlt : palette.panel },
    line: { color: palette.rule, width: 1 },
    objectName,
  });
}

export function addCurrentTask(pptx, slide, options) {
  const {
    slideNumber,
    text,
    x = SLIDE.marginX,
    y = SLIDE.contentY,
    w = SLIDE.widthIn - 2 * SLIDE.marginX,
    h = cm(1.0),
    theme = "dark",
  } = options;
  if (!String(text || "").trim()) {
    throw new Error(`Slide ${slideNumber}: current task text is required.`);
  }
  const palette = resolveTheme(theme);
  const prefix = `S${String(slideNumber).padStart(2, "0")}`;
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w: cm(0.16),
    h,
    fill: { color: palette.gold },
    line: { color: palette.gold, transparency: 100 },
    objectName: `${prefix}_CurrentTaskMarker`,
  });
  slide.addText(text, {
    x: x + cm(0.32),
    y,
    w: w - cm(0.32),
    h,
    fontFace: FONTS.title,
    fontSize: FONT_SIZE.titleMinimum,
    bold: true,
    color: palette.text,
    margin: 0.03,
    valign: "mid",
    objectName: `${prefix}_CurrentTask`,
  });
}

export function addConclusionBar(pptx, slide, options) {
  const { slideNumber, text, yCm = 16.3, hCm = 1.7, theme = "dark" } = options;
  const palette = resolveTheme(theme);
  const prefix = `S${String(slideNumber).padStart(2, "0")}`;
  slide.addText(text, {
    x: SLIDE.marginX,
    y: cm(yCm),
    w: SLIDE.widthIn - 2 * SLIDE.marginX,
    h: cm(hCm),
    shape: pptx.ShapeType.rect,
    fill: { color: palette.conclusion },
    line: { color: palette.conclusion, transparency: 100 },
    fontFace: FONTS.title,
    fontSize: FONT_SIZE.conclusionDefault,
    bold: true,
    color: palette.conclusionText,
    align: "center",
    valign: "mid",
    margin: 0.08,
    objectName: `${prefix}_Conclusion`,
  });
}

export function addMediaPlaceholder(pptx, slide, options) {
  const {
    slideNumber,
    index = 1,
    x,
    y,
    w,
    h,
    mediaType = "外部素材",
    description,
    focus = "",
    ratio = "按占位区域",
    theme = "dark",
  } = options;
  const palette = resolveTheme(theme);
  const prefix = `S${String(slideNumber).padStart(2, "0")}`;
  const mediaId = `MEDIA-P${String(slideNumber).padStart(2, "0")}-${String(index).padStart(2, "0")}`;
  const lines = [
    `【${mediaId}｜${mediaType}占位】`,
    `建议内容：${description}`,
    `推荐比例：${ratio}`,
  ];
  if (focus) lines.push(`观察重点：${focus}`);
  lines.push("教师交付后手动替换");

  slide.addText(lines.join("\n"), {
    x, y, w, h,
    shape: pptx.ShapeType.rect,
    fill: { color: palette.panelAlt, transparency: 4 },
    line: { color: palette.gold, width: 2 },
    fontFace: FONTS.body,
    fontSize: FONT_SIZE.bodyMinimum,
    bold: true,
    color: palette.text,
    align: "center",
    valign: "mid",
    margin: 0.16,
    breakLine: false,
    objectName: `${prefix}_MEDIA_${String(index).padStart(2, "0")}`,
  });

  return mediaId;
}

export function contentBounds() {
  return {
    x: SLIDE.marginX,
    y: SLIDE.contentY,
    w: SLIDE.widthIn - 2 * SLIDE.marginX,
    h: SLIDE.contentBottom - SLIDE.contentY,
  };
}
