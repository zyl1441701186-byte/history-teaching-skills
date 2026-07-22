export const CM_PER_INCH = 2.54;

export const SLIDE = Object.freeze({
  widthCm: 33.867,
  heightCm: 19.05,
  widthIn: 33.867 / CM_PER_INCH,
  heightIn: 19.05 / CM_PER_INCH,
  marginX: 1.2 / CM_PER_INCH,
  topBarY: 0,
  topBarH: 1.45 / CM_PER_INCH,
  level2Y: 1.45 / CM_PER_INCH,
  level2H: 1.4 / CM_PER_INCH,
  level3Y: 2.85 / CM_PER_INCH,
  level3H: 1.25 / CM_PER_INCH,
  contentY: 4.25 / CM_PER_INCH,
  contentBottom: 18.0 / CM_PER_INCH,
  footerY: 18.25 / CM_PER_INCH,
  footerH: 0.8 / CM_PER_INCH,
});

export const COLORS = Object.freeze({
  primary: "1E869F",
  deepBlue: "14525E",
  lightBlue: "DCEFF2",
  background: "F5FAFB",
  white: "FFFFFF",
  body: "111111",
  conclusion: "ED7D31",
  evidence: "FFC000",
});

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
    organization = "",
  } = options;

  const prefix = `S${String(slideNumber).padStart(2, "0")}`;
  const contentW = SLIDE.widthIn - 2 * SLIDE.marginX;
  slide.background = { color: COLORS.background };

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: SLIDE.topBarY,
    w: SLIDE.widthIn,
    h: SLIDE.topBarH,
    fill: { color: COLORS.primary },
    line: { color: COLORS.primary, transparency: 100 },
    objectName: `${prefix}_TopBar`,
  });

  slide.addText(level1, {
    x: SLIDE.marginX,
    y: SLIDE.topBarY,
    w: contentW,
    h: SLIDE.topBarH,
    fontFace: FONTS.title,
    fontSize: FONT_SIZE.titleDefault,
    bold: true,
    color: COLORS.white,
    margin: 0.05,
    valign: "mid",
    breakLine: false,
    objectName: `${prefix}_L1Title`,
  });

  slide.addText(level2, {
    x: SLIDE.marginX,
    y: SLIDE.level2Y,
    w: contentW,
    h: SLIDE.level2H,
    fontFace: FONTS.title,
    fontSize: FONT_SIZE.titleDefault,
    bold: true,
    color: COLORS.primary,
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
    color: COLORS.deepBlue,
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
    line: { color: COLORS.lightBlue, width: 1.25 },
    objectName: `${prefix}_HeaderRule`,
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: SLIDE.footerY,
    w: SLIDE.widthIn,
    h: SLIDE.footerH,
    fill: { color: COLORS.primary },
    line: { color: COLORS.primary, transparency: 100 },
    objectName: `${prefix}_Footer`,
  });

  slide.addText(courseName, {
    x: SLIDE.marginX,
    y: SLIDE.footerY,
    w: cm(18),
    h: SLIDE.footerH,
    fontFace: FONTS.title,
    fontSize: FONT_SIZE.metadata,
    color: COLORS.white,
    bold: true,
    margin: 0.03,
    valign: "mid",
    objectName: `${prefix}_FooterCourse`,
  });

  slide.addText(organization, {
    x: cm(20),
    y: SLIDE.footerY,
    w: cm(10),
    h: SLIDE.footerH,
    fontFace: FONTS.title,
    fontSize: FONT_SIZE.metadata,
    color: COLORS.white,
    align: "right",
    margin: 0.03,
    valign: "mid",
    objectName: `${prefix}_FooterOrg`,
  });

  slide.addText(String(slideNumber), {
    x: cm(30.6),
    y: SLIDE.footerY,
    w: cm(2.0),
    h: SLIDE.footerH,
    fontFace: FONTS.latin,
    fontSize: FONT_SIZE.metadata,
    color: COLORS.white,
    bold: true,
    align: "right",
    margin: 0.03,
    valign: "mid",
    objectName: `${prefix}_PageNumber`,
  });
}

export function addConclusionBar(pptx, slide, options) {
  const { slideNumber, text, yCm = 16.3, hCm = 1.7 } = options;
  const prefix = `S${String(slideNumber).padStart(2, "0")}`;
  slide.addText(text, {
    x: SLIDE.marginX,
    y: cm(yCm),
    w: SLIDE.widthIn - 2 * SLIDE.marginX,
    h: cm(hCm),
    shape: pptx.ShapeType.rect,
    fill: { color: COLORS.conclusion },
    line: { color: COLORS.conclusion, transparency: 100 },
    fontFace: FONTS.title,
    fontSize: FONT_SIZE.conclusionDefault,
    bold: true,
    color: COLORS.white,
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
  } = options;
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
    x,
    y,
    w,
    h,
    shape: pptx.ShapeType.rect,
    fill: { color: COLORS.lightBlue, transparency: 18 },
    line: { color: COLORS.primary, width: 2 },
    fontFace: FONTS.body,
    fontSize: FONT_SIZE.bodyMinimum,
    bold: true,
    color: COLORS.deepBlue,
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
