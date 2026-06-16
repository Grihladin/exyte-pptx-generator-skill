export const C = {
  paper: "#F7F2E8",
  paper2: "#FBF7EE",
  ink: "#151922",
  slate: "#283440",
  muted: "#65645F",
  soft: "#E9DED0",
  rule: "#D8CEC0",
  teal: "#256D85",
  rust: "#B63E2F",
  gold: "#C79A3D",
  green: "#3E6B4D",
  white: "#FFFFFF",
  none: "#00000000",
};

export function addBackground(slide, ctx, fill = C.paper) {
  ctx.addShape(slide, {
    x: 0,
    y: 0,
    w: ctx.W,
    h: ctx.H,
    fill,
    line: ctx.line(C.none, 0),
    name: "background",
  });
}

export function addKicker(slide, ctx, label, x, y, color = C.teal, onDark = false, suffix = "01") {
  ctx.addShape(slide, {
    x,
    y: y + 8,
    w: 9,
    h: 9,
    fill: color,
    line: ctx.line(C.none, 0),
    name: `kicker-${suffix}-marker`,
  });
  ctx.addText(slide, {
    text: label.toUpperCase(),
    x: x + 18,
    y,
    w: 260,
    h: 26,
    fontSize: 12,
    bold: true,
    color: onDark ? "#F2E7D5" : C.muted,
    valign: "middle",
    name: `kicker-${suffix}-label`,
  });
}

export function addClaimTitle(slide, ctx, text, x, y, w, color = C.ink, size = 42) {
  return ctx.addText(slide, {
    text,
    x,
    y,
    w,
    h: 112,
    fontSize: size,
    color,
    bold: true,
    face: ctx.fonts.title,
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
    name: "claim-title",
  });
}

export function addBody(slide, ctx, text, x, y, w, h, color = C.muted, size = 18, name = "body") {
  return ctx.addText(slide, {
    text,
    x,
    y,
    w,
    h,
    fontSize: size,
    color,
    face: ctx.fonts.body,
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
    name,
  });
}

export function addRule(slide, ctx, x, y, w, color = C.rule, h = 1, name = "rule") {
  return ctx.addShape(slide, {
    x,
    y,
    w,
    h,
    fill: color,
    line: ctx.line(C.none, 0),
    name,
  });
}

export function addFooter(slide, ctx, page, dark = false) {
  const y = 662;
  const textColor = dark ? "#CDBFAE" : "#8A8174";
  const ruleColor = dark ? "#46505C" : C.rule;
  addRule(slide, ctx, 64, y, 1032, ruleColor, 1, "footer-rule");
  ctx.addText(slide, {
    text: "Illustrative example data | Generated with artifact-tool PPTX skill",
    x: 64,
    y: y + 14,
    w: 700,
    h: 18,
    fontSize: 10,
    color: textColor,
    name: "footer-source",
  });
  ctx.addText(slide, {
    text: String(page).padStart(2, "0"),
    x: 1130,
    y: y + 12,
    w: 86,
    h: 20,
    fontSize: 12,
    bold: true,
    color: textColor,
    align: "right",
    name: "page-marker",
  });
}

export function addLabel(slide, ctx, text, x, y, w, h, color = C.muted, size = 13, name = "label") {
  return ctx.addText(slide, {
    text,
    x,
    y,
    w,
    h,
    fontSize: size,
    color,
    valign: "middle",
    insets: { left: 0, right: 0, top: 0, bottom: 0 },
    name,
  });
}
