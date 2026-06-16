import { C, addBackground, addBody, addClaimTitle, addFooter, addKicker, addLabel, addRule } from "./style.mjs";

function stage(slide, ctx, { number, title, detail, proof, color }, x, y) {
  ctx.addShape(slide, {
    x,
    y,
    w: 330,
    h: 304,
    fill: C.white,
    line: ctx.line(C.rule, 1),
    name: `stage-${number}-box`,
  });
  ctx.addShape(slide, {
    x: x + 26,
    y: y + 28,
    w: 44,
    h: 44,
    geometry: "ellipse",
    fill: color,
    line: ctx.line(C.none, 0),
    name: `stage-${number}-marker`,
  });
  ctx.addText(slide, {
    text: number,
    x: x + 26,
    y: y + 31,
    w: 44,
    h: 38,
    fontSize: 18,
    bold: true,
    color: C.white,
    align: "center",
    valign: "middle",
    name: `stage-${number}-number`,
  });
  ctx.addText(slide, {
    text: title,
    x: x + 88,
    y: y + 28,
    w: 205,
    h: 54,
    fontSize: 22,
    bold: true,
    color: C.ink,
    face: ctx.fonts.title,
    name: `stage-${number}-title`,
  });
  addRule(slide, ctx, x + 28, y + 104, 274, C.soft, 1, `stage-${number}-rule`);
  addBody(slide, ctx, detail, x + 30, y + 126, 270, 76, C.muted, 16, `stage-${number}-detail`);
  ctx.addShape(slide, {
    x: x + 28,
    y: y + 226,
    w: 274,
    h: 48,
    fill: "#F7F2E8",
    line: ctx.line(C.none, 0),
    name: `stage-${number}-proof-box`,
  });
  addLabel(slide, ctx, proof, x + 44, y + 235, 238, 28, C.ink, 13, `stage-${number}-proof`);
}

export async function slide02(presentation, ctx) {
  const slide = presentation.slides.add();
  addBackground(slide, ctx, C.paper2);

  addKicker(slide, ctx, "operating loop", 64, 50, C.teal, false, "01");
  addClaimTitle(
    slide,
    ctx,
    "The work moves from prompt to proof through three controlled loops.",
    64,
    92,
    790,
    C.ink,
    42,
  );
  addBody(
    slide,
    ctx,
    "Each loop produces an explicit artifact, which keeps the final deck editable and makes visual QA repeatable.",
    894,
    105,
    280,
    74,
    C.muted,
    17,
    "workflow-support-note",
  );

  const y = 226;
  stage(
    slide,
    ctx,
    {
      number: "1",
      title: "Story lock",
      detail: "Claims, omissions, and proof objects are written before drawing any slide.",
      proof: "Output: claim spine",
      color: C.rust,
    },
    78,
    y,
  );
  stage(
    slide,
    ctx,
    {
      number: "2",
      title: "Editable build",
      detail: "Artifact-tool modules create native text, shapes, and chart primitives.",
      proof: "Output: slide modules",
      color: C.teal,
    },
    475,
    y,
  );
  stage(
    slide,
    ctx,
    {
      number: "3",
      title: "Render QA",
      detail: "PNG previews, layout checks, and package checks catch export defects.",
      proof: "Output: final PPTX",
      color: C.green,
    },
    872,
    y,
  );

  addRule(slide, ctx, 408, y + 150, 67, C.teal, 3, "connector-1-2");
  addRule(slide, ctx, 805, y + 150, 67, C.teal, 3, "connector-2-3");
  ctx.addShape(slide, {
    x: 64,
    y: 582,
    w: 1090,
    h: 46,
    fill: C.ink,
    line: ctx.line(C.none, 0),
    name: "outcome-band",
  });
  ctx.addText(slide, {
    text: "Operating rule: do not move to export until the rendered contact sheet proves rhythm, readability, and package integrity.",
    x: 84,
    y: 594,
    w: 1020,
    h: 22,
    fontSize: 15,
    color: "#F2E7D5",
    name: "outcome-band-text",
  });

  addFooter(slide, ctx, 2, false);
  return slide;
}
