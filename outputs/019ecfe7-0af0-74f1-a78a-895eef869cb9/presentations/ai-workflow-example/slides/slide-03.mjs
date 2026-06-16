import { C, addBackground, addBody, addClaimTitle, addFooter, addKicker, addLabel, addRule } from "./style.mjs";

function pairedBar(slide, ctx, { label, before, after }, y, index) {
  const x0 = 576;
  const maxW = 424;
  const beforeW = (before / 5) * maxW;
  const afterW = (after / 5) * maxW;
  addLabel(slide, ctx, label, 420, y - 4, 132, 24, "#F2E7D5", 14, `bar-${index}-label`);
  ctx.addShape(slide, {
    x: x0,
    y,
    w: maxW,
    h: 14,
    fill: "#3D4650",
    line: ctx.line(C.none, 0),
    name: `bar-${index}-track-before`,
  });
  ctx.addShape(slide, {
    x: x0,
    y,
    w: beforeW,
    h: 14,
    fill: "#89929A",
    line: ctx.line(C.none, 0),
    name: `bar-${index}-before`,
  });
  ctx.addShape(slide, {
    x: x0,
    y: y + 28,
    w: maxW,
    h: 18,
    fill: "#3D4650",
    line: ctx.line(C.none, 0),
    name: `bar-${index}-track-after`,
  });
  ctx.addShape(slide, {
    x: x0,
    y: y + 28,
    w: afterW,
    h: 18,
    fill: index === 2 ? C.gold : "#81C7B5",
    line: ctx.line(C.none, 0),
    name: `bar-${index}-after`,
  });
  addLabel(slide, ctx, before.toFixed(1), x0 + beforeW + 10, y - 3, 42, 18, "#BAC1C7", 12, `bar-${index}-before-value`);
  addLabel(slide, ctx, after.toFixed(1), x0 + afterW + 10, y + 27, 42, 20, "#F2E7D5", 13, `bar-${index}-after-value`);
}

export async function slide03(presentation, ctx) {
  const slide = presentation.slides.add();
  addBackground(slide, ctx, C.ink);

  ctx.addShape(slide, {
    x: 0,
    y: 0,
    w: 360,
    h: 720,
    fill: C.slate,
    line: ctx.line(C.none, 0),
    name: "left-context-rail",
  });
  ctx.addShape(slide, {
    x: 360,
    y: 0,
    w: 8,
    h: 720,
    fill: C.gold,
    line: ctx.line(C.none, 0),
    name: "context-rail-accent",
  });

  addKicker(slide, ctx, "qa loop", 64, 52, C.gold, true, "01");
  ctx.addText(slide, {
    text: "Separate checks",
    x: 64,
    y: 114,
    w: 230,
    h: 44,
    fontSize: 30,
    bold: true,
    color: C.white,
    face: ctx.fonts.title,
    name: "rail-heading",
  });
  addBody(
    slide,
    ctx,
    "Story logic, visual rhythm, and render integrity fail for different reasons. Treating them as one review step hides the source of quality defects.",
    66,
    180,
    230,
    138,
    "#CDBFAE",
    16,
    "rail-body",
  );
  addRule(slide, ctx, 64, 358, 220, "#59616B", 1, "rail-rule");
  ctx.addText(slide, {
    text: "1. Fix the claim\n2. Simplify the proof\n3. Re-render the deck",
    x: 66,
    y: 386,
    w: 230,
    h: 104,
    fontSize: 16,
    color: "#F2E7D5",
    name: "rail-steps",
  });

  addKicker(slide, ctx, "example scorecard", 420, 52, C.gold, true, "02");
  addClaimTitle(
    slide,
    ctx,
    "Quality improves fastest when evidence and rendering checks are separated.",
    420,
    94,
    760,
    C.white,
    39,
  );
  addBody(
    slide,
    ctx,
    "Illustrative 1-5 rubric scores after a first artifact-tool pass and after a focused QA iteration.",
    420,
    192,
    640,
    44,
    "#CDBFAE",
    16,
    "chart-note",
  );

  addLabel(slide, ctx, "First pass", 576, 266, 120, 18, "#BAC1C7", 12, "legend-before");
  ctx.addShape(slide, {
    x: 548,
    y: 270,
    w: 18,
    h: 8,
    fill: "#89929A",
    line: ctx.line(C.none, 0),
    name: "legend-before-mark",
  });
  addLabel(slide, ctx, "After QA", 740, 266, 120, 18, "#F2E7D5", 12, "legend-after");
  ctx.addShape(slide, {
    x: 712,
    y: 268,
    w: 18,
    h: 12,
    fill: "#81C7B5",
    line: ctx.line(C.none, 0),
    name: "legend-after-mark",
  });

  pairedBar(slide, ctx, { label: "Story", before: 3.3, after: 4.5 }, 320, 1);
  pairedBar(slide, ctx, { label: "Rhythm", before: 3.0, after: 4.3 }, 410, 2);
  pairedBar(slide, ctx, { label: "Render", before: 3.6, after: 4.7 }, 500, 3);

  addRule(slide, ctx, 576, 596, 424, "#59616B", 1, "chart-axis");
  addLabel(slide, ctx, "0", 568, 606, 24, 16, "#BAC1C7", 10, "axis-0");
  addLabel(slide, ctx, "5", 992, 606, 24, 16, "#BAC1C7", 10, "axis-5");
  addFooter(slide, ctx, 3, true);
  return slide;
}
