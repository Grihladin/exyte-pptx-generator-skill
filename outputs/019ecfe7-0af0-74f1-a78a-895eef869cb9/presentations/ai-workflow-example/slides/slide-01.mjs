import { C, addBackground, addBody, addClaimTitle, addFooter, addKicker, addRule } from "./style.mjs";

function metric(slide, ctx, { value, label, note }, y, accent, index) {
  addRule(slide, ctx, 840, y - 18, 330, "#47515A", 1, `metric-${index}-rule`);
  ctx.addText(slide, {
    text: value,
    x: 840,
    y,
    w: 140,
    h: 54,
    fontSize: 38,
    bold: true,
    face: ctx.fonts.title,
    color: accent,
    name: `metric-${index}-value`,
  });
  ctx.addText(slide, {
    text: label,
    x: 990,
    y: y + 6,
    w: 185,
    h: 24,
    fontSize: 17,
    bold: true,
    color: C.white,
    name: `metric-${index}-label`,
  });
  ctx.addText(slide, {
    text: note,
    x: 990,
    y: y + 34,
    w: 190,
    h: 38,
    fontSize: 12,
    color: "#CDBFAE",
    name: `metric-${index}-context`,
  });
}

export async function slide01(presentation, ctx) {
  const slide = presentation.slides.add();
  addBackground(slide, ctx, C.paper);

  ctx.addShape(slide, {
    x: 782,
    y: 0,
    w: 498,
    h: 720,
    fill: C.ink,
    line: ctx.line(C.none, 0),
    name: "right-ink-rail",
  });
  ctx.addShape(slide, {
    x: 782,
    y: 0,
    w: 12,
    h: 720,
    fill: C.rust,
    line: ctx.line(C.none, 0),
    name: "right-rail-accent",
  });

  addKicker(slide, ctx, "workflow snapshot", 64, 55, C.rust, false, "01");
  addClaimTitle(
    slide,
    ctx,
    "A smaller team can produce a board-ready narrative in one working day.",
    64,
    116,
    640,
    C.ink,
    50,
  );
  addBody(
    slide,
    ctx,
    "Three editable slides that demonstrate the PPTX skill path: claim spine, design lock, artifact-tool build, render QA, and final PowerPoint export.",
    68,
    316,
    560,
    90,
    C.muted,
    19,
    "cover-support-note",
  );

  addRule(slide, ctx, 68, 458, 540, C.rule, 1, "cover-lower-rule");
  ctx.addText(slide, {
    text: "Example deck",
    x: 68,
    y: 484,
    w: 220,
    h: 32,
    fontSize: 17,
    bold: true,
    color: C.teal,
    name: "example-deck-label",
  });
  ctx.addText(slide, {
    text: "Built from native PowerPoint text, shapes, bars, and layout primitives.",
    x: 68,
    y: 520,
    w: 470,
    h: 44,
    fontSize: 15,
    color: C.muted,
    name: "example-deck-note",
  });

  addKicker(slide, ctx, "sample output", 840, 55, C.gold, true, "02");
  ctx.addText(slide, {
    text: "Editable proof rail",
    x: 840,
    y: 116,
    w: 320,
    h: 48,
    fontSize: 34,
    bold: true,
    face: ctx.fonts.title,
    color: C.white,
    name: "rail-title",
  });
  ctx.addText(slide, {
    text: "The values below are placeholders used to show how a metric system renders.",
    x: 842,
    y: 168,
    w: 310,
    h: 48,
    fontSize: 14,
    color: "#CDBFAE",
    name: "rail-note",
  });

  metric(
    slide,
    ctx,
    {
      value: "1 day",
      label: "brief to draft",
      note: "story, build, and QA kept as separate loops",
    },
    258,
    C.gold,
    1,
  );
  metric(
    slide,
    ctx,
    {
      value: "3x",
      label: "QA gates",
      note: "story logic, render quality, package checks",
    },
    376,
    "#E7D1A0",
    2,
  );
  metric(
    slide,
    ctx,
    {
      value: "100%",
      label: "editable",
      note: "text, shapes, and bars remain PowerPoint objects",
    },
    494,
    "#9ED4C4",
    3,
  );

  addFooter(slide, ctx, 1, false);
  return slide;
}
