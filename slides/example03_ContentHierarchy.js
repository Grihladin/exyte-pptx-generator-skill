module.exports = function (pptx, theme) {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addTitle(slide, "Hierarchy keeps dense slides readable.");

  theme.addSubheader(slide, "Example workflow", {
    x: theme.LAYOUT.FREE_X,
    y: 1.05,
    w: 3.2,
    h: 0.3,
  });

  const steps = [
    {
      number: "1",
      title: "Frame",
      body: "Start with the single message the slide needs to prove.",
    },
    {
      number: "2",
      title: "Structure",
      body: "Choose one proof object and place support text around it.",
    },
    {
      number: "3",
      title: "Polish",
      body: "Check contrast, spacing, and whether text stays within the 12-15 range.",
    },
  ];

  steps.forEach((step, index) => {
    const x = theme.LAYOUT.FREE_X + index * 2.85;

    slide.addShape("roundRect", {
      x,
      y: 1.55,
      w: 2.45,
      h: 2.25,
      rectRadius: 0.06,
      fill: { color: index === 1 ? theme.COLORS.MID_BLUE : theme.COLORS.BG_CARD },
      line: { color: index === 1 ? theme.COLORS.MID_BLUE : theme.COLORS.BORDER, width: 0.6 },
    });

    slide.addText(step.number, {
      x: x + 0.16,
      y: 1.72,
      w: 0.38,
      h: 0.38,
      fontSize: 15,
      fontFace: theme.FONTS.BODY,
      color: index === 1 ? theme.COLORS.MID_BLUE : theme.COLORS.WHITE,
      bold: true,
      align: "center",
      valign: "middle",
      fill: { color: index === 1 ? theme.COLORS.WHITE : theme.COLORS.MID_BLUE },
    });

    slide.addText(step.title, {
      x: x + 0.64,
      y: 1.73,
      w: 1.55,
      h: 0.35,
      fontSize: 15,
      fontFace: theme.FONTS.BODY,
      color: index === 1 ? theme.COLORS.WHITE : theme.COLORS.LIGHT_BLUE,
      bold: true,
      valign: "middle",
      fit: "shrink",
    });

    slide.addText(step.body, {
      x: x + 0.22,
      y: 2.35,
      w: 2.0,
      h: 0.9,
      fontSize: 12,
      fontFace: theme.FONTS.BODY,
      color: index === 1 ? theme.COLORS.WHITE : theme.COLORS.TEXT_BODY,
      valign: "top",
      fit: "shrink",
      breakLine: false,
    });
  });

  theme.addCalloutBox(
    slide,
    [
      theme.makeHighlightRun("Rule of thumb: "),
      theme.makeTextRun("one dominant idea, one dominant proof object, and no decorative text colors."),
    ],
    { y: 4.18, h: 0.58, fontSize: 13 }
  );
};
