module.exports = function (pptx, theme) {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addTitle(slide, "Color separates structure from reading text.");

  const panels = [
    {
      x: theme.LAYOUT.FREE_X,
      color: theme.COLORS.MID_BLUE,
      title: "Mid blue",
      body: "Primary headers, footer line, and strong structural emphasis.",
    },
    {
      x: theme.LAYOUT.FREE_X + 4.35,
      color: theme.COLORS.LIGHT_BLUE,
      title: "Light blue",
      body: "Secondary accents, subheaders, and highlighted words.",
    },
  ];

  panels.forEach((panel) => {
    slide.addShape("roundRect", {
      x: panel.x,
      y: 1.2,
      w: 4.05,
      h: 1.45,
      rectRadius: 0.06,
      fill: { color: panel.color },
      line: { color: panel.color, transparency: 100 },
    });

    slide.addText(panel.title, {
      x: panel.x + 0.25,
      y: 1.38,
      w: 3.5,
      h: 0.32,
      fontSize: 15,
      fontFace: theme.FONTS.BODY,
      color: theme.COLORS.WHITE,
      bold: true,
      valign: "middle",
    });

    slide.addText(panel.body, {
      x: panel.x + 0.25,
      y: 1.85,
      w: 3.45,
      h: 0.48,
      fontSize: 12,
      fontFace: theme.FONTS.BODY,
      color: theme.COLORS.WHITE,
      valign: "top",
      fit: "shrink",
    });
  });

  theme.addSubheader(slide, "Ordinary text rule", {
    x: theme.LAYOUT.FREE_X,
    y: 3.05,
    w: 3.0,
    h: 0.3,
  });

  slide.addText([
    theme.makeTextRun("Use "),
    theme.makeHighlightRun("black"),
    theme.makeTextRun(" for normal text on white and "),
    theme.makeHighlightRun("white"),
    theme.makeTextRun(" only when the text sits on a blue fill."),
  ], {
    x: theme.LAYOUT.FREE_X,
    y: 3.52,
    w: theme.LAYOUT.FREE_W,
    h: 0.5,
    fontSize: 13,
    fontFace: theme.FONTS.BODY,
    color: theme.COLORS.TEXT_BODY,
    valign: "top",
    fit: "shrink",
  });

  slide.addText("The result is a restrained deck system: blue for navigation and emphasis; black and white for readability.", {
    x: theme.LAYOUT.FREE_X,
    y: 4.32,
    w: theme.LAYOUT.FREE_W,
    h: 0.35,
    fontSize: 12,
    fontFace: theme.FONTS.BODY,
    color: theme.COLORS.TEXT_BODY,
    valign: "middle",
  });
};
