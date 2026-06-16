module.exports = function (pptx, theme) {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addTitle(slide, "Brand Typography Example");

  slide.addText("Updated presentation style", {
    x: theme.LAYOUT.FREE_X,
    y: 1.35,
    w: theme.LAYOUT.FREE_W,
    h: 0.4,
    fontSize: theme.FONTS.SUBHEADER_SIZE,
    fontFace: theme.FONTS.BODY,
    color: theme.COLORS.SECONDARY,
    bold: true,
    align: "center",
    valign: "middle",
  });

  slide.addText("A compact example deck showing the new Arial typography, Mid blue headers, Light blue secondary accents, and black/white body text rules.", {
    x: 1.45,
    y: 1.95,
    w: 7.1,
    h: 0.85,
    fontSize: 15,
    fontFace: theme.FONTS.BODY,
    color: theme.COLORS.TEXT_BODY,
    align: "center",
    valign: "middle",
    fit: "shrink",
    breakLine: false,
  });

  slide.addShape("rect", {
    x: 2.3,
    y: 3.25,
    w: 2.25,
    h: 0.62,
    fill: { color: theme.COLORS.MID_BLUE },
    line: { color: theme.COLORS.MID_BLUE, transparency: 100 },
  });
  slide.addText("Mid blue", {
    x: 2.3,
    y: 3.25,
    w: 2.25,
    h: 0.62,
    fontSize: 15,
    fontFace: theme.FONTS.BODY,
    color: theme.COLORS.WHITE,
    bold: true,
    align: "center",
    valign: "middle",
  });

  slide.addShape("rect", {
    x: 4.65,
    y: 3.25,
    w: 2.25,
    h: 0.62,
    fill: { color: theme.COLORS.LIGHT_BLUE },
    line: { color: theme.COLORS.LIGHT_BLUE, transparency: 100 },
  });
  slide.addText("Light blue", {
    x: 4.65,
    y: 3.25,
    w: 2.25,
    h: 0.62,
    fontSize: 15,
    fontFace: theme.FONTS.BODY,
    color: theme.COLORS.WHITE,
    bold: true,
    align: "center",
    valign: "middle",
  });

  slide.addText("Ordinary text stays black on white. White text is reserved for blue fills.", {
    x: 1.55,
    y: 4.2,
    w: 6.9,
    h: 0.35,
    fontSize: 12,
    fontFace: theme.FONTS.BODY,
    color: theme.COLORS.TEXT_BODY,
    align: "center",
    valign: "middle",
  });
};
