module.exports = function (pptx, theme) {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addTitle(slide, "The type scale is intentionally simple.");

  const rows = [
    {
      label: "Page header",
      sample: "Arial 24, Mid blue",
      options: {
        fontSize: theme.FONTS.TITLE_SIZE,
        fontFace: theme.FONTS.HEADING,
        color: theme.COLORS.MID_BLUE,
        bold: true,
      },
    },
    {
      label: "Subheader",
      sample: "Arial Bold 15, Light blue",
      options: {
        fontSize: theme.FONTS.SUBHEADER_SIZE,
        fontFace: theme.FONTS.BODY,
        color: theme.COLORS.LIGHT_BLUE,
        bold: true,
      },
    },
    {
      label: "Main text",
      sample: "Arial 12-15, black",
      options: {
        fontSize: 13,
        fontFace: theme.FONTS.BODY,
        color: theme.COLORS.TEXT_BODY,
      },
    },
    {
      label: "Highlight",
      sample: "Arial Bold 15",
      options: {
        fontSize: theme.FONTS.HIGHLIGHT_SIZE,
        fontFace: theme.FONTS.BODY,
        color: theme.COLORS.LIGHT_BLUE,
        bold: true,
      },
    },
  ];

  rows.forEach((row, index) => {
    const y = 1.15 + index * 0.78;

    slide.addText(row.label, {
      x: theme.LAYOUT.FREE_X,
      y,
      w: 2.0,
      h: 0.42,
      fontSize: 12,
      fontFace: theme.FONTS.BODY,
      color: theme.COLORS.TEXT_BODY,
      bold: true,
      valign: "middle",
    });

    slide.addText(row.sample, {
      x: 3.0,
      y: y - 0.05,
      w: 5.6,
      h: 0.5,
      valign: "middle",
      fit: "shrink",
      ...row.options,
    });

    slide.addShape("line", {
      x: theme.LAYOUT.FREE_X,
      y: y + 0.55,
      w: theme.LAYOUT.FREE_W,
      h: 0,
      line: { color: theme.COLORS.BORDER, width: 0.5 },
    });
  });

  slide.addText([
    theme.makeTextRun("Inline emphasis should be sparse: "),
    theme.makeHighlightRun("highlight the decision word"),
    theme.makeTextRun(", not the whole paragraph."),
  ], {
    x: theme.LAYOUT.FREE_X,
    y: 4.35,
    w: theme.LAYOUT.FREE_W,
    h: 0.35,
    fontFace: theme.FONTS.BODY,
    fontSize: 13,
    color: theme.COLORS.TEXT_BODY,
    valign: "middle",
  });
};
