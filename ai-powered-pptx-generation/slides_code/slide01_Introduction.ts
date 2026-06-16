import type { PptxDeck, ThemeApi } from "../../theme";

export default function buildSlide(pptx: PptxDeck, theme: ThemeApi): void {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addTitle(slide, "AI-Powered PPTX Generation");

  theme.addSubheader(slide, "From personal hack to product vision", {
    x: theme.LAYOUT.FREE_X,
    y: 1.02,
    w: 6.15,
    h: 0.35,
  });

  slide.addShape("roundRect", {
    x: theme.LAYOUT.FREE_X,
    y: 1.6,
    w: 8.4,
    h: 1.0,
    fill: { color: theme.COLORS.BG_SURFACE },
    line: { color: theme.COLORS.BORDER, pt: 0.8 },
    rectRadius: 0.04,
  });

  slide.addText(
    [
      theme.makeHighlightRun("Not an internal Exyte project. "),
      theme.makeTextRun("A side project that grew out of a real daily need."),
    ],
    {
      x: 1.05,
      y: 1.78,
      w: 7.9,
      h: 0.55,
      fontSize: 15,
      breakLine: false,
      fit: "shrink",
      valign: "middle",
    },
  );

  theme.addBody(
    slide,
    "Since joining Exyte, I have had to create many PowerPoint presentations for architecture reviews, technical reviews, and other meetings.\n\nA well-structured PPTX helps me stay on track and make sure I do not forget important points during the discussion.",
    {
      x: 1.0,
      y: 3.0,
      w: 7.9,
      h: 1.65,
      fontSize: 15,
      lineSpacingMultiple: 1.1,
      fit: "shrink",
    },
  );

  slide.addText("Michael Ratke | IPAI AI-Circle", {
    x: 1.0,
    y: 4.78,
    w: 7.9,
    h: 0.25,
    fontSize: theme.FONTS.SMALL_SIZE,
    fontFace: theme.FONTS.BODY,
    color: theme.COLORS.TEXT_MUTED,
    align: "right",
    lang: theme.PRESENTATION_LANGUAGE,
  });
}
