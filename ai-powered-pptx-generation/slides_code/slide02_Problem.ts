import type { PptxDeck, ThemeApi } from "../../theme";

export default function buildSlide(pptx: PptxDeck, theme: ThemeApi): void {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addTitle(slide, "The Problem: I've Never Used PowerPoint");

  slide.addShape("roundRect", {
    x: theme.LAYOUT.FREE_X,
    y: 1.1,
    w: 3.45,
    h: 3.75,
    fill: { color: theme.COLORS.BG_SURFACE },
    line: { color: theme.COLORS.BORDER, pt: 0.8 },
    rectRadius: 0.04,
  });

  slide.addText("Presentation tools I knew", {
    x: 1.05,
    y: 1.35,
    w: 2.95,
    h: 0.3,
    fontSize: 13,
    fontFace: theme.FONTS.BODY,
    color: theme.COLORS.LIGHT_BLUE,
    bold: true,
    lang: theme.PRESENTATION_LANGUAGE,
  });

  const knownTools = ["Canva", "Google Slides", "Plenty of presentations"];
  knownTools.forEach((tool, index) => {
    const y = 1.9 + index * 0.7;
    slide.addShape("roundRect", {
      x: 1.05,
      y,
      w: 2.95,
      h: 0.42,
      fill: { color: theme.COLORS.WHITE },
      line: { color: theme.COLORS.BORDER, pt: 0.6 },
      rectRadius: 0.03,
    });
    slide.addText(tool, {
      x: 1.22,
      y: y + 0.08,
      w: 2.6,
      h: 0.18,
      fontSize: 12,
      fontFace: theme.FONTS.BODY,
      color: theme.COLORS.TEXT_BODY,
      lang: theme.PRESENTATION_LANGUAGE,
    });
  });

  theme.addBody(
    slide,
    [
      theme.makeTextRun("The missing format was ", { fontSize: 16 }),
      theme.makeHighlightRun("PowerPoint", { fontSize: 16 }),
      theme.makeTextRun(
        ".\n\nI was born after PowerPoint had its peak popularity. I had done plenty of presentations, but always in Canva or Google Slides.\n\nI never touched PPTX. And now I have to.",
        { fontSize: 16 },
      ),
    ],
    {
      x: 4.55,
      y: 1.2,
      w: 4.15,
      h: 3.4,
      lineSpacing: 24,
      fit: "shrink",
    },
  );

  theme.addCalloutBox(
    slide,
    [
      theme.makeTextRun("The task was not just making slides. It was learning a legacy business format fast enough to use it in real meetings.", {
        fontSize: 12,
      }),
    ],
    { x: 4.55, y: 4.35, w: 4.15, h: 0.52 },
  );
}
