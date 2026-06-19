import type { PptxDeck, ThemeApi } from "../../../theme";

export default function buildSlide(pptx: PptxDeck, theme: ThemeApi): void {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addSlideTitle(slide, "Content remains inside the safe frame");

  const cards = [
    { x: 0.8, title: "Local theme", body: "Every generated deck carries its own theme and assets." },
    { x: 4.83, title: "Portable QA", body: "Structural validation does not depend on PowerPoint or macOS." },
    { x: 8.86, title: "Editable output", body: "Slides remain native PowerPoint shapes and text." },
  ];

  for (const card of cards) {
    slide.addShape("roundRect", {
      x: card.x,
      y: 1.65,
      w: 3.67,
      h: 3.55,
      fill: { color: theme.COLORS.BG_CARD },
      line: { color: theme.COLORS.BORDER, width: 0.8 },
      rectRadius: 0.04,
    });
    theme.addSubheading(slide, card.title, {
      x: card.x + 0.25,
      y: 2.05,
      w: 3.17,
      h: 0.35,
      align: "center",
    });
    theme.addBodyText(slide, card.body, {
      x: card.x + 0.3,
      y: 2.65,
      w: 3.07,
      h: 1.2,
      align: "center",
      valign: "middle",
      margin: 0,
    });
  }
}
