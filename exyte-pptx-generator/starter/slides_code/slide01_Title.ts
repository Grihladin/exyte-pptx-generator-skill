import type { PptxDeck, ThemeApi } from "../theme";

export default function buildSlide(pptx: PptxDeck, theme: ThemeApi): void {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addSlideTitle(slide, "__DECK_TITLE_TS__");
  theme.addSubheading(slide, "Replace this starter slide with presentation content.");
  theme.addBodyText(
    slide,
    [
      theme.createTextRun("Use the local "),
      theme.createEmphasisRun("theme.ts"),
      theme.createTextRun(" helpers and keep custom content inside theme.LAYOUT.FREE_*."),
    ],
    { y: theme.LAYOUT.FREE_Y + 0.65, h: 0.8 },
  );
}
