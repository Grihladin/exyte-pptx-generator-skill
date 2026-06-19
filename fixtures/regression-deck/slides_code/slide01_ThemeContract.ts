import type { PptxDeck, ThemeApi } from "../../../exyte-pptx-generator/starter/theme";

export default function buildSlide(pptx: PptxDeck, theme: ThemeApi): void {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addSlideTitle(slide, "Widescreen theme contract");
  theme.addSubheading(slide, "Stable canvas, typography, logo, and footer");
  theme.addBodyText(
    slide,
    [
      theme.createTextRun("The regression fixture confirms that "),
      theme.createEmphasisRun("all required chrome"),
      theme.createTextRun(" is present and remains inside the 13.333 × 7.5 inch canvas."),
    ],
    { y: 2.0, h: 0.8 },
  );
  theme.addCalloutBox(
    slide,
    [
      theme.createEmphasisRun("Validation: "),
      theme.createTextRun("object bounds, media signatures, fonts, paths, and slide metadata are checked."),
    ],
  );
}
