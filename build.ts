import path from "node:path";
import pptxgen from "pptxgenjs";
import * as theme from "./exyte-pptx-generator/starter/theme";

async function main(): Promise<void> {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Exyte";
  pptx.company = "Exyte";
  pptx.subject = "Smoke test for the Exyte PPTX generator";
  pptx.title = "PPTX Generator Smoke Test";
  pptx.theme = {
    headFontFace: theme.TYPOGRAPHY.HEADING,
    bodyFontFace: theme.TYPOGRAPHY.BODY,
  };

  const outputPath = path.join(__dirname, "output.pptx");
  const logoPath = path.join(__dirname, "exyte-pptx-generator", "starter", "assets", "exyte_logo.png");

  theme.configurePresentation("PPTX Generator Smoke Test", "19/06/2026", logoPath);
  theme.resetSlideCounter(0);

  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addSlideTitle(slide, "PPTX Generator Smoke Test");
  theme.addSubheading(slide, "Self-contained widescreen theme and validation contract");
  theme.addBodyText(
    slide,
    [
      theme.createTextRun("This deck verifies the "),
      theme.createEmphasisRun("13.333 × 7.5 inch"),
      theme.createTextRun(" canvas, local PNG logo, Arial typography, and stable Exyte chrome."),
    ],
    { y: theme.LAYOUT.FREE_Y + 0.65, h: 0.8 },
  );
  theme.addCalloutBox(
    slide,
    [
      theme.createEmphasisRun("Smoke check: "),
      theme.createTextRun("the generated package must pass strict structural validation before delivery."),
    ],
  );

  await pptx.writeFile({ fileName: outputPath, compression: true });
}

main().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});
