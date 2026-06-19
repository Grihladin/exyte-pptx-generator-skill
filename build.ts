import path from "node:path";
import pptxgen from "pptxgenjs";
import * as theme from "./theme";

async function main(): Promise<void> {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "Exyte";
  pptx.company = "Exyte";
  pptx.subject = "Smoke test for the TypeScript PPTX generator theme";
  pptx.title = "PPTX Generator Smoke Test";

  const outputPath = path.join(__dirname, "output.pptx");
  const logoPath = path.join(__dirname, "exyte-pptx-generator", "exyte_logo.svg");

  theme.configurePresentation("PPTX Generator Smoke Test", "16/06/2026", logoPath);
  theme.resetSlideCounter(0);

  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addSlideTitle(slide, "PPTX Generator Smoke Test");
  theme.addSubheading(slide, "TypeScript runtime and theme contract");
  theme.addBodyText(
    slide,
    [
      theme.createTextRun("This deck verifies that "),
      theme.createEmphasisRun("theme.ts"),
      theme.createTextRun(" can create a PPTX through "),
      theme.createEmphasisRun("pptxgenjs"),
      theme.createTextRun(" using the shared Exyte layout contract."),
    ],
    { y: theme.LAYOUT.FREE_Y + 0.6, h: 0.8 },
  );
  theme.addCalloutBox(
    slide,
    [
      theme.createTextRun("Smoke check: ", { bold: true, color: theme.COLORS.ACCENT }),
      theme.createTextRun("logo, footer, page number, date, and 16:9 slide sizing are applied by the theme."),
    ],
    { y: 4.1, h: 0.65 },
  );

  await pptx.writeFile({ fileName: outputPath });
}

main().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});
