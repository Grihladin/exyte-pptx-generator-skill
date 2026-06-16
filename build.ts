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
  const logoPath = path.join(__dirname, "assets", "Exyte_RGB.svg.png");

  theme.updatePresentationSettings("PPTX Generator Smoke Test", "16/06/2026", logoPath);
  theme.resetSlideCounter(0);

  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addTitle(slide, "PPTX Generator Smoke Test");
  theme.addSubheader(slide, "TypeScript runtime and theme contract");
  theme.addBody(
    slide,
    [
      theme.makeTextRun("This deck verifies that "),
      theme.makeHighlightRun("theme.ts"),
      theme.makeTextRun(" can create a PPTX through "),
      theme.makeHighlightRun("pptxgenjs"),
      theme.makeTextRun(" using the shared Exyte layout contract."),
    ],
    { y: theme.LAYOUT.FREE_Y + 0.6, h: 0.8 },
  );
  theme.addCalloutBox(
    slide,
    [
      theme.makeTextRun("Smoke check: ", { bold: true, color: theme.COLORS.ACCENT }),
      theme.makeTextRun("logo, footer, page number, date, and 16:9 slide sizing are applied by the theme."),
    ],
    { y: 4.1, h: 0.65 },
  );

  await pptx.writeFile({ fileName: outputPath });
}

main().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});
