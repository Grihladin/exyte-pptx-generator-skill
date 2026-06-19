import path from "node:path";
import pptxgen from "pptxgenjs";
import * as theme from "../../../exyte-pptx-generator/starter/theme";
import slide01 from "./slide01_ThemeContract";
import slide02 from "./slide02_ContentLayout";

const outputPath = path.join(__dirname, "..", "regression-deck.pptx");
const logoPath = path.join(
  __dirname,
  "..",
  "..",
  "..",
  "exyte-pptx-generator",
  "starter",
  "assets",
  "exyte_logo.png",
);

async function main(): Promise<void> {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Exyte";
  pptx.company = "Exyte";
  pptx.subject = "Neutral regression fixture";
  pptx.title = "Exyte PPTX Regression Fixture";
  pptx.theme = {
    headFontFace: theme.TYPOGRAPHY.HEADING,
    bodyFontFace: theme.TYPOGRAPHY.BODY,
  };

  theme.configurePresentation("Regression Fixture", "19/06/2026", logoPath);
  theme.resetSlideCounter(0);
  [slide01, slide02].forEach((buildSlide) => buildSlide(pptx, theme));

  await pptx.writeFile({ fileName: outputPath, compression: true });
}

main().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});
