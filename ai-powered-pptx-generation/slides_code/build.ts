import path from "node:path";
import pptxgen from "pptxgenjs";
import * as theme from "../../theme";
import slide01 from "./slide01_Introduction";
import slide02 from "./slide02_Problem";
import slide03 from "./slide03_Research";

const deckSlug = "ai-powered-pptx-generation";
const outputPath = path.join(__dirname, "..", `${deckSlug}.pptx`);
const logoPath = path.join(__dirname, "..", "..", "assets", "Exyte_RGB.svg");
const title = "AI-Powered PPTX Generation";

async function main(): Promise<void> {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "Exyte";
  pptx.company = "Exyte";
  pptx.subject = "From personal hack to product vision";
  pptx.title = title;

  theme.updatePresentationSettings(title, "16/06/2026", logoPath);
  theme.resetSlideCounter(0);

  const slides = [slide01, slide02, slide03];
  slides.forEach((buildSlide) => buildSlide(pptx, theme));

  await pptx.writeFile({ fileName: outputPath });
}

main().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});
