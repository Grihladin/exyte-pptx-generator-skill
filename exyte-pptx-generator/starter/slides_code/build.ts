import path from "node:path";
import pptxgen from "pptxgenjs";
import * as theme from "../theme";
import slide01 from "./slide01_Title";

const deckSlug = "__DECK_SLUG_TS__";
const deckTitle = "__DECK_TITLE_TS__";
const outputPath = path.join(__dirname, "..", `${deckSlug}.pptx`);
const logoPath = path.join(__dirname, "..", "assets", "exyte_logo.png");

async function main(): Promise<void> {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Exyte";
  pptx.company = "Exyte";
  pptx.subject = deckTitle;
  pptx.title = deckTitle;
  pptx.theme = {
    headFontFace: theme.TYPOGRAPHY.HEADING,
    bodyFontFace: theme.TYPOGRAPHY.BODY,
  };

  theme.configurePresentation(deckTitle, "__DECK_DATE_TS__", logoPath);
  theme.resetSlideCounter(0);

  const slides = [slide01];
  slides.forEach((buildSlide) => buildSlide(pptx, theme));

  await pptx.writeFile({ fileName: outputPath, compression: true });
}

main().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});
