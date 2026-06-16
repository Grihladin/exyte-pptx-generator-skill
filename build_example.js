const pptxgen = require("pptxgenjs");
const theme = require("./theme.js");
const pptx = new pptxgen();

pptx.layout = "LAYOUT_WIDE";
pptx.author = "Exyte";
pptx.subject = "Example presentation for the updated Exyte typography and color skill";
pptx.title = "Brand Typography Example";
pptx.company = "Exyte";
pptx.lang = "en-US";

theme.updatePresentationSettings(
  "Brand Typography Example",
  "16/06/2026",
  "./assets/Exyte_RGB.svg.png"
);
theme.resetSlideCounter(0);

const slides = [
  "./slides/example01_BrandTypographyCover.js",
  "./slides/example02_TypeScale.js",
  "./slides/example03_ContentHierarchy.js",
  "./slides/example04_ColorUsage.js",
];

theme.TOTAL_SLIDES = slides.length;
slides.forEach((path) => require(path)(pptx, theme));

pptx.writeFile({ fileName: "outputs/example_brand_typography.pptx" });
