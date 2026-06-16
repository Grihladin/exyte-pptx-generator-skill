const pptxgen = require("pptxgenjs");
const theme = require("./theme.js");
const pptx = new pptxgen();

theme.updatePresentationSettings(
  "AI-Powered PPTX Generation",
  "16/06/2026",
  "./assets/Exyte_RGB.svg.png"
);
theme.resetSlideCounter(0);

const slides = [
  "./slides/slide01_Introduction.js",
  "./slides/slide02_TheProblem.js",
  "./slides/slide03_PythonPptx.js",
];

theme.TOTAL_SLIDES = slides.length;
slides.forEach((path) => require(path)(pptx, theme));
pptx.writeFile({ fileName: "output.pptx" });
