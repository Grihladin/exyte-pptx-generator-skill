# Presentation Generator -- Exyte Corporate Blue

Generate a PowerPoint presentation from a Markdown content file using `pptxgenjs` and the design system in `theme.js`.

**Input:** `$ARGUMENTS` (path to a .md file containing slide content)

## Design Rules

- Design clean, readable slides. Use cards, bullet points, tables, and shapes/graphics where they fit.
- Vary the layout across slides so the deck does not feel repetitive.
- Use pptxgenjs shapes to build simple diagrams, flowcharts, or visual elements when they communicate an idea better than text alone.
- No emoji.
- Do not use vertical accent bars (left-side colored bars) on callout boxes or cards. Use background color and rounded corners instead.

## Workflow

1. Read the provided .md file for content.
2. Choose a concise presentation title/topic, convert it to a safe lowercase hyphen slug, and create `<topic-slug>/slides_code/` at the repo root.
3. Create a task/TODO for each slide before starting.
4. Build one slide at a time in `<topic-slug>/slides_code/`.
5. Create/update `<topic-slug>/slides_code/build.js` after all slides are done, then run `node <topic-slug>/slides_code/build.js` to verify.

## File Output Contract

Create a new folder for each presentation at the repo root. Use the same topic slug for the folder and final PPTX filename:

```text
<topic-slug>/
  <topic-slug>.pptx
  slides_code/
    build.js
    slide01_DescriptiveName.js
    slide02_DescriptiveName.js
```

If the topic folder already exists, append `-2`, `-3`, etc. to the folder and PPTX slug instead of overwriting.

## Slide File Pattern

One file per slide in `<topic-slug>/slides_code/`. Each exports `function(pptx, theme)`:

```javascript
module.exports = function (pptx, theme) {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide);        // bg, logo, footer (required)
  // theme.addTitle(slide, "Title");   // optional -- skip for statement/quote slides
  // ... free content below
};
```

## Layout Contract

- Start each slide with `theme.applySlideBase(slide)`; it owns the background, logo, and footer.
- Use `theme.addTitle()` for normal page headers.
- Place custom content inside `theme.LAYOUT.FREE_*`.
- Do not create manual footer text, page numbers, source labels, bottom rules, or competing header chrome.

## Typography And Color Contract

- Use the typography and color helpers in `theme.js`; do not reimplement standard text styling in slide files.
- Use `theme.addTitle()`, `theme.addSubheader()` / `theme.addSubtitle()`, `theme.addBody()`, `theme.makeTextRun()`, and `theme.makeHighlightRun()` for standard text.
- Use `theme.COLORS` constants for custom shapes, fills, borders, or manual text overrides.
- Only hand-specify fonts, sizes, or hex colors when a genuinely custom element is not covered by the theme helpers.

## Theme API

Prefer `theme.js` helpers before raw `slide.addText()` / `slide.addShape()` styling.

Helpers: `applySlideBase`, `addTitle`, `addSubheader`, `addSubtitle`, `addBody`, `addStyledTable`, `makeTextRun`, `makeHighlightRun`, `addCalloutBox`.

Constants: `theme.COLORS`, `theme.FONTS`, `theme.LAYOUT`.

For raw pptxgenjs options, inspect `node_modules/pptxgenjs/types/index.d.ts`.

## Slide Naming

Name each slide file `slideNN_DescriptiveName.js` where `NN` is the zero-padded number and `DescriptiveName` is a short PascalCase name based on the slide content (e.g. `slide01_TitleSlide.js`, `slide04_MarketOverview.js`).

## build.js Pattern

- **Title**: Choose a concise, descriptive presentation title based on the .md file content.
- **Slug**: Use the final topic slug for both the presentation folder and PPTX filename.
- **Date**: Use today's date in `DD/MM/YYYY` format.
- **Location**: Save this file as `<topic-slug>/slides_code/build.js`.
- **Logo**: Always load `assets/Exyte_RGB.svg.png` from the repo root.

```javascript
const path = require("path");
const pptxgen = require("pptxgenjs");
const theme = require(path.join(__dirname, "..", "..", "theme.js"));
const pptx = new pptxgen();

const deckSlug = "<topic-slug>";
const outputPath = path.join(__dirname, "..", `${deckSlug}.pptx`);
const logoPath = path.join(__dirname, "..", "..", "assets", "Exyte_RGB.svg.png");

theme.updatePresentationSettings("<title from content>", "<today DD/MM/YYYY>", logoPath);
theme.resetSlideCounter(0);

const slides = [
  "./slide01_TitleSlide.js",
  // ... all slides
];

slides.forEach((slidePath) => require(path.join(__dirname, slidePath))(pptx, theme));
pptx.writeFile({ fileName: outputPath });
```
