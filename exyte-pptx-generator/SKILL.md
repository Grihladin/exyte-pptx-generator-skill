---
name: exyte-pptx-generator
description: Generate and update self-contained Exyte-style PowerPoint PPTX deck projects from Markdown, notes, or user-provided content using TypeScript, PptxGenJS, a bundled widescreen theme, local brand assets, and strict structural validation. Use for corporate presentations, slide decks, generated slide source files, or PPTX build and verification workflows.
---

# Exyte PPTX Generator

Create an editable PowerPoint deck project from user-provided source material. Use the bundled initializer, local theme, PNG logo, TypeScript slide files, and validator.

## Initialize The Deck

Create the project from this skill directory:

```bash
node --import tsx scripts/init-deck.ts <topic-slug> \
  --output-dir <workspace-root> \
  --title "<presentation title>" \
  --date "<DD/MM/YYYY>" \
  --source "<optional-source-file>"
```

The initializer sanitizes the slug and never overwrites an existing directory. It appends `-2`, `-3`, and so on when necessary.

Initialization creates project source files only. It also creates a `.gitignore` for `node_modules/`, PPTX files, preview/output directories, and `.DS_Store`; keep `content/` tracked. The PPTX is created later by `npm run build`.

Enter the printed project directory and author the deck using the bundled project files.

## Inspect The Installed API

Before authoring slides:

1. Read `node_modules/pptxgenjs/package.json`.
2. Resolve the declaration entry from `exports.types` or `types`.
3. Search the declaration file for each API or option the deck will use.
4. Read the complete relevant interfaces and signatures.

Use installed declarations as the source of truth. Do not guess unfamiliar PptxGenJS options or copy examples from another version. Reading the entire declaration file is unnecessary.

## Author The Deck

1. Read all source material under `content/`.
2. Create a short TODO for every planned slide.
3. Keep one TypeScript file per slide under `slides_code/`.
4. Import types from the project-local `../theme`.
5. Register every slide in `slides_code/build.ts`.
6. Update `--expected-slides` in `package.json` to the registered slide count.
7. Keep `pptx.layout = "LAYOUT_WIDE"`.
8. Keep `pptx.theme` set to the local Arial heading and body fonts.
9. Use `assets/exyte_logo.png`; do not embed SVG logos.

Use this slide pattern:

```typescript
import type { PptxDeck, ThemeApi } from "../theme";

export default function buildSlide(pptx: PptxDeck, theme: ThemeApi): void {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addSlideTitle(slide, "Slide title");
  // Add custom content inside theme.LAYOUT.FREE_*.
}
```

## Layout Contract

- Use the 13.333 × 7.5 inch PowerPoint Widescreen canvas.
- Call `theme.applySlideBase(slide)` exactly once per slide to add the background, logo, footer, and page number.
- Call `theme.addSlideTitle()` separately for the normal slide title.
- Keep every custom object fully inside `theme.LAYOUT.FREE_*`. The validator enforces the exact rounded EMU boundaries with no overflow allowance.
- Respect `theme.LAYOUT.TITLE_W` and the reserved logo zone.
- Do not recreate the logo, footer line, footer text, date, divider, or page number.
- Do not use absolute local paths as alt text, object names, visible content, or notes.

The theme owns stable object names for validation. Do not rename its chrome objects.

## Typography And Color Contract

- Use Arial for all editable slide text.
- Keep normal body, table, card, and callout text at 15 pt.
- Use the theme's 24 pt standard title style.
- Use `theme.addSubheading()`, `theme.addBodyText()`, `theme.addThemedTable()`, `theme.createTextRun()`, and `theme.createEmphasisRun()` for standard content.
- Use bold black text for inline emphasis.
- Use `theme.COLORS` for custom fills, lines, and text.
- Hand-specify typography only for genuinely custom display elements.
- Do not use emoji.

## Project Contract

```text
<topic-slug>/
  .gitignore
  assets/
    exyte_logo.png
  content/
    source.md
  scripts/
    validate-pptx.ts
  slides_code/
    build.ts
    slide01_DescriptiveName.ts
  package.json
  package-lock.json
  theme.ts
  tsconfig.json
```

The `<topic-slug>.pptx` output appears after `npm run build`.

Keep user inputs under `content/`. Never modify or move the user's originals unless explicitly asked.

## Verify Before Delivery

Run:

```bash
npm run verify
```

The required validator checks:

- archive integrity and exact slide count;
- 13.333 × 7.5 inch slide dimensions;
- off-slide objects, footer intrusion, and strict containment of all custom objects within `LAYOUT.FREE_*`;
- title/logo separation;
- required logo and footer chrome;
- Arial typography;
- valid media signatures and the 1244 × 500 PNG logo;
- prohibited SVG relationships;
- absolute filesystem path leakage;
- ambiguous sibling PPTX outputs.

After structural verification, render every slide in PowerPoint, Quick Look, or LibreOffice and manually inspect text wrapping, clipping, overlap, alignment, image quality, and overall visual quality. If rendering is unavailable, report that caveat. Structural validation remains mandatory.

Fix every validation or visible rendering defect and rerun `npm run verify` before returning the final PPTX path.
