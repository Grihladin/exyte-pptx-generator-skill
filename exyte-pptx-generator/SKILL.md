---
name: exyte-pptx-generator
description: Generate branded Exyte-style PowerPoint PPTX decks from Markdown or user-provided slide content using TypeScript, pptxgenjs, theme.ts, and local presentation assets. Use when creating or updating corporate-style .pptx presentations, slide decks, generated slide source files, or PPTX build scripts.
---

# PPTX Generator

Generate a PowerPoint presentation from the Markdown file, inline notes, or source material requested by the user. Use `pptxgenjs`, TypeScript slide files, and the workspace `theme.ts` design system.

## Initial Setup

Before generating slides, check the local runtime:

```bash
node --version
npm --version
npm ls pptxgenjs typescript tsx @types/node --depth=0
```

If Node or npm is missing, stop and ask whether the user wants help setting it up. Do not install system software without approval.

If dependencies are missing, ask before running setup or repair commands. After approval:

- Run `npm install` when `package.json` already declares the needed packages.
- Run `npm install pptxgenjs` when the runtime dependency is absent.
- Run `npm install -D typescript tsx @types/node` when TypeScript tooling is absent.

If setup is declined, continue only when the deck can be generated and verified with existing dependencies.

## Workflow

1. Read the requested Markdown file, inline notes, or source material.
2. Choose a concise presentation title and safe lowercase hyphen slug.
3. Create `<topic-slug>/` at the workspace root. If it exists, append `-2`, `-3`, etc.
4. Create `<topic-slug>/content/` and put all deck-specific source material there.
5. Create `<topic-slug>/slides_code/`.
6. Create a task/TODO for each slide before writing slide source.
7. Build one `.ts` file per slide in `<topic-slug>/slides_code/`.
8. Create `<topic-slug>/slides_code/build.ts`.
9. Run `npx tsx <topic-slug>/slides_code/build.ts` to verify PPTX generation.

## Output Contract

```text
<topic-slug>/
  <topic-slug>.pptx
  content/
    source.md
    additional-user-content.ext
  slides_code/
    build.ts
    slide01_DescriptiveName.ts
    slide02_DescriptiveName.ts
```

Use the same slug for the folder and final PPTX filename. Never overwrite an existing topic folder.

Use `<topic-slug>/content/` for every deck-specific user input:

- If the user provides a Markdown file, copy it into `content/` using its original basename unless a normalized `source.md` is clearer.
- If the user provides inline notes instead of a file, save them as `content/source.md`.
- If the user provides images, data files, reference documents, or other supporting material, copy them into `content/`.
- Never move or edit the user's original source files unless explicitly asked.
- If a filename collision occurs inside `content/`, append `-2`, `-3`, etc.

## Slide File Pattern

One file per slide. Export a default function with typed PPTX and theme arguments:

```typescript
import type { PptxDeck, ThemeApi } from "../../theme";

export default function buildSlide(pptx: PptxDeck, theme: ThemeApi): void {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide); // required exactly once per slide
  // theme.addTitle(slide, "Title"); // optional for statement/quote slides
  // Add slide-specific content inside theme.LAYOUT.FREE_*.
}
```

## Layout Contract

- Set `pptx.layout = "LAYOUT_16x9"` in every generated `build.ts`.
- Start each slide with `theme.applySlideBase(slide)` exactly once; it owns the background, logo, footer, and page number.
- Use `theme.addTitle()` for normal page headers.
- Place custom content inside `theme.LAYOUT.FREE_*`.
- Do not create manual footer text, page numbers, source labels, bottom rules, or competing header chrome.
- Do not add new layout helper APIs as part of normal deck generation.

## Typography And Color Contract

- Use existing typography and color helpers from `theme.ts`; do not reimplement standard text styling in slide files.
- Use `theme.addTitle()`, `theme.addSubheader()` / `theme.addSubtitle()`, `theme.addBody()`, `theme.makeTextRun()`, and `theme.makeHighlightRun()` for standard text.
- Main slide headers created with `theme.addTitle()` are Arial, regular weight, dark blue, and may use the theme title size.
- Subheadings created with `theme.addSubheader()` / `theme.addSubtitle()` are Arial Bold, light blue, 15 pt.
- Body text, table text, callout text, and other normal slide text are Arial, black, 15 pt.
- Text highlighting means bold black text only. Use `theme.makeHighlightRun()` for inline emphasis; do not use heading or subheading styling as a highlighter, and do not turn highlighted words light blue.
- Use `theme.COLORS` constants for custom shapes, fills, borders, or manual text overrides.
- Only hand-specify fonts, sizes, or hex colors when a genuinely custom element is not covered by the current theme helpers; normal slide text must stay at 15 pt.
- No emoji.

## Build File Pattern

```typescript
import path from "node:path";
import pptxgen from "pptxgenjs";
import * as theme from "../../theme";
import slide01 from "./slide01_TitleSlide";

const deckSlug = "<topic-slug>";
const outputPath = path.join(__dirname, "..", `${deckSlug}.pptx`);
const logoPath = path.join(__dirname, "..", "..", "assets", "Exyte_RGB.svg");

async function main(): Promise<void> {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "Exyte";
  pptx.company = "Exyte";
  pptx.subject = "<short subject from content>";
  pptx.title = "<title from content>";

  theme.updatePresentationSettings("<title from content>", "<today DD/MM/YYYY>", logoPath);
  theme.resetSlideCounter(0);

  const slides = [slide01];
  slides.forEach((buildSlide) => buildSlide(pptx, theme));

  await pptx.writeFile({ fileName: outputPath });
}

main().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});
```

## Verification

Run:

```bash
npx tsc --noEmit
npx tsx <topic-slug>/slides_code/build.ts
```

Open or inspect the generated PPTX when possible — visually confirm no text is clipped and no shape bleeds off-slide or overlaps the footer/logo. Verify the footer, date, page numbers, and 16:9 slide size. Fix any issues found and rebuild before considering the deck done.
