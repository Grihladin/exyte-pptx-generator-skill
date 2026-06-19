# pptx-slide-generator

A prototype for AI-driven PowerPoint generation using [pptxgenjs](https://github.com/gitbrent/PptxGenJS), TypeScript, and a corporate theme system.

I built this during coffee breaks because I have never used PowerPoint and have no intention of ever learning it. Write your presentation content in Markdown, point an AI coding assistant at it, and get a branded `.pptx` file out -- no PowerPoint required.

## The Idea

Every slide has a hardcoded header with the corporate logo and a hardcoded footer with the title, date, and page number. These are applied automatically by `theme.applySlideBase()`. Everything in between is a free content area where the AI model builds the actual slide content: text, cards, tables, diagrams, or whatever fits.

You control the look and feel through `theme.ts`. It defines the brand colors, fonts, size presets, layout constants, and theme helpers in one place.

```text
+------------------------------------------+
|  [Logo]                          header   |
|------------------------------------------|
|                                          |
|         free content area                |
|                                          |
|------------------------------------------|
|  (c) Exyte | Title       date | page     |
+------------------------------------------+
```

## How It Works

```text
Markdown/user files -> deck content/ -> slide01.ts, slide02.ts, ... -> build.ts -> output.pptx
```

1. User source material for each deck lives in that deck's `<topic-slug>/content/` folder.
2. Generated slide files are TypeScript modules that receive a `pptxgenjs` instance and the theme.
3. `theme.ts` provides the Exyte Corporate Blue design system.
4. `build.ts` configures presentation metadata, runs slide builders, and writes the PPTX.

## Quick Start

```bash
npm install
npm run typecheck
npm run build
```

`npm run build` runs `tsx build.ts` and writes `output.pptx`.

## Skill

The portable Agent Skill lives in `exyte-pptx-generator/SKILL.md`. Use that file for Claude or Codex workflows. `PPTX_SKILL.md` is only a compatibility pointer for older workflows.

## Project Structure

```text
.
|-- assets/
|   `-- Exyte_RGB.svg
|-- exyte-pptx-generator/
|   |-- SKILL.md
|   `-- agents/openai.yaml
|-- build.ts
|-- theme.ts
|-- tsconfig.json
`-- PPTX_SKILL.md
```

Generated presentations should own their source material:

```text
<topic-slug>/
|-- <topic-slug>.pptx
|-- content/
|   |-- source.md
|   `-- additional-user-content.ext
`-- slides_code/
    |-- build.ts
    `-- slide01_DescriptiveName.ts
```

## Theme System

`theme.ts` exports constants, types, and helpers used by generated slides:

| Export | Purpose |
|---|---|
| `COLORS` | Exyte brand palette |
| `FONTS` | Font families and size presets |
| `LAYOUT` | Safe content area coordinates |
| `applySlideBase(slide)` | Applies background, logo, footer, and page number |
| `addTitle(slide, text)` | Standard title text |
| `addSubtitle(slide, text)` | Alias for subheader text |
| `addBody(slide, text, overrides?)` | Body text block |
| `addStyledTable(slide, headers, rows, opts?)` | Themed table |
| `makeTextRun(text, opts?)` | Inline styled text run |
| `makeHighlightRun(text, opts?)` | Highlighted inline text run |
| `addCalloutBox(slide, textRuns, opts?)` | Highlighted callout box |

## Writing A Slide

Generated slides should follow this TypeScript pattern:

```typescript
import type { PptxDeck, ThemeApi } from "../../theme";

export default function buildSlide(pptx: PptxDeck, theme: ThemeApi): void {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addTitle(slide, "My Title");

  theme.addBody(slide, "Hello", {
    x: theme.LAYOUT.FREE_X,
    y: theme.LAYOUT.FREE_Y,
    w: theme.LAYOUT.FREE_W,
    h: 1,
  });
}
```

## Dependencies

- [pptxgenjs](https://www.npmjs.com/package/pptxgenjs) -- PowerPoint file generation for Node.js
- TypeScript, tsx, and @types/node -- TypeScript authoring and direct TS execution

## License

MIT
