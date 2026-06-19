# pptx-slide-generator

A skill for PowerPoint generation using [pptxgenjs](https://github.com/gitbrent/PptxGenJS), TypeScript, and a corporate theme system.

Write your presentation content in Markdown, point an AI coding assistant at it (eg. codex, claude, opencode and so on...), and get a branded `.pptx` file out -- no PowerPoint required.

I built this during coffee breaks because I have never used PowerPoint and have no intention of ever learning it. 

## Quick Start

1. Clone the repo.
2. Inside the repo run AI Agent.
3. Ask it to use exyte-pptx-generator-skill.
4. Paste a Markdown with content or ask agent to build it with u, if u will have a content in text first it will increase the quality of presentatuion by a mile (i advice to pay attehcion to script it can enchance pptx a lot)
5. wait and done.

```bash
npm install
npm run typecheck
npm run build
```

`npm run build` runs `tsx build.ts` and writes `output.pptx`.

if u want to set it up as a skill in codex or claude follow these steps:

Codex.

Cladue


## How it works

Every slide has a hardcoded header with the corporate logo and a hardcoded footer with the title, date, and page number. These are applied automatically by `theme.applySlideBase()`.
Everything in between is a free content area where the AI model builds the actual slide content: text, cards, tables, diagrams, or whatever fits.

```text
+------------------------------------------+
|  [Logo]                          header  |
|------------------------------------------|
|                                          |
|         free content area                |
|                                          |
|------------------------------------------|
|  (c) Exyte | Title       date | page     |
+------------------------------------------+
```

You control the look and feel through `theme.ts`. It defines the brand colors, fonts, size presets, layout constants, and theme helpers in one place.

## Theme System

Here is a list of rules i defined, feel free to modify them as needed(here we need to balance between waht will be hardocded and how free AI model can generate content so do not overcomplicate ofc i could hardocde look of slides form design template, but i belive modern AI models can genretate better designs):

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

## Dependencies

- [pptxgenjs](https://www.npmjs.com/package/pptxgenjs) -- PowerPoint file generation for Node.js
- TypeScript, tsx, and @types/node -- TypeScript authoring and direct TS execution

## License

MIT
