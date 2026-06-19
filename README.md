# Exyte PPTX Generator

An AI skill for generating Exyte-style PowerPoint presentations from Markdown, notes, or other source material using [pptxgenjs](https://github.com/gitbrent/PptxGenJS).

## Quick Start

1. Install Node.js 22 or newer.
2. Clone and open this repository.
3. (Optional) Copy `exyte-pptx-generator/` into your agent's skills directory (each tool does it slightly differently, so please read the documentation for yours) to add it as a skill for further usage.
4. Run any AI coding agent of your choice (e.g., Codex, Claude, Cursor, OpenCode, etc.) inside the repo.
5. Provide your presentation content to the agent.
6. Let it cook.

Example:

```text
Use $exyte-pptx-generator to create a presentation from example_script.md.
```

For best results, describe the purpose and content of each slide in a script.md file.

I was using Codex with GPT-5.5, but I believe any frontier model in any modern harness can build a decent presentation with this skill.

## How it works

Every slide uses required theme chrome. `theme.applySlideBase()` adds the background, corporate logo, footer, and page number. `theme.addSlideTitle()` separately adds the slide title. Custom text, cards, tables, diagrams, and media must remain fully inside `theme.LAYOUT.FREE_*`.

You control the look and feel through `theme.ts`. It defines the brand colors, fonts, and size presets in one place.

```
+-------------------------------------------------------------------------+
|  Slide Title                                                     [LOGO] |  <- header (title + logo)
|-------------------------------------------------------------------------|
|                                                                         |
|                                                                         |
|                            FREE CONTENT AREA                            |
|                             (AI builds here)                            |
|                                                                         |
|                                                                         |
|-------------------------------------------------------------------------|
|  (c) Exyte | Presentation Title                     Date | Slide Page   |  <- footer
+-------------------------------------------------------------------------+
```

## Self-Contained Deck Projects

The skill initializes a complete project for every deck:

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

Initialization creates the project source files, not the PPTX. The generated `.gitignore` excludes dependencies, PPTX outputs, preview/output directories, and `.DS_Store` while keeping `content/` trackable.

Build and verify the generated deck:

```bash
cd <topic-slug>
npm ci
npm run verify
```

The `<topic-slug>.pptx` file appears when `npm run verify` runs its build step. You can also rebuild it directly with `npm run build`. Each project is self-contained and does not depend on this repository after initialization.

## Theme

The canonical theme lives in `exyte-pptx-generator/starter/theme.ts`. It is copied to the root of each generated deck project upon initialization.

The theme provides:

| Export | Purpose |
|---|---|
| `COLORS` | Exyte palette |
| `TYPOGRAPHY` | Arial role and size presets |
| `LAYOUT` | 13.333 × 7.5 inch safe geometry |
| `CHROME_OBJECT_NAMES` | Stable validator-facing names |
| `configurePresentation(...)` | Deck title, date, and local logo |
| `applySlideBase(slide)` | Background, PNG logo, footer, and page number |
| `addSlideTitle(...)` | Header constrained outside the logo zone |
| `addSubheading(...)` | Standard light-blue subheading |
| `addBodyText(...)` | Standard 15 pt body text |
| `addThemedTable(...)` | Branded editable table |
| `createTextRun(...)` | Standard inline text |
| `createEmphasisRun(...)` | Bold inline emphasis |
| `addCalloutBox(...)` | Branded callout |

## Verification

At repository level:

```bash
npm run verify
```

This typechecks the code, tests the validator and initializer, builds the smoke deck and neutral regression fixture, then validates both PPTX packages.

The validator rejects:

- off-slide or footer-overlapping objects;
- any custom object not fully contained within the exact rounded EMU boundaries derived from `LAYOUT.FREE_X`, `FREE_Y`, `FREE_W`, and `FREE_H`;
- title/logo collisions;
- missing theme chrome;
- non-Arial slide text;
- invalid or mislabeled media;
- SVG image relationships;
- filesystem paths embedded in XML;
- unexpected sibling PPTX files.

After structural verification, render and inspect every slide manually in PowerPoint, Quick Look, or LibreOffice. Check wrapping, clipping, overlaps, alignment, image quality, and overall visual balance; report when rendering is unavailable.

## Repository Structure

```text
.
├── build.ts
├── fixtures/regression-deck/
├── tests/
└── exyte-pptx-generator/
    ├── SKILL.md
    ├── agents/openai.yaml
    ├── scripts/init-deck.ts
    └── starter/
```

## License

[MIT](LICENSE)
