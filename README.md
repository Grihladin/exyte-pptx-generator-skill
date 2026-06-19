# Exyte PPTX Generator

An AI skill for generating editable Exyte-style PowerPoint presentations from Markdown, notes, or other source material.

## Quick Start

1. Clone and open this repository.
2. (Optional) Copy `exyte-pptx-generator/` into your agent's skills directory to add it as a skill for further usage.
3. Provide your presentation content to the agent.
4. Let it cook.

Example:

```text
Use $exyte-pptx-generator to create a presentation from example_script.md.
```

For best results, describe the purpose and content of each slide in a script.md file.

## Self-Contained Deck Projects

The skill initializes a complete project for every deck:

```text
<topic-slug>/
  <topic-slug>.pptx
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

Each deck can be copied elsewhere, installed with `npm ci`, rebuilt with `npm run build`, and checked with `npm run verify`. It does not depend on this repository after initialization.

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
- title/logo collisions;
- missing theme chrome;
- non-Arial slide text;
- invalid or mislabeled media;
- SVG image relationships;
- filesystem paths embedded in XML;
- unexpected sibling PPTX files.

Quick Look or LibreOffice rendering remains an additional visual QA step when available.

## Repository Structure

```text
.
├── build.ts
├── fixtures/regression-deck/
├── tests/
└── exyte-pptx-generator/
    ├── SKILL.md
    ├── agents/openai.yaml
    ├── scripts/init-deck.mjs
    └── starter/
```

## License

[MIT](LICENSE)
