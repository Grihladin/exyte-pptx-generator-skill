# Exyte PPTX Generator

An AI skill for generating Exyte-branded PowerPoint presentations from Markdown, notes, or other source material.

It uses [PptxGenJS](https://github.com/gitbrent/PptxGenJS), TypeScript, and a reusable corporate theme. Give the skill and your content to an AI coding agent, ask it to build the presentation, and get a ready-to-use `.pptx` file.

I built this during coffee breaks because I have never used PowerPoint and have no intention of starting now.

## Quick Start

1. Clone the repository and open it in your AI coding agent.
2. Give the agent your presentation content or ask it to help you write it.
3. Ask the agent to read `exyte-pptx-generator/SKILL.md` and build the presentation.
4. Review the generated `.pptx` file.

Example prompt:

```text
Use the exyte-pptx-generator skill to create a presentation from example_script.md
```

For the best results, provide a presentation script with the purpose and content of each slide described separately. The quality of the source material has a major impact on the quality of the final deck.

I used the skill with GPT-5.5, but any capable frontier coding model should be capable of generating a presentation.

## Adding it as a proper skill

The skill is not tied to a specific AI agent.

If your agent supports reusable skills, copy the entire `exyte-pptx-generator/` folder into its skills directory. Installation paths and naming conventions vary between agents, so follow the documentation for your tool.

## How It Works

Every slide receives the shared background, logo, footer, date, and page number through `theme.applySlideBase()`. Everything in between is a free content area where the AI agent builds the slide-specific layout.

```text
+------------------------------------------+
|  slide title                      [logo]  |
|                                          |
|         free content area                |
|                                          |
|------------------------------------------|
|  © Exyte | Title          date | page    |
+------------------------------------------+
```

Each generated deck is placed in a separate project folder:

```text
<topic-slug>/
  <topic-slug>.pptx
  content/
    source.md
  slides_code/
    build.ts
    slide01_DescriptiveName.ts
    slide02_DescriptiveName.ts
```

Keeping the content, generated slide code, and final presentation together makes each deck easy to inspect, edit, and rebuild.

## Theme System

`theme.ts` centralizes the design rules while leaving the main slide canvas flexible. The corporate elements stay consistent, while the AI agent remains free to choose the layout best suited to each slide.

The current rules are based on the Exyte corporate design template. Feel free to extend or adapt them for your needs.

| Export | Purpose |
|---|---|
| `COLORS` | Exyte brand palette |
| `TYPOGRAPHY` | Font families and size presets |
| `LAYOUT` | Slide dimensions and safe content coordinates |
| `configurePresentation(...)` | Sets the deck title, date, and logo |
| `applySlideBase(slide)` | Applies the background, logo, footer, and page number |
| `addSlideTitle(slide, text)` | Adds the standard slide title |
| `addSubheading(slide, text, options?)` | Adds a themed subheading |
| `addBodyText(slide, text, options?)` | Adds a standard body-text block |
| `addThemedTable(slide, headers, rows, options?)` | Adds a table using the shared theme |
| `createTextRun(text, options?)` | Creates a standard inline text run |
| `createEmphasisRun(text, options?)` | Creates a bold inline emphasis run |
| `addCalloutBox(slide, textRuns, options?)` | Adds a themed callout box |

## Local Setup and Smoke Test

Install the dependencies:

```bash
npm install
```

Check the TypeScript code and generate the example presentation:

```bash
npm run typecheck
npm run build
```

The build command runs `build.ts` and writes `output.pptx` to the repository root.

## Project Structure

```text
.
├── build.ts                         # Smoke-test deck
├── theme.ts                         # Shared presentation theme
├── package.json                     # Runtime and build commands
└── exyte-pptx-generator/
    ├── SKILL.md                     # Instructions followed by the AI agent
    ├── agents/openai.yaml           # Optional agent UI metadata
    └── exyte_logo.svg               # Presentation logo
```

## Dependencies

- [PptxGenJS](https://www.npmjs.com/package/pptxgenjs) for PowerPoint generation
- TypeScript, `tsx`, and `@types/node` for authoring and executing TypeScript

## License

[MIT](LICENSE)
