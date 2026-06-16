import type { PptxDeck, ThemeApi } from "../../theme";

export default function buildSlide(pptx: PptxDeck, theme: ThemeApi): void {
  const slide = pptx.addSlide();
  theme.applySlideBase(slide);
  theme.addTitle(slide, "Research: Can Code Build PPTX?");

  theme.addBody(
    slide,
    [
      theme.makeTextRun("The first automation question was simple: "),
      theme.makeHighlightRun("can we reach PPTX programmatically?"),
      theme.makeTextRun("\nThe answer came quickly: two well-maintained open-source frameworks."),
    ],
    {
      x: theme.LAYOUT.FREE_X,
      y: 1.0,
      w: theme.LAYOUT.FREE_W,
      h: 0.7,
      lineSpacing: 18,
      fit: "shrink",
    },
  );

  theme.addStyledTable(
    slide,
    ["Framework", "Strengths", "Important limitation"],
    [
      [
        { text: "python-pptx\nSteve Canny\nMIT license", opts: { bold: true, color: theme.COLORS.TEXT_PRIMARY } },
        "Deep XML control\nDefine slide design in Python\nCan read and write PPTX files",
        { text: "None for read/edit workflow", opts: { color: theme.COLORS.TEXT_BODY } },
      ],
      [
        { text: "pptxgenjs\nBrent Ely\n1.7M npm downloads", opts: { bold: true, color: theme.COLORS.TEXT_PRIMARY } },
        "Strong tables and charts\nPopular in JavaScript ecosystem\nGood fit for generation pipelines",
        { text: "Create-only: cannot read or edit existing PPTX files", opts: { bold: true, color: theme.COLORS.TEXT_PRIMARY } },
      ],
    ],
    {
      x: theme.LAYOUT.FREE_X,
      y: 1.92,
      w: theme.LAYOUT.FREE_W,
      colW: [2.25, 3.6, 2.55],
      rowH: 0.95,
      bodyFontSize: 10.4,
    },
  );

  theme.addCalloutBox(
    slide,
    [
      theme.makeTextRun("Keep the create-only limitation in mind - it becomes relevant later.", {
        fontSize: 12.5,
      }),
    ],
    { x: theme.LAYOUT.FREE_X, y: 4.55, w: theme.LAYOUT.FREE_W, h: 0.48 },
  );
}
