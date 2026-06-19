import type pptxgen from "pptxgenjs";

export type PptxDeck = InstanceType<typeof pptxgen>;
export type Slide = ReturnType<PptxDeck["addSlide"]>;
export type TextRuns = Extract<Parameters<Slide["addText"]>[0], unknown[]>;
export type TextRun = TextRuns[number];
export type TextOptions = NonNullable<Parameters<Slide["addText"]>[1]>;
export type TableRows = Parameters<Slide["addTable"]>[0];
export type TableRow = TableRows[number];
export type TableCellOptions = NonNullable<TableRow[number]["options"]>;
export type TableOptions = NonNullable<Parameters<Slide["addTable"]>[1]>;

export const COLORS = {
  MID_BLUE: "0085CA",
  LIGHT_BLUE: "00B5E2",
  SECONDARY: "00B5E2",
  BG: "FFFFFF",
  BG_CARD: "F2FBFD",
  BG_SURFACE: "E6F8FC",
  BORDER: "B6EAF7",
  TEXT_PRIMARY: "0085CA",
  TEXT_BODY: "000000",
  TEXT_MUTED: "0085CA",
  ACCENT: "0085CA",
  ACCENT_SECONDARY: "00B5E2",
  ACCENT_SOFT: "E6F8FC",
  WHITE: "FFFFFF",
} as const;

export const TYPOGRAPHY = {
  HEADING: "Arial",
  BODY: "Arial",
  MONO: "Consolas",
  TITLE_SIZE: 24,
  SUBHEADING_SIZE: 15,
  BODY_SIZE: 15,
  BODY_MIN_SIZE: 15,
  BODY_MAX_SIZE: 15,
  EMPHASIS_SIZE: 15,
  SMALL_SIZE: 15,
  TABLE_SIZE: 15,
  TABLE_HEADER_SIZE: 15,
} as const;

export const PRESENTATION_LANGUAGE = "en-US";

// PowerPoint Widescreen in inches for pptx.layout = "LAYOUT_WIDE".
export const LAYOUT = {
  SLIDE_W: 13.333,
  SLIDE_H: 7.5,

  MARGIN_L: 0.8,
  MARGIN_R: 0.8,
  CONTENT_X: 0.8,
  CONTENT_W: 11.733,

  TITLE_Y: 0.38,
  TITLE_H: 0.5,
  TITLE_W: 9.9,

  LOGO_X: 11.053,
  LOGO_Y: 0.3,
  LOGO_W: 1.48,
  LOGO_H: 0.595,
  LOGO_GAP: 0.35,

  FOOTER_LINE_Y: 7.12,
  FOOTER_LINE_H: 0.03,
  FOOTER_Y: 7.25,
  FOOTER_H: 0.14,
  FOOTER_LEFT_X: 0.13,
  FOOTER_LEFT_W: 7.0,
  FOOTER_DATE_X: 11.15,
  FOOTER_DATE_W: 0.9,
  FOOTER_DIVIDER_X: 12.22,
  FOOTER_DIVIDER_W: 0.08,
  FOOTER_PAGE_X: 12.53,
  FOOTER_PAGE_W: 0.25,
  FOOTER_FONT_SIZE: 7,

  FREE_X: 0.8,
  FREE_Y: 1.25,
  FREE_W: 11.733,
  FREE_H: 5.55,
} as const;

export const CHROME_OBJECT_NAMES = {
  LOGO: "Exyte Logo",
  FOOTER_LINE: "Exyte Footer Line",
  FOOTER_LEFT: "Exyte Footer Left",
  FOOTER_DATE: "Exyte Footer Date",
  FOOTER_DIVIDER: "Exyte Footer Divider",
  FOOTER_PAGE: "Exyte Footer Page",
  SLIDE_TITLE: "Exyte Slide Title",
} as const;

export interface ApplySlideBaseOptions {
  skipFooter?: boolean;
}

export interface PresentationConfig {
  title?: string;
  date?: string;
  logoPath?: string;
}

export interface ThemedTableCell {
  text: string;
  options?: Partial<TableCellOptions>;
}

export interface ThemedTableOptions {
  x?: number;
  y?: number;
  w?: number;
  colW?: TableOptions["colW"];
  rowH?: TableOptions["rowH"];
  bodyFontSize?: number;
}

export interface CalloutBoxOptions {
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  fontSize?: number;
}

let slideCounter = 0;
let presentationTitle = "";
let presentationDate = "";
let logoPath = "";

export function configurePresentation(
  titleOrConfig: string | PresentationConfig,
  date = "",
  logo = "",
): void {
  if (typeof titleOrConfig === "object") {
    presentationTitle = titleOrConfig.title ?? "";
    presentationDate = titleOrConfig.date ?? "";
    logoPath = titleOrConfig.logoPath ?? "";
    return;
  }

  presentationTitle = titleOrConfig || "";
  presentationDate = date || "";
  logoPath = logo || "";
}

export function resetSlideCounter(value = 0): void {
  slideCounter = value;
}

export function formatFooterDate(date: string | Date | undefined | null): string {
  if (!date) return "";

  if (date instanceof Date) {
    const day = String(date.getDate()).padStart(2, "0");
    const month = String(date.getMonth() + 1).padStart(2, "0");
    return `${day}/${month}/${date.getFullYear()}`;
  }

  const input = String(date).trim();
  const isoMatch = input.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (isoMatch) {
    const [, year, month, day] = isoMatch;
    return `${day.padStart(2, "0")}/${month.padStart(2, "0")}/${year}`;
  }

  const dmyMatch = input.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmyMatch) {
    const [, day, month, year] = dmyMatch;
    return `${day.padStart(2, "0")}/${month.padStart(2, "0")}/${year}`;
  }

  return input;
}

export function applySlideBase(slide: Slide, options: ApplySlideBaseOptions = {}): void {
  slideCounter += 1;
  const currentSlide = slideCounter;

  slide.background = { color: COLORS.BG };

  if (logoPath) {
    slide.addImage({
      path: logoPath,
      x: LAYOUT.LOGO_X,
      y: LAYOUT.LOGO_Y,
      w: LAYOUT.LOGO_W,
      h: LAYOUT.LOGO_H,
      altText: "Exyte logo",
      objectName: CHROME_OBJECT_NAMES.LOGO,
    });
  }

  if (options.skipFooter) return;

  slide.addShape("rect", {
    x: 0,
    y: LAYOUT.FOOTER_LINE_Y,
    w: LAYOUT.SLIDE_W,
    h: LAYOUT.FOOTER_LINE_H,
    fill: { color: COLORS.ACCENT },
    line: { color: COLORS.ACCENT, transparency: 100 },
    objectName: CHROME_OBJECT_NAMES.FOOTER_LINE,
  });

  const footerLeft = presentationTitle ? `\u00A9 Exyte | ${presentationTitle}` : "\u00A9 Exyte";
  slide.addText(footerLeft, {
    x: LAYOUT.FOOTER_LEFT_X,
    y: LAYOUT.FOOTER_Y,
    w: LAYOUT.FOOTER_LEFT_W,
    h: LAYOUT.FOOTER_H,
    fontSize: LAYOUT.FOOTER_FONT_SIZE,
    fontFace: TYPOGRAPHY.BODY,
    color: COLORS.TEXT_MUTED,
    align: "left",
    valign: "middle",
    margin: 0,
    lang: PRESENTATION_LANGUAGE,
    objectName: CHROME_OBJECT_NAMES.FOOTER_LEFT,
  });

  slide.addText(formatFooterDate(presentationDate), {
    x: LAYOUT.FOOTER_DATE_X,
    y: LAYOUT.FOOTER_Y,
    w: LAYOUT.FOOTER_DATE_W,
    h: LAYOUT.FOOTER_H,
    fontSize: LAYOUT.FOOTER_FONT_SIZE,
    fontFace: TYPOGRAPHY.BODY,
    color: COLORS.TEXT_MUTED,
    align: "right",
    valign: "middle",
    margin: 0,
    lang: PRESENTATION_LANGUAGE,
    objectName: CHROME_OBJECT_NAMES.FOOTER_DATE,
  });

  slide.addText("|", {
    x: LAYOUT.FOOTER_DIVIDER_X,
    y: LAYOUT.FOOTER_Y,
    w: LAYOUT.FOOTER_DIVIDER_W,
    h: LAYOUT.FOOTER_H,
    fontSize: LAYOUT.FOOTER_FONT_SIZE,
    fontFace: TYPOGRAPHY.BODY,
    color: COLORS.TEXT_MUTED,
    align: "center",
    valign: "middle",
    margin: 0,
    lang: PRESENTATION_LANGUAGE,
    objectName: CHROME_OBJECT_NAMES.FOOTER_DIVIDER,
  });

  slide.addText(String(currentSlide), {
    x: LAYOUT.FOOTER_PAGE_X,
    y: LAYOUT.FOOTER_Y,
    w: LAYOUT.FOOTER_PAGE_W,
    h: LAYOUT.FOOTER_H,
    fontSize: LAYOUT.FOOTER_FONT_SIZE,
    fontFace: TYPOGRAPHY.BODY,
    color: COLORS.TEXT_MUTED,
    align: "center",
    valign: "middle",
    margin: 0,
    lang: PRESENTATION_LANGUAGE,
    objectName: CHROME_OBJECT_NAMES.FOOTER_PAGE,
  });
}

export function addSlideTitle(slide: Slide, text: string): void {
  slide.addText(text, {
    x: LAYOUT.CONTENT_X,
    y: LAYOUT.TITLE_Y,
    w: LAYOUT.TITLE_W,
    h: LAYOUT.TITLE_H,
    fontSize: TYPOGRAPHY.TITLE_SIZE,
    fontFace: TYPOGRAPHY.HEADING,
    color: COLORS.TEXT_PRIMARY,
    bold: false,
    valign: "middle",
    margin: 0,
    fit: "shrink",
    lang: PRESENTATION_LANGUAGE,
    objectName: CHROME_OBJECT_NAMES.SLIDE_TITLE,
  });
}

export function addSubheading(slide: Slide, text: string, overrides: Partial<TextOptions> = {}): void {
  const { x = LAYOUT.FREE_X, y = LAYOUT.FREE_Y, w = LAYOUT.FREE_W, h = 0.3, ...textOpts } = overrides;

  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontSize: TYPOGRAPHY.SUBHEADING_SIZE,
    fontFace: TYPOGRAPHY.BODY,
    color: COLORS.LIGHT_BLUE,
    bold: true,
    valign: "top",
    lang: PRESENTATION_LANGUAGE,
    ...textOpts,
  });
}

export function addBodyText(slide: Slide, textOrRuns: string | TextRuns, overrides: Partial<TextOptions> = {}): void {
  const {
    x = LAYOUT.FREE_X,
    y = LAYOUT.FREE_Y + 0.4,
    w = LAYOUT.FREE_W,
    h = LAYOUT.FREE_H - 0.4,
    ...textOpts
  } = overrides;

  slide.addText(textOrRuns, {
    x,
    y,
    w,
    h,
    fontSize: TYPOGRAPHY.BODY_SIZE,
    fontFace: TYPOGRAPHY.BODY,
    color: COLORS.TEXT_BODY,
    valign: "top",
    lineSpacing: 22,
    lang: PRESENTATION_LANGUAGE,
    ...textOpts,
  });
}

export function addThemedTable(
  slide: Slide,
  headers: string[],
  dataRows: Array<Array<string | ThemedTableCell>>,
  options: ThemedTableOptions = {},
): void {
  const headerRow: TableRow = headers.map((header) => ({
    text: header,
    options: {
      bold: true,
      color: COLORS.TEXT_PRIMARY,
      fill: { color: COLORS.BG_SURFACE },
      fontFace: TYPOGRAPHY.BODY,
      fontSize: TYPOGRAPHY.TABLE_HEADER_SIZE,
      align: "left",
      valign: "middle",
      margin: [5, 8, 5, 8],
    },
  }));

  const bodyRows: TableRows = dataRows.map((row, rowIdx) =>
    row.map((cellContent) => {
      const cell = typeof cellContent === "string" ? { text: cellContent } : cellContent;
      return {
        text: cell.text,
        options: {
          fontFace: TYPOGRAPHY.BODY,
          fontSize: options.bodyFontSize ?? TYPOGRAPHY.TABLE_SIZE,
          color: COLORS.TEXT_BODY,
          fill: { color: rowIdx % 2 === 0 ? COLORS.BG : COLORS.BG_CARD },
          align: "left",
          valign: "top",
          margin: [5, 8, 5, 8],
          ...cell.options,
        },
      };
    }),
  );

  slide.addTable([headerRow, ...bodyRows], {
    x: options.x ?? LAYOUT.FREE_X,
    y: options.y ?? LAYOUT.FREE_Y + 0.4,
    w: options.w ?? LAYOUT.FREE_W,
    colW: options.colW,
    border: { type: "solid", pt: 0.5, color: COLORS.BORDER },
    fontFace: TYPOGRAPHY.BODY,
    rowH: options.rowH,
  });
}

export function createTextRun(text: string, options: Partial<TextOptions> = {}): TextRun {
  return {
    text,
    options: {
      fontFace: TYPOGRAPHY.BODY,
      fontSize: TYPOGRAPHY.BODY_SIZE,
      color: COLORS.TEXT_BODY,
      lang: PRESENTATION_LANGUAGE,
      ...options,
    },
  };
}

export function createEmphasisRun(text: string, options: Partial<TextOptions> = {}): TextRun {
  return createTextRun(text, {
    bold: true,
    color: COLORS.TEXT_BODY,
    ...options,
  });
}

export function addCalloutBox(slide: Slide, textRuns: TextRuns, options: CalloutBoxOptions = {}): void {
  const x = options.x ?? LAYOUT.CONTENT_X;
  const y = options.y ?? 5.85;
  const w = options.w ?? LAYOUT.CONTENT_W;
  const h = options.h ?? 0.75;

  slide.addShape("roundRect", {
    x,
    y,
    w,
    h,
    fill: { color: COLORS.BG_SURFACE },
    rectRadius: 0.05,
  });

  slide.addText(textRuns, {
    x: x + 0.15,
    y,
    w: w - 0.3,
    h,
    fontFace: TYPOGRAPHY.BODY,
    fontSize: options.fontSize ?? TYPOGRAPHY.BODY_SIZE,
    color: COLORS.TEXT_BODY,
    valign: "middle",
    lang: PRESENTATION_LANGUAGE,
  });
}

export interface ThemeApi {
  COLORS: typeof COLORS;
  TYPOGRAPHY: typeof TYPOGRAPHY;
  PRESENTATION_LANGUAGE: typeof PRESENTATION_LANGUAGE;
  LAYOUT: typeof LAYOUT;
  CHROME_OBJECT_NAMES: typeof CHROME_OBJECT_NAMES;
  configurePresentation: typeof configurePresentation;
  resetSlideCounter: typeof resetSlideCounter;
  applySlideBase: typeof applySlideBase;
  addSlideTitle: typeof addSlideTitle;
  addSubheading: typeof addSubheading;
  addBodyText: typeof addBodyText;
  addThemedTable: typeof addThemedTable;
  createTextRun: typeof createTextRun;
  createEmphasisRun: typeof createEmphasisRun;
  addCalloutBox: typeof addCalloutBox;
}
