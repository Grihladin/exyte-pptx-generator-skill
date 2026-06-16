// ============================================================
// Design System — Exyte Corporate Blue Theme
// ============================================================

const COLORS = {
  // Core brand colors
  MID_BLUE: '0085CA',
  LIGHT_BLUE: '00B5E2',
  SECONDARY: '00B5E2',

  // Backgrounds
  BG: 'FFFFFF',
  BG_CARD: 'F2FBFD',
  BG_SURFACE: 'E6F8FC',
  BORDER: 'B6EAF7',

  // Text
  TEXT_PRIMARY: '0085CA',
  TEXT_BODY: '000000',
  TEXT_MUTED: '0085CA',

  // Accent
  ACCENT: '0085CA',
  ACCENT_SECONDARY: '00B5E2',
  ACCENT_SOFT: 'E6F8FC',

  // Utility
  WHITE: 'FFFFFF',
};

const FONTS = {
  HEADING: 'Arial',
  BODY: 'Arial',
  MONO: 'Consolas',
  TITLE_SIZE: 24,
  SUBHEADER_SIZE: 15,
  SUBTITLE_SIZE: 15,
  BODY_SIZE: 13,
  BODY_MIN_SIZE: 12,
  BODY_MAX_SIZE: 15,
  HIGHLIGHT_SIZE: 15,
  SMALL_SIZE: 10,
  TABLE_SIZE: 11,
  TABLE_HEADER_SIZE: 12,
};

// Layout grid (inches, 10x5.63 canvas)
const LAYOUT = {
  SLIDE_W: 10,
  SLIDE_H: 5.625,

  // Chrome (header, logo, footer — managed by theme)
  MARGIN_L: 0.8,
  MARGIN_R: 0.8,
  CONTENT_X: 0.8,
  CONTENT_W: 8.4,
  TITLE_Y: 0.3,
  TITLE_H: 0.45,
  FOOTER_LINE_Y: 5.365,
  FOOTER_LINE_H: 0.03,
  FOOTER_Y: 5.455,
  FOOTER_H: 0.12,
  FOOTER_LEFT_X: 0.1,
  FOOTER_LEFT_W: 5.25,
  FOOTER_DATE_X: 8.55,
  FOOTER_DATE_W: 0.75,
  FOOTER_DIVIDER_X: 9.4,
  FOOTER_DIVIDER_W: 0.08,
  FOOTER_PAGE_X: 9.63,
  FOOTER_PAGE_W: 0.18,
  FOOTER_FONT_SIZE: 7,

  // Free content box — AI can place anything here without restrictions
  FREE_X: 0.8,
  FREE_Y: 1.0,
  FREE_W: 8.4,
  FREE_H: 4.1,
};

let slideCounter = 0;
let TOTAL_SLIDES = 25;

// Presentation-level settings
let presentationTitle = '';
let presentationDate = '';
let logoPath = '';

function updatePresentationSettings(title, date, logo) {
  presentationTitle = title || '';
  presentationDate = date || '';
  logoPath = logo || '';
}

function resetSlideCounter(val) {
  slideCounter = val !== undefined ? val : 0;
}

function formatFooterDate(date) {
  if (!date) return '';

  const input = String(date).trim();

  // ISO-like dates from scripts or metadata.
  const isoMatch = input.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (isoMatch) {
    const [, year, month, day] = isoMatch;
    return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${year}`;
  }

  // Preserve already-correct template format.
  const dmyMatch = input.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmyMatch) {
    const [, first, second, year] = dmyMatch;
    const firstNum = Number(first);
    const secondNum = Number(second);

    if (firstNum > 12 || secondNum <= 12) {
      return `${first.padStart(2, '0')}/${second.padStart(2, '0')}/${year}`;
    }

    return `${second.padStart(2, '0')}/${first.padStart(2, '0')}/${year}`;
  }

  return input;
}

// ── Base slide layout ──────────────────────────────────────

function applySlideBase(slide, opts = {}) {
  slideCounter++;
  const currentSlide = slideCounter;
  const skipFooter = opts.skipFooter || false;

  slide.background = { color: COLORS.BG };

  // Logo (top-right, natural aspect ratio ~2.7:1)
  if (logoPath) {
    slide.addImage({
      path: logoPath,
      x: 8.2,
      y: 0.3,
      w: 1.2,
      h: 0.45,
      sizing: { type: 'contain', w: 1.2, h: 0.45 },
    });
  }

  if (!skipFooter) {
    // Bottom accent line
    slide.addShape('rect', {
      x: 0,
      y: LAYOUT.FOOTER_LINE_Y,
      w: LAYOUT.SLIDE_W,
      h: LAYOUT.FOOTER_LINE_H,
      fill: { color: COLORS.ACCENT },
      line: { color: COLORS.ACCENT, transparency: 100 },
    });

    // © Exyte | {title} (left)
    const footerLeft = presentationTitle ? `\u00A9 Exyte | ${presentationTitle}` : '\u00A9 Exyte';
    slide.addText(footerLeft, {
      x: LAYOUT.FOOTER_LEFT_X,
      y: LAYOUT.FOOTER_Y,
      w: LAYOUT.FOOTER_LEFT_W,
      h: LAYOUT.FOOTER_H,
      fontSize: LAYOUT.FOOTER_FONT_SIZE,
      fontFace: FONTS.BODY,
      color: COLORS.TEXT_MUTED,
      align: 'left',
      valign: 'middle',
    });

    // {date} | {page} (right)
    const footerDate = formatFooterDate(presentationDate);
    slide.addText(footerDate, {
      x: LAYOUT.FOOTER_DATE_X,
      y: LAYOUT.FOOTER_Y,
      w: LAYOUT.FOOTER_DATE_W,
      h: LAYOUT.FOOTER_H,
      fontSize: LAYOUT.FOOTER_FONT_SIZE,
      fontFace: FONTS.BODY,
      color: COLORS.TEXT_MUTED,
      align: 'right',
      valign: 'middle',
    });

    slide.addText('|', {
      x: LAYOUT.FOOTER_DIVIDER_X,
      y: LAYOUT.FOOTER_Y,
      w: LAYOUT.FOOTER_DIVIDER_W,
      h: LAYOUT.FOOTER_H,
      fontSize: LAYOUT.FOOTER_FONT_SIZE,
      fontFace: FONTS.BODY,
      color: COLORS.TEXT_MUTED,
      align: 'center',
      valign: 'middle',
    });

    slide.addText(String(currentSlide), {
      x: LAYOUT.FOOTER_PAGE_X,
      y: LAYOUT.FOOTER_Y,
      w: LAYOUT.FOOTER_PAGE_W,
      h: LAYOUT.FOOTER_H,
      fontSize: LAYOUT.FOOTER_FONT_SIZE,
      fontFace: FONTS.BODY,
      color: COLORS.TEXT_MUTED,
      align: 'center',
      valign: 'middle',
    });
  }
}

// ── Typography helpers ──────────────────────────────────────

function addTitle(slide, text) {
  slide.addText(text, {
    x: LAYOUT.CONTENT_X,
    y: LAYOUT.TITLE_Y,
    w: LAYOUT.CONTENT_W,
    h: LAYOUT.TITLE_H,
    fontSize: FONTS.TITLE_SIZE,
    fontFace: FONTS.HEADING,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
    valign: 'middle',
  });
}

function addSubheader(slide, text, overrides = {}) {
  const {
    x = LAYOUT.FREE_X,
    y = LAYOUT.FREE_Y,
    w = LAYOUT.FREE_W,
    h = 0.3,
    ...textOpts
  } = overrides;

  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontSize: FONTS.SUBHEADER_SIZE,
    fontFace: FONTS.BODY,
    color: COLORS.SECONDARY,
    bold: true,
    valign: 'top',
    ...textOpts,
  });
}

function addSubtitle(slide, text, overrides = {}) {
  return addSubheader(slide, text, overrides);
}

function addBody(slide, textOrArray, overrides = {}) {
  const {
    x = LAYOUT.FREE_X,
    y = LAYOUT.FREE_Y + 0.4,
    w = LAYOUT.FREE_W,
    h = LAYOUT.FREE_H - 0.4,
    ...textOpts
  } = overrides;

  slide.addText(textOrArray, {
    x,
    y,
    w,
    h,
    fontSize: FONTS.BODY_SIZE,
    fontFace: FONTS.BODY,
    color: COLORS.TEXT_BODY,
    valign: 'top',
    lineSpacing: 22,
    ...textOpts,
  });
}

// ── Table helper ────────────────────────────────────────────

function addStyledTable(slide, headers, dataRows, opts = {}) {
  const headerRow = headers.map((h) => ({
    text: h,
    options: {
      bold: true,
      color: COLORS.TEXT_PRIMARY,
      fill: { color: COLORS.BG_SURFACE },
      fontFace: FONTS.BODY,
      fontSize: FONTS.TABLE_HEADER_SIZE,
      align: 'left',
      valign: 'middle',
      margin: [5, 8, 5, 8],
    },
  }));

  const bodyRows = dataRows.map((row, rowIdx) =>
    row.map((cellContent) => {
      const isObj = typeof cellContent === 'object' && cellContent !== null;
      return {
        text: isObj ? cellContent.text : cellContent,
        options: {
          fontFace: FONTS.BODY,
          fontSize: opts.bodyFontSize || FONTS.TABLE_SIZE,
          color: COLORS.TEXT_BODY,
          fill: { color: rowIdx % 2 === 0 ? COLORS.BG : COLORS.BG_CARD },
          align: 'left',
          valign: 'top',
          margin: [5, 8, 5, 8],
          ...(isObj ? cellContent.opts : {}),
        },
      };
    })
  );

  const allRows = [headerRow, ...bodyRows];

  slide.addTable(allRows, {
    x: opts.x || LAYOUT.FREE_X,
    y: opts.y || LAYOUT.FREE_Y + 0.4,
    w: opts.w || LAYOUT.FREE_W,
    colW: opts.colW,
    border: { type: 'solid', pt: 0.5, color: COLORS.BORDER },
    fontFace: FONTS.BODY,
    rowH: opts.rowH,
  });
}

// ── Rich text run helper ────────────────────────────────────

function makeTextRun(text, opts = {}) {
  return {
    text,
    options: {
      fontFace: FONTS.BODY,
      fontSize: FONTS.BODY_SIZE,
      color: COLORS.TEXT_BODY,
      ...opts,
    },
  };
}

function makeHighlightRun(text, opts = {}) {
  return makeTextRun(text, {
    fontSize: FONTS.HIGHLIGHT_SIZE,
    bold: true,
    color: COLORS.SECONDARY,
    ...opts,
  });
}

// ── Callout box helper ──────────────────────────────────────

function addCalloutBox(slide, textRuns, opts = {}) {
  const x = opts.x || LAYOUT.CONTENT_X;
  const y = opts.y || 4.2;
  const w = opts.w || LAYOUT.CONTENT_W;
  const h = opts.h || 0.7;

  // Background
  slide.addShape('roundRect', {
    x: x,
    y: y,
    w: w,
    h: h,
    fill: { color: COLORS.BG_SURFACE },
    rectRadius: 0.05,
  });

  // Text
  slide.addText(textRuns, {
    x: x + 0.15,
    y: y,
    w: w - 0.3,
    h: h,
    fontFace: FONTS.BODY,
    fontSize: opts.fontSize || 12,
    color: COLORS.TEXT_BODY,
    valign: 'middle',
  });
}

module.exports = {
  COLORS,
  FONTS,
  LAYOUT,
  TOTAL_SLIDES,
  updatePresentationSettings,
  applySlideBase,
  addTitle,
  addSubheader,
  addSubtitle,
  addBody,
  addStyledTable,
  makeTextRun,
  makeHighlightRun,
  addCalloutBox,
  resetSlideCounter,
};
