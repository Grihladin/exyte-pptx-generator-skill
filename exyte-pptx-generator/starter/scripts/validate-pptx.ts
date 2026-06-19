import fs from "node:fs/promises";
import path from "node:path";
import { inflateSync } from "node:zlib";
import JSZip from "jszip";
import { CHROME_OBJECT_NAMES, LAYOUT } from "../theme";

export const EXPECTED_SLIDE_WIDTH_EMU = 12_192_000;
export const EXPECTED_SLIDE_HEIGHT_EMU = 6_858_000;
export const EMU_PER_INCH = 914_400;

const FOOTER_LINE_Y_EMU = Math.round(LAYOUT.FOOTER_LINE_Y * EMU_PER_INCH);
const TITLE_LOGO_GAP_EMU = Math.round(LAYOUT.LOGO_GAP * EMU_PER_INCH);
const REQUIRED_CHROME = [
  CHROME_OBJECT_NAMES.LOGO,
  CHROME_OBJECT_NAMES.FOOTER_LINE,
  CHROME_OBJECT_NAMES.FOOTER_LEFT,
  CHROME_OBJECT_NAMES.FOOTER_DATE,
  CHROME_OBJECT_NAMES.FOOTER_DIVIDER,
  CHROME_OBJECT_NAMES.FOOTER_PAGE,
] as const;

const ABSOLUTE_PATH_PATTERNS = [
  /file:\/\/\//i,
  /(?:^|[\s"'=])[A-Za-z]:[\\/]/m,
  /(?:^|[\s"'=])\/(?:Users|home|private|tmp|Volumes|var\/folders)\//im,
];

export interface ValidationOptions {
  expectedSlides?: number;
  checkSiblingOutputs?: boolean;
}

export interface ValidationResult {
  errors: string[];
  warnings: string[];
  slideCount: number;
}

interface SlideObject {
  name: string;
  xml: string;
  x: number;
  y: number;
  w: number;
  h: number;
}

function decodeXml(value: string): string {
  return value
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&");
}

function attribute(xml: string, name: string): string {
  const match = xml.match(new RegExp(`\\b${name}="([^"]*)"`));
  return match ? decodeXml(match[1]) : "";
}

function parseSlideObjects(xml: string): SlideObject[] {
  const objects: SlideObject[] = [];
  const blockPattern = /<p:(sp|pic|graphicFrame)\b[\s\S]*?<\/p:\1>/g;

  for (const match of xml.matchAll(blockPattern)) {
    const block = match[0];
    const cNvPr = block.match(/<p:cNvPr\b[^>]*>/)?.[0] ?? "";
    const transform = block.match(
      /<[ap]:xfrm\b[^>]*>[\s\S]*?<a:off\b[^>]*\bx="(-?\d+)"[^>]*\by="(-?\d+)"[^>]*\/>[\s\S]*?<a:ext\b[^>]*\bcx="(-?\d+)"[^>]*\bcy="(-?\d+)"[^>]*\/>/,
    );
    if (!transform) continue;

    const rawX = Number(transform[1]);
    const rawY = Number(transform[2]);
    const rawW = Number(transform[3]);
    const rawH = Number(transform[4]);
    objects.push({
      name: attribute(cNvPr, "name"),
      xml: block,
      x: Math.min(rawX, rawX + rawW),
      y: Math.min(rawY, rawY + rawH),
      w: Math.abs(rawW),
      h: Math.abs(rawH),
    });
  }

  return objects;
}

function rectanglesOverlap(a: SlideObject, b: SlideObject): boolean {
  return a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
}

function isPng(buffer: Buffer): boolean {
  return buffer.length >= 24 && buffer.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
}

function isJpeg(buffer: Buffer): boolean {
  return buffer.length >= 3 && buffer[0] === 0xff && buffer[1] === 0xd8 && buffer[2] === 0xff;
}

function isGif(buffer: Buffer): boolean {
  const signature = buffer.subarray(0, 6).toString("ascii");
  return signature === "GIF87a" || signature === "GIF89a";
}

function isWebp(buffer: Buffer): boolean {
  return (
    buffer.length >= 12 &&
    buffer.subarray(0, 4).toString("ascii") === "RIFF" &&
    buffer.subarray(8, 12).toString("ascii") === "WEBP"
  );
}

function pngDimensions(buffer: Buffer): { width: number; height: number } | null {
  if (!isPng(buffer)) return null;
  return {
    width: buffer.readUInt32BE(16),
    height: buffer.readUInt32BE(20),
  };
}

function paethPredictor(left: number, above: number, upperLeft: number): number {
  const estimate = left + above - upperLeft;
  const leftDistance = Math.abs(estimate - left);
  const aboveDistance = Math.abs(estimate - above);
  const upperLeftDistance = Math.abs(estimate - upperLeft);
  if (leftDistance <= aboveDistance && leftDistance <= upperLeftDistance) return left;
  if (aboveDistance <= upperLeftDistance) return above;
  return upperLeft;
}

function pngHasTransparentPixel(buffer: Buffer): boolean {
  if (!isPng(buffer)) return false;

  let offset = 8;
  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = 0;
  let interlace = 0;
  const imageData: Buffer[] = [];

  while (offset + 12 <= buffer.length) {
    const length = buffer.readUInt32BE(offset);
    const type = buffer.subarray(offset + 4, offset + 8).toString("ascii");
    const data = buffer.subarray(offset + 8, offset + 8 + length);
    if (type === "IHDR") {
      width = data.readUInt32BE(0);
      height = data.readUInt32BE(4);
      bitDepth = data[8];
      colorType = data[9];
      interlace = data[12];
    } else if (type === "IDAT") {
      imageData.push(data);
    } else if (type === "IEND") {
      break;
    }
    offset += length + 12;
  }

  if (width === 0 || height === 0 || bitDepth !== 8 || colorType !== 6 || interlace !== 0) {
    return false;
  }

  const bytesPerPixel = 4;
  const rowLength = width * bytesPerPixel;
  const inflated = inflateSync(Buffer.concat(imageData));
  const previous = Buffer.alloc(rowLength);
  const current = Buffer.alloc(rowLength);
  let sourceOffset = 0;

  for (let row = 0; row < height; row += 1) {
    const filter = inflated[sourceOffset];
    sourceOffset += 1;
    for (let column = 0; column < rowLength; column += 1) {
      const raw = inflated[sourceOffset + column];
      const left = column >= bytesPerPixel ? current[column - bytesPerPixel] : 0;
      const above = previous[column];
      const upperLeft = column >= bytesPerPixel ? previous[column - bytesPerPixel] : 0;
      if (filter === 0) current[column] = raw;
      else if (filter === 1) current[column] = (raw + left) & 0xff;
      else if (filter === 2) current[column] = (raw + above) & 0xff;
      else if (filter === 3) current[column] = (raw + Math.floor((left + above) / 2)) & 0xff;
      else if (filter === 4) current[column] = (raw + paethPredictor(left, above, upperLeft)) & 0xff;
      else return false;
    }
    for (let alpha = 3; alpha < rowLength; alpha += bytesPerPixel) {
      if (current[alpha] < 255) return true;
    }
    current.copy(previous);
    sourceOffset += rowLength;
  }

  return false;
}

function relationshipTargets(xml: string): Map<string, string> {
  const relationships = new Map<string, string>();
  for (const match of xml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const id = attribute(match[0], "Id");
    const target = attribute(match[0], "Target");
    if (id && target) relationships.set(id, target);
  }
  return relationships;
}

async function readZipText(zip: JSZip, name: string): Promise<string> {
  const entry = zip.file(name);
  if (!entry) throw new Error(`Missing required PPTX part: ${name}`);
  return entry.async("string");
}

function slideNumber(name: string): number {
  return Number(name.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
}

export async function validatePptx(
  filePath: string,
  options: ValidationOptions = {},
): Promise<ValidationResult> {
  const errors: string[] = [];
  const warnings: string[] = [];
  const resolvedPath = path.resolve(filePath);

  let bytes: Buffer;
  try {
    bytes = await fs.readFile(resolvedPath);
  } catch {
    return { errors: [`PPTX does not exist: ${resolvedPath}`], warnings, slideCount: 0 };
  }
  if (bytes.length === 0) {
    return { errors: [`PPTX is empty: ${resolvedPath}`], warnings, slideCount: 0 };
  }

  if (options.checkSiblingOutputs !== false) {
    const siblings = (await fs.readdir(path.dirname(resolvedPath))).filter((name) =>
      name.toLowerCase().endsWith(".pptx"),
    );
    if (siblings.length !== 1 || siblings[0] !== path.basename(resolvedPath)) {
      errors.push(
        `Expected exactly one PPTX named ${path.basename(resolvedPath)} beside the build files; found: ${
          siblings.join(", ") || "none"
        }`,
      );
    }
  }

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  } catch (error) {
    return {
      errors: [`PPTX archive is invalid: ${error instanceof Error ? error.message : String(error)}`],
      warnings,
      slideCount: 0,
    };
  }

  const requiredParts = ["[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"];
  for (const requiredPart of requiredParts) {
    if (!zip.file(requiredPart)) errors.push(`Missing required PPTX part: ${requiredPart}`);
  }

  let presentationXml = "";
  try {
    presentationXml = await readZipText(zip, "ppt/presentation.xml");
  } catch (error) {
    errors.push(error instanceof Error ? error.message : String(error));
  }

  const sizeMatch = presentationXml.match(/<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
  if (!sizeMatch) {
    errors.push("Presentation does not declare a slide size.");
  } else if (
    Number(sizeMatch[1]) !== EXPECTED_SLIDE_WIDTH_EMU ||
    Number(sizeMatch[2]) !== EXPECTED_SLIDE_HEIGHT_EMU
  ) {
    errors.push(
      `Slide size must be ${EXPECTED_SLIDE_WIDTH_EMU} × ${EXPECTED_SLIDE_HEIGHT_EMU} EMU ` +
        `(13.333 × 7.5 in); found ${sizeMatch[1]} × ${sizeMatch[2]}.`,
    );
  }

  const slideNames = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => slideNumber(a) - slideNumber(b));
  const slideCount = slideNames.length;
  const presentationSlideCount = (presentationXml.match(/<p:sldId\b/g) ?? []).length;
  if (presentationSlideCount !== slideCount) {
    errors.push(
      `Presentation references ${presentationSlideCount} slides but the archive contains ${slideCount} slide parts.`,
    );
  }
  if (options.expectedSlides !== undefined && slideCount !== options.expectedSlides) {
    errors.push(`Expected ${options.expectedSlides} slides; found ${slideCount}.`);
  }

  const mediaNames = Object.keys(zip.files).filter((name) => name.startsWith("ppt/media/") && !name.endsWith("/"));
  for (const mediaName of mediaNames) {
    const extension = path.extname(mediaName).toLowerCase();
    const media = await zip.file(mediaName)!.async("nodebuffer");
    if (extension === ".svg") {
      errors.push(`SVG media is prohibited: ${mediaName}`);
    } else if (extension === ".png" && !isPng(media)) {
      errors.push(`PNG extension does not match file signature: ${mediaName}`);
    } else if ((extension === ".jpg" || extension === ".jpeg") && !isJpeg(media)) {
      errors.push(`JPEG extension does not match file signature: ${mediaName}`);
    } else if (extension === ".gif" && !isGif(media)) {
      errors.push(`GIF extension does not match file signature: ${mediaName}`);
    } else if (extension === ".webp" && !isWebp(media)) {
      errors.push(`WebP extension does not match file signature: ${mediaName}`);
    }
  }

  for (const [index, slideName] of slideNames.entries()) {
    const slideLabel = `Slide ${index + 1}`;
    const slideXml = await readZipText(zip, slideName);
    const objects = parseSlideObjects(slideXml);

    for (const object of objects) {
      if (object.x < 0 || object.y < 0) {
        errors.push(`${slideLabel}: "${object.name || "unnamed object"}" has negative geometry.`);
      }
      if (
        object.x + object.w > EXPECTED_SLIDE_WIDTH_EMU ||
        object.y + object.h > EXPECTED_SLIDE_HEIGHT_EMU
      ) {
        errors.push(`${slideLabel}: "${object.name || "unnamed object"}" extends off-slide.`);
      }
      if (
        !REQUIRED_CHROME.includes(object.name as (typeof REQUIRED_CHROME)[number]) &&
        object.y + object.h > FOOTER_LINE_Y_EMU
      ) {
        errors.push(`${slideLabel}: "${object.name || "unnamed object"}" intrudes into the footer area.`);
      }
      if (/<a:t>[\s\S]*?<\/a:t>/.test(object.xml) && !/typeface="Arial"/.test(object.xml)) {
        errors.push(`${slideLabel}: "${object.name || "unnamed text"}" does not explicitly use Arial.`);
      }
    }

    for (const requiredName of REQUIRED_CHROME) {
      const count = objects.filter((object) => object.name === requiredName).length;
      if (count !== 1) {
        errors.push(`${slideLabel}: expected one "${requiredName}" object; found ${count}.`);
      }
    }

    const logo = objects.find((object) => object.name === CHROME_OBJECT_NAMES.LOGO);
    const title = objects.find((object) => object.name === CHROME_OBJECT_NAMES.SLIDE_TITLE);
    if (logo && title) {
      if (rectanglesOverlap(logo, title) || title.x + title.w + TITLE_LOGO_GAP_EMU > logo.x) {
        errors.push(`${slideLabel}: slide title enters the reserved logo zone.`);
      }
    }

    if (logo) {
      const cNvPr = logo.xml.match(/<p:cNvPr\b[^>]*>/)?.[0] ?? "";
      if (attribute(cNvPr, "descr") !== "Exyte logo") {
        errors.push(`${slideLabel}: Exyte logo must use stable alt text instead of a source path.`);
      }

      const relationshipId = logo.xml.match(/<a:blip\b[^>]*\br:embed="([^"]+)"/)?.[1];
      const relsName = `ppt/slides/_rels/slide${index + 1}.xml.rels`;
      const relsXml = zip.file(relsName) ? await zip.file(relsName)!.async("string") : "";
      const target = relationshipId ? relationshipTargets(relsXml).get(relationshipId) : undefined;
      const mediaPath = target
        ? path.posix.normalize(path.posix.join(path.posix.dirname(slideName), target))
        : "";
      const logoMedia = mediaPath ? zip.file(mediaPath) : null;
      if (!logoMedia) {
        errors.push(`${slideLabel}: Exyte logo relationship does not resolve to packaged media.`);
      } else {
        const logoBytes = await logoMedia.async("nodebuffer");
        const dimensions = pngDimensions(logoBytes);
        if (!dimensions || dimensions.width !== 1244 || dimensions.height !== 500) {
          errors.push(`${slideLabel}: Exyte logo must be the valid 1244 × 500 PNG asset.`);
        } else if (!pngHasTransparentPixel(logoBytes)) {
          errors.push(`${slideLabel}: Exyte logo PNG must contain transparent pixels.`);
        }
      }
    }

    const nonArialFaces = [...slideXml.matchAll(/typeface="([^"]+)"/g)]
      .map((match) => match[1])
      .filter((face) => face !== "Arial");
    if (nonArialFaces.length > 0) {
      errors.push(`${slideLabel}: unexpected typefaces: ${[...new Set(nonArialFaces)].join(", ")}.`);
    }
    if (/asvg:svgBlip|relationships\/image[^"]*svg/i.test(slideXml)) {
      errors.push(`${slideLabel}: contains an SVG image relationship.`);
    }
  }

  for (const name of Object.keys(zip.files).filter((entryName) => /\.(?:xml|rels)$/i.test(entryName))) {
    const xml = await zip.file(name)!.async("string");
    if (ABSOLUTE_PATH_PATTERNS.some((pattern) => pattern.test(xml))) {
      errors.push(`Absolute filesystem path leaked into ${name}.`);
    }
  }

  return { errors, warnings, slideCount };
}

function parseCli(argv: string[]): { filePath: string; expectedSlides?: number } {
  if (argv.length === 0 || argv.includes("--help")) {
    console.error(
      "Usage: node --import tsx scripts/validate-pptx.ts <file.pptx> [--expected-slides N]",
    );
    process.exit(1);
  }

  const filePath = argv[0];
  let expectedSlides: number | undefined;
  for (let index = 1; index < argv.length; index += 1) {
    if (argv[index] !== "--expected-slides" || !argv[index + 1]) {
      throw new Error(`Unknown or incomplete option: ${argv[index]}`);
    }
    expectedSlides = Number(argv[index + 1]);
    if (!Number.isInteger(expectedSlides) || expectedSlides < 1) {
      throw new Error("--expected-slides must be a positive integer.");
    }
    index += 1;
  }
  return { filePath, expectedSlides };
}

async function main(): Promise<void> {
  const { filePath, expectedSlides } = parseCli(process.argv.slice(2));
  const result = await validatePptx(filePath, { expectedSlides });
  for (const warning of result.warnings) console.warn(`WARN: ${warning}`);
  if (result.errors.length > 0) {
    for (const error of result.errors) console.error(`ERROR: ${error}`);
    process.exitCode = 1;
    return;
  }
  console.log(`Validated ${path.resolve(filePath)} (${result.slideCount} slides).`);
}

if (require.main === module) {
  main().catch((error: unknown) => {
    console.error(error);
    process.exitCode = 1;
  });
}
