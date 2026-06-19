import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import {
  EXPECTED_SLIDE_HEIGHT_EMU,
  EXPECTED_SLIDE_WIDTH_EMU,
  validatePptx,
} from "../exyte-pptx-generator/starter/scripts/validate-pptx";

const VALID_PNG_PATH = path.resolve("exyte-pptx-generator/starter/assets/exyte_logo.png");

interface FixtureOptions {
  width?: number;
  height?: number;
  bodyX?: number;
  bodyY?: number;
  bodyW?: number;
  bodyH?: number;
  omitChrome?: string;
  media?: Buffer;
  logoDescription?: string;
}

function shape(name: string, x: number, y: number, w: number, h: number, text = ""): string {
  const textXml = text
    ? `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:latin typeface="Arial"/></a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody>`
    : "";
  return `<p:sp><p:nvSpPr><p:cNvPr id="2" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm></p:spPr>${textXml}</p:sp>`;
}

function logo(description: string): string {
  return `<p:pic><p:nvPicPr><p:cNvPr id="3" name="Exyte Logo" descr="${description}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/></p:blipFill><p:spPr><a:xfrm><a:off x="10100000" y="274320"/><a:ext cx="1353312" cy="544068"/></a:xfrm></p:spPr></p:pic>`;
}

async function createFixture(directory: string, options: FixtureOptions = {}): Promise<string> {
  const zip = new JSZip();
  const width = options.width ?? EXPECTED_SLIDE_WIDTH_EMU;
  const height = options.height ?? EXPECTED_SLIDE_HEIGHT_EMU;
  const required = [
    ["Exyte Footer Line", 0, 6_510_528, width, 27_432],
    ["Exyte Footer Left", 118_872, 6_629_400, 6_400_800, 128_016],
    ["Exyte Footer Date", 10_193_040, 6_629_400, 822_960, 128_016],
    ["Exyte Footer Divider", 11_173_968, 6_629_400, 73_152, 128_016],
    ["Exyte Footer Page", 11_456_832, 6_629_400, 228_600, 128_016],
  ] as const;
  const chrome = required
    .filter(([name]) => name !== options.omitChrome)
    .map(([name, x, y, w, h]) => shape(name, x, y, w, h, name))
    .join("");
  const body = shape(
    "Body",
    options.bodyX ?? 731_520,
    options.bodyY ?? 1_371_600,
    options.bodyW ?? 5_486_400,
    options.bodyH ?? 914_400,
    "Body",
  );
  const title = shape("Exyte Slide Title", 731_520, 347_472, 8_900_000, 457_200, "Title");
  const logoXml = options.omitChrome === "Exyte Logo" ? "" : logo(options.logoDescription ?? "Exyte logo");

  zip.file(
    "[Content_Types].xml",
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/></Types>`,
  );
  zip.file(
    "ppt/presentation.xml",
    `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst><p:sldId id="256"/></p:sldIdLst><p:sldSz cx="${width}" cy="${height}"/></p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
  );
  zip.file(
    "ppt/slides/slide1.xml",
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>${logoXml}${chrome}${title}${body}</p:spTree></p:cSld></p:sld>`,
  );
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/logo.png"/></Relationships>`,
  );
  zip.file("ppt/media/logo.png", options.media ?? (await fs.readFile(VALID_PNG_PATH)));

  const filePath = path.join(directory, "fixture.pptx");
  await fs.writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
  return filePath;
}

async function withFixture(
  options: FixtureOptions,
  callback: (errors: string[]) => void | Promise<void>,
): Promise<void> {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "exyte-validator-"));
  try {
    const filePath = await createFixture(directory, options);
    const result = await validatePptx(filePath, {
      expectedSlides: 1,
      checkSiblingOutputs: false,
    });
    await callback(result.errors);
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
}

test("accepts a structurally valid deck", async () => {
  await withFixture({}, (errors) => assert.deepEqual(errors, []));
});

test("rejects the legacy 10 × 5.625 canvas", async () => {
  await withFixture({ width: 9_144_000, height: 5_143_500 }, (errors) => {
    assert(errors.some((error) => error.includes("Slide size must be")));
  });
});

test("rejects off-slide geometry", async () => {
  await withFixture({ bodyX: 12_000_000, bodyW: 1_000_000 }, (errors) => {
    assert(errors.some((error) => error.includes("extends off-slide")));
  });
});

test("rejects footer intrusion", async () => {
  await withFixture({ bodyY: 6_300_000, bodyH: 400_000 }, (errors) => {
    assert(errors.some((error) => error.includes("intrudes into the footer")));
  });
});

test("rejects mislabeled PNG media", async () => {
  await withFixture({ media: Buffer.from("<svg/>") }, (errors) => {
    assert(errors.some((error) => error.includes("PNG extension does not match")));
  });
});

test("rejects absolute path leakage", async () => {
  await withFixture({ logoDescription: "/Users/example/logo.png" }, (errors) => {
    assert(errors.some((error) => error.includes("Absolute filesystem path leaked")));
  });
});

test("rejects missing required chrome", async () => {
  await withFixture({ omitChrome: "Exyte Footer Page" }, (errors) => {
    assert(errors.some((error) => error.includes('expected one "Exyte Footer Page"')));
  });
});
