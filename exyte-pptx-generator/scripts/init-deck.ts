#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const skillDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const starterDir = path.join(skillDir, "starter");

function usage(message?: string): never {
  if (message) console.error(message);
  console.error(
    "Usage: node --import tsx scripts/init-deck.ts <slug> [--output-dir DIR] [--title TITLE] [--date DD/MM/YYYY] [--source FILE]",
  );
  process.exit(1);
}

function sanitizeSlug(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 64);
}

function titleFromSlug(slug: string): string {
  return slug
    .split("-")
    .filter(Boolean)
    .map((word) => word[0].toUpperCase() + word.slice(1))
    .join(" ");
}

function todayDmy(): string {
  const now = new Date();
  return [
    String(now.getDate()).padStart(2, "0"),
    String(now.getMonth() + 1).padStart(2, "0"),
    now.getFullYear(),
  ].join("/");
}

const args = process.argv.slice(2);
if (args.length === 0 || args.includes("--help")) usage();

const requestedSlug = sanitizeSlug(args[0]);
if (!requestedSlug) usage("The slug must contain at least one letter or digit.");

interface InitOptions {
  outputDir: string;
  title: string;
  date: string;
  source: string;
}

const options: InitOptions = {
  outputDir: process.cwd(),
  title: titleFromSlug(requestedSlug),
  date: todayDmy(),
  source: "",
};

for (let index = 1; index < args.length; index += 1) {
  const flag = args[index];
  const value = args[index + 1];
  if (!["--output-dir", "--title", "--date", "--source"].includes(flag) || !value) {
    usage(`Unknown or incomplete option: ${flag}`);
  }
  if (flag === "--output-dir") options.outputDir = path.resolve(value);
  if (flag === "--title") options.title = value;
  if (flag === "--date") options.date = value;
  if (flag === "--source") options.source = path.resolve(value);
  index += 1;
}

if (!fs.existsSync(starterDir)) {
  throw new Error(`Starter project is missing: ${starterDir}`);
}
if (options.source && !fs.statSync(options.source).isFile()) {
  throw new Error(`Source file does not exist: ${options.source}`);
}

fs.mkdirSync(options.outputDir, { recursive: true });
let finalSlug = requestedSlug;
let destination = path.join(options.outputDir, finalSlug);
let suffix = 2;
while (fs.existsSync(destination)) {
  finalSlug = `${requestedSlug}-${suffix}`;
  destination = path.join(options.outputDir, finalSlug);
  suffix += 1;
}

fs.cpSync(starterDir, destination, { recursive: true, errorOnExist: true });

const replacements = new Map<string, string>([
  ["__DECK_SLUG__", finalSlug],
  ["__DECK_TITLE__", options.title],
  ["__DECK_DATE__", options.date],
]);
const textExtensions = new Set<string>([".json", ".md", ".ts"]);

function replaceTokens(directory: string): void {
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    const filePath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      replaceTokens(filePath);
      continue;
    }
    if (!textExtensions.has(path.extname(entry.name))) continue;
    let text = fs.readFileSync(filePath, "utf8");
    for (const [token, replacement] of replacements) {
      text = text.replaceAll(token, replacement);
    }
    fs.writeFileSync(filePath, text);
  }
}

replaceTokens(destination);

if (options.source) {
  const sourceName = path.basename(options.source);
  fs.rmSync(path.join(destination, "content", "source.md"), { force: true });
  fs.copyFileSync(options.source, path.join(destination, "content", sourceName));
}

console.log(destination);
