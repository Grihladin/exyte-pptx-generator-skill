import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { spawn } from "node:child_process";

const initializer = path.resolve("exyte-pptx-generator/scripts/init-deck.ts");

function runInitializer(args: string[]): Promise<string> {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, ["--import", "tsx", initializer, ...args], {
      stdio: ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => (stdout += chunk));
    child.stderr.on("data", (chunk) => (stderr += chunk));
    child.on("error", reject);
    child.on("close", (code) => {
      if (code === 0) resolve(stdout.trim());
      else reject(new Error(stderr || `Initializer exited with ${code}`));
    });
  });
}

test("initializer creates self-contained unique deck projects", async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "exyte-init-"));
  try {
    const first = await runInitializer([
      "Quarterly Update",
      "--output-dir",
      directory,
      "--title",
      "Quarterly Update",
      "--date",
      "19/06/2026",
    ]);
    const second = await runInitializer(["Quarterly Update", "--output-dir", directory]);

    assert.equal(path.basename(first), "quarterly-update");
    assert.equal(path.basename(second), "quarterly-update-2");
    for (const relativePath of [
      "assets/exyte_logo.png",
      "content/source.md",
      "package-lock.json",
      "package.json",
      "scripts/validate-pptx.ts",
      "slides_code/build.ts",
      "theme.ts",
      "tsconfig.json",
    ]) {
      await fs.access(path.join(first, relativePath));
    }
    const buildSource = await fs.readFile(path.join(first, "slides_code/build.ts"), "utf8");
    assert(!buildSource.includes("__DECK_"));
    assert(buildSource.includes('pptx.layout = "LAYOUT_WIDE"'));
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
});
