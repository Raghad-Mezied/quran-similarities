// Fetch similarities_new/1.json .. 114.json and save them to ./data
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const BASE_URL = "https://hefzmoyaser.com/data/similarities_new";
const OUT_DIR = path.join(__dirname, "data");
const START = 1;
const END = 114;

async function fetchOne(n) {
  const url = `${BASE_URL}/${n}.json`;
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`HTTP ${res.status} for ${url}`);
  }
  // Read as text so we save exactly what the server returned
  const text = await res.text();
  const outPath = path.join(OUT_DIR, `${n}.json`);
  fs.writeFileSync(outPath, text);
  return outPath;
}

async function main() {
  fs.mkdirSync(OUT_DIR, { recursive: true });

  let ok = 0;
  let failed = 0;

  for (let n = START; n <= END; n++) {
    try {
      const outPath = await fetchOne(n);
      ok++;
      console.log(`[${n}/${END}] saved -> ${outPath}`);
    } catch (err) {
      failed++;
      console.error(`[${n}/${END}] FAILED: ${err.message}`);
    }
  }

  console.log(`\nDone. Saved ${ok} file(s), ${failed} failure(s) into ${OUT_DIR}`);
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
