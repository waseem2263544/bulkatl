#!/usr/bin/env node
// Builds the prebuilt ATL data file consumed by the runtime.
// Output: dist/atl-ids.txt.gz — sorted, deduped, normalized NTNs/CNICs, gzipped.
//
// Invoked by .github/workflows/refresh-atl.yml every Monday.

import { promises as fs, createWriteStream } from "node:fs";
import path from "node:path";
import zlib from "node:zlib";
import { Readable } from "node:stream";
import { pipeline } from "node:stream/promises";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

const FBR_ATL_URL =
  process.env.FBR_ATL_URL ?? "https://download.fbr.gov.pk/IT/ATL_IT.xlsx";
const OUT_DIR = process.env.OUT_DIR ?? path.resolve("dist");
const RAW_PATH = path.join(OUT_DIR, "atl-raw.xlsx");
const IDS_GZ_PATH = path.join(OUT_DIR, "atl-ids.txt.gz");
const META_PATH = path.join(OUT_DIR, "atl-meta.json");

const CHUNK_BYTES = 4 * 1024 * 1024;
const MAX_CHUNK_RETRIES = 8;
const REQUEST_TIMEOUT_MS = 120_000;

await main();

async function main() {
  await fs.mkdir(OUT_DIR, { recursive: true });
  if (await fileExists(RAW_PATH) && process.env.ATL_SKIP_DOWNLOAD === "1") {
    log(`Reusing existing ${RAW_PATH} (ATL_SKIP_DOWNLOAD=1)`);
  } else {
    log(`Downloading ATL from ${FBR_ATL_URL}…`);
    await downloadWithRetry(FBR_ATL_URL, RAW_PATH);
  }
  const stat = await fs.stat(RAW_PATH);
  log(`Downloaded ${(stat.size / 1024 / 1024).toFixed(1)} MB`);

  log("Parsing xlsx…");
  const ids = await parseAllSheets(RAW_PATH);
  log(`Extracted ${ids.size.toLocaleString()} unique IDs`);

  log("Sorting & writing gzipped output…");
  const sorted = Array.from(ids).sort();
  await writeGzippedLines(sorted, IDS_GZ_PATH);
  const out = await fs.stat(IDS_GZ_PATH);
  log(`Wrote ${IDS_GZ_PATH} (${(out.size / 1024 / 1024).toFixed(1)} MB gzipped)`);

  const meta = {
    fetchedAt: Date.now(),
    rows: sorted.length,
    source: FBR_ATL_URL,
    fileSize: out.size,
    builtAt: new Date().toISOString(),
  };
  await fs.writeFile(META_PATH, JSON.stringify(meta, null, 2));
  log(`Wrote ${META_PATH}`);

  // Drop the raw 161 MB file — we only ship the gz output.
  await fs.unlink(RAW_PATH).catch(() => {});
  log("Done.");
}

function log(...args) {
  console.log("[atl-build]", ...args);
}

async function fileExists(p) {
  try {
    await fs.stat(p);
    return true;
  } catch {
    return false;
  }
}

function normalizeId(value) {
  if (value == null) return "";
  return String(value).replace(/[^0-9A-Za-z]/g, "").toUpperCase();
}

async function parseAllSheets(filePath) {
  const buf = await fs.readFile(filePath);
  const wb = XLSX.read(buf, {
    type: "buffer",
    dense: true,
    cellDates: false,
    cellFormula: false,
    cellStyles: false,
    cellHTML: false,
    cellNF: false,
  });
  const ids = new Set();
  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    if (!sheet) continue;
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      blankrows: false,
      defval: "",
      raw: true,
    });
    if (rows.length < 2) continue;
    const header = rows[0].map((c) => String(c ?? "").toLowerCase().trim());
    let idCols = header
      .map((h, i) => ({ h, i }))
      .filter(
        ({ h }) =>
          /\bntn\b/.test(h) ||
          /registration/.test(h) ||
          /\bcnic\b/.test(h) ||
          /reg\.?\s*no/.test(h),
      )
      .map(({ i }) => i);
    if (idCols.length === 0) idCols = [1];

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      for (const c of idCols) {
        const cell = row[c];
        const v = normalizeId(cell);
        if (v.length < 5) continue;
        ids.add(v);
        const raw = String(cell ?? "").trim();
        if (/^\d{1,9}-\d$/.test(raw)) {
          ids.add(normalizeId(raw.split("-")[0]));
        }
      }
    }
    // Free as we go.
    wb.Sheets[sheetName] = undefined;
  }
  return ids;
}

async function writeGzippedLines(lines, dest) {
  const gzip = zlib.createGzip({ level: 9 });
  const out = createWriteStream(dest);
  // Stream lines through gzip so we never hold the full text in memory twice.
  async function* gen() {
    const BATCH = 8192;
    for (let i = 0; i < lines.length; i += BATCH) {
      yield lines.slice(i, i + BATCH).join("\n") + "\n";
    }
  }
  await pipeline(Readable.from(gen()), gzip, out);
}

async function downloadWithRetry(url, dest) {
  const total = await getContentLength(url);
  if (!total) {
    return downloadWhole(url, dest);
  }
  await fs.mkdir(path.dirname(dest), { recursive: true });
  const fh = await fs.open(dest, "w");
  try {
    let offset = 0;
    while (offset < total) {
      const end = Math.min(offset + CHUNK_BYTES - 1, total - 1);
      const buf = await fetchRangeWithRetry(url, offset, end);
      await fh.write(buf, 0, buf.length, offset);
      offset = end + 1;
      const pct = ((offset / total) * 100).toFixed(1);
      if (offset === total || Number(pct) % 10 < 1) {
        log(`  ${pct}% (${(offset / 1024 / 1024).toFixed(0)} MB)`);
      }
    }
  } finally {
    await fh.close();
  }
}

async function getContentLength(url) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), 30_000);
  try {
    const res = await fetch(url, {
      method: "HEAD",
      headers: { "User-Agent": "Mozilla/5.0", Accept: "*/*" },
      redirect: "follow",
      signal: ctrl.signal,
    });
    if (!res.ok) return null;
    const len = res.headers.get("content-length");
    if (!len) return null;
    const n = Number.parseInt(len, 10);
    return Number.isFinite(n) && n > 0 ? n : null;
  } catch {
    return null;
  } finally {
    clearTimeout(t);
  }
}

async function fetchRangeWithRetry(url, start, end) {
  let lastErr = null;
  for (let attempt = 1; attempt <= MAX_CHUNK_RETRIES; attempt++) {
    try {
      return await fetchRange(url, start, end);
    } catch (err) {
      lastErr = err;
      const wait = Math.min(20_000, 500 * Math.pow(2, attempt - 1));
      await new Promise((r) => setTimeout(r, wait));
    }
  }
  throw new Error(
    `Failed to fetch bytes ${start}-${end} after ${MAX_CHUNK_RETRIES} attempts: ${
      lastErr?.message ?? lastErr
    }`,
  );
}

async function fetchRange(url, start, end) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), REQUEST_TIMEOUT_MS);
  try {
    const res = await fetch(url, {
      headers: {
        "User-Agent": "Mozilla/5.0",
        Accept: "*/*",
        Range: `bytes=${start}-${end}`,
        Connection: "keep-alive",
      },
      redirect: "follow",
      signal: ctrl.signal,
    });
    if (res.status !== 206 && res.status !== 200) {
      throw new Error(`HTTP ${res.status}`);
    }
    const buf = Buffer.from(await res.arrayBuffer());
    const expected = end - start + 1;
    if (res.status === 206 && buf.length !== expected) {
      throw new Error(`Short range: got ${buf.length} expected ${expected}`);
    }
    return buf;
  } finally {
    clearTimeout(t);
  }
}

async function downloadWhole(url, dest) {
  let lastErr = null;
  for (let attempt = 1; attempt <= MAX_CHUNK_RETRIES; attempt++) {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), REQUEST_TIMEOUT_MS * 4);
    try {
      const res = await fetch(url, {
        headers: { "User-Agent": "Mozilla/5.0", Accept: "*/*" },
        redirect: "follow",
        signal: ctrl.signal,
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const buf = Buffer.from(await res.arrayBuffer());
      await fs.writeFile(dest, buf);
      return;
    } catch (err) {
      lastErr = err;
      const wait = Math.min(20_000, 1000 * Math.pow(2, attempt - 1));
      await new Promise((r) => setTimeout(r, wait));
    } finally {
      clearTimeout(t);
    }
  }
  throw new Error(`Failed to download ATL: ${lastErr?.message ?? lastErr}`);
}
