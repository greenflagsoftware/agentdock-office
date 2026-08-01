// pdf.js needs its "standard fonts" glyph data available at a URL to render
// text using the base-14 PDF fonts (Helvetica, Times, etc.) that aren't
// embedded in the PDF itself — without this, page.render() hangs forever
// waiting on font data that never arrives (see PdfViewer.tsx).
// Copied from node_modules (not committed) into public/ so Vite serves it
// as a static asset alongside the SPA.
import { cpSync, existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const here = path.dirname(fileURLToPath(import.meta.url));
const src = path.join(here, "..", "node_modules", "pdfjs-dist", "standard_fonts");
const dest = path.join(here, "..", "public", "standard_fonts");

if (!existsSync(src)) {
  console.error("pdfjs-dist standard_fonts not found — run npm install first.");
  process.exit(1);
}

cpSync(src, dest, { recursive: true });
console.log("Copied pdf.js standard fonts to public/standard_fonts");
