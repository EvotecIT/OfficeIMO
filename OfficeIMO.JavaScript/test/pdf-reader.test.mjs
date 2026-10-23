import test from "node:test";
import assert from "node:assert/strict";
import { inspectPdf } from "./pdf-reader.mjs";

// Hand-authored WinAnsi text operands; no OfficeIMO writer or encoder participates.
// Character assignments: ISO 32000-1:2008, Appendix D.2, Table D.1 and notes 1-6.
// https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/PDF32000_2008.pdf
function winAnsiPdf(hexLines) {
  const content = hexLines.map((hex, i) => `BT /F1 12 Tf 1 0 0 1 20 ${100 - i * 14} Tm <${hex}> Tj ET\n`).join("");
  const bodies = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources 5 0 R /Contents 4 0 R >>",
    `<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}\nendstream`,
    "<< /Font << /F1 6 0 R >> >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
  ];
  let pdf = "%PDF-1.7\n";
  const offsets = [0];
  for (const [i, body] of bodies.entries()) {
    offsets.push(Buffer.byteLength(pdf));
    pdf += `${i + 1} 0 obj\n${body}\nendobj\n`;
  }
  const start = Buffer.byteLength(pdf);
  pdf += `xref\n0 ${offsets.length}\n0000000000 65535 f \n`;
  for (const offset of offsets.slice(1)) pdf += `${String(offset).padStart(10, "0")} 00000 n \n`;
  pdf += `trailer\n<< /Size ${offsets.length} /Root 1 0 R >>\nstartxref\n${start}\n%%EOF\n`;
  return new Blob([pdf], { type: "application/pdf" });
}

test("PDF inspector reads WinAnsi Euro and CP1252 special glyphs from independent operands", async () => {
  const pdf = await inspectPdf(winAnsiPdf([
    "8082838485868788898a8b8c8e",
    "9192939495969798999a9b9c9e9f"
  ]));
  assert.deepEqual(pdf.pages[0].lines, ["€‚ƒ„…†‡ˆ‰Š‹ŒŽ", "‘’“”•–—˜™š›œžŸ"]);
});

test("PDF inspector reads unused WinAnsi codes above space as bullet glyphs", async () => {
  const pdf = await inspectPdf(winAnsiPdf(["7f818d8f909d", "95"]));
  assert.deepEqual(pdf.pages[0].lines, ["••••••", "•"]);
});

test("PDF inspector reads WinAnsi space and hyphen aliases as their shared glyphs", async () => {
  const pdf = await inspectPdf(winAnsiPdf(["20a02dad", "21417ea1ace9aeff"]));
  assert.deepEqual(pdf.pages[0].lines, ["  --", "!A~¡¬é®ÿ"]);
});
