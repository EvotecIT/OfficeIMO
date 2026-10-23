import test from "node:test";
import assert from "node:assert/strict";
import { BlobByteSink } from "../dist/core/sinks.js";
import { PdfWriteSession } from "../dist/pdf/session.js";
import { settings } from "../dist/pdf/settings.js";
import { PdfTableLayout } from "../dist/pdf/table-layout.js";
import { captureCaption, prepareParagraph } from "../dist/pdf/paragraph-layout.js";
import { inspectPdf } from "./pdf-reader.mjs";

// Exercise the internal grouped-owner boundary with real page/layout/serialization owners.
// Public grouped fixtures separately verify the hierarchy that supplies this already-painted path.
async function contextTable(bands, precedingLines, detailLines) {
  const configuration = settings({ columns: [{ header: "Identifier" }, { header: "Detail" }, { header: "Metric" }],
    ...(bands ? { wideTable: { mode: "bands", repeatColumns: [0] } } : {}), columnWidths: [80, 100, 100],
    pageSize: { width: 300, height: 360 }, margins: 20, fontSize: 9, padding: 4,
    pageNumbers: false, compression: false, tagged: { language: "en" } });
  const sink = new BlobByteSink(), session = new PdfWriteSession(sink, configuration);
  await session.start();
  if (precedingLines) await session.pages.paragraph(Array.from({ length: precedingLines }, (_, index) => `Preface-${index}`).join("\n"));
  const caption = await prepareParagraph(captureCaption({ text: "Root\nLeaf", spaceAfter: 0 }, configuration, session.fonts), configuration, session.fonts);
  const continuation = await prepareParagraph(captureCaption({ text: "CONT Root\nCONT Leaf", spaceAfter: 0 }, configuration, session.fonts), configuration, session.fonts);
  await session.pages.paragraphs([caption], 0, false, { artifact: true });
  const layout = new PdfTableLayout(configuration, session.fonts), headers = await layout.headers();
  const row = await layout.data(["ROW-ID", Array.from({ length: detailLines }, (_, index) => `Detail-${index}`).join("\n"), "Metric-value"], 0);
  await session.pages.beginTable(layout, headers, layout.minimumRowHeight(row), caption, continuation, false, true, true);
  await session.pages.row(row); await session.pages.endTable(); await session.finish();
  return inspectPdf(sink.toBlob("application/pdf"));
}

test("already-painted artifact context is omitted only from the existing first table band", async () => {
  for (const bands of [false, true]) {
    const pdf = await contextTable(bands, 0, 2), first = pdf.pages[0];
    assert.equal(first.lines.filter(line => line === "Root").length, 1);
    assert.equal(first.lines.filter(line => line === "Leaf").length, 1);
    assert.deepEqual(first.lines.slice(0, 3), ["Root", "Leaf", "Identifier"]);
    if (bands) {
      assert.equal(pdf.pages.length, 2);
      assert.deepEqual(pdf.pages[1].lines.slice(0, 3), ["Root", "Leaf", "Identifier"], "A later horizontal band retains the full captured opening path");
    }
  }
});

test("moving an already-painted context table to a fresh page restores complete context and headings", async () => {
  for (const bands of [false, true]) {
    const pdf = await contextTable(bands, 8, 22);
    assert.ok(!pdf.pages[0].lines.includes("ROW-ID"), "The intact row moves off the partially occupied context page");
    const first = pdf.pages.find(page => page.lines.includes("ROW-ID"));
    assert.deepEqual(first.lines.slice(0, 3), bands ? ["Root", "Leaf", "Identifier"] : ["CONT Root", "CONT Leaf", "Identifier"],
      "A fresh table page cannot inherit the opening-caption omission");
    assert.deepEqual(pdf.pages.flatMap(page => page.lines).filter(line => /^Detail-/.test(line)),
      Array.from({ length: 22 }, (_, index) => `Detail-${index}`));
  }
});
