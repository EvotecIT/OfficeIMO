import test from "node:test";
import assert from "node:assert/strict";
import { writePdfDocument, writePdfDocumentTo } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";
import { pdfStructureElements } from "./pdf-structure-reader.mjs";

test("named paragraph styles snapshot before enumeration and explicit appearance merges by member", async () => {
  const paragraphStyles = { body: { fontSize: 12, presentation: { bold: true, italic: true, color: "123456", alignment: "right" },
    spaceBefore: 3, spaceAfter: 4, keepTogether: true } };
  function* blocks() {
    paragraphStyles.body.fontSize = 40; paragraphStyles.body.presentation.color = "FF0000";
    yield { kind: "paragraph", text: "Named", style: "body" };
    yield { kind: "paragraph", text: "Override", style: "body", fontSize: 10, presentation: { bold: false }, spaceBefore: 0, keepTogether: false };
  }
  const pdf = await inspectPdf(await writePdfDocument(blocks(), { paragraphStyles, pageNumbers: false, compression: false }));
  const content = pdf.pages[0].content;
  assert.deepEqual(pdf.pages[0].lines, ["Named", "Override"]);
  assert.match(content, /BT \/F\d+ 12 Tf/); assert.match(content, /BT \/F\d+ 10 Tf/); assert.doesNotMatch(content, /BT \/F\d+ 40 Tf/);
  assert.equal((content.match(/0\.0706 0\.2039 0\.3373 rg/g) ?? []).length, 2);
  const fonts = [...pdf.objects.values()].map(object => object.body);
  assert.ok(fonts.some(body => body.includes("/BaseFont /Helvetica-BoldOblique")));
  assert.ok(fonts.some(body => body.includes("/BaseFont /Helvetica-Oblique")));
  const positions = [...content.matchAll(/1 0 0 1 ([\d.]+) ([\d.]+) Tm/g)].map(match => [Number(match[1]), Number(match[2])]);
  assert.ok(positions.every(([left]) => left > 400), "Both paragraphs retain named right alignment");
  assert.ok(Math.abs(positions[0][1] - positions[1][1] - 18) < .001, "Explicit zero leading spacing and named trailing spacing both apply");
});

test("named appearance leaves semantic heading roles explicit", async () => {
  const pdf = await inspectPdf(await writePdfDocument([
    { kind: "paragraph", text: "Large body", style: "display" },
    { kind: "paragraph", text: "Small heading", style: "body", headingLevel: 2 }
  ], { paragraphStyles: { display: { fontSize: 20, presentation: { bold: true } }, body: { fontSize: 9 } },
    tagged: { language: "en" }, pageNumbers: false, compression: false }));
  const roles = pdfStructureElements(pdf).map(([, object]) => /\/S \/(\w+)/.exec(object.body)[1]);
  assert.equal(roles.filter(role => role === "P").length, 1); assert.equal(roles.filter(role => role === "H2").length, 1);
  assert.ok(!roles.includes("H1"));
});

test("named style definitions reject unsupported semantics and enforce their cap before source or destination work", async () => {
  for (const named of [{ style: "base" }, { headingLevel: 1 }, { presentation: { background: "FFFFFF" } }])
    await assert.rejects(writePdfDocument([], { paragraphStyles: { invalid: named } }), /paragraph styles|paragraph presentation/);
  for (const alignment of ["fill", "justify", "distributed"])
    await assert.rejects(writePdfDocument([], { paragraphStyles: { invalid: { presentation: { alignment } } } }), /alignment/);
  for (const style of ["missing", "toString"])
    await assert.rejects(writePdfDocument([{ kind: "paragraph", text: "Text", style }]), /Unknown PDF paragraph style/);
  let produced = 0, writes = 0;
  function* blocks() { produced++; yield { kind: "paragraph", text: "Never" }; }
  await assert.rejects(writePdfDocumentTo(blocks(), { write() { writes++; } }, {
    paragraphStyles: { first: {}, second: {} }, limits: { maxParagraphStyles: 1 }
  }), error => error.code === "RESOURCE_LIMIT" && /maxParagraphStyles/.test(error.message));
  assert.equal(produced, 0); assert.equal(writes, 0);
  const pdf = await inspectPdf(await writePdfDocument([{ kind: "paragraph", text: "Direct", fontSize: 11 }], { limits: { maxParagraphStyles: 0 }, pageNumbers: false }));
  assert.equal(pdf.text, "Direct");
});
