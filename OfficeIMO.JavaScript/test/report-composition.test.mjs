import test from "node:test";
import assert from "node:assert/strict";
import { writePdf, writePdfDocument, writePdfDocumentTo, pdfTable, ExportCell } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const small = { pageSize: { width: 200, height: 130 }, orientation: "landscape", margins: 10, pageNumbers: false, compression: false };
const paragraph = (text, options = {}) => ({ kind: "paragraph", text, ...options });
const joined = pdf => pdf.pages.flatMap(page => page.lines).join("");
const annotations = pdf => [...pdf.objects.values()].filter(object => /\/Subtype \/Link/.test(object.body));

test("rich paragraphs preserve contiguous text, emphasis, color and link fragments across wrapping", async () => {
  const label = "Read the detailed report with additional findings and context";
  const pdf = await inspectPdf(await writePdfDocument([paragraph([
    { text: "Status: " }, { text: "ready. ", bold: true, color: "008000" },
    { text: label, italic: true, link: { target: "https://example.com/report", tooltip: "Detailed report" } }
  ])], small));
  assert.equal(joined(pdf), "Status: ready. " + label);
  const fonts = [...pdf.objects.values()].filter(object => /\/Type \/Font /.test(object.body));
  assert.ok(fonts.some(font => /\/BaseFont \/Helvetica-Bold\b/.test(font.body)));
  assert.ok(fonts.some(font => /\/BaseFont \/Helvetica-Oblique\b/.test(font.body)));
  assert.match(pdf.pages[0].content, /0 0\.502 0 rg/);
  assert.ok(annotations(pdf).length >= 2, "A wrapped linked run needs one annotation for every painted fragment");
  for (const annotation of annotations(pdf)) {
    const rectangle = /\/Rect \[([^\]]+)\]/.exec(annotation.body)[1].split(" ").map(Number);
    assert.ok(rectangle[0] >= 10 && rectangle[2] <= 190.001 && rectangle[2] > rectangle[0]);
    assert.ok(rectangle[1] >= 10 && rectangle[3] <= 120.001 && rectangle[3] > rectangle[1]);
    assert.match(annotation.body, /68747470733a2f2f6578616d706c652e636f6d2f7265706f7274/);
  }
});

test("inline boundaries preserve spaces, cross-run CRLF and tabs without empty link annotations", async () => {
  const pdf = await inspectPdf(await writePdfDocument([paragraph([
    { text: "A\r" }, { text: "\nB\t", bold: true }, { text: " C" },
    { text: "", link: { target: "https://example.com/empty" } }
  ])], small));
  assert.deepEqual(pdf.pages[0].lines, ["A", "B    ", " C"]);
  assert.equal(annotations(pdf).length, 0);
  await assert.rejects(writePdfDocument([paragraph([{ text: "A\u0001B" }])], small), /control/);
  await assert.rejects(writePdfDocument([paragraph([{ text: "\ud800" }])], small), /surrogate/);
  await assert.rejects(writePdfDocument([paragraph([{ text: "\ud83e" }, { text: "\uddea" }])], small), /embedded Unicode TrueType font/,
    "A surrogate pair spanning two runs must be resolved as its original scalar");
});

test("PDF-only formatRuns captures once, keeps raw values and respects explicit display labels", async () => {
  const calls = [];
  const pdf = await inspectPdf(await writePdf([{ value: 12 }, { value: new ExportCell(23, { text: "Chosen label" }) },
    { value: new ExportCell(34, { text: "" }) }], {
    ...small, columns: [{ header: "Value", key: "value" }],
    formatRuns: (value, context) => { calls.push([value, context.rowIndex, context.columnIndex]); return [{ text: "Total " }, { text: String(value), bold: true }]; }
  }));
  assert.deepEqual(calls, [[12, 0, 0]]);
  assert.equal(joined(pdf), "ValueTotal 12Chosen label");
  assert.ok([...pdf.objects.values()].some(object => /\/BaseFont \/Helvetica-Bold\b/.test(object.body)));
  await assert.rejects(writePdf([{ value: 1 }], { ...small, columns: [{ header: "Value", key: "value" }], formatRuns: async () => [] }), /synchronous array/);
  await assert.rejects(writePdf([{ value: 1 }], { ...small, columns: [{ header: "Value", key: "value" }], formatRuns: () => "wrong" }), /synchronous array/);
});

test("rich cell links continue across pages and cannot overlap a whole-cell annotation", async () => {
  const words = "Linked words ".repeat(45);
  const pdf = await inspectPdf(await writePdf([{ value: 1 }], { ...small, columns: [{ header: "Value", key: "value" }],
    formatRuns: () => [{ text: words, link: { target: "https://example.com/value" } }] }));
  assert.ok(pdf.pages.length > 1);
  assert.equal(joined(pdf).replaceAll("Value", ""), words);
  assert.ok(pdf.pages.every(page => /\/Annots \[/.test(page.body)));
  await assert.rejects(writePdf([{ value: new ExportCell(1, { link: { target: "https://example.com/cell" } }) }], {
    ...small, columns: [{ header: "Value", key: "value" }], formatRuns: () => [{ text: "Linked", link: { target: "https://example.com/run" } }]
  }), /cell link.*inline run links/);
});

test("kept headings accompany the first actual table fragment without buffering its row source", async () => {
  let produced = 0, returned = 0;
  const words = "First actual row ".repeat(40);
  function* rows() { try { produced++; yield { value: words }; produced++; yield { value: "Last row" }; } finally { returned++; } }
  const pdf = await inspectPdf(await writePdfDocument([
    paragraph("One\nTwo\nThree\nFour\nFive\nSix"), paragraph("Section heading", { keepWithNext: true }),
    pdfTable(rows(), { columns: [{ header: "Table heading", key: "value" }] })
  ], small));
  assert.equal(produced, 2); assert.equal(returned, 1);
  assert.ok(!pdf.pages[0].lines.includes("Section heading"));
  assert.ok(pdf.pages[1].lines.includes("Section heading") && pdf.pages[1].lines.includes("Table heading"));
  assert.ok(pdf.pages[1].lines.some(line => line.includes("First actual row")));
  assert.equal(pdf.pages.flatMap(page => page.lines).filter(line => line === "Section heading").length, 1);
  assert.ok(joined(pdf).includes("Last row"));
});

test("kept headings reserve complete graphics and bounded groups move intact", async () => {
  const graphic = { kind: "vector", width: 40, height: 40, fill: "008000", commands: [
    { op: "move", x: 0, y: 0 }, { op: "line", x: 40, y: 40 }
  ] };
  const pdf = await inspectPdf(await writePdfDocument([
    paragraph("One\nTwo\nThree\nFour\nFive\nSix"), paragraph("Graphic heading", { keepWithNext: true }), graphic,
    { kind: "group", blocks: [paragraph("Grouped heading", { headingLevel: 2 }), paragraph("Grouped body\nSecond line")] }
  ], small));
  assert.ok(!pdf.pages[0].lines.includes("Graphic heading"));
  assert.ok(pdf.pages[1].lines.includes("Graphic heading"));
  assert.match(pdf.pages[1].content, /0 0 40 40 re W n/);
  const groupPage = pdf.pages.find(page => page.lines.includes("Grouped heading"));
  assert.ok(groupPage.lines.includes("Grouped body") && groupPage.lines.includes("Second line"));
});

test("non-rendering paragraphs and groups preserve keeps until the next visible block", async () => {
  const gaps = [paragraph(""), paragraph("", { keepWithNext: true }), paragraph([]), paragraph([{ text: "" }]),
    { kind: "group", blocks: [] }, { kind: "group", blocks: [paragraph(""), paragraph([])] },
    pdfTable([], { columns: [{ header: "Value", key: "value" }], includeHeader: false })];
  const following = [
    () => pdfTable([{ value: "First actual row ".repeat(40) }], { columns: [{ header: "Table heading", key: "value" }] }),
    () => paragraph("Visible body\nSecond body line"),
    () => ({ kind: "group", blocks: [paragraph("Visible body"), paragraph("Second body line")] })
  ];
  const prefix = () => [paragraph("One\nTwo\nThree\nFour\nFive\nSix"), paragraph("Section heading", { keepWithNext: true })];
  for (const next of following) {
    const expected = await inspectPdf(await writePdfDocument([...prefix(), next()], small));
    for (const gap of gaps) {
      const pdf = await inspectPdf(await writePdfDocument([...prefix(), gap, next()], small));
      assert.deepEqual(pdf.pages.map(page => page.lines), expected.pages.map(page => page.lines));
      const page = pdf.pages.find(page => page.lines.includes("Section heading"));
      assert.ok(page.lines.some(line => line.includes("First actual row") || line === "Visible body"));
      assert.equal(pdf.pages.flatMap(page => page.lines).filter(line => line === "Section heading").length, 1);
    }
  }
  await assert.rejects(writePdfDocument([paragraph("", { keepWithNext: "invalid" })], small), /keepWithNext must be boolean/);
  const footerOnly = await inspectPdf(await writePdfDocument([...prefix(),
    pdfTable([], { columns: [{ header: "Value", key: "value" }], includeHeader: false,
      footer: { values: [Array.from({ length: 8 }, (_, i) => "Visible footer " + (i + 1)).join("\n")] } })
  ], small));
  assert.ok(!footerOnly.pages[0].lines.includes("Section heading"));
  assert.ok(footerOnly.pages[1].lines.includes("Section heading") && footerOnly.pages[1].lines.includes("Visible footer 1"));
  assert.ok(footerOnly.pages.at(-1).lines.includes("Visible footer 8"));
});

test("paragraph keep controls, groups and inline metadata have explicit bounded failure contracts", async () => {
  await assert.rejects(writePdfDocument([paragraph("Long\n".repeat(15), { keepTogether: true })], small), /fresh page/);
  await assert.rejects(writePdfDocument([{ kind: "group", blocks: [paragraph("Long\n".repeat(15))] }], small), /fresh page/);
  await assert.rejects(writePdfDocument([{ kind: "group", blocks: [{ kind: "pageBreak" }] }], small), /paragraph blocks only/);
  await assert.rejects(writePdfDocument([{ kind: "group", blocks: [paragraph("Heading", { keepWithNext: true })] }], small), /group's boundary/);
  await assert.rejects(writePdfDocument([{ kind: "group", blocks: [paragraph("A"), paragraph("B")] }], { ...small, limits: { maxGroupBlocks: 1 } }), /maxGroupBlocks/);
  await assert.rejects(writePdfDocument([paragraph([{ text: "A" }, { text: "B" }])], { ...small, limits: { maxTextRuns: 1 } }), /maxTextRuns/);
  await assert.rejects(writePdfDocument([{ kind: "group", blocks: [paragraph("ABC"), paragraph("DEF")] }], { ...small, limits: { maxCellCharacters: 5 } }), /maxCellCharacters/);
  const plain = await inspectPdf(await writePdfDocument([paragraph("Plain")], { ...small, limits: { maxTextRuns: 0, maxGroupBlocks: 0 } }));
  assert.equal(joined(plain), "Plain");
  let advanced = 0;
  function* source() { yield paragraph("Heading", { keepWithNext: true }); advanced++; yield paragraph("Body"); }
  await assert.rejects(writePdfDocument(source(), { ...small, limits: { maxTextCharacters: 0 } }), /maxTextCharacters/);
  assert.equal(advanced, 0, "A kept paragraph must reserve its text budget before looking ahead");
});

test("tagged rich reports retain one logical paragraph and native inline link ownership", async () => {
  const pdf = await inspectPdf(await writePdfDocument([paragraph([{ text: "Normal " }, { text: "strong", bold: true }], { headingLevel: 2 })],
    { ...small, tagged: { language: "en-US" } }));
  assert.equal([...pdf.objects.values()].filter(object => /\/Type \/StructElem \/S \/H2 /.test(object.body)).length, 1);
  assert.equal(joined(pdf), "Normal strong");
  const linked = await inspectPdf(await writePdfDocument([paragraph([{ text: "Link", link: { target: "https://example.com" } }])],
    { ...small, tagged: { language: "en-US" } }));
  assert.equal(annotations(linked).length, 1);
  assert.ok([...linked.objects.values()].some(object => /\/S \/Link /.test(object.body) && /\/Type \/OBJR/.test(object.body)));
});

test("rich paragraph snapshots survive cooperative layout and stream output matches Blob output", async () => {
  const runs = [{ text: "A".repeat(4096), bold: true, link: { target: "https://example.com/original" } }];
  const now = Object.getOwnPropertyDescriptor(performance, "now"), scheduler = globalThis.scheduler;
  let yields = 0;
  try {
    let clock = 0; Object.defineProperty(performance, "now", { configurable: true, value: () => clock += 40 });
    globalThis.scheduler = { postTask: async callback => { if (++yields === 1) { runs[0].text = "changed"; runs[0].bold = false; runs[0].link.target = "https://example.com/changed"; } callback(); } };
    const pdf = await inspectPdf(await writePdfDocument([paragraph(runs)], small));
    assert.ok(yields > 0); assert.equal(joined(pdf), "A".repeat(4096));
    assert.ok(annotations(pdf).every(annotation => !annotation.body.includes(Buffer.from("changed").toString("hex"))));
  } finally { if (now) Object.defineProperty(performance, "now", now); else delete performance.now; globalThis.scheduler = scheduler; }
  const blocks = () => [paragraph([{ text: "Report " }, { text: "link", link: { target: "https://example.com" } }], { keepWithNext: true }),
    pdfTable([{ value: 7 }], { columns: [{ header: "Value", key: "value" }], formatRuns: value => [{ text: String(value), bold: true }] })];
  const blob = await writePdfDocument(blocks(), small), chunks = [];
  const result = await writePdfDocumentTo(blocks(), { write: bytes => chunks.push(bytes.slice()) }, small);
  assert.deepEqual(Buffer.from(await blob.arrayBuffer()), Buffer.concat(chunks));
  assert.equal(result.rows, 1); assert.equal(result.columns, 1); assert.equal(result.bytes, blob.size);
});
