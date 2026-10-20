import test from "node:test";
import assert from "node:assert/strict";
import { writePdf, writePdfDocument, writePdfDocumentTo, pdfTable, ExportCell } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const columns = [{ header: "Name", key: "name" }, { header: "Amount", key: "amount" }];
const tagged = { language: "en-US" };
const image = { format: "rgb", width: 1, height: 1, data: Uint8Array.of(0, 128, 200) };
const vector = { kind: "vector", width: 10, height: 10, fill: "008000", commands: [
  { op: "move", x: 0, y: 0 }, { op: "line", x: 10, y: 10 }, { op: "line", x: 0, y: 10 }, { op: "close" }
] };
const refs = body => [...body.matchAll(/(\d+) 0 R/g)].map(match => Number(match[1]));
const ref = (body, name) => Number(new RegExp("/" + name + " (\\d+) 0 R").exec(body)?.[1]);

// Walk the actual xref-resolved graph in both directions; all painted MCIDs must have one owner.
async function inspectStructure(blob) {
  const pdf = await inspectPdf(blob), catalog = [...pdf.objects.values()].find(object => /\/Type \/Catalog /.test(object.body));
  const rootId = ref(catalog.body, "StructTreeRoot"), root = pdf.objects.get(rootId);
  assert.match(catalog.body, /\/MarkInfo << \/Marked true >>/);
  assert.match(catalog.body, /\/Lang <feff0065006e002d00550053>/);
  assert.match(root.body, /\/Type \/StructTreeRoot /);
  const elements = new Map([...pdf.objects].filter(([, object]) => /\/Type \/StructElem /.test(object.body)).map(([id, object]) => [id, {
    id, body: object.body, role: /\/S \/(\w+)/.exec(object.body)[1], parent: ref(object.body, "P"),
    kids: /\/K \[([\s\S]*?)\](?: |$)/.exec(object.body)[1]
  }]));
  const document = elements.get(refs(/\/K \[([^\]]*)\]/.exec(root.body)[1])[0]);
  assert.equal(document.role, "Document"); assert.equal(document.parent, rootId);
  const entries = [];
  function numberTree(id, root = true) {
    const body = pdf.objects.get(id).body, kids = /\/Kids \[([^\]]*)\]/.exec(body);
    const before = entries.length;
    if (kids) {
      const nodes = refs(kids[1]); assert.ok(nodes.length <= 64);
      for (const child of nodes) numberTree(child, false);
    } else {
      const numbers = /\/Nums \[([^\]]*)\]/.exec(body)[1];
      const values = [...numbers.matchAll(/(\d+) (\d+) 0 R/g)].map(match => [Number(match[1]), Number(match[2])]);
      assert.ok(values.length <= 64); entries.push(...values);
    }
    if (!root) assert.match(body, new RegExp("/Limits \\[" + entries[before][0] + " " + entries.at(-1)[0] + "\\]"));
  }
  numberTree(ref(root.body, "ParentTree"));
  const pageEntries = new Map(entries), keys = entries.map(([key]) => key);
  assert.equal(new Set(keys).size, keys.length, "Page and annotation ParentTree keys must be unique");
  assert.deepEqual(keys, [...keys].sort((left, right) => left - right));
  assert.equal(Number(/\/ParentTreeNextKey (\d+)/.exec(root.body)[1]), Math.max(...keys) + 1);
  const marked = new Map();
  const associatedAnnotations = new Set();
  for (const page of pdf.pages) {
    assert.match(page.body, /\/Tabs \/S/);
    const key = Number(/\/StructParents (\d+)/.exec(page.body)[1]);
    const parents = refs(pdf.objects.get(pageEntries.get(key)).body);
    const content = [...page.content.matchAll(/\/(\w+) << \/MCID (\d+) >> BDC\n([\s\S]*?)EMC/g)];
    assert.equal(content.length, parents.length);
    for (const [index, match] of content.entries()) {
      const mcid = Number(match[2]); assert.equal(mcid, index);
      const owner = elements.get(parents[mcid]); assert.equal(owner.role, match[1]);
      assert.ok(owner.kids.includes("<< /Type /MCR /Pg " + page.id + " 0 R /MCID " + mcid + " >>"));
      marked.set(page.id + ":" + mcid, { owner: owner.id, text: [...match[3].matchAll(/<([a-f0-9]*)> Tj/g)].map(m => Buffer.from(m[1], "hex").toString("latin1")).join("") });
    }
    assert.equal((page.content.match(/\bBDC\b/g) ?? []).length, (page.content.match(/\bEMC\b/g) ?? []).length);
    for (const id of refs(/\/Annots \[([^\]]*)\]/.exec(page.body)?.[1] ?? "")) {
      const annotation = pdf.objects.get(id), key = Number(/\/StructParent (\d+)/.exec(annotation.body)?.[1]);
      const owner = elements.get(pageEntries.get(key));
      assert.equal(owner?.role, "Link", "Every tagged annotation must point to its logical Link");
      assert.equal(ref(annotation.body, "P"), page.id);
      assert.ok(owner.kids.includes("<< /Type /OBJR /Obj " + id + " 0 R /Pg " + page.id + " 0 R >>"));
      assert.ok(["P", "H1", "H2", "H3", "H4", "H5", "H6", "TD", "TH"].includes(elements.get(owner.parent)?.role));
      assert.ok(!associatedAnnotations.has(id)); associatedAnnotations.add(id);
    }
  }
  let contentCount = 0;
  for (const element of elements.values()) {
    const children = element.kids.replace(/<<[\s\S]*?>>/g, "");
    for (const child of refs(children)) assert.equal(elements.get(child).parent, element.id);
    element.content = [...element.kids.matchAll(/<< \/Type \/MCR \/Pg (\d+) 0 R \/MCID (\d+) >>/g)].map(match => {
      const entry = marked.get(match[1] + ":" + match[2]); assert.equal(entry.owner, element.id); contentCount++; return entry;
    });
  }
  assert.equal(contentCount, marked.size);
  for (const element of elements.values()) for (const match of element.kids.matchAll(/<< \/Type \/OBJR \/Obj (\d+) 0 R \/Pg (\d+) 0 R >>/g))
    assert.ok(associatedAnnotations.has(Number(match[1])), "Every OBJR must resolve to one annotated page");
  return { ...pdf, elements, document, marked };
}

test("tagged composed reports retain logical paragraphs and table cells across page fragments", async () => {
  const long = Array.from({ length: 45 }, (_, index) => "Left" + index).join("\n");
  const blocks = [
    { kind: "paragraph", text: "Report", headingLevel: 1 },
    { kind: "paragraph", text: Array.from({ length: 22 }, (_, i) => "Paragraph" + i).join("\n") },
    pdfTable([{ name: long, amount: "Right" }, { name: "Last", amount: 9 }], { columns, footer: { values: ["Total"], totals: { amount: "sum" } } }),
    { kind: "image", image, alternativeText: "Blue square" },
    { ...vector, alternativeText: "Green triangle" },
    { ...vector, decorative: true },
    { kind: "paragraph", text: "End", headingLevel: 2 }
  ];
  const options = { tagged, pageSize: { width: 220, height: 180 }, margins: 24, pageHeader: "Running", pageFooter: "Footer", compression: false };
  const blob = await writePdfDocument(blocks, options), pdf = await inspectStructure(blob);
  assert.deepEqual(Buffer.from(await blob.arrayBuffer()), Buffer.from(await (await writePdfDocument(blocks, options)).arrayBuffer()), "Tagged output is deterministic");
  assert.ok(pdf.pages.length > 5);
  const ordered = refs(pdf.document.kids).map(id => pdf.elements.get(id));
  assert.deepEqual(ordered.map(element => element.role), ["H1", "P", "Table", "Figure", "Figure", "H2"]);
  const paragraph = ordered[1]; assert.equal(paragraph.content.length, 22); assert.equal(paragraph.content.map(part => part.text).join("|"), Array.from({ length: 22 }, (_, i) => "Paragraph" + i).join("|"));
  const table = ordered[2], rows = refs(table.kids).map(id => pdf.elements.get(id));
  assert.equal(rows.length, 4, "One header, two data rows and one ordinary footer; no repeated semantic headers");
  assert.ok(rows.every(row => row.role === "TR"));
  const header = refs(rows[0].kids).map(id => pdf.elements.get(id));
  assert.ok(header.every(cell => cell.role === "TH" && cell.body.includes("/A << /O /Table /Scope /Column >>")));
  const cells = refs(rows[1].kids).map(id => pdf.elements.get(id));
  assert.equal(cells[0].role, "TD"); assert.equal(cells[0].content.length, 45);
  assert.equal(cells[0].content.map(part => part.text).join("|"), long.replaceAll("\n", "|"));
  assert.equal(cells[1].content.map(part => part.text).join(""), "Right");
  assert.match(ordered[3].body, /\/Alt <feff0042006c007500650020007300710075006100720065>/);
  assert.equal([...pdf.elements.values()].filter(element => element.role === "Figure").length, 2);
  assert.ok(pdf.pages.every(page => page.content.includes("/Artifact << /Type /Pagination >> BDC")));
  assert.ok(![...pdf.marked.values()].some(part => /Running|Footer|Page /.test(part.text)), "Decorations are excluded from logical content");
});

test("tagged table convenience exports share structure ownership and preserve empty leaf cells", async () => {
  const pdf = await inspectStructure(await writePdf([{ name: "Item", amount: 3 }], { columns: [{ header: "", key: "name" }, columns[1]], tagged,
    title: "Title", messageTop: "Before", messageBottom: "After", pageNumbers: false, compression: false }));
  assert.deepEqual(refs(pdf.document.kids).map(id => pdf.elements.get(id).role), ["H1", "P", "Table", "P"]);
  assert.equal([...pdf.elements.values()].filter(element => element.role === "TH").length, 1);
  assert.equal([...pdf.elements.values()].filter(element => element.role === "TD").length, 3);
  const untagged = await inspectPdf(await writePdfDocument([{ kind: "paragraph", text: "Plain" }], { compression: false }));
  assert.ok(![...untagged.objects.values()].some(object => /\/StructTreeRoot|\/StructParents|\/Marked|\/MCID|\/Artifact/.test(object.body)));
});

test("tagged profile rejects unsupported semantics instead of silently flattening them", async () => {
  for (const options of [{ columns: [{ header: "Name", key: "name", groups: ["Group"] }] },
    { columns, headerRows: [[{ value: "Group", columnSpan: 2 }, null]] },
    { columns, footer: { rows: [[{ value: "Total", columnSpan: 2 }, null]] } }])
    await assert.rejects(writePdf([{ name: "Item", amount: 1 }], { ...options, tagged }), /Tagged PDF supports leaf/);
  await assert.rejects(writePdf([{ name: new ExportCell(" ", { link: { target: "https://example.com" } }), amount: 1 }], { columns, tagged }), /nonempty displayed text/);
  for (const block of [{ kind: "image", image }, vector, { ...vector, alternativeText: " " }, { ...vector, alternativeText: "Chart", decorative: true }])
    await assert.rejects(writePdfDocument([block], { tagged }), /alternativeText|decorative/);
  for (const language of ["", "not_a_tag", "en--US", "en-US\0", 1])
    await assert.rejects(writePdfDocument([], { tagged: { language } }), /tagged.language/);
  await assert.rejects(writePdfDocument([{ kind: "paragraph", text: "Heading", headingLevel: 7 }], { tagged }), /headingLevel/);
});

test("tagged paragraph and table links associate text and each cross-page native annotation", async () => {
  const long = "Linked report details ".repeat(65), cellText = "Linked table line\n".repeat(30);
  const pdf = await inspectStructure(await writePdfDocument([
    { kind: "paragraph", text: [{ text: "Before " }, { text: long, bold: true, link: { target: "https://example.com/report", tooltip: "Report details" } }, { text: "after" }] },
    pdfTable([{ name: new ExportCell(cellText, { text: cellText, link: { target: "https://example.com/cell" } }), amount: 7 }], {
      columns, formatRuns: (value, context) => context.columnIndex === 1
        ? [{ text: String(value), link: { target: "https://example.com/amount" } }] : [{ text: String(value ?? "") }],
      footer: { values: [new ExportCell("Footer", { text: "Footer", link: { target: "https://example.com/footer" } })], totals: { amount: "sum" } }
    })
  ], { tagged, pageSize: { width: 220, height: 180 }, margins: 24, pageNumbers: false, compression: false }));
  const links = [...pdf.elements.values()].filter(element => element.role === "Link");
  assert.equal(links.length, 5, "Each authored paragraph/cell run keeps one logical link identity");
  const paragraphLink = links.find(element => pdf.elements.get(element.parent).role === "P");
  assert.equal(paragraphLink.content.map(part => part.text).join(""), long);
  assert.ok(new Set([...paragraphLink.kids.matchAll(/\/Pg (\d+) 0 R/g)].map(match => match[1])).size > 1);
  const cellLink = links.find(element => element.content.map(part => part.text).join("") === cellText.replaceAll("\n", ""));
  assert.ok(cellLink && cellLink.content.length > 20);
  assert.ok([...pdf.objects.values()].filter(object => /\/Subtype \/Link/.test(object.body)).length > links.length);
  await assert.rejects(writePdfDocument([{ kind: "paragraph", text: [{ text: " ", link: { target: "https://example.com" } }] }], { tagged }), /nonempty displayed text/);
  await assert.rejects(writePdfDocument([{ kind: "paragraph", text: [{ text: "Link", link: { target: "https://example.com" } }] }],
    { tagged, limits: { maxHyperlinks: 0 }, pageNumbers: false }), /maxHyperlinks/);
});

test("tagged structure and marked-content budgets fail visibly and release active sources", async () => {
  const blocks = () => [{ kind: "paragraph", text: "Intro" }, pdfTable([{ name: "Item", amount: 1 }], { columns })];
  const limits = { maxStructureElements: 9, maxMarkedContentReferences: 5 };
  await inspectStructure(await writePdfDocument(blocks(), { tagged, limits, pageNumbers: false, compression: false }));
  for (const [key, value] of [["maxStructureElements", 8], ["maxMarkedContentReferences", 4]])
    await assert.rejects(writePdfDocument(blocks(), { tagged, limits: { ...limits, [key]: value }, pageNumbers: false }), new RegExp(key));
  let returned = 0, blocksReturned = 0;
  const rows = { *[Symbol.iterator]() { try { yield { name: "One", amount: 1 }; yield { name: "Two", amount: 2 }; } finally { returned++; } } };
  const source = { *[Symbol.iterator]() { try { yield pdfTable(rows, { columns }); } finally { blocksReturned++; } } };
  const destination = new WritableStream({ write() {} });
  await assert.rejects(writePdfDocumentTo(source, destination, { tagged, limits: { maxStructureElements: 8 }, pageNumbers: false }), /maxStructureElements/);
  assert.equal(returned, 1); assert.equal(blocksReturned, 1); assert.equal(destination.locked, false);
  await assert.rejects(writePdfDocument([{ ...vector, alternativeText: "Described" }], { tagged, limits: { maxCellCharacters: 3 } }), /maxCellCharacters/);
});

test("graphic alternative getters preserve the validated text and decorative intent", async () => {
  let textReads = 0, decorativeReads = 0;
  const described = { ...vector, get alternativeText() { return ++textReads === 1 ? "A" : "BBBB"; } };
  const decorative = { ...vector, get decorative() { return ++decorativeReads === 1 ? true : false; } };
  const pdf = await inspectPdf(await writePdfDocument([described, decorative], {
    tagged, compression: false, pageNumbers: false, limits: { maxCellCharacters: 1, maxCells: 1, maxTextCharacters: 1 }
  }));
  const figures = [...pdf.objects.values()].filter(o => /\/S \/Figure /.test(o.body));
  assert.equal(figures.length, 1);
  assert.match(figures[0].body, /\/Alt <feff0041>/);
  assert.doesNotMatch(figures[0].body, /0042004200420042/);
  assert.equal((pdf.pages[0].content.match(/\/Artifact << \/Type \/Layout >> BDC/g) ?? []).length, 1);
});

test("composed paragraph getters retain validated heading and text values", async () => {
  let textReads = 0, levelReads = 0, presentationReads = 0;
  const paragraph = { kind: "paragraph", get text() { return ++textReads === 1 ? "A" : "BBBB"; },
    get headingLevel() { return ++levelReads === 1 ? 2 : 9; },
    get presentation() { return ++presentationReads === 1 ? { bold: true } : { background: "FF0000" }; } };
  const pdf = await inspectPdf(await writePdfDocument([paragraph], {
    tagged, compression: false, pageNumbers: false, limits: { maxCellCharacters: 1, maxTextCharacters: 1 }
  }));
  assert.equal(pdf.text, "A");
  const heading = [...pdf.objects.values()].find(o => /\/S \/H2 /.test(o.body));
  assert.ok(heading, "The captured heading level owns the emitted text");
  assert.doesNotMatch([...pdf.objects.values()].map(o => o.body).join("\n"), /\/S \/H9 /);
});

test("paragraph presentation snapshots nested getters before validating its emitted style", async () => {
  const reads = { bold: 0, italic: 0, color: 0, alignment: 0 };
  const presentation = {
    get bold() { return ++reads.bold <= 2 ? false : true; },
    get italic() { return ++reads.italic <= 2 ? false : true; },
    get color() { return ++reads.color <= 2 ? "008000" : "0000FF"; },
    get alignment() { return ++reads.alignment <= 2 ? "left" : "right"; }
  };
  const pdf = await inspectPdf(await writePdfDocument([{ kind: "paragraph", text: "A", presentation }], { compression: false, pageNumbers: false }));
  assert.deepEqual(reads, { bold: 1, italic: 1, color: 1, alignment: 1 });
  assert.match(pdf.pages[0].content, /q 0 0\.502 0 rg 0 0\.502 0 RG/);
  assert.match(pdf.pages[0].content, /1 0 0 1 36 [\d.]+ Tm <41> Tj/);
  const fonts = [...pdf.objects.values()].filter(o => /\/Type \/Font /.test(o.body));
  assert.ok(fonts.some(font => /\/BaseFont \/Helvetica /.test(font.body)));
  assert.ok(fonts.every(font => !/Helvetica-(?:Bold|Oblique)/.test(font.body)));
  for (const [key, invalid, valid, error] of [["bold", "yes", false, /bold must be boolean/], ["italic", "yes", false, /italic must be boolean/],
    ["color", "invalid", "008000", /six-digit RGB/], ["alignment", "justify", "left", /alignment supports/]]) {
    let count = 0;
    const invalidFirst = Object.defineProperty({}, key, { enumerable: true, get() { return ++count === 1 ? invalid : valid; } });
    await assert.rejects(writePdfDocument([{ kind: "paragraph", text: "A", presentation: invalidFirst }], { pageNumbers: false }), error);
    assert.equal(count, 1);
  }
});

test("cancelling tagged page output returns rows and blocks and releases the caller destination", async () => {
  const abort = new AbortController(); let rowsReturned = 0, blocksReturned = 0, writes = 0;
  async function* rows() { try { for (let i = 0; i < 100; i++) yield { name: "Row " + i, amount: i }; } finally { rowsReturned++; } }
  async function* blocks() { try { yield pdfTable(rows(), { columns }); } finally { blocksReturned++; } }
  const destination = new WritableStream({ write() { writes++; abort.abort(Error("tagged output cancelled")); } });
  await assert.rejects(writePdfDocumentTo(blocks(), destination, { tagged, signal: abort.signal, compression: false, pageSize: "A5" }), /tagged output cancelled/);
  assert.equal(writes, 1); assert.equal(rowsReturned, 1); assert.equal(blocksReturned, 1); assert.equal(destination.locked, false);
});

test("tagged streamed tables finalize completed rows and bound multi-level ParentTree nodes", async () => {
  const chunks = [], shared = new ExportCell("Open record", { text: "Open record", link: { target: "https://example.com/record" } });
  let completedRowsObserved = false, sourceClosed = false;
  async function* rows() {
    try {
      for (let i = 0; i < 2100; i++) {
        if (i === 200) assert.ok(completedRowsObserved, "Finished logical rows must reach the sink before source exhaustion");
        yield { name: shared, amount: shared };
      }
    } finally { sourceClosed = true; }
  }
  const result = await writePdfDocumentTo([pdfTable(rows(), { columns })], { write(bytes) {
    const chunk = bytes.slice(); chunks.push(chunk);
    completedRowsObserved ||= Buffer.from(chunk).toString("latin1").includes("/S /TR ");
  } }, { tagged, compression: false, pageNumbers: false });
  assert.equal(result.rows, 2100); assert.ok(sourceClosed);
  const pdf = await inspectStructure(new Blob(chunks));
  const table = [...pdf.elements.values()].find(element => element.role === "Table");
  assert.equal(refs(table.kids).length, 2101);
  assert.equal([...pdf.elements.values()].filter(element => element.role === "Link").length, 4200);
  assert.equal([...pdf.objects.values()].filter(object => /\/Subtype \/Link/.test(object.body)).length, 4200);
});
