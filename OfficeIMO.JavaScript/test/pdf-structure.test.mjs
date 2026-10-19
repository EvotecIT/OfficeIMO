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
  const numbers = pdf.objects.get(ref(root.body, "ParentTree")).body;
  const pageEntries = new Map([...numbers.matchAll(/(\d+) (\d+) 0 R/g)].map(match => [Number(match[1]), Number(match[2])]));
  const marked = new Map();
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
  await assert.rejects(writePdf([{ name: new ExportCell("Linked", { link: { target: "https://example.com" } }), amount: 1 }], { columns, tagged }), /Tagged PDF link/);
  await assert.rejects(writePdfDocument([pdfTable([], { columns, footer: { values: [new ExportCell("Linked", { link: { target: "https://example.com" } })] } })], { tagged }), /Tagged PDF link/);
  for (const block of [{ kind: "image", image }, vector, { ...vector, alternativeText: " " }, { ...vector, alternativeText: "Chart", decorative: true }])
    await assert.rejects(writePdfDocument([block], { tagged }), /alternativeText|decorative/);
  for (const language of ["", "not_a_tag", "en--US", "en-US\0", 1])
    await assert.rejects(writePdfDocument([], { tagged: { language } }), /tagged.language/);
  await assert.rejects(writePdfDocument([{ kind: "paragraph", text: "Heading", headingLevel: 7 }], { tagged }), /headingLevel/);
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

test("cancelling tagged page output returns rows and blocks and releases the caller destination", async () => {
  const abort = new AbortController(); let rowsReturned = 0, blocksReturned = 0, writes = 0;
  async function* rows() { try { for (let i = 0; i < 100; i++) yield { name: "Row " + i, amount: i }; } finally { rowsReturned++; } }
  async function* blocks() { try { yield pdfTable(rows(), { columns }); } finally { blocksReturned++; } }
  const destination = new WritableStream({ write() { writes++; abort.abort(Error("tagged output cancelled")); } });
  await assert.rejects(writePdfDocumentTo(blocks(), destination, { tagged, signal: abort.signal, compression: false, pageSize: "A5" }), /tagged output cancelled/);
  assert.equal(writes, 1); assert.equal(rowsReturned, 1); assert.equal(blocksReturned, 1); assert.equal(destination.locked, false);
});
