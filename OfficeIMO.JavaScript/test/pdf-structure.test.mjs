import test from "node:test";
import assert from "node:assert/strict";
import { writePdf, writePdfDocument, writePdfDocumentTo, pdfTable, pdfGroupedTable, createGroupedReport, ExportCell } from "../dist/pdf/index.js";
import { inspectPdf, pdfStructureAttributes } from "./pdf-reader.mjs";

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
    id, body: object.body, attributes: pdfStructureAttributes(pdf, object.body), role: /\/S \/(\w+)/.exec(object.body)[1], parent: ref(object.body, "P"),
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
      assert.ok(["P", "H1", "H2", "H3", "H4", "H5", "H6", "TD", "TH", "Caption"].includes(elements.get(owner.parent)?.role));
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
  assert.ok(header.every(cell => cell.role === "TH" && cell.attributes === "<< /O /Table /Scope /Column >>"));
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

test("tagged linked continuation captions retain one first table owner across pages and bands", async () => {
  for (const bands of [false, true]) for (const primary of [false, true]) {
    const reportColumns = ["ID", "Detail", "Amount", "Count"].map(header => ({ header }));
    const rows = Array.from({ length: 5 }, (_, index) => ["R" + index, "Detail", index, 2]);
    const table = pdfTable(rows, { columns: reportColumns, ...(bands ? {
      columnWidths: [60, 90, 90, 90], wideTable: { mode: "bands", repeatColumns: [0] }
    } : {}) });
    const pdf = await inspectStructure(await writePdfDocument([{ ...table,
      ...(primary ? { caption: { text: [{ text: "Primary documentation", link: { target: "https://example.org/main" } }] } } : {}),
      continuationCaption: { text: [{ text: "Continuation documentation", link: { target: "https://example.org/continued" } }] }
    }], { tagged, compression: false, pageNumbers: false, orientation: "landscape", margins: 20,
      pageSize: { width: 300, height: bands ? 220 : 112 }, limits: { maxTableWindowRows: 2 } }));
    assert.ok(pdf.pages.length > 1, "The fixture must paint actual continuation context");
    const owner = [...pdf.elements.values()].find(element => element.role === "Table");
    const children = refs(owner.kids).map(id => pdf.elements.get(id));
    assert.equal(children[0].role, "Caption", "A continuation-only owner also precedes the first logical row");
    assert.equal(children.filter(child => child.role === "Caption").length, 1, "Primary and continuation links share one table caption");
    assert.ok(children.slice(1).every(child => child.role === "TR"));
    const caption = children[0], links = [...pdf.elements.values()].filter(element => element.role === "Link");
    assert.equal(links.length, primary ? 2 : 1);
    assert.ok(links.every(link => link.parent === caption.id));
    const continuation = links.find(link => /\/Alt /.test(link.body));
    assert.match(continuation.body, /\/Alt <feff0043006f006e00740069006e0075006100740069006f006e00200064006f00630075006d0065006e0074006100740069006f006e>/);
    assert.ok((continuation.kids.match(/\/Type \/OBJR/g) ?? []).length > 0);
    const logicalText = [...pdf.marked.values()].map(part => part.text).join("");
    assert.equal(logicalText.includes("Continuation documentation"), false, "Repeated context remains an artifact");
    assert.equal(logicalText.split("Primary documentation").length - 1, primary ? 1 : 0);
    for (const row of rows) assert.equal(logicalText.split(row[0]).length - 1, 1, "Source identifiers retain one logical occurrence");
  }
});

test("unused linked continuation captions create no phantom logical owners", async () => {
  for (const bands of [false, true]) {
    const pdf = await inspectStructure(await writePdfDocument([{ ...pdfTable([["R1", 7, 9]], {
      columns: ["ID", "Amount", "Count"].map(header => ({ header })),
      ...(bands ? { columnWidths: [80, 120, 120], wideTable: { mode: "bands", repeatColumns: [0] } } : {}) }),
      continuationCaption: { text: [{ text: "Unused navigation", link: { target: "https://example.org/unused" } }] }
    }], { tagged, compression: false, pageNumbers: false, pageSize: { width: 300, height: 220 },
      orientation: "landscape", margins: 20, limits: { maxTableWindowRows: 2 } }));
    assert.equal(pdf.pages.length, bands ? 2 : 1, "Every physical band still belongs to the first logical window");
    assert.equal([...pdf.elements.values()].filter(element => element.role === "Caption" || element.role === "Link").length, 0);
    assert.equal([...pdf.objects.values()].filter(object => /\/Subtype \/Link /.test(object.body)).length, 0);
    assert.ok([...pdf.elements.values()].filter(element => element.role === "Table").every(table =>
      refs(table.kids).every(id => pdf.elements.get(id).role === "TR")));
  }
});

test("tagged profile validates authored link, figure and language semantics", async () => {
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

test("cached nonband table headings consume physical cell/text budgets without recapturing source callbacks", async () => {
  for (const semantic of [false, true]) for (const composed of [false, true]) {
    const run = async limits => {
      let headerReads = 0, formatted = 0, returned = 0;
      const table = { columns: [
        { key: "left", get header() { headerReads++; return "HEAD"; } },
        { key: "right", get header() { headerReads++; return "TAIL"; } }
      ], formatValue(value) { formatted++; return String(value); } };
      const rows = { *[Symbol.iterator]() { try { yield { left: "A", right: "C" }; yield { left: "B", right: "D" }; } finally { returned++; } } };
      const settings = { pageSize: { width: 200, height: 60 }, orientation: "landscape", fontSize: 10, padding: 2, margins: 10,
        pageNumbers: false, compression: false, ...(semantic ? { tagged } : {}), limits };
      try {
        return composed ? await writePdfDocument([pdfTable(rows, table)], settings) : await writePdf(rows, { ...table, ...settings });
      } finally {
        assert.equal(headerReads, 2, "Each authored heading is captured once across continuation pages");
        assert.equal(formatted, 4, "Repeated headings never call the data formatter again");
        assert.equal(returned, 1, "A budget failure closes the active source");
      }
    };
    const pdf = await inspectPdf(await run({ maxCells: 100, maxTextCharacters: 100 }));
    assert.equal(pdf.pages.length, 2, "The fixture must physically repeat both cached headings");
    for (const heading of ["HEAD", "TAIL"])
      assert.equal(pdf.pages.flatMap(page => page.lines).filter(line => line === heading).length, 2);
    // Six original cells contain twelve characters; either ceiling must reject their physical repeats.
    for (const limits of [{ maxCells: 6 }, { maxTextCharacters: 12 }])
      await assert.rejects(run(limits), error => error.code === "RESOURCE_LIMIT" && error.message.includes(Object.keys(limits)[0]));
  }
});

test("tagged leaf subtotal shares captured headings and charges only actually painted content", async () => {
  const run = async limits => {
    let headerReads = 0, projected = 0, subtotalLabels = 0, continuationLabels = 0, completed = false;
    const declared = Array.from({ length: 8 }, (_, index) => ({ key: "m" + index, type: "number",
      get header() { headerReads++; return "Metric" + index + "-" + "X".repeat(32); },
      value(row) { projected++; return row[index]; } }));
    const report = createGroupedReport([Array(8).fill(1)], { columns: declared,
      groupBy: [{ key: "team", value: () => "Group-A" }], totals: { m0: "sum" } });
    try {
      const document = await writePdfDocument([pdfGroupedTable(report, { labelColumn: { header: "Group" }, includeGrandTotal: false,
        subtotalLabel(group) { subtotalLabels++; return "Subtotal " + group.label; }, grandTotalLabel: () => "Unused",
        continuationLabel(group) { continuationLabels++; return group.label; } })], {
        tagged, pageSize: { width: 1000, height: 600 }, orientation: "landscape", margins: 20,
        pageNumbers: false, compression: false, limits
      });
      completed = true; return document;
    } finally {
      assert.equal(headerReads, 8, "The leaf subtotal reuses captured headings without invoking authored getters");
      assert.ok(projected <= 8, "The record is never projected twice, including early budget failures");
      assert.ok(subtotalLabels <= 1);
      assert.ok(continuationLabels <= 1, "An early resource rejection must not recapture continuation text");
      if (completed) {
        assert.equal(projected, 8, "A completed export projects the one record exactly once");
        assert.equal(subtotalLabels, 1); assert.equal(continuationLabels, 1);
      }
    }
  };
  const pdf = await inspectStructure(await run({ maxCells: 34, maxTextCharacters: 500 }));
  assert.equal(pdf.pages.length, 1, "The detail and its leaf subtotal share one physical table");
  assert.equal([...pdf.elements.values()].filter(element => element.role === "Table").length, 1);
  const heading = "Metric0-" + "X".repeat(32);
  assert.equal([...pdf.elements.values()].filter(element => element.role === "TH" && element.content.map(part => part.text).join("") === heading).length, 1);
  for (const limits of [{ maxCells: 26 }, { maxTextCharacters: 320 }])
    await assert.rejects(run(limits), error => error.code === "RESOURCE_LIMIT" && error.message.includes(Object.keys(limits)[0]));
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

test("scalar and rich paragraph presentations capture getters before validating their emitted style", async () => {
  for (const text of ["A", [{ text: "A" }]]) {
    const reads = { bold: 0, italic: 0, color: 0, alignment: 0 };
    const presentation = {
      get bold() { return ++reads.bold <= 2 ? false : true; },
      get italic() { return ++reads.italic <= 2 ? false : true; },
      get color() { return ++reads.color <= 2 ? "008000" : "0000FF"; },
      get alignment() { return ++reads.alignment <= 2 ? "left" : "right"; }
    };
    const pdf = await inspectPdf(await writePdfDocument([{ kind: "paragraph", text, presentation }], { compression: false, pageNumbers: false }));
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
      await assert.rejects(writePdfDocument([{ kind: "paragraph", text, presentation: invalidFirst }], { pageNumbers: false }), error);
      assert.equal(count, 1);
    }
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
