import test from "node:test";
import assert from "node:assert/strict";
import { writePdf, writePdfDocument, pdfTable, writeGroupedPdf } from "../dist/pdf/index.js";
import { createGroupedReport } from "../dist/core/index.js";
import { inspectPdf, pdfStructureAttributes } from "./pdf-reader.mjs";
import { pdfStructureChildren } from "./pdf-structure-reader.mjs";

const references = body => [...body.matchAll(/(\d+) 0 R/g)].map(match => Number(match[1]));
const reference = (body, name) => Number(new RegExp("/" + name + " (\\d+) 0 R").exec(body)?.[1]);
const headerIds = body => [...(/\/Headers \[([^\]]*)\]/.exec(body)?.[1] ?? "").matchAll(/\(([^)]*)\)/g)].map(match => match[1]);

async function tableStructure(blob) {
  const pdf = await inspectPdf(blob), elements = new Map();
  for (const [id, object] of pdf.objects) if (/\/Type \/StructElem /.test(object.body)) {
    const text = pdfStructureChildren(object.body).entries.filter(entry => entry.kind === "mark").map(entry => {
      const page = pdf.pages.find(page => page.id === entry.pageId);
      assert.ok(page, "Every semantic cell reference resolves to a physical page");
      const marked = new RegExp("/\\w+ << /MCID " + entry.mcid + "(?: /ActualText <[^>]*>)? >> BDC\\n([\\s\\S]*?)EMC").exec(page.content);
      assert.ok(marked, "Every semantic cell reference resolves to marked page content");
      return [...marked[1].matchAll(/<([a-f0-9]*)> Tj/g)].map(value => Buffer.from(value[1], "hex").toString("latin1")).join("");
    }).join("");
    const attributes = pdfStructureAttributes(pdf, object.body);
    elements.set(id, { id, body: object.body, attributes, role: /\/S \/(\w+)/.exec(object.body)[1], text,
      identifier: /\/ID \(([^)]*)\)/.exec(object.body)?.[1], headers: headerIds(attributes) });
  }
  const catalog = [...pdf.objects.values()].find(object => /\/Type \/Catalog /.test(object.body));
  const root = pdf.objects.get(reference(catalog.body, "StructTreeRoot"));
  const entries = [];
  function names(id, top = true) {
    const body = pdf.objects.get(id).body, before = entries.length, kids = /\/Kids \[([^\]]*)\]/.exec(body);
    if (kids) {
      const children = references(kids[1]); assert.ok(children.length <= 64);
      for (const child of children) names(child, false);
    } else {
      const values = [.../\/Names \[([^\]]*)\]/.exec(body)[1].matchAll(/\(([^)]*)\) (\d+) 0 R/g)];
      assert.ok(values.length <= 64); entries.push(...values.map(match => [match[1], Number(match[2])]));
    }
    if (!top) assert.match(body, new RegExp("/Limits \\[\\(" + entries[before][0] + "\\) \\(" + entries.at(-1)[0] + "\\)\\]"));
  }
  names(reference(root.body, "IDTree"));
  assert.deepEqual(entries.map(([key]) => key), entries.map(([key]) => key).sort());
  assert.equal(new Set(entries.map(([key]) => key)).size, entries.length);
  const indexed = new Map(entries);
  for (const element of elements.values()) {
    if (element.identifier) assert.equal(indexed.get(element.identifier), element.id, "Every structure ID resolves through the name tree");
    for (const key of element.headers) assert.equal(elements.get(indexed.get(key))?.role, "TH", "Header associations resolve to actual header elements");
  }
  return { ...pdf, elements, indexed };
}

test("tagged span matrices associate each data/footer cell with its actual hierarchical column headers", async () => {
  const columns = ["ID", "Amount", "Count", "Healthy", "Offline"].map(header => ({ header }));
  const pdf = await tableStructure(await writePdf([["SVC-1", 12, 3, 7, 2]], {
    columns, tagged: { language: "en" }, compression: false, pageNumbers: false,
    headerRows: [
      [{ value: "Identifier", rowSpan: 3 }, { value: "Business", columnSpan: 4 }, null, null, null],
      [null, { value: "Metrics", columnSpan: 2 }, null, { value: "Availability", columnSpan: 2 }, null],
      [null, { value: "Amount" }, { value: "Count" }, { value: "Healthy" }, { value: "Offline" }]
    ],
    footer: { rows: [
      [{ value: "Totals", columnSpan: 3 }, null, null, { value: 7, rowSpan: 2 }, { value: 2, rowSpan: 2 }],
      [{ value: "Reviewed", columnSpan: 3 }, null, null, null, null]
    ] }
  }));
  const headers = [...pdf.elements.values()].filter(element => element.role === "TH");
  const ids = new Map(headers.map(element => [element.text, element.identifier]));
  assert.equal(headers.length, 8);
  assert.match(headers.find(element => element.text === "Identifier").attributes, /\/Scope \/Column \/RowSpan 3/);
  assert.match(headers.find(element => element.text === "Business").attributes, /\/Scope \/Column \/ColSpan 4/);
  const expected = names => names.map(name => ids.get(name));
  assert.deepEqual(headers.find(element => element.text === "Metrics").headers, expected(["Business"]));
  assert.deepEqual(headers.find(element => element.text === "Healthy").headers, expected(["Business", "Availability"]));
  const data = [...pdf.elements.values()].filter(element => element.role === "TD");
  assert.deepEqual(data.find(element => element.text === "SVC-1").headers, expected(["Identifier"]));
  assert.deepEqual(data.find(element => element.text === "12").headers, expected(["Business", "Metrics", "Amount"]));
  const total = data.find(element => element.text === "Totals");
  assert.match(total.attributes, /\/ColSpan 3/);
  assert.deepEqual(total.headers, expected(["Identifier", "Business", "Metrics", "Amount", "Count"]));
  const reviewed = data.find(element => element.text === "Reviewed");
  assert.match(reviewed.attributes, /\/ColSpan 3/);
  assert.deepEqual(reviewed.headers, total.headers);
  assert.equal([...pdf.elements.values()].filter(element => element.role === "TR").length, 6);
});

test("many composed tables retain unique resolvable header IDs in a bounded sorted name tree", async () => {
  const columns = [{ header: "Record" }, { header: "Amount" }];
  const pdf = await tableStructure(await writePdfDocument(Array.from({ length: 40 }, (_, index) =>
    pdfTable([["R" + index, index]], { columns })), { tagged: { language: "en" }, compression: false, pageNumbers: false }));
  assert.equal(pdf.indexed.size, 80);
  assert.equal([...pdf.elements.values()].filter(element => element.role === "Table").length, 40);
  assert.equal([...pdf.elements.values()].filter(element => element.role === "TH").length, 80);
  assert.equal([...pdf.elements.values()].filter(element => element.role === "TD").length, 80);
  for (const cell of [...pdf.elements.values()].filter(element => element.role === "TD")) assert.equal(cell.headers.length, 1);
});

test("ordinary table cells share immutable column attributes across rows and retain table-local header identities", async () => {
  const columns = [{ header: "Record" }, { header: "Amount" }];
  for (const compression of [false, true]) {
    const pdf = await tableStructure(await writePdfDocument([
      pdfTable([["A1", 1], ["A2", 2], ["A3", 3]], { columns }),
      pdfTable([["B1", 4], ["B2", 5], ["B3", 6]], { columns })
    ], { tagged: { language: "en" }, compression, pageNumbers: false }));
    const tables = [...pdf.elements.values()].filter(element => element.role === "Table"), shared = new Set();
    assert.equal(tables.length, 2);
    for (const table of tables) {
      const rows = references(/\/K \[([^\]]*)\]/.exec(table.body)[1]).map(id => pdf.elements.get(id));
      const cells = row => references(/\/K \[([^\]]*)\]/.exec(row.body)[1]).map(id => pdf.elements.get(id));
      const headers = cells(rows[0]);
      assert.deepEqual(headers.map(cell => cell.text), ["Record", "Amount"]);
      for (let column = 0; column < columns.length; column++) {
        const data = rows.slice(1).map(row => cells(row)[column]), ids = data.map(cell => reference(cell.body, "A"));
        assert.ok(ids.every(Number.isSafeInteger), "Every ordinary cell points to an indirect attribute dictionary");
        assert.equal(new Set(ids).size, 1, "Rows reuse the same immutable attributes for their column");
        assert.ok(!shared.has(ids[0]), "Another table or column cannot reuse this table's header association");
        shared.add(ids[0]);
        assert.match(pdf.objects.get(ids[0]).body, /^<< \/O \/Table \/Headers \[/);
        for (const cell of data) assert.deepEqual(cell.headers, [headers[column].identifier]);
      }
    }
    assert.equal(shared.size, 4);
  }
});

test("explicit detail row headers associate preceding cells without changing footer or column-heading roles", async () => {
  const columns = [{ header: "Wartosc", key: "amount", type: "number" },
    { header: "Usluga", key: "service", rowHeader: true }, { header: "Miejsce", key: "site", rowHeader: true }];
  const pdf = await tableStructure(await writePdf([[12, "SVC-1", "HQ"]], {
    columns, tagged: { language: "pl" }, compression: false, pageNumbers: false,
    footer: { rows: [[{ value: "Totals", columnSpan: 2 }, null, { value: 12 }]] }
  }));
  const elements = [...pdf.elements.values()], rowHeaders = elements.filter(element => /\/Scope \/Row/.test(element.attributes));
  assert.deepEqual(rowHeaders.map(element => element.text), ["SVC-1", "HQ"]);
  const ids = rowHeaders.map(element => element.identifier);
  const amount = elements.find(element => element.role === "TD" && element.text === "12");
  assert.deepEqual(amount.headers.slice(-2), ids, "A row heading after its metric still identifies that earlier data cell");
  for (const heading of rowHeaders) assert.ok(heading.headers.every(id => !ids.includes(id)), "Row headings do not create circular associations");
  const footer = elements.find(element => element.text === "Totals");
  assert.equal(footer.role, "TD");
  assert.ok(footer.headers.every(id => !ids.includes(id)), "Footer matrices do not inherit a detail record's identifiers");
  const row = elements.find(element => element.id === reference(amount.body, "P"));
  assert.deepEqual(references(/\/K \[([^\]]*)\]/.exec(row.body)[1]), [amount.id, ...rowHeaders.map(element => element.id)]);
});

test("captured key-selected bands retain one row heading across repeated visual identifiers", async () => {
  let roleReads = 0, selectorReads = 0, formatted = 0, opened = 0;
  const identifier = Object.assign(Object.create({ get rowHeader() { roleReads++; return true; } }), { header: "Usluga", key: "0" });
  const columns = [{ header: "Wartosc", key: "amount" }, identifier, { header: "Miejsce", key: "site" }, { header: "Opis", key: "detail" }];
  const selectors = ["site", 1], policy = { mode: "bands", get repeatColumns() { selectorReads++; return selectors; } };
  function* rows() { opened++; yield [12, "SVC-1", "HQ", "Detail"]; }
  const block = pdfTable(rows(), { columns, columnWidths: [80, 80, 80, 80], wideTable: policy,
    formatValue(value) { formatted++; return String(value); } });
  identifier.key = "changed"; selectors.length = 0;
  const pdf = await tableStructure(await writePdfDocument([block], { tagged: { language: "pl" }, compression: false,
    pageNumbers: false, pageSize: { width: 330, height: 210 }, orientation: "landscape", margins: 20 }));
  assert.equal(roleReads, 1); assert.equal(selectorReads, 1); assert.equal(opened, 1); assert.equal(formatted, 4);
  assert.equal(pdf.pages.length, 2);
  for (const page of pdf.pages) assert.deepEqual(page.lines.slice(0, 2), ["Usluga", "Miejsce"], "Repeated prefix follows selected source order");
  const rowHeaders = [...pdf.elements.values()].filter(element => /\/Scope \/Row/.test(element.attributes));
  assert.equal(rowHeaders.length, 1); assert.equal(rowHeaders[0].text, "SVC-1");
  assert.ok(pdf.pages.every(page => page.lines.includes("SVC-1") && page.lines.includes("HQ")));
  for (const cell of [...pdf.elements.values()].filter(element => element.role === "TD")) assert.ok(cell.headers.includes(rowHeaders[0].identifier));
});

test("grouped summaries use their authored label rather than inheriting numeric detail row-header roles", async () => {
  const columns = [{ header: "Name", key: "name", rowHeader: true }, { header: "Amount", key: "amount", type: "number", rowHeader: true }];
  const source = createGroupedReport([["A", 12], ["A", 1]], { columns, groupBy: "name", totals: { amount: "sum" } });
  const pdf = await tableStructure(await writeGroupedPdf(source, { labelColumn: { header: "Summary" },
    subtotalLabel: () => "Group A", grandTotalLabel: () => "All records", tagged: { language: "en" }, compression: false, pageNumbers: false }));
  const elements = [...pdf.elements.values()];
  for (const text of ["Group A", "All records"]) {
    const label = elements.find(element => element.text === text && /\/Scope \/Row/.test(element.attributes));
    assert.ok(label);
    const row = reference(label.body, "P"), siblings = elements.filter(element => reference(element.body, "P") === row);
    assert.equal(siblings.filter(element => element.role === "TH").length, 1);
    assert.ok(siblings.filter(element => element.id !== label.id).every(element => element.role === "TD"));
    assert.ok(siblings.some(element => element.text === "13" && element.headers.includes(label.identifier)));
  }
});
