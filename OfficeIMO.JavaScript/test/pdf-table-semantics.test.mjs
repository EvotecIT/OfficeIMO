import test from "node:test";
import assert from "node:assert/strict";
import { writePdf, writePdfDocument, pdfTable } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const references = body => [...body.matchAll(/(\d+) 0 R/g)].map(match => Number(match[1]));
const reference = (body, name) => Number(new RegExp("/" + name + " (\\d+) 0 R").exec(body)?.[1]);
const headerIds = body => [...(/\/Headers \[([^\]]*)\]/.exec(body)?.[1] ?? "").matchAll(/\(([^)]*)\)/g)].map(match => match[1]);

async function tableStructure(blob) {
  const pdf = await inspectPdf(blob), elements = new Map();
  for (const [id, object] of pdf.objects) if (/\/Type \/StructElem /.test(object.body)) {
    const text = [...object.body.matchAll(/\/Type \/MCR \/Pg (\d+) 0 R \/MCID (\d+)/g)].map(match => {
      const content = pdf.pages.find(page => page.id === Number(match[1])).content;
      const marked = new RegExp("/\\w+ << /MCID " + match[2] + "(?: /ActualText <[^>]*>)? >> BDC\\n([\\s\\S]*?)EMC").exec(content);
      assert.ok(marked, "Every semantic cell reference resolves to marked page content");
      return [...marked[1].matchAll(/<([a-f0-9]*)> Tj/g)].map(value => Buffer.from(value[1], "hex").toString("latin1")).join("");
    }).join("");
    elements.set(id, { id, body: object.body, role: /\/S \/(\w+)/.exec(object.body)[1], text,
      identifier: /\/ID \(([^)]*)\)/.exec(object.body)?.[1], headers: headerIds(object.body) });
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
  assert.match(headers.find(element => element.text === "Identifier").body, /\/Scope \/Column \/RowSpan 3/);
  assert.match(headers.find(element => element.text === "Business").body, /\/Scope \/Column \/ColSpan 4/);
  const expected = names => names.map(name => ids.get(name));
  assert.deepEqual(headers.find(element => element.text === "Metrics").headers, expected(["Business"]));
  assert.deepEqual(headers.find(element => element.text === "Healthy").headers, expected(["Business", "Availability"]));
  const data = [...pdf.elements.values()].filter(element => element.role === "TD");
  assert.deepEqual(data.find(element => element.text === "SVC-1").headers, expected(["Identifier"]));
  assert.deepEqual(data.find(element => element.text === "12").headers, expected(["Business", "Metrics", "Amount"]));
  const total = data.find(element => element.text === "Totals");
  assert.match(total.body, /\/ColSpan 3/);
  assert.deepEqual(total.headers, expected(["Identifier", "Business", "Metrics", "Amount", "Count"]));
  const reviewed = data.find(element => element.text === "Reviewed");
  assert.match(reviewed.body, /\/ColSpan 3/);
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
