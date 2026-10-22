import test from "node:test";
import assert from "node:assert/strict";
import { createGroupedReport } from "../dist/core/index.js";
import { writeGroupedPdf, writeGroupedPdfTo, pdfGroupedTable } from "../dist/pdf/grouped.js";
import { writePdfDocument } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const columns = [{ header: "Service", key: "name" }, { header: "Amount", key: "amount", type: "number", groups: ["Metrics"] }];
const levels = ["region", "site", "environment", "team"];
const labels = { labelColumn: { header: "Summary" }, subtotalLabel: group => `SUM:${group.label}`,
  grandTotalLabel: summary => `ALL:${summary.rows}` };
const options = { tagged: { language: "en" }, pageSize: { width: 320, height: 260 }, orientation: "portrait",
  margins: 20, pageNumbers: false, compression: false, fontSize: 9 };
const source = (rows, selectedColumns = columns) => createGroupedReport(rows, { columns: selectedColumns, groupBy: levels.map(key => ({ key, value: row => row[key] })),
  totals: { amount: "sum" } });
const references = body => [...body.matchAll(/(\d+) 0 R/g)].map(match => Number(match[1]));

/** Inspect the actual structure object graph and the marked text painted on its referenced pages.
 * These ASCII fixtures use the standard font; native reader/AT acceptance is a separate gate. */
async function semanticPdf(blob) {
  const pdf = await inspectPdf(blob), marked = new Map();
  for (const page of pdf.pages) for (const match of page.content.matchAll(/\/(\w+) << \/MCID (\d+)(?: \/ActualText <[^>]*>)? >> BDC\n([\s\S]*?)EMC/g))
    marked.set(`${page.id}:${match[2]}`, [...match[3].matchAll(/<([a-f0-9]*)> Tj/g)].map(part => Buffer.from(part[1], "hex").toString("latin1")).join(""));
  const nodes = new Map([...pdf.objects].filter(([, object]) => /\/Type \/StructElem /.test(object.body)).map(([id, object]) => {
    const kids = /\/K \[([\s\S]*?)\](?: |$)/.exec(object.body)[1];
    return [id, { id, body: object.body, role: /\/S \/(\w+)/.exec(object.body)[1], parent: Number(/\/P (\d+) 0 R/.exec(object.body)[1]),
      children: references(kids.replace(/<<[\s\S]*?>>/g, "")), text: [...kids.matchAll(/\/MCR \/Pg (\d+) 0 R \/MCID (\d+)/g)]
        .map(part => { const value = marked.get(`${part[1]}:${part[2]}`); assert.notEqual(value, undefined); return value; }).join("") }];
  }));
  for (const node of nodes.values()) for (const child of node.children) assert.equal(nodes.get(child)?.parent, node.id);
  return { ...pdf, nodes, find: (role, text) => [...nodes.values()].filter(node => node.role === role && node.text === text) };
}

test("tagged four-level reports keep group headings, details and ancestor summary tables in their own logical sections", async () => {
  const long = Array.from({ length: 34 }, (_, index) => `Detail-${index}`).join("\n");
  const rows = [
    { region: "Region-A", site: "Site-A", environment: "Env-A", team: "Team-A", name: long, amount: 2 },
    { region: "Region-A", site: "Site-A", environment: "Env-A", team: "Team-B", name: "Service-B", amount: 3 },
    { region: "Region-B", site: "Site-B", environment: "Env-B", team: "Team-C", name: "Service-C", amount: 7 }
  ];
  const expected = [
    ["Region-A", 0, undefined], ["Site-A", 1, "Region-A"], ["Env-A", 2, "Site-A"],
    ["Team-A", 3, "Env-A"], ["Team-B", 3, "Env-A"],
    ["Region-B", 0, undefined], ["Site-B", 1, "Region-B"], ["Env-B", 2, "Site-B"], ["Team-C", 3, "Env-B"]
  ];
  for (const composed of [false, true]) {
    const groupHeadingLevel = composed ? 3 : 1;
    const blob = composed ? await writePdfDocument([
      { kind: "paragraph", text: "Report", headingLevel: 1 },
      { kind: "section", title: "Analysis", level: 2, bookmark: false },
      pdfGroupedTable(source(rows), { ...labels, groupHeadingLevel }), { kind: "paragraph", text: "Recommendations" }
    ], options) : await writeGroupedPdf(source(rows), { ...options, ...labels, groupHeadingLevel });
    const pdf = await semanticPdf(blob), sections = new Map();
    assert.ok(pdf.pages.length > 2, "The long detail must exercise physical continuation pages");
    for (const [label, level] of expected) {
      const headings = pdf.find(`H${level + groupHeadingLevel}`, label); assert.equal(headings.length, 1, `One semantic heading: ${label}`);
      const section = pdf.nodes.get(headings[0].parent); assert.equal(section.role, "Sect"); sections.set(label, section);
      const summaries = pdf.find("TH", `SUM:${label}`); assert.equal(summaries.length, 1, `One row-heading subtotal: ${label}`);
      assert.match(summaries[0].body, /\/Scope \/Row/);
      const row = pdf.nodes.get(summaries[0].parent), table = pdf.nodes.get(row.parent);
      assert.equal(row.role, "TR"); assert.equal(table.role, "Table"); assert.equal(table.parent, section.id);
      assert.ok(table.children.some(id => pdf.nodes.get(id).children.some(cell => pdf.nodes.get(cell).text === "Amount")), "A summary's own visible column heading has its own semantic owner");
      const metrics = row.children.map(id => pdf.nodes.get(id)).filter(cell => cell.role === "TD");
      assert.ok(metrics.every(cell => /\/Headers \[/.test(cell.body)), "Summary values retain explicit header associations");
    }
    for (const [label, , parent] of expected) if (parent) assert.equal(sections.get(label).parent, sections.get(parent).id, label);
    for (const [text, group] of [[long.replaceAll("\n", ""), "Team-A"], ["Service-B", "Team-B"], ["Service-C", "Team-C"]]) {
      const [cell] = pdf.find("TD", text); assert.ok(cell, text);
      const table = pdf.nodes.get(pdf.nodes.get(cell.parent).parent); assert.equal(table.parent, sections.get(group).id);
    }
    const [grand] = pdf.find("TH", "ALL:3"); assert.ok(grand);
    assert.equal(pdf.nodes.get(pdf.nodes.get(grand.parent).parent).parent, sections.get("Region-A").parent,
      "The grand total belongs outside every analytical group");
    assert.ok(pdf.text.split("Region-A").length > 2, "Repeated visual group context remains on continued pages");
    assert.ok([...pdf.objects.values()].some(object => /\/Type \/Outlines/.test(object.body)), "Group headings provide native PDF outline navigation");
    if (composed) {
      const [after] = pdf.find("P", "Recommendations"); assert.equal(after.parent, sections.get("Region-A").parent,
        "Analytical group ownership does not leak into later document content");
    }
  }
});

test("empty tagged analytical reports preserve optional headings and grand summaries without manufactured groups", async () => {
  for (const includeHeader of [false, true]) for (const includeGrandTotal of [false, true]) {
    const chunks = [];
    const selectedColumns = includeHeader ? columns : columns.map(({ groups: _groups, ...column }) => column);
    const result = await writeGroupedPdfTo(source([], selectedColumns), { write(bytes) { chunks.push(bytes.slice()); } },
      { ...options, ...labels, includeHeader, includeGrandTotal, groupHeadingLevel: 6 });
    const pdf = await semanticPdf(new Blob(chunks));
    assert.equal(result.rows, 0); assert.equal(result.columns, 3);
    assert.equal([...pdf.nodes.values()].filter(node => /^H[2-5]$/.test(node.role)).length, 0);
    assert.equal([...pdf.nodes.values()].filter(node => node.role === "Sect").length, 0);
    assert.equal(pdf.find("TH", "ALL:0").length, includeGrandTotal ? 1 : 0);
    assert.equal(pdf.text.includes("Amount"), includeHeader);
    assert.equal(pdf.text.includes("Metrics"), includeHeader);
  }
});

test("tagged grouping captures projected values once and returns the source when structure limits reject", async () => {
  let reads = 0, returned = 0;
  function* rows() {
    try {
      for (let index = 0; index < 2; index++) yield {
        region: "Region", site: "Site", environment: "Environment", team: "Team",
        get name() { reads++; return `Service-${index}`; }, get amount() { reads++; return index + 1; }
      };
    } finally { returned++; }
  }
  const formatted = [], chunks = [];
  const result = await writeGroupedPdfTo(source(rows()), { write(bytes) { chunks.push(bytes.slice()); } }, {
    ...options, ...labels, formatValue(value, position) {
      if (position.column.key === "amount") formatted.push([position.rowIndex, position.columnIndex, value]);
      return String(value ?? "");
    }
  });
  assert.equal(reads, 4); assert.equal(returned, 1); assert.equal(result.rows, 2);
  assert.deepEqual(formatted.slice(0, 2), [[0, 1, 1], [1, 1, 2]]);
  assert.deepEqual(formatted.slice(2), Array.from({ length: 5 }, () => [-1, 1, 3]), "Four ancestor summaries and the grand total use captured numeric values");
  assert.equal((await semanticPdf(new Blob(chunks))).find("TH", "ALL:2").length, 1);
  reads = 0; returned = 0;
  const destination = new WritableStream({ write() {} });
  await assert.rejects(writeGroupedPdfTo(source(rows()), destination, { ...options, ...labels, limits: { maxStructureElements: 12 } }), /maxStructureElements/);
  assert.equal(reads, 4); assert.equal(returned, 1); assert.equal(destination.locked, false);
  reads = 0; returned = 0;
  const overflowDestination = new WritableStream({ write() {} });
  await assert.rejects(writeGroupedPdfTo(source(rows()), overflowDestination, { ...options, ...labels, groupHeadingLevel: 4 }),
    /groupHeadingLevel.*opened group depth/);
  assert.equal(returned, 1); assert.equal(overflowDestination.locked, false);
  for (const groupHeadingLevel of [0, 7, 1.5, "2"])
    await assert.rejects(writeGroupedPdf(source([]), { ...options, ...labels, groupHeadingLevel }), /groupHeadingLevel/);
});
