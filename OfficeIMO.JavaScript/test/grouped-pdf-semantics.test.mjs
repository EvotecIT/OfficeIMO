import test from "node:test";
import assert from "node:assert/strict";
import { createGroupedReport } from "../dist/core/index.js";
import { writeGroupedPdf, writeGroupedPdfTo, pdfGroupedTable } from "../dist/pdf/grouped.js";
import { writePdfDocument } from "../dist/pdf/index.js";
import { inspectPdf, pdfStructureAttributes } from "./pdf-reader.mjs";
import { pdfStructureChildren } from "./pdf-structure-reader.mjs";

const columns = [{ header: "Service", key: "name" }, { header: "Amount", key: "amount", type: "number", groups: ["Metrics"] }];
const levels = ["region", "site", "environment", "team"];
const labels = { labelColumn: { header: "Summary" }, subtotalLabel: group => `SUM:${group.label}`,
  grandTotalLabel: summary => `ALL:${summary.rows}` };
const options = { tagged: { language: "en" }, pageSize: { width: 320, height: 260 }, orientation: "portrait",
  margins: 20, pageNumbers: false, compression: false, fontSize: 9 };
const source = (rows, selectedColumns = columns) => createGroupedReport(rows, { columns: selectedColumns, groupBy: levels.map(key => ({ key, value: row => row[key] })),
  totals: { amount: "sum" } });

test("tagged group context charges only painted labels and the owned bookmark", async () => {
  const report = () => createGroupedReport([{ site: "A", name: "ROW" }], {
    columns: [{ header: "H", key: "name" }], groupBy: [{ key: "site", value: row => row.site }], totals: {}
  });
  const configuration = { tagged: { language: "en" }, labelColumn: { header: "S" },
    subtotalLabel: () => "SUM", grandTotalLabel: () => "ALL", includeGrandTotal: false,
    pageNumbers: false, compression: false };
  // Heading A, two column headings, two detail cells and two subtotal cells; one bookmark A.
  const blob = await writeGroupedPdf(report(), { ...configuration, limits: { maxCells: 8, maxTextCharacters: 10 } });
  const pdf = await inspectPdf(blob);
  assert.equal(pdf.pages.length, 1);
  assert.deepEqual(pdf.pages[0].lines, ["A", "S", "H", "", "ROW", "SUM", ""]);
  await assert.rejects(writeGroupedPdf(report(), { ...configuration, limits: { maxCells: 7 } }), /maxCells/);
  await assert.rejects(writeGroupedPdf(report(), { ...configuration, limits: { maxTextCharacters: 9 } }), /maxTextCharacters/);
});

/** Inspect the actual structure object graph and the marked text painted on its referenced pages.
 * These ASCII fixtures use the standard font; native reader/AT acceptance is a separate gate. */
async function semanticPdf(blob) {
  const pdf = await inspectPdf(blob), marked = new Map();
  const pages = new Set(pdf.pages.map(page => page.id));
  for (const page of pdf.pages) for (const match of page.content.matchAll(/\/(\w+) << \/MCID (\d+)(?: \/ActualText <[^>]*>)? >> BDC\n([\s\S]*?)EMC/g))
    marked.set(`${page.id}:${match[2]}`, [...match[3].matchAll(/<([a-f0-9]*)> Tj/g)].map(part => Buffer.from(part[1], "hex").toString("latin1")).join(""));
  const nodes = new Map([...pdf.objects].filter(([, object]) => /\/Type \/StructElem /.test(object.body)).map(([id, object]) => {
    const { entries } = pdfStructureChildren(object.body);
    return [id, { id, body: object.body, attributes: pdfStructureAttributes(pdf, object.body), role: /\/S \/(\w+)/.exec(object.body)[1], parent: Number(/\/P (\d+) 0 R/.exec(object.body)[1]),
      children: entries.filter(entry => entry.kind === "element").map(entry => entry.id),
      text: entries.filter(entry => entry.kind === "mark").map(entry => {
        assert.ok(pages.has(entry.pageId), "Every structure mark resolves to a physical page");
        const value = marked.get(`${entry.pageId}:${entry.mcid}`);
        assert.notEqual(value, undefined, "Every structure mark resolves to marked page content"); return value;
      }).join("") }];
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
      assert.match(summaries[0].attributes, /\/Scope \/Row/);
      const row = pdf.nodes.get(summaries[0].parent), table = pdf.nodes.get(row.parent);
      assert.equal(row.role, "TR"); assert.equal(table.role, "Table"); assert.equal(table.parent, section.id);
      assert.ok(table.children.some(id => pdf.nodes.get(id).children.some(cell => pdf.nodes.get(cell).text === "Amount")), "A summary's own visible column heading has its own semantic owner");
      const metrics = row.children.map(id => pdf.nodes.get(id)).filter(cell => cell.role === "TD");
      assert.ok(metrics.every(cell => /\/Headers \[/.test(cell.attributes)), "Summary values retain explicit header associations");
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

test("tagged leaf subtotals share their own detail table while ancestor totals keep independent ownership", async () => {
  const rows = [
    { region: "Root", site: "Site", environment: "Env", team: "LeafA", name: "REC-A", amount: 2 },
    { region: "Root", site: "Site", environment: "Env", team: "LeafB", name: "REC-B", amount: 3 }
  ];
  for (const bands of [false, true]) {
    const pdf = await semanticPdf(await writeGroupedPdf(source(rows), { ...options, ...labels,
      pageSize: { width: 300, height: 380 }, columnWidths: [90, 90, 120],
      ...(bands ? { wideTable: { mode: "bands", repeatColumns: [0], minimumColumnWidth: 80 } } : {}) }));
    const tableFor = cell => pdf.nodes.get(pdf.nodes.get(cell.parent).parent);
    for (const [name, leaf] of [["REC-A", "LeafA"], ["REC-B", "LeafB"]]) {
      const [detail] = pdf.find("TD", name), [summary] = pdf.find("TH", "SUM:" + leaf);
      assert.ok(detail); assert.ok(summary);
      assert.equal(tableFor(detail).id, tableFor(summary).id, "The leaf total reuses its own logical detail table");
      const [heading] = pdf.find("H5", leaf);
      assert.equal(tableFor(summary).parent, heading.parent);
      assert.equal(pdf.find("TD", name).length, 1, "Repeated physical bands do not repeat logical source cells");
    }
    const tables = [...pdf.nodes.values()].filter(node => node.role === "Table");
    assert.equal(tables.length, 6, "Two leaf tables, three ancestor summary tables and one grand-total table");
    for (const ancestor of ["Root", "Site", "Env"]) {
      const [summary] = pdf.find("TH", "SUM:" + ancestor), table = tableFor(summary);
      const heading = [...pdf.nodes.values()].find(node => /^H[2-4]$/.test(node.role) && node.text === ancestor);
      assert.equal(table.parent, heading.parent);
      assert.equal(table.children.map(id => pdf.nodes.get(id)).filter(node => node.role === "TR").length, 3,
        "The ancestor has its own two-row header and one summary, without the final child's data");
    }
    assert.equal(pdf.find("TH", "ALL:2").length, 1);
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

test("group headings and native outlines omit blank labels while preserving unnamed sibling sections", async () => {
  for (const authoredLabels of [false, true]) for (const composed of [false, true]) {
    const namedLabel = authoredLabels ? "  Named  " : "Named", capturedLabels = [];
    const rows = [
      { outer: "A", inner: "X", name: "Service-X", amount: 1 },
      { outer: null, inner: "Y", name: "Service-Y", amount: 2 },
      { outer: authoredLabels ? "hidden-blank" : "", inner: "Blank-child", name: "Service-blank", amount: 3 },
      { outer: authoredLabels ? "hidden-space" : "   ", inner: "Space-child", name: "Service-space", amount: 4 },
      { outer: "Named", inner: "Named-first", name: "Service-first", amount: 5 },
      { outer: "Named", inner: "Named-second", name: "Service-second", amount: 6 }
    ];
    const grouped = createGroupedReport(rows, { columns, totals: { amount: "sum" }, groupBy: [
      { key: "outer", value: row => row.outer, ...(authoredLabels ? { label: value => {
        capturedLabels.push(value);
        return value === "hidden-blank" ? "" : value === "hidden-space" ? "   " : value === "Named" ? namedLabel : String(value ?? "");
      } } : {}) },
      { key: "inner", value: row => row.inner }
    ] });
    const pdf = await semanticPdf(composed ? await writePdfDocument([
      { kind: "section", title: "Scope", level: 1 }, pdfGroupedTable(grouped, { ...labels, groupHeadingLevel: 2 }),
      { kind: "section", title: "After", level: 1 }, { kind: "paragraph", text: "Recommendations" }
    ], options) : await writeGroupedPdf(grouped, { ...options, ...labels, groupHeadingLevel: 2 }));
    const groupHeadings = [...pdf.nodes.values()].filter(node => /^H[23]$/.test(node.role));
    assert.ok(groupHeadings.every(heading => heading.text.trim()), "Unnamed groups must not manufacture empty or whitespace-only semantic headings");
    assert.deepEqual(groupHeadings.map(heading => heading.text), ["A", "X", "Y", "Blank-child", "Space-child", namedLabel, "Named-first", "Named-second"],
      "Every nonblank authored heading is preserved exactly once without trimming its displayed text");
    for (const title of ["Y", "Blank-child", "Space-child"]) {
      const [heading] = pdf.find("H3", title), child = pdf.nodes.get(heading.parent), ancestor = pdf.nodes.get(child.parent);
      assert.equal(child.role, "Sect"); assert.equal(ancestor.role, "Sect", "Blank groups retain their real section depth");
      assert.ok(ancestor.children.every(id => !/^H[1-6]$/.test(pdf.nodes.get(id).role)), "An unnamed parent section has no manufactured heading");
    }
    assert.equal(capturedLabels.length, authoredLabels ? 5 : 0, "Semantic heading preparation does not replay group label callbacks");
    const root = [...pdf.objects].find(([, object]) => /\/Type \/Outlines /.test(object.body));
    assert.ok(root);
    const outlines = new Map([...pdf.objects].filter(([, object]) => /\/Title <feff[0-9a-f]*> \/Parent \d+ 0 R.*\/Dest \[/i.test(object.body)).map(([id, object]) => {
      const title = Buffer.from(/\/Title <feff([0-9a-f]*)>/i.exec(object.body)[1], "hex").swap16().toString("utf16le");
      const destination = /\/Dest \[(\d+) 0 R \/XYZ null ([\d.-]+) null\]/.exec(object.body);
      return [title, { id, parent: Number(/\/Parent (\d+) 0 R/.exec(object.body)[1]), page: Number(destination[1]), y: Number(destination[2]) }];
    }));
    assert.deepEqual([...outlines.keys()], [...(composed ? ["Scope"] : []), "A", "X", "Y", "Blank-child", "Space-child", namedLabel,
      "Named-first", "Named-second", ...(composed ? ["After"] : [])]);
    const parent = title => outlines.get(title).parent;
    const enclosing = composed ? outlines.get("Scope").id : root[0];
    if (composed) {
      assert.equal(parent("Scope"), root[0]); assert.equal(parent("After"), root[0], "Later composed sections retain their own root ancestry");
    }
    assert.equal(parent("A"), enclosing); assert.equal(parent("X"), outlines.get("A").id);
    for (const title of ["Y", "Blank-child", "Space-child", namedLabel])
      assert.equal(parent(title), enclosing, `${title} must not become a child of a closed named sibling`);
    for (const title of ["Named-first", "Named-second"])
      assert.equal(parent(title), outlines.get(namedLabel).id, "An actual open named ancestor retains all of its children");
    for (const [title, destination] of outlines) {
      const page = pdf.pages.find(page => page.id === destination.page); assert.ok(page, `${title} targets a real page`);
      assert.ok(page.lines.includes(title), `${title} targets the page where its real heading is painted`);
      const height = Number(/\/MediaBox \[0 0 [\d.]+ ([\d.]+)\]/.exec(page.body)[1]);
      assert.ok(Number.isFinite(destination.y) && destination.y >= 0 && destination.y <= height, `${title} has a bounded native destination`);
    }
  }
});

test("tagged grouped detail and subtotal tables repeat complete physical context without duplicate logical headings", async () => {
  const selected = [{ header: "ID", key: "id" }, { header: "Detail", key: "detail" }, { header: "Value", key: "amount", type: "number" }];
  const long = Array.from({ length: 24 }, (_, index) => `Detail-${index}`).join("\n");
  const rows = [
    { region: "Root", site: "Site", environment: "Env", team: "LeafA", id: "REC-A0", detail: long, amount: 1 },
    { region: "Root", site: "Site", environment: "Env", team: "LeafA", id: "REC-A1", detail: "Next", amount: 2 },
    { region: "Root", site: "Site", environment: "Env", team: "LeafB", id: "REC-B0", detail: "Last", amount: 3 }
  ];
  const paths = new Map(rows.map(row => [row.id, levels.map(level => row[level])]));
  const summaries = new Map([...["LeafA", "LeafB"].map(leaf => [`Tot3${leaf}`, ["Root", "Site", "Env", leaf]]),
    ["Tot2Env", ["Root", "Site", "Env"]], ["Tot1Site", ["Root", "Site"]], ["Tot0Root", ["Root"]]]);
  for (const bands of [false, true]) for (const composed of [false, true]) {
    const calls = [], table = { labelColumn: { header: "Summary", width: 80 }, subtotalLabel: group => `Tot${group.level}${group.label}`,
      grandTotalLabel: () => "GRAND", continuationLabel(group) { calls.push(group.groupIndex); return `CONT:${group.label}`; },
      columnWidths: [80, 80, 100, 100], ...(bands ? { wideTable: { mode: "bands", minimumColumnWidth: 80, repeatColumns: [0, 1] } } : {}) };
    const grouped = source(rows, selected), configuration = { ...options, pageSize: { width: 300, height: 340 }, limits: { maxTableWindowRows: 1 } };
    const blob = composed ? await writePdfDocument([
      { kind: "paragraph", text: "Report", headingLevel: 1 },
      { kind: "paragraph", text: Array.from({ length: 17 }, (_, index) => `Preface-${index}`).join("\n") },
      pdfGroupedTable(grouped, table)
    ], configuration) : await writeGroupedPdf(grouped, { ...configuration, ...table });
    const pdf = await semanticPdf(blob);
    const assertContext = (page, index, path) => {
      const header = page.lines.lastIndexOf("Summary", index);
      assert.ok(header >= 0, "Every physical data/summary band has its column headings");
      const context = page.lines.slice(0, header).slice(-path.length);
      assert.ok(JSON.stringify(context) === JSON.stringify(path) || JSON.stringify(context) === JSON.stringify(path.map(label => `CONT:${label}`)),
        `Full ordered context must immediately precede this table's headings: ${JSON.stringify({ bands, composed, context, path, text: page.lines[index] })}`);
    };
    let records = 0, totals = 0, continued = false;
    for (const page of pdf.pages) for (const [index, line] of page.lines.entries()) {
      if (paths.has(line)) { assertContext(page, index, paths.get(line)); records++; }
      if (summaries.has(line)) { assertContext(page, index, summaries.get(line)); totals++; }
      continued ||= line.startsWith("CONT:");
    }
    assert.ok(records >= rows.length && totals >= summaries.size);
    assert.ok(continued, "A continued row/window must keep the authored continuation context");
    for (const [level, names] of [[2, ["Root"]], [3, ["Site"]], [4, ["Env"]], [5, ["LeafA", "LeafB"]]])
      for (const name of names) assert.equal(pdf.find(`H${level}`, name).length, 1, "Physical repetition does not duplicate logical group headings");
    assert.equal([...pdf.nodes.values()].filter(node => node.role === "Caption").length, 0, "Plain group context has no duplicate logical caption");
    assert.equal(calls.length, 5); assert.equal(new Set(calls).size, 5, "Each opened group's continuation callback is captured once across its leaf and summary tables");
    const [grand] = pdf.find("TH", "GRAND"); assert.ok(grand);
    const grandTable = pdf.nodes.get(pdf.nodes.get(grand.parent).parent);
    assert.equal(pdf.nodes.get(grandTable.parent).role, "Document", "Grand totals retain no closed group ownership");
    if (composed) {
      const first = pdf.pages.find(page => page.lines.includes("Root"));
      assert.ok(pdf.pages.indexOf(first) > 0, "The complete group-heading/context/first-fragment chain moves off the nearly full preface page");
      assert.ok(first.lines.includes("REC-A0"), "A real group heading is not orphaned from its kept first fragment");
    }
  }
});

test("tagged leaf subtotal shares its detail headers and retains context through a long-label continuation", async () => {
  const selected = [{ header: "Item", key: "name" }, { header: "Amount", key: "amount", type: "number" }];
  const summaryLines = Array.from({ length: 22 }, (_, index) => `Subtotal-${index}`);
  const preface = Array.from({ length: 6 }, (_, index) => `Preface-${index}`).join("\n");
  const table = { labelColumn: { header: "Summary", width: 100 }, columnWidths: [100, 90, 90],
    subtotalLabel: () => summaryLines.join("\n"), grandTotalLabel: () => "Grand" };
  const configuration = { ...options, orientation: "landscape" };
  for (const groupLabel of ["Group", ""]) for (const composed of [false, true]) {
    const grouped = createGroupedReport([{ group: groupLabel || null, name: "Only", amount: 2 }], {
      columns: selected, groupBy: [{ key: "group", value: row => row.group }], totals: { amount: "sum" }
    });
    const pdf = await semanticPdf(composed
      ? await writePdfDocument([{ kind: "paragraph", text: preface }, pdfGroupedTable(grouped, table)], configuration)
      : await writeGroupedPdf(grouped, { ...configuration, ...table, messageTop: preface }));
    const first = pdf.pages.find(page => page.lines.includes("Only"));
    assert.ok(first);
    assert.equal(first.lines.filter(line => line === "Summary").length, 1,
      "The leaf subtotal shares its detail headings rather than creating another table");
    assert.ok(first.lines.includes("Subtotal-0"),
      "An overheight leaf subtotal starts below its detail while retaining its existing context/headings");
    const summaryPages = pdf.pages.filter(page => page.lines.some(line => summaryLines.includes(line)));
    assert.ok(summaryPages.length > 1, "The wrapped subtotal must exercise continuation");
    for (const page of summaryPages) {
      for (const header of ["Summary", "Item", "Amount"]) assert.ok(page.lines.includes(header), header);
      if (groupLabel) assert.ok(page.lines.includes(groupLabel), "Each summary fragment retains its active physical context");
    }
    assert.deepEqual(pdf.pages.flatMap(page => page.lines).filter(line => summaryLines.includes(line)), summaryLines);
    assert.equal(pdf.find("TH", summaryLines.join("")).length, 1, "Every physical summary fragment retains one logical row heading");
    assert.equal(pdf.find("H2", "Group").length, groupLabel ? 1 : 0);
    assert.equal([...pdf.nodes.values()].filter(node => node.role === "Caption").length, 0);
    const [grand] = pdf.find("TH", "Grand"); assert.ok(grand);
    assert.equal(pdf.nodes.get(pdf.nodes.get(pdf.nodes.get(grand.parent).parent).parent).role, "Document",
      "Grand totals remain outside the analytical section");
  }
});

test("new group headings paint the opening path once while subsequent bands retain complete context", async () => {
  const selected = [{ header: "ID", key: "id", rowHeader: true }, { header: "Detail", key: "detail" }, { header: "Amount", key: "amount", type: "number" }];
  const grouped = createGroupedReport([
    { outer: "Region", inner: "First", id: "REC-A", detail: "Alpha", amount: 1 },
    { outer: "Region", inner: "Second", id: "REC-B", detail: "Beta", amount: 2 }
  ], { columns: selected, groupBy: ["outer", "inner"].map(key => ({ key, value: row => row[key] })), totals: { amount: "sum" } });
  const pdf = await semanticPdf(await writeGroupedPdf(grouped, { ...options, ...labels,
    pageSize: { width: 360, height: 500 }, columnWidths: [80, 80, 100, 100],
    wideTable: { mode: "bands", repeatColumns: [0, "id"] } }));
  for (const [record, leaf] of [["REC-A", "First"], ["REC-B", "Second"]]) {
    const physical = pdf.pages.filter(page => page.lines.includes(record));
    assert.equal(physical.length, 2, "Both selected data bands retain the complete identifier");
    for (const page of physical) {
      const header = page.lines.lastIndexOf("Summary", page.lines.indexOf(record));
      assert.deepEqual(page.lines.slice(0, header).slice(-2), ["Region", leaf], "Each band's opening path is complete and ordered");
    }
    const first = physical[0], header = first.lines.lastIndexOf("Summary", first.lines.indexOf(record));
    const opening = first.lines.slice(0, header);
    assert.notDeepEqual(opening.slice(-4), ["Region", leaf, "Region", leaf],
      "The current opening path is not painted again by a full caption; earlier finished groups may share this page");
    assert.equal(pdf.find("H3", leaf).length, 1, "The logical heading remains unique");
  }
});

test("separate analytical summaries retain only declared aggregate fields with original formatter positions", async () => {
  const selected = [
    { key: "id", header: "Identifier", rowHeader: true }, { key: "detail", header: "Description" },
    { key: "zero", header: "Zero sum", type: "number" }, { key: "unused", header: "Other detail" },
    { key: "empty", header: "Empty average", type: "number" }
  ];
  let enumerations = 0, reads = 0;
  function* rows() { enumerations++; yield { outer: "Region", inner: "Team", get id() { reads++; return "REC-A"; }, detail: "Detail-A", zero: 0, unused: "Other-A", empty: null }; }
  const calls = [];
  const grouped = createGroupedReport(rows(), { columns: selected, groupBy: ["outer", "inner"].map(key => ({ key, value: row => row[key] })), totals: { zero: "sum", empty: "average" } });
  const pdf = await semanticPdf(await writeGroupedPdf(grouped, { ...options, ...labels,
    pageSize: { width: 360, height: 500 }, columnWidths: [80, 80, 100, 100, 100, 100],
    wideTable: { mode: "bands", repeatColumns: [0, "id"] },
    formatValue(value, position) { calls.push([position.rowIndex, position.columnIndex, position.column.key, value]); return value === null ? "empty" : String(value); }
  }));
  assert.equal(enumerations, 1); assert.equal(reads, 1);
  for (const label of ["SUM:Region", "ALL:1"]) {
    const [summary] = pdf.find("TH", label), table = pdf.nodes.get(pdf.nodes.get(summary.parent).parent);
    const cells = table.children.flatMap(id => pdf.nodes.get(id).children.map(child => pdf.nodes.get(child)));
    assert.deepEqual(cells.filter(cell => cell.role === "TH" && /\/Scope \/Column/.test(cell.attributes)).map(cell => cell.text),
      ["Summary", "Zero sum", "Empty average"], "Standalone headings describe only the actual aggregate values");
    assert.deepEqual(cells.filter(cell => cell.role === "TD").map(cell => cell.text), ["0", "empty"], "Declared zero and null results are retained");
    const pages = pdf.pages.filter(page => page.lines.includes(label));
    assert.equal(pages.length, 1, "Non-aggregate fields do not manufacture empty summary bands");
    assert.ok(pages[0].lines.includes("Zero sum") && pages[0].lines.includes("Empty average"));
  }
  for (const key of ["zero", "empty"]) {
    const position = selected.findIndex(column => column.key === key);
    assert.ok(calls.filter(call => call[0] === -1 && call[2] === key).every(call => call[1] === position), "Projected callbacks keep original source-column positions");
  }
});

test("untagged projected grand summaries paint their own aggregate headings on a shared page", async () => {
  const grouped = createGroupedReport([{ group: "Region", id: "REC-A", amount: 2 }], {
    columns: [{ key: "id", header: "Identifier" }, { key: "amount", header: "Amount", type: "number" }],
    groupBy: [{ key: "group", value: row => row.group }], totals: { amount: "sum" }
  });
  const { tagged: _tagged, ...untagged } = options;
  const pdf = await inspectPdf(await writeGroupedPdf(grouped, { ...untagged, ...labels, pageSize: { width: 500, height: 700 } }));
  const page = pdf.pages.find(page => page.lines.includes("ALL:1")), total = page.lines.indexOf("ALL:1");
  assert.ok(page.lines.includes("REC-A"), "The separate summary can share the detail page");
  assert.deepEqual(page.lines.slice(total - 2, total), ["Summary", "Amount"], "Projected values cannot reuse the preceding three-column detail heading positions");
});
