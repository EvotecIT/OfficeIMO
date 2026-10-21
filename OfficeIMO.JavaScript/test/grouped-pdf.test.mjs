import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { createGroupedReport } from "../dist/core/index.js";
import { writeGroupedPdf, writeGroupedPdfTo, pdfGroupedTable } from "../dist/pdf/grouped.js";
import { writePdf, writePdfDocument, writePdfDocumentTo, pdfTable, PdfFont } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const regular = new PdfFont(new Uint8Array(await readFile(new URL("../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf", import.meta.url))));
const columns = [{ header: "Zespół", key: "team" }, { header: "Usługa", key: "name" }, { header: "Czas", key: "amount", type: "number" }];
const options = { labelColumn: { header: "Podsumowanie", width: 22 }, subtotalLabel: group => `${group.label} (${group.rows})`,
  grandTotalLabel: summary => `Razem (${summary.rows})`, fonts: { regular }, pageSize: "A5", compression: false, pageNumbers: false };
const report = rows => createGroupedReport(rows, { columns, groupBy: "team", totals: { amount: "average" } });

test("grouped PDF grand summaries preserve grouped column headings and original formatting context", async () => {
  const groupedColumns = columns.map(column => column.key === "amount" ? { ...column, groups: ["Miary"] } : column);
  const rows = [{ team: "A", name: "Pierwsza", amount: 2 }, { team: "B", name: "Druga", amount: 6 }];
  const source = values => createGroupedReport(values, { columns: groupedColumns, groupBy: "team", totals: { amount: "average" } });
  for (const includeGrandTotal of [true, false]) {
    const pdf = await inspectPdf(await writeGroupedPdf(source(rows), { ...options, includeGrandTotal, formatValue(value, context) {
      if (context.column.key === "amount") {
        assert.deepEqual(context.column.groups, ["Miary"], "The original source context retains heading metadata");
        return value == null ? "" : value.toFixed(1);
      }
      return String(value ?? "");
    } }));
    const lines = pdf.pages.flatMap(page => page.lines);
    assert.equal(lines.filter(line => line === "Miary").length, 2, "Only the two detail tables repeat their declared headings");
    assert.equal(lines.includes("Razem (2)"), includeGrandTotal);
    if (includeGrandTotal) assert.ok(lines.includes("4.0"), "The final average is computed from source values");
  }
  const empty = await inspectPdf(await writeGroupedPdf(source([]), options));
  assert.ok(empty.pages.flatMap(page => page.lines).includes("Miary"), "An empty report retains its grouped schema headings");
});

test("grouped PDF streams matching shared totals, source-key formatting and repeated group context over long rows", async () => {
  let produced = 0, returned = 0, reads = 0;
  const long = Array.from({ length: 90 }, (_, i) => `Wiersz ${i}`).join("\n"), seen = [], progress = [], chunks = [];
  function* source() { try {
    for (const [team, name, amount] of [["Łódź", long, 2], ["Łódź", "Druga", 4], ["Łódź", "Brak", null], ["Gdynia", "Ostatnia", 12]]) {
      produced++;
      yield { get team() { reads++; return team; }, get name() { reads++; return name; }, get amount() { reads++; return amount; } };
    }
  } finally { returned++; } }
  const result = await writeGroupedPdfTo(report(source()), { async write(bytes) {
    const before = produced; await Promise.resolve(); assert.equal(produced, before, "Rows stop while output is accepted"); chunks.push(bytes.slice());
  } }, { ...options, onProgress: p => progress.push(p), formatValue(value, context) {
    if (context.column.key === "amount") {
      assert.equal(context.columnIndex, 2, "The added label column does not shift source formatting coordinates");
      seen.push([context.rowIndex, value]); return value == null ? "" : value.toFixed(2);
    }
    return value == null ? "" : String(value);
  } });
  assert.equal(reads, 12); assert.equal(returned, 1);
  assert.equal(result.rows, 4); assert.equal(result.columns, 4);
  assert.equal(result.bytes, chunks.reduce((n, chunk) => n + chunk.length, 0));
  assert.equal(progress.at(-1).phase, "complete"); assert.equal(progress.at(-1).rows, 4);
  const pdf = await inspectPdf(new Blob(chunks));
  for (const label of ["Łódź (3)", "Gdynia (1)", "Razem (4)", "3.00", "12.00", "6.00"]) assert.ok(pdf.text.includes(label), label);
  assert.ok(pdf.pages.length >= 3);
  const longPages = pdf.pages.filter(page => page.lines.some(line => /^Wiersz /.test(line)));
  assert.ok(longPages.length >= 2);
  for (const page of longPages) {
    assert.ok(page.lines.includes("Łódź"), "Each long-row continuation retains its group context");
    for (const header of ["Podsumowanie", "Zespół", "Usługa", "Czas"]) assert.ok(page.lines.includes(header));
    assert.ok(page.lines.indexOf("Łódź") < page.lines.indexOf("Podsumowanie"));
    assert.ok(page.lines.indexOf("Czas") < page.lines.findIndex(line => /^Wiersz /.test(line)), "Group context precedes repeated leaf headings and continued body");
  }
  assert.deepEqual(pdf.pages.flatMap(page => page.lines).filter(line => /^Wiersz /.test(line)), long.split("\n"));
  assert.deepEqual(seen.filter(([index]) => index >= 0), [[0, 2], [1, 4], [2, null], [3, 12]]);
  assert.deepEqual(seen.filter(([index]) => index === -1).map(([, value]) => value), [3, 12, 6]);
});

test("grouped PDF budgets generated rows but completion reports only body rows, including an empty source", async () => {
  const one = [{ team: "A", name: "One", amount: 2 }];
  const pdf = await inspectPdf(await writeGroupedPdf(report(one), { ...options, limits: { maxRows: 3 } }));
  assert.ok(pdf.text.includes("A (1)") && pdf.text.includes("Razem (1)"));
  await assert.rejects(writeGroupedPdf(report(one), { ...options, limits: { maxRows: 2 } }), /maxRows/);
  const chunks = [], progress = [];
  const empty = await writeGroupedPdfTo(report([]), { write: bytes => { chunks.push(bytes); } }, { ...options, onProgress: p => progress.push(p) });
  assert.equal(empty.rows, 0); assert.equal(progress.at(-1).rows, 0);
  assert.ok((await inspectPdf(new Blob(chunks))).text.includes("Razem (0)"));
  await assert.rejects(writeGroupedPdf(report(one), { ...options, tagged: { language: "pl" } }), /Tagged grouped/);
  await assert.rejects(writeGroupedPdf(report(one), { ...options, subtotalLabel: async () => "Async" }), /synchronous/);
});

test("grouped summary rich-value formatting retains original source coordinates for every numeric aggregate", async () => {
  const contexts = [];
  const pdf = await inspectPdf(await writeGroupedPdf(report([{ team: "A", name: "One", amount: 2 }]), {
    ...options, formatRuns(value, context) {
      if (context.column.key === "amount") {
        contexts.push([context.rowIndex, context.columnIndex, value]); return [{ text: `Value-${value}`, bold: true }];
      }
      return [{ text: String(value ?? "") }];
    }
  }));
  assert.deepEqual(contexts, [[0, 2, 2], [-1, 2, 2], [-1, 2, 2]]);
  assert.equal(pdf.pages.flatMap(page => page.lines).filter(line => line === "Value-2").length, 3);
});

test("empty grouped PDFs retain selected headings independently of optional grand totals", async () => {
  const headings = [options.labelColumn.header, ...columns.map(column => column.header)];
  for (const includeHeader of [true, false]) for (const includeGrandTotal of [false, true]) {
    const chunks = [], progress = [];
    const result = await writeGroupedPdfTo(report([]), { write: bytes => chunks.push(bytes.slice()) },
      { ...options, includeHeader, includeGrandTotal, onProgress: event => progress.push(event) });
    assert.equal(result.rows, 0);
    assert.equal(result.columns, columns.length + 1);
    assert.equal(result.bytes, chunks.reduce((sum, chunk) => sum + chunk.length, 0));
    assert.equal(progress.at(-1).phase, "complete"); assert.equal(progress.at(-1).rows, 0);
    const pdf = await inspectPdf(new Blob(chunks));
    for (const heading of headings) assert.equal(pdf.text.includes(heading), includeHeader, heading);
    assert.equal(pdf.text.includes("Razem (0)"), includeGrandTotal);
    const blob = await writeGroupedPdf(report([]), { ...options, includeHeader, includeGrandTotal });
    assert.equal(blob.type, "application/pdf");
    const copied = await inspectPdf(blob);
    for (const heading of headings) assert.equal(copied.text.includes(heading), includeHeader, heading);
    assert.equal(copied.text.includes("Razem (0)"), includeGrandTotal);
  }
  const defaults = await inspectPdf(await writeGroupedPdf(report([]), { ...options, includeGrandTotal: false }));
  for (const heading of headings) assert.ok(defaults.text.includes(heading), "Headers are included by default: " + heading);
  for (const limits of [{ maxCells: 0 }, { maxTextCharacters: 0 }])
    await assert.rejects(writeGroupedPdf(report([]), { ...options, includeGrandTotal: false, limits }), { code: "RESOURCE_LIMIT" });
  const populated = await inspectPdf(await writeGroupedPdf(report([{ team: "A", name: "One", amount: 1 }]), options));
  for (const heading of headings) assert.equal(populated.text.split(heading).length - 1, 1, "A nonempty grand total does not repeat headings");
});

test("grand totals retain metric headings when they move or split onto fresh pages", async () => {
  const columns = [{ header: "Item", key: "item" }, { header: "Amount", key: "amount", type: "number" }];
  const labels = { labelColumn: { header: "Summary" }, subtotalLabel: () => "Subtotal", grandTotalLabel: () => "Grand" };
  const compact = { pageSize: { width: 320, height: 220 }, orientation: "landscape", margins: 20,
    pageNumbers: false, compression: false, fontSize: 10 };
  const source = (count, depth = 1) => createGroupedReport(Array.from({ length: count }, (_, index) => ({ item: `Row-${index}`, amount: index })), {
    columns, groupBy: Array.from({ length: depth }, (_, level) => ({ key: `hidden-${level}`, value: () => `Group-${level}` })), totals: { amount: "sum" }
  });
  for (const depth of [1, 4]) for (const composed of [false, true]) {
    const pdf = await inspectPdf(await (composed
      ? writePdfDocument([pdfGroupedTable(source(5, depth), labels)], compact)
      : writeGroupedPdf(source(5, depth), { ...compact, ...labels })));
    const grand = pdf.pages.find(page => page.lines.includes("Grand"));
    assert.ok(pdf.pages.indexOf(grand) > 0, "The fixture must place the grand total on a later page");
    for (const heading of ["Summary", "Item", "Amount"]) assert.ok(grand.lines.includes(heading), heading);
  }
  const samePage = await inspectPdf(await writeGroupedPdf(source(1), { ...compact, ...labels }));
  assert.equal(samePage.pages.length, 1);
  for (const heading of ["Summary", "Item", "Amount"])
    assert.equal(samePage.pages[0].lines.filter(line => line === heading).length, 1, "No duplicate same-page headings");
  const unheaded = await inspectPdf(await writeGroupedPdf(source(5), { ...compact, ...labels, includeHeader: false }));
  for (const heading of ["Summary", "Item", "Amount"]) assert.ok(!unheaded.text.includes(heading), "Explicit header omission is preserved");
  const split = await inspectPdf(await writeGroupedPdf(source(5), { ...compact, ...labels,
    grandTotalLabel: () => Array.from({ length: 36 }, (_, index) => `Grand-${index}`).join("\n") }));
  const totalPages = split.pages.filter(page => page.lines.some(line => /^Grand-/.test(line)));
  assert.ok(totalPages.length > 1);
  for (const page of totalPages) for (const heading of ["Summary", "Item", "Amount"]) assert.ok(page.lines.includes(heading), heading);
  assert.deepEqual(totalPages.flatMap(page => page.lines).filter(line => /^Grand-/.test(line)), Array.from({ length: 36 }, (_, index) => `Grand-${index}`));
});

test("writer-only cancellation returns an unresolved grouped producer and releases the borrowed destination", async () => {
  let returned = 0, entered;
  const started = new Promise(resolve => { entered = resolve; });
  const source = { [Symbol.asyncIterator]() { return {
    next() { entered(); return new Promise(() => {}); },
    return() { returned++; return Promise.resolve({ done: true }); }
  }; } };
  const abort = new AbortController(), reason = Error("Grouped writer cancelled"), destination = new WritableStream({ write() {} });
  const writing = writeGroupedPdfTo(report(source), destination, { ...options, signal: abort.signal });
  await started; abort.abort(reason);
  await assert.rejects(writing, error => error === reason);
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(returned, 1); assert.equal(destination.locked, false);
});

test("source-only cancellation interrupts a stalled grouped PDF sink and returns the active input", { timeout: 3000 }, async () => {
  const abort = new AbortController(), reason = Error("Source report cancelled");
  let produced = 0, returned = 0, entered;
  const stalled = new Promise(resolve => { entered = resolve; });
  function* source() { try { for (let i = 0; i < 200; i++) { produced++; yield { team: "A", name: `Item ${i}`, amount: i }; } }
    finally { returned++; } }
  const grouped = createGroupedReport(source(), { columns, groupBy: "team", signal: abort.signal, totals: { amount: "average" } });
  const destination = new WritableStream({ write() { if (produced > 10) { entered(); return new Promise(() => {}); } } });
  const writing = writeGroupedPdfTo(grouped, destination, options);
  await stalled; abort.abort(reason);
  await assert.rejects(writing, error => error === reason);
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(returned, 1); assert.equal(destination.locked, false);
});

test("composed hierarchical PDFs repeat all hidden-key ancestors and leaf headings in one document session", async () => {
  const long = Array.from({ length: 90 }, (_, index) => `Detail ${index}`).join("\n");
  const data = [{ region: "R", site: "S", environment: "E", team: "T", name: long, amount: 2 },
    { region: "R", site: "S", environment: "E", team: "T", name: "Second", amount: 4 },
    { region: "R", site: "S", environment: "E", team: "U", name: "Last", amount: 12 }];
  const selected = [{ header: "Item", key: "name" }, { header: "Measure", key: "amount", type: "number" }];
  const make = () => createGroupedReport(data, { columns: selected, totals: { amount: "average" },
    groupBy: ["region", "site", "environment", "team"].map((key, level) => ({ key, value: row => row[key], label: value => `${level}:${value}` })) });
  const labels = { labelColumn: { header: "Summary" }, subtotalLabel: group => `S${group.level}:${group.label}(${group.rows})`,
    grandTotalLabel: summary => `All(${summary.rows})`, continuationLabel(group) {
      assert.equal(group.path.length, group.level + 1); assert.ok(Object.isFrozen(group)); return group.label;
    } };
  const blocks = grouped => [
    { kind: "paragraph", text: "Executive summary" },
    { kind: "vector", width: 60, height: 20, commands: [{ op: "move", x: 0, y: 20 }, { op: "line", x: 60, y: 0 }], stroke: "1155aa", decorative: true },
    pdfTable([["Overview"]], { columns: [{ header: "Status" }] }),
    { kind: "section", title: "Appendix", includeHeading: false, bookmark: false }, pdfGroupedTable(grouped, labels),
    { kind: "paragraph", text: "Recommendations" }, { kind: "list", items: [{ text: "Continue monitoring" }] }
  ];
  const chunks = [], progress = [];
  const result = await writePdfDocumentTo(blocks(make()), { write: bytes => chunks.push(bytes.slice()) },
    { ...options, limits: { maxRows: 10 }, onProgress: value => progress.push(value) });
  assert.equal(result.rows, 4, "Ordinary and analytical source rows share result counts");
  assert.equal(result.columns, 3); assert.equal(progress.at(-1).rows, 4);
  const pdf = await inspectPdf(new Blob(chunks));
  for (const text of ["Executive summary", "Overview", "S3:3:T(2)", "S3:3:U(1)", "S2:2:E(3)", "S1:1:S(3)", "S0:0:R(3)", "All(3)", "Recommendations", "Continue monitoring"])
    assert.ok(pdf.text.includes(text), text);
  const continued = pdf.pages.filter(page => page.lines.some(line => /^Detail /.test(line)));
  assert.ok(continued.length >= 2);
  for (const page of continued) for (const label of ["0:R", "1:S", "2:E", "3:T", "Summary", "Item", "Measure"])
    assert.ok(page.lines.includes(label), `Repeated context: ${label}`);
  assert.deepEqual(continued.flatMap(page => page.lines).filter(line => /^Detail /.test(line)), long.split("\n"));
  assert.ok(pdf.pages.every(page => !page.lines.includes("region") && !page.lines.includes("environment")), "Hidden grouping keys never become columns");
  await assert.rejects(writePdfDocument(blocks(make()), { ...options, limits: { maxRows: 9 } }), /maxRows/);
  await assert.rejects(writePdfDocument([pdfGroupedTable(make(), labels)], { ...options, tagged: { language: "en" } }), /Tagged grouped/);
});

test("last detail and every closing subtotal stay together when they fit, with bounded oversized fallback", async () => {
  const selected = [{ header: "Item", key: "name" }, { header: "Measure", key: "amount", type: "number" }];
  const source = rows => createGroupedReport(rows, { columns: selected, groupBy: [
    { key: "region", value: row => row.region }, { key: "team", value: row => row.team }
  ], totals: { amount: "sum" } });
  const labels = { labelColumn: { header: "Summary" }, subtotalLabel: group => `Subtotal-${group.level}`, grandTotalLabel: () => "Grand" };
  const compact = { pageSize: { width: 320, height: 220 }, orientation: "landscape", margins: { top: 20, right: 20, bottom: 20, left: 20 },
    pageNumbers: false, compression: false, fontSize: 10 };
  const rows = Array.from({ length: 8 }, (_, index) => ({ region: "Outer", team: "Inner", name: index === 7 ? "LAST" : `Row-${index}`, amount: 1 }));
  const pdf = await inspectPdf(await writeGroupedPdf(source(rows), { ...compact, ...labels }));
  assert.ok(pdf.pages.length > 1);
  const last = pdf.pages.find(page => page.lines.includes("LAST"));
  for (const subtotal of ["Subtotal-1", "Subtotal-0"]) assert.ok(last.lines.includes(subtotal), subtotal);
  const one = await inspectPdf(await writePdfDocument([
    { kind: "paragraph", text: Array.from({ length: 6 }, (_, index) => `Preface-${index}`).join("\n") },
    pdfGroupedTable(source([{ region: "Outer", team: "Inner", name: "Only", amount: 1 }]), { ...labels, includeGrandTotal: false })
  ], compact));
  assert.equal(one.pages.length, 2);
  assert.ok(!one.pages[0].lines.includes("Outer"), "A kept one-row leaf does not leave its ancestor caption/headings on the preceding page");
  for (const text of ["Outer", "Inner", "Only", "Subtotal-1", "Subtotal-0"]) assert.ok(one.pages[1].lines.includes(text), text);
  const body = Array.from({ length: 80 }, (_, index) => `Long-${index}`).join("\n");
  const oversized = await inspectPdf(await writeGroupedPdf(source([{ region: "Outer", team: "Inner", name: body, amount: 1 }]), { ...compact, ...labels }));
  assert.deepEqual(oversized.pages.flatMap(page => page.lines).filter(line => /^Long-/.test(line)), body.split("\n"));
  for (const page of oversized.pages.filter(page => page.lines.some(line => /^Long-/.test(line))))
    for (const label of ["Outer", "Inner", "Summary", "Item", "Measure"]) assert.ok(page.lines.includes(label), label);
  assert.ok(oversized.text.includes("Subtotal-1") && oversized.text.includes("Subtotal-0"));
});

test("unreached and completed grouped blocks leave no source listeners or iterator ownership behind", async () => {
  const abort = new AbortController(), signal = abort.signal;
  const add = signal.addEventListener.bind(signal), remove = signal.removeEventListener.bind(signal), active = new Set();
  signal.addEventListener = (name, listener, settings) => { if (name === "abort") active.add(listener); return add(name, listener, settings); };
  signal.removeEventListener = (name, listener, settings) => { if (name === "abort") active.delete(listener); return remove(name, listener, settings); };
  let produced = 0, returned = 0;
  function* source() { try { produced++; yield { team: "A", name: "One", amount: 2 }; } finally { returned++; } }
  const grouped = createGroupedReport(source(), { columns, groupBy: "team", signal, totals: { amount: "sum" } });
  const labelColumn = { header: "Captured" }, block = pdfGroupedTable(grouped, { labelColumn, subtotalLabel: () => "Subtotal", grandTotalLabel: () => "Grand" });
  labelColumn.header = "Changed";
  assert.equal(active.size, 0);
  await assert.rejects(writePdfDocument([{ kind: "paragraph", text: 123 }, block], options), /text/i);
  assert.equal(produced, 0); assert.equal(active.size, 0);
  async function* blocks() {
    yield block;
    assert.equal(active.size, 0, "Source listeners end with their analytical block");
    abort.abort(Error("Completed source no longer owns the document"));
    yield { kind: "paragraph", text: "After source completion" };
  }
  const pdf = await inspectPdf(await writePdfDocument(blocks(), options));
  assert.equal(produced, 1); assert.equal(returned, 1); assert.equal(active.size, 0);
  assert.ok(pdf.text.includes("Captured") && !pdf.text.includes("Changed") && pdf.text.includes("After source completion"));
});

test("composed grouped cancellation returns a pending source and releases destination without waiting for return", { timeout: 3000 }, async () => {
  let returned = 0, entered;
  const started = new Promise(resolve => { entered = resolve; });
  const source = { [Symbol.asyncIterator]() { return {
    next() { entered(); return new Promise(() => {}); },
    return() { returned++; return new Promise(() => {}); }
  }; } };
  const abort = new AbortController(), reason = Error("Composed grouped cancellation"), destination = new WritableStream({ write() {} });
  const writing = writePdfDocumentTo([{ kind: "paragraph", text: "Before" }, pdfGroupedTable(report(source), {
    labelColumn: { header: "Summary" }, subtotalLabel: () => "Subtotal", grandTotalLabel: () => "Grand"
  })], destination, { ...options, signal: abort.signal });
  await started; abort.abort(reason);
  await assert.rejects(writing, error => error === reason);
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(returned, 1); assert.equal(destination.locked, false);
});

test("a composed grouped destination failure closes the active source and preserves the sink's error", async () => {
  let produced = 0, returned = 0;
  function* source() { try {
    for (let index = 0; index < 200; index++) { produced++; yield { team: "A", name: `Item ${index}`, amount: index }; }
  } finally { returned++; } }
  const reason = Error("Destination refused grouped page"), destination = new WritableStream({ write() { if (produced > 10) throw reason; } });
  await assert.rejects(writePdfDocumentTo([pdfGroupedTable(report(source()), {
    labelColumn: { header: "Summary" }, subtotalLabel: () => "Subtotal", grandTotalLabel: () => "Grand"
  })], destination, options), error => error === reason);
  await new Promise(resolve => setImmediate(resolve));
  assert.ok(produced < 200); assert.equal(returned, 1); assert.equal(destination.locked, false);
});

test("page-number literals are captured, language independent, inherited by sections and geometrically bounded", async () => {
  let reads = 0;
  const labels = { get prefix() { reads++; return "Strona "; }, separator: " z ", suffix: " stron" };
  const pdf = await inspectPdf(await writePdfDocument([
    pdfTable([["First"]], { columns: [{ header: "Value" }] }),
    { kind: "section", title: "Other", includeHeading: false, bookmark: false },
    pdfTable([["Last"]], { columns: [{ header: "Value" }] })
  ], { pageNumberLabels: labels, compression: false }));
  assert.equal(reads, 1); assert.equal(pdf.pages.length, 2);
  for (const [i, page] of pdf.pages.entries()) assert.ok(page.lines.includes(`Strona ${i + 1} z `));
  const total = [...pdf.objects.values()].find(object => /\/Subtype \/Form/.test(object.body));
  assert.ok(total.stream.toString().includes(`<${Buffer.from("2 stron").toString("hex")}>`), "The total-count form retains the authored suffix");
  const source = [["One"]], base = { columns: [{ header: "Value" }], compression: false };
  const defaultPdf = await inspectPdf(await writePdf(source, base));
  const explicit = await inspectPdf(await writePdf(source, { ...base, pageNumberLabels: { prefix: "Page ", separator: " of ", suffix: "" } }));
  assert.deepEqual(explicit.pages.map(page => page.content), defaultPdf.pages.map(page => page.content));
  await assert.rejects(writePdf(source, { ...base, pageNumberLabels: { prefix: "X".repeat(1000) } }), /too narrow/);
  await assert.rejects(writePdf(source, { ...base, pageNumberLabels: { suffix: 3 } }), /must be text/);
  await assert.rejects(writePdf(source, { ...base, pageNumberLabels: { suffix: "LONG" }, limits: { maxCellCharacters: 3 } }), /maxCellCharacters/);
});
