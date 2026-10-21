import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { createGroupedReport } from "../dist/core/index.js";
import { writeGroupedPdf, writeGroupedPdfTo } from "../dist/pdf/grouped.js";
import { writePdf, writePdfDocument, pdfTable, PdfFont } from "../dist/pdf/index.js";
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
