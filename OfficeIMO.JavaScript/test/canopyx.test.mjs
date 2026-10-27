import test from "node:test";
import assert from "node:assert/strict";
import { writeFile } from "node:fs/promises";
import { createCanopyExport, exportCanopy, writeCanopyTo } from "../dist/integrations/canopyx/index.js";
import { createGroupedReport } from "../dist/core/index.js";
import { readZip } from "./zip-reader.mjs";
import { inspectPdf } from "./pdf-reader.mjs";

const column = (id, kind = "text") => ({ id, title: id, kind, width: { unit: "css-px", preferred: 160, minimum: 80, maximum: 240 }, wrap: true, alignment: kind === "number" ? "right" : "left" });
const cell = (value, text = String(value ?? ""), extra = {}) => ({ value, text, ...extra });
function capture(rows, columns = [column("Name")], request = {}) {
  return { request: { columns, recordCount: rows.length, values: "raw", presentation: "semantic", timeZone: "utc", revision: "snapshot-1", ...request },
    async *rows({ signal } = {}) { for (const row of rows) { signal?.throwIfAborted(); yield row; } } };
}
const collect = async source => { const rows = []; for await (const row of source.rows()) rows.push(row); return rows; };
async function fixture(name, blob) {
  if (process.env.OFFICEIMO_CANOPY_FIXTURES) await writeFile(process.env.OFFICEIMO_CANOPY_FIXTURES + "/" + name, new Uint8Array(await blob.arrayBuffer()));
}

test("Canopy report links and histories export visible labels and frozen latest values", async () => {
  const input = capture([{ id: "record-1", cells: {
    Name: cell("#record/42", "Record 42", { diagnostics: ["relative-link"] }),
    History: cell(12.5, "12.50 ms", { diagnostics: ["history-series"] }),
    EmptyHistory: cell(null, "No samples", { diagnostics: ["history-series"] })
  } }], [column("Name"), column("History", "number"), column("EmptyHistory", "number")]);
  const diagnostics = [];
  const blob = await exportCanopy(input, "xlsx", { onDiagnostic: item => diagnostics.push(item), xlsx: { compression: "store" } });
  const zip = await readZip(blob), xml = zip.get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /Record 42/); assert.doesNotMatch(xml, /#record\/42/);
  assert.match(xml, /<c r="B2"[^>]*><v>12\.5<\/v><\/c>/);
  assert.doesNotMatch(xml, /<c r="C2"[^>]*><v>/);
  assert.doesNotMatch(xml, /<hyperlink\b/);
  assert.deepEqual(diagnostics.map(item => item.code), ["relative-link", "history-series", "history-series"]);
  assert.ok(diagnostics.every(Object.isFrozen));
  assert.equal(await (await exportCanopy(input, "csv")).text(), "Name,History,EmptyHistory\r\nRecord 42,12.5,\r\n");
  await fixture("canopy-report-presentations.xlsx", blob);
});

test("Canopy unmapped tones retain status words and typed values without inventing styles", async () => {
  const input = capture([{ id: "status-1", tone: "unmapped-row", cells: {
    Status: cell("Not evaluated", "Not evaluated", { tone: "could-not-evaluate" }),
    Amount: cell(4, "4 items", { tone: "custom-count", diagnostics: ["relative-link"] })
  } }], [column("Status"), column("Amount", "number")]);
  const diagnostics = [];
  const rows = await collect(createCanopyExport(input, "xlsx", { onDiagnostic: item => diagnostics.push(item) }));
  assert.equal(rows[0][0].value, "Not evaluated"); assert.equal(rows[0][1].value, 4);
  assert.equal(rows[0][0].presentation, undefined); assert.equal(rows[0][1].presentation, undefined);
  assert.deepEqual(diagnostics.map(item => item.code), ["UNMAPPED_TONE:unmapped-row", "UNMAPPED_TONE:could-not-evaluate", "relative-link", "UNMAPPED_TONE:custom-count"]);
  const zip = await readZip(await exportCanopy(input, "xlsx"));
  const xml = zip.get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /Not evaluated/); assert.match(xml, /<c r="B2"[^>]*><v>4<\/v><\/c>/);
});

test("Canopy report fallbacks keep unsafe links and genuine unsupported presentation rejected", async () => {
  const input = diagnostics => capture([{ id: "r", cells: { Name: cell("#r", "=visible", { diagnostics }) } }]);
  assert.equal(await (await exportCanopy(input(["relative-link"]), "csv")).text(), "Name\r\n'=visible\r\n");
  await assert.rejects(exportCanopy(input(["relative-link", "unsafe-link"]), "xlsx"), /unsafe-link/);
  await assert.rejects(exportCanopy(input(["history-series", "custom-renderer"]), "xlsx"), /custom-renderer/);
  const unsafe = capture([{ id: "r", cells: { Name: cell("#r", "Visible", { diagnostics: ["relative-link"], link: { target: "javascript:alert(1)" } }) } }]);
  await assert.rejects(exportCanopy(unsafe, "xlsx"), /absolute HTTP/);
  await assert.rejects(exportCanopy(input(["relative-link"]), "xlsx", { onDiagnostic: async () => {} }), /synchronous/);
});

test("Canopy relative link labels include null and non-string text-column inputs", async () => {
  const rows = [null, false, 42].map((value, index) => ({ id: "r" + index,
    cells: { Name: cell(value, "View record " + index, { diagnostics: ["relative-link"] }) } }));
  const input = capture(rows);
  const zip = await readZip(await exportCanopy(input, "xlsx"));
  const xml = zip.get("xl/worksheets/sheet1.xml").content;
  for (let index = 0; index < rows.length; index++) assert.ok(xml.includes("View record " + index));
  assert.equal(await (await exportCanopy(input, "csv")).text(), "Name\r\nView record 0\r\nView record 1\r\nView record 2\r\n");
});

test("Canopy grouped relative links preserve raw identities and reject mismatched captured ranges", async () => {
  const make = (values) => capture([
    { id: "r1", cells: { Name: cell(values[0], "Shared label", { diagnostics: ["relative-link"] }), "Name#raw": cell("Unrelated") } },
    { id: "r2", cells: { Name: cell(values[1], "Shared label", { diagnostics: ["relative-link"] }), "Name#raw": cell("Unrelated") } }
  ], [column("Name"), column("Name#raw")], { grouping: { levels: [{ column: "Name", direction: "asc" }], aggregates: [], totals: {}, groups: [
    { key: "first", value: values[0], text: "First record", offset: 0, count: 1, aggregates: {} },
    { key: "second", value: values[1], text: "Second record", offset: 1, count: 1, aggregates: {} }
  ] } });
  for (const values of [["#record/42", "#record/43"], [null, false]]) for (const format of ["xlsx", "pdf"]) {
    const blob = await exportCanopy(make(values), format);
    const text = format === "xlsx" ? (await readZip(blob)).get("xl/worksheets/sheet1.xml").content : (await inspectPdf(blob)).text;
    assert.ok(text.includes("First record") && text.includes("Second record"));
    assert.equal((text.match(/Shared label/g) ?? []).length, 2);
    assert.ok(!text.includes("#record/42") && !text.includes("#record/43"));
    const bad = make(values); bad.request.grouping.groups[0].value = "#record/wrong";
    await assert.rejects(exportCanopy(bad, format), /disagrees with its captured group/);
  }
});

test("Canopy capture preserves column order, own keys, nulls and reusable enumeration without query evaluation", async () => {
  const columns = [column("Amount", "number"), column("__proto__"), column("Name")];
  const rows = [{ id: "r1", cells: { Name: cell(null, "Missing"), ["__proto__"]: cell("Literal"), Amount: cell(12.5, "12.50 USD") } }];
  const input = capture(rows, columns), source = createCanopyExport(input, "csv");
  columns.reverse(); columns[0].title = "Mutated";
  assert.deepEqual(source.columns.map(c => [c.key, c.header, c.width]), [["Amount", "Amount", undefined], ["__proto__", "__proto__", undefined], ["Name", "Name", undefined]]);
  assert.deepEqual(await collect(source), [[12.5, "Literal", null]]);
  assert.deepEqual(await collect(source), [[12.5, "Literal", null]]);
  assert.ok(Object.isFrozen(source.columns) && Object.isFrozen(source.columns[0]));
  assert.equal(createCanopyExport(capture([], [column("Name")]), "xlsx", { columnOptions: { Amount: { format: "0.00" } } }).columns.length, 1);
  await assert.rejects(collect(createCanopyExport(capture([{ id: "r", cells: {} }]), "csv")), /missing column/);
  for (const value of [NaN, Infinity, Number.MAX_SAFE_INTEGER + 1, -Number.MAX_SAFE_INTEGER - 1])
    await assert.rejects(collect(createCanopyExport(capture([{ id: "r", cells: { Name: cell(value) } }]), "csv")), /scalar value/);
});

test("Canopy raw and display CSV remain literal, including formula protection after projection", async () => {
  const row = { id: "r", cells: { Name: cell("Raw", "=display"), Amount: cell(12.5, "12.50 USD"), Seen: cell("2026-10-08T12:34:56.1234567+02:00", "Yesterday") } };
  const cols = [column("Name"), column("Amount", "number"), column("Seen", "datetime")];
  const raw = await exportCanopy(capture([row], cols), "csv"), display = await exportCanopy(capture([row], cols, { values: "display" }), "csv");
  assert.equal(await raw.text(), "Name,Amount,Seen\r\nRaw,12.5,2026-10-08T12:34:56.1234567+02:00\r\n");
  assert.equal(await display.text(), "Name,Amount,Seen\r\n'=display,12.50 USD,Yesterday\r\n");
  await fixture("canopy-raw.csv", raw); await fixture("canopy-display.csv", display);
});

test("Canopy raw numeric metadata supports grouped metrics without typing rendered display or preserved dates", async () => {
  const precise = "2026-10-08T12:34:56.123456789Z";
  const cols = [column("Site"), { ...column("Latency", "number"), title: "Opóźnienie" }, column("Seen", "datetime")];
  const rows = [{ id: "a", cells: { Site: cell("A"), Latency: cell(12.5, "12,50 ms"), Seen: cell(precise, "Local date") } },
    { id: "b", cells: { Site: cell("A"), Latency: cell(null, "Missing"), Seen: cell(precise, "Local date") } }];
  for (const format of ["csv", "xlsx", "pdf"]) {
    const raw = createCanopyExport(capture(rows, cols), format);
    assert.deepEqual(raw.columns.map(c => c.type), [undefined, "number", undefined]);
    const grouped = createGroupedReport(raw.rows(), { columns: raw.columns, groupBy: "Site", totals: { Latency: "average" } });
    const events = []; for await (const event of grouped.events) events.push(event);
    assert.deepEqual(events.at(-1).totals, [null, 12.5, null]);
    const firstDate = events.find(event => event.kind === "row").values[2];
    assert.equal(format === "csv" ? firstDate : firstDate.value, precise, "Date preservation must not force a Date column");
    const display = createCanopyExport(capture(rows, cols, { values: "display" }), format);
    assert.ok(display.columns.every(c => c.type === undefined));
    const rendered = (await collect(display))[0][1];
    assert.equal(format === "csv" ? rendered : rendered.value, "12,50 ms");
  }
  const contradictory = capture([{ id: "bad", cells: { Site: cell("A"), Latency: cell("12.5"), Seen: cell(precise) } }], cols);
  await assert.rejects(exportCanopy(contradictory, "xlsx"), /does not match column type number/);
});

test("Canopy selected column roles and presentation overrides are captured once without reading unselected options", async () => {
  let roleReads = 0, groupReads = 0, optionsReads = 0, role = true;
  const groups = ["Identity"], override = Object.create({
    get rowHeader() { roleReads++; return role; }, get groups() { groupReads++; return groups; }
  });
  const columnOptions = { Name: override, get Private() { throw Error("Unselected options were read"); } };
  const options = { get columnOptions() { optionsReads++; return columnOptions; } };
  const cols = [column("Amount", "number"), column("Name")];
  const input = capture([{ id: "r", cells: { Name: cell("SVC-1"), Amount: cell(12.5) } }], cols);
  const source = createCanopyExport(input, "csv", options);
  role = false; groups[0] = "Changed"; cols.reverse();
  assert.equal(optionsReads, 1); assert.equal(roleReads, 1); assert.equal(groupReads, 1);
  assert.deepEqual(source.columns.map(column => [column.key, column.rowHeader]), [["Amount", undefined], ["Name", true]]);
  assert.deepEqual(source.columns[1].groups, ["Identity"]);
  assert.deepEqual(await collect(source), [[12.5, "SVC-1"]]);
  assert.throws(() => createCanopyExport(input, "csv", { columnOptions: { Name: { rowHeader: "yes" } } }), /rowHeader must be boolean/);
});

test("Canopy row and cell tones compose without replacing values or number formats", async () => {
  const style = { background: "E2F0D9", bold: true, numberFormat: "0.00" };
  const input = capture([{ id: "r", tone: "good", cells: { Name: cell(12.5, "12.50 USD", { tone: "alert", link: { target: "https://example.com/report" } }) } }]);
  const source = createCanopyExport(input, "xlsx", { tones: { good: style, alert: { color: "9C0006", bold: false } } });
  style.background = "FFFFFF";
  const [[value]] = await collect(source);
  assert.equal(value.value, 12.5); assert.equal(value.text, "12.50 USD");
  assert.deepEqual(value.presentation, { background: "E2F0D9", bold: false, numberFormat: "0.00", color: "9C0006" });
  assert.equal(value.link.target, "https://example.com/report");
});

test("Canopy tones capture structural presentation once before rows are consumed", async () => {
  let reads = 0, background = "FFF2CC", bold = true;
  class Warning { get background() { reads++; return background; } get bold() { reads++; return bold; } }
  const input = capture([{ id: "r", tone: "warning", cells: { Name: cell(12.5, "12.50") } }]);
  for (const format of ["xlsx", "pdf"]) {
    reads = 0; background = "FFF2CC"; bold = true;
    const source = createCanopyExport(input, format, { tones: { warning: new Warning() } });
    assert.equal(reads, 2);
    background = "FFFFFF"; bold = false;
    const [[value]] = await collect(source);
    assert.deepEqual(value.presentation, { background: "FFF2CC", bold: true });
    assert.equal(value.value, 12.5); assert.equal(reads, 2);
  }
  assert.throws(() => createCanopyExport(input, "xlsx", { tones: { warning: { bold: "yes" } } }), /must be boolean/);
});

test("Canopy unsupported presentation rejects by default and explicit text mode reports bounded diagnostics", async () => {
  const input = capture([{ id: "r", diagnostics: ["CUSTOM_ROW"], tone: "custom", cells: { Name: cell("Raw", "Visible", { diagnostics: ["CUSTOM_CELL"] }) } }]);
  await assert.rejects(collect(createCanopyExport(input, "pdf")), /CUSTOM_ROW/);
  const received = [];
  assert.equal((await collect(createCanopyExport(input, "pdf", { unsupportedPresentation: "text", onDiagnostic: d => received.push(d) })))[0][0].text, "Visible");
  assert.deepEqual(received.map(d => [d.code, d.rowId, d.columnId]), [["CUSTOM_ROW", "r", undefined], ["UNMAPPED_TONE:custom", "r", undefined], ["CUSTOM_CELL", "r", "Name"]]);
  assert.ok(received.every(Object.isFrozen));
  const ordinary = capture([], [{ id: "Name", title: "Name", kind: "text" }], { presentation: "text" });
  assert.equal(createCanopyExport(ordinary, "xlsx").rowCount, 0);
  assert.equal(createCanopyExport(ordinary, "xlsx", { unsupportedPresentation: "text" }).rowCount, 0);
  assert.equal(await (await exportCanopy(ordinary, "csv")).text(), "Name\r\n");
  assert.equal(createCanopyExport(ordinary, "csv", { unsupportedPresentation: "reject" }).rowCount, 0);
  await assert.rejects(collect(createCanopyExport(input, "pdf", { unsupportedPresentation: "text", onDiagnostic: async () => { throw new Error("async notification"); } })), /synchronous/);
});

test("Canopy explicit datetime intent preserves ISO precision and rejects invalid typed dates", async () => {
  const input = value => capture([{ id: "r", cells: { Seen: cell(value, "Resolved clock") } }], [column("Seen", "datetime")]);
  for (const iso of ["2024-02-29T12:34:56.123+02:00", "2026-10-08T01:02:03Z", "2026-10-08T01:02:03.123000000-05:30", "2026-10-08T12:34Z", "2026-12-31T24:00Z"]) {
    const [[value]] = await collect(createCanopyExport(input(iso), "xlsx"));
    assert.ok(value.value instanceof Date); assert.equal(value.value.getTime(), new Date(iso).getTime());
    assert.equal(value.text, "Resolved clock");
  }
  const precise = "2026-10-08T12:34:56.1234567+02:00";
  const nano = "2026-10-08T12:34:56.123456789Z";
  assert.equal((await collect(createCanopyExport(input(nano), "xlsx")))[0][0].value, nano);
  await assert.rejects(collect(createCanopyExport(input(nano), "xlsx", { datetime: "typed" })), /sub-millisecond/);
  assert.equal((await collect(createCanopyExport(input(precise), "xlsx")))[0][0].value, precise);
  await assert.rejects(collect(createCanopyExport(input(precise), "xlsx", { datetime: "typed" })), /sub-millisecond/);
  const ancient = "0001-01-01T00:00:00Z";
  assert.equal((await collect(createCanopyExport(input(ancient), "xlsx")))[0][0].value, ancient);
  await assert.rejects(collect(createCanopyExport(input(ancient), "xlsx", { datetime: "typed" })), /1900 through 9999/);
  assert.equal((await collect(createCanopyExport(input(precise), "xlsx", { datetime: "text" })))[0][0].value, precise);
  for (const invalid of ["2023-02-29T01:00:00Z", "2026-10-08T24:00:01Z", "2026-10-08T24:00:00.000001Z", "2026-10-08T01:00:00+24:00", "Yesterday", "2026-10-08T01:00:00"]) {
    await assert.rejects(collect(createCanopyExport(input(invalid), "xlsx")), /datetime/);
  }
  assert.equal((await collect(createCanopyExport(input(precise), "xlsx", { datetime: "text" })))[0][0].value, precise);
});

test("Canopy preservation uses the writer's effective local or UTC year at both Excel boundaries", async () => {
  const previous = process.env.TZ;
  try {
    for (const [zone, iso, year] of [["America/New_York", "1900-01-01T00:00:00Z", 1899], ["Pacific/Auckland", "9999-12-31T23:00:00Z", 10000]]) {
      process.env.TZ = zone;
      assert.equal(new Date(iso).getFullYear(), year);
      const local = capture([{ id: "r", cells: { Seen: cell(iso) } }], [column("Seen", "datetime")], { timeZone: "local" });
      assert.equal((await collect(createCanopyExport(local, "xlsx")))[0][0].value, iso);
      await assert.rejects(exportCanopy(local, "xlsx", { datetime: "typed" }), /1900 through 9999/);
      const workbook = await readZip(await exportCanopy(local, "xlsx"));
      assert.ok(workbook.get("xl/worksheets/sheet1.xml").content.includes(iso));
      assert.ok((await collect(createCanopyExport(local, "xlsx", { xlsx: { dateMode: "utc" } })))[0][0].value instanceof Date);
      await exportCanopy(local, "xlsx", { xlsx: { dateMode: "utc" } });
      const utc = { ...local, request: { ...local.request, timeZone: "utc" } };
      assert.equal((await collect(createCanopyExport(utc, "xlsx", { xlsx: { dateMode: "local" } })))[0][0].value, iso);
    }
  } finally { if (previous === undefined) delete process.env.TZ; else process.env.TZ = previous; }
});

test("Canopy recordCount is checked without retaining IDs or recalculating selection", async () => {
  const rows = [{ id: "r", cells: { Name: cell("Value") } }];
  for (const count of [0, 2]) await assert.rejects(collect(createCanopyExport(capture(rows, undefined, { recordCount: count }), "csv")), /recordCount/);
  assert.deepEqual(await collect(createCanopyExport(capture(rows, undefined, { recordCount: null }), "csv")), [["Value"]]);
  for (const count of [-1, 1.5, NaN, undefined]) assert.throws(() => createCanopyExport(capture([], undefined, { recordCount: count }), "csv"), /metadata/);
});

test("Canopy cancellation reaches the in-flight native source and returns the iterator", async () => {
  const controller = new AbortController(), reason = new Error("Cancelled page"), input = capture([], undefined, { recordCount: null });
  let seen, returned = false;
  input.rows = ({ signal }) => {
    seen = signal;
    return { [Symbol.asyncIterator]() { return { next() { return new Promise((_, reject) => signal.addEventListener("abort", () => reject(signal.reason), { once: true })); },
      return() { returned = true; return Promise.resolve({ done: true }); } }; } };
  };
  const pending = collect(createCanopyExport(input, "csv", { signal: controller.signal }));
  await new Promise(resolve => setImmediate(resolve)); controller.abort(reason);
  await assert.rejects(pending, error => error === reason);
  assert.equal(seen, controller.signal); assert.equal(returned, true);
});

test("Canopy streamed destinations enforce backpressure, resource limits and source cleanup", async () => {
  let read = 0, closed = 0, release, started;
  const blocked = new Promise(resolve => { release = resolve; }), writing = new Promise(resolve => { started = resolve; });
  const input = capture([], undefined, { recordCount: 100 });
  input.rows = async function* () { try { for (let i = 0; i < 100; i++) { read++; yield { id: String(i), cells: { Name: cell("x".repeat(10000)) } }; } } finally { closed++; } };
  let writes = 0;
  const pending = writeCanopyTo(input, "csv", { async write() { if (++writes === 1) { started(); await blocked; } } });
  await writing; assert.ok(read < 100); release();
  assert.equal((await pending).rows, 100); assert.equal(closed, 1);
  read = 0; closed = 0;
  await assert.rejects(exportCanopy(input, "csv", { limits: { maxRows: 2 } }), /maxRows/);
  assert.equal(closed, 1); assert.ok(read < 100);
  const failure = new Error("Sink failed");
  await assert.rejects(writeCanopyTo(input, "csv", { write() { throw failure; } }), error => error === failure);
  assert.equal(closed, 2);
});

test("Canopy writer formats preserve typed XLSX, resolved PDF text, links and known progress totals", async () => {
  const columns = [column("Amount", "number"), column("Seen", "datetime")], rows = [{ id: "r", tone: "good", cells: {
    Amount: cell(12.5, "12.50 USD", { link: { target: "https://example.com/report" } }), Seen: cell("2026-10-08T01:02:03.123Z", "08 Oct 2026") } }];
  const input = capture(rows, columns), options = { tones: { good: { background: "E2F0D9" } }, columnOptions: { Amount: { format: "0.00", width: 16 }, Seen: { format: "yyyy-mm-dd hh:mm:ss.000", width: 28 } } };
  const progress = [], xlsx = await exportCanopy(input, "xlsx", { ...options, onProgress: p => progress.push(p), xlsx: { sheet: { title: { text: "Canopy report" }, table: {} }, compression: "store" } });
  const parts = await readZip(xlsx), xml = parts.get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /<c r="A3"[^>]*><v>12.5<\/v>/); assert.match(xml, /<c r="B3"[^>]*><v>\d+/); assert.match(xml, /<hyperlink ref="A3"/);
  assert.match(parts.get("xl/styles.xml").content, /E2F0D9/);
  assert.equal(progress.at(-1).phase, "complete"); assert.equal(progress.at(-1).totalRows, 1);
  const pdf = await exportCanopy(input, "pdf", { ...options, pdf: { compression: false } }), read = await inspectPdf(pdf);
  assert.ok(read.pages[0].lines.includes("12.50 USD") && read.pages[0].lines.includes("08 Oct 2026"));
  assert.match([...read.objects.values()].map(o => o.body).join("\n"), /\/URI <68747470733a/);
  await fixture("canopy-report.xlsx", xlsx); await fixture("canopy-report.pdf", pdf);
  await assert.rejects(exportCanopy(input, "csv", { ...options, onProgress: async () => { throw new Error("async progress"); } }), /synchronous/);
});

test("Canopy adapter retains format ceilings while explicit common limits override them", async () => {
  const input = capture([{ id: "r1", cells: { Name: cell("one") } }, { id: "r2", cells: { Name: cell("two") } }]);
  for (const format of ["xlsx", "csv", "pdf"]) {
    class Limits { get maxRows() { return 1; } }
    await assert.rejects(exportCanopy(input, format, { [format]: { limits: new Limits() }, limits: { maxRows: undefined } }), /maxRows/);
    await exportCanopy(input, format, { [format]: { limits: { maxRows: 1 } }, limits: { maxRows: 2 } });
  }
});
