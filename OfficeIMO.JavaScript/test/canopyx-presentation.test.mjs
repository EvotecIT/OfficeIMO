import test from "node:test";
import assert from "node:assert/strict";
import { exportCanopy } from "../dist/integrations/canopyx/index.js";
import { ExportCell } from "../dist/core/index.js";
import { readZip } from "./zip-reader.mjs";
import { inspectPdf } from "./pdf-reader.mjs";

const column = (id, kind = "text") => ({ id, title: id, kind, width: { unit: "css-px", preferred: 160, minimum: 80 }, wrap: false, alignment: kind === "number" ? "right" : "left" });
const cell = (value, text = String(value ?? ""), extra = {}) => ({ value, text, ...extra });
function capture(rows, columns = [column("Name")]) {
  return { request: { columns, recordCount: rows.length, values: "raw", presentation: "semantic", timeZone: "utc", revision: "r1" },
    async *rows() { yield* rows; } };
}

// Resolve a worksheet cell's style the way Excel does: cellXfs index -> font, fill, border and alignment records.
async function workbook(blob) {
  const zip = await readZip(blob), sheet = zip.get("xl/worksheets/sheet1.xml").content, styles = zip.get("xl/styles.xml").content;
  const section = name => new RegExp("<" + name + "\\b[^>]*>([\\s\\S]*?)</" + name + ">").exec(styles)[1];
  const fonts = [...section("fonts").matchAll(/<font>([\s\S]*?)<\/font>/g)].map(m => m[1]);
  const fills = [...section("fills").matchAll(/<fill>([\s\S]*?)<\/fill>/g)].map(m => m[1]);
  const borders = [...section("borders").matchAll(/<border>([\s\S]*?)<\/border>/g)].map(m => m[1]);
  const xfs = [...section("cellXfs").matchAll(/<xf\b([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g)].map(m => m);
  const strings = zip.has("xl/sharedStrings.xml") ? [...zip.get("xl/sharedStrings.xml").content.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => m[1].replace(/<[^>]+>/g, "")) : [];
  return { sheet, style(reference) {
    const element = new RegExp('<c r="' + reference + '"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)').exec(sheet);
    assert.ok(element, "missing cell " + reference);
    const index = Number(/ s="(\d+)"/.exec(element[1])?.[1] ?? 0), xf = xfs[index], attribute = name => new RegExp(name + '="(\\d+)"').exec(xf[1])?.[1];
    const body = element[2] ?? "", text = / t="s"/.test(element[1]) ? strings[Number(/<v>(\d+)<\/v>/.exec(body)[1])] : body.replace(/<[^>]+>/g, "");
    return { text, font: fonts[Number(attribute("fontId"))], fill: fills[Number(attribute("fillId"))],
      border: borders[Number(attribute("borderId"))], alignment: xf[2] ?? "" };
  } };
}
const fillColor = style => /<fgColor rgb="FF([0-9A-F]{6})"/i.exec(style.fill)?.[1]?.toUpperCase();
const fontColor = style => /<color rgb="FF([0-9A-F]{6})"/i.exec(style.font)?.[1]?.toUpperCase();
const pdfFill = hex => [0, 2, 4].map(i => Math.round(parseInt(hex.slice(i, i + 2), 16) / 255 * 10000) / 10000).join(" ");

const badgeCapture = () => capture([
  { id: "u1", cells: { Name: cell("Alice", "Alice", { badges: [{ text: "Disabled", tone: "danger" }, { text: "Stale" }], diagnostics: ["badges"] }),
    Logons: cell(12, "12", { badges: ["Rare"], diagnostics: ["badges", "custom-class"] }) } },
  { id: "u2", diagnostics: ["custom-class"], cells: { Name: cell("Bob", "Bob", { diagnostics: ["custom-renderer"] }), Logons: cell(3, "3") } }
], [column("Name"), column("Logons", "number")]);

test("Canopy badge, class and renderer presentations export their text in XLSX and PDF with diagnostics", async () => {
  for (const format of ["xlsx", "pdf"]) {
    const diagnostics = [];
    const blob = await exportCanopy(badgeCapture(), format, { onDiagnostic: item => diagnostics.push(item.code) });
    assert.deepEqual(diagnostics, ["badges", "badges", "custom-class", "custom-class", "custom-renderer"]);
    if (format === "xlsx") {
      const book = await workbook(blob);
      assert.equal(book.style("A2").text, "Alice · Disabled · Stale");
      // A typed number keeps its scalar; the badge stays in the display text only.
      assert.match(book.sheet, /<c r="B2"[^>]*><v>12<\/v><\/c>/);
      assert.equal(book.style("A3").text, "Bob");
    } else {
      const lines = (await inspectPdf(blob)).pages.flatMap(page => page.lines);
      assert.ok(lines.includes("Alice · Disabled · Stale"), lines.join("|"));
    }
  }
  assert.equal(await (await exportCanopy(badgeCapture(), "csv")).text(), "Name,Logons\r\nAlice · Disabled · Stale,12\r\nBob,3\r\n");
});

test("Canopy badges without labels still export, and malformed badges are rejected", async () => {
  const labelOnly = capture([{ id: "r", cells: { Name: cell(null, "", { badges: [{ text: "Locked" }], diagnostics: ["badges"] }) } }]);
  assert.equal((await workbook(await exportCanopy(labelOnly, "xlsx"))).style("A2").text, "Locked");
  const flagOnly = capture([{ id: "r", cells: { Name: cell("X", "X", { diagnostics: ["badges"] }) } }]);
  assert.equal((await workbook(await exportCanopy(flagOnly, "xlsx"))).style("A2").text, "X");
  const malformed = capture([{ id: "r", cells: { Name: cell("x", "X", { badges: [{ tone: "danger" }] }) } }]);
  await assert.rejects(exportCanopy(malformed, "xlsx"), /badges/);
});

test("Canopy cells with more than 64 badges still export, with one count diagnostic per cell", async () => {
  const labels = Array.from({ length: 100 }, (_, index) => "B" + index);
  const many = () => capture([{ id: "r", cells: { Name: cell("Alice", "Alice", { badges: labels.map(text => ({ text, tone: "info" })), diagnostics: ["badges"] }) } }]);
  const expected = ["Alice", ...labels.slice(0, 64)].join(" · ");
  for (const format of ["xlsx", "pdf", "csv"]) {
    const diagnostics = [];
    // Default policies: XLSX and PDF reject unsupported presentation, so the truncation must stay a text fallback.
    const blob = await exportCanopy(many(), format, { onDiagnostic: item => diagnostics.push([item.code, item.rowId, item.columnId]) });
    assert.deepEqual(diagnostics, [["badges", "r", "Name"], ["BADGES_TRUNCATED:36", "r", "Name"]], format);
    if (format === "xlsx") assert.equal((await workbook(blob)).style("A2").text, expected);
    else if (format === "csv") assert.equal(await blob.text(), "Name\r\n" + expected + "\r\n");
    else assert.ok((await inspectPdf(blob)).pages.flatMap(page => page.lines).join(" ").includes("B63"));
  }
});

test("Canopy text cells keep their captured label when every badge is left out", async () => {
  const text = "z".repeat(32767);
  const full = () => capture([{ id: "r", cells: { Name: cell("id-7", text, { badges: ["x"] }) } }]);
  for (const format of ["xlsx", "csv"]) {
    const diagnostics = [], blob = await exportCanopy(full(), format, { onDiagnostic: item => diagnostics.push(item.code) });
    const exported = format === "xlsx" ? (await workbook(blob)).style("A2").text : (await blob.text()).split("\r\n")[1];
    assert.equal(exported, text, format);
    assert.deepEqual(diagnostics, ["BADGES_TRUNCATED:1"]);
  }
});

test("Canopy badge labels never push a cell past the Excel text limit", async () => {
  const text = "x".repeat(32760);
  const long = capture([{ id: "r", cells: { Name: cell(text, text, { badges: ["Kept", "Dropped"] }) } }]);
  const diagnostics = [];
  const book = await workbook(await exportCanopy(long, "xlsx", { onDiagnostic: item => diagnostics.push(item.code) }));
  assert.equal(book.style("A2").text, text + " · Kept");
  assert.deepEqual(diagnostics, ["BADGES_TRUNCATED:1"]);
});

const linkRows = count => Array.from({ length: count }, (_, index) => ({ id: "r" + index, cells: { Name: cell("n" + index, "n" + index, { link: { target: "https://example.com/" + index } }) } }));
async function linkedSheet(input, options = {}) {
  const diagnostics = [];
  const zip = await readZip(await exportCanopy(input, "xlsx", { ...options, onDiagnostic: item => diagnostics.push(item.code) }));
  return { sheet: zip.get("xl/worksheets/sheet1.xml").content, diagnostics };
}

// Every source of worksheet hyperlinks, with the links it reserves. Captured data-cell links fill only the room left.
const long = "y".repeat(40000), help = target => new ExportCell("Help", { link: { target } });
const linkSources = [
  { name: "no reserved links", rows: 65532, dropped: 2 },
  { name: "configured sheet hyperlinks", rows: 65530, dropped: 1, kept: ["A1"], xlsx: { sheet: { hyperlinks: [{ cell: "A1", target: "https://example.com/help" }] } } },
  { name: "linked header rows", rows: 65530, dropped: 1, kept: ["A1"], xlsx: { sheet: { headerRows: [[{ value: new ExportCell("Name", { link: { target: "https://example.com/header" } }) }]] } } },
  { name: "linked footer values", rows: 65530, dropped: 1, kept: ["A65532"], xlsx: { sheet: { footer: { values: [help("https://example.com/footer")] } } } },
  { name: "a preserved cell before the links", rows: 65530, dropped: 1, internal: ["B2"], long: 0, xlsx: { oversizedText: "preserve" } },
  { name: "a preserved cell after the links", rows: 65530, dropped: 1, internal: ["B65531"], long: 65529, xlsx: { oversizedText: "preserve" } }
];

for (const source of linkSources) {
  test("Canopy XLSX link truncation reserves worksheet links for " + source.name, async () => {
    const rows = linkRows(source.rows).map((row, index) => source.long === undefined ? row :
      { ...row, cells: { ...row.cells, Note: index === source.long ? cell(long, long) : cell("", "") } });
    const columns = source.long === undefined ? [column("Name")] : [column("Name"), column("Note")];
    const { sheet, diagnostics } = await linkedSheet(capture(rows, columns), source.xlsx ? { xlsx: source.xlsx } : {});
    const refs = [...sheet.matchAll(/<hyperlink ref="([A-Z]+\d+)"([^>]*)/g)];
    assert.equal(refs.length, 65530);
    for (const reference of source.kept ?? []) assert.ok(refs.some(([, ref, rest]) => ref === reference && / r:id=/.test(rest)), reference);
    for (const reference of source.internal ?? []) assert.ok(refs.some(([, ref, rest]) => ref === reference && / location=/.test(rest)), reference);
    // Captured links keep row order: the first data rows keep theirs and the last ones keep only the label.
    assert.match(sheet, /<hyperlink ref="A2"/);
    assert.doesNotMatch(sheet, new RegExp('<hyperlink ref="A' + (source.rows + 1) + '"'));
    assert.match(sheet, new RegExp('<c r="A' + (source.rows + 1) + '"'));
    assert.deepEqual(diagnostics, ["LINKS_TRUNCATED:" + source.dropped]);
  });
}

test("Canopy XLSX still rejects an unsafe link that the worksheet limit would leave out", async () => {
  const rows = [...linkRows(65530), { id: "unsafe", cells: { Name: cell("x", "x", { link: { target: "javascript:alert(1)" } }) } }];
  await assert.rejects(exportCanopy(capture(rows), "xlsx"), /absolute HTTP, HTTPS or mailto/);
});

const toneCapture = () => capture([
  { id: "ok", cells: { Name: cell("a", "A", { tone: "success" }), State: cell("Fine", "Fine", { tone: "info" }) } },
  { id: "bad", tone: "danger", cells: { Name: cell("b", "B"), State: cell("Late", "Late", { tone: "warning" }) } },
  { id: "quiet", cells: { Name: cell("c", "C", { tone: "muted" }), State: cell("Plain", "Plain", { tone: "neutral" }) } }
], [column("Name"), column("State")]);

test("Canopy standard tones get default XLSX fills and readable font colours, overridable per tone", async () => {
  const diagnostics = [];
  const book = await workbook(await exportCanopy(toneCapture(), "xlsx", { onDiagnostic: item => diagnostics.push(item.code) }));
  assert.deepEqual(diagnostics, []);
  const expect = (reference, fill, font) => { const style = book.style(reference); assert.equal(fillColor(style), fill, reference); assert.equal(fontColor(style), font, reference); };
  expect("A2", "E2F0D9", "375623"); expect("B2", "DDEBF7", "1F4E78");
  // Row tone paints every cell; a cell tone wins over it.
  expect("A3", "FFC7CE", "9C0006"); expect("B3", "FFF2CC", "7F5000");
  expect("A4", undefined, "595959"); expect("B4", undefined, undefined);
  const custom = await workbook(await exportCanopy(toneCapture(), "xlsx", { tones: { success: { background: "00B050", color: "FFFFFF" } } }));
  assert.equal(fillColor(custom.style("A2")), "00B050"); assert.equal(fontColor(custom.style("A2")), "FFFFFF");
  assert.equal(fillColor(custom.style("B2")), "DDEBF7");
});

test("Canopy standard tones paint default PDF cell backgrounds", async () => {
  const content = (await inspectPdf(await exportCanopy(toneCapture(), "pdf"))).pages.map(page => page.content).join("\n");
  for (const hex of ["E2F0D9", "DDEBF7", "FFC7CE", "FFF2CC"]) assert.ok(content.includes(pdfFill(hex) + " rg"), hex);
});

test("Canopy XLSX header is shaded, bold, wrapped and ruled like the PDF heading", async () => {
  const book = await workbook(await exportCanopy(toneCapture(), "xlsx"));
  for (const reference of ["A1", "B1"]) {
    const header = book.style(reference);
    assert.equal(fillColor(header), "E7EDF5"); assert.match(header.font, /<b\/>/);
    assert.match(header.border, /<bottom style="thin"><color rgb="FFA6B4C8"\/><\/bottom>/);
    assert.match(header.alignment, /wrapText="1"/);
  }
  // Autofilter and the frozen header row stay in place.
  assert.match(book.sheet, /<autoFilter ref="A1:B4"\/>/); assert.match(book.sheet, /<pane [^>]*ySplit="1"[^>]*state="frozen"/);
  const plain = await workbook(await exportCanopy(toneCapture(), "xlsx", { xlsx: { sheet: { headerFill: "FFFFFF", headerWrap: false } } }));
  assert.equal(fillColor(plain.style("A1")), "FFFFFF"); assert.doesNotMatch(plain.style("A1").alignment, /wrapText/);
});

test("Canopy cell composition picks value or label per column kind, rule and writer", async () => {
  const columns = [column("Plain"), column("Badged"), column("Ruled"), column("Link"), column("LinkRuled"), column("Count", "number")];
  const row = { id: "r", cells: {
    Plain: cell("pending", "Awaiting approval"),
    Badged: cell("pending", "Awaiting approval", { badges: ["Active"] }),
    Ruled: cell("pending", "Awaiting approval", { badges: ["Active"] }),
    Link: cell("/users/1", "Alice", { diagnostics: ["relative-link"] }),
    LinkRuled: cell("/users/2", "Bob", { diagnostics: ["relative-link"] }),
    Count: cell(12, "12", { badges: ["Rare"], tone: "danger" }) } };
  const label = "Awaiting approval · Active";
  // Expected written cell per writer: XLSX and CSV write the value, PDF prints the text.
  const cases = [
    { values: "raw", xlsx: ["pending", label, "pending", "Alice", "/users/2", "12"], csv: ["pending", label, label, "Alice", "Bob", "12"] },
    { values: "display", xlsx: ["Awaiting approval", label, label, "Alice", "Bob", "12 · Rare"], csv: ["Awaiting approval", label, label, "Alice", "Bob", "12 · Rare"] }
  ];
  const pdf = ["Awaiting approval", label, label, "Alice", "Bob", "12 · Rare"];
  for (const { values, xlsx, csv } of cases) {
    const input = () => {
      const result = capture([row], columns);
      result.request.values = values;
      // Only XLSX emits these rules, and only for raw values; they protect Ruled and LinkRuled.
      result.request.formatRules = ["Ruled", "LinkRuled"].map(id => ({ id: "rule-" + id, column: id, target: "cell",
        when: { operator: "and", filters: [{ column: id, operator: "isNotNull", values: [] }], groups: [] }, style: { tone: "info" } }));
      return result;
    };
    const options = { unsupportedPresentation: "text" };
    const book = await workbook(await exportCanopy(input(), "xlsx", options));
    assert.deepEqual(columns.map((_, i) => book.style(String.fromCharCode(65 + i) + "2").text), xlsx, "xlsx " + values);
    assert.equal(fillColor(book.style("F2")), "FFC7CE", "tone fill " + values);
    if (values === "raw") assert.match(book.sheet, /<c r="F2"[^>]*><v>12<\/v><\/c>/);
    const text = await (await exportCanopy(input(), "csv", options)).text();
    assert.deepEqual(text.split("\r\n")[1].split(","), csv, "csv " + values);
    const lines = (await inspectPdf(await exportCanopy(input(), "pdf", options))).pages.flatMap(page => page.lines);
    for (const expected of pdf) assert.ok(lines.includes(expected), "pdf " + values + ": " + expected + " in " + lines.join("|"));
  }
});

test("Canopy badges leave native rule operands and typed group identities unchanged", async () => {
  const rule = { id: "locked", column: "Name", target: "cell", when: { operator: "and", filters: [{ column: "Name", operator: "isNull", values: [] }], groups: [] }, style: { tone: "danger" } };
  const ruled = capture([{ id: "r", cells: { Name: cell(null, "", { badges: ["Locked"], diagnostics: ["badges"] }), Note: cell("n", "n", { badges: ["Seen"] }) } }],
    [column("Name"), column("Note")]);
  ruled.request.formatRules = [rule];
  const book = await workbook(await exportCanopy(ruled, "xlsx"));
  // The isNull rule still sees a blank operand; columns no rule reads still carry their labels.
  assert.match(book.sheet, /ISBLANK\(/); assert.equal(book.style("A2").text, ""); assert.equal(book.style("B2").text, "n · Seen");

  const when = ["2026-01-02T03:04:05.000Z", "2026-02-03T04:05:06.000Z"];
  const grouped = capture(when.map((value, index) => ({ id: "r" + index, cells: { When: cell(value, value, { badges: ["Late"], diagnostics: ["badges"] }), Name: cell("n" + index) } })),
    [column("When", "datetime"), column("Name")]);
  grouped.request.grouping = { levels: [{ column: "When", direction: "asc" }], aggregates: [], totals: {},
    groups: when.map((value, index) => ({ key: "g" + index, value, offset: index, count: 1, aggregates: {} })) };
  const blob = await exportCanopy(grouped, "xlsx", { datetime: "typed" });
  assert.ok((await readZip(blob)).get("xl/worksheets/sheet1.xml").content.includes("n1"));
});

test("Canopy PDF numeric formatting keeps appended badge labels", async () => {
  const amount = { ...column("Amount", "number"), format: { minimumFractionDigits: 0, maximumFractionDigits: 0, suffix: " USD" } };
  const input = capture([
    { id: "a", cells: { Amount: cell(12, "12 USD", { badges: ["Rare"], diagnostics: ["badges"] }) } },
    { id: "b", cells: { Amount: cell(7, "7 USD") } },
    { id: "c", cells: { Amount: cell(12, "", { badges: ["Level 2"] }) } },
    { id: "d", cells: { Amount: cell(12, "", { badges: ["2"] }) } }
  ], [amount]);
  const lines = (await inspectPdf(await exportCanopy(input, "pdf"))).pages.flatMap(page => page.lines);
  assert.ok(lines.includes("12 USD · Rare"), lines.join("|")); assert.ok(lines.includes("7 USD"));
  // A badge-only label is not a number to re-render: digits inside the badge stay as written.
  assert.ok(lines.includes("Level 2") && lines.includes("2"), lines.join("|"));
  assert.ok(!lines.includes("Level 12 USD") && !lines.includes("12 USD"), lines.join("|"));
});

test("Canopy relative-link cells read by an emitted native rule keep their raw scalar", async () => {
  const input = capture([{ id: "r", cells: { Flag: cell(true, "Yes", { diagnostics: ["relative-link"] }), Note: cell("/x", "Open", { diagnostics: ["relative-link"] }) } }],
    [column("Flag"), column("Note")]);
  input.request.formatRules = [{ id: "on", column: "Flag", target: "cell", when: { operator: "and", filters: [{ column: "Flag", operator: "eq", values: [true] }], groups: [] }, style: { tone: "success" } }];
  const book = await workbook(await exportCanopy(input, "xlsx", { unsupportedPresentation: "text" }));
  // The rule's ISLOGICAL test needs the boolean; columns no rule reads still show the label.
  assert.match(book.sheet, /ISLOGICAL\(/);
  assert.match(book.sheet, /<c r="A2"[^>]* t="b"[^>]*><v>1<\/v><\/c>/);
  assert.equal(book.style("B2").text, "Open");
});

test("Canopy badges stay on text columns read only by rejected native rules", async () => {
  const input = capture([{ id: "r", cells: { Status: cell("Ready", "Ready", { badges: ["Active"], diagnostics: ["badges"] }) } }], [column("Status")]);
  // Native text comparisons are not emitted, so the cell keeps its badge label.
  input.request.formatRules = [{ id: "ready", column: "Status", target: "cell", when: { operator: "and", filters: [{ column: "Status", operator: "eq", values: ["Ready"] }], groups: [] }, style: { tone: "success" } }];
  const diagnostics = [];
  const book = await workbook(await exportCanopy(input, "xlsx", { unsupportedPresentation: "text", onDiagnostic: item => diagnostics.push(item.code) }));
  assert.ok(diagnostics.includes("UNQUALIFIED_NATIVE_TEXT_COMPARISON:ready"), diagnostics.join("|"));
  assert.doesNotMatch(book.sheet, /conditionalFormatting/);
  assert.equal(book.style("A2").text, "Ready · Active");
});
