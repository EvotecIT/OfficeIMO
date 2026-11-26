import test from "node:test";
import assert from "node:assert/strict";
import { exportCanopy } from "../dist/integrations/canopyx/index.js";
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
    { id: "b", cells: { Amount: cell(7, "7 USD") } }
  ], [amount]);
  const lines = (await inspectPdf(await exportCanopy(input, "pdf"))).pages.flatMap(page => page.lines);
  assert.ok(lines.includes("12 USD · Rare"), lines.join("|")); assert.ok(lines.includes("7 USD"));
});

test("Canopy rule operand scan stops at the compiler's predicate budget", async () => {
  let reads = 0;
  const leaf = { operator: "and", get filters() { reads++; return [{ column: "Name", operator: "isNull", values: [] }]; }, groups: [] };
  let node = leaf;
  // A small shared graph whose full expansion is 10^5 nodes; the compiler rejects it after 64.
  for (let level = 0; level < 5; level++) node = { operator: "or", filters: [], groups: Array(10).fill(node) };
  const input = capture([{ id: "r", cells: { Name: cell("A") } }]);
  input.request.formatRules = [{ id: "wide", column: "Name", target: "cell", when: node, style: { tone: "info" } }];
  const diagnostics = [];
  await exportCanopy(input, "xlsx", { unsupportedPresentation: "text", onDiagnostic: item => diagnostics.push(item.code) });
  assert.ok(reads <= 128, "predicate reads: " + reads);
  assert.deepEqual(diagnostics, ["RULE_PREDICATE_LIMIT:wide"]);
});
