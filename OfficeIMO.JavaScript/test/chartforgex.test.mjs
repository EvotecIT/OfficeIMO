import test from "node:test";
import assert from "node:assert/strict";
import { chartForgeXDrawing, exportChartForgeX } from "../dist/integrations/chartforgex/index.js";
import { readZip } from "./zip-reader.mjs";
import { inspectPdf } from "./pdf-reader.mjs";

const capture = (series, x = { kind: "date", title: "Week" }, extra = {}) => ({ version: 1, title: "Sign-ins", description: "Weekly sign-ins by method.", x,
  y: { title: "Sign-ins", unit: "count", format: "#,##0" }, series, ...extra });
const weekly = [
  { id: "passkey", name: "Passkey", points: [{ x: "2026-01-12", xCaption: "W2", y: 140, caption: "140" }, { x: "2026-01-05T00:00:00Z", xCaption: "W1", y: 120, caption: "120" }] },
  { id: "password", name: "Password", format: "0.0", points: [{ x: "2026-01-05", y: 300.5 }, { x: "2026-01-19", y: null }] }
];
async function sheet(blob) {
  const zip = await readZip(blob), xml = zip.get("xl/worksheets/sheet1.xml").content, styles = zip.get("xl/styles.xml").content;
  const formats = new Map([...styles.matchAll(/<numFmt numFmtId="(\d+)" formatCode="([^"]*)"/g)].map(m => [m[1], m[2].replace(/&quot;/g, "\"")]));
  const xfs = [.../<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)[1].matchAll(/<xf\b[^>]*numFmtId="(\d+)"/g)].map(m => m[1]);
  const cells = new Map([...xml.matchAll(/<c r="([A-Z]+\d+)"(?: s="(\d+)")?( t="inlineStr")?(?:\/>|>([\s\S]*?)<\/c>)/g)].map(([, ref, style, inline, body]) => [ref, {
    type: inline ? "string" : body ? "number" : "empty", value: inline ? /<t[^>]*>([^<]*)<\/t>/.exec(body)[1] : body ? Number(/<v>([^<]*)<\/v>/.exec(body)[1]) : null,
    format: formats.get(xfs[Number(style ?? 0)]) }]));
  return { zip, xml, cells };
}

test("a ChartForgeX capture becomes a typed wide sheet with captions beside underlying values", async () => {
  const { zip, cells } = await sheet(await exportChartForgeX(capture(weekly), "xlsx"));
  assert.deepEqual(["A1", "A2", "B2", "C2", "D2", "E2"].map(ref => cells.get(ref).value),
    ["Sign-ins", "Week", "Week (display)", "Passkey (count)", "Passkey (count) (display)", "Password (count)"]);
  // Rows follow ascending dates, whichever series supplied them first; dates are serial numbers with a date format.
  assert.deepEqual(["A3", "A4", "A5"].map(ref => [cells.get(ref).type, cells.get(ref).value, cells.get(ref).format]),
    [["number", 46027, "yyyy-mm-dd"], ["number", 46034, "yyyy-mm-dd"], ["number", 46041, "yyyy-mm-dd"]]);
  assert.deepEqual(["B3", "B4", "C3", "D3", "C4", "D4"].map(ref => cells.get(ref).value), ["W1", "W2", 120, "120", 140, "140"]);
  assert.equal(cells.get("C3").format, "#,##0");
  assert.deepEqual([cells.get("E3").value, cells.get("E3").format, cells.get("E5").type], [300.5, "0.0", "empty"]);
  assert.match(zip.get("docProps/core.xml").content, /<dc:title>Sign-ins<\/dc:title><dc:description>Weekly sign-ins by method\.<\/dc:description>/);
  assert.match(zip.get("xl/workbook.xml").content, /<sheet name="Chart data"/);
});

test("repeated x values and categories choose the long or first-seen layout and keep series identities", async () => {
  const scatter = [{ id: "a", name: "Hosts", points: [{ x: 1, y: 2 }, { x: 1, y: 3, caption: "3 ms" }] }, { id: "b", name: "Hosts", points: [{ x: 0.5, y: 4 }] }];
  const long = await sheet(await exportChartForgeX(capture(scatter, { kind: "number", format: "0.00" }), "xlsx"));
  assert.deepEqual(["A2", "B2", "C2", "D2", "E2"].map(ref => long.cells.get(ref).value), ["Series", "Series ID", "X", "Sign-ins (count)", "Sign-ins (count) (display)"]);
  assert.deepEqual([3, 4, 5].map(row => ["A", "B", "C", "D", "E"].map(column => long.cells.get(column + row).value)),
    [["Hosts", "a", 1, 2, null], ["Hosts", "a", 1, 3, "3 ms"], ["Hosts", "b", 0.5, 4, null]]);
  assert.equal(long.cells.get("C5").format, "0.00");
  await assert.rejects(exportChartForgeX(capture(scatter, { kind: "number" }), "xlsx", { layout: "wide" }), /unique x values/);

  const sites = [{ id: "a", name: "Sites", points: [{ x: "Warsaw", y: 3 }, { x: "Berlin", y: 1 }] }, { id: "b", name: "Sites", points: [{ x: "Austin", y: 2 }, { x: "Berlin", y: 5 }] }];
  const wide = await sheet(await exportChartForgeX(capture(sites, { kind: "category", title: "Site" }), "xlsx"));
  assert.deepEqual(["A2", "B2", "C2"].map(ref => wide.cells.get(ref).value), ["Site", "Sites [a] (count)", "Sites [b] (count)"]);
  assert.deepEqual([3, 4, 5].map(row => ["A", "B", "C"].map(column => wide.cells.get(column + row).value)),
    [["Warsaw", 3, null], ["Berlin", 1, 5], ["Austin", null, 2]]);
  assert.equal(wide.cells.get("A3").type, "string");
});

test("capture values are validated once before writing", async () => {
  for (const [input, error] of [
    [{ ...capture(weekly), version: 2 }, /version 1/],
    [capture(weekly, { kind: "time" }), /x\.kind/],
    [capture([{ id: "a", name: "A", points: [{ x: "05/01/2026", y: 1 }] }]), /ISO 8601/],
    [capture([{ id: "a", name: "A", points: [{ x: 1, y: Infinity }] }], { kind: "number" }), /finite number or null/],
    [capture([{ name: "A", points: [] }]), /id is required/],
    [capture(weekly, undefined, { svg: 1 }), /svg must be markup/]
  ]) await assert.rejects(exportChartForgeX(input, "xlsx"), error);
  await assert.rejects(exportChartForgeX(capture(weekly), "pdf"), /needs the chart's SVG/);
  await assert.rejects(exportChartForgeX(capture(weekly), "csv"), /xlsx or pdf/);
  // Times without an offset are UTC wall-clock values, not the reader's local time.
  const { cells } = await sheet(await exportChartForgeX(capture([{ id: "a", name: "A", points: [{ x: "2026-01-05T06:30", y: 1 }] }]), "xlsx"));
  assert.deepEqual([cells.get("A3").value, cells.get("A3").format], [46027.270833333336, "yyyy-mm-dd hh:mm"]);
});

test("long sheets preserve repeated missing categories, per-series units and conflicting x captions", async () => {
  const series = [
    { id: "a", name: "First", unit: "s", points: [{ x: null, y: 2 }, { x: null, y: 3 }, { x: "A", xCaption: "Alpha", y: 4 }] },
    { id: "b", name: "Second", points: [{ x: "A", xCaption: "Region A", y: 500 }] }
  ];
  const input = capture(series, { kind: "category" }, { y: { title: "Duration", unit: "ms" } });
  const { cells } = await sheet(await exportChartForgeX(input, "xlsx"));
  const row = index => ["B", "C", "D", "E", "F"].map(column => cells.get(column + index).value);
  assert.deepEqual([2, 3, 4, 5, 6].map(row), [
    ["Series ID", "X", "X (display)", "Duration", "Unit"],
    ["a", null, null, 2, "s"], ["a", null, null, 3, "s"], ["a", "A", "Alpha", 4, "s"], ["b", "A", "Region A", 500, "ms"]
  ]);
  await assert.rejects(exportChartForgeX(input, "xlsx", { layout: "wide" }), /unique x values/);
  const conflictOnly = { ...input, series: [{ ...series[0], points: series[0].points.slice(2) }, series[1]] };
  await assert.rejects(exportChartForgeX(conflictOnly, "xlsx", { layout: "wide" }), /consistent x captions/);
  const singleUnit = await sheet(await exportChartForgeX({ ...input, series: [series[0]] }, "xlsx"));
  assert.equal(singleUnit.cells.get("E2").value, "Duration (s)");
  assert.equal(singleUnit.cells.has("F2"), false);
});

test("XLSX row limits reject a sparse wide projection before writing", async () => {
  const series = Array.from({ length: 100 }, (_, index) =>
    ({ id: String(index), name: "Series " + index, points: [{ x: index, y: index }] }));
  const limits = { maxRows: 2 };
  const input = capture(series, { kind: "number" });
  await assert.rejects(exportChartForgeX(input, "xlsx", { limits }), { code: "RESOURCE_LIMIT", message: /maxRows exceeded \(3 > 2\)/ });
  await assert.rejects(exportChartForgeX(input, "xlsx", { layout: "long", limits }), { code: "RESOURCE_LIMIT", message: /maxRows exceeded \(100 > 2\)/ });
});

const chartSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="400" height="200" viewBox="0 0 800 400">
  <defs><clipPath id="plot" clipPathUnits="userSpaceOnUse"><rect x="40" y="20" width="720" height="340"/></clipPath>
    <linearGradient id="fade"><stop offset="0" style="stop-color: rgb(42, 120, 214); stop-opacity: 1"/></linearGradient></defs>
  <style>.mark { fill: red }</style>
  <rect width="800" height="400" style="fill: rgb(255, 255, 255)"/>
  <g clip-path="url(#plot)" transform="translate(0 0)">
    <path d="M40 300 L200 200 C300 100 400 100 500 150 Q600 200 700 120 A20 20 0 0 1 740 120 Z" fill="url(#fade)" stroke="#16181C" stroke-width="2"/>
    <line x1="40" y1="360" x2="760" y2="360" style="stroke: oklch(0.7 0.02 250); stroke-dasharray: 4 2; opacity: 0.5"/>
    <circle class="mark" cx="200" cy="200" r="4" fill="#2A78D6"/><circle cx="500" cy="150" r="4" fill="#2A78D6"/>
  </g>
  <text x="400" y="390" text-anchor="middle" font-size="24" fill="#4D525B">Week −1 → 2</text>
  <text transform="translate(20 200) rotate(-90)" text-anchor="middle" style="font-size: 20px; font-weight: 600; fill: color(srgb 0.1 0.1 0.1)">Sign-ins</text>
  <text x="60" y="40" font-size="20"><tspan x="60">First</tspan><tspan x="60" dy="1.2em">Second</tspan></text>
  <image href="data:image/png;base64,AA==" width="10" height="10"/>
</svg>`;

test("chart SVG becomes a vector drawing in points with clips, curves, turned text and diagnostics", () => {
  const diagnostics = [];
  const drawing = chartForgeXDrawing(chartSvg, { onDiagnostic: diagnostic => diagnostics.push(diagnostic.code) });
  assert.deepEqual([drawing.version, drawing.width, drawing.height, drawing.standardFonts], [1, 300, 150, true]);
  const [background, plot, label, axis, first, second] = drawing.elements;
  assert.deepEqual([background.kind, background.fill], ["path", "FFFFFF"]);
  assert.deepEqual(plot.clip, { x: 15, y: 7.5, width: 270, height: 127.5 });
  const [area, line, marks] = plot.elements;
  assert.deepEqual([area.fill, area.stroke, area.strokeWidth], ["2A78D6", "16181C", .75]);
  assert.deepEqual(area.commands.map(c => c.op), ["move", "line", "cubic", "cubic", "cubic", "cubic", "close"]);
  assert.deepEqual([area.commands.at(-2).x, area.commands.at(-2).y], [277.5, 45], "The half-circle arc ends at its endpoint");
  assert.deepEqual([line.dash, line.strokeWidth], [[1.5, .75], .375]);
  assert.match(line.stroke, /^[0-9A-F]{6}$/);
  assert.equal(marks.commands.filter(c => c.op === "move").length, 2, "Same-paint circles share one path element");
  assert.deepEqual(label, { kind: "text", text: "Week -1 ? 2", x: 150, y: 146.25, font: "regular", fontSize: 9, color: "4D525B", anchor: "middle" });
  assert.deepEqual([axis.text, axis.font, axis.rotation, axis.anchor, axis.x, axis.y, axis.color], ["Sign-ins", "bold", -90, "middle", 7.5, 75, "1A1A1A"]);
  assert.deepEqual([first.text, first.y, second.text, second.y], ["First", 15, "Second", 24]);
  assert.deepEqual(diagnostics, ["STYLESHEET_IGNORED", "GRADIENT_FLATTENED", "OPACITY_FLATTENED", "TEXT_SUBSTITUTED", "IMAGE_SKIPPED"]);
  assert.throws(() => chartForgeXDrawing(chartSvg, { unsupported: "reject" }), /STYLESHEET_IGNORED/);
});

test("path data, transforms and markup edge cases parse like SVG renderers", () => {
  const draw = body => chartForgeXDrawing(`<svg width="100" height="100">${body}</svg>`, { onDiagnostic: () => {} }).elements;
  // Compact arc flags: "a10 10 0 011 1" is large=0, sweep=1, x=1, y=1.
  const [arc] = draw(`<path d="M0 0a10 10 0 011 1" stroke="black"/>`);
  assert.equal(arc.commands.at(-1).op, "cubic");
  assert.deepEqual([arc.commands.at(-1).x, arc.commands.at(-1).y], [.75, .75]);
  const [truncated] = draw(`<path d="M10 10 L20 20 Z l5 0 D0 0" stroke="black"/>`);
  assert.deepEqual(truncated.commands.slice(-2), [{ op: "move", x: 7.5, y: 7.5 }, { op: "line", x: 11.25, y: 7.5 }]);
  const [scaled] = draw(`<g transform="matrix(2 0 0 2 10 0) scale(.5)"><rect x="0" y="0" width="10" height="10" rx="2" fill="#000"/></g>`);
  assert.deepEqual(scaled.commands[0], { op: "move", x: 9, y: 0 });
  assert.equal(scaled.commands.filter(c => c.op === "cubic").length, 4);
  assert.equal(draw(`<g style="display: none"><rect width="10" height="10"/></g><rect width="10" height="10" fill="none"/>`).length, 0);
  assert.deepEqual(draw(`<!-- note --><text x="1" y="2" fill="currentColor" style="color: rgba(0,0,0,.5)">a &amp; b&#x21;</text>`)[0].color, "808080");
  assert.throws(() => chartForgeXDrawing(`<svg width="10" height="10"><g></svg>`), /does not match/);
  assert.throws(() => chartForgeXDrawing(`<html/>`), /<svg> root/);
  assert.throws(() => chartForgeXDrawing(`<!DOCTYPE svg [<!ENTITY x "y">]><svg width="1" height="1"/>`), /internal DTD/);
});

test("SVG conversion preserves symbol viewports, resolved paint and painter order", () => {
  const diagnostics = [];
  const result = chartForgeXDrawing(`<?export x > y?><svg width="100" height="100">
    <defs><symbol id="mark" viewBox="0 0 10 10"><rect width="100%" height="100%"/></symbol></defs>
    <use href="#mark" width="20" height="20" transform="translate(10)" style="transform: translate(20px)"/>
    <rect width="25" height="25" x="5" fill="white" stroke="black"/>
    <rect width="25" height="25" x="15" fill="white" stroke="black"/>
    <rect width="10" height="10" opacity="50%" style="fill: red !important; fill: green"/>
    <rect width="10" height="10" fill="rgb(255 0 0 / none)"/>
    <switch><text y="40">A<tspan visibility="hidden">Hidden</tspan><tspan>B</tspan></text><text y="40">Extra</text></switch>
    <text x="0" y="70">A<tspan x="20" font-weight="bold">B</tspan>C</text>
  </svg>`, { onDiagnostic: d => diagnostics.push(d.code) });
  assert.equal(result.elements.length, 7);
  assert.deepEqual(result.elements[0].commands[2], { op: "line", x: 30, y: 15 });
  assert.deepEqual(result.elements.slice(1, 3).map(path => [path.commands[0].x, path.fill, path.stroke]),
    [[3.75, "FFFFFF", "000000"], [11.25, "FFFFFF", "000000"]]);
  assert.equal(result.elements[3].fill, "FF8080");
  assert.equal(result.elements[4].text, "AB");
  assert.deepEqual(diagnostics, ["OPACITY_FLATTENED", "TEXT_STYLE_MERGED"]);
});

test("SVG percentages use the active viewBox for geometry, clips and labels", () => {
  const drawing = chartForgeXDrawing(`<svg width="200" height="100" viewBox="0 0 100 50">
    <defs><clipPath id="half"><rect width="50%" height="50%"/></clipPath></defs>
    <g clip-path="url(#half)"><rect width="50%" height="50%"/></g>
    <circle cx="25%" cy="50%" r="10%" fill="blue"/>
    <svg x="50%" y="50%" width="50%" height="50%" viewBox="0 0 40 20">
      <text x="50%" y="50%">Label</text>
    </svg>
  </svg>`);
  const [group, circle, label] = drawing.elements;
  assert.deepEqual(group.clip, { x: 0, y: 0, width: 75, height: 37.5 });
  assert.deepEqual(group.elements[0].commands[2], { op: "line", x: 75, y: 37.5 });
  assert.ok(Math.abs(circle.commands[0].x - (37.5 + Math.hypot(100, 50) / Math.SQRT2 * .15)) < 1e-9);
  assert.deepEqual([label.x, label.y], [112.5, 56.25]);
});

test("SVG conversion applies drawing ceilings during reference expansion", async () => {
  const marks = `<svg width="40" height="40">
    <defs><circle id="mark" r="2"/></defs><use href="#mark"/><use href="#mark" x="10"/>
  </svg>`;
  const input = capture(weekly, undefined, { svg: marks });
  await assert.rejects(exportChartForgeX(input, "pdf", { pdf: { limits: { maxVectorCommands: 6 } } }),
    { code: "RESOURCE_LIMIT", message: /maxVectorCommands/ });
  assert.equal(chartForgeXDrawing(marks, { limits: { maxDrawingElements: 1, maxVectorCommands: 12 } }).elements.length, 1);

  const refs = Array.from({ length: 12 }, (_, i) => `<g id="r${i + 1}"><use href="#r${i}"/><use href="#r${i}"/></g>`).join("");
  const emptyExpansion = `<svg width="40" height="40"><defs><g id="r0"/>${refs}</defs><use href="#r12"/></svg>`;
  assert.throws(() => chartForgeXDrawing(emptyExpansion, { limits: { maxDrawingElements: 1, maxVectorCommands: 12 } }),
    { code: "RESOURCE_LIMIT", message: /conversion work limit/ });
});

test("a ChartForgeX PDF carries the title, description and the chart as vector paths and text", async () => {
  const diagnostics = [];
  const pdf = await inspectPdf(await exportChartForgeX(capture(weekly, undefined, { svg: chartSvg, title: "Sign-ins → Weekly", description: "Weekly sign-ins − by method." }), "pdf",
    { pdf: { compression: false, tagged: { language: "en-US" } }, onDiagnostic: diagnostic => diagnostics.push(diagnostic.code) }));
  assert.equal(pdf.pages.length, 1);
  const content = pdf.pages[0].content;
  assert.match(pdf.text, /^Sign-ins \? Weekly\nWeekly sign-ins - by method\./);
  assert.match(pdf.text, /Week -1 \? 2/);
  assert.match(content, / c\n/); assert.match(content, / re W n\n/); assert.match(content, /\[1\.5 0\.75\] 0 d\n/);
  assert.match(content, /0 -1 -1 0 [\d.]+ [\d.]+ Tm/, "The y axis title stays vertical");
  assert.ok(!/\/XObject|\/Image/.test([...pdf.objects.values()].map(o => o.body).join("\n")), "The chart is vector, not an embedded image");
  assert.match([...pdf.objects.values()].map(o => o.body).join("\n"), /\/Title <feff005300690067006e002d0069006e0073002021920020005700650065006b006c0079>/);
  assert.match([...pdf.objects.values()].map(o => o.body).join("\n"), /\/Alt <feff/);
  assert.ok(diagnostics.includes("TEXT_SUBSTITUTED") && diagnostics.includes("IMAGE_SKIPPED"));
});
