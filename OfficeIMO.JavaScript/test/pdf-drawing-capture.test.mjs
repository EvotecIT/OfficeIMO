import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { PdfFont, writePdfDocument } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const regular = new PdfFont(new Uint8Array(await readFile(new URL("../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf", import.meta.url))));
const options = { fonts: { regular }, compression: false, pageNumbers: false };
const path = commands => ({ kind: "path", stroke: "008000", commands });
const text = value => ({ kind: "text", text: value, x: 10, y: 25, font: "regular", fontSize: 14, color: "000000" });
const block = elements => ({ kind: "drawing", drawing: { version: 1, width: 200, height: 100, elements } });
const longPath = () => path([{ op: "move", x: 0, y: 0 }, ...Array.from({ length: 1100 }, () => ({ op: "line", x: 1, y: 1 })), { op: "line", x: 123, y: 45 }]);

async function mutateAtYield(drawingBlock, mutate) {
  const now = Object.getOwnPropertyDescriptor(performance, "now"), scheduler = globalThis.scheduler;
  let clock = 0, mutations = 0;
  try {
    Object.defineProperty(performance, "now", { configurable: true, value: () => clock += 40 });
    globalThis.scheduler = { postTask: async callback => { if (!mutations++) mutate(); callback(); } };
    const pdf = await inspectPdf(await writePdfDocument([drawingBlock], options));
    assert.ok(mutations > 0, "The mutation ran during cooperative drawing work");
    return pdf;
  } finally { if (now) Object.defineProperty(performance, "now", now); else delete performance.now; globalThis.scheduler = scheduler; }
}

test("drawing capture owns later command coordinates before cooperative path work", async () => {
  const first = longPath(), final = first.commands.at(-1);
  const pdf = await mutateAtYield(block([first]), () => { final.x = 901; final.y = 902; });
  assert.equal(/123 45 l\nS/.test(pdf.pages[0].content), true, "The captured final command reaches the PDF");
  assert.equal(pdf.pages[0].content.includes("901 902 l"), false);
});

for (const phase of ["path", "font"]) {
  test("drawing capture owns later elements, nested clips and text before " + phase + " work", async () => {
    const leading = phase === "path" ? longPath() : text("A".repeat(2100));
    const final = { op: "cubic", x1: 20, y1: 21, x2: 22, y2: 23, x: 24, y: 25 };
    const laterPath = { ...path([{ op: "move", x: 11, y: 12 }, final]), fill: "FF0000", strokeWidth: 3, lineCap: "round", lineJoin: "bevel", fillRule: "evenodd" };
    const laterText = text("Owned");
    const clip = { x: 3, y: 4, width: 120, height: 50 }, children = [laterPath, laterText];
    const group = { kind: "group", clip, elements: [{ kind: "group", elements: children }] };
    const replacedGroup = { kind: "group", elements: [text("Sibling")] };
    const drawingBlock = block([leading, group, replacedGroup]);
    const pdf = await mutateAtYield(drawingBlock, () => {
      Object.assign(final, { x1: 901, y1: 902, x2: 903, y2: 904, x: 905, y: 906 });
      Object.assign(laterPath, { fill: "0000FF", stroke: "0000FF", strokeWidth: 20, lineCap: "square", lineJoin: "round", fillRule: "nonzero" });
      Object.assign(laterText, { text: "Changed", x: 90, y: 91, font: "bold", fontSize: 30, color: "0000FF" });
      Object.assign(clip, { x: 70, y: 71, width: 72, height: 73 });
      children.push(text("Added")); replacedGroup.elements = []; drawingBlock.drawing.elements.push(text("Appended"));
    });
    assert.equal(pdf.text, (leading.kind === "text" ? leading.text + "\n" : "") + "Owned\nSibling");
    assert.match(pdf.pages[0].content, /3 4 120 50 re W n/);
    assert.match(pdf.pages[0].content, /1 0 0 rg\n0 0\.502 0 RG\n3 w\n1 J\n2 j\n10 M\n11 12 m\n20 21 22 23 24 25 c\nB\*/);
    assert.match(pdf.pages[0].content, /BT \/F1 14 Tf 0 Tr 1 0 0 -1 10 25 Tm/);
  });
}

test("drawing capture finishes before an earlier element's queued microtask", async () => {
  const later = path([{ op: "move", x: 31, y: 32 }, { op: "line", x: 33, y: 34 }]);
  let mutated = false;
  const first = { kind: "path", stroke: "000000", get commands() {
    queueMicrotask(() => { mutated = true; later.commands[1].x = 900; });
    return [{ op: "move", x: 0, y: 0 }];
  } };
  const pdf = await inspectPdf(await writePdfDocument([block([first, later])], options));
  assert.equal(mutated, true);
  assert.match(pdf.pages[0].content, /31 32 m\n33 34 l/);
});

test("drawing capture rejects over-budget arrays before visiting their records", async () => {
  let reads = 0;
  const command = { get op() { reads++; return "move"; }, x: 0, y: 0 };
  await assert.rejects(writePdfDocument([block([path([command, command])])], { ...options, limits: { maxVectorCommands: 1 } }), /maxVectorCommands/);
  const element = { get kind() { reads++; return "path"; }, stroke: "000000", commands: [command] };
  await assert.rejects(writePdfDocument([block([element, element])], { ...options, limits: { maxDrawingElements: 1 } }), /maxDrawingElements/);
  assert.equal(reads, 0);
});
