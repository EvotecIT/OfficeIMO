import test from "node:test";
import assert from "node:assert/strict";
import { writePdfDocument, writePdfDocumentTo, pdfTable } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const columns = [{ header: "Asset", key: "name" }, { header: "Count", key: "count" }];
const options = { pageSize: "A5", margins: 24, pageNumbers: true, compression: false };
function outlines(pdf) {
  const catalog = [...pdf.objects.values()].find(object => /\/Type \/Catalog /.test(object.body));
  const root = Number(/\/Outlines (\d+) 0 R/.exec(catalog.body)?.[1]);
  if (!root) return [];
  const result = [], visited = new Set();
  function siblings(id, parent) {
    let previous;
    while (id) {
      assert.ok(!visited.has(id)); visited.add(id);
      const body = pdf.objects.get(id).body;
      assert.match(body, new RegExp("/Parent " + parent + " 0 R"));
      if (previous) assert.match(body, new RegExp("/Prev " + previous + " 0 R"));
      const title = /\/Title <feff([0-9a-f]+)>/.exec(body)[1];
      const text = [...title.matchAll(/.{4}/g)].map(match => String.fromCharCode(parseInt(match[0], 16))).join("");
      const destination = /\/Dest \[(\d+) 0 R \/XYZ null ([\d.]+) null\]/.exec(body);
      const page = pdf.pages.findIndex(page => page.id === Number(destination[1])); assert.ok(page >= 0);
      result.push({ id, title: text, page, y: Number(destination[2]), body });
      const first = Number(/\/First (\d+) 0 R/.exec(body)?.[1]); if (first) siblings(first, id);
      previous = id; id = Number(/\/Next (\d+) 0 R/.exec(body)?.[1]);
    }
  }
  siblings(Number(/\/First (\d+) 0 R/.exec(pdf.objects.get(root).body)?.[1]), root);
  return result;
}

test("sections stream portrait, landscape and inherited portrait pages with continuous numbering and native bookmarks", async () => {
  const context = [], blocks = [
    { kind: "section", title: "Overview" }, { kind: "paragraph", text: "Summary body" },
    { kind: "section", title: "Inventory", level: 2, orientation: "landscape", pageHeader: "Wide inventory" },
    pdfTable([{ name: "NODE-001", count: 7 }], { columns }),
    { kind: "section", title: "Findings" }, { kind: "paragraph", text: "Follow-up body" }
  ];
  const pdf = await inspectPdf(await writePdfDocument(blocks, { ...options, pageHeader(value) { context.push(value); return "Global header"; } }));
  assert.equal(pdf.pages.length, 3);
  assert.match(pdf.pages[0].body, /\/MediaBox \[0 0 419\.528 595\.276\]/);
  assert.match(pdf.pages[1].body, /\/MediaBox \[0 0 595\.276 419\.528\]/);
  assert.match(pdf.pages[2].body, /\/MediaBox \[0 0 419\.528 595\.276\]/);
  assert.deepEqual(context.map(({ pageNumber, sectionNumber, sectionPageNumber, sectionTitle }) => [pageNumber, sectionNumber, sectionPageNumber, sectionTitle]), [[1,1,1,"Overview"],[3,3,1,"Findings"]]);
  assert.deepEqual(outlines(pdf).map(item => [item.title, item.page]), [["Overview",0],["Inventory",1],["Findings",2]]);
  assert.ok(outlines(pdf)[1].body.includes("/Parent " + outlines(pdf)[0].id + " 0 R"));
  for (let index = 0; index < 3; index++) assert.ok(pdf.pages[index].lines.includes("Page " + (index + 1) + " of "));
});

test("first landscape, empty paragraphs and consecutive unpainted sections never manufacture transition pages", async () => {
  for (const prefix of [[], [{ kind: "paragraph", text: "" }], [{ kind: "group", blocks: [] }]]) {
    const pdf = await inspectPdf(await writePdfDocument([...prefix,
      { kind: "section", title: "Empty section", includeHeading: false },
      { kind: "section", title: "Wide section", orientation: "landscape" },
      { kind: "paragraph", text: "Only body" }
    ], options));
    assert.equal(pdf.pages.length, 1); assert.match(pdf.pages[0].body, /\/MediaBox \[0 0 595\.276 419\.528\]/);
    assert.deepEqual(outlines(pdf).map(item => item.page), [0,0]);
  }
});

test("bookmarks follow actual kept headings and tagged sections retain one shared Document", async () => {
  const pdf = await inspectPdf(await writePdfDocument([
    { kind: "section", title: "Intro" }, { kind: "paragraph", text: "Line\n".repeat(6) },
    { kind: "paragraph", text: "Kept destination", headingLevel: 2, keepWithNext: true, bookmark: { title: "Kept destination", level: 2 } },
    pdfTable([{ name: "Next table", count: 1 }], { columns }),
    { kind: "section", title: "Next" }, { kind: "paragraph", text: "Done" }
  ], { pageSize: { width: 240, height: 150 }, margins: 20, pageNumbers: false, compression: false, tagged: { language: "en" } }));
  const destination = outlines(pdf).find(item => item.title === "Kept destination");
  assert.ok(pdf.pages[destination.page].lines.includes("Kept destination"));
  assert.ok(pdf.pages[destination.page].lines.includes("Next table"));
  const structures = [...pdf.objects.entries()].filter(([, object]) => /\/Type \/StructElem/.test(object.body));
  const documents = structures.filter(([, object]) => /\/S \/Document /.test(object.body)); assert.equal(documents.length, 1);
  const sections = structures.filter(([, object]) => /\/S \/Sect /.test(object.body)); assert.equal(sections.length, 2);
  for (const [, section] of sections) assert.match(section.body, new RegExp("/P " + documents[0][0] + " 0 R"));
});

test("section and bookmark limits remain cumulative and cancellation releases active sources and the destination", async () => {
  await assert.rejects(writePdfDocument([{ kind: "section", title: "One" }, { kind: "section", title: "Two" }], { ...options, limits: { maxSections: 1 } }), /maxSections/);
  await assert.rejects(writePdfDocument([{ kind: "section", title: "One", includeHeading: false }], { ...options, limits: { maxBookmarks: 0 } }), /maxBookmarks/);
  await assert.rejects(writePdfDocument([{ kind: "section", title: "One" }, { kind: "section", title: "Two" }], { ...options, limits: { maxPages: 1 } }), /maxPages/);
  const abort = new AbortController(), reason = Error("Section export stopped"); let returned = 0;
  async function* blocks() { try { yield { kind: "section", title: "Inventory", orientation: "landscape" }; for (let i = 0; i < 2000; i++) yield { kind: "paragraph", text: "Record " + i }; } finally { returned++; } }
  const destination = new WritableStream({ write() { abort.abort(reason); } });
  await assert.rejects(writePdfDocumentTo(blocks(), destination, { ...options, signal: abort.signal }), error => error === reason);
  assert.equal(returned, 1); assert.equal(destination.locked, false);
});


test("sections restore captured custom geometry and margins despite caller mutation", async () => {
  const pageSize = { width: 240, height: 320 }, margins = { left: 20, right: 20, top: 20, bottom: 20 };
  async function* blocks() {
    yield { kind: "section", title: "Wide", orientation: "landscape" };
    yield { kind: "paragraph", text: "First body" };
    pageSize.width = 800; margins.left = 300;
    yield { kind: "section", title: "Restored" };
    yield { kind: "paragraph", text: "Captured default" };
  }
  const pdf = await inspectPdf(await writePdfDocument(blocks(), { pageSize, margins, pageNumbers: false, compression: false }));
  assert.equal(pdf.pages.length, 2);
  assert.match(pdf.pages[0].body, /\/MediaBox \[0 0 320 240\]/);
  assert.match(pdf.pages[1].body, /\/MediaBox \[0 0 240 320\]/);
  assert.match(pdf.pages[1].content, /1 0 0 1 20 /);
});


test("document title and preface survive a first landscape section without a blank portrait page", async () => {
  const pdf = await inspectPdf(await writePdfDocument([
    { kind: "section", title: "Inventory", orientation: "landscape" },
    pdfTable([{ name: "First asset", count: 1 }], { columns })
  ], { ...options, title: "Global report title", messageTop: "Document preface" }));
  assert.equal(pdf.pages.length, 1); assert.match(pdf.pages[0].body, /\/MediaBox \[0 0 595\.276 419\.528\]/);
  for (const text of ["Global report title", "Document preface", "Inventory", "First asset"]) assert.ok(pdf.pages[0].lines.includes(text));
});


test("outline-only section titles consume cumulative text and cell budgets before being retained", async () => {
  const blocks = [{ kind: "section", title: "One", includeHeading: false }, { kind: "section", title: "Two", includeHeading: false }];
  for (const [limits, expected] of [[{ maxCells: 1 }, /maxCells/], [{ maxTextCharacters: 5 }, /maxTextCharacters/]])
    await assert.rejects(writePdfDocument(blocks, { ...options, pageNumbers: false, limits }), expected);
  const pdf = await inspectPdf(await writePdfDocument(blocks, { ...options, pageNumbers: false,
    limits: { maxCells: 2, maxTextCharacters: 6 } }));
  assert.equal(pdf.pages.length, 1); assert.deepEqual(outlines(pdf).map(item => item.title), ["One", "Two"]);
});
