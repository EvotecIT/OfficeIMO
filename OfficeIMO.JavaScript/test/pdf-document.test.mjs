import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { writePdfDocument, writePdfDocumentTo, pdfTable } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const columns = [{ header: "Item", key: "name" }, { header: "Amount", key: "amount" }];
const logo = { format: "rgba", width: 2, height: 2, data: Uint8Array.of(255,0,0,255, 0,255,0,128, 0,0,255,255, 255,255,0,0) };
const vector = { kind: "vector", width: 80, height: 30, fill: "E2F0D9", stroke: "008000", commands: [
  { op: "move", x: 1, y: 1 }, { op: "line", x: 79, y: 1 }, { op: "cubic", x1: 79, y1: 15, x2: 50, y2: 29, x: 1, y: 29 }, { op: "close" }
] };
// Synthetic four-color 8x8 JPEG, encoded once by Poppler's JPEG writer.
const jpeg = Uint8Array.from(Buffer.from("/9j/4AAQSkZJRgABAQEACAAIAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAAIAAgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwCv8Lbb/hIv7W+f7P5Hk9t+7dv+mOlFFFfCccwWX5/iMLhfdhHlst94Rb1d3u31M8Rl2GzWo8bjY81SW7u1e2i0TS2S6H//2Q==", "base64"));

test("PDF documents compose graphics and independently paginated tables in one resource tree", async () => {
  const rows = Array.from({ length: 45 }, (_, i) => ({ name: "First row " + i, amount: 2 }));
  const pdf = await inspectPdf(await writePdfDocument([
    { kind: "paragraph", text: "Quarterly report", fontSize: 18, presentation: { bold: true, alignment: "center" } },
    { kind: "image", image: logo, width: 60, height: 60 }, vector,
    pdfTable(rows, { columns, footer: { values: ["First total"], totals: { amount: "sum" } } }),
    { kind: "pageBreak" }, { kind: "paragraph", text: "Second section" },
    { kind: "image", image: logo, width: 20, height: 20 },
    pdfTable([{ name: "Second row", amount: 7 }], { columns: [{ header: "Other", key: "name" }, columns[1]], footer: { values: ["Second total"], totals: { amount: "sum" } } })
  ], { pageSize: "A5", pageHeader: "Shared report", compression: false }));
  for (const text of ["Quarterly report", "First row 44", "First total", "90", "Second section", "Second row", "Second total", "7"]) assert.ok(pdf.text.includes(text), text);
  assert.ok(pdf.pages.length >= 4);
  assert.ok(pdf.pages.slice(1, -1).every(page => page.lines.includes("Item")), "The active table repeats its own heading");
  assert.ok(!pdf.pages.at(-1).lines.includes("Item") && pdf.pages.at(-1).lines.includes("Other"));
  const images = [...pdf.objects.values()].filter(o => /\/Subtype \/Image/.test(o.body));
  assert.equal(images.length, 2, "Repeated images share one RGB object and one alpha mask");
  const rgb = images.find(o => /DeviceRGB/.test(o.body)), alpha = images.find(o => /DeviceGray/.test(o.body));
  assert.deepEqual([...rgb.stream], [255,0,0, 0,255,0, 0,0,255, 255,255,0]);
  assert.deepEqual([...alpha.stream], [255,128,255,0]);
  assert.match(pdf.pages[0].content, /79 15 50 29 1 29 c/);
});

test("PDF documents embed valid JPEG bytes without re-encoding and contain oversized graphics", async () => {
  const pdf = await inspectPdf(await writePdfDocument([{ kind: "image", image: { format: "jpeg", data: jpeg }, width: 1000, height: 1000 }, vector],
    { pageSize: "A5", pageNumbers: false, compression: false }));
  const image = [...pdf.objects.values()].find(o => /\/Filter \/DCTDecode/.test(o.body));
  assert.deepEqual(image.stream, Buffer.from(jpeg));
  assert.match(image.body, /\/Width 8 \/Height 8 \/ColorSpace \/DeviceRGB/);
  assert.equal(pdf.pages.length, 1);
  assert.match(pdf.pages[0].content, /q 347\.528 0 0 347\.528 36 /, "The 1,000-point graphic scales down to the printable A5 width");
  // Synthetic 12×8 quadrant images, Pillow 11.3.0 quality=95/subsampling=0:
  // L=0/85/170/255 (baseline); RGB=240,20,20/20,230,20/20,20,235/235,220,20 (progressive).
  for (const [file, space] of [["gray-baseline.jpg", "DeviceGray"], ["rgb-progressive.jpg", "DeviceRGB"]]) {
    const data = new Uint8Array(await readFile(new URL("./fixtures/pdf/" + file, import.meta.url)));
    const document = await inspectPdf(await writePdfDocument([{ kind: "image", image: { format: "jpeg", data } }], { compression: false }));
    const resource = [...document.objects.values()].find(o => /\/Filter \/DCTDecode/.test(o.body));
    assert.deepEqual(resource.stream, Buffer.from(data));
    assert.match(resource.body, new RegExp("/Width 12 /Height 8 /ColorSpace /" + space));
  }
});

test("PDF image display dimensions fit the page after admitting very wide and explicitly oversized images", async () => {
  const panorama = { format: "rgb", width: 20000, height: 1, data: new Uint8Array(60000) };
  for (const block of [{ kind: "image", image: panorama }, { kind: "image", image: logo, width: 50000, height: 50000 }]) {
    const pdf = await inspectPdf(await writePdfDocument([block], { pageSize: "A5", pageNumbers: false, compression: false }));
    assert.equal(pdf.pages.length, 1);
    const transform = /q ([\d.]+) 0 0 ([\d.]+) 36 [\d.]+ cm \/Image1 Do/.exec(pdf.pages[0].content);
    assert.ok(transform, "A finite fitted image is painted");
    assert.ok(Number(transform[1]) > 0 && Number(transform[1]) <= 347.529);
    assert.ok(Number(transform[2]) > 0 && Number(transform[2]) <= 523.276);
  }
  for (const width of [0, -1, Infinity, NaN])
    await assert.rejects(writePdfDocument([{ kind: "image", image: logo, width }]), /finite and positive/);
});

test("PDF images capture byte, format and display getters before validation and output", async () => {
  let dataReads = 0, formatReads = 0, heightReads = 0;
  const image = { width: 1, height: 1, get format() { return ++formatReads === 1 ? "rgb" : "jpeg"; },
    get data() { return ++dataReads === 1 ? Uint8Array.of(1, 2, 3) : Uint8Array.of(4, 5, 6, 7, 8, 9); } };
  const pdf = await inspectPdf(await writePdfDocument([{ kind: "image", image, get height() { return ++heightReads === 1 ? 17 : 34; } }], {
    compression: false, pageNumbers: false, limits: { maxImageBytes: 3, maxImagePixels: 1 }
  }));
  const embedded = [...pdf.objects.values()].find(o => /\/Subtype \/Image/.test(o.body));
  assert.deepEqual(embedded.stream, Buffer.from([1, 2, 3]));
  assert.match(embedded.body, /\/Width 1 \/Height 1 \/ColorSpace \/DeviceRGB/);
  assert.doesNotMatch(embedded.body, /DCTDecode/);
  assert.match(pdf.pages[0].content, /q 17 0 0 17 /);
  assert.doesNotMatch(pdf.pages[0].content, /q 34 0 0 34 /);
  await assert.rejects(writePdfDocument([{ kind: "image", image: { format: "rgb", width: 1, height: 1, get data() { return new Uint8Array(6); } } }], {
    limits: { maxImageBytes: 3 }
  }), /maxImageBytes/);
});

test("PDF vectors capture geometry, path records and command counts before cooperative serialization", async () => {
  const now = Object.getOwnPropertyDescriptor(performance, "now"), scheduler = globalThis.scheduler;
  const commands = [{ op: "move", x: 0, y: 0 }, ...Array.from({ length: 1024 }, () => ({ op: "line", x: 1, y: 1 }))];
  const mutable = { kind: "vector", width: 10, height: 10, commands, fill: "FFFFFF", stroke: "000000", fillRule: "evenodd" };
  const path = [{ op: "move", x: 0, y: 0 }, { op: "line", x: 2, y: 2 }];
  try {
    let clock = 0, yields = 0;
    Object.defineProperty(performance, "now", { configurable: true, value: () => clock += 40 });
    globalThis.scheduler = { postTask: async callback => {
      if (++yields === 1) {
        commands.push({ op: "line", x: 9, y: 9 }); commands[1024].x = 9;
        mutable.width = 20; mutable.fillRule = "nonzero";
      }
      callback();
    } };
    const pdf = await inspectPdf(await writePdfDocument([mutable, { kind: "vector", width: 10, height: 10, commands: path, stroke: "000000" },
      { kind: "drawing", drawing: { version: 1, width: 10, height: 10, elements: [{ kind: "path", commands: path, stroke: "000000" }] } }],
      { pageNumbers: false, compression: false, limits: { maxVectorCommands: 1029 } }));
    assert.ok(yields > 0);
    const content = pdf.pages[0].content;
    assert.equal((content.match(/0 0 10 10 re W n/g) ?? []).length, 3);
    assert.doesNotMatch(content, /0 0 20 10 re W n|0\.5 0 0 -1/);
    assert.match(content, /B\*/);
    assert.doesNotMatch(content, /9 9 l|9 1 l/);
    assert.equal((content.match(/1 1 l/g) ?? []).length, 1024);
  } finally { if (now) Object.defineProperty(performance, "now", now); else delete performance.now; globalThis.scheduler = scheduler; }
});

test("PDF headings moved to a fresh page consume the rendered text budget once", async () => {
  for (const [options, cells, characters] of [
    [{ columns: [{ header: "Name", key: "name" }] }, 3, 26],
    [{ columns: [{ header: "Name", key: "name", groups: ["Group"] }] }, 4, 31],
    [{ columns: [{ header: "Name", key: "name" }], headerRows: [[{ value: "Group" }], [{ value: "Name" }]] }, 4, 31]
  ]) {
    const blocks = () => [{ kind: "paragraph", text: "One\nTwo\nThree\nFour" }, pdfTable([{ name: "Item" }], options)];
    const configuration = { pageSize: { width: 100, height: 100 }, margins: 10, pageNumbers: false, compression: false,
      limits: { maxCells: cells, maxTextCharacters: characters } };
    const blob = await writePdfDocument(blocks(), configuration), chunks = [];
    await writePdfDocumentTo(blocks(), { write: bytes => { chunks.push(bytes.slice()); } }, configuration);
    assert.deepEqual(Buffer.from(await blob.arrayBuffer()), Buffer.from(await new Blob(chunks).arrayBuffer()));
    const pdf = await inspectPdf(blob);
    assert.equal(pdf.pages.length, 2);
    assert.ok(!pdf.pages[0].lines.includes("Name") && pdf.pages[1].lines.includes("Name"));
    assert.ok(pdf.pages[1].lines.includes("Item"));
    await assert.rejects(writePdfDocument(blocks(), { ...configuration, limits: { maxCells: cells - 1 } }), /maxCells/);
  }
});

test("PDF documents share row, text, page, image and vector budgets across blocks", async () => {
  const table = () => pdfTable([{ name: "Row", amount: 1 }], { columns });
  for (const [blocks, limits, error] of [
    [[table(), table()], { maxRows: 1 }, /maxRows/],
    [[{ kind: "paragraph", text: "One" }, { kind: "paragraph", text: "Two" }], { maxTextCharacters: 5 }, /maxTextCharacters/],
    [[{ kind: "pageBreak" }], { maxPages: 1 }, /maxPages/],
    [[{ kind: "image", image: logo }], { maxImageBytes: 15 }, /maxImageBytes/],
    [[{ kind: "image", image: logo }, { kind: "image", image: { ...logo } }], { maxImageBytes: 20 }, /maxImageBytes/],
    [[{ kind: "image", image: logo }], { maxImagePixels: 3 }, /maxImagePixels/],
    [[vector, vector], { maxVectorCommands: 7 }, /maxVectorCommands/],
    [[vector], { maxOutputBytes: 10 }, /maxOutputBytes/]
  ]) await assert.rejects(writePdfDocument(blocks, { limits, pageNumbers: false }), error);
  const pdf = await inspectPdf(await writePdfDocument([{ kind: "image", image: logo }, { kind: "image", image: logo }], { limits: { maxImageBytes: 16 } }));
  assert.equal([...pdf.objects.values()].filter(o => /\/Subtype \/Image/.test(o.body)).length, 2);
});

test("PDF caption budgets count painted occurrences rather than unused continuation definitions", async () => {
  for (const bands of [false, true]) {
    const table = () => ({ ...pdfTable([["ROW"]], { columns: [{ header: "H" }], ...(bands ? {
      columnWidths: [60], wideTable: { mode: "bands", repeatColumns: [] }
    } : {}) }), continuationCaption: { text: "UNUSED" } });
    const pdf = await inspectPdf(await writePdfDocument([table()], { pageNumbers: false, compression: false,
      limits: { maxCells: 2, maxTextCharacters: 4 } }));
    assert.equal(pdf.pages.length, 1); assert.ok(pdf.text.includes("ROW")); assert.ok(!pdf.text.includes("UNUSED"));
    await assert.rejects(writePdfDocument([{ ...table(), caption: { text: "USED" } }], { pageNumbers: false,
      limits: { maxTextCharacters: 7 } }), /maxTextCharacters/);
  }
  const blocks = () => [{ ...pdfTable(Array.from({ length: 6 }, () => ["ROW"]), { columns: [{ header: "H" }] }),
    caption: { text: "START" }, continuationCaption: { text: "NEXT" } }];
  const options = { pageSize: { width: 140, height: 120 }, margins: 10, pageNumbers: false, compression: false };
  const reference = await inspectPdf(await writePdfDocument(blocks(), options));
  assert.ok(reference.pages.length > 1);
  const occurrences = reference.pages.flatMap(page => page.lines), cells = occurrences.length,
    characters = occurrences.reduce((sum, value) => sum + value.length, 0);
  assert.ok(occurrences.includes("START") && occurrences.includes("NEXT"));
  await writePdfDocument(blocks(), { ...options, limits: { maxCells: cells, maxTextCharacters: characters } });
  await assert.rejects(writePdfDocument(blocks(), { ...options, limits: { maxCells: cells - 1 } }), /maxCells/);
  await assert.rejects(writePdfDocument(blocks(), { ...options, limits: { maxTextCharacters: characters - 1 } }), /maxTextCharacters/);
});

test("PDF documents reject unsupported graphics before embedding their payload", async () => {
  assert.throws(() => pdfTable([], { columns, fontSize: 36 }), /document settings/);
  await assert.rejects(writePdfDocument([{ kind: "table", rows: [], options: { columns, padding: 40 } }]), /document settings/);
  for (const image of [{ ...logo, width: 3 }, { ...logo, width: Number.MAX_SAFE_INTEGER }, { format: "png", data: Uint8Array.of(1) }])
    await assert.rejects(writePdfDocument([{ kind: "image", image }]), /Raw image|dimensions|support|PNG signature/);
  for (const block of [{ ...vector, commands: [{ op: "line", x: 1, y: 1 }] }, { ...vector, commands: [{ op: "move", x: 81, y: 1 }] },
    { ...vector, commands: [{ op: "svg", data: "<svg/>" }] }]) await assert.rejects(writePdfDocument([block]), /subpath|bounds|command/);
  const sof = jpeg.findIndex((b, i) => b === 255 && jpeg[i + 1] === 192), cmyk = jpeg.slice(); cmyk[sof + 9] = 4;
  await assert.rejects(writePdfDocument([{ kind: "image", image: { format: "jpeg", data: cmyk } }]), /CMYK/);
  for (const [marker, signature, error] of [[226, "ICC_PROFILE\0", /ICC/], [225, "Exif\0\0", /EXIF/]]) {
    const metadata = Buffer.from(signature), segment = Buffer.from([255, marker, 0, metadata.length + 2]);
    const data = new Uint8Array(Buffer.concat([jpeg.subarray(0, 2), segment, metadata, jpeg.subarray(2)]));
    await assert.rejects(writePdfDocument([{ kind: "image", image: { format: "jpeg", data } }]), error);
  }
});

test("PDF document table streaming applies backpressure and returns active block and row sources", async () => {
  let writes = 0, accepted = 0, returned = 0, produced = 0;
  async function* rows() { try { for (let i = 0; i < 120; i++) { if (i === 100) assert.ok(writes > 0); produced++; yield { name: "Row " + i, amount: i }; } } finally { returned++; } }
  const chunks = [], destination = new WritableStream({ async write(b) {
    writes++; const before = produced; await Promise.resolve(); assert.equal(produced, before, "Input must stop while the destination accepts bytes");
    accepted += b.length; chunks.push(b.slice());
  } });
  const result = await writePdfDocumentTo([pdfTable(rows(), { columns })], destination, { compression: false });
  assert.deepEqual(result, { rows: 120, columns: 2, bytes: accepted }); assert.equal(returned, 1); assert.equal(destination.locked, false);
  assert.ok((await inspectPdf(new Blob(chunks))).text.includes("Row 119"));
  let blocksReturned = 0, rowsReturned = 0;
  const badRows = { *[Symbol.iterator]() { try { yield { name: "First", amount: 1 }; yield { name: "Second", amount: 2 }; } finally { rowsReturned++; } } };
  const blocks = { *[Symbol.iterator]() { try { yield pdfTable(badRows, { columns }); yield { kind: "paragraph", text: "Never" }; } finally { blocksReturned++; } } };
  await assert.rejects(writePdfDocument(blocks, { limits: { maxRows: 1 } }), /maxRows/);
  assert.equal(blocksReturned, 1); assert.equal(rowsReturned, 1);
});

test("PDF document cancellation releases the destination during RGBA preparation", async () => {
  const now = Object.getOwnPropertyDescriptor(performance, "now"), scheduler = globalThis.scheduler, abort = new AbortController();
  const destination = new WritableStream({ write() { throw Error("No page should be emitted"); } });
  try {
    let clock = 0; Object.defineProperty(performance, "now", { configurable: true, value: () => clock += 40 });
    globalThis.scheduler = { postTask: async callback => { abort.abort(Error("image abort")); callback(); } };
    await assert.rejects(writePdfDocumentTo([{ kind: "image", image: { format: "rgba", width: 256, height: 128, data: new Uint8Array(256 * 128 * 4) } }], destination,
      { signal: abort.signal, pageNumbers: false }), /image abort/);
    assert.equal(destination.locked, false);
  } finally { if (now) Object.defineProperty(performance, "now", now); else delete performance.now; globalThis.scheduler = scheduler; }
});
