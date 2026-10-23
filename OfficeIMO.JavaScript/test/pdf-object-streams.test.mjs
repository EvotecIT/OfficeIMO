import test from "node:test";
import assert from "node:assert/strict";
import { writePdfDocument, writePdfDocumentTo, pdfGroupedTable, createGroupedReport, ExportCell } from "../dist/pdf/index.js";
import { PdfObjects } from "../dist/pdf/objects.js";
import { pdfXrefChunks } from "../dist/pdf/xref.js";
import { inspectPdf, pdfStructureAttributes } from "./pdf-reader.mjs";
import { pdfStructureChildren } from "./pdf-structure-reader.mjs";

function report(compression, { sink, signal, limits, returned } = {}) {
  const columns = [{ header: "ID", key: "id" }, { header: "Site", key: "site" }, { header: "Amount", key: "amount", type: "number" }];
  function* rows() {
    try {
      for (let index = 0; index < 80; index++) yield { id: new ExportCell("REC-" + index, { link: { target: "https://example.org/record/" + index } }),
        site: index < 40 ? "Alpha" : "Beta", amount: index + .5 };
    } finally { returned?.(); }
  }
  const source = createGroupedReport(rows(), { columns, groupBy: "site", totals: { amount: "sum" } });
  const blocks = [{ kind: "paragraph", text: "Inventory report", headingLevel: 1 },
    pdfGroupedTable(source, { labelColumn: { header: "Summary" }, subtotalLabel: group => "Total " + group.label,
      grandTotalLabel: () => "Grand total", continuationLabel: group => group.label + " continued", columnWidths: [60, 70, 100, 100],
      wideTable: { mode: "bands", repeatColumns: [0, 1] } })];
  const options = { tagged: { language: "en" }, pageSize: { width: 280, height: 320 }, margins: 20, pageNumbers: false,
    compression, signal, limits: { maxTableWindowRows: 7, ...limits } };
  return sink === undefined ? writePdfDocument(blocks, options) : writePdfDocumentTo(blocks, sink, options);
}

/** Object numbers may change when containers are reserved. Compare the xref-resolved logical graph,
 * including its page-local marks, header identities and native annotation owners. */
function semanticGraph(pdf) {
  const nodes = [...pdf.objects].filter(([, object]) => /\/Type \/StructElem\b/.test(object.body));
  const identities = new Map(nodes.map(([id], index) => [id, index])), pages = new Map(pdf.pages.map((page, index) => [page.id, index]));
  const annotations = new Map([...pdf.objects].filter(([, object]) => /\/Subtype \/Link\b/.test(object.body)).map(([id], index) => [id, index]));
  const headers = new Map(nodes.flatMap(([id, object]) => {
    const identifier = /\/ID \(([^)]+)\)/.exec(object.body)?.[1]; return identifier ? [[identifier, identities.get(id)]] : [];
  }));
  return nodes.map(([, object]) => {
    const body = object.body, attributes = pdfStructureAttributes(pdf, body), { entries } = pdfStructureChildren(body);
    const headerNames = [...(/\/Headers \[([^\]]*)\]/.exec(attributes)?.[1] ?? "").matchAll(/\(([^)]+)\)/g)].map(match => {
      assert.ok(headers.has(match[1])); return headers.get(match[1]);
    });
    return { role: /\/S \/(\w+)/.exec(body)[1], parent: identities.get(Number(/\/P (\d+) 0 R/.exec(body)[1])) ?? "root",
      children: entries.filter(child => child.kind === "element").map(child => { assert.ok(identities.has(child.id)); return identities.get(child.id); }),
      marks: entries.filter(child => child.kind === "mark").map(child => { assert.ok(pages.has(child.pageId)); return [pages.get(child.pageId), child.mcid]; }),
      annotations: entries.filter(child => child.kind === "annotation").map(child => {
        assert.ok(annotations.has(child.id) && pages.has(child.pageId)); return [annotations.get(child.id), pages.get(child.pageId)];
      }),
      scope: /\/Scope \/(\w+)/.exec(attributes)?.[1], colSpan: /\/ColSpan (\d+)/.exec(attributes)?.[1], rowSpan: /\/RowSpan (\d+)/.exec(attributes)?.[1], headers: headerNames };
  });
}

/** Resolve every synthetic value through the emitted xref, including ordinary overflow values. */
async function compactFixture(bodies, afterMember) {
  const chunks = [], expected = new Map(), objects = new PdfObjects({ write(bytes) { chunks.push(bytes.slice()); } }, undefined, undefined, true);
  await objects.text("%PDF-1.7\n");
  const catalog = objects.reserve(), pages = objects.reserve(), page = objects.reserve(), resources = objects.reserve(), content = objects.reserve(), info = objects.reserve();
  await objects.object(catalog, "<< /Type /Catalog /Pages " + pages + " 0 R >>");
  await objects.object(pages, "<< /Type /Pages /Kids [" + page + " 0 R] /Count 1 >>");
  await objects.object(page, "<< /Type /Page /Parent " + pages + " 0 R /MediaBox [0 0 100 100] /Resources " + resources + " 0 R /Contents " + content + " 0 R >>");
  await objects.object(resources, "<< >>"); await objects.stream(content, "", "", false); await objects.object(info, "<< >>");
  for (const value of bodies) {
    const id = objects.reserve(), body = typeof value === "function" ? value(id) : value;
    expected.set(id, Buffer.from(body, "utf8").toString("latin1"));
    await objects.compactObject(id, body);
    await afterMember?.(objects);
  }
  await objects.finish(catalog, info);
  const pdf = await inspectPdf(new Blob(chunks));
  for (const [id, body] of expected) assert.equal(pdf.objects.get(id)?.body, body, "The xref must resolve each complete UTF-8 value exactly once");
  const containers = [...pdf.objects.values()].filter(object => /\/Type \/ObjStm\b/.test(object.body));
  for (const container of containers) {
    assert.ok(container.stream.length <= 32768, "The decoded index, member bytes and newlines share one byte ceiling");
    assert.ok(Number(/\/N (\d+)/.exec(container.body)[1]) <= 256);
  }
  return { pdf, containers, ids: [...expected.keys()] };
}

test("compressed structure containers preserve banded grouped reports, headers, links and reading order", async () => {
  const plain = await inspectPdf(await report(false)), compact = await inspectPdf(await report(true));
  assert.deepEqual(compact.pages.map(page => page.content), plain.pages.map(page => page.content));
  assert.deepEqual(semanticGraph(compact), semanticGraph(plain));
  assert.ok(compact.bytes.length < plain.bytes.length, "This repeated tagged workload benefits from structure compression");
  const containers = [...compact.objects.values()].filter(object => /\/Type \/ObjStm\b/.test(object.body));
  assert.ok(containers.length > 1, "The report must exercise several completed batches");
  for (const container of containers) {
    assert.ok(Number(/\/N (\d+)/.exec(container.body)[1]) <= 256);
    assert.ok(container.stream.length <= 32768, "Decoded batches retain a bounded member index and body");
    const lengthId = Number(/\/Length (\d+) 0 R/.exec(container.body)[1]);
    assert.match(compact.objects.get(lengthId).body, /^\d+$/, "A container length stays an ordinary scalar object");
  }
});

test("partial UTF-8 encodings preserve committed members and retry whole multibyte values", async () => {
  for (const [scalar, padding] of [["é", 32748], ["€", 32747], ["😀", 32747], ["\ud800", 32747]]) {
    const { containers } = await compactFixture(["(seed)", "(" + "x".repeat(padding) + scalar + ")"]);
    assert.equal(containers.length, 2, "A value that exceeds the remaining suffix must fit intact in a fresh container");
    assert.ok(containers[0].stream.length < 100, "A failed tentative encoding must not extend the committed prefix");
  }
});

test("the decoded byte ceiling includes the index and newlines, with UTF-8 overflow remaining ordinary", async () => {
  const { pdf, containers, ids } = await compactFixture([
    id => "(" + "x".repeat(32768 - String(id).length - 7) + ")",
    "(" + "😀".repeat(8191) + ")",
    "(following compact value)"
  ]);
  assert.equal(containers[0].stream.length, 32768, "A complete first member may exactly fill the decoded byte ceiling");
  assert.equal(containers.length, 2);
  assert.equal(pdf.objects.get(ids[0]).bodyOffset, undefined);
  assert.notEqual(pdf.objects.get(ids[1]).bodyOffset, undefined, "UTF-16 length alone cannot qualify a UTF-8 value for packing");
  assert.equal(pdf.objects.get(ids[2]).bodyOffset, undefined, "An ordinary overflow must leave subsequent compact xref membership valid");
});

test("completed payloads retain their bytes while later batches reuse the bounded member storage", async () => {
  const { containers } = await compactFixture(Array.from({ length: 517 }, (_, index) =>
    "(" + String(index).padStart(3, "0") + ":" + "Z".repeat(24) + (index % 2 ? "é" : "😀") + ")"));
  assert.deepEqual(containers.map(container => Number(/\/N (\d+)/.exec(container.body)[1])), [256, 256, 5]);
});

test("serialized compact xrefs preserve index 255 and ordinary entries across membership storage chunks", async () => {
  const { pdf, containers, ids } = await compactFixture(Array.from({ length: 4352 }, (_, index) => "(record-" + index + ")"));
  const xref = [...pdf.objects.values()].find(object => /\/Type \/XRef\b/.test(object.body));
  const after = ids.findIndex(id => id >= 4096), before = after - 1;
  assert.ok(before >= 0 && after < ids.length, "Real object IDs must cross the membership chunk boundary");
  for (const member of [255, before, after, ids.length - 1]) {
    const position = ids[member] * 11;
    assert.equal(xref.stream[position], 2);
    assert.equal(xref.stream.readUInt16BE(position + 9), member % 256);
    assert.equal(pdf.objects.get(ids[member]).body, "(record-" + member + ")");
  }
  for (const container of containers) {
    assert.equal(Number(/\/N (\d+)/.exec(container.body)[1]), 256);
    const lengthId = Number(/\/Length (\d+) 0 R/.exec(container.body)[1]);
    for (const id of [container.id, lengthId]) {
      assert.equal(xref.stream[id * 11], 1, "Untouched membership slots retain ordinary xref entries");
      assert.equal(xref.stream.readUInt16BE(id * 11 + 9), 0);
    }
  }
});

test("page-byte delivery preserves pending structure members across bounded containers", async () => {
  const bodies = Array.from({ length: 517 }, (_, index) => "(" + index + ":" + "Z".repeat(24) + ")");
  const ordinary = await compactFixture(bodies), pageFlushed = await compactFixture(bodies, objects => objects.flushBytes());
  assert.deepEqual(pageFlushed.containers.map(container => container.stream), ordinary.containers.map(container => container.stream));
  assert.deepEqual(pageFlushed.pdf.pages.map(page => page.content), ordinary.pdf.pages.map(page => page.content));
});

test("asynchronous native codecs consume each arena batch before the next batch overwrites it", async () => {
  const native = globalThis.CompressionStream;
  try {
    globalThis.CompressionStream = class {
      constructor(format) {
        const stream = new native(format), writer = stream.writable.getWriter();
        this.readable = stream.readable;
        this.writable = new WritableStream({
          async write(chunk) { await new Promise(resolve => setImmediate(resolve)); await writer.write(chunk); },
          async close() { try { await writer.close(); } finally { writer.releaseLock(); } },
          async abort(reason) { try { await writer.abort(reason); } finally { writer.releaseLock(); } }
        });
      }
    };
    await compactFixture(Array.from({ length: 517 }, (_, index) =>
      "(" + String(index).padStart(3, "0") + ":" + "Z".repeat(24) + (index % 2 ? "é" : "😀") + ")"));
  } finally { globalThis.CompressionStream = native; }
});

test("TextEncoder implementations without encodeInto preserve complete packed UTF-8 values", async () => {
  const bodies = ["(seed)", "(" + "x".repeat(32747) + "😀é\ud800)", "(tail)"], native = await compactFixture(bodies);
  const prototype = TextEncoder.prototype, descriptor = Object.getOwnPropertyDescriptor(prototype, "encodeInto");
  try {
    Object.defineProperty(prototype, "encodeInto", { value: undefined, configurable: true, writable: true });
    const fallback = await compactFixture(bodies);
    assert.deepEqual(fallback.containers.map(container => container.stream), native.containers.map(container => container.stream));
  } finally {
    if (descriptor) Object.defineProperty(prototype, "encodeInto", descriptor);
    else delete prototype.encodeInto;
  }
});

test("disabled and unavailable native compression preserve the ordinary-object classic-xref representation", async () => {
  const plain = Buffer.from(await (await report(false)).arrayBuffer()), native = globalThis.CompressionStream;
  try {
    for (const replacement of [undefined, class { constructor() { throw new TypeError("Deflate unsupported"); } }]) {
      globalThis.CompressionStream = replacement;
      const bytes = Buffer.from(await (await report(true)).arrayBuffer());
      assert.deepEqual(bytes, plain);
      await inspectPdf(new Blob([bytes]));
    }
  } finally { globalThis.CompressionStream = native; }
});

test("pending compact output awaits a slow sink and preserves sink failures and cancellation", async () => {
  for (const mode of ["slow", "sink", "abort"]) {
    const controller = new AbortController(), reason = new Error(mode + " stopped");
    let enter, accept, completed = false, borrowed, captured;
    const observed = new Promise(resolve => { enter = resolve; }), held = new Promise(resolve => { accept = resolve; });
    const objects = new PdfObjects({ async write(bytes) {
      borrowed = bytes; captured = bytes.slice();
      enter();
      if (mode === "sink") throw reason;
      if (mode === "abort") { controller.abort(reason); return new Promise(() => {}); }
      await held;
    } }, controller.signal, undefined, true);
    await objects.compactObject(objects.reserve(), "<< /Type /StructElem /S /Document /K [] >>");
    const pending = objects.flush().then(() => { completed = true; });
    const failure = mode === "slow" ? undefined : assert.rejects(pending, error => error === reason);
    await observed;
    if (mode === "slow") { assert.equal(completed, false); accept(); await pending; }
    else await failure;
    assert.deepEqual(borrowed, captured, "Pending or rejected output must keep its original byte view");
  }
});

test("PDF output stays intact until each asynchronous destination acknowledgement", async () => {
  const expected = Buffer.from(await (await report(true)).arrayBuffer()), chunks = [];
  await report(true, { sink: { async write(bytes) {
    const captured = bytes.slice();
    await new Promise(resolve => setImmediate(resolve));
    assert.deepEqual(bytes, captured, "The next PDF bytes must not overwrite a pending destination write");
    chunks.push(bytes.slice());
  } } });
  assert.ok(chunks.length > 1, "The report must exercise consecutive output windows");
  assert.deepEqual(Buffer.concat(chunks), expected);
  await inspectPdf(new Blob(chunks));
});

test("compressed structure output ceilings include containers, length objects and the final xref", async () => {
  const expected = await report(true), chunks = [];
  let returned = 0;
  const result = await report(true, { sink: { write(bytes) { chunks.push(bytes.slice()); } }, limits: { maxOutputBytes: expected.size } });
  assert.equal(result.bytes, expected.size); await inspectPdf(new Blob(chunks));
  await assert.rejects(report(true, { sink: { write() {} }, limits: { maxOutputBytes: expected.size - 1 }, returned() { returned++; } }), { code: "RESOURCE_LIMIT" });
  assert.equal(returned, 1, "An output rejection still returns the single-use source");
});

test("compaction propagates codec failures and never silently changes representation after a completed batch", async () => {
  const native = globalThis.CompressionStream, reason = new Error("Compaction codec stopped");
  try {
    globalThis.CompressionStream = class {
      constructor() { const stream = new TransformStream({ transform() { throw reason; } }); this.readable = stream.readable; this.writable = stream.writable; }
    };
    const failing = new PdfObjects({ write() {} }, undefined, undefined, true);
    await failing.compactObject(failing.reserve(), "<< /Type /StructElem /S /Document /K [] >>");
    await assert.rejects(failing.flush(), error => error === reason);
    globalThis.CompressionStream = native;
    const changing = new PdfObjects({ write() {} }, undefined, undefined, true);
    await changing.compactObject(changing.reserve(), "<< /Type /StructElem /S /Document /K [] >>");
    await changing.flush();
    globalThis.CompressionStream = class { constructor() { throw reason; } };
    await changing.compactObject(changing.reserve(), "<< /Type /StructElem /S /P /K [] >>");
    await assert.rejects(changing.flush(), error => error === reason);
  } finally { globalThis.CompressionStream = native; }
});

test("xref entries preserve offsets above 32 bits and compact reservations still reject invalid object state", async () => {
  const bytes = Buffer.concat([...pdfXrefChunks([0, 4294967297, 9999999999, 7], [undefined, undefined, undefined, 255])]);
  assert.equal(bytes[11], 1); assert.equal(bytes.readBigUInt64BE(12), 4294967297n);
  assert.equal(bytes[22], 1); assert.equal(bytes.readBigUInt64BE(23), 9999999999n);
  assert.equal(bytes[33], 2); assert.equal(bytes.readBigUInt64BE(34), 7n); assert.equal(bytes.readUInt16BE(42), 255);
  assert.throws(() => [...pdfXrefChunks([0, 1], [undefined])], /locations must align/);
  const objects = new PdfObjects({ write() {} }, undefined, undefined, true), id = objects.reserve();
  await objects.compactObject(id, "<< /Type /StructElem /S /Document /K [] >>");
  await assert.rejects(objects.compactObject(id, "<< >>"), /Invalid PDF object state/);
  await assert.rejects(objects.object(id, "<< >>"), /Invalid PDF object state/);
  objects.reserve(); await assert.rejects(objects.finish(id, id), /unwritten objects/);
});

test("oversized structure values and iterable parents remain forward-only ordinary objects", async () => {
  const chunks = [], objects = new PdfObjects({ write(bytes) { chunks.push(bytes.slice()); } }, undefined, undefined, true);
  const compact = objects.reserve(); await objects.compactObject(compact, "<< /Type /StructElem /S /Document /K [] >>");
  const large = objects.reserve(), body = "<< /Type /StructElem /S /Figure /Alt <" + "0041".repeat(10000) + "> /K [] >>";
  await objects.compactObject(large, body);
  const parent = objects.reserve(), parts = ["<< /Type /StructElem /S /Table /K [", (compact + " 0 R ").repeat(10000), "] >>"];
  await objects.object(parent, parts); await objects.flush();
  const bytes = Buffer.concat(chunks).toString("latin1");
  assert.ok(bytes.includes(large + " 0 obj\n" + body + "\nendobj\n"));
  assert.ok(bytes.includes(parent + " 0 obj\n" + parts.join("") + "\nendobj\n"));
  assert.ok(chunks.every(chunk => chunk.length <= 32768));
});
