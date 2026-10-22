import test from "node:test";
import assert from "node:assert/strict";
import { writePdfDocument, writePdfDocumentTo, pdfGroupedTable, createGroupedReport, ExportCell } from "../dist/pdf/index.js";
import { PdfObjects } from "../dist/pdf/objects.js";
import { pdfXrefChunks } from "../dist/pdf/xref.js";
import { inspectPdf } from "./pdf-reader.mjs";

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
    const body = object.body, kids = /\/K \[([\s\S]*?)\](?: |$)/.exec(body)[1];
    const headerNames = [...(/\/Headers \[([^\]]*)\]/.exec(body)?.[1] ?? "").matchAll(/\(([^)]+)\)/g)].map(match => {
      assert.ok(headers.has(match[1])); return headers.get(match[1]);
    });
    return { role: /\/S \/(\w+)/.exec(body)[1], parent: identities.get(Number(/\/P (\d+) 0 R/.exec(body)[1])) ?? "root",
      children: [...kids.replace(/<<[\s\S]*?>>/g, "").matchAll(/(\d+) 0 R/g)].map(match => identities.get(Number(match[1]))),
      marks: [...kids.matchAll(/\/MCR \/Pg (\d+) 0 R \/MCID (\d+)/g)].map(match => [pages.get(Number(match[1])), Number(match[2])]),
      annotations: [...kids.matchAll(/\/OBJR \/Obj (\d+) 0 R \/Pg (\d+) 0 R/g)].map(match => [annotations.get(Number(match[1])), pages.get(Number(match[2]))]),
      scope: /\/Scope \/(\w+)/.exec(body)?.[1], colSpan: /\/ColSpan (\d+)/.exec(body)?.[1], rowSpan: /\/RowSpan (\d+)/.exec(body)?.[1], headers: headerNames };
  });
}

test("compressed structure containers preserve banded grouped reports, headers, links and reading order", async () => {
  const plain = await inspectPdf(await report(false)), compact = await inspectPdf(await report(true));
  assert.deepEqual(compact.pages.map(page => page.content), plain.pages.map(page => page.content));
  assert.deepEqual(semanticGraph(compact), semanticGraph(plain));
  assert.ok(compact.bytes.length < plain.bytes.length, "This repeated tagged workload benefits from structure compression");
  const containers = [...compact.objects.values()].filter(object => /\/Type \/ObjStm\b/.test(object.body));
  assert.ok(containers.length > 1, "The report must exercise several completed batches");
  for (const container of containers) {
    assert.ok(Number(/\/N (\d+)/.exec(container.body)[1]) <= 100);
    assert.ok(container.stream.length <= 32768, "Decoded batches retain a bounded member index and body");
    const lengthId = Number(/\/Length (\d+) 0 R/.exec(container.body)[1]);
    assert.match(compact.objects.get(lengthId).body, /^\d+$/, "A container length stays an ordinary scalar object");
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
    let enter, accept, completed = false;
    const observed = new Promise(resolve => { enter = resolve; }), held = new Promise(resolve => { accept = resolve; });
    const objects = new PdfObjects({ async write() {
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
  }
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
    await assert.rejects(changing.compactObject(changing.reserve(), "<< /Type /StructElem /S /P /K [] >>"), error => error === reason);
  } finally { globalThis.CompressionStream = native; }
});

test("xref entries preserve offsets above 32 bits and compact reservations still reject invalid object state", async () => {
  const bytes = Buffer.concat([...pdfXrefChunks([0, 4294967297, 7], [undefined, undefined, 99])]);
  assert.equal(bytes[11], 1); assert.equal(bytes.readBigUInt64BE(12), 4294967297n);
  assert.equal(bytes[22], 2); assert.equal(bytes.readBigUInt64BE(23), 7n); assert.equal(bytes.readUInt16BE(31), 99);
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
