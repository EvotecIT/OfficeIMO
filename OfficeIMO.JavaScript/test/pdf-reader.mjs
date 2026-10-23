// Independent narrow fixture inspector: byte offsets, stream lengths, zlib, font Unicode maps and page content.
// General PDF interoperability is separately checked by OfficeIMO.Pdf, Poppler and MuPDF.
import assert from "node:assert/strict";
import { inflateSync } from "node:zlib";

function plainObject(text, offset, expected) {
  const heading = /^(\d+) 0 obj\n/.exec(text.slice(offset));
  assert.ok(heading, `An xref offset must target an indirect object: ${offset}`);
  const id = Number(heading[1]); if (expected !== undefined) assert.equal(id, expected);
  const bodyOffset = offset + heading[0].length, end = text.indexOf("\nendobj\n", bodyOffset);
  assert.ok(end > bodyOffset);
  return { id, bodyOffset, body: text.slice(bodyOffset, end) };
}

function readStream(object, objects, bytes, text) {
  const marker = object.body.indexOf("\nstream\n");
  if (marker < 0) return;
  const dictionary = object.body.slice(0, marker), declared = /\/Length (\d+)( 0 R)?/.exec(dictionary);
  assert.ok(declared);
  const length = declared[2] ? Number(objects.get(Number(declared[1]))?.body) : Number(declared[1]);
  assert.ok(Number.isSafeInteger(length) && length >= 0);
  const begin = object.bodyOffset + marker + 8;
  assert.equal(text.slice(begin + length, begin + length + 18), "\nendstream\nendobj\n");
  let stream = bytes.subarray(begin, begin + length);
  if (/\/Filter \/FlateDecode/.test(dictionary)) stream = inflateSync(stream);
  object.body = text.slice(object.bodyOffset, begin + length + 10);
  object.stream = stream;
}

function xrefEntries(bytes, text, start) {
  const entries = new Map();
  if (text.slice(start, start + 5) === "xref\n") {
    const xref = text.slice(start).split("\n"), count = Number(xref[1].split(" ")[1]);
    for (let id = 1; id < count; id++) entries.set(id, { type: 1, offset: Number(xref[id + 2].slice(0, 10)) });
    return { count, entries };
  }
  const object = plainObject(text, start);
  assert.match(object.body, /\/Type \/XRef\b/); readStream(object, new Map(), bytes, text);
  const count = Number(/\/Size (\d+)/.exec(object.body)[1]);
  const widths = /\/W \[([^\]]+)\]/.exec(object.body)[1].trim().split(/\s+/).map(Number);
  const sections = /\/Index \[([^\]]+)\]/.exec(object.body)?.[1].trim().split(/\s+/).map(Number) ?? [0, count];
  assert.equal(widths.length, 3); assert.equal(sections.length % 2, 0);
  let position = 0;
  for (let section = 0; section < sections.length; section += 2) for (let index = 0; index < sections[section + 1]; index++) {
    const id = sections[section] + index, fields = widths.map((width, field) => {
      if (!width) return field === 0 ? 1 : 0;
      let value = 0;
      for (let byte = 0; byte < width; byte++) { assert.ok(position < object.stream.length); value = value * 256 + object.stream[position++]; }
      assert.ok(Number.isSafeInteger(value)); return value;
    });
    assert.ok(!entries.has(id));
    if (fields[0] === 0) { assert.equal(id, 0); assert.equal(fields[2], 65535); }
    else if (fields[0] === 1) { assert.equal(fields[2], 0); entries.set(id, { type: 1, offset: fields[1] }); }
    else { assert.equal(fields[0], 2); entries.set(id, { type: 2, stream: fields[1], index: fields[2] }); }
  }
  assert.equal(position, object.stream.length);
  assert.equal(entries.get(object.id)?.offset, start, "The xref stream must index itself");
  return { count, entries };
}

/** Resolve the fixture writer's inline or indirect structure attribute dictionary through its xref. */
export function pdfStructureAttributes(pdf, body) {
  const reference = /\/A (\d+) 0 R\b/.exec(body);
  if (!reference) return /\/A (<<[\s\S]*?>>)/.exec(body)?.[1] ?? "";
  const attributes = pdf.objects.get(Number(reference[1]));
  assert.ok(attributes, "Every structure attribute reference must resolve through the xref");
  assert.match(attributes.body, /^<< \/O \//);
  return attributes.body;
}

export async function inspectPdf(blob) {
  const bytes = Buffer.from(await blob.arrayBuffer()), text = bytes.toString("latin1");
  assert.ok(text.startsWith("%PDF-1.7\n"));
  const start = Number(/startxref\n(\d+)\n%%EOF\n$/.exec(text)?.[1]);
  const { count, entries } = xrefEntries(bytes, text, start);
  let objects = new Map();
  for (let id = 1; id < count; id++) {
    const entry = entries.get(id); assert.ok(entry);
    if (entry.type === 1) objects.set(id, plainObject(text, entry.offset, id));
  }
  // Stream Length may be an indirect object emitted after the streamed bytes.
  for (const object of objects.values()) readStream(object, objects, bytes, text);
  const unpacked = new Set();
  for (const [id, entry] of entries) if (entry.type === 2) {
    if (unpacked.has(entry.stream)) continue;
    const container = objects.get(entry.stream); assert.ok(container?.stream);
    assert.match(container.body, /\/Type \/ObjStm\b/);
    const size = Number(/\/N (\d+)/.exec(container.body)[1]), first = Number(/\/First (\d+)/.exec(container.body)[1]);
    const index = container.stream.subarray(0, first).toString("ascii").trim().split(/\s+/).map(Number);
    assert.equal(index.length, size * 2); assert.ok(entry.index < size); assert.equal(index[entry.index * 2], id);
    for (let member = 0; member < size; member++) {
      const memberId = index[member * 2], begin = first + index[member * 2 + 1], end = member + 1 < size ? first + index[member * 2 + 3] : container.stream.length;
      assert.ok(begin >= first && begin < end && end <= container.stream.length);
      assert.ok(!objects.has(memberId));
      const owner = entries.get(memberId); assert.equal(owner?.type, 2); assert.equal(owner.stream, entry.stream); assert.equal(owner.index, member);
      objects.set(memberId, { body: container.stream.subarray(begin, end).toString("latin1").trim() });
    }
    unpacked.add(entry.stream);
  }
  objects = new Map([...objects].sort(([left], [right]) => left - right));
  const fontMaps = new Map();
  for (const [id, object] of objects) if (/\/Subtype \/Type0/.test(object.body)) {
    const unicode = Number(/\/ToUnicode (\d+) 0 R/.exec(object.body)[1]), mapping = new Map();
    const cmap = objects.get(unicode).stream.toString();
    for (const block of cmap.matchAll(/\d+ beginbfchar\n([\s\S]*?)endbfchar/g)) for (const m of block[1].matchAll(/<([0-9a-f]+)> <([0-9a-f]+)>/gi)) {
      let value = ""; for (let i = 0; i < m[2].length; i += 4) value += String.fromCharCode(parseInt(m[2].slice(i, i + 4), 16));
      mapping.set(parseInt(m[1], 16), value);
    }
    fontMaps.set(id, mapping);
  }
  const pageTree = [...objects.values()].find(o => /\/Type \/Pages /.test(o.body));
  const pageIds = [.../\/Kids \[([^\]]*)\]/.exec(pageTree.body)[1].matchAll(/(\d+) 0 R/g)].map(m => Number(m[1]));
  assert.equal(pageIds.length, Number(/\/Count (\d+)/.exec(pageTree.body)[1]));
  const pages = pageIds.map(id => {
    const object = objects.get(id), content = objects.get(Number(/\/Contents (\d+) 0 R/.exec(object.body)[1])).stream.toString();
    const resources = objects.get(Number(/\/Resources (\d+) 0 R/.exec(object.body)[1])).body;
    const names = new Map([...resources.matchAll(/\/(F\d+) (\d+) 0 R/g)].map(m => [m[1], Number(m[2])]));
    const lines = [];
    for (const m of content.matchAll(/BT \/(F\d+) [\s\S]*?<([a-f0-9]*)> Tj ET/g)) {
      const map = fontMaps.get(names.get(m[1]));
      if (map) { let line = ""; for (let i = 0; i < m[2].length; i += 4) { const cp = map.get(parseInt(m[2].slice(i, i + 4), 16)); assert.notEqual(cp, undefined); line += cp; } lines.push(line); }
      else lines.push(new TextDecoder("windows-1252").decode(Buffer.from(m[2], "hex")));
    }
    return { id, content, lines, body: object.body };
  });
  return { bytes, objects, pages, text: pages.flatMap(p => p.lines).join("\n") };
}
