// Independent narrow fixture inspector: byte offsets, stream lengths, zlib, font Unicode maps and page content.
// General PDF interoperability is separately checked by OfficeIMO.Pdf, Poppler and MuPDF.
import assert from "node:assert/strict";
import { inflateSync } from "node:zlib";

export async function inspectPdf(blob) {
  const bytes = Buffer.from(await blob.arrayBuffer()), text = bytes.toString("latin1");
  assert.ok(text.startsWith("%PDF-1.7\n"));
  const start = Number(/startxref\n(\d+)\n%%EOF\n$/.exec(text)?.[1]);
  assert.equal(text.slice(start, start + 5), "xref\n");
  const xref = text.slice(start).split("\n"), count = Number(xref[1].split(" ")[1]), objects = new Map();
  for (let id = 1; id < count; id++) {
    const offset = Number(xref[id + 2].slice(0, 10));
    assert.equal(text.slice(offset, offset + String(id).length + 6), id + " 0 obj");
    const end = text.indexOf("\nendobj\n", offset);
    assert.ok(end > offset);
    const body = text.slice(offset + String(id).length + 7, end), marker = body.indexOf("\nstream\n");
    let stream;
    if (marker >= 0) {
      const length = Number(/\/Length (\d+)/.exec(body)[1]), begin = offset + String(id).length + 7 + marker + 8;
      stream = bytes.subarray(begin, begin + length);
      assert.equal(text.slice(begin + length, begin + length + 10), "\nendstream");
      if (/\/Filter \/FlateDecode/.test(body.slice(0, marker))) stream = inflateSync(stream);
    }
    objects.set(id, { body, stream });
  }
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
