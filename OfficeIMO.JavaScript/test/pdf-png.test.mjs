import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { constants, deflateSync } from "node:zlib";
import { writePdfDocument, writePdfDocumentTo } from "../dist/pdf/index.js";
import { Crc32 } from "../dist/zip/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const signature = Buffer.from([137,80,78,71,13,10,26,10]);
const colors = [[240,20,20], [20,230,20], [20,20,235], [235,220,20]], alphas = [255,128,0,64];
function chunk(type, data = Buffer.alloc(0)) {
  const name = Buffer.from(type), crc = new Crc32(); crc.update(name); crc.update(data);
  const size = Buffer.alloc(4), checksum = Buffer.alloc(4); size.writeUInt32BE(data.length); checksum.writeUInt32BE(crc.value);
  return Buffer.concat([size, name, data, checksum]);
}
function png({ width = 1, height = 1, depth = 8, color = 2, interlace = 0, scanlines = Buffer.from([0,10,20,30]), before = [], after = [], split = false, compression = {}, trailing = Buffer.alloc(0) } = {}) {
  const header = Buffer.alloc(13); header.writeUInt32BE(width); header.writeUInt32BE(height, 4); header[8] = depth; header[9] = color; header[12] = interlace;
  const data = Buffer.concat([deflateSync(scanlines, compression), trailing]), idat = split ? [chunk("IDAT"), ...[...data].map(byte => chunk("IDAT", Buffer.from([byte]))), chunk("IDAT")] : [chunk("IDAT", data)];
  return Buffer.concat([signature, chunk("IHDR", header), ...before, ...idat, ...after, chunk("IEND")]);
}
const imageBlock = data => ({ kind: "image", image: { format: "png", data } });
const fixture = file => readFile(new URL("./fixtures/pdf/" + file, import.meta.url));
const images = pdf => [...pdf.objects.values()].filter(object => /\/Subtype \/Image/.test(object.body));

test("PDF PNG intake preserves independent Pillow RGB and straight RGBA samples and reuses resources", async () => {
  for (const format of ["rgb", "rgba"]) {
    // Pillow 11.3.0, 12x8 four quadrants, compress_level=9; no color/orientation metadata.
    const data = new Uint8Array(await fixture(format + "-pillow.png")), block = imageBlock(data);
    const pdf = await inspectPdf(await writePdfDocument([block, block], { compression: false, limits: { maxImageBytes: data.length } }));
    const resources = images(pdf), rgb = resources.find(object => /DeviceRGB/.test(object.body));
    assert.equal(resources.length, format === "rgba" ? 2 : 1);
    assert.match(rgb.body, /\/Width 12 \/Height 8/);
    const expected = [], alpha = [];
    for (let y = 0; y < 8; y++) for (let x = 0; x < 12; x++) { const q = Math.floor(y / 4) * 2 + Math.floor(x / 6); expected.push(...colors[q]); alpha.push(alphas[q]); }
    assert.deepEqual([...rgb.stream], expected);
    if (format === "rgba") assert.deepEqual([...resources.find(object => /DeviceGray/.test(object.body)).stream], alpha);
    await assert.rejects(writePdfDocument([block, imageBlock(data)], { limits: { maxImageBytes: data.length } }), /maxImageBytes/);
  }
  const split = await inspectPdf(await writePdfDocument([imageBlock(png({ split: true }))], { compression: false }));
  assert.deepEqual([...images(split)[0].stream], [10,20,30], "IDAT boundaries may split the zlib header, payload and checksum");
});

test("PDF PNG intake reverses all five filters with independent Pillow qualification", async () => {
  // The reproducible fixture uses filter types 0..4 on its five rows; Pillow independently checks all decoded samples.
  const pdf = await inspectPdf(await writePdfDocument([imageBlock(new Uint8Array(await fixture("rgba-filters.png")))], { compression: false }));
  const resources = images(pdf), expected = [], alpha = [];
  for (let y = 0; y < 5; y++) for (let x = 0; x < 4; x++) {
    expected.push((x * 67 + y * 29) % 256, (x * 13 + y * 83) % 256, (x * 113 + y * 7) % 256); alpha.push((x * 51 + y * 37) % 256);
  }
  assert.deepEqual([...resources.find(object => /DeviceRGB/.test(object.body)).stream], expected);
  assert.deepEqual([...resources.find(object => /DeviceGray/.test(object.body)).stream], alpha);
});

test("PDF PNG intake rejects corrupt structure, unsupported color semantics and unsupported profiles", async () => {
  const valid = png(), corrupt = valid.slice(); corrupt[corrupt.length - 1] ^= 1;
  const cases = [
    [corrupt, /CRC/], [valid.subarray(0, valid.length - 3), /Truncated/], [Buffer.concat([valid, Buffer.from([0])]), /IEND/],
    [png({ after: [chunk("tEXt"), chunk("IDAT")] }), /consecutive/],
    [png({ before: [chunk("IHDR", Buffer.alloc(13))] }), /IHDR/],
    [png({ scanlines: Buffer.from([5,1,2,3]) }), /filter/]
  ];
  for (const [data, error] of cases) await assert.rejects(writePdfDocument([imageBlock(data)]), error);
  for (const type of ["sRGB", "gAMA", "cHRM", "iCCP", "cICP", "mDCV", "cLLI", "sBIT", "PLTE", "tRNS", "eXIf", "acTL", "fcTL", "fdAT", "ABCD"])
    await assert.rejects(writePdfDocument([imageBlock(png({ before: [chunk(type)] }))]), error => error.code === "NOT_SUPPORTED" && error.message.includes(type));
  for (const options of [{ depth: 16 }, { color: 0 }, { color: 3 }, { color: 4 }, { interlace: 1 }])
    await assert.rejects(writePdfDocument([imageBlock(png(options))]), error => error.code === "NOT_SUPPORTED");
});

test("PDF PNG intake bounds source, IHDR dimensions and exact inflated bytes", async () => {
  const data = png();
  await assert.rejects(writePdfDocument([imageBlock(data)], { limits: { maxImageBytes: data.length - 1 } }), error => error.code === "RESOURCE_LIMIT");
  await assert.rejects(writePdfDocument([imageBlock(png({ width: 0x7fffffff, height: 0x7fffffff }))]), error => error.code === "RESOURCE_LIMIT" && /maxImagePixels/.test(error.message));
  await assert.rejects(writePdfDocument([imageBlock(png({ width: 2, height: 2 }))], { limits: { maxImagePixels: 3 } }), /maxImagePixels/);
  // Tiny 1-pixel IHDR paired with four MiB of highly compressible data must stop at the first oversized output chunk.
  for (const scanlines of [Buffer.alloc(4 * 1024 * 1024), Buffer.from([0,1,2])])
    await assert.rejects(writePdfDocument([imageBlock(png({ scanlines }))]), /exact IHDR scanline length/);
  const damagedZlib = png({ scanlines: Buffer.from([0,1,2,3]) });
  const idat = 33, size = damagedZlib.readUInt32BE(idat), payload = damagedZlib.subarray(idat + 8, idat + 8 + size).slice(); payload[payload.length - 1] ^= 1;
  const bad = Buffer.concat([damagedZlib.subarray(0, idat), chunk("IDAT", payload), chunk("IEND")]);
  await assert.rejects(writePdfDocument([imageBlock(bad)]), /PNG zlib image data/);
  const extra = Buffer.concat([data.subarray(41, 41 + data.readUInt32BE(33)), Buffer.from([0])]);
  await assert.rejects(writePdfDocument([imageBlock(Buffer.concat([data.subarray(0, 33), chunk("IDAT", extra), chunk("IEND")]))]), /PNG zlib image data/);
});

test("PDF PNG intake rejects trailing zlib bytes for stored, fixed and dynamic blocks", async () => {
  const scanlines = Buffer.alloc(6145);
  for (let i = 1; i < scanlines.length; i++) scanlines[i] = i % 13 === 0 ? i % 251 : 42;
  for (const [compression, blockType] of [[{ level: 0 }, 0], [{ strategy: constants.Z_FIXED }, 1], [{}, 2]]) {
    const options = { width: 2048, scanlines, compression };
    const compressed = deflateSync(scanlines, compression);
    assert.equal((compressed[2] >> 1) & 3, blockType, "the fixture exercises the specified DEFLATE block type");
    for (const split of [false, true]) {
      const pdf = await inspectPdf(await writePdfDocument([imageBlock(png({ ...options, split }))], { compression: false }));
      assert.deepEqual([...images(pdf)[0].stream], [...scanlines.subarray(1)]);
      for (const trailing of [Buffer.from([0]), compressed.subarray(-4), deflateSync(Buffer.alloc(0))])
        await assert.rejects(writePdfDocument([imageBlock(png({ ...options, split, trailing }))]), /PNG zlib image data/);
    }
  }
});

test("PDF PNG boundary verification preserves the abort reason", async () => {
  const native = globalThis.DecompressionStream, abort = new AbortController(), reason = Error("PNG boundary verification abort");
  let inflaters = 0;
  try {
    globalThis.DecompressionStream = class extends native {
      constructor(format) { super(format); if (++inflaters === 2) abort.abort(reason); }
    };
    await assert.rejects(writePdfDocument([imageBlock(png())], { signal: abort.signal }), error => error === reason);
    assert.equal(inflaters, 2);
  } finally { globalThis.DecompressionStream = native; }
});

test("PDF PNG intake reports unavailable native inflation and cancels during pixel reconstruction", async () => {
  const native = globalThis.DecompressionStream;
  try {
    globalThis.DecompressionStream = undefined;
    await assert.rejects(writePdfDocument([imageBlock(png())]), error => error.code === "PLATFORM_UNAVAILABLE");
  } finally { globalThis.DecompressionStream = native; }
  const now = Object.getOwnPropertyDescriptor(performance, "now"), scheduler = globalThis.scheduler, abort = new AbortController();
  let inflating = false, returned = 0;
  const destination = new WritableStream({ write() { throw Error("No page should be emitted"); } });
  const blocks = { *[Symbol.iterator]() { try { yield imageBlock(png({ width: 512, height: 256, color: 6, scanlines: Buffer.alloc((512 * 4 + 1) * 256) })); } finally { returned++; } } };
  try {
    let clock = 0; Object.defineProperty(performance, "now", { configurable: true, value: () => clock += 40 });
    globalThis.DecompressionStream = class extends native { constructor(format) { super(format); inflating = true; } };
    globalThis.scheduler = { postTask: async callback => { if (inflating) abort.abort(Error("PNG reconstruction abort")); callback(); } };
    await assert.rejects(writePdfDocumentTo(blocks, destination, { signal: abort.signal, pageNumbers: false }), /PNG reconstruction abort/);
    assert.equal(returned, 1); assert.equal(destination.locked, false);
  } finally {
    if (now) Object.defineProperty(performance, "now", now); else delete performance.now;
    globalThis.scheduler = scheduler; globalThis.DecompressionStream = native;
  }
});
