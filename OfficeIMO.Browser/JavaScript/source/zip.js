const crcTable = new Uint32Array(256).map((_, index) => {
  let value = index;
  for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  return value >>> 0;
});
const zipLimit = 0xffffffff;

function zipSize(value) {
  if (value >= zipLimit) throw new RangeError("Export exceeds the classic ZIP 4 GiB limit; split the workbook.");
  return value;
}

// Each worksheet feeds one compressor while its readable side is drained concurrently.
// Stored mode retains the same bounded UTF-8 chunks without requiring Streams support.
function zipEntry(compression, signal) {
  checkAbort(signal);
  let stream;
  if (compression !== "store" && typeof CompressionStream === "function") {
    try { stream = new CompressionStream("deflate-raw"); } catch (error) {
      if (!(error instanceof TypeError)) throw error;
    }
  }
  let parts = [], crc = 0xffffffff, size = 0, compressedSize = 0, failure;
  const writer = stream?.writable.getWriter(), reader = stream?.readable.getReader();
  const abort = () => {
    failure ??= signal?.reason ?? new DOMException("Export cancelled.", "AbortError");
    parts = [];
    writer?.abort(failure).catch(() => {});
    reader?.cancel(failure).catch(() => {});
  };
  signal?.addEventListener("abort", abort, { once: true });
  const drain = reader ? (async () => {
    try {
      while (true) {
        const { value, done } = await reader.read();
        if (done) return;
        checkAbort(signal);
        compressedSize = zipSize(compressedSize + value.length);
        parts.push(value);
      }
    } catch (error) {
      failure = error;
      writer.abort(error).catch(() => {});
    }
  })() : Promise.resolve();
  function cleanup() { signal?.removeEventListener("abort", abort); }
  return {
    async write(bytes) {
      checkAbort(signal);
      if (failure) throw failure;
      size = zipSize(size + bytes.length);
      for (const byte of bytes) crc = crcTable[(crc ^ byte) & 255] ^ (crc >>> 8);
      if (writer) await writer.write(bytes);
      else { parts.push(bytes); compressedSize = size; }
    },
    async close() {
      try {
        checkAbort(signal);
        if (writer) await writer.close();
        await drain;
        if (failure) throw failure;
        checkAbort(signal);
        const blob = new Blob(parts);
        parts = [];
        return { blob, crc: (crc ^ 0xffffffff) >>> 0, size, compressedSize, method: stream ? 8 : 0 };
      } finally { cleanup(); }
    },
    async discard(error) {
      failure = error;
      parts = [];
      cleanup();
      if (writer) await writer.abort(error).catch(() => {});
      if (reader) await reader.cancel(error).catch(() => {});
      await drain;
    }
  };
}

function zipHeader(signature, length, name) {
  const bytes = new Uint8Array(length + name.length), data = new DataView(bytes.buffer);
  data.setUint32(0, signature, true);
  bytes.set(name, length);
  return { bytes, data };
}

function zipArchive(entries) {
  if (entries.length >= 65535) throw new RangeError("Too many ZIP entries.");
  const parts = [], central = [];
  let offset = 0;
  for (const { name, entry } of entries) {
    const encoded = utf8.encode(name);
    const local = zipHeader(0x04034b50, 30, encoded);
    local.data.setUint16(4, 20, true); local.data.setUint16(6, 0x800, true);
    local.data.setUint16(8, entry.method, true); local.data.setUint16(12, 33, true);
    local.data.setUint32(14, entry.crc, true); local.data.setUint32(18, entry.compressedSize, true);
    local.data.setUint32(22, entry.size, true); local.data.setUint16(26, encoded.length, true);
    const record = zipHeader(0x02014b50, 46, encoded);
    record.data.setUint16(4, 20, true); record.data.setUint16(6, 20, true);
    record.data.setUint16(8, 0x800, true); record.data.setUint16(10, entry.method, true);
    record.data.setUint16(14, 33, true); record.data.setUint32(16, entry.crc, true);
    record.data.setUint32(20, entry.compressedSize, true); record.data.setUint32(24, entry.size, true);
    record.data.setUint16(28, encoded.length, true); record.data.setUint32(42, offset, true);
    parts.push(local.bytes, entry.blob); central.push(record.bytes);
    offset = zipSize(offset + local.bytes.length + entry.compressedSize);
  }
  const directorySize = central.reduce((sum, bytes) => sum + bytes.length, 0);
  zipSize(offset + directorySize + 22);
  const end = zipHeader(0x06054b50, 22, new Uint8Array());
  end.data.setUint16(8, entries.length, true); end.data.setUint16(10, entries.length, true);
  end.data.setUint32(12, directorySize, true); end.data.setUint32(16, offset, true);
  return new Blob([...parts, ...central, end.bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
