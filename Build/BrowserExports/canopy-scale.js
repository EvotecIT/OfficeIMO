// Native Canopy pages are generated on demand; the adapter writes to an acknowledged 64 KiB file bridge.
globalThis.runCanopyScale = async function ({ rows, columns, format, navigation, slowSink, fallback }) {
  const specs = Array.from({ length: columns }, (_, c) => ({ id: 'c' + c, title: 'Column ' + c,
    kind: c % 4 === 0 ? 'number' : c % 4 === 3 ? 'datetime' : 'text', width: 160 }));
  const view = { protocol: 'canopyx/2', type: 'grid', id: 'canopy-scale', label: 'Paged records', source: { kind: 'host', navigation },
    rowHeight: 28, pageSize: 256, selection: 'none', presentation: 'paged', pageLength: 5, reporting: {}, columns: specs };
  let pages = 0, maxPageRows = 0, consumed = 0, bytes = 0, writes = 0, firstByteRows = null;
  const source = { navigation, async query(query, { signal }) {
    if (signal.aborted) throw signal.reason;
    pages++; const first = navigation === 'cursor' ? Number(query.cursor || 0) : query.offset;
    const items = Array.from({ length: Math.min(query.limit, rows - first) }, (_, i) => {
      const r = first + i, cells = Object.create(null);
      for (let c = 0; c < columns; c++) cells['c' + c] = c % 4 === 0 ? r * columns + c : c % 4 === 1 ? 'Unique ' + r + ': Łódź🧪' :
        c % 4 === 2 ? r % 2 === 0 : '2026-01-' + String(1 + r % 28).padStart(2, '0') + 'T00:00:00.000Z';
      return { id: String(r), cells };
    });
    maxPageRows = Math.max(maxPageRows, items.length);
    return { protocol: 'canopyx/2', type: navigation === 'cursor' ? 'gridCursorPage' : 'gridPage', view: query.view,
      revision: 'scale-r1', offset: first, total: navigation === 'cursor' ? null : rows, matched: navigation === 'cursor' ? null : rows,
      items, ...(navigation === 'cursor' && first + items.length < rows ? { nextCursor: String(first + items.length) } : {}) };
  } };
  const grid = await canopyMount(view, { source });
  const original = grid.prepareExport({ scope: 'all', values: 'raw', timeZone: 'utc' });
  const capture = { request: original.request, async *rows(options) { for await (const row of original.rows(options)) { consumed++; yield row; } } };
  let buffer = new Uint8Array(65536), used = 0;
  async function flush() {
    if (!used) return;
    let binary = ''; for (let i = 0; i < used; i += 8192) binary += String.fromCharCode(...buffer.subarray(i, Math.min(i + 8192, used)));
    await acceptCanopyScaleChunk(btoa(binary)); used = 0;
    if (slowSink) await new Promise(resolve => setTimeout(resolve, 1));
  }
  const compression = globalThis.CompressionStream;
  if (fallback) globalThis.CompressionStream = undefined;
  const started = performance.now();
  try {
    const result = await OfficeIMO.writeCanopyTo(capture, format, { async write(chunk) {
      firstByteRows ??= consumed; bytes += chunk.length; writes++;
      for (let offset = 0; offset < chunk.length;) {
        const size = Math.min(chunk.length - offset, buffer.length - used); buffer.set(chunk.subarray(offset, offset + size), used); used += size; offset += size;
        if (used === buffer.length) await flush();
      }
    } }, { xlsx: { sheet: { autoSize: false } }, limits: { maxRows: rows, maxCells: (rows + 1) * columns } });
    await flush();
    canopyAssert(result.rows === rows && consumed === rows && maxPageRows <= 256, 'Native page bound or final row count differs.');
    return { rows, columns, format, navigation, pages, maxPageRows, firstByteRows, bytes, writes, elapsedMs: performance.now() - started,
      sourceBookkeeping: 'CanopyX retains seen IDs and cursor tokens; output bridge retains at most 64 KiB' };
  } finally { globalThis.CompressionStream = compression; grid.destroy(); }
};
