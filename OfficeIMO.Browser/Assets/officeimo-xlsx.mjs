// Generated: Build/assets.mjs. MIT.
const utf8 = new TextEncoder();

function checkAbort(signal) {
  if (signal?.aborted) throw signal.reason ?? new DOMException("Export cancelled.", "AbortError");
}

function withAbort(promise, signal) {
  if (!signal) return promise;
  checkAbort(signal);
  return new Promise((resolve, reject) => {
    const abort = () => reject(signal.reason ?? new DOMException("Export cancelled.", "AbortError"));
    signal.addEventListener("abort", abort, { once: true });
    Promise.resolve(promise).then(resolve, reject).finally(() => signal.removeEventListener("abort", abort));
  });
}

async function* inputRows(input, signal) {
  checkAbort(signal);
  const iterator = input?.[Symbol.asyncIterator]?.() ?? input?.[Symbol.iterator]?.();
  if (!iterator) throw new TypeError("Rows must be a synchronous or asynchronous iterable.");
  let done = false;
  try {
    while (true) {
      checkAbort(signal);
      const next = iterator.next();
      const item = next?.then ? await withAbort(next, signal) : next;
      checkAbort(signal);
      if (item.done) { done = true; return; }
      yield item.value;
    }
  } finally {
    if (!done && iterator.return) {
      const returned = iterator.return();
      // I/O-bound producers must also observe the signal.
      if (signal?.aborted) Promise.resolve(returned).catch(() => {});
      else await returned;
    }
  }
}

function rowValues(row, columns) {
  if (Array.isArray(row)) {
    if (row.length > columns.length) throw new RangeError("Row has more values than declared columns.");
    return row;
  }
  if (!row || typeof row !== "object" || row instanceof Date) throw new TypeError("A row must be an array or object.");
  return columns.map(c => row[c.key ?? c.header]);
}

function copyColumns(columns) {
  if (!Array.isArray(columns)) throw new TypeError("Declare the columns in export order.");
  return columns.map(c => {
    if (!c || typeof c.header !== "string" || (c.key !== undefined && typeof c.key !== "string"))
      throw new TypeError("Each column needs a string header and an optional string key.");
    return { ...c };
  });
}

function pause() { return new Promise(resolve => setTimeout(resolve, 0)); }

function textChunks(write, signal) {
  let text = "", deadline = performance.now() + 8;
  async function flush() {
    checkAbort(signal);
    if (text) { const bytes = utf8.encode(text); text = ""; await write(bytes); }
    if (performance.now() >= deadline) { await pause(); deadline = performance.now() + 8; }
    checkAbort(signal);
  }
  return {
    append(value) { text += value; return text.length >= 32768 || performance.now() >= deadline; },
    flush
  };
}

function saveBlob(blob, fileName) {
  if (!(blob instanceof Blob) || typeof fileName !== "string" || !fileName) throw new TypeError("A Blob and file name are required.");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url; link.download = fileName;
  document.body.append(link);
  try { link.click(); } finally {
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }
}

const spreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const relationshipsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
const officeRelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const xmlDeclaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function cleanXml(value) {
  // Unicode mode preserves valid surrogate pairs while removing lone surrogates.
  return String(value).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff\ufffe\uffff]/gu, "");
}

function xml(value) {
  return cleanXml(value).replace(/[&<>"'\r\n\t]/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;",
    "\r": "&#13;", "\n": "&#10;", "\t": "&#9;"
  })[c]);
}

function cellText(value) {
  const text = cleanXml(value);
  if (text.length > 32767) throw new RangeError("Excel cell text exceeds 32,767 UTF-16 code units.");
  // Protect literal OOXML escape sequences from Excel's string decoder.
  return xml(text.replace(/_x[0-9a-f]{4}_/gi, match => "_x005F_" + match.slice(1)));
}

function columnName(index) {
  let name = "";
  while (index) { index--; name = String.fromCharCode(65 + index % 26) + name; index = Math.floor(index / 26); }
  return name;
}

function clipName(text, length) { return text.slice(0, length).replace(/[\ud800-\udbff]$/, ""); }

function sheetName(requested, names) {
  if (typeof requested !== "string") throw new TypeError("Sheet name must be a string.");
  let base = cleanXml(requested).replace(/[\[\]:*?/\\]/g, "_").trim().replace(/^'+|'+$/g, "");
  if (!base) base = "Sheet";
  if (base.toLowerCase() === "history") base += "_";
  base = clipName(base, 31);
  let name = base, suffix = 2;
  while (names.has(name.toLowerCase())) {
    const tail = " (" + suffix++ + ")";
    name = clipName(base, 31 - tail.length) + tail;
  }
  names.add(name.toLowerCase());
  return name;
}

function excelDate(date, mode) {
  if (!Number.isFinite(date.getTime())) return null;
  const p = mode === "utc" ? "getUTC" : "get";
  const year = date[p + "FullYear"]();
  if (year < 1900 || year > 9999) throw new RangeError("Excel dates must be in years 1900 through 9999.");
  const wall = new Date(0);
  wall.setUTCFullYear(year, date[p + "Month"](), date[p + "Date"]());
  wall.setUTCHours(date[p + "Hours"](), date[p + "Minutes"](), date[p + "Seconds"](), date[p + "Milliseconds"]());
  const time = wall.getTime();
  return (time - Date.UTC(1899, 11, 31)) / 86400000 + (time >= Date.UTC(1900, 2, 1) ? 1 : 0);
}

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

function styleRegistry() {
  const formats = new Map(), fills = new Map(), indexes = new Map(), styles = [];
  function add(column = {}, header = false, fill, date = false) {
    const format = column.format ?? (column.type === "date" || date ? "yyyy-mm-dd hh:mm:ss" : "");
    if (typeof format !== "string") throw new TypeError("Column format must be a string.");
    const key = JSON.stringify([format, !!column.wrapText, column.alignment ?? "", header, fill ?? ""]);
    if (indexes.has(key)) return indexes.get(key);
    if (styles.length >= 64000) throw new RangeError("Workbook exceeds Excel's cell style limit.");
    let numFmtId = 0, fillId = 0;
    if (format) {
      if (!formats.has(format)) formats.set(format, formats.size + 164);
      numFmtId = formats.get(format);
    }
    if (fill) {
      if (!fills.has(fill)) fills.set(fill, fills.size + 2);
      fillId = fills.get(fill);
    }
    const alignment = column.wrapText || column.alignment
      ? '<alignment' + (column.wrapText ? ' wrapText="1"' : "") +
        (column.alignment ? ' horizontal="' + xml(column.alignment) + '"' : "") + '/>' : "";
    styles.push('<xf numFmtId="' + numFmtId + '" fontId="' + (header ? 1 : 0) + '" fillId="' + fillId +
      '" borderId="0" xfId="0"' + (numFmtId ? ' applyNumberFormat="1"' : "") +
      (header ? ' applyFont="1"' : "") + (fillId ? ' applyFill="1"' : "") +
      (alignment ? ' applyAlignment="1"' : "") + '>' + alignment + '</xf>');
    const index = styles.length - 1;
    indexes.set(key, index);
    return index;
  }
  add();
  return {
    add,
    xml() {
      return xmlDeclaration + '<styleSheet xmlns="' + spreadsheetNs + '">' +
        '<numFmts count="' + formats.size + '">' + [...formats].map(([format, id]) =>
          '<numFmt numFmtId="' + id + '" formatCode="' + xml(format) + '"/>').join("") + '</numFmts>' +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
        '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="' + (fills.size + 2) + '"><fill><patternFill patternType="none"/></fill>' +
        '<fill><patternFill patternType="gray125"/></fill>' + [...fills.keys()].map(fill =>
          '<fill><patternFill patternType="solid"><fgColor rgb="' + fill +
          '"/><bgColor indexed="64"/></patternFill></fill>').join("") + '</fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="' + styles.length + '">' + styles.join("") + '</cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    }
  };
}

function createWorkbook(options = {}) {
  const settings = { ...options };
  const dateMode = settings.dateMode ?? "local", compression = settings.compression ?? "auto";
  if (!["local", "utc"].includes(dateMode)) throw new RangeError("dateMode must be local or utc.");
  if (!["auto", "store"].includes(compression)) throw new RangeError("compression must be auto or store.");
  const created = settings.created ?? new Date(), modified = settings.modified ?? created;
  if (!(created instanceof Date) || !(modified instanceof Date) ||
      !Number.isFinite(created.getTime()) || !Number.isFinite(modified.getTime())) throw new TypeError("Invalid workbook property date.");
  settings.created = created.toISOString(); settings.modified = modified.toISOString();
  const signal = settings.signal, styles = styleRegistry(), names = new Set(), sheets = [];
  let state = "open", result;
  function open() {
    checkAbort(signal);
    if (state !== "open") throw new Error("Workbook is already finalized.");
  }
  function addSheet(requested, sheetOptions = {}) {
    open();
    if (sheets.length >= 65527) throw new RangeError("Workbook has too many sheets.");
    const columns = copyColumns(sheetOptions.columns ?? []), opts = { ...sheetOptions };
    if (columns.length > 16384) throw new RangeError("Excel supports at most 16,384 columns.");
    if (opts.includeHeader === false && (opts.freezeHeader || opts.autoFilter)) throw new TypeError("A frozen header or autofilter requires a header row.");
    let fill = opts.headerFill;
    if (fill !== undefined) {
      if (typeof fill !== "string" || !/^#?(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(fill)) throw new TypeError("headerFill must be RGB or ARGB hex.");
      fill = fill.replace(/^#/, "").toUpperCase();
      if (fill.length === 6) fill = "FF" + fill;
    }
    for (const column of columns) {
      if (column.width !== undefined && (!Number.isFinite(column.width) || column.width < 0 || column.width > 255))
        throw new RangeError("Column width must be from 0 through 255 characters.");
      if (column.type !== undefined && !["string", "number", "boolean", "date"].includes(column.type)) throw new TypeError("Invalid column type.");
      if (column.alignment !== undefined && !["left", "center", "right", "fill", "justify", "distributed"].includes(column.alignment))
        throw new TypeError("Invalid horizontal alignment.");
      // Validate header text before allocating compressor resources.
      cellText(column.header);
    }
    const name = sheetName(requested, names);
    const declared = columns.map((column, index) => ({
      column, letter: columnName(index + 1), style: styles.add(column), dateStyle: styles.add(column, false, undefined, true),
      headerStyle: styles.add({ wrapText: column.wrapText, alignment: column.alignment }, opts.boldHeader !== false, fill)
    }));
    // Allocate the compressor lazily so unused/discarded workbooks hold no native stream.
    let entry, buffer, started = false, busy = false, error, count = 0;
    const headerRows = columns.length && opts.includeHeader !== false ? 1 : 0;
    function cell(value, index, row, header = false) {
      const col = declared[index], type = value instanceof Date ? "date" : typeof value;
      if (!header && value != null && col.column.type && col.column.type !== type)
        throw new TypeError("Cell " + col.letter + row + " does not match column type " + col.column.type + ".");
      const style = header ? col.headerStyle : type === "date" ? col.dateStyle : col.style;
      const prefix = '<c r="' + col.letter + row + '" s="' + style + '"';
      if (value == null || (type === "number" && !Number.isFinite(value))) return prefix + '/>';
      if (type === "date") value = excelDate(value, dateMode);
      if (value == null) return prefix + '/>';
      if (type === "string") return prefix + ' t="inlineStr"><is><t xml:space="preserve">' + cellText(value) + '</t></is></c>';
      if (type === "boolean") return prefix + ' t="b"><v>' + (value ? 1 : 0) + '</v></c>';
      if (type !== "number" && type !== "date") throw new TypeError("Excel cells must be strings, numbers, booleans, Dates or null.");
      return prefix + '><v>' + value + '</v></c>';
    }
    async function writeRow(values, number, header) {
      if (buffer.append('<row r="' + number + '">')) await buffer.flush();
      for (let i = 0; i < columns.length; i++) if (buffer.append(cell(values[i], i, number, header))) await buffer.flush();
      if (buffer.append('</row>')) await buffer.flush();
    }
    async function start() {
      if (started) return;
      started = true;
      entry = zipEntry(compression, signal);
      buffer = textChunks(bytes => entry.write(bytes), signal);
      buffer.append(xmlDeclaration + '<worksheet xmlns="' + spreadsheetNs + '"><sheetViews><sheetView workbookViewId="0">' +
        (opts.freezeHeader && headerRows ? '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
          '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>' : "") + '</sheetView></sheetViews>');
      if (columns.some(c => c.width !== undefined)) {
        buffer.append('<cols>');
        for (let i = 0; i < columns.length; i++) {
          const column = columns[i];
          if (column.width !== undefined && buffer.append('<col min="' + (i + 1) + '" max="' + (i + 1) +
            '" width="' + column.width + '" customWidth="1"/>')) await buffer.flush();
        }
        buffer.append('</cols>');
      }
      buffer.append('<sheetData>');
      if (headerRows) await writeRow(columns.map(c => c.header), 1, true);
    }
    function progress() { settings.onProgress?.({ phase: "rows", rows: count, sheetName: name }); }
    const internalSheet = {
      get busy() { return busy; },
      get rowCount() { return count; },
      async finish() {
        if (error) throw error;
        try {
          await start();
          buffer.append('</sheetData>');
          if (opts.autoFilter && headerRows) buffer.append('<autoFilter ref="A1:' + columnName(columns.length) + (count + 1) + '"/>');
          buffer.append('</worksheet>');
          await buffer.flush();
          return await entry.close();
        } catch (failure) { await entry?.discard(failure); throw failure; }
      },
      async discard(failure) { await entry?.discard(failure); }
    };
    sheets.push({ name, sheet: internalSheet });
    return Object.freeze({
      name,
      async addRows(rows) {
        open();
        if (busy) throw new Error("Await the current addRows call before writing more rows to this sheet.");
        if (error) throw error;
        busy = true;
        try {
          await start();
          let checkpoint = performance.now();
          for await (const row of inputRows(rows, signal)) {
            if (count + headerRows >= 1048576) throw new RangeError("Excel supports at most 1,048,576 rows including the header; split the sheet.");
            if (!columns.length) throw new RangeError("Declare columns before adding rows.");
            await writeRow(rowValues(row, columns), count + headerRows + 1, false);
            count++;
            if (performance.now() - checkpoint >= 50) { progress(); checkpoint = performance.now(); }
          }
          await buffer.flush(); progress(); checkAbort(signal);
        } catch (failure) { error = failure; await entry?.discard(failure); throw failure; }
        finally { busy = false; }
      }
    });
  }
  function toBlob() {
    checkAbort(signal);
    if (result) return result;
    open();
    if (sheets.some(s => s.sheet.busy)) throw new Error("Await addRows before finalizing the workbook.");
    if (!sheets.length) addSheet("Sheet1");
    state = "finalizing";
    result = (async () => {
      const entries = [];
      async function part(name, text) {
        const entry = zipEntry(compression, signal);
        try {
          // Static parts can be large for wide/multi-sheet workbooks; use the same chunk/yield contract.
          const buffer = textChunks(bytes => entry.write(bytes), signal);
          for (let i = 0; i < text.length;) {
            // Do not split a surrogate pair at a UTF-8 encoding boundary.
            let end = Math.min(text.length, i + 16384);
            if (end < text.length && /[\ud800-\udbff]/.test(text[end - 1])) end--;
            buffer.append(text.slice(i, end)); await buffer.flush(); i = end;
          }
          await buffer.flush();
          entries.push({ name, entry: await entry.close() });
        } catch (failure) { await entry.discard(failure); throw failure; }
      }
      try {
        for (let i = 0; i < sheets.length; i++) entries.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', entry: await sheets[i].sheet.finish() });
        const rel = (id, type, target) => '<Relationship Id="' + id + '" Type="' + officeRelationshipsNs + '/' + type + '" Target="' + target + '"/>';
        await part('xl/workbook.xml', xmlDeclaration + '<workbook xmlns="' + spreadsheetNs + '" xmlns:r="' + officeRelationshipsNs +
          '"><workbookPr date1904="0"/><bookViews><workbookView/></bookViews><sheets>' + sheets.map((s, i) =>
            '<sheet name="' + xml(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join("") + '</sheets></workbook>');
        await part('xl/_rels/workbook.xml.rels', xmlDeclaration + '<Relationships xmlns="' + relationshipsNs + '">' +
          sheets.map((_, i) => rel('rId' + (i + 1), 'worksheet', 'worksheets/sheet' + (i + 1) + '.xml')).join("") +
          rel('styles', 'styles', 'styles.xml') + '</Relationships>');
        await part('xl/styles.xml', styles.xml());
        await part('_rels/.rels', xmlDeclaration + '<Relationships xmlns="' + relationshipsNs + '">' + rel('workbook', 'officeDocument', 'xl/workbook.xml') +
          '<Relationship Id="core" Type="' + relationshipsNs + '/metadata/core-properties" Target="docProps/core.xml"/></Relationships>');
        await part('docProps/core.xml', xmlDeclaration + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
          ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
          '<dc:creator>' + xml(settings.creator ?? "") + '</dc:creator><dc:title>' + xml(settings.title ?? "") + '</dc:title>' +
          '<dcterms:created xsi:type="dcterms:W3CDTF">' + settings.created + '</dcterms:created>' +
          '<dcterms:modified xsi:type="dcterms:W3CDTF">' + settings.modified + '</dcterms:modified></cp:coreProperties>');
        const override = (name, type) => '<Override PartName="/' + name + '" ContentType="application/vnd.openxmlformats-officedocument.' + type + '"/>';
        await part('[Content_Types].xml', xmlDeclaration + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          override('xl/workbook.xml', 'spreadsheetml.sheet.main+xml') + override('xl/styles.xml', 'spreadsheetml.styles+xml') +
          sheets.map((_, i) => override('xl/worksheets/sheet' + (i + 1) + '.xml', 'spreadsheetml.worksheet+xml')).join("") +
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>');
        checkAbort(signal);
        const blob = zipArchive(entries);
        settings.onProgress?.({ phase: "complete", rows: sheets.reduce((sum, s) => sum + s.sheet.rowCount, 0), bytes: blob.size });
        checkAbort(signal);
        state = "complete";
        return blob;
      } catch (failure) {
        state = "failed";
        for (const sheet of sheets) await sheet.sheet.discard(failure);
        throw failure;
      }
    })();
    return result;
  }
  return Object.freeze({ addSheet, toBlob });
}

export { createWorkbook, saveBlob };
