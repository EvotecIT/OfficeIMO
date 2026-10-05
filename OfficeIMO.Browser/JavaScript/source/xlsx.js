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
