// Match OfficeIMO.CSV's ASCII-space/trigger rule.
function csvField(value, delimiter, protect) {
  let text;
  if (value == null) text = "";
  else if (value instanceof Date) text = Number.isFinite(value.getTime()) ? value.toISOString() : "";
  else if (typeof value === "boolean") text = value ? "True" : "False";
  else if (typeof value === "number") text = Number.isFinite(value) ? String(value) : "";
  else if (typeof value === "string") text = value;
  else throw new TypeError("CSV cells must be strings, numbers, booleans, Dates or null.");
  if (protect && typeof value === "string" && /^ *[=+\-@\t\r\n]/.test(text)) text = "'" + text;
  return text.includes(delimiter) || /["\r\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
}

async function writeCsv(rows, options = {}) {
  const columns = copyColumns(options.columns);
  const delimiter = options.delimiter ?? ",", lineEnding = options.lineEnding ?? "\r\n";
  if (![",", ";", "\t"].includes(delimiter)) throw new RangeError("Delimiter must be comma, semicolon or tab.");
  if (!["\r\n", "\n", "\r"].includes(lineEnding)) throw new RangeError("Invalid line ending.");
  const signal = options.signal, protect = options.formulaInjectionProtection !== false;
  checkAbort(signal);
  const parts = options.bom ? [new Uint8Array([239, 187, 191])] : [];
  const buffer = textChunks(bytes => { parts.push(bytes); }, signal);
  let count = 0;
  async function record(values) {
    for (let i = 0; i < columns.length; i++) {
      if (buffer.append((i ? delimiter : "") + csvField(values[i], delimiter, protect))) await buffer.flush();
    }
    if (buffer.append(lineEnding)) { await buffer.flush(); options.onProgress?.({ phase: "rows", rows: count }); }
  }
  if (options.includeHeader !== false && columns.length) await record(columns.map(c => c.header));
  for await (const row of inputRows(rows, signal)) {
    await record(rowValues(row, columns));
    count++;
  }
  await buffer.flush();
  const blob = new Blob(parts, { type: "text/csv;charset=utf-8" });
  options.onProgress?.({ phase: "complete", rows: count, bytes: blob.size });
  checkAbort(signal);
  return blob;
}
