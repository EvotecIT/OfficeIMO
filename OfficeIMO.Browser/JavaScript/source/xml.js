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
  return text;
}

function inlineText(value) {
  const text = cellText(value);
  const node = part => '<t xml:space="preserve">' + xml(part) + '</t>';
  // Escape sequences are decoded per text run. Split their initial underscore so
  // consumers with and without an OOXML escape decoder read the same literal text.
  if (/_x[0-9a-f]{4}_/i.test(text))
    return text.split(/(?<=_)(?=x[0-9a-f]{4}_)/gi).map(part => '<r>' + node(part) + '</r>').join("");
  return node(text);
}

function columnName(index) {
  let name = "";
  while (index) { index--; name = String.fromCharCode(65 + index % 26) + name; index = Math.floor(index / 26); }
  return name;
}

function clipName(text, length) { return text.slice(0, length).replace(/[\ud800-\udbff]$/, ""); }

function sheetName(requested, names) {
  if (typeof requested !== "string") throw new TypeError("Sheet name must be a string.");
  let base = cleanXml(requested).replace(/[\[\]:*?/\\]/g, "_").trim().replace(/^'+|'+$/g, "").trim();
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
