// Opt-in public API adapters only. PowerForge owns every repetition and warmup.
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";
import { readFile } from "node:fs/promises";
import { PassThrough } from "node:stream";

export const packages = Object.freeze({
  papaparse: { package: "papaparse", version: "5.7.0", formats: ["csv"] },
  exceljs: { package: "exceljs", version: "4.4.0", formats: ["xlsx"], variant: "Document Workbook to complete buffer" },
  "exceljs-stream": { package: "exceljs", version: "4.4.0", formats: ["xlsx"], nodeOnly: true,
    variant: "Node WorkbookWriter, committed rows, no shared strings/styles, complete buffered output" },
  sheetjs: { package: "xlsx", version: "0.20.3", formats: ["xlsx"] },
  pdfkit: { package: "pdfkit", version: "0.20.2", formats: ["pdf"] },
  pdfmake: { package: "pdfmake", version: "0.3.12", formats: ["pdf"] },
  jspdf: { package: "jspdf", version: "4.2.1", formats: ["pdf"], plugin: { package: "jspdf-autotable", version: "5.0.9" } }
});

export const fixtures = Object.freeze([
  { name: "flat-csv-100000x6-buffer", format: "csv", rows: 100000, columns: 6 },
  { name: "flat-csv-100000x20-buffer", format: "csv", rows: 100000, columns: 20 },
  { name: "flat-xlsx-10000x6-buffer", format: "xlsx", rows: 10000, columns: 6 },
  { name: "flat-xlsx-100000x20-buffer", format: "xlsx", rows: 100000, columns: 20 },
  { name: "flat-pdf-10000x6-buffer", format: "pdf", rows: 10000, columns: 6 },
  { name: "flat-pdf-2000x20-buffer", format: "pdf", rows: 2000, columns: 20 },
  ...["csv", "xlsx", "pdf"].map(format => ({ name: `bootstrap-${format}-1000x6-buffer`, format, rows: 1000, columns: 6 }))
]);

// The separate managed reader implements this same published deterministic fixture.
export function valueAt(row, column, spec) {
  return column % 3 === 0 ? row * 10 + column
    : column % 3 === 1 ? spec.format === "pdf"
      ? `Łódź${spec.unique ? row : row % 100}-${column}`
      : `Łódź${spec.textProfile === "bmp" ? "" : " 🧪"} ${spec.unique ? "row-" + row : "group-" + row % 100}-column-${column}`
    : row % 10000 + column / 100;
}
function* rows(spec) { for (let row = 0; row < spec.rows; row++) yield Array.from({ length: spec.columns }, (_, column) => valueAt(row, column, spec)); }
const headings = spec => Array.from({ length: spec.columns }, (_, column) => (spec.format === "pdf" ? "Column" : "Column ") + (column + 1));
const width = spec => (1190.551 - 72) / spec.columns;
export function lineSpacing(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength); let units, ascent, descent;
  for (let index = 0; index < view.getUint16(4); index++) {
    const at = 12 + index * 16, offset = view.getUint32(at + 8);
    const tag = String.fromCharCode(...bytes.subarray(at, at + 4));
    if (tag === "head") units = view.getUint16(offset + 18);
    if (tag === "hhea") { ascent = view.getInt16(offset + 4); descent = view.getInt16(offset + 6); }
  }
  if (!units || !Number.isFinite(ascent - descent)) throw Error("Comparison font lacks TrueType line metrics.");
  const ratio = (ascent - descent) / units;
  return Math.max(1.3, ratio + .1) / ratio;
}

/** Import only the selected public package, through its exact isolated installation. */
export async function loadAdapter(name, manifest) {
  if (name === "officeimo") return { name, api: await import(pathToFileURL(manifest.officeimo.path).href) };
  const entry = packages[name]; if (!entry) throw Error("Unknown comparison engine: " + name);
  const require = createRequire(manifest.dependencies.packageJson);
  const imported = await import(pathToFileURL(require.resolve(entry.package)).href);
  const api = imported.default ?? imported;
  const plugin = entry.plugin ? await import(pathToFileURL(require.resolve(entry.plugin.package)).href) : undefined;
  return { name, api, imported, plugin };
}

async function officeimo(api, spec, fonts) {
  const columns = headings(spec).map(header => ({ header, width: 20 }));
  if (spec.format === "csv") return new Uint8Array(await (await api.writeCsv(rows(spec), {
    columns, delimiter: ",", lineEnding: "\r\n", quote: "all", bom: false, formulaInjectionProtection: true
  })).arrayBuffer());
  if (spec.format === "xlsx") return new Uint8Array(await (await api.writeXlsx(rows(spec), { columns,
    sheet: { name: "Data", autoSize: { sampleRows: 0 }, autoFilter: false, freezeHeader: false, boldHeader: false }
  })).arrayBuffer());
  return new Uint8Array(await (await api.writePdf(rows(spec), { columns, fonts: {
    regular: new api.PdfFont(fonts.regular), bold: new api.PdfFont(fonts.bold)
  }, pageSize: "A3", orientation: "landscape", fontSize: 6, pageNumbers: false,
    columnWidths: Array(spec.columns).fill(width(spec))
  })).arrayBuffer());
}

async function pdfmake(api, spec, fonts) {
  // Use the documented Node filename contract; every PDF lane includes file intake.
  api.setUrlAccessPolicy(() => false);
  api.setLocalAccessPolicy(file => file === fonts.regularPath || file === fonts.boldPath);
  api.addFonts({ Report: { normal: fonts.regularPath, bold: fonts.boldPath, italics: fonts.regularPath, bolditalics: fonts.boldPath } });
  const doc = { pageSize: "A3", pageOrientation: "landscape", pageMargins: [36, 36, 36, 36],
    defaultStyle: { font: "Report", fontSize: 6, lineHeight: fonts.regularSpacing, color: "#111827" },
    content: [{ table: { headerRows: 1, widths: Array(spec.columns).fill(width(spec) - 8 - .4 / spec.columns),
      body: [headings(spec).map(text => ({ text, bold: true, lineHeight: fonts.boldSpacing })), ...rows(spec)] },
      layout: { hLineWidth: () => .4, vLineWidth: () => .4, hLineColor: () => "#d1d9e3", vLineColor: () => "#d1d9e3",
        paddingLeft: () => 3.8, paddingRight: () => 3.8, paddingTop: () => 3.8, paddingBottom: () => 3.8,
        fillColor: row => row === 0 ? "#e7edf5" : null } }] };
  return new Uint8Array(await api.createPdf(doc).getBuffer());
}

async function jspdf(adapter, spec, fonts) {
  const Pdf = adapter.imported.jsPDF ?? adapter.api.jsPDF;
  const autoTable = adapter.plugin.autoTable ?? adapter.plugin.default?.autoTable;
  if (typeof Pdf !== "function" || typeof autoTable !== "function") throw Error("Pinned jsPDF/AutoTable public exports differ.");
  const doc = new Pdf({ orientation: "landscape", unit: "pt", format: "a3", compress: true, putOnlyUsedFonts: true });
  doc.addFileToVFS("Report-Regular.ttf", Buffer.from(fonts.regular).toString("base64"));
  doc.addFileToVFS("Report-Bold.ttf", Buffer.from(fonts.bold).toString("base64"));
  doc.addFont("Report-Regular.ttf", "Report", "normal"); doc.addFont("Report-Bold.ttf", "Report", "bold");
  autoTable(doc, { head: [headings(spec)], body: Array.from(rows(spec)), theme: "grid", margin: 36, showHead: "everyPage",
    styles: { font: "Report", fontStyle: "normal", fontSize: 6, cellPadding: 3.8, textColor: [17, 24, 39],
      lineColor: [209, 217, 227], lineWidth: .4, overflow: "linebreak" },
    headStyles: { fontStyle: "bold", fillColor: [231, 237, 245], textColor: [17, 24, 39], halign: "left" },
    columnStyles: Object.fromEntries(Array.from({ length: spec.columns }, (_, column) => [column, { cellWidth: width(spec) }])) });
  return new Uint8Array(doc.output("arraybuffer"));
}

async function pdfkit(Pdf, spec, fonts) {
  // Continuation rows are measured from the top margin before pageAdded paints the header.
  const doc = new Pdf({ size: "A3", layout: "landscape", margins: { top: 52, bottom: 36, left: 36, right: 36 },
    autoFirstPage: false, compress: true });
  const chunks = [], completed = new Promise((resolve, reject) => {
    doc.on("data", chunk => chunks.push(Buffer.from(chunk)));
    doc.once("end", () => resolve(Buffer.concat(chunks))); doc.once("error", reject);
  });
  void completed.catch(() => {});
  doc.registerFont("Report", Buffer.from(fonts.regular)); doc.registerFont("Report-Bold", Buffer.from(fonts.bold));
  doc.on("pageAdded", () => {
    const headerHeight = 16;
    doc.save().rect(36, 36, 1190.551 - 72, headerHeight).fillAndStroke("#e7edf5", "#d1d9e3").restore();
    doc.font("Report-Bold").fontSize(6).fillColor("#111827");
    headings(spec).forEach((text, column) => doc.text(text, 36 + column * width(spec) + 3.8, 39.8,
      { width: width(spec) - 7.6, lineBreak: false }));
    doc.font("Report").fontSize(6); doc.x = 36; doc.y = 36 + headerHeight;
  });
  try {
    doc.addPage();
    await doc.table({ data: rows(spec), columnStyles: Array(spec.columns).fill(width(spec)),
      defaultStyle: { font: "Report", fontSize: 6, padding: 3.8, border: .4, borderColor: "#d1d9e3", textColor: "#111827" } });
    doc.end(); return new Uint8Array(await completed);
  } catch (error) { doc.destroy(error); throw error; }
}

async function exceljsStream(api, spec) {
  const output = new PassThrough(), chunks = [];
  let total = 0;
  const completed = new Promise((resolve, reject) => {
    output.on("data", chunk => {
      total += chunk.length;
      if (total > 512 * 1024 * 1024) { output.destroy(new RangeError("Streaming comparison output exceeded 512 MiB.")); return; }
      chunks.push(Buffer.from(chunk));
    });
    output.once("end", () => resolve(Buffer.concat(chunks, total)));
    output.once("error", reject);
    output.once("close", () => { if (!output.readableEnded) reject(new Error("Streaming XLSX output closed before EOF.")); });
  });
  void completed.catch(() => {});
  try {
    const workbook = new api.stream.xlsx.WorkbookWriter({ stream: output, useSharedStrings: false, useStyles: false,
      zip: { zlib: { level: 6 } } });
    const sheet = workbook.addWorksheet("Data");
    sheet.columns = headings(spec).map(header => ({ header, width: 20 }));
    sheet.getRow(1).commit();
    for (const values of rows(spec)) sheet.addRow(values).commit();
    sheet.commit();
    const commit = workbook.commit();
    void commit.catch(error => output.destroy(error));
    // Workbook commit waits for writable finish; EOF and concatenation also belong inside generation.
    const [, bytes] = await Promise.all([commit, completed]);
    return new Uint8Array(bytes);
  } catch (error) { output.destroy(error); throw error; }
  finally { if (!output.destroyed) output.destroy(); chunks.length = 0; }
}

/** Materialize the same complete UTF-8/binary buffer for every selected engine. */
export async function exportBuffer(adapter, spec, fonts) {
  if (spec.format === "pdf" && adapter.name !== "pdfmake") fonts = { ...fonts,
    regular: new Uint8Array(await readFile(fonts.regularPath)), bold: new Uint8Array(await readFile(fonts.boldPath)) };
  if (adapter.name === "officeimo") return officeimo(adapter.api, spec, fonts);
  if (adapter.name === "papaparse") return new TextEncoder().encode(adapter.api.unparse({ fields: headings(spec), data: Array.from(rows(spec)) },
    { quotes: true, delimiter: ",", newline: "\r\n", header: true, escapeFormulae: true }) + "\r\n");
  if (adapter.name === "exceljs") {
    const workbook = new adapter.api.Workbook(), sheet = workbook.addWorksheet("Data");
    sheet.columns = headings(spec).map(header => ({ header, width: 20 }));
    for (const values of rows(spec)) sheet.addRow(values);
    return new Uint8Array(await workbook.xlsx.writeBuffer({ zip: { zlib: { level: 6 } } }));
  }
  if (adapter.name === "exceljs-stream") return exceljsStream(adapter.api, spec);
  if (adapter.name === "sheetjs") {
    const sheet = adapter.api.utils.aoa_to_sheet([headings(spec), ...rows(spec)], { dense: true });
    sheet["!cols"] = Array.from({ length: spec.columns }, () => ({ width: 20 }));
    const workbook = adapter.api.utils.book_new(); adapter.api.utils.book_append_sheet(workbook, sheet, "Data");
    return new Uint8Array(adapter.api.write(workbook, { type: "buffer", bookType: "xlsx", bookSST: false, compression: true }));
  }
  if (adapter.name === "pdfmake") return pdfmake(adapter.api, spec, fonts);
  if (adapter.name === "jspdf") return jspdf(adapter, spec, fonts);
  if (adapter.name === "pdfkit") return pdfkit(adapter.api, spec, fonts);
  throw Error("No public adapter for " + adapter.name);
}
