// Shared executable example: the website, workers and artifact qualification call this operation.
// Hosts supply rows; document policy stays here and transport stays in OfficeIMO/execution.
export const reportColumns = Object.freeze([
  { header: "Service", key: "service", width: 20 },
  { header: "Site", key: "site", width: 14 },
  { header: "Status", key: "status", width: 14 },
  { header: "Latency (ms)", key: "latency", type: "number", width: 16, format: "0.0" },
  { header: "Availability", key: "availability", type: "number", width: 16, format: "0.00%" },
  { header: "Checked (UTC)", key: "checked", type: "date", width: 22, format: "yyyy-mm-dd hh:mm" }
].map(column => Object.freeze(column)));

/** Illustrative operational data, generated lazily with stable numeric and date values. */
export function* operationalRows(count) {
  if (!Number.isSafeInteger(count) || count < 0 || count > 100000) throw new RangeError("Choose 0 through 100,000 services.");
  const sites = ["Warsaw", "Łódź", "Praha", "Berlin"];
  for (let index = 0; index < count; index++) yield [
    "WEB-" + String(index + 1).padStart(6, "0"), sites[index % sites.length],
    index % 37 === 0 ? "Offline" : index % 11 === 0 ? "Attention" : "Healthy",
    (index % 97) / 10 + 4.75, 0.995 - (index % 17) / 10000,
    new Date(Date.UTC(2026, 9, 19, 8, index % 60))
  ];
}

export async function writeOperationalReport(api, { input, rows, sink, signal, onProgress }) {
  const options = { signal, onProgress, limits: { maxRows: 100000, maxCells: 700000, maxOutputBytes: 96 * 1024 * 1024 } };
  if (input.format === "csv") return api.writeCsvTo(rows, sink, { ...options, columns: reportColumns, bom: true });
  let count = 0;
  const statuses = { Healthy: 0, Attention: 0, Offline: 0 };
  async function* track() {
    for await (const row of rows) {
      count++;
      if (!(row[2] in statuses)) throw new TypeError("Unexpected service status.");
      statuses[row[2]]++;
      yield row;
    }
  }
  if (input.format === "xlsx") {
    const book = new api.Workbook({ ...options, sink, dateMode: "utc", creator: "OfficeIMO report workbench" });
    try {
      const sheet = book.addWorksheet("Services", {
        columns: reportColumns, title: { text: "Service health report", height: 32 },
        freezeHeader: true, freezeColumns: 1, table: { name: "Services", style: "TableStyleMedium9" },
        autoSize: { sampleRows: 128 }, rowHeight: 22, print: { repeatHeaders: true },
        footer: { values: ["Services"], totals: { latency: "average", availability: "average" } },
        cellStyle: ({ column, value }) => column.key === "status" && value !== "Healthy"
          ? { font: { bold: true, color: value === "Offline" ? "991B1B" : "92400E" }, fill: { color: value === "Offline" ? "FEE2E2" : "FEF3C7" } } : undefined
      });
      await sheet.addRows(track()); await sheet.close();
      const summary = book.addWorksheet("Overview", { columns: [{ header: "Status", width: 20 }, { header: "Services", width: 14, type: "number" }] });
      await summary.addRows(Object.entries(statuses));
      summary.addChart({ type: "columnClustered", title: "Service health", row: 6, column: 1, width: 640, height: 340,
        data: { categories: Object.keys(statuses), series: [{ name: "Services", values: Object.values(statuses), color: "2563EB" }] },
        legend: false, dataLabels: { showValue: true }, valueAxis: { title: "Services", minimum: 0 },
        description: "Editable service counts by status, backed by worksheet cells" });
      const result = await book.finish();
      return { rows: count, columns: reportColumns.length, bytes: result.bytes };
    } catch (error) { await book.discard(error); throw error; }
  }
  if (input.format !== "pdf") throw new TypeError("Choose Excel, PDF or CSV.");
  if (!input.fonts?.regular || !input.fonts?.bold) throw new TypeError("PDF requires the example's embedded regular and bold fonts.");
  const fonts = { regular: await api.PdfFont.create(input.fonts.regular, { signal }), bold: await api.PdfFont.create(input.fonts.bold, { signal }) };
  const formats = [null, null, null, new Intl.NumberFormat("en", { minimumFractionDigits: 1, maximumFractionDigits: 1 }),
    new Intl.NumberFormat("en", { style: "percent", minimumFractionDigits: 2, maximumFractionDigits: 2 })];
  function* blocks() {
    yield { kind: "paragraph", headingLevel: 1, fontSize: 22, keepWithNext: true,
      text: [{ text: "Service health ", bold: true, color: "163C75" }, { text: "report", color: "2563EB" }] };
    yield { kind: "group", blocks: [
      { kind: "paragraph", text: "Operational reporting · 19 October 2026", presentation: { color: "475569" } },
      { kind: "paragraph", text: [{ text: "Illustrative data. ", bold: true }, { text: "Typed values, embedded fonts and a streamed source. Generated locally." }] }
    ] };
    yield { kind: "paragraph", headingLevel: 2, fontSize: 13, text: "Service detail", keepWithNext: true, spaceBefore: 10 };
    yield api.pdfTable(track(), { columns: reportColumns, alternateRowColor: "F3F6FB",
      headerPresentation: { background: "163C75", color: "FFFFFF", bold: true },
      formatRuns(value, context) {
        if (context.columnIndex === 2) return [{ text: String(value), bold: true,
          color: value === "Offline" ? "991B1B" : value === "Attention" ? "92400E" : "166534" }];
        return [{ text: value instanceof Date ? value.toISOString().slice(0, 16).replace("T", " ") : formats[context.columnIndex]?.format(value) ?? String(value ?? "") }];
      } });
    yield { kind: "paragraph", fontSize: 11, keepTogether: true, spaceBefore: 12,
      text: [{ text: count.toLocaleString("en") + " services. ", bold: true },
        { text: Object.entries(statuses).map(([name, value]) => name + ": " + value.toLocaleString("en")).join(" · ") }] };
  }
  return api.writePdfDocumentTo(blocks(), sink, { ...options, fonts, pageSize: "A4", orientation: "landscape", margins: 32,
    tagged: { language: "en" }, pageHeader: "OFFICEIMO / OPERATIONS", pageFooter: "Browser-local reference report", pageNumbers: true });
}
