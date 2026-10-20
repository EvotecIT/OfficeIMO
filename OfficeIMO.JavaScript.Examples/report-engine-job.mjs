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
  const document = await serviceReliabilityDocument(api, { input, rows, signal, onProgress });
  return api.writePdfDocumentTo(document.blocks, sink, document.options);
}

/** Constant-space summary for the illustrative source, without opening or consuming its iterator. */
export function operationalSummary(count) {
  if (!Number.isSafeInteger(count) || count < 0 || count > 100000) throw new RangeError("Choose 0 through 100,000 services.");
  const multiples = divisor => count ? Math.floor((count - 1) / divisor) + 1 : 0;
  const Offline = multiples(37), Attention = multiples(11) - multiples(407);
  return Object.freeze({ total: count, Healthy: count - Offline - Attention, Attention, Offline });
}

/**
 * Build one report from caller-owned summary metadata and a single-use selected row source.
 * Hosts can pass the returned blocks/options to either PDF delivery API. Graphics are supplied
 * by their existing owner; this example does not render charts or inspect a live grid.
 */
export async function serviceReliabilityDocument(api, { input, rows, signal, onProgress, columns = reportColumns }) {
  if (!input.fonts?.regular || !input.fonts?.bold) throw new TypeError("PDF requires the example's embedded regular and bold fonts.");
  const supplied = input.summary ?? operationalSummary(input.rowCount);
  const summary = Object.freeze(Object.fromEntries(["total", "Healthy", "Attention", "Offline"].map(key => [key, supplied[key]])));
  for (const key of ["total", "Healthy", "Attention", "Offline"])
    if (!Number.isSafeInteger(summary[key]) || summary[key] < 0) throw new TypeError("Report summary requires nonnegative integer service counts.");
  if (summary.Healthy + summary.Attention + summary.Offline !== summary.total) throw new TypeError("Report summary counts must match its total.");
  const statusColumn = columns.findIndex(column => column.header === "Status");
  if (statusColumn < 0) throw new TypeError("The service report requires its Status column.");
  const fonts = { regular: await api.PdfFont.create(input.fonts.regular, { signal }), bold: await api.PdfFont.create(input.fonts.bold, { signal }) };
  const decimal = new Intl.NumberFormat("en", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const percent = new Intl.NumberFormat("en", { style: "percent", minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const statuses = { Healthy: 0, Attention: 0, Offline: 0 }; let count = 0;
  async function* track() {
    for await (const row of rows) {
      const cell = row[statusColumn], status = cell instanceof api.ExportCell ? cell.value : cell;
      if (!Object.hasOwn(statuses, status)) throw new TypeError("Unexpected service status.");
      statuses[status]++; count++; yield row;
    }
    if (count !== summary.total || Object.entries(statuses).some(([name, value]) => summary[name] !== value))
      throw new TypeError("The selected rows differ from the supplied report summary.");
  }
  const scope = input.scope ?? "Illustrative operational services · 19 October 2026";
  const findings = [
    { text: [{ text: "Restore unavailable services. ", bold: true, color: "991B1B" },
      { text: summary.Offline.toLocaleString("en") + " offline services require an operations owner. Verify the last successful check, restore the service, and record the recovery evidence. " },
      { text: "Recovery guidance", link: { target: "https://example.com/operations/recovery" } }] },
    { text: [{ text: "Review services needing attention. ", bold: true, color: "92400E" },
      { text: summary.Attention.toLocaleString("en") + " services need a latency and availability review before the next change window. Compare the selected inventory with the service's agreed operating target, assign the follow-up, and confirm the outcome at the next check." }] },
    { text: [{ text: "Retain the healthy baseline. ", bold: true, color: "166534" },
      { text: summary.Healthy.toLocaleString("en") + " healthy services remain in routine verification. Keep the exported scope and check times with this report so the next review can explain any change." }], keepTogether: true }
  ];
  async function* blocks() {
    yield { kind: "section", title: "Service health report", level: 1, headingStyle: "reportTitle" };
    yield { kind: "paragraph", style: "muted", text: scope };
    yield { kind: "paragraph", style: "body", text: [{ text: summary.total.toLocaleString("en") + " services. ", bold: true },
      { text: "Healthy: " + summary.Healthy.toLocaleString("en") + " · Attention: " + summary.Attention.toLocaleString("en") + " · Offline: " + summary.Offline.toLocaleString("en") }] };
    yield { kind: "paragraph", style: "sectionHeading", headingLevel: 2, text: "Findings and ownership", keepWithNext: true };
    yield { kind: "list", ordered: true, start: 1, itemStyle: "finding", items: findings };
    if (input.figure) yield { kind: "image", image: input.figure.image, width: input.figure.width ?? 420,
      alternativeText: input.figure.alternativeText,
      caption: { text: input.figure.caption, style: "caption" } };
    yield { kind: "section", title: "Inventory appendix", level: 2, headingStyle: "sectionHeading", orientation: "landscape" };
    yield { ...api.pdfTable(track(), { columns, alternateRowColor: "F3F6FB",
      headerPresentation: { background: "163C75", color: "FFFFFF", bold: true },
      formatRuns(value, context) {
        if (context.columnIndex === statusColumn) return [{ text: String(value), bold: true,
          color: value === "Offline" ? "991B1B" : value === "Attention" ? "92400E" : "166534" }];
        const header = columns[context.columnIndex].header;
        return [{ text: value instanceof Date ? value.toISOString().slice(0, 16).replace("T", " ")
          : header === "Latency (ms)" && typeof value === "number" ? decimal.format(value)
          : header === "Availability" && typeof value === "number" ? percent.format(value) : String(value ?? "") }];
      } }), caption: { text: "Table 1. Selected service inventory; check times are UTC.", style: "caption" } };
    yield { kind: "section", title: "Next actions", level: 2, headingStyle: "sectionHeading" };
    yield { kind: "paragraph", style: "body", text: "Assign an owner to each unresolved finding, retain the selected inventory, and compare the next review against the same scope." };
    yield { kind: "paragraph", style: "muted", text: "Stream verification: " + count.toLocaleString("en") + " selected services matched the supplied summary. Check times and report scope describe this export, not a live monitoring guarantee." };
  }
  return { blocks: blocks(), options: { signal, onProgress,
    limits: { maxRows: 100000, maxCells: 700000, maxOutputBytes: 96 * 1024 * 1024 }, fonts,
    pageSize: "A4", orientation: "portrait", margins: 36, fontSize: 10,
    paragraphStyles: {
      reportTitle: { fontSize: 22, presentation: { bold: true, color: "163C75" }, spaceAfter: 12 },
      sectionHeading: { fontSize: 15, presentation: { bold: true, color: "163C75" }, spaceBefore: 8, spaceAfter: 8 },
      body: { fontSize: 10, spaceAfter: 10 },
      muted: { fontSize: 9, presentation: { color: "475569" }, spaceAfter: 10 },
      finding: { fontSize: 10, spaceAfter: 9 },
      caption: { fontSize: 9, presentation: { italic: true, color: "475569" }, spaceBefore: 8, spaceAfter: 5 }
    },
    tagged: { language: "en" }, pageHeader: context => "OFFICEIMO / OPERATIONS · " + context.sectionTitle,
    pageFooter: context => "Browser-local reference report · Section " + context.sectionNumber + " / " + context.sectionPageNumber,
    pageNumbers: true } };
}
