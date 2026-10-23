// Shared executable example: the website, workers and artifact qualification call this operation.
// Hosts supply rows; document policy stays here and transport stays in OfficeIMO/execution.
export const reportColumns = Object.freeze([
  { header: "Service", key: "service", width: 20, rowHeader: true },
  { header: "Site", key: "site", width: 14 },
  { header: "Status", key: "status", width: 14 },
  { header: "Latency (ms)", key: "latency", type: "number", width: 16, format: "0.0" },
  { header: "Availability", key: "availability", type: "number", width: 16, format: "0.00%" },
  { header: "Checked (UTC)", key: "checked", type: "date", width: 22, format: "yyyy-mm-dd hh:mm" }
].map(column => Object.freeze(column)));

/** Illustrative operational data, generated lazily with stable numeric and date values. */
export function* operationalRows(count) {
  if (!Number.isSafeInteger(count) || count < 0 || count > 100000) throw new RangeError("Choose 0 through 100,000 services.");
  for (let index = 0; index < count; index++) yield operationalRow(index);
}

function operationalRow(index) {
  const sites = ["Warsaw", "Łódź", "Praha", "Berlin"];
  return [
    "WEB-" + String(index + 1).padStart(6, "0"), sites[index % sites.length],
    index % 37 === 0 ? "Offline" : index % 11 === 0 ? "Attention" : "Healthy",
    (index % 97) / 10 + 4.75, 0.995 - (index % 17) / 10000,
    new Date(Date.UTC(2026, 9, 19, 8, index % 60))
  ];
}

/** The host supplies the intended group ordering; the document engine never sorts or buffers records. */
export function* operationalRowsBySite(count) {
  if (!Number.isSafeInteger(count) || count < 0 || count > 100000) throw new RangeError("Choose 0 through 100,000 services.");
  for (const residue of [3, 1, 2, 0]) for (let index = residue; index < count; index += 4) yield operationalRow(index);
}

// Explicit example translation resources. Engine grouping uses stable column keys and raw values.
const groupedReportText = Object.freeze({
  en: Object.freeze({ headers: { service: "Service", site: "Site", status: "Status", latency: "Latency (ms)", availability: "Availability", checked: "Checked (UTC)" },
    group: "Group", site: "Site: ", subtotal: "Group total", total: "All selected services", records: "services", continued: " (continued)",
    title: "Services by site", page: { prefix: "Page ", separator: " of ", suffix: "" }, booleanText: { true: "True", false: "False" } }),
  pl: Object.freeze({ headers: { service: "Usługa", site: "Lokalizacja", status: "Stan", latency: "Opóźnienie (ms)", availability: "Dostępność", checked: "Sprawdzone (UTC)" },
    group: "Grupa", site: "Lokalizacja: ", subtotal: "Suma grupy", total: "Wszystkie wybrane usługi", records: "usług", continued: " (ciąg dalszy)",
    title: "Usługi według lokalizacji", page: { prefix: "Strona ", separator: " z ", suffix: "" }, booleanText: { true: "Tak", false: "Nie" } })
});

/** A thin grouped-report operation over an already ordered, selected portable capture. */
export async function writeGroupedOperationalReport(api, { input = {}, rows, columns = reportColumns, sink, signal, onProgress }) {
  const language = input.language ?? "en";
  if (!Object.hasOwn(groupedReportText, language)) throw new TypeError("The grouped example supports en and pl report resources.");
  const text = groupedReportText[language];
  const selectedColumns = columns.map(column => ({ ...column,
    ...(input.language !== undefined && text.headers[column.key] ? { header: text.headers[column.key] } : {}) }));
  const totals = Object.fromEntries(["latency", "availability"].filter(key => selectedColumns.some(column => column.key === key)).map(key => [key, "average"]));
  const report = api.createGroupedReport(rows, { columns: selectedColumns, groupBy: "site", totals, signal,
    groupLabel: value => text.site + String(value ?? ""), maxGroups: 1000, limits: { maxRows: 100000, maxCells: 700000 } });
  const common = { signal, onProgress, labelColumn: { header: text.group, width: 30 },
    subtotalLabel: group => text.subtotal + " · " + group.rows.toLocaleString(language) + " " + text.records,
    grandTotalLabel: summary => text.total + " · " + summary.rows.toLocaleString(language) + " " + text.records,
    limits: { maxRows: 103001, maxCells: 800008, maxOutputBytes: 96 * 1024 * 1024 } };
  if (input.format === "xlsx") {
    const options = { ...common, dateMode: "utc", compression: input.compression ?? "auto",
      sheet: { name: text.title, title: { text: text.title }, freezeHeader: true, freezeColumns: 1, rowHeight: 22,
        headerFill: "163C75", print: { orientation: "landscape", repeatHeaders: true } },
      groupStyle: { font: { bold: true, color: "163C75" }, fill: { color: "DBEAFE" } },
      subtotalStyle: { font: { bold: true }, fill: { color: "EEF2F6" } }, grandTotalStyle: { font: { bold: true }, fill: { color: "DBEAFE" } } };
    return sink === undefined ? api.writeGroupedXlsx(report, options) : api.writeGroupedXlsxTo(report, sink, options);
  }
  if (input.format !== "pdf") throw new TypeError("Grouped reports support Excel or PDF; CSV exports ordinary records.");
  if (!input.fonts?.regular || !input.fonts?.bold) throw new TypeError("PDF requires the example's embedded regular and bold fonts.");
  const fonts = { regular: await api.PdfFont.create(input.fonts.regular, { signal }), bold: await api.PdfFont.create(input.fonts.bold, { signal }) };
  const decimal = new Intl.NumberFormat(language, { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const percent = new Intl.NumberFormat(language, { style: "percent", minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const options = { ...common, labelColumn: { header: text.group, width: 26 }, fonts, pageSize: "A4", orientation: "landscape", margins: 30,
    tagged: { language }, title: text.title,
    fontSize: 9, pageNumberLabels: text.page, booleanText: text.booleanText, headerPresentation: { background: "163C75", color: "FFFFFF", bold: true },
    continuationLabel: group => group.label + text.continued,
    formatValue(value, context) { return value instanceof Date ? value.toISOString().slice(0, 16).replace("T", " ")
      : typeof value === "number" && context.column.key === "latency" ? decimal.format(value)
      : typeof value === "number" && context.column.key === "availability" ? percent.format(value) : String(value ?? ""); } };
  return sink === undefined ? api.writeGroupedPdf(report, options) : api.writeGroupedPdfTo(report, sink, options);
}

export async function writeOperationalReport(api, { input, rows, sink, signal, onProgress }) {
  if (input.layout !== undefined && input.layout !== "analysis" && input.layout !== "grouped") throw new TypeError("Choose analysis or grouped report layout.");
  if (input.layout === "grouped") return writeGroupedOperationalReport(api, { input, rows, sink, signal, onProgress });
  const options = { signal, onProgress, limits: { maxRows: 100000, maxCells: 700000, maxOutputBytes: 96 * 1024 * 1024 } };
  if (input.format === "csv") return api.writeCsvTo(rows, sink, { ...options, columns: reportColumns, bom: true });
  if (input.format === "xlsx") return writeServiceAnalysis(api, { input, rows, sink, signal, onProgress });
  if (input.format !== "pdf") throw new TypeError("Choose Excel, PDF or CSV.");
  const document = await serviceReliabilityDocument(api, { input, rows, signal, onProgress });
  return api.writePdfDocumentTo(document.blocks, sink, document.options);
}

/**
 * Export one selected portable capture as an editable analysis workbook. Initial counts are
 * collected in one pass; Excel recalculates the explicit formulas after an inventory edit.
 * Counts cover every exported row, including rows later hidden by an Excel filter.
 * Omitting sink returns a Blob; a supplied destination receives ordinary completion counts.
 */
export async function writeServiceAnalysis(api, { input = {}, rows, columns = reportColumns, sink, signal, onProgress }) {
  const statuses = ["Healthy", "Attention", "Offline"], counts = { Healthy: 0, Attention: 0, Offline: 0 };
  const statusColumns = columns.map((column, index) => column.key === "status" ? index : -1).filter(index => index >= 0);
  if (statusColumns.length !== 1) throw new TypeError("Service analysis requires one selected column with key 'status'.");
  const statusColumn = statusColumns[0];
  const colors = { Healthy: { color: "166534", background: "DCFCE7" }, Attention: { color: "92400E", background: "FEF3C7" }, Offline: { color: "991B1B", background: "FEE2E2" } };
  let count = 0;
  async function* track() {
    for await (const row of rows) {
      const cell = row[statusColumn], status = cell instanceof api.ExportCell ? cell.value : cell;
      if (!Object.hasOwn(counts, status)) throw new TypeError("Unexpected service status.");
      if (count === 100000) throw new RangeError("Service analysis supports at most 100,000 selected services.");
      counts[status]++; count++; yield row;
    }
  }
  const book = new api.Workbook({ sink, signal, dateMode: "utc", compression: input.compression ?? "auto", creator: "OfficeIMO report workbench",
    limits: { maxRows: 100003, maxCells: 700006, maxOutputBytes: 96 * 1024 * 1024 },
    onProgress: event => event.phase === "rows" && event.sheetName !== "Services" ? undefined : onProgress?.({ ...event, rows: count }) });
  try {
    const totals = Object.fromEntries(["latency", "availability"].filter(key => columns.some(column => column.key === key)).map(key => [key, "average"]));
    const services = book.addWorksheet("Services", { columns, title: { text: "Service health report", height: 32 },
      freezeHeader: true, freezeColumns: 1, table: { name: "Services", style: "TableStyleMedium9" },
      autoSize: { sampleRows: 128 }, rowHeight: 22, print: { repeatHeaders: true },
      footer: { values: ["Services"], totals },
      dataValidation: [{ type: "list", column: "status", items: statuses, allowBlank: false, showErrorMessage: true,
        errorStyle: "stop", errorTitle: "Choose a service status", errorMessage: "Choose Healthy, Attention or Offline." }],
      conditionalFormats: statuses.map(status => ({ type: "cellIs", range: { column: "status" }, operator: "equal", value: status,
        style: { font: { color: colors[status].color, bold: status !== "Healthy" }, fill: { color: colors[status].background } }, stopIfTrue: true }))
    });
    await services.addRows(track()); await services.close();
    const sourceRange = services.dataRange("status");
    const overview = book.addWorksheet("Overview", { columns: [
      { header: "Status", key: "status", width: 20 }, { header: "Services", key: "count", width: 14, type: "number" }
    ] });
    overview.addChart({ type: "columnClustered", title: "Service health", row: 6, column: 1, width: 640, height: 340,
      data: { kind: "range", sheet: overview, rowCount: 3, categoryColumn: "status", series: [{ column: "count", color: "2563EB" }] },
      legend: false, dataLabels: { showValue: true }, valueAxis: { title: "Services", minimum: 0 },
      description: "Service counts calculated from the exported inventory; edit Services status to update this chart" });
    await overview.addRows(statuses.map(status => [status,
      new api.FormulaCell(sourceRange ? "COUNTIF(" + sourceRange + ",\"" + status + "\")" : "0", counts[status])]));
    if (sink === undefined) return await book.toBlob();
    const result = await book.finish(); return { rows: count, columns: columns.length, bytes: result.bytes };
  } catch (error) { await book.discard(error); throw error; }
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
 * Column keys identify service roles; display headings may be translated or renamed.
 */
export async function serviceReliabilityDocument(api, { input, rows, signal, onProgress, columns = reportColumns }) {
  if (!input.fonts?.regular || !input.fonts?.bold) throw new TypeError("PDF requires the example's embedded regular and bold fonts.");
  const supplied = input.summary ?? operationalSummary(input.rowCount);
  const summary = Object.freeze(Object.fromEntries(["total", "Healthy", "Attention", "Offline"].map(key => [key, supplied[key]])));
  for (const key of ["total", "Healthy", "Attention", "Offline"])
    if (!Number.isSafeInteger(summary[key]) || summary[key] < 0) throw new TypeError("Report summary requires nonnegative integer service counts.");
  if (summary.Healthy + summary.Attention + summary.Offline !== summary.total) throw new TypeError("Report summary counts must match its total.");
  const statusColumns = columns.map((column, index) => column.key === "status" ? index : -1).filter(index => index >= 0);
  if (statusColumns.length !== 1) throw new TypeError("The service report requires one selected column with key 'status'.");
  const statusColumn = statusColumns[0];
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
        const key = columns[context.columnIndex].key;
        return [{ text: value instanceof Date ? value.toISOString().slice(0, 16).replace("T", " ")
          : key === "latency" && typeof value === "number" ? decimal.format(value)
          : key === "availability" && typeof value === "number" ? percent.format(value) : String(value ?? "") }];
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
